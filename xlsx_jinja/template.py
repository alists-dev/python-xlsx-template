from __future__ import annotations

import base64
import binascii
import copy
import datetime as dt
import io
import math
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, BinaryIO
from xml.etree import ElementTree as ET

from jinja2 import Environment, meta
from PIL import Image

from ._tags import (
    BLOCK_MARKER_RE as _BLOCK_MARKER_RE,
    IMAGE_TAG_RE as _IMAGE_TAG_RE,
    JINJA_RE as _JINJA_RE,
    STRUCTURAL_ROW_RE as _STRUCTURAL_ROW_RE,
    TR_RE as _TR_RE,
    XV_RE as _XV_RE,
)
from ._xml import (
    CONTENT_TYPES_NS,
    DRAWING_MAIN_NS,
    DRAWING_NS,
    MAIN_NS,
    NS,
    PACKAGE_REL_NS,
    REL_NS,
    q as _q,
    qn as _qn,
)
from .introspection import find_undeclared_variables
from .media import apply_media_replacements, compute_crc32
from .richtext import RichText, apply_richtext_cell, make_richtext_finalize

MAX_XLSX_ROW = 1_048_576
DRAWING_REL = f"{REL_NS}/drawing"
IMAGE_REL = f"{REL_NS}/image"

_FOR_STATEMENT_RE = re.compile(r"^for\s+(.+?)\s+in\s+(.+)$", re.DOTALL | re.IGNORECASE)
_LEADING_CONTROL_RE = re.compile(
    r"^\s*{%\s*((?:for|if|elif|else|endfor|endif)\b.*?)\s*%}",
    re.DOTALL | re.IGNORECASE,
)
_LEADING_SET_RE = re.compile(
    r"^\s*{%\s*(set\s+[^%]*?=.*?)\s*%}", re.DOTALL | re.IGNORECASE
)
_CONTROL_TAG_RE = re.compile(
    r"{%\s*(for|endfor|if|endif)\b.*?%}", re.DOTALL | re.IGNORECASE
)
_TRAILING_END_RE = re.compile(
    r"{%\s*((?:endfor|endif)\b.*?)\s*%}\s*$", re.DOTALL | re.IGNORECASE
)
_CELL_REF_RE = re.compile(r"^(\$?[A-Z]{1,3})\$?(\d+)$", re.IGNORECASE)
_FORMULA_REF_RE = re.compile(
    r"(?<![A-Z0-9_])(?P<c1>\$?[A-Z]{1,3})(?P<ra1>\$?)(?P<r1>\d+)"
    r"(?:\:(?P<c2>\$?[A-Z]{1,3})(?P<ra2>\$?)(?P<r2>\d+))?",
    re.IGNORECASE,
)
_RANGE_RE = re.compile(
    r"^(?P<c1>\$?[A-Z]{1,3})(?P<ra1>\$?)(?P<r1>\d+):"
    r"(?P<c2>\$?[A-Z]{1,3})(?P<ra2>\$?)(?P<r2>\d+)$",
    re.IGNORECASE,
)
_NATIVE_TOKEN_RE = re.compile(r"^__XLSX_JINJA_NATIVE_(\d+)__$")
_IMAGE_TOKEN_RE = re.compile(r"__XLSX_JINJA_IMAGE_(\d+)__")
_XMLNS_RE = re.compile(br'xmlns(?::([A-Za-z_][\w.-]*))?="([^"]+)"')


class XlsxJinjaError(Exception):
    pass


class UnsupportedFeatureError(XlsxJinjaError):
    pass


def _column_number(column: str) -> int:
    value = 0
    for char in column.replace("$", "").upper():
        value = value * 26 + ord(char) - 64
    return value


def _column_name(column: int) -> str:
    value = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        value = chr(65 + remainder) + value
    return value


def _source_bytes(source: str | Path | BinaryIO | bytes | bytearray) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, "read"):
        source.seek(0)
        return source.read()
    return Path(source).read_bytes()


def _register_namespaces(xml: bytes) -> list[tuple[bytes, bytes]]:
    declarations = _XMLNS_RE.findall(xml.split(b">", 1)[0])
    for prefix, uri in declarations:
        try:
            ET.register_namespace(prefix.decode(), uri.decode())
        except ValueError:
            pass
    return declarations


def _restore_missing_namespaces(
    xml: bytes, declarations: list[tuple[bytes, bytes]]
) -> bytes:
    declaration, separator, body = xml.partition(b"?>")
    head, close, tail = body.partition(b">")
    for prefix, uri in declarations:
        attribute = b"xmlns" + (b":" + prefix if prefix else b"") + b'="' + uri + b'"'
        if attribute not in head:
            head += b" " + attribute
    return declaration + separator + head + close + tail


def _png_bytes(value: Any) -> bytes | None:
    if not value:
        return None
    if isinstance(value, Image.Image):
        image = value.copy()
    else:
        if hasattr(value, "read"):
            value.seek(0)
            data = value.read()
        elif isinstance(value, (str, Path)):
            path = Path(value)
            try:
                data = path.read_bytes() if path.is_file() else str(value).encode()
            except OSError:
                data = str(value).encode()
        elif isinstance(value, bytes):
            data = value
        else:
            raise UnsupportedFeatureError(
                f"Unsupported image value: {type(value).__name__}."
            )
        candidates = [data]
        try:
            candidates.append(base64.b64decode(data, validate=False))
        except (ValueError, binascii.Error):
            pass
        image = None
        for candidate in candidates:
            try:
                image = Image.open(io.BytesIO(candidate))
                image.load()
                break
            except Exception:
                continue
        if image is None:
            raise UnsupportedFeatureError("Image value is not a valid image.")
    if image.mode not in {"1", "L", "LA", "P", "RGB", "RGBA", "I"}:
        image = image.convert("RGBA")
    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


class XlsxTemplate:
    """Render Jinja expressions directly in XLSX OOXML worksheet parts."""

    def __init__(self, template_file: str | Path | BinaryIO | bytes | bytearray):
        raw = _source_bytes(template_file)
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                self._infos = [copy.copy(info) for info in archive.infolist()]
                self._original_entries = {
                    info.filename: archive.read(info.filename) for info in self._infos
                }
        except zipfile.BadZipFile as exc:
            raise XlsxJinjaError("Template is not a valid XLSX package.") from exc

        if "xl/workbook.xml" not in self._original_entries:
            raise XlsxJinjaError("Template does not contain xl/workbook.xml.")

        self._entries = dict(self._original_entries)
        self._globals: dict[str, Any] = {}
        self._media_replacements: dict[int, bytes] = {}
        self._rendered = False

    @property
    def sheet_names(self) -> list[str]:
        return [name for name, _ in self._worksheet_parts()]

    def set_jinja_globals(self, **globals_: Any) -> None:
        self._globals.update(globals_)

    def replace_media(self, source: bytes, replacement: bytes) -> None:
        """Swap an embedded media part by content, matched via CRC32.

        Useful to change a picture (e.g. a per-tenant logo) without a full
        render pass. ``source`` must be the exact original bytes so its CRC32
        can be looked up among ``xl/media/*`` parts at save time.
        """
        self._media_replacements[compute_crc32(source)] = replacement

    def reset_replacements(self) -> None:
        """Clear any pending ``replace_media`` registrations."""
        self._media_replacements = {}

    def get_undeclared_template_variables(
        self, jinja_env: Environment | None = None
    ) -> set[str]:
        """Return Jinja variable names referenced anywhere in the workbook."""
        shared_strings = self._shared_strings()
        sheet_texts = [
            self._sheet_cell_texts(sheet_path, shared_strings)
            for _, sheet_path in self._worksheet_parts()
        ]
        return find_undeclared_variables(sheet_texts, jinja_env)

    def _sheet_cell_texts(
        self, sheet_path: str, shared_strings: list[str]
    ) -> list[str]:
        data = self._original_entries.get(sheet_path)
        if not data:
            return []
        root = ET.fromstring(data)
        sheet_data = root.find("x:sheetData", NS)
        if sheet_data is None:
            return []
        return [
            text
            for cell in sheet_data.iter(_q("c"))
            if (text := self._cell_text(cell, shared_strings)) is not None
        ]

    def render(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
        autoescape: bool = True,
    ) -> None:
        self._entries = dict(self._original_entries)
        shared_strings = self._shared_strings()
        sheets = self._worksheet_parts()
        comments = self._comment_controls(sheets)
        base_finalize = getattr(jinja_env, "finalize", None) if jinja_env else None

        for index, (sheet_name, sheet_path) in enumerate(sheets):
            if sheet_path not in self._entries:
                continue
            env = jinja_env or Environment(autoescape=autoescape)
            if jinja_env is not None:
                env.autoescape = autoescape
            env.globals.update(self._globals)
            richtext_values: list[RichText] = []
            env.finalize = make_richtext_finalize(richtext_values, base_finalize)
            sheet_context = dict(context)
            sheet_context.update(sheet_name=sheet_name, tpl_idx=index)
            self._entries[sheet_path] = self._render_sheet(
                sheet_path,
                self._entries[sheet_path],
                shared_strings,
                comments.get(sheet_path, {}),
                sheet_context,
                env,
                richtext_values,
            )

        self._rendered = True

    def save(self, output: str | Path | BinaryIO) -> None:
        target: BinaryIO
        close_target = False
        if hasattr(output, "write"):
            target = output
            target.seek(0)
            target.truncate(0)
        else:
            target = open(output, "wb")
            close_target = True

        entries = apply_media_replacements(self._entries, self._media_replacements)
        try:
            with zipfile.ZipFile(target, "w") as archive:
                original_names = {info.filename for info in self._infos}
                for info in self._infos:
                    archive.writestr(info, entries[info.filename])
                for name, data in entries.items():
                    if name not in original_names:
                        archive.writestr(name, data, zipfile.ZIP_DEFLATED)
        finally:
            if close_target:
                target.close()

    def _shared_strings(self) -> list[str]:
        data = self._entries.get("xl/sharedStrings.xml")
        if not data:
            return []
        root = ET.fromstring(data)
        return ["".join(node.text or "" for node in item.iter(_q("t"))) for item in root]

    def _worksheet_parts(self) -> list[tuple[str, str]]:
        workbook = ET.fromstring(self._entries["xl/workbook.xml"])
        rels_data = self._entries.get("xl/_rels/workbook.xml.rels")
        rel_targets: dict[str, str] = {}
        if rels_data:
            rels = ET.fromstring(rels_data)
            for rel in rels:
                target = rel.get("Target")
                if target:
                    rel_targets[rel.get("Id", "")] = self._resolve_part(
                        "xl/workbook.xml", target
                    )

        parts = []
        for sheet in workbook.findall("x:sheets/x:sheet", NS):
            target = rel_targets.get(sheet.get(f"{{{REL_NS}}}id", ""))
            if target and target.startswith("xl/worksheets/"):
                parts.append((sheet.get("name", "Sheet"), target))
        return parts

    @staticmethod
    def _resolve_part(source_part: str, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))

    def _comment_controls(
        self, sheets: list[tuple[str, str]]
    ) -> dict[str, dict[int, list[str]]]:
        result: dict[str, dict[int, list[str]]] = {}
        for _, sheet_path in sheets:
            rel_path = posixpath.join(
                posixpath.dirname(sheet_path),
                "_rels",
                posixpath.basename(sheet_path) + ".rels",
            )
            rel_data = self._entries.get(rel_path)
            if not rel_data:
                continue
            rels = ET.fromstring(rel_data)
            for rel in rels:
                if not rel.get("Type", "").endswith("/comments"):
                    continue
                target = rel.get("Target")
                if not target:
                    continue
                comment_path = self._resolve_part(sheet_path, target)
                comment_data = self._entries.get(comment_path)
                if not comment_data:
                    continue
                root = ET.fromstring(comment_data)
                controls = result.setdefault(sheet_path, {})
                for comment in root.findall("x:commentList/x:comment", NS):
                    ref = comment.get("ref", "")
                    match = _CELL_REF_RE.match(ref)
                    if not match:
                        continue
                    text = "".join(node.text or "" for node in comment.iter(_q("t")))
                    marker = re.match(r"^\s*beforerow\s*:?\s*(.*?)\s*$", text, re.DOTALL)
                    if marker and marker.group(1):
                        controls.setdefault(int(match.group(2)), []).append(marker.group(1))
        return result

    def _render_sheet(
        self,
        sheet_path: str,
        xml: bytes,
        shared_strings: list[str],
        comment_controls: dict[int, list[str]],
        context: dict[str, Any],
        env: Environment,
        richtext_values: list[RichText],
    ) -> bytes:
        namespaces = _register_namespaces(xml)
        root = ET.fromstring(xml)
        sheet_data = root.find("x:sheetData", NS)
        if sheet_data is None or not comment_controls and not any(
            _JINJA_RE.search(text)
            for cell in sheet_data.iter(_q("c"))
            if (text := self._cell_text(cell, shared_strings)) is not None
        ):
            return xml

        if any(
            _BLOCK_MARKER_RE.match(text)
            for cell in sheet_data.iter(_q("c"))
            if (text := self._cell_text(cell, shared_strings)) is not None
        ):
            expanded = _restore_missing_namespaces(
                self._expand_regions(
                    sheet_path, root, shared_strings, context, env
                ),
                namespaces,
            )
            return self._render_sheet(
                sheet_path,
                expanded,
                shared_strings,
                comment_controls,
                context,
                env,
                richtext_values,
            )

        source_rows = []
        removed_rows: set[int] = set()
        compact_rows: set[int] = set()
        template_parts: list[str] = []

        for row in list(sheet_data):
            old_row = int(row.get("r", "0") or 0)
            source_rows.append(old_row)
            self._normalize_template_cells(row, shared_strings)

            structural = self._structural_row_tag(row)
            if structural is not None:
                template_parts.append("{% " + structural + " %}")
                removed_rows.add(old_row)
                continue

            generated_before = row.attrib.pop("data-xlsx-jinja-before", "")
            before = ([generated_before] if generated_before else []) + list(
                comment_controls.get(old_row, [])
            )
            inline_before, inline_after, row_scoped = self._hoist_row_controls(row)
            if row_scoped:
                compact_rows.add(old_row)
            before.extend(inline_before)
            template_parts.extend(before)
            template_parts.append(ET.tostring(row, encoding="unicode"))
            template_parts.extend(inline_after)

        native_values: list[Any] = []
        image_values: list[tuple[str, Any, int]] = []

        def native(value: Any) -> str:
            native_values.append(value)
            return f"__XLSX_JINJA_NATIVE_{len(native_values) - 1}__"

        def image(mode: str, value: Any, index: int = 0) -> str:
            image_values.append((mode, value, int(index)))
            return f"__XLSX_JINJA_IMAGE_{len(image_values) - 1}__"

        env.globals.update(__xlsx_native=native, __xlsx_image=image)
        template = _XV_RE.sub(r"{{ __xlsx_native(\1) }}", "".join(template_parts))
        template = _IMAGE_TAG_RE.sub(
            lambda match: "{{ __xlsx_image(%r, %s) }}"
            % (match.group(1).lower(), match.group(2)),
            template,
        )

        rendered = env.from_string(template).render(context)
        try:
            rendered_root = ET.fromstring(f'<rows xmlns="{MAIN_NS}">{rendered}</rows>')
        except ET.ParseError as exc:
            raise XlsxJinjaError(f"Rendered worksheet XML is invalid: {exc}") from exc

        output_rows = list(rendered_root)
        output_old_rows = [int(row.get("r", "0") or 0) for row in output_rows]
        removed_rows.update(set(source_rows) - set(output_old_rows))
        row_map = self._renumber_rows(output_rows, removed_rows, compact_rows)

        image_requests: list[dict[str, Any]] = []
        for row in output_rows:
            old_row = int(row.get("data-xlsx-jinja-old-r", row.get("r", "0")))
            new_row = int(row.get("r", "0"))
            row.attrib.pop("data-xlsx-jinja-old-r", None)
            self._apply_native_values(row, native_values, richtext_values)
            self._collect_image_values(
                row, old_row, image_values, image_requests
            )
            self._translate_formulas(row, old_row, new_row, row_map)

        sheet_data[:] = output_rows
        self._render_images(sheet_path, root, row_map, image_requests)
        self._shift_merges(root, row_map)
        self._update_dimension(root, output_rows)
        rendered_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        return _restore_missing_namespaces(rendered_xml, namespaces)

    def _expand_regions(
        self,
        sheet_path: str,
        root: ET.Element,
        shared_strings: list[str],
        context: dict[str, Any],
        env: Environment,
    ) -> bytes:
        sheet_data = root.find("x:sheetData", NS)
        rows = {
            int(row.get("r", "0")): row
            for row in sheet_data
            if row.tag == _q("row")
        }
        stacks: dict[str, dict[str, Any]] = {}
        regions: list[dict[str, Any]] = []

        for row_number, row in rows.items():
            for cell in row.findall("x:c", NS):
                text = self._cell_text(cell, shared_strings)
                match = _BLOCK_MARKER_RE.match(text or "")
                if not match:
                    continue
                range_name = match.group(1).upper()
                start_name, end_name = range_name.split(":")
                start_col, end_col = _column_number(start_name), _column_number(end_name)
                cell_ref = _CELL_REF_RE.match(cell.get("r", ""))
                if start_col > end_col or not cell_ref:
                    raise UnsupportedFeatureError(
                        f"Invalid region range {range_name}."
                    )
                marker_col = _column_number(cell_ref.group(1))
                if not start_col <= marker_col <= end_col:
                    raise UnsupportedFeatureError(
                        f"Region marker {cell.get('r')} must be inside {range_name}."
                    )
                statement = match.group(2).strip()
                for_match = _FOR_STATEMENT_RE.match(statement)
                if for_match:
                    if range_name in stacks:
                        raise UnsupportedFeatureError(
                            f"Nested region {range_name} is not supported."
                        )
                    stacks[range_name] = {
                        "range": range_name,
                        "start_col": start_col,
                        "end_col": end_col,
                        "open_row": row_number,
                        "target": for_match.group(1).strip(),
                        "expression": for_match.group(2).strip(),
                    }
                elif statement.lower() == "endfor":
                    region = stacks.pop(range_name, None)
                    if region is None:
                        raise UnsupportedFeatureError(
                            f"Region {range_name} has endfor without for."
                        )
                    region["close_row"] = row_number
                    regions.append(region)
                else:
                    raise UnsupportedFeatureError(
                        "Region scope currently supports for/endfor only."
                    )

        if stacks:
            raise UnsupportedFeatureError(
                "Unclosed region(s): " + ", ".join(sorted(stacks))
            )
        if not regions:
            raise UnsupportedFeatureError("Region marker could not be parsed.")

        regions.sort(key=lambda item: (item["open_row"], item["start_col"]))
        for index, region in enumerate(regions):
            if region["close_row"] <= region["open_row"] + 1:
                raise UnsupportedFeatureError(
                    f"Region {region['range']} requires at least one body row."
                )
            region["id"] = f"b{index}"
            region["body_height"] = region["close_row"] - region["open_row"] - 1
            try:
                region["items"] = list(
                    env.compile_expression(region["expression"])(**context)
                )
            except Exception as exc:
                raise XlsxJinjaError(
                    f"Cannot evaluate region {region['range']}: {exc}"
                ) from exc
            region["output_height"] = len(region["items"]) * region["body_height"]
            region["delta"] = region["output_height"] - (
                region["close_row"] - region["open_row"] + 1
            )

        for index, left in enumerate(regions):
            for right in regions[index + 1:]:
                columns_overlap = not (
                    left["end_col"] < right["start_col"]
                    or right["end_col"] < left["start_col"]
                )
                rows_overlap = not (
                    left["close_row"] < right["open_row"]
                    or right["close_row"] < left["open_row"]
                )
                if columns_overlap and rows_overlap:
                    raise UnsupportedFeatureError(
                        f"Regions {left['range']} and {right['range']} overlap."
                    )
                if columns_overlap and left["range"] != right["range"]:
                    raise UnsupportedFeatureError(
                        "Partially overlapping region columns are not supported."
                    )

        for region in regions:
            region["target_start"] = region["open_row"] + sum(
                previous["delta"]
                for previous in regions
                if previous["range"] == region["range"]
                and previous["close_row"] < region["open_row"]
            )
            region_values = context.setdefault("__xlsx_region_values", {})
            region_loops = context.setdefault("__xlsx_region_loops", {})
            region_values[region["id"]] = region["items"]
            length = len(region["items"])
            region_loops[region["id"]] = [
                {
                    "index": item_index + 1,
                    "index0": item_index,
                    "first": item_index == 0,
                    "last": item_index == length - 1,
                    "length": length,
                    "revindex": length - item_index,
                    "revindex0": length - item_index - 1,
                }
                for item_index in range(length)
            ]

        def covering_region(row_number: int, column: int) -> dict[str, Any] | None:
            return next(
                (
                    region
                    for region in regions
                    if region["open_row"] <= row_number <= region["close_row"]
                    and region["start_col"] <= column <= region["end_col"]
                ),
                None,
            )

        def shift_for(row_number: int, column: int) -> int:
            return sum(
                region["delta"]
                for region in regions
                if region["close_row"] < row_number
                and region["start_col"] <= column <= region["end_col"]
            )

        cell_map: dict[tuple[int, int], ET.Element] = {}
        row_attributes: dict[int, dict[str, str]] = {}
        retained_rows: set[int] = set()

        def add_cell(cell: ET.Element, row_number: int, column: int) -> None:
            if not 1 <= row_number <= MAX_XLSX_ROW:
                raise UnsupportedFeatureError(
                    "Region output exceeds Excel's 1,048,576-row limit."
                )
            key = (row_number, column)
            if key in cell_map:
                raise UnsupportedFeatureError(
                    f"Region output collides at {_column_name(column)}{row_number}."
                )
            cell.set("r", f"{_column_name(column)}{row_number}")
            cell_map[key] = cell

        for row_number, row in rows.items():
            attrs = dict(row.attrib)
            attrs.pop("r", None)
            attrs.pop("spans", None)
            has_retained_cell = False
            for source_cell in row.findall("x:c", NS):
                ref = _CELL_REF_RE.match(source_cell.get("r", ""))
                if not ref:
                    continue
                column = _column_number(ref.group(1))
                if covering_region(row_number, column):
                    continue
                target_row = row_number + shift_for(row_number, column)
                cell = copy.deepcopy(source_cell)
                self._shift_formula_rows(cell, target_row - row_number)
                add_cell(cell, target_row, column)
                has_retained_cell = True
            if has_retained_cell or not any(
                region["open_row"] <= row_number <= region["close_row"]
                for region in regions
            ):
                retained_rows.add(row_number)
                row_attributes.setdefault(row_number, attrs)

        for region in regions:
            body_start = region["open_row"] + 1
            loop_name = f"__xlsx_loop_{region['id']}"
            for item_index, _item in enumerate(region["items"]):
                iteration_start = (
                    region["target_start"] + item_index * region["body_height"]
                )
                before = (
                    "{% set "
                    + region["target"]
                    + f" = __xlsx_region_values['{region['id']}'][{item_index}] %}}"
                    + "{% set "
                    + loop_name
                    + f" = __xlsx_region_loops['{region['id']}'][{item_index}] %}}"
                )
                for offset in range(region["body_height"]):
                    source_row_number = body_start + offset
                    source_row = rows.get(source_row_number)
                    target_row = iteration_start + offset
                    retained_rows.add(target_row)
                    if source_row is not None:
                        attrs = dict(source_row.attrib)
                        attrs.pop("r", None)
                        attrs.pop("spans", None)
                        existing = row_attributes.setdefault(target_row, attrs)
                        if attrs.get("ht") and float(attrs["ht"]) > float(
                            existing.get("ht", "0")
                        ):
                            existing.update(ht=attrs["ht"], customHeight="1")
                        for source_cell in source_row.findall("x:c", NS):
                            ref = _CELL_REF_RE.match(source_cell.get("r", ""))
                            if not ref:
                                continue
                            column = _column_number(ref.group(1))
                            if not region["start_col"] <= column <= region["end_col"]:
                                continue
                            cell = copy.deepcopy(source_cell)
                            text = self._cell_text(cell, shared_strings)
                            if text is not None and _JINJA_RE.search(text):
                                text = re.sub(
                                    r"\bloop\.", f"{loop_name}.", text
                                )
                                self._set_inline_text(cell, text)
                            self._shift_formula_rows(
                                cell, target_row - source_row_number
                            )
                            add_cell(cell, target_row, column)
                    if offset == 0:
                        attrs = row_attributes.setdefault(target_row, {})
                        attrs["data-xlsx-jinja-before"] = (
                            attrs.get("data-xlsx-jinja-before", "") + before
                        )

        cells_by_row: dict[int, list[tuple[int, ET.Element]]] = {}
        for (row_number, column), cell in cell_map.items():
            cells_by_row.setdefault(row_number, []).append((column, cell))
        output_rows: list[ET.Element] = []
        for row_number in sorted(retained_rows | cells_by_row.keys()):
            row = ET.Element(_q("row"), {"r": str(row_number)})
            row.attrib.update(row_attributes.get(row_number, {}))
            for _column, cell in sorted(cells_by_row.get(row_number, [])):
                row.append(cell)
            output_rows.append(row)
        sheet_data[:] = output_rows

        self._expand_region_merges(root, regions)
        self._expand_region_drawings(sheet_path, root, regions)
        self._update_dimension(root, output_rows)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    @staticmethod
    def _shift_formula_rows(cell: ET.Element, delta: int) -> None:
        if not delta:
            return
        formula = cell.find("x:f", NS)
        if formula is None or not formula.text or "!" in formula.text:
            return

        def replace(match: re.Match[str]) -> str:
            def shifted(row: str, absolute: str) -> int:
                return int(row) if absolute else max(1, int(row) + delta)

            result = (
                f"{match.group('c1')}{match.group('ra1')}"
                f"{shifted(match.group('r1'), match.group('ra1'))}"
            )
            if match.group("r2") is not None:
                result += (
                    f":{match.group('c2')}{match.group('ra2')}"
                    f"{shifted(match.group('r2'), match.group('ra2'))}"
                )
            return result

        parts = re.split(r'("(?:[^"]|"")*")', formula.text)
        formula.text = "".join(
            part if index % 2 else _FORMULA_REF_RE.sub(replace, part)
            for index, part in enumerate(parts)
        )

    def _expand_region_merges(
        self, root: ET.Element, regions: list[dict[str, Any]]
    ) -> None:
        merges = root.find("x:mergeCells", NS)
        if merges is None:
            return
        output: list[str] = []
        for merge in list(merges):
            match = _RANGE_RE.match(merge.get("ref", ""))
            if not match:
                output.append(merge.get("ref", ""))
                continue
            start_col = _column_number(match.group("c1"))
            end_col = _column_number(match.group("c2"))
            start_row, end_row = int(match.group("r1")), int(match.group("r2"))
            body_region = next(
                (
                    region
                    for region in regions
                    if region["start_col"] <= start_col <= end_col <= region["end_col"]
                    and region["open_row"] < start_row <= end_row < region["close_row"]
                ),
                None,
            )
            if body_region:
                for item_index in range(len(body_region["items"])):
                    delta = (
                        body_region["target_start"]
                        + item_index * body_region["body_height"]
                        - body_region["open_row"]
                        - 1
                    )
                    output.append(
                        f"{match.group('c1')}{start_row + delta}:"
                        f"{match.group('c2')}{end_row + delta}"
                    )
                continue

            shifts = {
                sum(
                    region["delta"]
                    for region in regions
                    if region["close_row"] < row
                    and region["start_col"] <= column <= region["end_col"]
                )
                for row, column in (
                    (start_row, start_col),
                    (end_row, end_col),
                )
            }
            intersects = any(
                not (end_col < region["start_col"] or start_col > region["end_col"])
                and not (end_row < region["open_row"] or start_row > region["close_row"])
                for region in regions
            )
            if intersects:
                contained_marker = any(
                    region["start_col"] <= start_col <= end_col <= region["end_col"]
                    and region["open_row"] <= start_row <= end_row <= region["close_row"]
                    for region in regions
                )
                if contained_marker:
                    continue
                raise UnsupportedFeatureError(
                    f"Merged range {merge.get('ref')} crosses a region boundary."
                )
            if len(shifts) != 1:
                raise UnsupportedFeatureError(
                    f"Merged range {merge.get('ref')} crosses independently shifted regions."
                )
            shift = shifts.pop()
            output.append(
                f"{match.group('c1')}{start_row + shift}:"
                f"{match.group('c2')}{end_row + shift}"
            )

        merges[:] = []
        for ref in output:
            ET.SubElement(merges, _q("mergeCell"), {"ref": ref})
        merges.set("count", str(len(output)))

    def _expand_region_drawings(
        self,
        sheet_path: str,
        root: ET.Element,
        regions: list[dict[str, Any]],
    ) -> None:
        parts = self._drawing_parts(sheet_path, root, create=False)
        if parts is None:
            return
        drawing_path, drawing_rel_path, drawing_root, drawing_rels = parts
        output: list[ET.Element] = []
        for source_anchor in list(drawing_root):
            position = self._anchor_position(source_anchor)
            if position is None:
                output.append(source_anchor)
                continue
            start_row, start_col = position
            end_marker = source_anchor.find("xdr:to", NS)
            end_row = (
                int(end_marker.findtext("xdr:row", namespaces=NS)) + 1
                if end_marker is not None
                else start_row
            )
            end_col = (
                int(end_marker.findtext("xdr:col", namespaces=NS)) + 1
                if end_marker is not None
                else start_col
            )
            body_region = next(
                (
                    region
                    for region in regions
                    if region["start_col"] <= start_col <= end_col <= region["end_col"]
                    and region["open_row"] < start_row <= end_row < region["close_row"]
                ),
                None,
            )
            if body_region:
                for item_index in range(len(body_region["items"])):
                    anchor = copy.deepcopy(source_anchor)
                    target_row = (
                        body_region["target_start"]
                        + item_index * body_region["body_height"]
                        + start_row
                        - body_region["open_row"]
                        - 1
                    )
                    self._shift_anchor(anchor, target_row - start_row)
                    output.append(anchor)
                continue
            crossing_region = next(
                (
                    region
                    for region in regions
                    if not (end_col < region["start_col"] or start_col > region["end_col"])
                    and not (end_row < region["open_row"] or start_row > region["close_row"])
                ),
                None,
            )
            if crossing_region and not (
                crossing_region["start_col"] <= start_col <= end_col <= crossing_region["end_col"]
            ):
                raise UnsupportedFeatureError(
                    "Drawing crosses a region column boundary."
                )
            marker_region = next(
                (
                    region
                    for region in regions
                    if region["start_col"] <= start_col <= end_col <= region["end_col"]
                    and region["open_row"] <= start_row <= end_row <= region["close_row"]
                ),
                None,
            )
            if marker_region:
                continue
            start_shift = sum(
                region["delta"]
                for region in regions
                if region["close_row"] < start_row
                and region["start_col"] <= start_col <= region["end_col"]
            )
            end_shift = sum(
                region["delta"]
                for region in regions
                if region["close_row"] < end_row
                and region["start_col"] <= end_col <= region["end_col"]
            )
            if start_shift != end_shift:
                raise UnsupportedFeatureError(
                    "Drawing crosses independently shifted regions."
                )
            shift = start_shift
            anchor = copy.deepcopy(source_anchor)
            self._shift_anchor(anchor, shift)
            output.append(anchor)

        drawing_root[:] = output
        for number, anchor in enumerate(output, 1):
            properties = anchor.find(".//xdr:cNvPr", NS)
            if properties is not None:
                properties.set("id", str(number))
                properties.set("name", f"Picture {number}")
        self._entries[drawing_path] = ET.tostring(
            drawing_root, encoding="utf-8", xml_declaration=True
        )
        self._entries[drawing_rel_path] = ET.tostring(
            drawing_rels, encoding="utf-8", xml_declaration=True
        )

    def _normalize_template_cells(self, row: ET.Element, shared_strings: list[str]) -> None:
        for cell in row.findall("x:c", NS):
            text = self._cell_text(cell, shared_strings)
            if text is None or not _JINJA_RE.search(text):
                continue
            self._set_inline_text(cell, text)

    @staticmethod
    def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str | None:
        cell_type = cell.get("t")
        if cell_type == "s":
            value = cell.findtext("x:v", default="", namespaces=NS)
            try:
                return shared_strings[int(value)]
            except (ValueError, IndexError):
                return None
        if cell_type == "inlineStr":
            return "".join(node.text or "" for node in cell.iter(_q("t")))
        if cell_type == "str":
            return cell.findtext("x:v", namespaces=NS)
        return None

    @staticmethod
    def _set_inline_text(cell: ET.Element, text: str) -> None:
        for child in list(cell):
            if child.tag != _q("f"):
                cell.remove(child)
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, _q("is"))
        text_node = ET.SubElement(inline, _q("t"))
        if text[:1].isspace() or text[-1:].isspace():
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_node.text = text

    def _structural_row_tag(self, row: ET.Element) -> str | None:
        for cell in row.findall("x:c", NS):
            text = self._cell_text(cell, [])
            if text is None:
                continue
            match = _STRUCTURAL_ROW_RE.match(text)
            if match:
                return match.group(1)
        return None

    @staticmethod
    def _control_closes_in_cell(text: str, opening: str) -> bool:
        closing = "endfor" if opening == "for" else "endif"
        depth = 0
        for match in _CONTROL_TAG_RE.finditer(text):
            tag = match.group(1).lower()
            if tag == opening:
                depth += 1
            elif tag == closing:
                depth -= 1
                if depth == 0:
                    return True
        return False

    def _hoist_row_controls(
        self, row: ET.Element
    ) -> tuple[list[str], list[str], bool]:
        before: list[str] = []
        after: list[str] = []
        row_openings: list[str] = []
        row_scoped = False
        for cell in row.findall("x:c", NS):
            text = self._cell_text(cell, [])
            if text is None or not _JINJA_RE.search(text):
                continue
            had_tr = bool(_TR_RE.search(text))
            text = _TR_RE.sub("", text)

            while True:
                set_match = _LEADING_SET_RE.match(text)
                match = set_match or _LEADING_CONTROL_RE.match(text)
                if not match:
                    break
                statement = match.group(1).strip()
                opening = statement.split(None, 1)[0].lower()
                if (
                    not had_tr
                    and opening in {"for", "if"}
                    and self._control_closes_in_cell(text, opening)
                ):
                    break
                tag = "{% " + statement + " %}"
                if opening in {"for", "if"}:
                    before.append(tag)
                    row_openings.append("endfor" if opening == "for" else "endif")
                    row_scoped = True
                elif (
                    opening in {"endfor", "endif"}
                    and row_openings
                    and row_openings[-1] == opening
                ):
                    row_openings.pop()
                    after.append(tag)
                else:
                    before.append(tag)
                text = text[match.end():]

            trailing: list[tuple[str, int]] = []
            remaining = text
            while True:
                match = _TRAILING_END_RE.search(remaining)
                if not match:
                    break
                statement = match.group(1).strip()
                trailing.insert(0, (statement, match.start()))
                remaining = remaining[:match.start()]
            if trailing and len(trailing) <= len(row_openings):
                expected = list(reversed(row_openings[-len(trailing):]))
                endings = [statement.split(None, 1)[0].lower() for statement, _ in trailing]
                inline_opening = any(
                    self._control_closes_in_cell(text, "for" if ending == "endfor" else "if")
                    for ending in endings
                )
                if endings == expected and (had_tr or not inline_opening):
                    del row_openings[-len(trailing):]
                    after.extend("{% " + statement + " %}" for statement, _ in trailing)
                    text = remaining

            self._set_inline_text(cell, text)
        return before, after, row_scoped

    @staticmethod
    def _renumber_rows(
        rows: list[ET.Element],
        removed_rows: set[int],
        compact_rows: set[int],
    ) -> dict[int, list[int]]:
        row_map: dict[int, list[int]] = {}
        valid_rows: list[ET.Element] = []
        previous_old: int | None = None
        previous_new: int | None = None

        for row in rows:
            old = int(row.get("r", "0") or 0)
            if previous_old is None:
                new = old - sum(1 for removed in removed_rows if removed <= old)
                new = max(1, new)
            elif old == previous_old:
                new = previous_new + 1
            else:
                skipped_between = {
                    removed
                    for removed in removed_rows
                    if previous_old < removed <= old
                }
                skipped_between.update(
                    compact
                    for compact in compact_rows
                    if previous_old < compact < old
                )
                new = previous_new + max(
                    1, old - previous_old - len(skipped_between)
                )

            if new > MAX_XLSX_ROW:
                if any(
                    cell.find("x:f", NS) is not None
                    or cell.find("x:v", NS) is not None
                    or any((text.text or "") for text in cell.iter(_q("t")))
                    for cell in row.findall("x:c", NS)
                ):
                    raise UnsupportedFeatureError(
                        "Rendered worksheet exceeds Excel's 1,048,576-row limit."
                    )
                continue

            row.set("data-xlsx-jinja-old-r", str(old))
            row.set("r", str(new))
            for cell in row.findall("x:c", NS):
                ref = cell.get("r", "")
                match = _CELL_REF_RE.match(ref)
                if match:
                    cell.set("r", f"{match.group(1).replace('$', '')}{new}")
            row_map.setdefault(old, []).append(new)
            valid_rows.append(row)
            previous_old, previous_new = old, new
        rows[:] = valid_rows
        return row_map

    @staticmethod
    def _apply_native_values(
        row: ET.Element,
        native_values: list[Any],
        richtext_values: list[RichText],
    ) -> None:
        for cell in row.findall("x:c", NS):
            if cell.get("t") != "inlineStr":
                continue
            full_text = "".join(node.text or "" for node in cell.iter(_q("t")))
            if apply_richtext_cell(cell, full_text, richtext_values):
                continue
            text = full_text.strip()
            match = _NATIVE_TOKEN_RE.match(text)
            if not match:
                continue
            value = native_values[int(match.group(1))]
            for child in list(cell):
                if child.tag != _q("f"):
                    cell.remove(child)
            value_node = ET.SubElement(cell, _q("v"))
            if value is None:
                cell.attrib.pop("t", None)
                value_node.text = ""
            elif isinstance(value, bool):
                cell.set("t", "b")
                value_node.text = "1" if value else "0"
            elif isinstance(value, (int, float)):
                cell.set("t", "n")
                value_node.text = str(value)
            elif isinstance(value, (dt.date, dt.datetime)):
                cell.set("t", "d")
                value_node.text = value.isoformat()
            else:
                cell.set("t", "inlineStr")
                cell.remove(value_node)
                inline = ET.SubElement(cell, _q("is"))
                ET.SubElement(inline, _q("t")).text = str(value)

    def _collect_image_values(
        self,
        row: ET.Element,
        old_row: int,
        image_values: list[tuple[str, Any, int]],
        requests: list[dict[str, Any]],
    ) -> None:
        for cell in row.findall("x:c", NS):
            if cell.get("t") != "inlineStr":
                continue
            text = "".join(node.text or "" for node in cell.iter(_q("t")))
            matches = list(_IMAGE_TOKEN_RE.finditer(text))
            if not matches:
                continue
            ref = _CELL_REF_RE.match(cell.get("r", ""))
            if not ref:
                continue
            for match in matches:
                mode, value, index = image_values[int(match.group(1))]
                requests.append(
                    {
                        "mode": mode,
                        "value": value,
                        "index": index,
                        "old_row": old_row,
                        "new_row": int(ref.group(2)),
                        "old_col": _column_number(ref.group(1)),
                        "new_col": _column_number(ref.group(1)),
                    }
                )
            self._set_inline_text(cell, _IMAGE_TOKEN_RE.sub("", text))

    @staticmethod
    def _relationship_path(part_path: str) -> str:
        return posixpath.join(
            posixpath.dirname(part_path),
            "_rels",
            posixpath.basename(part_path) + ".rels",
        )

    @staticmethod
    def _next_relationship_id(root: ET.Element) -> str:
        used = {rel.get("Id", "") for rel in root}
        index = 1
        while f"rId{index}" in used:
            index += 1
        return f"rId{index}"

    def _next_part_path(self, pattern: str, template: str) -> str:
        numbers = [
            int(match.group(1))
            for path in self._entries
            if (match := re.match(pattern, path))
        ]
        return template.format(max(numbers, default=0) + 1)

    def _ensure_content_type(
        self, *, extension: str | None = None, part: str | None = None
    ) -> None:
        path = "[Content_Types].xml"
        root = ET.fromstring(self._entries[path])
        if extension and not any(
            node.get("Extension", "").lower() == extension.lower() for node in root
        ):
            ET.SubElement(
                root,
                _qn(CONTENT_TYPES_NS, "Default"),
                {"Extension": extension, "ContentType": f"image/{extension}"},
            )
        if part and not any(node.get("PartName") == f"/{part}" for node in root):
            ET.SubElement(
                root,
                _qn(CONTENT_TYPES_NS, "Override"),
                {
                    "PartName": f"/{part}",
                    "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
                },
            )
        self._entries[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _add_image_part(self, drawing_rels: ET.Element, value: Any) -> str | None:
        data = _png_bytes(value)
        if data is None:
            return None
        media_path = self._next_part_path(
            r"xl/media/image(\d+)\.[^.]+$", "xl/media/image{}.png"
        )
        self._entries[media_path] = data
        self._ensure_content_type(extension="png")
        relationship_id = self._next_relationship_id(drawing_rels)
        ET.SubElement(
            drawing_rels,
            _qn(PACKAGE_REL_NS, "Relationship"),
            {
                "Id": relationship_id,
                "Type": IMAGE_REL,
                "Target": posixpath.relpath(media_path, "xl/drawings"),
            },
        )
        return relationship_id

    def _drawing_parts(
        self, sheet_path: str, sheet_root: ET.Element, create: bool
    ) -> tuple[str, str, ET.Element, ET.Element] | None:
        sheet_rel_path = self._relationship_path(sheet_path)
        sheet_rels = ET.fromstring(
            self._entries.get(
                sheet_rel_path,
                f'<Relationships xmlns="{PACKAGE_REL_NS}"/>'.encode(),
            )
        )
        drawing_node = sheet_root.find("x:drawing", NS)
        drawing_path = None
        if drawing_node is not None:
            drawing_id = drawing_node.get(_qn(REL_NS, "id"), "")
            for relationship in sheet_rels:
                if relationship.get("Id") == drawing_id:
                    drawing_path = self._resolve_part(
                        sheet_path, relationship.get("Target", "")
                    )
                    break
        if drawing_path is None and not create:
            return None
        if drawing_path is None:
            drawing_path = self._next_part_path(
                r"xl/drawings/drawing(\d+)\.xml$", "xl/drawings/drawing{}.xml"
            )
            drawing_id = self._next_relationship_id(sheet_rels)
            ET.SubElement(
                sheet_rels,
                _qn(PACKAGE_REL_NS, "Relationship"),
                {
                    "Id": drawing_id,
                    "Type": DRAWING_REL,
                    "Target": posixpath.relpath(
                        drawing_path, posixpath.dirname(sheet_path)
                    ),
                },
            )
            drawing_node = ET.Element(_q("drawing"), {_qn(REL_NS, "id"): drawing_id})
            trailing = {
                "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects",
                "controls", "webPublishItems", "tableParts", "extLst",
            }
            insert_at = next(
                (
                    index
                    for index, child in enumerate(sheet_root)
                    if child.tag.rsplit("}", 1)[-1] in trailing
                ),
                len(sheet_root),
            )
            sheet_root.insert(insert_at, drawing_node)
            self._ensure_content_type(part=drawing_path)
        self._entries[sheet_rel_path] = ET.tostring(
            sheet_rels, encoding="utf-8", xml_declaration=True
        )
        drawing_rel_path = self._relationship_path(drawing_path)
        drawing_root = ET.fromstring(
            self._entries.get(
                drawing_path, f'<xdr:wsDr xmlns:xdr="{DRAWING_NS}"/>'.encode()
            )
        )
        drawing_rels = ET.fromstring(
            self._entries.get(
                drawing_rel_path,
                f'<Relationships xmlns="{PACKAGE_REL_NS}"/>'.encode(),
            )
        )
        return drawing_path, drawing_rel_path, drawing_root, drawing_rels

    @staticmethod
    def _anchor_position(anchor: ET.Element) -> tuple[int, int] | None:
        marker = anchor.find("xdr:from", NS)
        if marker is None:
            return None
        row = marker.findtext("xdr:row", namespaces=NS)
        column = marker.findtext("xdr:col", namespaces=NS)
        if row is None or column is None:
            return None
        return int(row) + 1, int(column) + 1

    @staticmethod
    def _shift_anchor(anchor: ET.Element, row_delta: int) -> None:
        for marker_name in ("from", "to"):
            marker = anchor.find(f"xdr:{marker_name}", NS)
            row = marker.find("xdr:row", NS) if marker is not None else None
            if row is not None:
                row.text = str(max(0, int(row.text or 0) + row_delta))

    @staticmethod
    def _set_anchor_image(anchor: ET.Element, relationship_id: str) -> None:
        blip = anchor.find(".//a:blip", NS)
        if blip is None:
            raise UnsupportedFeatureError("Image placeholder has no embedded image.")
        blip.set(_qn(REL_NS, "embed"), relationship_id)

    @staticmethod
    def _cell_size(root: ET.Element, row: int, column: int) -> tuple[int, int]:
        sheet_format = root.find("x:sheetFormatPr", NS)
        width = float(sheet_format.get("defaultColWidth", "8.43")) if sheet_format is not None else 8.43
        height = float(sheet_format.get("defaultRowHeight", "15")) if sheet_format is not None else 15
        for col in root.findall("x:cols/x:col", NS):
            if int(col.get("min", "0")) <= column <= int(col.get("max", "0")):
                width = float(col.get("width", width))
                break
        row_node = root.find(f"x:sheetData/x:row[@r='{row}']", NS)
        if row_node is not None and row_node.get("ht"):
            height = float(row_node.get("ht"))
        width_px = max(math.floor(((width * 256 + 128) / 256) * 7), 1)
        height_px = max(math.ceil(height * 96 / 72), 1)
        return width_px * 9525, height_px * 9525

    def _new_image_anchor(
        self,
        sheet_root: ET.Element,
        request: dict[str, Any],
        relationship_id: str,
    ) -> ET.Element:
        anchor = ET.Element(_qn(DRAWING_NS, "oneCellAnchor"))
        marker = ET.SubElement(anchor, _qn(DRAWING_NS, "from"))
        for name, value in (
            ("col", request["new_col"] - 1), ("colOff", 0),
            ("row", request["new_row"] - 1), ("rowOff", 0),
        ):
            ET.SubElement(marker, _qn(DRAWING_NS, name)).text = str(value)
        width, height = self._cell_size(
            sheet_root, request["new_row"], request["new_col"]
        )
        ET.SubElement(
            anchor, _qn(DRAWING_NS, "ext"), {"cx": str(width), "cy": str(height)}
        )
        picture = ET.SubElement(anchor, _qn(DRAWING_NS, "pic"))
        properties = ET.SubElement(picture, _qn(DRAWING_NS, "nvPicPr"))
        ET.SubElement(properties, _qn(DRAWING_NS, "cNvPr"), {"id": "0", "name": "Picture"})
        lock = ET.SubElement(properties, _qn(DRAWING_NS, "cNvPicPr"))
        ET.SubElement(lock, _qn(DRAWING_MAIN_NS, "picLocks"), {"noChangeAspect": "1"})
        fill = ET.SubElement(picture, _qn(DRAWING_NS, "blipFill"))
        ET.SubElement(
            fill, _qn(DRAWING_MAIN_NS, "blip"), {_qn(REL_NS, "embed"): relationship_id}
        )
        stretch = ET.SubElement(fill, _qn(DRAWING_MAIN_NS, "stretch"))
        ET.SubElement(stretch, _qn(DRAWING_MAIN_NS, "fillRect"))
        shape = ET.SubElement(picture, _qn(DRAWING_NS, "spPr"))
        geometry = ET.SubElement(shape, _qn(DRAWING_MAIN_NS, "prstGeom"), {"prst": "rect"})
        ET.SubElement(geometry, _qn(DRAWING_MAIN_NS, "avLst"))
        ET.SubElement(anchor, _qn(DRAWING_NS, "clientData"))
        return anchor

    def _render_images(
        self,
        sheet_path: str,
        sheet_root: ET.Element,
        row_map: dict[int, list[int]],
        requests: list[dict[str, Any]],
    ) -> None:
        parts = self._drawing_parts(
            sheet_path,
            sheet_root,
            create=any(request["mode"] == "insert_img" and request["value"] for request in requests),
        )
        if parts is None:
            if any(request["mode"] == "img" for request in requests):
                raise UnsupportedFeatureError(
                    "{% img %} requires a placeholder image anchored to the same cell."
                )
            return
        drawing_path, drawing_rel_path, drawing_root, drawing_rels = parts
        source_anchors = list(drawing_root)
        anchor_indexes: dict[tuple[int, int], int] = {}
        output_anchors: list[ET.Element] = []

        for source_anchor in source_anchors:
            position = self._anchor_position(source_anchor)
            if position is None:
                output_anchors.append(source_anchor)
                continue
            old_row, old_col = position
            is_picture = source_anchor.find(".//a:blip", NS) is not None
            image_index = anchor_indexes.get(position, 0) if is_picture else None
            if is_picture:
                anchor_indexes[position] = image_index + 1
            for new_row in row_map.get(old_row, [old_row]):
                anchor = copy.deepcopy(source_anchor)
                self._shift_anchor(anchor, new_row - old_row)
                request = next(
                    (
                        item for item in requests
                        if item["mode"] == "img"
                        and item["old_row"] == old_row
                        and item["old_col"] == old_col
                        and item["new_row"] == new_row
                        and image_index is not None
                        and item["index"] == image_index
                    ),
                    None,
                )
                if request is not None:
                    request["handled"] = True
                    relationship_id = self._add_image_part(
                        drawing_rels, request["value"]
                    )
                    if relationship_id:
                        self._set_anchor_image(anchor, relationship_id)
                output_anchors.append(anchor)

        for request in requests:
            if request["mode"] == "img" and not request.get("handled"):
                raise UnsupportedFeatureError(
                    "{% img %} requires a placeholder image anchored to the same cell and index."
                )
            if request["mode"] != "insert_img" or not request["value"]:
                continue
            relationship_id = self._add_image_part(drawing_rels, request["value"])
            if relationship_id:
                output_anchors.append(
                    self._new_image_anchor(sheet_root, request, relationship_id)
                )

        drawing_root[:] = output_anchors
        for number, anchor in enumerate(output_anchors, 1):
            properties = anchor.find(".//xdr:cNvPr", NS)
            if properties is not None:
                properties.set("id", str(number))
                properties.set("name", f"Picture {number}")
        self._entries[drawing_path] = ET.tostring(
            drawing_root, encoding="utf-8", xml_declaration=True
        )
        self._entries[drawing_rel_path] = ET.tostring(
            drawing_rels, encoding="utf-8", xml_declaration=True
        )

    def _translate_formulas(
        self,
        row: ET.Element,
        old_row: int,
        new_row: int,
        row_map: dict[int, list[int]],
    ) -> None:
        for formula in row.findall("x:c/x:f", NS):
            if (
                not formula.text
                or formula.get("t") in {"array", "dataTable", "shared"}
                or any(marker in formula.text for marker in ("!", "[", "]", "#"))
            ):
                continue

            def replace(match: re.Match[str]) -> str:
                row1 = int(match.group("r1"))
                row2_text = match.group("r2")
                if row2_text is not None:
                    row2 = int(row2_text)
                    if row1 == row2 == old_row:
                        start = end = new_row
                    else:
                        start = row_map.get(row1, [row1])[0]
                        end = row_map.get(row2, [row2])[-1]
                    return (
                        f"{match.group('c1')}{match.group('ra1')}{start}:"
                        f"{match.group('c2')}{match.group('ra2')}{end}"
                    )

                if row1 == old_row and not match.group("ra1"):
                    target = new_row
                elif row1 in row_map:
                    target = row_map[row1][0]
                elif not match.group("ra1"):
                    target = max(1, row1 + new_row - old_row)
                else:
                    target = row1
                return f"{match.group('c1')}{match.group('ra1')}{target}"

            parts = re.split(r'("(?:[^"]|"")*")', formula.text)
            formula.text = "".join(
                part if index % 2 else _FORMULA_REF_RE.sub(replace, part)
                for index, part in enumerate(parts)
            )

    @staticmethod
    def _shift_merges(root: ET.Element, row_map: dict[int, list[int]]) -> None:
        merges = root.find("x:mergeCells", NS)
        if merges is None:
            return
        refs: list[str] = []
        for merge in list(merges):
            match = _RANGE_RE.match(merge.get("ref", ""))
            if not match:
                refs.append(merge.get("ref", ""))
                continue
            start_row = int(match.group("r1"))
            end_row = int(match.group("r2"))
            if start_row == end_row and len(row_map.get(start_row, [])) > 1:
                for mapped in row_map[start_row]:
                    refs.append(f"{match.group('c1')}{mapped}:{match.group('c2')}{mapped}")
            else:
                mapped_start = row_map.get(start_row, [start_row])[0]
                mapped_end = row_map.get(end_row, [end_row])[-1]
                refs.append(
                    f"{match.group('c1')}{mapped_start}:{match.group('c2')}{mapped_end}"
                )
        merges[:] = []
        for ref in refs:
            ET.SubElement(merges, _q("mergeCell"), {"ref": ref})
        merges.set("count", str(len(refs)))

    @staticmethod
    def _update_dimension(root: ET.Element, rows: list[ET.Element]) -> None:
        dimension = root.find("x:dimension", NS)
        refs = [
            cell.get("r", "")
            for row in rows
            for cell in row.findall("x:c", NS)
            if _CELL_REF_RE.match(cell.get("r", ""))
        ]
        if dimension is None or not refs:
            return
        parsed = [_CELL_REF_RE.match(ref) for ref in refs]
        min_row = min(int(match.group(2)) for match in parsed)
        max_row = max(int(match.group(2)) for match in parsed)
        min_col = min(parsed, key=lambda match: _column_number(match.group(1))).group(1)
        max_col = max(parsed, key=lambda match: _column_number(match.group(1))).group(1)
        dimension.set(
            "ref",
            f"{min_col.replace('$', '')}{min_row}:{max_col.replace('$', '')}{max_row}",
        )
