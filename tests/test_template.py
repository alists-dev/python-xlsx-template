import base64
import io
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from xlsx_jinja import UnsupportedFeatureError, XlsxTemplate

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
MAIN = NS["x"]
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWING_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def package(sheet_xml, shared_strings=(), extra=None):
    shared = "".join(f"<si><t>{value}</t></si>" for value in shared_strings)
    entries = {
        "[Content_Types].xml": b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        "_rels/.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        "xl/workbook.xml": (
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Report &amp; Data" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ).encode(),
        "xl/_rels/workbook.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/></Relationships>'
        ).encode(),
        "xl/worksheets/sheet1.xml": sheet_xml.encode(),
        "xl/sharedStrings.xml": (
            f'<sst xmlns="{MAIN}" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
            f"{shared}</sst>"
        ).encode(),
        "custom/preserved.bin": b"\x00preserve-me\xff",
    }
    entries.update(extra or {})
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in entries.items():
            archive.writestr(name, value)
    return output.getvalue()


def render(source, context):
    output = io.BytesIO()
    template = XlsxTemplate(source)
    template.render(context)
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        preserved = archive.read("custom/preserved.bin")
    return root, preserved


def cells(root):
    return {cell.get("r"): cell for cell in root.findall(".//x:c", NS)}


def text(cell):
    return "".join(node.text or "" for node in cell.findall(".//x:t", NS))


def value(cell):
    return cell.findtext("x:v", namespaces=NS)


def test_shared_and_inline_strings_are_rendered_and_xml_escaped():
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:B1"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>{{{{ sheet_name }}}}: {{{{ value }}}}</t></is></c></row>
    </sheetData></worksheet>'''
    root, preserved = render(package(sheet, ["{{ value }}"]), {"value": "A&B <C>"})
    result = cells(root)
    assert text(result["A1"]) == "A&B <C>"
    assert text(result["B1"]) == "Report & Data: A&B <C>"
    assert preserved == b"\x00preserve-me\xff"


def test_structural_rows_loop_condition_native_formula_and_merge_shift():
    strings = [
        "{%r for item in items %}",
        "{{ item.name }}",
        "{% xv item.qty %}",
        "{%r endfor %}",
        "{%r if show_note %}",
        "Visible",
        "{%r endif %}",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:F7"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c><c r="C2"><f>IF(B2=&quot;B2&quot;,B2,0)</f><v>0</v></c><c r="D2" t="inlineStr"><is><t>merged</t></is></c><c r="F2"><f>'Other Sheet'!B2</f><v>0</v></c></row>
      <row r="3"><c r="A3" t="s"><v>3</v></c></row>
      <row r="4"><c r="A4" t="s"><v>4</v></c></row>
      <row r="5"><c r="A5" t="s"><v>5</v></c></row>
      <row r="6"><c r="A6" t="s"><v>6</v></c></row>
      <row r="7"><c r="B7"><f>SUM(B2:B2)</f><v>0</v></c></row>
    </sheetData><mergeCells count="1"><mergeCell ref="D2:E2"/></mergeCells></worksheet>'''
    root, _ = render(
        package(sheet, strings),
        {"items": [{"name": "One", "qty": 2}, {"name": "Two", "qty": 3}], "show_note": True},
    )
    result = cells(root)
    assert [row.get("r") for row in root.findall("x:sheetData/x:row", NS)] == ["1", "2", "3", "4"]
    assert text(result["A1"]) == "One"
    assert text(result["A2"]) == "Two"
    assert result["B1"].get("t") == "n" and value(result["B1"]) == "2"
    assert result["B2"].get("t") == "n" and value(result["B2"]) == "3"
    assert result["C1"].findtext("x:f", namespaces=NS) == 'IF(B1="B2",B1,0)'
    assert result["C2"].findtext("x:f", namespaces=NS) == 'IF(B2="B2",B2,0)'
    assert result["F1"].findtext("x:f", namespaces=NS) == "'Other Sheet'!B2"
    assert result["F2"].findtext("x:f", namespaces=NS) == "'Other Sheet'!B2"
    assert result["A3"].get("t") == "s" and value(result["A3"]) == "5"
    assert result["B4"].findtext("x:f", namespaces=NS) == "SUM(B1:B2)"
    assert [merge.get("ref") for merge in root.findall("x:mergeCells/x:mergeCell", NS)] == ["D1:E1", "D2:E2"]
    assert root.find("x:dimension", NS).get("ref") == "A1:F4"


def test_false_structural_condition_removes_body_row():
    strings = ["{%r if shown %}", "secret", "{%r endif %}"]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:A4"/><sheetData>
      <row r="1"><c r="A1" t="inlineStr"><is><t>header</t></is></c></row>
      <row r="2"><c r="A2" t="s"><v>0</v></c></row>
      <row r="3"><c r="A3" t="s"><v>1</v></c></row>
      <row r="4"><c r="A4" t="s"><v>2</v></c></row>
    </sheetData></worksheet>'''
    root, _ = render(package(sheet, strings), {"shown": False})
    result = cells(root)
    assert list(result) == ["A1"]
    assert text(result["A1"]) == "header"


def test_inline_tr_loop_is_backward_compatible():
    strings = [
        "{% tr %}{% for item in items %}{{ item }}",
        "{% tr %}{% endfor %}",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:B1"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    </sheetData></worksheet>'''
    root, _ = render(package(sheet, strings), {"items": ["A", "B"]})
    result = cells(root)
    assert text(result["A1"]) == "A"
    assert text(result["A2"]) == "B"
    assert text(result["B1"]) == ""
    assert text(result["B2"]) == ""


def test_cross_row_for_keeps_complete_inline_if_blocks_inside_cells():
    strings = [
        "{% set ns = namespace(counter=0) %}{% for line in lines %}",
        "{% if line.product %}{% set ns.counter = ns.counter + 1 %}{{ ns.counter }}{% endif %}",
        "{% if line.product %}{{ line.name }}{% else %}{{ line.name }}{% endif %}",
        "{% endfor %}",
        "TOTAL",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:C2"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
      <row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2" t="s"><v>4</v></c></row>
    </sheetData></worksheet>'''
    lines = [
        type("Line", (), {"product": True, "name": "Product"})(),
        type("Line", (), {"product": False, "name": "Section"})(),
    ]
    root, _ = render(package(sheet, strings), {"lines": lines})
    result = cells(root)
    assert text(result["B1"]) == "1"
    assert text(result["C1"]) == "Product"
    assert text(result["B2"]) == ""
    assert text(result["C2"]) == "Section"
    assert result["B3"].get("t") == "s" and value(result["B3"]) == "4"


def test_cross_cell_if_is_merged_at_row_scope():
    strings = [
        "{%r for line in lines %}",
        "{% if line.product %}{{ line.name }}",
        "{{ line.qty }}{% endif %}",
        "{% if not line.product %}",
        "{{ line.name }}",
        "{% endif %}",
        "{%r endfor %}",
        "FOOTER",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:C5"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c></row>
      <row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3" t="s"><v>4</v></c><c r="C3" t="s"><v>5</v></c></row>
      <row r="4"><c r="A4" t="s"><v>6</v></c></row>
      <row r="5"><c r="B5" t="s"><v>7</v></c></row>
    </sheetData></worksheet>'''
    lines = [
        type("Line", (), {"product": True, "name": "Product", "qty": 2})(),
        type("Line", (), {"product": False, "name": "Section", "qty": 0})(),
    ]
    root, _ = render(package(sheet, strings), {"lines": lines})
    result = cells(root)
    assert text(result["A1"]) == "Product"
    assert text(result["B1"]) == "2"
    assert text(result["B2"]) == "Section"
    assert result["B3"].get("t") == "s" and value(result["B3"]) == "7"
    assert len(root.findall("x:sheetData/x:row", NS)) == 3


def test_row_expansion_drops_only_empty_rows_pushed_past_excel_limit():
    strings = [
        "{% tr %}{% for item in items %}{{ item }}{% tr %}{% endfor %}",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:A1"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="1048575" ht="12.75" customHeight="1"/>
      <row r="1048576" ht="12.75" customHeight="1"/>
    </sheetData></worksheet>'''
    root, _ = render(package(sheet, strings), {"items": ["One", "Two"]})
    rows = root.findall("x:sheetData/x:row", NS)
    assert [int(row.get("r")) for row in rows] == [1, 2, 1048576]


def test_worksheet_without_template_tags_is_preserved_byte_for_byte():
    sheet = f'''<worksheet xmlns="{MAIN}"><sheetData>
      <row r="1"><c r="A1" t="inlineStr"><is><t>unchanged</t></is></c></row>
    </sheetData></worksheet>'''
    source = package(sheet)
    output = io.BytesIO()
    template = XlsxTemplate(source)
    template.render({"unused": "value"})
    template.save(output)
    with zipfile.ZipFile(io.BytesIO(source)) as before, zipfile.ZipFile(output) as after:
        assert after.read("xl/worksheets/sheet1.xml") == before.read(
            "xl/worksheets/sheet1.xml"
        )


def test_ignorable_namespace_prefixes_survive_rendering():
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    x14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
    sheet = f'''<worksheet xmlns="{MAIN}" xmlns:mc="{mc}" xmlns:x14ac="{x14ac}" mc:Ignorable="x14ac"><sheetData>
      <row r="1" x14ac:dyDescent="0.25"><c r="A1" t="inlineStr"><is><t>{{{{ value }}}}</t></is></c></row>
    </sheetData></worksheet>'''
    output = io.BytesIO()
    template = XlsxTemplate(package(sheet))
    template.render({"value": "rendered"})
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        rendered = archive.read("xl/worksheets/sheet1.xml")
    assert b'xmlns:mc="' + mc.encode() + b'"' in rendered
    assert b'xmlns:x14ac="' + x14ac.encode() + b'"' in rendered
    assert b'mc:Ignorable="x14ac"' in rendered
    ET.fromstring(rendered)


def test_placeholder_image_is_replaced_and_cloned_with_row_loop():
    strings = ["{%r for item in items %}", "{% img item.image %}", "{%r endfor %}"]
    sheet = f'''<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c></row>
      <row r="3"><c r="A3" t="s"><v>2</v></c></row>
    </sheetData><drawing r:id="rId1"/></worksheet>'''
    drawing = f'''<xdr:wsDr xmlns:xdr="{DRAWING}" xmlns:a="{DRAWING_MAIN}" xmlns:r="{REL}">
      <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Placeholder"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
    </xdr:wsDr>'''
    extra = {
        "xl/worksheets/_rels/sheet1.xml.rels": f'''<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>'''.encode(),
        "xl/drawings/drawing1.xml": drawing.encode(),
        "xl/drawings/_rels/drawing1.xml.rels": f'''<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/image" Target="../media/image1.png"/></Relationships>'''.encode(),
        "xl/media/image1.png": PNG,
    }
    output = io.BytesIO()
    template = XlsxTemplate(package(sheet, strings, extra))
    image = base64.b64encode(PNG)
    template.render({"items": [{"image": image}, {"image": image}]})
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        drawing_root = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
        names = set(archive.namelist())
    anchors = list(drawing_root)
    drawing_ns = {"xdr": DRAWING, "a": DRAWING_MAIN}
    assert [anchor.findtext("xdr:from/xdr:row", namespaces=drawing_ns) for anchor in anchors] == ["0", "1"]
    assert len({anchor.find(".//a:blip", drawing_ns).get(f"{{{REL}}}embed") for anchor in anchors}) == 2
    assert {"xl/media/image2.png", "xl/media/image3.png"} <= names


def test_insert_image_creates_drawing_relationships_and_media():
    sheet = f'''<worksheet xmlns="{MAIN}"><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="10"/></cols><sheetData>
      <row r="1" ht="20"><c r="A1" t="inlineStr"><is><t>{{% insert_img logo %}}</t></is></c></row>
    </sheetData></worksheet>'''
    output = io.BytesIO()
    template = XlsxTemplate(package(sheet))
    template.render({"logo": base64.b64encode(PNG)})
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
        sheet_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        drawing_root = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
        media = archive.read("xl/media/image1.png")
    assert sheet_root.find("x:drawing", {"x": MAIN}) is not None
    assert {
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/media/image1.png",
    } <= names
    assert drawing_root.find(f"{{{DRAWING}}}oneCellAnchor") is not None
    assert media.startswith(b"\x89PNG")


def test_independent_block_regions_expand_and_shift_their_own_columns():
    strings = [
        "{%b A:D for fruit in fruits %}",
        "{%b F:H for vegetable in vegetables %}",
        "{{ loop.index }}. {{ fruit.name }}",
        "{% xv fruit.qty %}",
        "{{ loop.index }}. {{ vegetable.name }}",
        "{%b A:D endfor %}",
        "{%b F:H endfor %}",
        "FRUIT FOOTER",
        "VEGETABLE FOOTER",
        "UNCHANGED",
        "{% insert_img vegetable.image %}",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:K4"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="F1" t="s"><v>1</v></c></row>
      <row r="2" ht="24" customHeight="1"><c r="A2" t="s"><v>2</v></c><c r="C2" t="s"><v>3</v></c><c r="D2"><f>C2*2</f><v>0</v></c><c r="F2" t="s"><v>4</v></c><c r="H2" t="s"><v>10</v></c></row>
      <row r="3"><c r="A3" t="s"><v>5</v></c><c r="F3" t="s"><v>6</v></c></row>
      <row r="4"><c r="A4" t="s"><v>7</v></c><c r="F4" t="s"><v>8</v></c><c r="K4" t="s"><v>9</v></c></row>
    </sheetData><mergeCells count="2"><mergeCell ref="A2:B2"/><mergeCell ref="F2:G2"/></mergeCells></worksheet>'''
    output = io.BytesIO()
    template = XlsxTemplate(package(sheet, strings))
    template.render(
        {
            "fruits": [
                {"name": "Apple", "qty": 1},
                {"name": "Mango", "qty": 2},
                {"name": "Orange", "qty": 3},
            ],
            "vegetables": [
                {"name": "Spinach", "image": base64.b64encode(PNG)},
                {"name": "Carrot", "image": base64.b64encode(PNG)},
            ],
        }
    )
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
    result = cells(root)
    assert [text(result[f"A{row}"]) for row in (1, 2, 3)] == [
        "1. Apple", "2. Mango", "3. Orange"
    ]
    assert [value(result[f"C{row}"]) for row in (1, 2, 3)] == ["1", "2", "3"]
    assert [
        result[f"D{row}"].findtext("x:f", namespaces=NS) for row in (1, 2, 3)
    ] == ["C1*2", "C2*2", "C3*2"]
    assert [text(result[f"F{row}"]) for row in (1, 2)] == [
        "1. Spinach", "2. Carrot"
    ]
    assert value(result["F3"]) == "8"
    assert value(result["A4"]) == "7"
    assert value(result["K4"]) == "9"
    assert [
        merge.get("ref")
        for merge in root.findall("x:mergeCells/x:mergeCell", NS)
    ] == ["A1:B1", "A2:B2", "A3:B3", "F1:G1", "F2:G2"]
    assert all(root.find(f"x:sheetData/x:row[@r='{row}']", NS).get("ht") == "24" for row in (1, 2, 3))
    assert [
        anchor.findtext("xdr:from/xdr:row", namespaces={"xdr": DRAWING})
        for anchor in drawing
    ] == ["0", "1"]


def test_overlapping_block_regions_fail_clearly():
    strings = [
        "{%b A:D for left in lefts %}",
        "{%b C:F for right in rights %}",
        "{{ left }}",
        "{{ right }}",
        "{%b A:D endfor %}",
        "{%b C:F endfor %}",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>
      <row r="2"><c r="A2" t="s"><v>2</v></c><c r="C2" t="s"><v>3</v></c></row>
      <row r="3"><c r="A3" t="s"><v>4</v></c><c r="C3" t="s"><v>5</v></c></row>
    </sheetData></worksheet>'''
    template = XlsxTemplate(package(sheet, strings))
    with unittest.TestCase().assertRaisesRegex(
        UnsupportedFeatureError, "overlap"
    ):
        template.render({"lefts": [1], "rights": [2]})


def test_block_region_clones_and_replaces_placeholder_images():
    strings = [
        "{%b A:B for item in items %}",
        "{% img item.image %}",
        "{%b A:B endfor %}",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c></row>
      <row r="3"><c r="A3" t="s"><v>2</v></c></row>
    </sheetData><drawing r:id="rId1"/></worksheet>'''
    drawing = f'''<xdr:wsDr xmlns:xdr="{DRAWING}" xmlns:a="{DRAWING_MAIN}" xmlns:r="{REL}">
      <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Placeholder"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
    </xdr:wsDr>'''
    extra = {
        "xl/worksheets/_rels/sheet1.xml.rels": f'''<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>'''.encode(),
        "xl/drawings/drawing1.xml": drawing.encode(),
        "xl/drawings/_rels/drawing1.xml.rels": f'''<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/image" Target="../media/image1.png"/></Relationships>'''.encode(),
        "xl/media/image1.png": PNG,
    }
    output = io.BytesIO()
    template = XlsxTemplate(package(sheet, strings, extra))
    image = base64.b64encode(PNG)
    template.render({"items": [{"image": image}, {"image": image}]})
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        drawing_root = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
    assert [
        anchor.findtext("xdr:from/xdr:row", namespaces={"xdr": DRAWING})
        for anchor in drawing_root
    ] == ["0", "1"]
    assert len(
        {
            anchor.find(".//a:blip", {"a": DRAWING_MAIN}).get(f"{{{REL}}}embed")
            for anchor in drawing_root
        }
    ) == 2


def test_empty_block_region_removes_its_span_without_shifting_other_columns():
    strings = [
        "{%b A:B for item in items %}",
        "{{ item }}",
        "{%b A:B endfor %}",
        "LOCAL FOOTER",
        "GLOBAL CELL",
    ]
    sheet = f'''<worksheet xmlns="{MAIN}"><dimension ref="A1:K4"/><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c></row>
      <row r="3"><c r="A3" t="s"><v>2</v></c></row>
      <row r="4"><c r="A4" t="s"><v>3</v></c><c r="K4" t="s"><v>4</v></c></row>
    </sheetData></worksheet>'''
    root, _ = render(package(sheet, strings), {"items": []})
    result = cells(root)
    assert value(result["A1"]) == "3"
    assert value(result["K4"]) == "4"


def test_shipped_invoice_example_template_renders():
    output = io.BytesIO()
    template = XlsxTemplate(
        Path(__file__).parents[1] / "examples" / "invoice_template.xlsx"
    )
    template.render(
        {
            "company": "DLI & Co",
            "report_date": "2026-08-24",
            "lines": [
                {"name": "Freight", "qty": 2, "price": 10, "active": True, "image": base64.b64encode(PNG)},
                {"name": "Handling", "qty": 1, "price": 5, "active": False, "image": base64.b64encode(PNG)},
            ],
        }
    )
    template.save(output)
    with zipfile.ZipFile(output) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml")
        drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
    assert b"DLI &amp; Co" in sheet
    assert b"Freight" in sheet and b"(inactive) Handling" in sheet
    assert b"SUM(E5:E6)" in sheet
    assert len(list(drawing)) == 2


def test_img_without_placeholder_fails_clearly():
    sheet = f'''<worksheet xmlns="{MAIN}"><sheetData>
      <row r="1"><c r="A1" t="inlineStr"><is><t>{{% img logo %}}</t></is></c></row>
    </sheetData></worksheet>'''
    template = XlsxTemplate(package(sheet))
    with unittest.TestCase().assertRaisesRegex(
        UnsupportedFeatureError, "requires a placeholder image"
    ):
        template.render({"logo": base64.b64encode(PNG)})


if __name__ == "__main__":
    tests = [value for name, value in globals().items() if name.startswith("test_")]
    for test in tests:
        test()
    print(f"{len(tests)} tests passed")
