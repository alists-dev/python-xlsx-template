# xlsx-jinja

Jinja2-based `.xlsx` template engine. It directly edits OOXML parts inside the XLSX ZIP package, preserving original styles, formulas, merged cells, drawings, print setups, and non-rendered workbook parts without reconstructing through full spreadsheet object models.

API design and structural tags follow the approach of [`python-docx-template`](https://github.com/elapouya/python-docx-template).

For documentation in Bahasa Indonesia, see [README_ID.md](README_ID.md).

---

## Installation

```bash
pip install xlsx-jinja
```

Runtime dependencies:
- Python 3.10+
- Jinja2
- Pillow

## Quick Start

```python
from xlsx_jinja import XlsxTemplate

context = {
    "company": "Abracadabra!",
    "lines": [
        {"name": "Freight", "qty": 2, "price": 150_000},
        {"name": "Handling", "qty": 1, "price": 75_000},
    ],
}

template = XlsxTemplate("invoice_template.xlsx")
template.render(context)
template.save("invoice_output.xlsx")
```

Accepts file paths, raw `bytes`, or binary file-like objects (such as `BytesIO`) for both input and output.

---

## Basic Syntax

### Variables
Write standard Jinja expressions inside cells:
```jinja2
{{ company }}
{{ customer.name }}
{{ value | default('') }}
```
Rendered text is automatically XML-escaped by default.

### Native Excel Values
`{{ ... }}` renders values as text strings. Use `{% xv ... %}` to write numbers, booleans, dates, or datetimes as native Excel data types:
```jinja2
{% xv line.qty %}
{% xv line.price %}
{% xv invoice.invoice_date %}
```
Native numeric cells work directly with Excel formulas (e.g. `=SUM()`).

---

## Row-Scope Looping (`{%r ... %}`)

### Dedicated Marker Rows (Recommended)
Place control tags in dedicated marker rows. Marker rows are stripped from final output:

| Row | A | B | C |
|---|---|---|---|
| 4 | `{%r for line in lines %}` | | |
| 5 | `{{ loop.index }}` | `{{ line.name }}` | `{% xv line.qty %}` |
| 6 | `{%r endfor %}` | | |

Row 5 is repeated for every item.

A single loop item can span multiple template rows:
```text
Row 4: {%r for line in lines %}
Row 5: {{ line.name }} | {% xv line.qty %}
Row 6: Notes            | {{ line.notes or '' }}
Row 7: {%r endfor %}
```

### Inline Row Loop (`{% tr %}`)
Inline tags on the same row are supported for simple repeats:
```text
A5: {% tr %}{% for line in lines %}{{ line.name }}
B5: {% xv line.qty %}
C5: {% xv line.price %}{% tr %}{% endfor %}
```

### Legacy Cell Notes (`beforerow`)
Excel cell notes (comments) are also supported:
```text
Cell A5 note: beforerow{% for line in lines %}
Cell A6 note: beforerow{% endfor %}
```

---

## Independent Block/Region Scope (`{%b RANGE ... %}`)

While `{%r ... %}` repeats the entire row globally, `{%b RANGE ... %}` repeats a vertical column subset independently without forcing adjacent columns to shift by the same amount.

Example: Fruits in `A:D` and Vegetables in `F:I`:

| Row | A:D | F:I |
|---|---|---|
| 10 | `{%b A:D for fruit in fruits %}` | `{%b F:I for vegetable in vegetables %}` |
| 11 | `{{ loop.index }}. {{ fruit.name }}` | `{{ loop.index }}. {{ vegetable.name }}` |
| 12 | `{%b A:D endfor %}` | `{%b F:I endfor %}` |

If `fruits` has 4 items and `vegetables` has 2 items:
- Columns `A:D` produce 4 output rows.
- Columns `F:I` produce 2 output rows.
- Footers/content below `A:D` shift by 4 items; content below `F:I` shifts by 2 items. Other columns do not shift.

### Region Rules & Boundaries
- Range format: column boundaries (e.g. `A:D`), not 2D cell blocks (`A10:D12`).
- Both opening and closing markers must specify the same range name (`{%b A:D for ... %}` and `{%b A:D endfor %}`).
- Region tags must be placed inside their declared column range.
- At least one template body row must exist between markers.
- Supported tags: `for/endfor`. Place conditionals (`if`) inside body cells.
- Regions must not overlap columns partially or nest inside other regions.
- Merged cells and images must reside entirely inside the region boundary.

---

## Conditions (`if / else`)

### Single-Cell Inline Condition
```jinja2
{% if line.active %}{{ line.name }}{% else %}(inactive) {{ line.name }}{% endif %}
```

### Cross-Cell Row-Scope Condition
Opening and closing tags across different cells in the same row are merged at row scope (the entire row disappears when condition is falsy):
```text
A6: {% if line.active %}{{ line.name }}
B6: {% xv line.qty %}
C6: {% xv line.price %}{% endif %}
```

### Dedicated Marker Row Condition
```text
Row 8:  {%r if show_total %}
Row 9:  TOTAL | {% xv total %}
Row 10: {%r endif %}
```

---

## Stateful Iterations with Namespace

```text
Row 4: {% set ns = namespace(counter=0) %}
Row 5: {%r for line in lines %}
Row 6: {% if line.active %}{% set ns.counter = ns.counter + 1 %}{{ ns.counter }}{% endif %}
Row 7: {%r endfor %}
```

---

## Images

### Replace Placeholder Image
Place a placeholder image in the Excel template with its top-left corner on the tag cell:
```jinja2
{% img line.image %}
```
If multiple placeholders share the same top-left anchor, use a zero-based index:
```jinja2
{% img line.image, 1 %}
```

### Insert Image Without Placeholder
```jinja2
{% insert_img line.image %}
```
Automatically scales image to fit target cell dimensions. Images inside loop rows are cloned and anchored per output row.

Accepted values:
- Base64-encoded `bytes` (e.g. standard Odoo binary fields)
- Raw image `bytes`
- `BytesIO` / binary stream
- Pillow `Image`
- Local file path

All output images are normalized to PNG in-memory (including WebP inputs).

---

## Custom Jinja Globals & Environment

```python
from xlsx_jinja import XlsxTemplate

book = XlsxTemplate("template.xlsx")
book.set_jinja_globals(
    currency=lambda val: f"Rp {val:,.0f}",
    zip=zip,
)
book.render({"total": 125000})
book.save("output.xlsx")
```

Custom Jinja `Environment`:
```python
from jinja2 import Environment, StrictUndefined

book.render(context, jinja_env=Environment(undefined=StrictUndefined))
```

## Multi-Sheet Rendering
All worksheets are rendered using the same context. Every worksheet also receives:
- `sheet_name`: title of the active worksheet
- `tpl_idx`: zero-based index of the worksheet (`0, 1, ...`)

---

## Runnable Examples

Explore the [`examples/`](examples/) directory:
- `invoice_template.xlsx` — Full row-scope loop example template
- `render_example.py` — Renders `invoice_template.xlsx` to `invoice_output.xlsx`
- `create_template.py` — Script to generate `invoice_template.xlsx`
- `region_template.xlsx` — Two independent block regions (`A:D` and `F:I`)
- `render_region_example.py` — Renders `region_template.xlsx` to `region_output.xlsx`
- `create_region_template.py` — Script to generate `region_template.xlsx`

Run examples:
```bash
python examples/render_example.py
python examples/render_region_example.py
```

## Running Tests

Run the test suite (uses standard library `unittest` under the hood; no pytest required):
```bash
python tests/test_template.py
```

---

## Current Scope & Limitations

### Supported
- Variables & standard Jinja2 expressions/filters
- Global row scope (`{%r %}`) & inline row scope (`{% tr %}`)
- Independent block/region scope (`{%b RANGE %}`)
- Multi-row loop blocks & nested row loops
- Native numbers, booleans, dates, datetimes (`{% xv %}`)
- Image placeholder replacement (`{% img %}`) & auto-scaled insertion (`{% insert_img %}`)
- Automatic row-loop image replication & positioning
- Merged ranges inside loop rows & independent regions
- Relative formula row translation on repeated rows
- Byte-for-byte preservation of unrendered workbook package parts

### Out of Scope / Not Rewritten Automatically on Row Expansion
- Cross-sheet formulas
- Structured table references
- Dynamic spill / array formula ranges
- Data validation ranges
- Conditional formatting ranges
- Chart series ranges & Pivot cache ranges

These parts remain intact in the workbook output, but their range references are not expanded automatically.
