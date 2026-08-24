# xlsx-jinja

Template engine `.xlsx` berbasis Jinja2. Library mengedit OOXML di dalam package XLSX secara langsung, sehingga style, formula, merged cells, drawing, dan bagian workbook yang tidak dirender tetap dipertahankan.

Desain API dan structural tag mengikuti pendekatan [`python-docx-template`](https://github.com/elapouya/python-docx-template).

## Instalasi

```bash
pip install xlsx-jinja
```

Dependency runtime:

- Python 3.10+
- Jinja2
- Pillow

## Quick start

```python
from xlsx_jinja import XlsxTemplate

context = {
    "company": "Dia Logistics Indonesia",
    "lines": [
        {"name": "Freight", "qty": 2, "price": 150_000},
        {"name": "Handling", "qty": 1, "price": 75_000},
    ],
}

template = XlsxTemplate("invoice_template.xlsx")
template.render(context)
template.save("invoice_output.xlsx")
```

Input dan output dapat berupa path, `bytes`, atau binary file-like object seperti `BytesIO`.

## Syntax dasar

### Variable

Isi cell Excel dengan Jinja biasa:

```jinja2
{{ company }}
{{ customer.name }}
{{ value | default('') }}
```

Output text di-XML-escape secara default.

### Native Excel value

`{{ ... }}` menghasilkan text. Gunakan `{% xv ... %}` agar number, boolean, date, atau datetime tetap bertipe native Excel:

```jinja2
{% xv line.qty %}
{% xv line.price %}
{% xv invoice.invoice_date %}
```

Native number dapat dipakai oleh formula Excel seperti `SUM()` dan perkalian.

## Looping

### Loop satu atau beberapa row — direkomendasikan

Letakkan pembuka dan penutup pada dedicated marker rows. Marker rows akan dihapus dari output.

| Row | A | B | C |
|---|---|---|---|
| 4 | `{%r for line in lines %}` | | |
| 5 | `{{ loop.index }}` | `{{ line.name }}` | `{% xv line.qty %}` |
| 6 | `{%r endfor %}` | | |

Row 5 digandakan untuk setiap item.

Satu item juga boleh memakai beberapa template rows:

```text
Row 4: {%r for line in lines %}
Row 5: {{ line.name }} | {% xv line.qty %}
Row 6: Notes            | {{ line.notes or '' }}
Row 7: {%r endfor %}
```

Setiap item menghasilkan row data dan notes.

### Loop satu row dengan `{% tr %}`

Syntax lama tetap didukung untuk loop sederhana:

```text
A5: {% tr %}{% for line in lines %}{{ line.name }}
B5: {% xv line.qty %}
C5: {% xv line.price %}{% tr %}{% endfor %}
```

Gunakan `{%r ... %}` untuk template baru karena lebih mudah dibaca dan mendukung multi-row block.

### Cell notes `beforerow`

Control tag lama pada Excel Note juga didukung:

```text
Cell A5 note: beforerow{% for line in lines %}
Cell A6 note: beforerow{% endfor %}
```

## Independent block/region scope

`{%r ... %}` mengulang seluruh row. Gunakan `{%b RANGE ... %}` jika beberapa vertical list harus berkembang independen pada worksheet sama.

Contoh buah pada `A:D` dan sayur pada `F:I`:

| Row | A:D | F:I |
|---|---|---|
| 10 | `{%b A:D for fruit in fruits %}` | `{%b F:I for vegetable in vegetables %}` |
| 11 | `{{ loop.index }}. {{ fruit.name }}` | `{{ loop.index }}. {{ vegetable.name }}` |
| 12 | `{%b A:D endfor %}` | `{%b F:I endfor %}` |

Jika buah berisi empat item dan sayur dua item, `A:D` menghasilkan empat rows sedangkan `F:I` hanya dua. Cell di bawah `A:D` bergeser empat-item flow; cell di bawah `F:I` bergeser dua-item flow. Kolom lain tidak bergeser.

Marker pembuka dan penutup wajib menyebut range sama:

```jinja2
{%b A:D for fruit in fruits %}
{%b A:D endfor %}
```

### Multi-row region body

Satu item dapat memakai beberapa rows:

```text
A10: {%b A:D for fruit in fruits %}
A11: {{ fruit.name }}
A12: Notes: {{ fruit.notes or '' }}
A13: {%b A:D endfor %}
```

Rows 11–12 digandakan sebagai satu block untuk setiap buah.

### Region features

Di dalam region tersedia:

```jinja2
{{ loop.index }}
{{ loop.index0 }}
{{ loop.first }}
{{ loop.last }}
{{ loop.length }}
{{ loop.revindex }}
{{ loop.revindex0 }}
```

Region juga mendukung:

- `{% xv ... %}` native values
- inline `if/else`
- merged cells yang sepenuhnya berada dalam range
- formula relatif di template body
- `{% img ... %}` dan `{% insert_img ... %}`
- body styles, borders, row height, dan column width
- list kosong; region span hilang dan content lokal di bawahnya naik
- beberapa region berurutan dengan range kolom sama

### Region rules dan batas

- Range menggunakan column boundaries, misalnya `A:D`, bukan cell range `A10:D12`.
- Marker harus berada di dalam range yang dideklarasikan.
- Minimal satu body row harus berada di antara marker.
- Region saat ini mendukung `for/endfor`; `if` tetap ditulis di body.
- Region tidak boleh nested atau overlap; nested `for` di dalam body juga belum didukung.
- Partial column overlap seperti `A:E` dengan `D:H` ditolak.
- Merged cell atau drawing yang melintasi region boundary ditolak agar output tidak korup.
- Row height adalah properti global Excel. Jika dua region menulis target row sama, tinggi terbesar dipakai.
- Jangan nest `{%r %}` di dalam `{%b %}`. Gunakan salah satu scope untuk area yang sama.
- Formula di dalam body digeser seperti copy-cell. Aggregate formula di luar region yang harus mencakup seluruh hasil sebaiknya dihitung dari context atau diperiksa setelah ekspansi.
- Data-validation, conditional-formatting, table, chart, dan pivot ranges belum digeser secara lokal.

Template siap pakai tersedia di [`examples/region_template.xlsx`](examples/region_template.xlsx).

Jalankan contoh:

```bash
python examples/render_region_example.py
```

## Condition

Condition lengkap dapat ditulis dalam satu cell:

```jinja2
{% if line.active %}{{ line.name }}{% else %}(inactive) {{ line.name }}{% endif %}
```

Condition yang dibuka dan ditutup pada cell berbeda dalam row sama digabung menjadi row scope. Seluruh row hilang saat kondisi salah:

```text
A6: {% if line.active %}{{ line.name }}
B6: {% xv line.qty %}
C6: {% xv line.price %}{% endif %}
```

Atau condition satu row memakai marker:

```text
Row 8:  {%r if show_total %}
Row 9:  TOTAL | {% xv total %}
Row 10: {%r endif %}
```

Condition dalam loop:

```jinja2
{% if line.display_type == 'product' %}{{ line.name }}{% endif %}
```

## Counter dengan namespace

```text
Row 4: {% set ns = namespace(counter=0) %}
Row 5: {%r for line in lines %}
Row 6: {% if line.active %}{% set ns.counter = ns.counter + 1 %}{{ ns.counter }}{% endif %}
Row 7: {%r endfor %}
```

## Images

### Mengganti placeholder

Tambahkan dummy image di Excel. Posisi top-left image harus berada pada cell berisi tag:

```jinja2
{% img line.image %}
```

Jika beberapa placeholder memiliki top-left cell sama, gunakan index berbasis nol:

```jinja2
{% img line.image, 1 %}
```

### Insert tanpa placeholder

```jinja2
{% insert_img line.image %}
```

Image baru mengikuti ukuran row dan column target. Dalam loop, image otomatis dibuat dan diposisikan untuk setiap output row.

Nilai image yang didukung:

- Odoo base64 `bytes`
- raw image `bytes`
- `BytesIO` atau binary stream
- Pillow `Image`
- local file path

Image output dinormalisasi menjadi PNG, termasuk input WebP.

## Custom Jinja globals

```python
from xlsx_jinja import XlsxTemplate

book = XlsxTemplate("template.xlsx")
book.set_jinja_globals(
    currency=lambda value: f"Rp {value:,.0f}",
)
book.render({"total": 125000})
book.save("output.xlsx")
```

Template:

```jinja2
{{ currency(total) }}
```

Custom `Environment` juga dapat diberikan:

```python
from jinja2 import Environment, StrictUndefined

book.render(
    context,
    jinja_env=Environment(undefined=StrictUndefined),
)
```

## Multi-sheet

Semua worksheet dirender dengan context sama. Library menambahkan:

- `sheet_name`: nama worksheet aktif
- `tpl_idx`: index worksheet, mulai dari `0`

```jinja2
{{ tpl_idx + 1 }} - {{ sheet_name }}
```

## Error handling

```python
from xlsx_jinja import UnsupportedFeatureError, XlsxJinjaError

try:
    template.render(context)
    template.save("output.xlsx")
except UnsupportedFeatureError as error:
    print(f"Template feature tidak didukung: {error}")
except XlsxJinjaError as error:
    print(f"XLSX tidak valid: {error}")
```

Renderer menolak output berisi data di atas batas Excel `1,048,576` rows. Empty formatting rows yang terdorong melewati batas akibat loop dibuang otomatis.

## Contoh siap jalan

Folder [`examples/`](examples/) berisi:

- `invoice_template.xlsx` — template row-scope siap diedit
- `render_example.py` — render menjadi `invoice_output.xlsx`
- `create_template.py` — membuat ulang invoice template
- `region_template.xlsx` — dua independent regions dalam worksheet sama
- `render_region_example.py` — render menjadi `region_output.xlsx`
- `create_region_template.py` — membuat ulang region template

Jalankan:

```bash
python examples/render_example.py
python examples/render_region_example.py
```

Buka `examples/invoice_output.xlsx` dan `examples/region_output.xlsx` memakai Excel atau LibreOffice.

Untuk membuat ulang template contoh:

```bash
pip install -e '.[examples]'
python examples/create_template.py
python examples/create_region_template.py
```

## Menjalankan test

Test memakai standard library dan dapat dijalankan tanpa pytest:

```bash
python tests/test_template.py
```

## Batas saat ini

Didukung:

- variable dan standard Jinja expressions
- global row scope `{%r ... %}`
- independent block/region scope `{%b RANGE ... %}`
- condition dan multi-row loop body
- native number/date/boolean
- image replacement dan insertion
- repeated images
- shared strings dan inline strings
- straightforward formula translation
- merged ranges pada repeated rows
- preservation bagian ZIP yang tidak berubah

Belum ditulis ulang setelah row expansion:

- cross-sheet formula
- structured table references
- array/shared/spill formula ranges
- data-validation ranges
- conditional-formatting ranges
- chart data ranges
- pivot cache ranges

Bagian tersebut tetap dipertahankan, tetapi range-nya tidak otomatis diperluas.
