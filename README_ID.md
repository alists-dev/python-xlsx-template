# xlsx-jinja

Template engine `.xlsx` berbasis Jinja2. Engine ini mengedit OOXML langsung di dalam file package ZIP XLSX, sehingga format cell/style, formula, merged cells, drawing/image, dan sheet yang tidak dirender tetap terjaga tanpa melalui rekonstruksi object model Excel yang berat.

Desain API dan tag struktural mengikuti konsep [`python-docx-template`](https://github.com/elapouya/python-docx-template).

Untuk dokumentasi bahasa Inggris (utama), lihat [README.md](README.md).

---

## Instalasi

```bash
pip install xlsx-jinja
```

Dependensi runtime:
- Python 3.10+
- Jinja2
- Pillow

## Panduan Cepat

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

Input dan output mendukung file path, raw `bytes`, atau objek stream binary seperti `BytesIO`.

---

## Syntax Dasar

### Variabel Text
Tulis ekspresi Jinja biasa di dalam cell:
```jinja2
{{ company }}
{{ customer.name }}
{{ value | default('') }}
```
Teks yang dihasilkan otomatis di-escape agar aman terhadap XML OOXML.

### Nilai Native Excel (`{% xv %}`)
`{{ ... }}` mencetak nilai sebagai teks (string). Gunakan `{% xv ... %}` agar tipe data angka, tanggal, datetime, atau boolean disimpan sebagai data native Excel:
```jinja2
{% xv line.qty %}
{% xv line.price %}
{% xv invoice.invoice_date %}
```
Cell angka native bisa langsung dihitung oleh formula Excel seperti `=SUM()`.

---

## Looping Baris (`{%r ... %}`)

### Baris Marker Terpisah (Direkomendasikan)
Tulis tag pembuka dan penutup pada baris khusus (marker row). Baris marker otomatis dihapus dari output:

| Row | A | B | C |
|---|---|---|---|
| 4 | `{%r for line in lines %}` | | |
| 5 | `{{ loop.index }}` | `{{ line.name }}` | `{% xv line.qty %}` |
| 6 | `{%r endfor %}` | | |

Baris 5 akan digandakan untuk setiap data.

Satu perulangan juga bisa mencakup beberapa baris template sekaligus:
```text
Row 4: {%r for line in lines %}
Row 5: {{ line.name }} | {% xv line.qty %}
Row 6: Notes            | {{ line.notes or '' }}
Row 7: {%r endfor %}
```

### Looping Inline Satu Baris (`{% tr %}`)
Tag inline pada baris yang sama tetap didukung untuk kasus perulangan satu baris sederhana:
```text
A5: {% tr %}{% for line in lines %}{{ line.name }}
B5: {% xv line.qty %}
C5: {% xv line.price %}{% tr %}{% endfor %}
```

### Komentar / Catatan Cell (`beforerow`)
Mendukung sintaks legacy via catatan cell (Excel comment/note):
```text
Catatan Cell A5: beforerow{% for line in lines %}
Catatan Cell A6: beforerow{% endfor %}
```

---

## Looping Region / Blok Kolom Mandiri (`{%b RANGE ... %}`)

Jika `{%r ... %}` mengulang seluruh baris secara horizontal penuh, `{%b RANGE ... %}` hanya mengulang subset kolom tertentu secara vertikal tanpa memaksa kolom di sebelahnya bergeser sama panjang.

Contoh: Daftar Buah di kolom `A:D` dan Daftar Sayur di kolom `F:I`:

| Row | A:D | F:I |
|---|---|---|
| 10 | `{%b A:D for fruit in fruits %}` | `{%b F:I for vegetable in vegetables %}` |
| 11 | `{{ loop.index }}. {{ fruit.name }}` | `{{ loop.index }}. {{ vegetable.name }}` |
| 12 | `{%b A:D endfor %}` | `{%b F:I endfor %}` |

Jika data `fruits` berisi 4 baris dan `vegetables` berisi 2 baris:
- Kolom `A:D` menghasilkan 4 baris output.
- Kolom `F:I` menghasilkan 2 baris output.
- Footer / konten di bawah kolom `A:D` bergeser turun sebanyak 4 baris; konten di bawah `F:I` bergeser turun sebanyak 2 baris. Kolom lain tidak terpengaruh.

### Aturan & Batasan Region
- Format range menggunakan batas kolom (contoh: `A:D`, bukan koordinat cell 2D `A10:D12`).
- Marker pembuka dan penutup wajib menyebut nama range yang sama (`{%b A:D for ... %}` dan `{%b A:D endfor %}`).
- Tag marker harus berada di dalam batas kolom yang dideklarasikan.
- Minimal harus ada 1 baris template body di antara baris marker.
- Tag yang didukung pada level region: `for/endfor`. Kondisi `if` diletakkan di cell body.
- Region tidak boleh saling bertumpuk (overlap sebagian) atau bersarang (nested).
- Merged cell dan image wajib berada penuh di dalam batas kolom region.

---

## Kondisi (`if / else`)

### Kondisi Inline Satu Cell
```jinja2
{% if line.active %}{{ line.name }}{% else %}(inactive) {{ line.name }}{% endif %}
```

### Kondisi Melintasi Cell dalam Satu Baris
Pembuka `if` dan penutup `endif` di cell berbeda pada baris yang sama otomatis digabung menjadi row scope (seluruh baris hilang jika kondisi false):
```text
A6: {% if line.active %}{{ line.name }}
B6: {% xv line.qty %}
C6: {% xv line.price %}{% endif %}
```

### Baris Marker Khusus Kondisi
```text
Row 8:  {%r if show_total %}
Row 9:  TOTAL | {% xv total %}
Row 10: {%r endif %}
```

---

## Namespace untuk Counter Antar-Iterasi

```text
Row 4: {% set ns = namespace(counter=0) %}
Row 5: {%r for line in lines %}
Row 6: {% if line.active %}{% set ns.counter = ns.counter + 1 %}{{ ns.counter }}{% endif %}
Row 7: {%r endfor %}
```

---

## Gambar / Image

### Mengganti Gambar Dummy (Placeholder)
Pasang gambar dummy di template Excel dengan sudut kiri atas (top-left) menempel di cell berisi tag:
```jinja2
{% img line.image %}
```
Jika ada beberapa placeholder yang menempel pada cell yang sama, gunakan indeks:
```jinja2
{% img line.image, 1 %}
```

### Menyisipkan Gambar Baru Tanpa Placeholder
```jinja2
{% insert_img line.image %}
```
Otomatis menyesuaikan ukuran lebar kolom dan tinggi baris target. Gambar di dalam looping baris otomatis digandakan per iterasi.

Format input yang diterima:
- Base64 `bytes` (misal field Binary standar Odoo)
- Raw image `bytes`
- Stream binary (`BytesIO`)
- Objek Pillow `Image`
- Path file lokal

Semua format gambar (termasuk WebP) otomatis dikonversi in-memory ke PNG.

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
Seluruh worksheet otomatis dirender dengan context yang sama. Setiap sheet menerima variabel tambahan:
- `sheet_name`: nama judul worksheet yang sedang aktif
- `tpl_idx`: indeks numerik sheet (`0, 1, ...`)

---

## Contoh Template Siap Pakai

Folder [`examples/`](examples/):
- `invoice_template.xlsx` — Template contoh looping baris standar (`{%r %}`)
- `render_example.py` — Script render `invoice_template.xlsx` ke `invoice_output.xlsx`
- `create_template.py` — Script generator pembuat `invoice_template.xlsx`
- `region_template.xlsx` — Template dua region mandiri berdampingan (`A:D` dan `F:I`)
- `render_region_example.py` — Script render `region_template.xlsx` ke `region_output.xlsx`
- `create_region_template.py` — Script generator pembuat `region_template.xlsx`

Jalankan:
```bash
python examples/render_example.py
python examples/render_region_example.py
```

## Menjalankan Unit Test

```bash
python tests/test_template.py
```

---

## Batasan & Cakupan Fitur

### Didukung
- Variabel & ekspresi filter standar Jinja2
- Row scope global (`{%r %}`) & inline (`{% tr %}`)
- Independent block/region scope (`{%b RANGE %}`)
- Blok loop multi-baris & nested loop baris
- Tipe data native number, boolean, date, datetime (`{% xv %}`)
- Penggantian gambar placeholder (`{% img %}`) & auto-insert (`{% insert_img %}`)
- Duplikasi dan penataan koordinat gambar dalam looping baris
- Merged range di dalam baris looping & independent region
- Translasi baris formula relatif pada baris perulangan
- Preservasi byte-for-byte untuk bagian file XLSX yang tidak dirender

### Di Luar Cakupan (Belum Disesuaikan Otomatis Saat Baris Bertambah)
- Formula lintas sheet (cross-sheet formulas)
- Referensi structured table
- Formula dinamis spill / array
- Range Data Validation
- Range Conditional Formatting
- Range data Chart series & Pivot Table cache

Komponen di atas tidak rusak dan tetap tersimpan di file output, tetapi koordinat range-nya tidak melar secara otomatis jika baris bertambah.
