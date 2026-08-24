"""Render invoice_template.xlsx into invoice_output.xlsx."""

import base64
from dataclasses import dataclass
from datetime import date
from io import BytesIO
from pathlib import Path

from PIL import Image

from xlsx_jinja import XlsxTemplate

HERE = Path(__file__).parent


@dataclass
class Line:
    name: str
    qty: float
    price: float
    image: bytes
    active: bool = True


def image(color: str) -> bytes:
    output = BytesIO()
    Image.new("RGB", (64, 64), color).save(output, "PNG")
    # Odoo Binary fields expose base64 bytes, so example uses same shape.
    return base64.b64encode(output.getvalue())


def main() -> None:
    context = {
        "company": "Dia Logistics Indonesia",
        "report_date": date.today().isoformat(),
        "lines": [
            Line("Ocean Freight", 2, 1_500_000, image("#5B9BD5")),
            Line("Handling", 1, 750_000, image("#70AD47")),
            Line("Optional Service", 3, 125_000, image("#A5A5A5"), active=False),
        ],
    }

    template = XlsxTemplate(HERE / "invoice_template.xlsx")
    template.render(context)
    output = HERE / "invoice_output.xlsx"
    template.save(output)
    print(f"Created {output}")


if __name__ == "__main__":
    main()
