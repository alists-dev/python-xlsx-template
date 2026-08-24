"""Render two independent block regions in one worksheet."""

import base64
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from PIL import Image

from xlsx_jinja import XlsxTemplate

HERE = Path(__file__).parent


@dataclass
class Item:
    name: str
    qty: float
    image: bytes


def image(color: str) -> bytes:
    output = BytesIO()
    Image.new("RGB", (48, 48), color).save(output, "PNG")
    return base64.b64encode(output.getvalue())


def main() -> None:
    context = {
        "fruits": [
            Item("Apple", 10, image("#C00000")),
            Item("Mango", 7, image("#FFC000")),
            Item("Orange", 12, image("#ED7D31")),
            Item("Grape", 20, image("#7030A0")),
        ],
        "vegetables": [
            Item("Spinach", 5, image("#70AD47")),
            Item("Carrot", 8, image("#ED7D31")),
        ],
    }

    template = XlsxTemplate(HERE / "region_template.xlsx")
    template.render(context)
    output = HERE / "region_output.xlsx"
    template.save(output)
    print(f"Created {output}")


if __name__ == "__main__":
    main()
