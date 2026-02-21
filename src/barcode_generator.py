from barcode import Code128
from barcode.writer import ImageWriter
from pathlib import Path


def generate_barcode(value, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    file_path = output_dir / f"{value}"

    options = {
        "module_width": 0.28,
        "module_height": 10.5,     # Controlled height
        "quiet_zone": 6.5,
        "font_size": 0,          # No auto text
        "text_distance": 1,
        "dpi": 300,
        "write_text": False,
    }

    barcode = Code128(value, writer=ImageWriter())
    barcode.save(str(file_path), options=options)

    return str(file_path) + ".png"
