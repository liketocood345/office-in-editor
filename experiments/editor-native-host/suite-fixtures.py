"""Generate original, disposable test documents without starting desktop Office."""
import hashlib
import io
import json
from pathlib import Path
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


def main(root: Path, excel_run: str | None = None) -> None:
    if excel_run is not None and not re.fullmatch(r"[a-f0-9]{32}", excel_run):
        raise ValueError("Invalid Excel fixture run ID")
    excel_file = f"synthetic-{excel_run}.xlsx" if excel_run else "synthetic.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Native fixture"
    sheet.append(["Item", "Count", "Unit", "Total"])
    sheet.append(["Draft", 3, 12])
    sheet.append(["Review", 2, 20])
    sheet["D2"] = "=B2*C2"
    sheet["D3"] = "=B3*C3"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="28544B")
    sheet.column_dimensions["A"].width = 18
    sheet.freeze_panes = "B2"
    workbook = io.BytesIO()
    book.save(workbook)

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(0.8), Inches(1), Inches(8), Inches(2))
    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = "PowerPoint native slot fixture"
    paragraph.font.size = Pt(28)
    paragraph.font.bold = True
    paragraph.font.color.rgb = RGBColor.from_string("28544B")
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_textbox(Inches(0.8), Inches(1), Inches(8), Inches(2)).text = "Second synthetic slide"
    slides = io.BytesIO()
    presentation.save(slides)

    checksums = {}
    for filename, content in (
        (excel_file, workbook.getvalue()),
        ("synthetic.pptx", slides.getvalue()),
    ):
        with (root / filename).open("xb") as stream:
            stream.write(content)
        checksums[filename] = hashlib.sha256(content).hexdigest()
    with (root / "fixtures.json").open("x", encoding="utf-8") as stream:
        json.dump(checksums, stream)


if __name__ == "__main__":
    main(Path(sys.argv[1]), *sys.argv[2:])
