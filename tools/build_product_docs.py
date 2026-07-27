from __future__ import annotations

import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / "docs"
OUTPUT_DIR = ROOT / "output" / "pdf"
FONT_DIR = Path("C:/Windows/Fonts")

DOCUMENTS = {
    "DEVELOPER_GUIDE_RU.md": "Khalil_Audit_Developer_Guide_RU.pdf",
    "ADMIN_GUIDE_RU.md": "Khalil_Audit_Administrator_Guide_RU.pdf",
}

ACCENT = colors.HexColor("#12877C")
ORANGE = colors.HexColor("#FF6F16")
INK = colors.HexColor("#18212D")
MUTED = colors.HexColor("#5D6B7A")
LINE = colors.HexColor("#D9E1E8")
PALE = colors.HexColor("#F4F7F9")


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont("Arial", str(FONT_DIR / "arial.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Bold", str(FONT_DIR / "arialbd.ttf")))
    pdfmetrics.registerFontFamily(
        "Arial",
        normal="Arial",
        bold="Arial-Bold",
    )


def linkify(text: str) -> str:
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    code_values: list[str] = []

    def hold_code(match: re.Match) -> str:
        code_values.append(match.group(1))
        return f"@@CODE{len(code_values) - 1}@@"

    text = re.sub(r"`([^`]+)`", hold_code, text)
    text = re.sub(
        r"(https?://[^\s<)]+)",
        lambda match: f'<link href="{match.group(1)}" color="#0878A4">{match.group(1)}</link>',
        text,
    )
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    for index, code in enumerate(code_values):
        text = text.replace(
            f"@@CODE{index}@@",
            f'<font name="Arial-Bold">{code}</font>',
        )
    return text


def styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "Title",
            parent=base["Title"],
            fontName="Arial-Bold",
            fontSize=26,
            leading=31,
            textColor=INK,
            spaceAfter=7 * mm,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            fontName="Arial",
            fontSize=12,
            leading=18,
            textColor=MUTED,
            spaceAfter=8 * mm,
        ),
        "h2": ParagraphStyle(
            "Heading2",
            fontName="Arial-Bold",
            fontSize=14,
            leading=17,
            textColor=INK,
            spaceBefore=3.8 * mm,
            spaceAfter=2.2 * mm,
        ),
        "body": ParagraphStyle(
            "Body",
            fontName="Arial",
            fontSize=8.6,
            leading=12.1,
            textColor=INK,
            spaceAfter=1.8 * mm,
        ),
        "bullet": ParagraphStyle(
            "Bullet",
            fontName="Arial",
            fontSize=8.6,
            leading=12.1,
            textColor=INK,
            leftIndent=5 * mm,
            firstLineIndent=-3.5 * mm,
            spaceAfter=1.1 * mm,
        ),
        "small": ParagraphStyle(
            "Small",
            fontName="Arial",
            fontSize=7.7,
            leading=10.5,
            textColor=MUTED,
        ),
        "table_head": ParagraphStyle(
            "TableHead",
            fontName="Arial-Bold",
            fontSize=7.3,
            leading=9.3,
            textColor=colors.white,
        ),
        "table": ParagraphStyle(
            "Table",
            fontName="Arial",
            fontSize=7.1,
            leading=9.2,
            textColor=INK,
        ),
    }


def header_footer(canvas, doc) -> None:
    canvas.saveState()
    width, height = A4
    canvas.setFillColor(ACCENT)
    canvas.rect(0, height - 4 * mm, width, 4 * mm, stroke=0, fill=1)
    canvas.setFont("Arial", 7.5)
    canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 10 * mm, "Khalil Audit System · by Ivan Rudoy")
    canvas.drawRightString(width - 18 * mm, 10 * mm, f"{doc.page}")
    canvas.restoreState()


def parse_table(lines: list[str], style_map: dict) -> Table:
    rows = []
    for line in lines:
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if all(re.fullmatch(r"[-: ]+", cell or "-") for cell in cells):
            continue
        rows.append(cells)
    width = 174 * mm
    columns = max(len(row) for row in rows)
    col_widths = [width / columns] * columns
    data = []
    for row_index, row in enumerate(rows):
        style_name = "table_head" if row_index == 0 else "table"
        data.append(
            [Paragraph(linkify(cell), style_map[style_name]) for cell in row]
        )
    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def markdown_story(path: Path) -> list:
    style_map = styles()
    lines = path.read_text(encoding="utf-8").splitlines()
    story = []
    first_heading = True
    index = 0
    while index < len(lines):
        raw = lines[index].rstrip()
        stripped = raw.strip()
        if not stripped:
            story.append(Spacer(1, 0.8 * mm))
            index += 1
            continue
        if stripped.startswith("|"):
            block = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                block.append(lines[index])
                index += 1
            story.append(parse_table(block, style_map))
            story.append(Spacer(1, 3 * mm))
            continue
        if stripped.startswith("# "):
            if first_heading:
                story.append(Spacer(1, 11 * mm))
                story.append(Paragraph(linkify(stripped[2:]), style_map["title"]))
                first_heading = False
            index += 1
            continue
        if stripped.startswith("## "):
            if not story or first_heading:
                story.append(Paragraph(linkify(stripped[3:]), style_map["title"]))
                first_heading = False
            else:
                story.append(Paragraph(linkify(stripped[3:]), style_map["h2"]))
            index += 1
            continue
        if stripped in {"### 12. Диагностика", "### 13. Если что-то не работает"}:
            story.append(PageBreak())
            story.append(Paragraph(linkify(stripped[4:]), style_map["h2"]))
            index += 1
            continue
        if stripped.startswith("### "):
            story.append(Paragraph(linkify(stripped[4:]), style_map["h2"]))
            index += 1
            continue
        if re.match(r"^\d+\.\s+", stripped):
            number, text = stripped.split(".", 1)
            story.append(
                Paragraph(f"<b>{number}.</b> {linkify(text.strip())}", style_map["bullet"])
            )
            index += 1
            continue
        if stripped.startswith("- "):
            story.append(Paragraph(f"• {linkify(stripped[2:])}", style_map["bullet"]))
            index += 1
            continue
        if stripped.startswith("Версия документа:") or stripped.startswith("Владелец") or stripped.startswith("Назначение:"):
            story.append(Paragraph(linkify(stripped), style_map["subtitle"]))
            index += 1
            continue
        story.append(Paragraph(linkify(stripped), style_map["body"]))
        index += 1
    return story


def build_pdf(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    doc = BaseDocTemplate(
        str(target),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=17 * mm,
        title=source.stem,
        author="Ivan Rudoy",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="main")
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=header_footer)])
    doc.build(markdown_story(source))


def main() -> None:
    register_fonts()
    for source_name, target_name in DOCUMENTS.items():
        build_pdf(SOURCE_DIR / source_name, OUTPUT_DIR / target_name)
        print(OUTPUT_DIR / target_name)


if __name__ == "__main__":
    main()
