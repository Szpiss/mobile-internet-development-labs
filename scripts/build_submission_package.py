from __future__ import annotations

import os
import shutil
import zipfile
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
PACKAGE_DIR = ROOT / "移动互联网开发基础_四个实验提交包"
PDF_NAME = "2315302125_崔子霖_移动互联网开发基础实验一至四合并报告.pdf"
ZIP_NAME = "2315302125_崔子霖_移动互联网开发基础实验一至四提交包"
TMP_DIR = ROOT / "tmp" / "combined_report_assets"

EXPERIMENTS = [
    ("实验一", ROOT / "实验一" / "实验文档" / "2315302125 崔子霖1.docx"),
    ("实验二", ROOT / "实验二" / "实验文档" / "2315302125 崔子霖2.docx"),
    ("实验三", ROOT / "实验三" / "实验文档" / "2315302125 崔子霖3.docx"),
    ("实验四", ROOT / "实验四" / "实验文档" / "2315302125 崔子霖4.docx"),
]

CODE_COPY_RULES = {
    "实验一": ["index.html"],
    "实验二": ["index.html"],
    "实验三": ["index.html", "css", "img"],
    "实验四": ["index.html", "css", "js", "img"],
}


def register_fonts() -> str:
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/STHeiti Medium.ttc",
    ]
    for path in candidates:
        if Path(path).exists():
            pdfmetrics.registerFont(TTFont("ChineseMain", path))
            return "ChineseMain"
    return "Helvetica"


FONT_NAME = register_fonts()


def make_styles() -> dict[str, ParagraphStyle]:
    base = {
        "fontName": FONT_NAME,
        "wordWrap": "CJK",
    }
    return {
        "title": ParagraphStyle(
            "title",
            **base,
            fontSize=19,
            leading=25,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#1F467A"),
            spaceAfter=14,
        ),
        "section": ParagraphStyle(
            "section",
            **base,
            fontSize=15,
            leading=22,
            textColor=colors.HexColor("#245F4B"),
            spaceBefore=10,
            spaceAfter=7,
        ),
        "body": ParagraphStyle(
            "body",
            **base,
            fontSize=10.8,
            leading=17,
            firstLineIndent=22,
            alignment=TA_LEFT,
            spaceAfter=4,
        ),
        "caption": ParagraphStyle(
            "caption",
            **base,
            fontSize=9.8,
            leading=14,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#444444"),
            spaceBefore=3,
            spaceAfter=7,
        ),
        "divider": ParagraphStyle(
            "divider",
            **base,
            fontSize=21,
            leading=30,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#1F467A"),
            spaceAfter=18,
        ),
    }


STYLES = make_styles()


def escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br/>")
    )


def iter_block_items(parent: DocumentType) -> Iterable[DocxParagraph | DocxTable]:
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield DocxTable(child, parent)


def paragraph_image_ids(paragraph: DocxParagraph) -> list[str]:
    ids: list[str] = []
    for blip in paragraph._p.xpath(".//a:blip"):
        rid = blip.get(qn("r:embed"))
        if rid:
            ids.append(rid)
    return ids


def save_image(doc: DocumentType, rid: str, out_dir: Path, prefix: str, index: int) -> Path | None:
    part = doc.part.related_parts.get(rid)
    if not part:
        return None
    content_type = getattr(part, "content_type", "")
    if "image" not in content_type:
        return None
    ext = Path(part.partname).suffix.lower()
    if ext not in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}:
        return None
    out_path = out_dir / f"{prefix}_{index:03d}{ext}"
    out_path.write_bytes(part.blob)
    return out_path


def image_flowable(path: Path, max_width: float = 15.6 * cm, max_height: float = 18.5 * cm) -> Image | None:
    try:
        with PILImage.open(path) as img:
            width_px, height_px = img.size
    except Exception:
        return None
    if width_px <= 0 or height_px <= 0:
        return None
    ratio = min(max_width / width_px, max_height / height_px, 1.0)
    return Image(str(path), width=width_px * ratio, height=height_px * ratio, hAlign="CENTER")


def table_flowable(table: DocxTable) -> Table:
    data = []
    for row in table.rows:
        data.append([escape(cell.text.strip()) for cell in row.cells])
    if not data:
        data = [[""]]
    col_count = max(len(row) for row in data)
    usable_width = A4[0] - 4 * cm
    flow = Table(data, colWidths=[usable_width / col_count] * col_count, repeatRows=1)
    flow.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                ("FONTSIZE", (0, 0), (-1, -1), 9.4),
                ("LEADING", (0, 0), (-1, -1), 13),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#888888")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7F0F6")),
            ]
        )
    )
    return flow


def paragraph_style(text: str, previous_count: int) -> ParagraphStyle:
    if previous_count < 3 and ("实验报告" in text or "南京理工大学" in text):
        return STYLES["title"]
    if text.startswith(("一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、")):
        return STYLES["section"]
    if text.startswith("实验") and len(text) < 35:
        return STYLES["section"]
    return STYLES["body"]


def add_docx_to_story(story: list, label: str, path: Path, image_dir: Path) -> None:
    doc = Document(path)
    story.append(Paragraph(label, STYLES["divider"]))
    story.append(Spacer(1, 0.3 * cm))

    paragraph_count = 0
    image_index = 1
    for block in iter_block_items(doc):
        if isinstance(block, DocxParagraph):
            text = block.text.strip()
            image_ids = paragraph_image_ids(block)
            if text:
                style = paragraph_style(text, paragraph_count)
                story.append(Paragraph(escape(text), style))
                paragraph_count += 1
            for rid in image_ids:
                image_path = save_image(doc, rid, image_dir, label, image_index)
                image_index += 1
                if image_path:
                    flow = image_flowable(image_path)
                    if flow:
                        story.append(KeepTogether([flow, Spacer(1, 0.12 * cm)]))
        elif isinstance(block, DocxTable):
            story.append(table_flowable(block))
            story.append(Spacer(1, 0.25 * cm))


def add_page_numbers(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT_NAME, 9)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawCentredString(A4[0] / 2, 1.15 * cm, f"第 {doc.page} 页")
    canvas.restoreState()


def build_pdf() -> Path:
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    pdf_path = PACKAGE_DIR / PDF_NAME
    story: list = []

    for idx, (label, report_path) in enumerate(EXPERIMENTS):
        if idx:
            story.append(PageBreak())
        add_docx_to_story(story, label, report_path, TMP_DIR)

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=1.8 * cm,
        title="移动互联网开发基础实验一至四合并报告",
        author="崔子霖",
    )
    doc.build(story, onFirstPage=add_page_numbers, onLaterPages=add_page_numbers)
    return pdf_path


def copy_source_tree() -> Path:
    source_root = PACKAGE_DIR / "实验源码"
    if source_root.exists():
        shutil.rmtree(source_root)
    source_root.mkdir(parents=True, exist_ok=True)

    for exp_name, entries in CODE_COPY_RULES.items():
        dest = source_root / exp_name
        dest.mkdir(parents=True, exist_ok=True)
        src_root = ROOT / exp_name
        for entry in entries:
            src = src_root / entry
            if not src.exists():
                continue
            dst = dest / entry
            if src.is_dir():
                shutil.copytree(src, dst, ignore=shutil.ignore_patterns(".DS_Store", "__pycache__"))
            else:
                shutil.copy2(src, dst)

    readme = source_root / "README.txt"
    readme.write_text(
        "移动互联网开发基础实验一至四源码说明\n\n"
        "实验一：index.html，包含基础 JavaScript 函数页面。\n"
        "实验二：index.html，包含事件与表单验证页面。\n"
        "实验三：index.html、css、img，包含装修网站首页源码与运行图片。\n"
        "实验四：index.html、css、js、img，包含 Swiper 邀请函源码与运行图片。\n\n"
        "运行方式：使用浏览器分别打开各实验目录下的 index.html。\n",
        encoding="utf-8",
    )
    return source_root


def clean_package_dir() -> None:
    if PACKAGE_DIR.exists():
        shutil.rmtree(PACKAGE_DIR)
    PACKAGE_DIR.mkdir(parents=True, exist_ok=True)


def make_zip() -> Path:
    zip_base = ROOT / ZIP_NAME
    zip_path = Path(shutil.make_archive(str(zip_base), "zip", root_dir=PACKAGE_DIR.parent, base_dir=PACKAGE_DIR.name))
    return zip_path


def main() -> None:
    clean_package_dir()
    copy_source_tree()
    pdf_path = build_pdf()
    zip_path = make_zip()

    print(f"PDF: {pdf_path}")
    print(f"Package folder: {PACKAGE_DIR}")
    print(f"ZIP: {zip_path}")


if __name__ == "__main__":
    main()
