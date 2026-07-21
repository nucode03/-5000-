from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
CSV_PATH = ROOT / "build" / "oxford_5000_entries.csv"
PDF_PATH = ROOT / "output" / "Oxford_5000_B2-C1_Korean_Vocabulary.pdf"
FONT = Path(r"C:\Windows\Fonts\malgun.ttf")
FONT_BOLD = Path(r"C:\Windows\Fonts\malgunbd.ttf")


def footer(canvas, doc) -> None:
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#D3DAE4"))
    canvas.setLineWidth(0.4)
    canvas.line(14 * mm, 289 * mm, 196 * mm, 289 * mm)
    canvas.setFillColor(colors.HexColor("#687385"))
    canvas.setFont("MalgunGothic", 7.5)
    canvas.drawString(14 * mm, 292 * mm, "Oxford 5000 B2-C1 | Korean Vocabulary Book")
    canvas.drawRightString(196 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def make_styles():
    pdfmetrics.registerFont(TTFont("MalgunGothic", str(FONT)))
    pdfmetrics.registerFont(TTFont("MalgunGothicBold", str(FONT_BOLD)))
    base = getSampleStyleSheet()
    return {
        "cover_title": ParagraphStyle("cover_title", parent=base["Title"], fontName="MalgunGothicBold", fontSize=30, leading=38, textColor=colors.HexColor("#163A5F"), alignment=TA_CENTER, spaceAfter=4),
        "cover_subtitle": ParagraphStyle("cover_subtitle", parent=base["Normal"], fontName="MalgunGothicBold", fontSize=21, leading=28, textColor=colors.HexColor("#2E74B5"), alignment=TA_CENTER),
        "cover_note": ParagraphStyle("cover_note", parent=base["Normal"], fontName="MalgunGothic", fontSize=9.5, leading=15, textColor=colors.HexColor("#687385"), alignment=TA_CENTER),
        "h1": ParagraphStyle("h1", parent=base["Heading1"], fontName="MalgunGothicBold", fontSize=19, leading=24, textColor=colors.HexColor("#163A5F"), spaceBefore=0, spaceAfter=5),
        "body": ParagraphStyle("body", parent=base["Normal"], fontName="MalgunGothic", fontSize=10, leading=15, textColor=colors.HexColor("#1F2937"), spaceAfter=7),
        "small": ParagraphStyle("small", parent=base["Normal"], fontName="MalgunGothic", fontSize=8.5, leading=11, textColor=colors.HexColor("#687385")),
        "word": ParagraphStyle("word", parent=base["Normal"], fontName="MalgunGothicBold", fontSize=8.9, leading=11, textColor=colors.HexColor("#1F2937")),
        "cell": ParagraphStyle("cell", parent=base["Normal"], fontName="MalgunGothic", fontSize=8.7, leading=11, textColor=colors.HexColor("#1F2937")),
        "center": ParagraphStyle("center", parent=base["Normal"], fontName="MalgunGothic", fontSize=8.4, leading=10.5, textColor=colors.HexColor("#394B63"), alignment=TA_CENTER),
        "header": ParagraphStyle("header", parent=base["Normal"], fontName="MalgunGothicBold", fontSize=8.8, leading=11, textColor=colors.HexColor("#163A5F"), alignment=TA_CENTER),
        "index": ParagraphStyle("index", parent=base["Normal"], fontName="MalgunGothicBold", fontSize=8.4, leading=12, textColor=colors.HexColor("#163A5F"), alignment=TA_CENTER),
    }


def paragraph(text: str, style) -> Paragraph:
    return Paragraph(escape(text).replace("\n", "<br/>"), style)


def load_entries():
    with CSV_PATH.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def make_pdf(entries) -> None:
    styles = make_styles()
    doc = SimpleDocTemplate(
        str(PDF_PATH), pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm, topMargin=18 * mm, bottomMargin=16 * mm,
        title="Oxford 5000 B2-C1 한글 단어장", author="Codex",
    )
    story = []
    story.extend([
        Spacer(1, 78 * mm),
        paragraph("Oxford 5000", styles["cover_title"]),
        paragraph("B2-C1 한글 단어장", styles["cover_subtitle"]),
        Spacer(1, 9 * mm),
        paragraph("American English | 2,000개 고급 학습 단어", styles["cover_note"]),
        Spacer(1, 100 * mm),
        paragraph("출처: The Oxford 5000™ (American English)<br/>원문 목록을 바탕으로 한국어 대표 뜻을 덧붙여 학습용으로 재구성함", styles["cover_note"]),
        PageBreak(),
        paragraph("이 단어장 사용법", styles["h1"]),
    ])
    guide = [
        "각 항목은 영단어, 원문의 품사, CEFR 레벨, 대표 한글 뜻 순서입니다.",
        "단어는 알파벳순이며, 매 알파벳을 새 페이지에서 시작합니다.",
        "복수 품사와 동형어는 원문 표기를 보존했습니다. 문맥에 따라 뜻은 달라질 수 있습니다.",
        "레벨은 원문 표기를 그대로 따릅니다. 원문에 포함된 B1 표기도 수정하지 않았습니다.",
    ]
    for text in guide:
        story.append(paragraph("• " + text, styles["body"]))

    groups = defaultdict(list)
    for entry in entries:
        groups[entry["word"][0].upper()].append(entry)
    letters = sorted(groups)
    story.append(Spacer(1, 4 * mm))
    story.append(paragraph("알파벳 색인", styles["h1"]))
    index_data = []
    for i in range(0, len(letters), 4):
        row = []
        for letter in letters[i:i + 4]:
            words = groups[letter]
            row.append(paragraph(f"{letter}<br/>{words[0]['word']} - {words[-1]['word']}<br/>{len(words)}개", styles["index"]))
        while len(row) < 4:
            row.append("")
        index_data.append(row)
    index = Table(index_data, colWidths=[45.5 * mm] * 4)
    index.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C5D3E1")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EAF0F7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.extend([index, PageBreak()])

    widths = [52 * mm, 36 * mm, 18 * mm, 76 * mm]
    for i, letter in enumerate(letters):
        if i:
            story.append(PageBreak())
        story.append(paragraph(letter, styles["h1"]))
        story.append(paragraph(f"{len(groups[letter])}개 단어", styles["small"]))
        data = [[paragraph("영단어", styles["header"]), paragraph("품사", styles["header"]), paragraph("레벨", styles["header"]), paragraph("대표 한글 뜻", styles["header"])]]
        for entry in groups[letter]:
            data.append([
                paragraph(entry["word"], styles["word"]),
                paragraph(entry["pos"], styles["cell"]),
                paragraph(entry["level"], styles["center"]),
                paragraph(entry["meaning"], styles["cell"]),
            ])
        table = Table(data, colWidths=widths, repeatRows=1, splitByRow=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C7D2DE")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E5F2")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (2, 1), (2, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)

    doc.build(story, onFirstPage=footer, onLaterPages=footer)


if __name__ == "__main__":
    make_pdf(load_entries())
    print(PDF_PATH)
