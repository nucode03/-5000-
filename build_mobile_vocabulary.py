from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A5
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
CSV_PATH = ROOT / "build" / "oxford_5000_entries.csv"
PDF_PATH = ROOT / "output" / "Oxford_5000_B2-C1_Korean_Vocabulary_ver2_mobile.pdf"
FONT = Path(r"C:\Windows\Fonts\malgun.ttf")
FONT_BOLD = Path(r"C:\Windows\Fonts\malgunbd.ttf")


class MobileDoc(SimpleDocTemplate):
    def afterFlowable(self, flowable):
        name = getattr(flowable, "bookmark_name", None)
        if name:
            self.canv.bookmarkPage(name)
            self.canv.addOutlineEntry(flowable.outline_label, name, level=0, closed=False)


def load_entries():
    with CSV_PATH.open(encoding="utf-8-sig", newline="") as f:
        entries = list(csv.DictReader(f))
    if len(entries) != 2000 or any(not entry["meaning"].strip() for entry in entries):
        raise ValueError("Mobile source data must contain all 2,000 translated entries.")
    return entries


def make_styles():
    pdfmetrics.registerFont(TTFont("MalgunGothic", str(FONT)))
    pdfmetrics.registerFont(TTFont("MalgunGothicBold", str(FONT_BOLD)))
    base = getSampleStyleSheet()
    def style(name, font, size, leading, color, alignment=0, **extra):
        return ParagraphStyle(name, parent=base["Normal"], fontName=font, fontSize=size, leading=leading, textColor=colors.HexColor("#" + color), alignment=alignment, **extra)
    return {
        "cover_title": style("cover_title", "MalgunGothicBold", 28, 35, "163A5F", TA_CENTER),
        "cover_sub": style("cover_sub", "MalgunGothicBold", 18, 25, "2E74B5", TA_CENTER),
        "cover_note": style("cover_note", "MalgunGothic", 9.5, 14, "687385", TA_CENTER),
        "guide": style("guide", "MalgunGothic", 10.5, 16, "1F2937", 0, spaceAfter=8),
        "letter": style("letter", "MalgunGothicBold", 23, 28, "163A5F", 0, spaceAfter=3),
        "letter_count": style("letter_count", "MalgunGothic", 9.5, 12, "687385", 0, spaceAfter=7),
        "word": style("word", "MalgunGothicBold", 16.5, 21, "163A5F"),
        "pos": style("pos", "MalgunGothicBold", 10.5, 14, "4F5D75", TA_RIGHT),
        "meaning": style("meaning", "MalgunGothic", 14, 19, "1F2937"),
    }


POS_LABEL = {
    "n.": "명사", "v.": "동사", "adj.": "형용사", "adv.": "부사",
    "prep.": "전치사", "conj.": "접속사", "pron.": "대명사", "det.": "한정사",
    "exclam.": "감탄사", "number": "수사",
}


def korean_pos(detail: str) -> str:
    import re
    clean = re.sub(r"\b(?:B1|B2|C1)\b", "", detail)
    tokens = re.findall(r"n\.|v\.|adj\.|adv\.|prep\.|conj\.|pron\.|det\.|exclam\.|number", clean)
    return " / ".join(POS_LABEL[token] for token in tokens)


def para(text, style):
    return Paragraph(escape(text).replace("\n", "<br/>"), style)


def card(entry, styles):
    data = [
        [para(entry["word"], styles["word"]), para(korean_pos(entry["source_detail"]), styles["pos"])],
        [para(entry["meaning"], styles["meaning"]), ""],
    ]
    table = Table(data, colWidths=[88 * mm, 36 * mm], hAlign="LEFT")
    table.setStyle(TableStyle([
        ("SPAN", (0, 1), (1, 1)),
        ("BOX", (0, 0), (-1, -1), 0.65, colors.HexColor("#B9CBE0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EDF4FB")),
        ("LINEBELOW", (0, 0), (-1, 0), 0.35, colors.HexColor("#D7E4F1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 5), ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
        ("TOPPADDING", (0, 1), (-1, 1), 5), ("BOTTOMPADDING", (0, 1), (-1, 1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    return KeepTogether([table, Spacer(1, 3.2 * mm)])


def footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#D3DAE4")); canvas.setLineWidth(0.35)
    canvas.line(12 * mm, 202 * mm, 136 * mm, 202 * mm)
    canvas.setFillColor(colors.HexColor("#687385")); canvas.setFont("MalgunGothic", 7)
    canvas.drawString(12 * mm, 205 * mm, "Oxford 5000 B2-C1 | Mobile Edition")
    canvas.drawRightString(136 * mm, 7 * mm, f"Page {doc.page}")
    canvas.restoreState()


def make_pdf(entries):
    styles = make_styles()
    document = MobileDoc(
        str(PDF_PATH), pagesize=A5,
        leftMargin=12 * mm, rightMargin=12 * mm, topMargin=15 * mm, bottomMargin=14 * mm,
        title="Oxford 5000 B2-C1 한글 단어장 ver2 모바일", author="Codex",
    )
    story = [
        Spacer(1, 54 * mm), para("Oxford 5000", styles["cover_title"]), para("B2-C1 한글 단어장", styles["cover_sub"]),
        Spacer(1, 8 * mm), para("모바일 큰 글씨 카드형", styles["cover_note"]),
        Spacer(1, 63 * mm), para("영단어 · 한글 품사 · 대표 한글 뜻\n2,000개 항목 | American English", styles["cover_note"]),
        PageBreak(),
        para("모바일 버전 사용법", styles["letter"]),
        para("한 화면에서 쉽게 읽도록 큰 글씨 카드형으로 구성했습니다.", styles["guide"]),
        para("단어는 알파벳순이며, PDF 목차에서 알파벳을 선택해 빠르게 이동할 수 있습니다.", styles["guide"]),
        para("각 카드는 영단어, 한글 품사, 대표 한글 뜻 순서입니다.", styles["guide"]),
        PageBreak(),
    ]
    groups = defaultdict(list)
    for entry in entries:
        groups[entry["word"][0].upper()].append(entry)
    for index, letter in enumerate(sorted(groups)):
        if index:
            story.append(PageBreak())
        heading = para(letter, styles["letter"])
        heading.bookmark_name = f"letter_{letter}"
        heading.outline_label = letter
        story += [heading, para(f"{len(groups[letter])}개 단어", styles["letter_count"])]
        for entry in groups[letter]:
            story.append(card(entry, styles))
    document.build(story, onFirstPage=footer, onLaterPages=footer)


if __name__ == "__main__":
    make_pdf(load_entries())
    print(PDF_PATH)
