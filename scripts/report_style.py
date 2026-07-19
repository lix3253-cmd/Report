# -*- coding: utf-8 -*-
"""
상무 월간보고 스타일 재사용 엔진.

글머리 4단계 고정:
  L1  5. 6.   (굵게, 대제목, 핵심어 노란 형광 가능)
  L2  1) 2)   (굵게, 들여쓰기)
  L3  ①~⑳    (일반, 내어쓰기)
  L4  a) b)   (일반, 더 깊은 들여쓰기 / '요약:' 등)
  -   대시      (단일 서술 항목)

글자체: 바탕체(ascii + w:eastAsia), 본문 양쪽 정렬.
표: 머리행 파랑 음영, 강조행 주황 음영.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "바탕체"          # 사내 표준이 다르면 이 값만 교체 (예: "맑은 고딕")
BODY_SIZE = 10.5
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

HDR_FILL = "DDEBF7"      # 머리행 파랑 음영
EMPH_FILL = "FCE4D6"     # 강조행 주황 음영


def _set_run_font(run, size=BODY_SIZE, bold=False, highlight=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)
    rfonts.set(qn("w:eastAsia"), FONT)
    if highlight:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def new_document():
    """바탕체·양쪽정렬 기본 스타일이 적용된 새 문서."""
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY_SIZE)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = 1.3
    pf.space_after = Pt(2)
    return doc


def _para(doc, left_cm, hanging_cm=0.0, space_before=0.0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(left_cm)
    if hanging_cm:
        pf.first_line_indent = Cm(-hanging_cm)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(space_before)
    return p


def _emit(p, marker, text, size, bold, highlight_key=None):
    """marker(번호기호) + 텍스트. highlight_key 문자열이 text에 있으면 그 부분만 형광."""
    r = p.add_run(marker)
    _set_run_font(r, size=size, bold=bold)
    if highlight_key and highlight_key in text:
        before, after = text.split(highlight_key, 1)
        if before:
            _set_run_font(p.add_run(before), size=size, bold=bold)
        _set_run_font(p.add_run(highlight_key), size=size, bold=bold, highlight=True)
        if after:
            _set_run_font(p.add_run(after), size=size, bold=bold)
    else:
        _set_run_font(p.add_run(text), size=size, bold=bold)
    return p


def h1(doc, num, text, highlight_key=None):
    p = _para(doc, 0.0, space_before=10)
    return _emit(p, f"{num}. ", text, size=13, bold=True, highlight_key=highlight_key)


def h2(doc, num, text):
    p = _para(doc, 0.5, space_before=6)
    return _emit(p, f"{num}) ", text, size=11.5, bold=True)


def h3(doc, idx, text):
    """idx: 1..20 -> ①..⑳"""
    marker = CIRCLED[idx - 1] + " "
    p = _para(doc, 1.2, hanging_cm=0.7)
    return _emit(p, marker, text, size=BODY_SIZE, bold=False)


def h4(doc, letter, text):
    p = _para(doc, 1.9, hanging_cm=0.7)
    return _emit(p, f"{letter}) ", text, size=BODY_SIZE, bold=False)


def dash(doc, text, left_cm=1.2):
    p = _para(doc, left_cm, hanging_cm=0.5)
    return _emit(p, "- ", text, size=BODY_SIZE, bold=False)


def label(doc, level, tag, body):
    """(라벨) 본문 — level: 3 또는 4 들여쓰기 사용, tag는 괄호 없이."""
    left = 1.2 if level == 3 else 1.9
    p = _para(doc, left, hanging_cm=0.7)
    r = p.add_run(f"({tag}) ")
    _set_run_font(r, bold=True)
    _set_run_font(p.add_run(body))
    return p


def _shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _cell_text(cell, text, bold=False, center=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    _set_run_font(p.add_run(str(text)), bold=bold)


def shaded_table(doc, headers, rows, emphasis_rows=None):
    """머리행 파랑 음영 + 선택 강조행 주황 음영. emphasis_rows: 0-based 데이터행 인덱스."""
    emphasis_rows = set(emphasis_rows or [])
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        _cell_text(c, h, bold=True, center=True)
        _shade(c, HDR_FILL)
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, val in enumerate(row):
            _cell_text(cells[j], val, center=(j != 1))
            if i in emphasis_rows:
                _shade(cells[j], EMPH_FILL)
    return t
