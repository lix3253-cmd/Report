# -*- coding: utf-8 -*-
"""Claude Docs에서 내보낸 docx를 사내 보고서 서식으로 다시 만드는 변환기.

- 글꼴 '바탕', 본문 12pt, 표 11pt
- 번호(1. / 1) / ① / a) / i))는 Word 다단계 자동 번호로 변환
  → 다른 보고서에 붙여 넣으면 그 위치의 번호 흐름에 맞춰 다시 매겨짐
- Docs 표시용 줄(작성일 · 작성자)과 빈 표 행은 제거

사용: python3 scripts/docs_to_word.py <Docs 내보내기.docx> <출력.docx>
"""
import re
import sys

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

FONT = "바탕"
BODY_PT = 12
TABLE_PT = 11
HDR_FILL = "DDEBF7"
NUM_ID = 901

# (단계, 텍스트 앞 번호 패턴, Word 번호 형식, 번호 문자열, 들여쓰기 cm, 굵게)
LEVELS = [
    (0, r"^\d+\.\s*", "decimal", "%1.", 0.0, True),
    (1, r"^\d+\)\s*", "decimal", "%2)", 0.5, True),
    (2, r"^[①-⑳]\s*", "decimalEnclosedCircle", "%3", 1.2, False),
    (3, r"^[a-z]\)\s*", "lowerLetter", "%4)", 1.9, False),
    (4, r"^(?:i|ii|iii|iv|v|vi|vii|viii|ix|x)\)\s*", "lowerRoman", "%5)", 2.6, False),
]
HANG = 0.7
CAPTION = re.compile(r"^\[(표|그림)\. \d+\]")
BYLINE = re.compile(r"^[A-Z][a-z]{2} \d{1,2}, \d{4} · ")


def set_font(run, size, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = FONT
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), FONT)


def add_numbering(doc):
    """다단계 번호 정의를 numbering.xml에 추가."""
    lvls = []
    for lvl, _, fmt, text, left, bold in LEVELS:
        ind_left = int(Cm(left + HANG).twips)
        hang = int(Cm(HANG).twips)
        lvls.append(
            f'<w:lvl w:ilvl="{lvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{ind_left}" w:hanging="{hang}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}"/>'
            f'{"<w:b/>" if bold else ""}<w:sz w:val="{BODY_PT * 2}"/></w:rPr></w:lvl>'
        )
    numbering = doc.part.numbering_part.element
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{NUM_ID}">'
        f'<w:multiLevelType w:val="multilevel"/>{"".join(lvls)}</w:abstractNum>'
    )
    existing = numbering.findall(qn("w:abstractNum"))
    if existing:
        existing[-1].addnext(abstract)
    else:
        numbering.insert(0, abstract)
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{NUM_ID}"><w:abstractNumId w:val="{NUM_ID}"/></w:num>'
    ))


def numbered(doc, level, text):
    _, _, _, _, left, bold = LEVELS[level]
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl"); ilvl.set(qn("w:val"), str(level))
    nid = OxmlElement("w:numId"); nid.set(qn("w:val"), str(NUM_ID))
    numPr.append(ilvl); numPr.append(nid); pPr.append(numPr)
    pf = p.paragraph_format
    pf.left_indent = Cm(left + HANG)
    pf.first_line_indent = Cm(-HANG)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if level <= 1:
        pf.space_before = Pt(6)
    set_font(p.add_run(text), BODY_PT, bold)


def plain(doc, text, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), BODY_PT, bold)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def copy_table(doc, src):
    rows = [[c.text.strip() for c in r.cells] for r in src.rows]
    rows = [r for r in rows if any(r)]  # 빈 행 제거
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(rows):
        for j, val in enumerate(r):
            cell = t.rows[i].cells[j]
            cell.text = ""
            lines = val.split("\n")
            for k, line in enumerate(lines):
                p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (i == 0 or j == 0) else WD_ALIGN_PARAGRAPH.LEFT
                set_font(p.add_run(line), TABLE_PT, bold=(i == 0 or j == 0))
            if i == 0:
                shade(cell, HDR_FILL)
    # 첫 열(구분)은 좁게, 나머지는 본문 폭(16cm)을 나눔
    n = len(rows[0])
    widths = [2.4] + [(16.0 - 2.4) / (n - 1)] * (n - 1)
    t.autofit = False
    for col, w in zip(t.columns, widths):
        col.width = Cm(w)
    for r in t.rows:
        for c, w in zip(r.cells, widths):
            c.width = Cm(w)
    return t


def classify(line):
    s = line.strip().lstrip("  ")
    for lvl, pat, *_ in LEVELS:
        m = re.match(pat, s)
        if m:
            return lvl, s[m.end():].strip()
    return None, s


def build(src_path, out_path):
    src = docx.Document(src_path)
    doc = docx.Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY_PT)
    rf = normal.element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.space_after = Pt(2)
    add_numbering(doc)

    for el in src.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "tbl":
            copy_table(doc, Table(el, src))
        elif tag == "p":
            text = Paragraph(el, src).text
            for line in text.split("\n"):
                if not line.strip() or BYLINE.match(line.strip()):
                    continue
                if CAPTION.match(line.strip()):
                    plain(doc, line.strip(), bold=True)
                    continue
                lvl, body = classify(line)
                if lvl is None:
                    plain(doc, body)
                else:
                    numbered(doc, lvl, body)
    doc.save(out_path)
    print("saved:", out_path)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
