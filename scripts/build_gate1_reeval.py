# -*- coding: utf-8 -*-
"""11월 월간보고 — Gate #1 미달 항목 재평가 절 docx 생성기.

다른 보고서에 붙여넣기용. 번호(1) / ① / a) / i))는 문자로 입력하지 않고
Word 자동 번호(다단계 목록)로 넣는다. 각 단계는 단락 스타일
("보고 1단계" ~ "보고 5단계")에 연결되어 있어, 붙여넣을 때 대상 문서의
목록/스타일을 따르게 할 수 있다.

글꼴: 바탕, 11pt (본문과 표 모두).
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

S.FONT = "바탕"
SIZE = 11
S.BODY_SIZE = SIZE

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "11월_월간보고_Gate1_재평가.docx"))

# 단계: (번호 형식, 표시 문자, 왼쪽 들여쓰기 cm, 내어쓰기 cm, 굵게)
LEVELS = [
    ("decimal",               "%1.", 0.0, 0.6, True),
    ("decimal",               "%2)", 0.5, 0.6, True),
    ("decimalEnclosedCircle", "%3",  1.2, 0.6, False),
    ("lowerLetter",           "%4)", 1.9, 0.6, False),
    ("lowerRoman",            "%5)", 2.6, 0.7, False),
]
STYLE_NAMES = [f"보고 {i + 1}단계" for i in range(len(LEVELS))]
EMOJI = ("🔴", "🟢")
EMOJI_FONT = "Segoe UI Emoji"
UP = "^"      # 위 칸과 세로 병합
LEFT = None   # 왼쪽 칸과 가로 병합


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def setup_numbering(doc):
    """다단계 목록 정의 + 단계별 단락 스타일 생성."""
    numbering = doc.part.numbering_part.element
    abs_id, num_id = 90, 90
    absn = _el("w:abstractNum", **{"w:abstractNumId": abs_id})
    absn.append(_el("w:multiLevelType", **{"w:val": "multilevel"}))
    for i, (fmt, text, left, hang, _b) in enumerate(LEVELS):
        lvl = _el("w:lvl", **{"w:ilvl": i})
        lvl.append(_el("w:start", **{"w:val": 1}))
        lvl.append(_el("w:numFmt", **{"w:val": fmt}))
        lvl.append(_el("w:pStyle", **{"w:val": f"Report{i + 1}"}))
        lvl.append(_el("w:lvlText", **{"w:val": text}))
        lvl.append(_el("w:lvlJc", **{"w:val": "left"}))
        ppr = _el("w:pPr")
        ppr.append(_el("w:ind", **{"w:left": int((left + hang) * 567),
                                   "w:hanging": int(hang * 567)}))
        lvl.append(ppr)
        absn.append(lvl)
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(absn)
    else:
        numbering.append(absn)
    num = _el("w:num", **{"w:numId": num_id})
    num.append(_el("w:abstractNumId", **{"w:val": abs_id}))
    numbering.append(num)

    styles = doc.styles
    for i, (fmt, text, left, hang, bold) in enumerate(LEVELS):
        st = styles.add_style(STYLE_NAMES[i], WD_STYLE_TYPE.PARAGRAPH)
        st.element.set(qn("w:styleId"), f"Report{i + 1}")
        st.base_style = styles["Normal"]
        st.font.bold = bold
        st.font.size = Pt(SIZE)
        ppr = st.element.get_or_add_pPr()
        numpr = _el("w:numPr")
        if i:
            numpr.append(_el("w:ilvl", **{"w:val": i}))
        numpr.append(_el("w:numId", **{"w:val": num_id}))
        ppr.insert(0, numpr)
        st.paragraph_format.left_indent = Cm(left + hang)
        st.paragraph_format.first_line_indent = Cm(-hang)
        if i <= 1:
            st.paragraph_format.space_before = Pt(6)


def _emoji_run(p, ch):
    r = p.add_run(ch)
    S._set_run_font(r, size=SIZE)
    rf = r._element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), EMOJI_FONT)


def _runs(p, text, bold=False):
    """**굵게** 표기와 신호등 이모지를 처리해 run으로 넣는다."""
    for k, part in enumerate(text.split("**")):
        b = bold or (k % 2 == 1)
        buf = ""
        for ch in part:
            if ch in EMOJI:
                if buf:
                    S._set_run_font(p.add_run(buf), size=SIZE, bold=b)
                    buf = ""
                _emoji_run(p, ch)
            else:
                buf += ch
        if buf:
            S._set_run_font(p.add_run(buf), size=SIZE, bold=b)


def item(doc, level, text):
    """level: 1=1., 2=1), 3=①, 4=a), 5=i). 번호는 스타일의 자동 번호."""
    p = doc.add_paragraph(style=STYLE_NAMES[level - 1])
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _runs(p, text, LEVELS[level - 1][4])
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    S._set_run_font(p.add_run(text), size=SIZE, bold=True)
    return p


def _fill(cell, lines, bold=False, center=False):
    """lines: 문자열 또는 문자열 리스트(셀 안 문단)."""
    if isinstance(lines, str):
        lines = [lines]
    cell.text = ""
    for k, line in enumerate(lines):
        p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.space_after = Pt(0)
        if line:
            _runs(p, line, bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def grid_table(doc, headers, rows, widths_cm, bold_rows=(), center_cols=(0,)):
    """rows: 셀 값 리스트. LEFT(None)=왼쪽 칸과 가로 병합, UP('^')=위 칸과 세로 병합."""
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        _fill(c, h, bold=True, center=True)
        S._shade(c, S.HDR_FILL)
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, val in enumerate(row):
            if val is LEFT or val == UP:
                continue
            _fill(cells[j], val, bold=(i in bold_rows), center=(j in center_cols))
        S._shade(cells[0], S.HDR_FILL)
    t.autofit = False
    for j, w in enumerate(widths_cm):
        t.columns[j].width = Cm(w)
    for row in t.rows:
        for j, w in enumerate(widths_cm):
            row.cells[j].width = Cm(w)
    # 세로 병합
    for j in range(len(headers)):
        i = 0
        while i < len(rows):
            k = i + 1
            while k < len(rows) and rows[k][j] == UP:
                k += 1
            if k - i > 1:
                t.cell(i + 1, j).merge(t.cell(k, j))
            i = k
    # 가로 병합
    for i, row in enumerate(rows):
        j = 0
        while j < len(row):
            k = j + 1
            while k < len(row) and row[k] is LEFT:
                k += 1
            if k - j > 1:
                t.cell(i + 1, j).merge(t.cell(i + 1, k - 1))
            j = k
    return t




# ── Claude Docs HTML 내보내기 → docx ─────────────────────────────
import re
import html as _html
from html.parser import HTMLParser

SRC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "content", "11월_Gate1_재평가.html"))
TABLE_WIDTHS = [[2.6, 4.0, 2.6, 4.4, 3.4], [2.0, 3.9, 4.5, 3.3, 3.3]]

MARKERS = [
    (re.compile(r"^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\)\s*"), 5),
    (re.compile(r"^[a-z]\)\s*"), 4),
    (re.compile(r"^[①-⑳]\s*"), 3),
    (re.compile(r"^\d+\)\s*"), 2),
    (re.compile(r"^\d+\.\s+"), 1),
]


class _DocParser(HTMLParser):
    """블록 목록: ('h1'|'h2'|'p', runs) 또는 ('table', rows).
    runs = [(text, bold)], rows = [[{'paras': [runs], 'rs', 'cs', 'th'}]]"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.blocks, self.runs, self.bold = [], None, 0
        self.table = self.row = self.cell = None
        self.kind = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("h1", "h2", "p"):
            if self.cell is not None:
                self.runs = []
                self.cell["paras"].append(self.runs)
            else:
                self.kind, self.runs = tag, []
        elif tag == "strong":
            self.bold += 1
        elif tag == "table":
            self.table = []
        elif tag == "tr":
            self.row = []
        elif tag in ("td", "th"):
            self.cell = {"paras": [], "rs": int(a.get("rowspan", 1)),
                         "cs": int(a.get("colspan", 1)), "th": tag == "th"}
            self.runs = []
            self.cell["paras"].append(self.runs)
        elif tag == "br" and self.runs is not None:
            self.runs.append(("\n", False))

    def handle_endtag(self, tag):
        if tag in ("h1", "h2", "p") and self.cell is None and self.runs is not None:
            self.blocks.append((self.kind, self.runs))
            self.runs = None
        elif tag == "strong":
            self.bold -= 1
        elif tag in ("td", "th"):
            self.cell["paras"] = [r for r in self.cell["paras"] if r] or [[]]
            self.row.append(self.cell)
            self.cell, self.runs = None, None
        elif tag == "tr":
            self.table.append(self.row)
        elif tag == "table":
            self.blocks.append(("table", self.table))
            self.table = None

    def handle_data(self, data):
        if self.runs is not None:
            self.runs.append((data, self.bold > 0))


def _put_runs(p, runs, bold=False):
    for text, b in runs:
        _runs(p, text, bold or b)


def _strip_marker(runs):
    """첫 run 앞의 공백과 번호 기호를 떼고 단계를 돌려준다."""
    if not runs:
        return None, runs
    first = runs[0][0].lstrip("   \t")
    for rx, lvl in MARKERS:
        m = rx.match(first)
        if m:
            return lvl, [(first[m.end():], runs[0][1])] + runs[1:]
    return None, [(first, runs[0][1])] + runs[1:]


def html_table(doc, rows, widths):
    ncols = sum(c["cs"] for c in rows[0])
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    taken, spans = set(), []
    for r, row in enumerate(rows):
        c = 0
        for cell in row:
            while (r, c) in taken:
                c += 1
            for dr in range(cell["rs"]):
                for dc in range(cell["cs"]):
                    taken.add((r + dr, c + dc))
            spans.append((r, c, cell))
            c += cell["cs"]
    for j, w in enumerate(widths[:ncols]):
        t.columns[j].width = Cm(w)
        for row in t.rows:
            row.cells[j].width = Cm(w)
    for r, c, cell in spans:
        tc = t.cell(r, c)
        tc.text = ""
        for k, runs in enumerate(cell["paras"]):
            p = tc.paragraphs[0] if k == 0 else tc.add_paragraph()
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.space_after = Pt(0)
            center = cell["th"] or c == 0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
            _put_runs(p, runs, bold=cell["th"])
        tc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if cell["th"] or c == 0:
            S._shade(tc, S.HDR_FILL)
        if cell["rs"] > 1 or cell["cs"] > 1:
            tc.merge(t.cell(r + cell["rs"] - 1, c + cell["cs"] - 1))
    return t


def build(src=SRC, path=OUT):
    parser = _DocParser()
    parser.feed(open(src, encoding="utf-8").read())
    doc = S.new_document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.0)
    setup_numbering(doc)
    n_table = 0
    for kind, body in parser.blocks:
        if kind == "h1":
            continue
        if kind == "table":
            html_table(doc, body, TABLE_WIDTHS[min(n_table, len(TABLE_WIDTHS) - 1)])
            n_table += 1
            continue
        text = "".join(t for t, _ in body).strip()
        if not text or re.match(r"^\d{4}-\d{2}-\d{2}", text):
            continue  # 빈 줄, 작성일·작성자 줄
        if text.startswith("[표.") or text.startswith("[그림."):
            caption(doc, text)
            continue
        lvl, runs = _strip_marker(body)
        if lvl is None:
            p = doc.add_paragraph()
            _put_runs(p, runs)
            continue
        p = doc.add_paragraph(style=STYLE_NAMES[lvl - 1])
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _put_runs(p, runs, bold=LEVELS[lvl - 1][4])
    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build(*sys.argv[1:2])
