# -*- coding: utf-8 -*-
"""보고서 본문(md 또는 Claude Docs HTML 내보내기) → 붙여넣기용 docx.

사용:
    python3 docx_report.py <입력.md|입력.html> <출력.docx> [--widths "2.4,3.8,3.2,4.6,2.6;3.6,6.4,6.6"]

입력 규칙 (md)
- 한 줄 = 한 단락. 줄 앞의 번호 글자로 단계를 정한다.
    `1.` → 1단계, `1)` → 2단계, `①` → 3단계, `a)` → 4단계, `i)` → 5단계, `-` → 바로 위 번호 아래 불릿
  번호 글자는 떼고 Word 다단계 자동 번호로 넣는다 (붙여넣어도 번호 유지).
- `[표. N] 제목`, `[그림. N] 제목` 줄 → 표·그림 제목.
- `|`로 시작하는 줄 → 표. 첫 줄은 머리행, `|---|` 구분 줄은 무시.
    셀 값 `^^` = 위 칸과 세로 병합, `<<` = 왼쪽 칸과 가로 병합. 셀 안 줄바꿈은 `<br>`.
- `**굵게**` 지원. `#`로 시작하는 줄(문서 제목)과 빈 줄은 건너뛴다.
입력 규칙 (html): Claude Docs의 Export → HTML 파일. 단락, 표(rowspan/colspan), 굵게를 같은 방식으로 읽는다.

서식 (주간, 월간 공통)
- 본문: 바탕 12pt, 줄 간격 1줄 / 표: 바탕 11pt, 줄 간격 1줄 / 여백 2cm
- 붙여넣을 파일의 기본 서식에 흔들리지 않도록 줄 간격, 격자 맞춤 해제, 행 높이, 테두리를 직접 지정한다.
"""
import html as _html
import re
import sys
from html.parser import HTMLParser

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONT = "바탕"
BODY_SIZE = 12
TABLE_SIZE = 11
HDR_FILL = "DDEBF7"
EMOJI = set("🔴🟢🟡🟠⚪⚫")
EMOJI_FONT = "Segoe UI Emoji"
PAGE_WIDTH_CM = 17.0  # A4 21cm - 여백 2cm x 2

# 단계: (번호 형식, 표시 문자, 왼쪽 cm, 내어쓰기 cm, 굵게)
LEVELS = [
    ("decimal",               "%1.", 0.0, 0.7, True),   # 1.
    ("decimal",               "%2)", 0.5, 0.7, True),   # 1)
    ("decimalEnclosedCircle", "%3",  1.2, 0.7, False),  # ①
    ("lowerLetter",           "%4)", 1.9, 0.7, False),  # a)
    ("lowerRoman",            "%5)", 2.6, 0.8, False),  # i)
]
STYLE_NAMES = [f"보고 {i + 1}단계" for i in range(len(LEVELS))]
BULLET_STYLE = "보고 불릿"

MARKERS = [
    (re.compile(r"^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\)\s*"), 5),
    (re.compile(r"^[a-z]\)\s*"), 4),
    (re.compile(r"^[①-⑳]\s*"), 3),
    (re.compile(r"^\d+\)\s*"), 2),
    (re.compile(r"^\d+\.\s+"), 1),
    (re.compile(r"^[-–]\s+"), 0),  # 0 = 불릿
]
SPACES = " \t   "

PPR_AFTER_SNAP = ("w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
                  "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
                  "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
                  "w:rPr", "w:sectPr", "w:pPrChange")


# ── 공통 서식 ─────────────────────────────────────────────
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def _font(run, size, bold=False, face=FONT):
    run.font.name = face
    run.font.size = Pt(size)
    run.font.bold = bold
    rf = run._element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), face)


def _runs(p, text, size, bold=False):
    """**굵게**와 신호등 이모지를 run으로 나눠 넣는다."""
    for k, part in enumerate(text.split("**")):
        b = bold or (k % 2 == 1)
        buf = ""
        for ch in part:
            if ch in EMOJI:
                if buf:
                    _font(p.add_run(buf), size, b)
                    buf = ""
                _font(p.add_run(ch), size, False, EMOJI_FONT)
            else:
                buf += ch
        if buf:
            _font(p.add_run(buf), size, b)


def _fixed_spacing(p, before=0, after=0, reset_indent=True):
    """붙여넣을 파일의 기본 단락 서식(줄 간격 고정값, 문서 격자)을 받지 않도록 직접 지정."""
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if reset_indent:
        pf.left_indent = Cm(0)
        pf.first_line_indent = Cm(0)
    p._element.get_or_add_pPr().insert_element_before(
        _el("w:snapToGrid", **{"w:val": 0}), *PPR_AFTER_SNAP)


def new_document():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY_SIZE)
    rf = normal.element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.0)
    _setup_numbering(doc)
    return doc


def _setup_numbering(doc):
    """다단계 목록 정의 + 단계별 단락 스타일("보고 N단계"). 번호 글꼴도 직접 지정."""
    numbering = doc.part.numbering_part.element
    abs_id = num_id = 90
    absn = _el("w:abstractNum", **{"w:abstractNumId": abs_id})
    absn.append(_el("w:multiLevelType", **{"w:val": "multilevel"}))
    for i, (fmt, text, left, hang, bold) in enumerate(LEVELS):
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
        rpr = _el("w:rPr")
        rpr.append(_el("w:rFonts", **{"w:ascii": FONT, "w:hAnsi": FONT,
                                      "w:eastAsia": FONT, "w:cs": FONT}))
        if bold:
            rpr.append(_el("w:b"))
        rpr.append(_el("w:sz", **{"w:val": BODY_SIZE * 2}))
        lvl.append(rpr)
        absn.append(lvl)
    first = numbering.find(qn("w:num"))
    (first.addprevious(absn) if first is not None else numbering.append(absn))
    num = _el("w:num", **{"w:numId": num_id})
    num.append(_el("w:abstractNumId", **{"w:val": abs_id}))
    numbering.append(num)

    for i, (fmt, text, left, hang, bold) in enumerate(LEVELS):
        st = doc.styles.add_style(STYLE_NAMES[i], WD_STYLE_TYPE.PARAGRAPH)
        st.element.set(qn("w:styleId"), f"Report{i + 1}")
        st.base_style = doc.styles["Normal"]
        st.font.bold = bold
        st.font.size = Pt(BODY_SIZE)
        numpr = _el("w:numPr")
        if i:
            numpr.append(_el("w:ilvl", **{"w:val": i}))
        numpr.append(_el("w:numId", **{"w:val": num_id}))
        st.element.get_or_add_pPr().insert(0, numpr)
        st.paragraph_format.left_indent = Cm(left + hang)
        st.paragraph_format.first_line_indent = Cm(-hang)
    bs = doc.styles.add_style(BULLET_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    bs.element.set(qn("w:styleId"), "ReportBullet")
    bs.base_style = doc.styles["Normal"]


# ── 단락 ──────────────────────────────────────────────────
def item(doc, level, text, prev_level=1):
    """level 1~5 = 1. / 1) / ① / a) / i). level 0 = '-' 불릿(바로 위 번호 단계 아래)."""
    if level == 0:
        p = doc.add_paragraph(style=BULLET_STYLE)
        _, _, left, hang, _ = LEVELS[min(prev_level, len(LEVELS) - 1)]
        _fixed_spacing(p, after=2)
        p.paragraph_format.left_indent = Cm(left + 0.5)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _runs(p, "- " + text, BODY_SIZE)
        return p
    p = doc.add_paragraph(style=STYLE_NAMES[level - 1])
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _fixed_spacing(p, before=6 if level <= 2 else 0, after=2, reset_indent=False)
    _runs(p, text, BODY_SIZE, LEVELS[level - 1][4])
    return p


def plain(doc, text):
    p = doc.add_paragraph(style="Normal")
    _fixed_spacing(p, after=2)
    _runs(p, text, BODY_SIZE)
    return p


def caption(doc, text):
    p = doc.add_paragraph(style="Normal")
    _fixed_spacing(p, before=8, after=3)
    p.paragraph_format.keep_with_next = True
    _runs(p, text, BODY_SIZE, True)
    return p


# ── 표 ────────────────────────────────────────────────────
def _shade(cell):
    cell._element.get_or_add_tcPr().append(
        _el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": HDR_FILL}))


def _table_props(t):
    tblpr = t._element.tblPr
    borders = _el("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(f"w:{side}", **{"w:val": "single", "w:sz": 4,
                                           "w:space": 0, "w:color": "000000"}))
    tblpr.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
    mar = _el("w:tblCellMar")
    for side, w in (("top", 28), ("left", 85), ("bottom", 28), ("right", 85)):
        mar.append(_el(f"w:{side}", **{"w:w": w, "w:type": "dxa"}))
    tblpr.insert_element_before(mar, "w:tblLook")


def table(doc, rows, widths=None):
    """rows: [[{'paras': [텍스트, ...], 'rs': n, 'cs': n, 'th': bool}, ...], ...] (병합 칸은 생략된 구조)."""
    ncols = sum(c["cs"] for c in rows[0])
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _table_props(t)
    for row in t.rows:
        row.height = Cm(0.7)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    widths = widths or [PAGE_WIDTH_CM / ncols] * ncols
    for row in t.rows:
        for j, w in enumerate(widths[:ncols]):
            row.cells[j].width = Cm(w)
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
    for r, c, cell in spans:
        tc = t.cell(r, c)
        for k, text in enumerate(cell["paras"] or [""]):
            p = tc.paragraphs[0] if k == 0 else tc.add_paragraph()
            p.style = doc.styles["Normal"]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _fixed_spacing(p)
            if text:
                _runs(p, text, TABLE_SIZE, cell["th"])
        tc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if cell["th"]:
            _shade(tc)
        if cell["rs"] > 1 or cell["cs"] > 1:
            m = tc.merge(t.cell(r + cell["rs"] - 1, c + cell["cs"] - 1))
            for p in m.paragraphs[len(cell["paras"] or [""]):]:  # 병합으로 생긴 빈 단락 제거
                if not p.text:
                    p._element.getparent().remove(p._element)
    _fixed_spacing(doc.add_paragraph(style="Normal"))
    return t


# ── 입력 읽기 → 블록 [('p', text) | ('table', rows)] ──────────
def _grid_to_rows(grid):
    """md 표(문자열 2차원, ^^ / << 병합 표시)를 rowspan/colspan 구조로 바꾼다."""
    n, m = len(grid), max(len(r) for r in grid)
    grid = [r + [""] * (m - len(r)) for r in grid]
    rows = []
    for i in range(n):
        out = []
        for j in range(m):
            v = grid[i][j]
            if v in ("^^", "<<"):
                continue
            cs = 1
            while j + cs < m and grid[i][j + cs] == "<<":
                cs += 1
            rs = 1
            while i + rs < n and grid[i + rs][j] == "^^":
                rs += 1
            paras = [s.strip() for s in re.split(r"<br\s*/?>", v)]
            out.append({"paras": paras, "rs": rs, "cs": cs, "th": i == 0})
        rows.append(out)
    return rows


def read_md(path):
    blocks, grid = [], None
    for raw in open(path, encoding="utf-8").read().splitlines() + [""]:
        line = raw.strip(SPACES)
        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            grid = (grid or []) + [cells]
            continue
        if grid:
            blocks.append(("table", _grid_to_rows(grid)))
            grid = None
        if line and not line.startswith("#"):
            blocks.append(("p", line))
    return blocks


class _HtmlParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.blocks, self.buf, self.bold = [], None, 0
        self.table = self.row = self.cell = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("h1", "h2", "h3", "h4", "p", "li"):
            if self.cell is not None:
                self.buf = []
                self.cell["paras"].append(self.buf)
            else:
                self.kind, self.buf = tag, []
        elif tag in ("strong", "b"):
            self.bold += 1
            if self.buf is not None:
                self.buf.append("**")
        elif tag == "table":
            self.table = []
        elif tag == "tr":
            self.row = []
        elif tag in ("td", "th"):
            self.cell = {"paras": [], "rs": int(a.get("rowspan", 1)),
                         "cs": int(a.get("colspan", 1)), "th": tag == "th"}
            self.buf = []
            self.cell["paras"].append(self.buf)
        elif tag == "br" and self.buf is not None:
            if self.cell is not None:
                self.buf = []
                self.cell["paras"].append(self.buf)
            else:
                self.buf.append("\n")

    def handle_endtag(self, tag):
        if tag in ("h1", "h2", "h3", "h4", "p", "li") and self.cell is None and self.buf is not None:
            text = "".join(self.buf)
            if self.kind != "h1":
                for part in text.split("\n"):
                    self.blocks.append(("p", part.strip(SPACES)))
            self.buf = None
        elif tag in ("strong", "b"):
            self.bold -= 1
            if self.buf is not None:
                self.buf.append("**")
        elif tag in ("td", "th"):
            self.cell["paras"] = ["".join(b).strip(SPACES) for b in self.cell["paras"]]
            self.cell["paras"] = [s for s in self.cell["paras"] if s] or [""]
            self.row.append(self.cell)
            self.cell = self.buf = None
        elif tag == "tr":
            self.table.append(self.row)
        elif tag == "table":
            if self.table:
                self.table[0] = [dict(c, th=True) for c in self.table[0]]
            self.blocks.append(("table", self.table))
            self.table = None

    def handle_data(self, data):
        if self.buf is not None:
            self.buf.append(data)


def read_html(path):
    p = _HtmlParser()
    p.feed(open(path, encoding="utf-8").read())
    return [(k, v.replace("****", "")) if k == "p" else (k, v) for k, v in p.blocks]


# ── 조립 ──────────────────────────────────────────────────
def _split_marker(text):
    for rx, lvl in MARKERS:
        m = rx.match(text)
        if m:
            return lvl, text[m.end():]
    return None, text


def build(src, out, widths=None):
    blocks = read_html(src) if src.lower().endswith((".html", ".htm")) else read_md(src)
    doc = new_document()
    prev, n_table = 1, 0
    for kind, body in blocks:
        if kind == "table":
            w = widths[min(n_table, len(widths) - 1)] if widths else None
            table(doc, body, w)
            n_table += 1
            continue
        text = _html.unescape(body).strip(SPACES)
        if not text or re.match(r"^\d{4}-\d{2}-\d{2}", text) or re.match(r"^[A-Z][a-z]{2} \d{1,2}, \d{4}", text):
            continue  # 빈 줄, 작성일과 작성자 줄
        if re.match(r"^\[(표|그림)\.", text):
            caption(doc, text)
            continue
        lvl, rest = _split_marker(text)
        if lvl is None:
            plain(doc, text)
        else:
            item(doc, lvl, rest, prev)
            if lvl:
                prev = lvl
    doc.save(out)
    print("saved:", out)


if __name__ == "__main__":
    args = sys.argv[1:]
    widths = None
    if "--widths" in args:
        i = args.index("--widths")
        widths = [[float(x) for x in grp.split(",")] for grp in args[i + 1].split(";")]
        del args[i:i + 2]
    if len(args) != 2:
        sys.exit(__doc__)
    build(args[0], args[1], widths)
