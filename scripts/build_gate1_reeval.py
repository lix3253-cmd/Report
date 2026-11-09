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


TABLE1_HEADERS = ["구분", "점검 항목 (목표)", "Gate #1 최초 평가",
                  "1차 재평가 (11/5)", "잔여 과제 조치 (2차 재평가 12/4)"]
TABLE1_ROWS = [
    ["양산 전개", "양산 공정 셀 문서 (셀라인 PFD 준비)", "🟢 통과",
     "셀(JC 파일럿 기준) 양산 공정 문서 작성 완료", LEFT],
    [UP, "양산 공정 모듈 문서 (모듈라인 PFD 준비)", "🟢 통과",
     "모듈(G12R 48FC PR) 양산 공정 문서 작성 완료", LEFT],
    [UP, "건축 및 유틸리티 (레이아웃 확보)", "🟢 통과",
     "PFD 및 폼팩터 기준 레이아웃 확보", LEFT],
    ["상업화 Risk", "Cell High-Risk (공정/설비 Risk 해소)", ["🟢 통과", "(완료 6)"],
     "Cell High-Risk 전체 완료", LEFT],
    [UP, "Module High-Risk (공정/설비 Risk 해소)", ["🔴 미달", "(완료 3, 미달 2)"],
     "🔴 미달: Light Soaker 공정 신뢰성 검증으로 진척, 2항목 잔여 (CTM 재현성, 모듈 설비 배치 사양)",
     ["CTM: 분류 기준 BKM 반영 및 검증",
      "모듈 설비 배치 사양: offset 측정 계획 수립 및 검증",
      "(12/4, 기술본부)"]],
    ["제품", "양산 폼팩터 (폼팩터 확정)", "🟢 통과",
     "Residential 향 G12R 48FC Portrait", LEFT],
    [UP, "IEC/UL 인증 (인증 획득)", "🟢 통과",
     "인증 획득 완료 (THA R&D M10 54 FC LS)", LEFT],
    ["Bankability", "LCOE 파라미터 (TCOE, LLB, 열화율 확보)", "🟢 통과",
     "TCOE -0.22%/K, LLB 93.1% 데이터 확보", LEFT],
    [UP, "글로벌 필드 테스트 (필드 실증 수행)", "🟢 통과",
     "R&D 주도 필드 실증 진행 중, '26년 4Q 열화율 결과 확보 예정", LEFT],
    ["JC 파일럿", "cell 효율 (≥ 28.0%)", "🔴 미달 (26.8%)", "🟢 **해결** (28.6%)", "–"],
    [UP, "FpG (≥ 95.0%)", "🔴 미달 (93.7%)", "🔴 미달: 50% (wet station 공정 이슈)",
     "11/13 공정 정상화 후 재측정"],
    [UP, "신뢰성 (LeTID 0%, 48h)", "🟢 통과", "±0% (48h), -0.7% (100hr)", LEFT],
    [UP, "열적 열화 (0%)", "🟢 통과", "-0.7%", LEFT],
    [UP, "모듈 출력 (480W)", "🟢 통과", "481W", LEFT],
    ["합계", LEFT, "19 / 23 통과 (미달 4)", "20 / 23 통과 (1건 해결, 미달 3)", "23 / 23 목표 (12/4)"],
]

TABLE2_HEADERS = ["구분", "Risk #1 CTM 재현성", "Risk #2 모듈 설비 배치 사양",
                  "JC 파일럿 cell 효율", "JC 파일럿 FpG"]
TABLE2_ROWS = [
    ["문제 현상", "동일 등급 cell 간 모듈 출력 편차, 관리 기준 부재",
     "모듈 라벨 출력 측정 방법 부재", "테스트 수준 목표 달성, BKM 전환 지연", LEFT],
    ["문제 원인", "cell 분류 기준 미정립",
     "1 sun 방식 안정화 시간 과다(>10시간)로 양산 적용 불가", "T05 설비 nozzle 막힘", LEFT],
    ["개선 계획", "PL, EL, HyR, IV 중 모듈 출력 연계 핵심 지표 선정 후 cell 분류 기준 반영",
     ["Light Soaker: 10 sun 고출력 사양 적용, Gsolar 社 설비 유효성 및 신뢰성 검토",
      "offset: Offline 측정기 15시간 MPPT 측정으로 라벨 출력 보정값 산정 (First Solar 벤치마킹)"],
     "신규 노즐 도입, CW44부터 공증착 BKM 대량 생산 검증", LEFT],
    ["실행 및 결과",
     "Toplimit Forward scan 분류 적용, Full Module 7개 CTM 오차 ±0.3% (기준 ±0.5%), 공정 시간 0.2초",
     ["Light Soaker: 2x2 미니 모듈 UV, LeTID 약 99%", "offset: 35매 측정, 적정 등급 74%"],
     "12,686장 효율 28.6% (기준 ≥ 28.0%)",
     "12,686장 FpG 50% (기준 ≥ 95.0%), 후면 pad 불량"],
    ["결론", "🔴 미달. 검증 수량 부족, BKM 미전환. 분류 기준 BKM 반영 후 재검증",
     "🔴 미달. 신뢰성 검증 완료, 라벨 측정 방안 미완료. 계획 재수립 및 추가 검증",
     "🟢 통과", "🔴 미달. 11/13 공정 정상화 후 재측정"],
    ["TRL", "TRL 5 (JC 파일럿 검증)",
     ["Light Soaker TRL 3 (R&D 검증)", "offset TRL 2 (아이디어 단계)"],
     "TRL 6 (JC 파일럿 BKM)", "TRL 5 (JC 파일럿 검증)"],
]


def build(path=OUT):
    doc = S.new_document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.0)
    setup_numbering(doc)

    item(doc, 2, "Gate #1 미달 항목 재평가 결과")
    item(doc, 3, "Gate #1 미달 항목 4건(상업화 Risk 2건, JC 파일럿 2건) 재평가를 11/5 실시")
    item(doc, 4, "변경 사유: 파일럿 라인 정전(CW41), 연차보수(CW42), 설비 PM(CW43)으로 cell 생산 일정 순연, "
                 "당초 10/30에서 변경")
    item(doc, 3, "재평가 결과, 미달 항목 4개 중 1개 통과, 3개 미달로 Gate #1 미달 상태 지속")
    item(doc, 3, "Gate #1 전 항목 통과 시까지 Gate #2 이행 보류. 12/4 2차 재평가에서 전 항목 통과 확보 후 "
                 "Gate #2 진행 추진")

    caption(doc, "[표. 1] Gate #1 점검 항목별 평가 경과 및 2차 재평가 계획")
    grid_table(doc, TABLE1_HEADERS, TABLE1_ROWS, [2.2, 4.2, 2.8, 4.4, 3.4],
               bold_rows=(len(TABLE1_ROWS) - 1,))

    item(doc, 3, "조치 결과에 따른 항목별 진척 및 잔여 과제")
    item(doc, 4, "1차 재평가에서 cell 효율 1건 해결, 대량 생산 조건에서 목표 효율 확보")
    item(doc, 4, "잔여 미달 3건 중 2건은 기술본부 기술 확보 과제, 1건은 공정 안정화 과제")
    item(doc, 5, "Risk #1 CTM 재현성: 분류 기준의 양산 적용 가능성은 확인, BKM 반영 및 검증 수량 확대 필요")
    item(doc, 5, "Risk #2 모듈 설비 배치 사양: 10 sun light soaking 유효성 및 신뢰성은 검증, 라벨 출력 측정 방안 미완료")
    item(doc, 5, "JC 파일럿 FpG: 일시적 공정 이슈로 미달, 11/13 공정 정상화 후 재측정 예정")

    caption(doc, "[표. 2] Gate #1 미달 항목별 원인 분석 및 조치 경과")
    grid_table(doc, TABLE2_HEADERS, TABLE2_ROWS, [2.0, 3.9, 4.5, 3.3, 3.3])

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build()
