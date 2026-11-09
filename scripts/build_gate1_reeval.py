# -*- coding: utf-8 -*-
"""11월 월간보고 — Gate #1 미달 항목 재평가 절 docx 생성기.

다른 보고서에 붙여넣기용. 번호(1) / ① / a) / i))는 문자로 입력하지 않고
Word 자동 번호(다단계 목록)로 넣는다. 각 단계는 단락 스타일
("보고 1단계" ~ "보고 5단계")에 연결되어 있어, 붙여넣을 때 대상 문서의
목록/스타일을 따르게 할 수 있다.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "11월_월간보고_Gate1_재평가.docx"))

# 단계: (번호 형식, 표시 문자, 왼쪽 들여쓰기 cm, 내어쓰기 cm, 굵게, 글자 크기)
LEVELS = [
    ("decimal",               "%1.", 0.0, 0.6, True, 13),
    ("decimal",               "%2)", 0.5, 0.6, True, 11.5),
    ("decimalEnclosedCircle", "%3",  1.2, 0.6, False, S.BODY_SIZE),
    ("lowerLetter",           "%4)", 1.9, 0.6, False, S.BODY_SIZE),
    ("lowerRoman",            "%5)", 2.6, 0.7, False, S.BODY_SIZE),
]
STYLE_NAMES = [f"보고 {i + 1}단계" for i in range(len(LEVELS))]
EMOJI = ("🔴", "🟢")
EMOJI_FONT = "Segoe UI Emoji"
TABLE_SIZE = 9


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
    for i, (fmt, text, left, hang, _b, _s) in enumerate(LEVELS):
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
    for i, (fmt, text, left, hang, bold, size) in enumerate(LEVELS):
        st = styles.add_style(STYLE_NAMES[i], WD_STYLE_TYPE.PARAGRAPH)
        st.element.set(qn("w:styleId"), f"Report{i + 1}")
        st.base_style = styles["Normal"]
        st.font.bold = bold
        st.font.size = Pt(size)
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


def _runs(p, text, size, bold=False):
    """텍스트를 넣되 신호등 이모지는 컬러 이모지 글꼴로 따로 넣는다."""
    buf = ""
    for ch in text:
        if ch in EMOJI:
            if buf:
                S._set_run_font(p.add_run(buf), size=size, bold=bold)
                buf = ""
            r = p.add_run(ch)
            S._set_run_font(r, size=size)
            rf = r._element.get_or_add_rPr().get_or_add_rFonts()
            for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rf.set(qn(a), EMOJI_FONT)
        else:
            buf += ch
    if buf:
        S._set_run_font(p.add_run(buf), size=size, bold=bold)


def item(doc, level, text):
    """level: 1=1., 2=1), 3=①, 4=a), 5=i). 번호는 스타일의 자동 번호."""
    p = doc.add_paragraph(style=STYLE_NAMES[level - 1])
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _runs(p, text, LEVELS[level - 1][5], LEVELS[level - 1][4])
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    S._set_run_font(p.add_run(text), bold=True)
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
        if line:
            _runs(p, line, TABLE_SIZE, bold)


def grid_table(doc, headers, rows, widths_cm, bold_rows=(), center_cols=(0,)):
    """rows: 각 행은 셀 값 리스트. 값이 None이면 왼쪽 셀과 가로 병합."""
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
            if val is None:
                continue
            _fill(cells[j], val, bold=(i in bold_rows), center=(j in center_cols))
        if j == 0 or True:
            S._shade(cells[0], S.HDR_FILL)
    t.autofit = False
    for j, w in enumerate(widths_cm):
        t.columns[j].width = Cm(w)
    for row in t.rows:
        for j, w in enumerate(widths_cm):
            row.cells[j].width = Cm(w)
    # 가로 병합 (값이 None인 칸을 왼쪽 칸에 합침)
    for i, row in enumerate(rows):
        r = t.rows[i + 1]
        j = 0
        while j < len(row):
            k = j + 1
            while k < len(row) and row[k] is None:
                k += 1
            if k - j > 1:
                r.cells[j].merge(r.cells[k - 1])
            j = k
    return t


TABLE1_HEADERS = ["항목", "Gate #1 최초 평가", "1차 재평가 (11/5)", "2차 재평가 (12/5)까지 조치"]
TABLE1_ROWS = [
    ["통과 / 전체", "18 / 22", "19 / 22", "22 / 22 (목표)"],
    ["Risk #1 CTM 재현성", "🔴 미달",
     "🔴 미달: 오차 ±0.3%로 기준(±0.5%) 이내, 검증 7개로 수량 부족",
     "신규 cell 분류 기준 BKM 반영 및 모듈 측정 계획 검증 (12/5, 기술본부)"],
    ["Risk #2 설비 배치 사양", "🔴 미달",
     "🔴 미달: 10 sun light soaking 신뢰성 검증, 라벨 측정 방안 미완료",
     "offset 측정 추가 계획 수립 및 검증 완료 (12/5, 기술본부)"],
    ["JC 파일럿 cell 효율", "🔴 미달", "🟢 통과 (28.6%)", "–"],
    ["JC 파일럿 FpG", "🔴 미달",
     "🔴 미달: 50% (wet station 공정 이슈), 과거 95% 달성 이력 보유",
     "wet station 공정 정상화 후 재측정 (11/13)"],
]

TABLE2_HEADERS = ["구분", "Risk #1 CTM 재현성", "Risk #2 모듈 설비 배치 사양",
                  "JC 파일럿 cell 효율", "JC 파일럿 FpG (Cell 양품률)"]
TABLE2_ROWS = [
    ["문제 현상",
     "동일 등급의 Cell에서 다른 Module 출력 등급이 나타나 관리 방법 부재",
     "Module 라벨 규격을 위한 모듈 측정 방법 부재",
     "테스트 수준에서 목표 확보는 검증했으나 BKM 전환 지연", None],
    ["문제 원인",
     "CTM 재현성 확보를 위한 cell 분류 기준 정의 및 반영 미구현",
     "R&D에서 사용 중인 1 sun light soaker 기반의 안정화 시간이 양산 적용 불가 수준으로 사양을 가짐 (>10시간)",
     "T05 설비 nozzle 막힘 현상 발생으로 설비 안정화 검증 필요", None],
    ["개선방안",
     "Cell 품질 기준에서 Module 출력과 연계되는 핵심 지표를 원리적으로 규명하고 양산 적용 가능한 측정 방법 도입",
     ["Light Soaker: Light 고출력 사양 적용하여 양산 LOB(생산밸런스)를 고려한 공정 사양 검토",
      "offset 측정: Offline 측정기를 이용한 offset 라벨 출력 보정 (First Solar 사례 벤치마킹)"],
     "신규 노즐 도입으로 설비 안정성 테스트", None],
    ["계획 수립",
     "Cell 품질 측정하는 PL, EL, HyR, IV 측정에서 핵심 지표를 대변하는 측정방법을 찾고 cell 분류 기준에 반영",
     ["Light Soaker: 중국 Gsolar 社 10 sun light soaker 적용 유효성 및 신뢰성 이슈 검토",
      "offset 측정: Offline 측정기에서 15시간 MPPT 측정하여 보정값 산정"],
     "CW44부터 공증착 기반 BKM 으로 대량 생산 및 검증", None],
    ["실행 및 결과",
     ["Hysteresis index 원리를 응용한 Toplimit Forward scan 분류 기준 적용 테스트",
      "Full Module 7개 측정, CTM 오차 ±0.3% (기준 ±0.5%) 결과 확보",
      "필요 공정 시간 약 0.2 초로 양산 적용 가능 확인"],
     ["Light Soaker: 2x2 미니 모듈 UV, LeTID 약 99% 검증",
      "offset 측정: 4회 총 35매 측정, 5W bin 기준 적정 등급 74% (측정 당시 보정 기준)"],
     "cell 12,686장 기준 효율 28.6% (기준 ≥ 28.0%)",
     "cell 12,686장 기준 FpG 50% (기준 ≥ 95.0%), Wet station 공정 이슈로 후면 pad 불량"],
    ["결론",
     "🔴 미달. 검증 수량 부족 및 BKM 전환 미완료. 기술본부 BKM cell 분류 기준 마련 계획 수립 및 검증",
     ["🔴 미달. 10 sun light soaking 신뢰성 테스트는 검증했으나 라벨 측정 방안 미완료",
      "모듈 라벨 출력 측정 방안 마련 계획 재수립 및 추가 검증 진행"],
     "🟢 통과",
     "🔴 미달. 11/13 Wet station 공정 정상화 후 FpG 재확보 예정"],
    ["TRL 수준 (양산 공정 기준, 별첨 참조)",
     "JC 파일럿 내 검증 완료 (TRL 5)",
     ["Light Soaker: R&D 검증 완료 (TRL 3)",
      "offset 측정: 타기관 정보 기반의 아이디어 (TRL 2)"],
     "JC 파일럿 BKM (TRL 6)",
     "JC 파일럿 내 검증 완료 (TRL 5)"],
]


def build(path=OUT):
    doc = S.new_document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.0)
    setup_numbering(doc)

    item(doc, 2, "Gate #1 미달 항목 재평가 결과")
    item(doc, 3, "Gate #1 미달 항목 4건(상업화 Risk 2건, JC 파일럿 2건) 재평가를 11/5 실시 (당초 10/30 계획)")
    item(doc, 4, "변경 사유: 파일럿 라인 정전(CW41), 연차보수(CW42), 설비 PM(CW43)으로 cell 생산 일정이 "
                 "순연되어, 검증용 cell 확보(상업화 Risk 2건)와 대량 생산 데이터(cell 효율, FpG) 확보 지연")
    item(doc, 3, "재평가 결과, 미달 항목 4개 중 1개 통과, 3개 미달로 Gate #1 미달 상태 지속")
    item(doc, 3, "Gate #1 전 항목 통과 시까지 Gate #2 이행 보류. 12/5 2차 재평가에서 전 항목 통과 확보 후 "
                 "Gate #2 진행 및 FEED 일정 재수립")

    caption(doc, "[표. 1] Gate #1 평가 경과 및 2차 재평가 계획")
    grid_table(doc, TABLE1_HEADERS, TABLE1_ROWS, [3.2, 2.6, 5.6, 5.6], bold_rows=(0,))

    item(doc, 3, "조치 결과에 따른 항목별 진척 및 잔여 과제")
    item(doc, 4, "1차 재평가에서 cell 효율 1건 해결, 대량 생산 조건에서 목표 효율 확보")
    item(doc, 4, "잔여 미달 3건 중 2건은 기술본부 기술 확보 과제, 1건은 공정 안정화 과제")
    item(doc, 5, "Risk #1 CTM 재현성: 분류 기준의 양산 적용 가능성은 확인, BKM 반영 및 검증 수량 확대 필요")
    item(doc, 5, "Risk #2 모듈 설비 배치 사양: 10 sun light soaking 유효성 및 신뢰성은 검증, 라벨 출력 측정 방안 미완료")
    item(doc, 5, "JC 파일럿 FpG: 일시적 공정 이슈로 미달, 11/13 공정 정상화 후 재검증 예정")

    caption(doc, "[표. 2] Gate #1 미달 항목별 원인 분석 및 조치 경과")
    grid_table(doc, TABLE2_HEADERS, TABLE2_ROWS, [2.0, 3.9, 4.5, 3.3, 3.3])

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build()
