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


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def setup_numbering(doc):
    """다단계 목록 정의 + 단계별 단락 스타일 생성. numId 반환."""
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
    # abstractNum은 num보다 앞에 와야 한다
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(absn)
    else:
        numbering.append(absn)
    num = _el("w:num", **{"w:numId": num_id})
    num.append(_el("w:abstractNumId", **{"w:val": abs_id}))
    numbering.append(num)

    styles = doc.styles
    from docx.enum.style import WD_STYLE_TYPE
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
    return num_id


def item(doc, level, text):
    """level: 1=1., 2=1), 3=①, 4=a), 5=i). 번호는 스타일의 자동 번호."""
    p = doc.add_paragraph(style=STYLE_NAMES[level - 1])
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _b, size = LEVELS[level - 1][4], LEVELS[level - 1][5]
    S._set_run_font(p.add_run(text), size=size, bold=_b)
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    S._set_run_font(p.add_run(text), bold=True)
    return p


def cell_lines(cell, lines, bold=False, center=False):
    cell.text = ""
    for k, line in enumerate(lines):
        p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(0)
        if line:
            S._set_run_font(p.add_run(line), size=9, bold=bold)


HEADERS = ["구분", "항목", "통과 기준", "재평가 결과", "세부 내용", "판정", "비고"]
ROWS = [
    ["상업화 Risk", ["#1 품질 불확실성 (CTM 재현성)"], ["CTM 오차 ±0.5% (> 90%)"],
     ["CTM 오차 0.3% 수준 (54 Full Module 4개 측정 결과)"],
     ["Top Limit Forward bias 방식의 신규 cell 분류 적용하여 cell 출력 기대값과 실제 모듈 출력값 비교 검증 완료. "
      "모듈 수량 부족으로 검증 신뢰성 미달로 판정"],
     "미달", ["기술본부 대량 생산 계획 수립 진행"]],
    ["상업화 Risk", ["#2 모듈 양산 설비 배치 사양 (Light Soaker 공정 확정)"],
     ["기존 열화율 대비 > 99% (10 sun 안정성)"],
     ["단일 cell 모듈 UV, LeTID > 100%, 2x2 미니 모듈 UV, LeTID ~ 99%"],
     ["Ref. 대비 열화율: LeTID100 2 min 98.6%, UV60 2 min 98.5%, 3 min 100.8%"],
     "통과",
     ["'27년 4월 목표로 10sun light soaker 도입 예정으로 모듈 대량 측정 가능 (JC 파일럿 생산 모듈 100% 측정)"]],
    ["상업화 Risk", ["#3 모듈 양산 설비 배치 사양 (offset 측정)"], ["실측값 대비 offset > 99%"],
     ["4회 총 35매 측정"],
     ["4회에 걸쳐 full module offline 출력 측정을 통한 offset 검증 진행. 측정 시마다 offset 산정식 변경으로 "
      "기준 마련 미완료. 5W bin 기준 적정 등급 74% 수준으로 미달"],
     "미달",
     ["offset 측정 방식에 대한 기술본부 내부 결론 도출 및 출력 측정 전략 재수립 예정",
      "10 sun soaking 기반 측정 시 15시간 MPPT 대비 상대 출력 ≥ 100% 확보 (2x2 미니 모듈)",
      "모듈 라벨 출력에 대한 표준 측정 방식 재검토 (offset 측정 vs. 10 sun light soaking 최적화)"]],
    ["JC 파일럿", ["cell 효율"], ["≥ 28.0% (BKM)"], ["효율 28.7% (cell 12,000장 기준)"], ["-"],
     "통과", [""]],
    ["JC 파일럿", ["FpG(cell 양품률)"], ["≥ 95.0% (BKM)"],
     ["FpG 50% (cell 12,000장 기준, 업데이트 예정)"],
     ["Wet station 공정 이슈로 후면 pad 분량 이슈 발생"],
     "미달",
     ["차주 Wet station 공정 정상화 및 FpG 95% 재확보 예정",
      "Co-dep 기준 X회(데이터 확보 후 업데이트 예정) 총 6,000장에 대해 목표 FpG ≥ 95.0% 결과 확보. "
      "일시적인 공정 이슈는 별도 Risk로 관리"]],
]
WIDTHS_CM = [1.6, 2.6, 2.4, 2.8, 4.2, 1.1, 3.3]


def result_table(doc):
    t = doc.add_table(rows=1, cols=len(HEADERS))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(HEADERS):
        c = t.rows[0].cells[j]
        cell_lines(c, [h], bold=True, center=True)
        S._shade(c, S.HDR_FILL)
    body = []
    for r in ROWS:
        cells = t.add_row().cells
        body.append(cells)
        cell_lines(cells[0], [r[0]], center=True)
        cell_lines(cells[1], r[1], center=True)
        cell_lines(cells[2], r[2], center=True)
        cell_lines(cells[3], r[3], center=True)
        cell_lines(cells[4], r[4])
        cell_lines(cells[5], [r[5]], center=True)
        cell_lines(cells[6], r[6])
    # 구분 열 세로 병합
    start = 0
    for i in range(1, len(ROWS) + 1):
        if i == len(ROWS) or ROWS[i][0] != ROWS[start][0]:
            if i - 1 > start:
                for k in range(start + 1, i):
                    body[k][0].text = ""
                m = body[start][0].merge(body[i - 1][0])
                for p in m.paragraphs[1:]:
                    p._element.getparent().remove(p._element)
            start = i
    t.autofit = False
    for j, w in enumerate(WIDTHS_CM):
        t.columns[j].width = Cm(w)
    for row in t.rows:
        for j, w in enumerate(WIDTHS_CM):
            row.cells[j].width = Cm(w)
    return t


def build(path=OUT):
    doc = S.new_document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.0)
    setup_numbering(doc)

    item(doc, 2, "Gate #1 미달 항목 재평가 결과")
    item(doc, 3, "Gate #1 미달 항목 5건(상업화 Risk 3건, JC 파일럿 2건) 재평가 일정을 10/30에서 11/5로 변경하여 실시")
    item(doc, 4, "변경 사유: 파일럿 라인 정전(CW41) 이후 연차보수(CW42)와 설비 PM(CW43)이 이어지며 cell 생산 일정이 "
                 "순연되어, 검증용 cell 확보(상업화 Risk 3건)와 대량 생산 데이터 안정화(cell 효율, FpG)가 지연")
    item(doc, 3, "재평가 결과, 세부 항목 5개 중 2개 통과, 3개 미달에 따라 해당 항목 완료 계획 재수립")
    item(doc, 4, "Risk #1. CTM 재현성: Full Module 기준 4개 측정 시 CTM 오차 0.3% 수준이나 소량 검증으로 "
                 "검증 신뢰성 측면에서 미달로 구분")
    item(doc, 4, "Risk #3. offset 측정")
    item(doc, 5, "4회에 걸쳐 총 35매 측정을 하였으나 측정마다 offset 산정식이 달라 모듈 등급 차이 발생. "
                 "5W bin 기준 적정 등급 74% 수준으로 미달")
    item(doc, 5, "10 sun light soaking 측정 시 15시간 MPPT 대비 모듈 상대 출력 100% 이상 확인(2x2 미니 모듈 기준). "
                 "모듈 라벨 출력 측정 방식(offset 측정 또는 10 sun light soaking)은 기술본부 내부 검토 후 결정 예정")
    item(doc, 4, "JC 파일럿 FpG")
    item(doc, 5, "50% 수준(업데이트 예정)의 결과 확보로 미달. Wet station 공정 이슈 발생으로 후면 pad 불량")
    item(doc, 5, "기존 테스트 시행 X회(데이터 확보 후 업데이트 예정) 총 6,000장 수량에서 FpG 95% 달성 이력 보유. "
                 "차주 wet station 공정 정상화 후 FpG 목표값 재확보 예정")

    caption(doc, "[표. 1] 상업화 Gate #1 미달 항목 재평가 결과")
    result_table(doc)

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build()
