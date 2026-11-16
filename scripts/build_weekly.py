# -*- coding: utf-8 -*-
"""제조본부장 주간보고(11/16) docx 빌더 — 다른 Word 파일에 붙여넣기용.

번호(1) / ① / a) / i))는 글자로 넣지 않고 Word 자동 번호(다단계 목록)로 넣는다.
각 단계는 단락 스타일 "보고 1단계" ~ "보고 4단계"에 연결되어 있어 붙여넣어도 번호가 유지된다.
(origin/claude/lucid-feynman-mclwvn 의 build_gate1_reeval.py 와 같은 방식)
글꼴: 바탕체 12pt. 내용은 Claude Docs 문서(제조본부장 주간보고) rev 264 기준.
"""
import os

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONT = "바탕체"
SIZE = 12
HDR_FILL = "DDEBF7"
OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "제조본부장_주간보고_1116.docx"))

# 단계: (번호 형식, 표시 문자, 왼쪽 cm, 내어쓰기 cm, 굵게)
LEVELS = [
    ("decimal",               "%1)", 0.0, 0.7, True),   # 1)
    ("decimalEnclosedCircle", "%2",  0.5, 0.7, True),   # ①
    ("lowerLetter",           "%3)", 1.2, 0.7, False),  # a)
    ("lowerRoman",            "%4)", 1.9, 0.8, False),  # i)
]
STYLE_NAMES = [f"보고 {i + 1}단계" for i in range(len(LEVELS))]


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def font(run, bold=False):
    run.font.name = FONT
    run.font.size = Pt(SIZE)
    run.font.bold = bold
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)


def setup_document():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(SIZE)
    rfonts = normal.element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    normal.paragraph_format.space_after = Pt(2)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.0)
    return doc


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
        rpr = _el("w:rPr")
        rpr.append(_el("w:rFonts", **{"w:ascii": FONT, "w:hAnsi": FONT,
                                      "w:eastAsia": FONT, "w:cs": FONT}))
        if _b:
            rpr.append(_el("w:b"))
        rpr.append(_el("w:sz", **{"w:val": SIZE * 2}))
        lvl.append(rpr)
        absn.append(lvl)
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(absn)
    else:
        numbering.append(absn)
    num = _el("w:num", **{"w:numId": num_id})
    num.append(_el("w:abstractNumId", **{"w:val": abs_id}))
    numbering.append(num)

    for i, (fmt, text, left, hang, bold) in enumerate(LEVELS):
        st = doc.styles.add_style(STYLE_NAMES[i], WD_STYLE_TYPE.PARAGRAPH)
        st.element.set(qn("w:styleId"), f"Report{i + 1}")
        st.base_style = doc.styles["Normal"]
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


def item(doc, level, text):
    """level: 1=1), 2=①, 3=a), 4=i). 번호는 스타일의 자동 번호."""
    p = doc.add_paragraph(style=STYLE_NAMES[level - 1])
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    font(p.add_run(text), bold=LEVELS[level - 1][4])
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    font(p.add_run(text), bold=True)


def shade(cell):
    tcpr = cell._element.get_or_add_tcPr()
    tcpr.append(_el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": HDR_FILL}))


def table(doc, headers, rows, widths, merge_first=True):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, values in enumerate([headers] + rows):
        for c, v in enumerate(values):
            cell = t.cell(r, c)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            font(p.add_run(v), bold=(r == 0))
            if r == 0:
                shade(cell)
    if merge_first:  # 구분 열의 같은 값 세로 병합
        start = 1
        for i in range(2, len(rows) + 2):
            if i == len(rows) + 1 or rows[i - 1][0] != rows[start - 1][0]:
                if i - 1 > start:
                    for k in range(start + 1, i):
                        t.cell(k, 0).text = ""
                    m = t.cell(start, 0).merge(t.cell(i - 1, 0))
                    for p in m.paragraphs[1:]:
                        p._element.getparent().remove(p._element)
                start = i
    t.autofit = False
    for row in t.rows:
        for j, w in enumerate(widths):
            row.cells[j].width = Cm(w)
    doc.add_paragraph()


def build(path=OUT):
    doc = setup_document()
    setup_numbering(doc)

    item(doc, 1, "Gate 2차 재평가 계획 및 catch-up plan 수립 현황")
    item(doc, 2, "미충족 항목 확인을 위한 1차 재평가 결과 요약")
    item(doc, 3, "Gate #1 미달 4건(상업화 Risk 2건, JC 파일럿 2건) 대상 1차 재평가 (11/5) "
                 "결과 미완료 항목 3건에 대한 2차 재평가 계획 수립")
    item(doc, 3, "미달 항목은 2차 재평가에서 검증 완료 예정 (12/4)")
    caption(doc, "[표. 1] Gate #1 미달 항목별 재평가 결과 및 2차 재평가 일정")
    table(doc,
          ["구분", "점검 항목(목표)", "1차 재평가(11/5)", "미달 사유", "2차 재평가"],
          [
              ["상업화 Risk#1", "Module High-Risk: CTM 재현성", "미달", "샘플 모수 부족", "12/4 예정"],
              ["상업화 Risk#2", "Module High-Risk: 모듈 설비 배치 사양", "미달",
               "라벨 출력 측정 방안(offset 기법) 결과 신뢰성 낮음", "12/4 예정"],
              ["JC 파일럿", "cell 효율(≥ 28.0%)", "통과(28.6%)", "-", "-"],
              ["JC 파일럿", "FpG(셀 양품률, ≥ 95.0%)", "미달(50%, 11/13 재측정 80%)",
               "공증착 Nozzle 막힘으로 인한 FpG 목표 대비 미달", "12/4 예정"],
          ],
          [2.4, 3.8, 3.2, 4.6, 2.6])
    item(doc, 2, "미달 항목 해소를 위한 2차 재평가 계획")
    item(doc, 3, "미달 3건 모두 12/4 2차 재평가로 Gate #1 전 항목 통과 목표")
    item(doc, 3, "파일럿 FpG catch-up plan 수립 완료")
    item(doc, 4, "Wet station 공정으로 인한 후면 패드 정상 반영하였으나 최종 FpG 80%로 목표 (95%) 미달")
    item(doc, 4, "공증착 공정 시 nozzle 막힘현상 반복 발생으로 FpG저하 발생. 2차 재평가까지 해소 계획 수립")
    item(doc, 3, "상업화 Risk 2건은 계획 수정 중으로 11/16 완료 예정")
    item(doc, 4, "CTM 재현성: 신규 cell sorting criteria 기반 대량 모듈 결과 확보, JC BKM 반영 여부 검증")
    item(doc, 4, "모듈 설비 배치 사양: 기술본부 모듈 라벨 측정 계획(offset 측정) 결론, 모듈 라벨 출력 방안 검증")

    item(doc, 1, "선익시스템 양산 증착기 개발 현황")
    item(doc, 2, "신규 양산 설비 개발 현황")
    item(doc, 3, "고가 재료인 C60을 대상으로 Material Recycle System 개발")
    item(doc, 4, "증착 중 외벽에 버려지는 재료를 basket에 회수해 재사용")
    item(doc, 4, "Basket과 linear source 구조를 동일하게 설계해 별도 조치 없이 회수 재료 재투입 가능")
    item(doc, 2, "양산 적용 시 유효성")
    item(doc, 3, "버려지던 C60을 회수해 재료 사용 효율을 높이는 구조로, 양산 적용 시 재료비 절감 가능")
    item(doc, 3, "단, 재사용 재료의 공정 재현성과 효율 수치는 업체 제시 기준으로 당사 검증 전 단계")
    caption(doc, "[표. 2] Material Recycle System 검증 항목 및 업체 제시 효과")
    table(doc,
          ["구분", "검증 항목", "업체 제시 현황"],
          [
              ["공정 재현성", "증착 Rate 재현", "검증 진행 [확인필요: 결과]"],
              ["물질 특성 재현성", "C60 물질 특성(Powder XRD, TOF-SIMS 분석)", "검증 진행 [확인필요: 결과]"],
              ["재료 사용 효율", "재료 사용 유효율", "일반 증착기 40% → 선익 신규 공정 60%"],
          ],
          [3.6, 6.4, 6.6], merge_first=False)
    item(doc, 2, "후속 검토 사항")
    item(doc, 3, "재료 사용 유효율 산정 근거 및 타사 비교 데이터 확보 (선익 추가 검토 예정)")
    item(doc, 3, "재사용 재료 재현성 검증 결과 확인 후 당사 양산 적용 여부 검토")

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build()
