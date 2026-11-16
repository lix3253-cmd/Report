"""제조본부장 주간보고(11/16) docx 빌더 — 바탕체 12pt 단일 서식.

내용은 Claude Docs 문서(제조본부장 주간보고) rev 264 기준.
"""
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONT = "바탕체"
SIZE = 12
OUT = "제조본부장_주간보고_1116.docx"


def font(run, bold=False):
    run.font.name = FONT
    run.font.size = Pt(SIZE)
    run.font.bold = bold
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)


def para(doc, text, left_cm=0.0, bold=False, space_before=0, center=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(left_cm)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(2)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run(text), bold=bold)
    return p


def shade(cell, fill="D9E2F3"):
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def table(doc, caption, headers, rows):
    para(doc, caption, space_before=4)
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
    doc.add_paragraph()


def build():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(SIZE)
    rfonts = normal.element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.0)
        s.top_margin = s.bottom_margin = Cm(2.0)

    L1, L2, L3, L4 = 0.0, 0.5, 1.0, 1.6  # 1) / ① / a) / i)

    para(doc, "제조본부장 주간보고", bold=True, center=True)
    para(doc, "2026.11.16", center=True)

    # 1)
    para(doc, "1) Gate 2차 재평가 계획 및 catch-up plan 수립 현황", L1, bold=True, space_before=8)
    para(doc, "① 미충족 항목 확인을 위한 1차 재평가 결과 요약", L2, bold=True, space_before=4)
    para(doc, "a) Gate #1 미달 4건(상업화 Risk 2건, JC 파일럿 2건) 대상 1차 재평가 (11/5) "
              "결과 미완료 항목 3건에 대한 2차 재평가 계획 수립", L3)
    para(doc, "b) 미달 항목은 2차 재평가에서 검증 완료 예정 (12/4)", L3)
    table(doc, "[표. 1] Gate #1 미달 항목별 재평가 결과 및 2차 재평가 일정",
          ["구분", "점검 항목(목표)", "1차 재평가(11/5)", "미달 사유", "2차 재평가"],
          [
              ["상업화 Risk#1", "Module High-Risk: CTM 재현성", "미달", "샘플 모수 부족", "12/4 예정"],
              ["상업화 Risk#2", "Module High-Risk: 모듈 설비 배치 사양", "미달",
               "라벨 출력 측정 방안(offset 기법) 결과 신뢰성 낮음", "12/4 예정"],
              ["JC 파일럿", "cell 효율(≥ 28.0%)", "통과(28.6%)", "-", "-"],
              ["JC 파일럿", "FpG(셀 양품률, ≥ 95.0%)", "미달(50%, 11/13 재측정 80%)",
               "공증착 Nozzle 막힘으로 인한 FpG 목표 대비 미달", "12/4 예정"],
          ])
    para(doc, "② 미달 항목 해소를 위한 2차 재평가 계획", L2, bold=True, space_before=4)
    para(doc, "a) 미달 3건 모두 12/4 2차 재평가로 Gate #1 전 항목 통과 목표", L3)
    para(doc, "b) 파일럿 FpG catch-up plan 수립 완료", L3)
    para(doc, "i) Wet station 공정으로 인한 후면 패드 정상 반영하였으나 최종 FpG 80%로 목표 (95%) 미달", L4)
    para(doc, "ii) 공증착 공정 시 nozzle 막힘현상 반복 발생으로 FpG저하 발생. 2차 재평가까지 해소 계획 수립", L4)
    para(doc, "c) 상업화 Risk 2건은 계획 수정 중으로 11/16 완료 예정", L3)
    para(doc, "i) CTM 재현성: 신규 cell sorting criteria 기반 대량 모듈 결과 확보, JC BKM 반영 여부 검증", L4)
    para(doc, "ii) 모듈 설비 배치 사양: 기술본부 모듈 라벨 측정 계획(offset 측정) 결론, 모듈 라벨 출력 방안 검증", L4)

    # 2)
    para(doc, "2) 선익시스템 양산 증착기 개발 현황", L1, bold=True, space_before=10)
    para(doc, "① 신규 양산 설비 개발 현황", L2, bold=True, space_before=4)
    para(doc, "a) 고가 재료인 C60을 대상으로 Material Recycle System 개발", L3)
    para(doc, "i) 증착 중 외벽에 버려지는 재료를 basket에 회수해 재사용", L4)
    para(doc, "ii) Basket과 linear source 구조를 동일하게 설계해 별도 조치 없이 회수 재료 재투입 가능", L4)
    para(doc, "② 양산 적용 시 유효성", L2, bold=True, space_before=4)
    para(doc, "a) 버려지던 C60을 회수해 재료 사용 효율을 높이는 구조로, 양산 적용 시 재료비 절감 가능", L3)
    para(doc, "b) 단, 재사용 재료의 공정 재현성과 효율 수치는 업체 제시 기준으로 당사 검증 전 단계", L3)
    table(doc, "[표. 2] Material Recycle System 검증 항목 및 업체 제시 효과",
          ["구분", "검증 항목", "업체 제시 현황"],
          [
              ["공정 재현성", "증착 Rate 재현", "검증 진행 [확인필요: 결과]"],
              ["물질 특성 재현성", "C60 물질 특성(Powder XRD, TOF-SIMS 분석)", "검증 진행 [확인필요: 결과]"],
              ["재료 사용 효율", "재료 사용 유효율", "일반 증착기 40% → 선익 신규 공정 60%"],
          ])
    para(doc, "③ 후속 검토 사항", L2, bold=True, space_before=4)
    para(doc, "a) 재료 사용 유효율 산정 근거 및 타사 비교 데이터 확보 (선익 추가 검토 예정)", L3)
    para(doc, "b) 재사용 재료 재현성 검증 결과 확인 후 당사 양산 적용 여부 검토", L3)

    doc.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    build()
