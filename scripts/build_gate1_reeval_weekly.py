# -*- coding: utf-8 -*-
"""제조본부장 주간보고 — Gate #1 미달 항목 재평가 docx 생성기.

build_gate1_reeval.py(11월 월간보고)와 같은 조건으로 만든다.
- 번호(1) / ① / a) / i))는 Word 자동 번호(다단계 목록, "보고 N단계" 스타일)
- 표: Table Grid, 9pt, 머리행 음영, 구분 열 세로 병합, 여백 2cm
내용은 Claude Docs 사본(Gate #1 미달 항목 재평가)을 그대로 옮긴다.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S
from build_gate1_reeval import setup_numbering, item, caption, cell_lines
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "제조본부장_주간보고_Gate1_재평가.docx"))

# [표. 1] 상업화 Gate #1 미달 항목 재평가 결과
T1_HEADERS = ["구분", "항목", "통과 기준", "재평가 결과", "세부 내용", "판정", "비고"]
T1_ROWS = [
    ["상업화 Risk", ["#1 품질 불확실성 (CTM 재현성)"], ["CTM 오차 ±0.5% (> 90%)"],
     ["CTM 오차 0.3% 수준 (54 Full Module 4개 측정 결과)"],
     ["Toplimit Forward scan 방식의 신규 cell 분류 적용하여 cell 출력 기대값과 실제 모듈 출력값 비교 검증 완료. "
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
    ["JC 파일럿", ["FpG (Cell 양품률)"], ["≥ 95.0% (BKM)"],
     ["FpG 50% (cell 12,000장 기준, 업데이트 예정)"],
     ["Wet station 공정 이슈로 후면 pad 분량 이슈 발생"],
     "미달",
     ["차주 Wet station 공정 정상화 및 FpG 95% 재확보 예정",
      "Co-dep 기준 X회(데이터 확보 후 업데이트 예정) 총 6,000장에 대해 목표 FpG ≥ 95.0% 결과 확보. "
      "일시적인 공정 이슈는 별도 Risk로 관리"]],
]
T1_WIDTHS = [1.6, 2.6, 2.4, 2.8, 4.2, 1.1, 3.3]

# [표. 2] 재평가 항목별 문제 원인부터 결론까지 단계별 상세 (행: 단계, 열: 항목)
T2_HEADERS = ["구분", "Risk #1 CTM 재현성",
              "Risk #2 모듈 설비 배치 사양 - Light Soaker 공정 확정",
              "Risk #3 모듈 설비 배치 사양 - offset 측정",
              "JC 파일럿 cell 효율", "JC 파일럿 FpG (Cell 양품률)"]
T2_ROWS = [
    ["문제 현상",
     ["동일 등급의 Cell에서 다른 Module 출력 등급이 나타나 관리 방법 부재"],
     ["Module 라벨 규격을 위한 모듈 측정 방법 부재"],
     ["좌동"],
     ["테스트 수준에서 목표 확보는 검증했으나 BKM 전환 지연"],
     [""]],
    ["문제 원인",
     ["CTM 재현성 확보를 위한 cell 분류 기준 추가 마련 필요"],
     ["R&D에서 사용하는 1 sun light soaker 기반의 안정화 시간이 양산 적용 불가 수준으로 사양을 가짐"],
     ["좌동"],
     ["T05 설비 nozzle 막힘 현상 발생으로 설비 안정화 검증 필요"],
     [""]],
    ["개선방안",
     ["Cell 품질 기준에서 Module 출력과 핵심 파라미터를 원리적으로 규명하고 양산 적용 가능한 측정 방법 도입"],
     ["Light 고출력 사양 적용하여 양산 LOB(생산밸런스)를 고려한 공정 사양 검토"],
     ["Offline 측정기를 이용한 offset 라벨 출력 보정 (First Solar 사례 벤치마킹)"],
     ["신규 노즐 도입으로 설비 안정성 테스트"],
     [""]],
    ["계획 수립",
     ["Cell 품질 측정하는 PL, EL, HyR, IV 측정에서 핵심 파라미터를 대변하는 측정방법을 찾고 cell 분류 기준에 반영"],
     ["중국 Gsolar 社 10 sun light soaker 적용 유효성 및 신뢰성 이슈 검토"],
     ["Offline 측정기에서 15시간 MPPT 측정하여 offset 보정값 산정"],
     ["CW44부터 공증착 기반 BKM 으로 대량 생산 및 검증"],
     ["좌동"]],
    ["실행 및 결과",
     ["Hysteresis index 원리를 응용한 Toplimit Forward scan 분류 기준 적용 테스트",
      "Full Module 7개 측정, CTM 오차 0.3% (기준 ±0.5%) 결과 확보",
      "필요 공정 시간 약 0.2 초로 양산 적용 가능 확인"],
     ["단일 cell 모듈 UV, LeTID > 100% 검증",
      "2x2 미니 모듈 UV, LeTID 약 99% 검증",
      "Ref. 대비 열화율 LeTID100 2분 98.6%, UV60 2분 98.5%, 3분 100.8%"],
     ["4회 총 35매 측정, 5W bin 기준 적정 등급 74% (측정 당시 offset 보정 기준)"],
     ["cell 12,000장 기준 효율 28.7% (기준 ≥ 28.0%)"],
     ["cell 12,000장 기준 FpG 50% (기준 ≥ 95.0%), Wet station 공정 이슈로 후면 pad 불량"]],
    ["결론",
     ["미달. 검증 수량 부족, 기술본부 대량 생산 계획 수립 후 재검증 예정"],
     ["통과. 10 sun light soaking 공정 유효성 및 신뢰성 risk 검증 완료",
      "'27년 4월 10 sun light soaker 도입 예정. 이후 생산 모듈 전수 측정 가능 (현재 측정 시간 : 2시간/EA in JC 파일럿)"],
     ["미달",
      "10 sun soaking 측정 시 15시간 MPPT 대비 상대 출력 ≥ 100% (2x2 미니 모듈)",
      "기술본부 개발 로드맵 재검토 후 출력 측정 방식 확정 예정"],
     ["통과"],
     ["미달. 차주 Wet station 공정 정상화 후 FpG 재확보 예정"]],
    ["TRL 수준 (양산 공정 기준, 별첨 참조)",
     ["JC 파일럿 내 검증 완료 (TRL 5)"],
     ["R&D 검증 완료 (TRL 3)"],
     ["타기관 정보 기반의 아이디어 (TRL 2)"],
     ["JC 파일럿 BKM (TRL 6)"],
     ["JC 파일럿 BKM (TRL 6)"]],
]
T2_WIDTHS = [2.0, 3.0, 3.0, 3.0, 3.0, 3.0]


def _widths(t, widths):
    t.autofit = False
    for j, w in enumerate(widths):
        t.columns[j].width = Cm(w)
    for row in t.rows:
        for j, w in enumerate(widths):
            row.cells[j].width = Cm(w)


def _header(t, headers):
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        cell_lines(c, [h], bold=True, center=True)
        S._shade(c, S.HDR_FILL)


def table1(doc):
    t = doc.add_table(rows=1, cols=len(T1_HEADERS))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _header(t, T1_HEADERS)
    body = []
    for r in T1_ROWS:
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
    for i in range(1, len(T1_ROWS) + 1):
        if i == len(T1_ROWS) or T1_ROWS[i][0] != T1_ROWS[start][0]:
            if i - 1 > start:
                for k in range(start + 1, i):
                    body[k][0].text = ""
                m = body[start][0].merge(body[i - 1][0])
                for p in m.paragraphs[1:]:
                    p._element.getparent().remove(p._element)
            start = i
    _widths(t, T1_WIDTHS)
    return t


def table2(doc):
    t = doc.add_table(rows=1, cols=len(T2_HEADERS))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _header(t, T2_HEADERS)
    for r in T2_ROWS:
        cells = t.add_row().cells
        cell_lines(cells[0], [r[0]], bold=True, center=True)
        for j, lines in enumerate(r[1:], start=1):
            cell_lines(cells[j], lines)
    _widths(t, T2_WIDTHS)
    return t


def build(path=OUT):
    doc = S.new_document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.0)
    setup_numbering(doc)

    item(doc, 2, "Gate #1 미달 항목 재평가 결과")
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
    table1(doc)

    item(doc, 2, "재평가 근거 확인을 위한 항목별 상세")
    item(doc, 3, "통과 2건은 양산 적용 가능성까지 확인")
    item(doc, 4, "Risk #2 Light Soaker 공정: R&D 안정화 시간이 양산에 맞지 않던 문제를 10 sun 고출력 사양으로 해소, "
                 "설비 도입 후 생산 모듈 전수 측정 체계 확보 예정")
    item(doc, 4, "JC 파일럿 cell 효율: 공증착 기반 BKM 대량 생산에서 목표 효율 확보")
    item(doc, 3, "미달 3건은 원인의 성격이 달라 항목별로 후속 조치 구분")
    item(doc, 4, "Risk #1 CTM 재현성: 신규 cell 분류 기준의 양산 적용 가능성은 확인했으나 검증 수량이 부족해 "
                 "대량 생산 기반 재검증 필요")
    item(doc, 4, "Risk #3 offset 측정: 아이디어 단계 기술로 보정 기준이 미확립되어, 기술본부 개발 로드맵 재검토 후 "
                 "출력 측정 방식 확정 예정")
    item(doc, 4, "JC 파일럿 cell FpG: 일시적 공정 이슈로 인한 미달로, 이전 테스트에서 목표를 달성한 이력이 있어 "
                 "공정 정상화 후 재확보 예정")
    caption(doc, "[표. 2] 재평가 항목별 문제 원인부터 결론까지 단계별 상세")
    table2(doc)

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build()
