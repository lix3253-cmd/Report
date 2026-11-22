# -*- coding: utf-8 -*-
"""주간보고 절: Gate #1 미달 항목 2차 재평가 및 catch-up plan 추진 현황.

내용 원본: workflow/3_Gate1_2차재평가_개요.md
자리표시([확인필요], [협의중], [MM/DD 결과])는 사실 확정 후 채운다.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "주간보고_Gate1_2차재평가.docx"))
CK, NG = "[확인필요]", "[협의중]"

T1 = {
    "headers": ["구분", "점검 항목", "1차 재평가(11/5)", "2차 재평가 계획", "제약 사항", "완료 전망"],
    "rows": [
        ["상업화 Risk#1", "CTM 재현성", "미달\n(샘플 모수 부족)",
         "신규 cell sorting criteria 기반 대량 모듈 결과 확보, JC BKM 반영 여부 검증", CK, NG],
        ["상업화 Risk#2", "모듈 설비 배치 사양", "미달\n(offset 기법 신뢰성 낮음)",
         "기술본부 모듈 라벨 측정 계획(offset 측정) 결론, 라벨 출력 방안 검증", CK, NG],
        ["JC 파일럿", "FpG(셀 양품률)\n≥ 95.0%", "미달\n(50%, 11/13 재측정 80%)",
         "공증착 nozzle 막힘 해소 후 BKM 전환", f"증착기 반복 이슈 {CK}", "12/4 예정"],
    ],
}

T2 = {
    "headers": ["항목", "검증 내용", "담당", "당초 일정", "제약 사유", "조정 일정"],
    "rows": [
        ["Risk#1\nCTM 재현성", "대량 모듈 결과 확보, JC BKM 반영 여부", f"기술본부 {CK}", "12/4", CK, NG],
        ["Risk#2\n모듈 설비 배치 사양", "offset 측정 결론, 라벨 출력 방안", "기술본부", "12/4", CK, NG],
    ],
}

T3 = {
    "headers": ["이슈", "발생 시점", "FpG 영향", "조치 내용", "상태"],
    "rows": [
        ["공증착 nozzle 막힘", CK, CK, CK, CK],
        [CK, CK, CK, CK, CK],
    ],
}


def widths(t, cms):
    """열 너비(cm) 지정. 서술 열을 넓혀 줄바꿈을 줄인다."""
    from docx.shared import Cm
    t.autofit = False
    for col, w in zip(t.columns, cms):
        col.width = Cm(w)
    for row in t.rows:
        for c, w in zip(row.cells, cms):
            c.width = Cm(w)


def build(path=OUT):
    doc = S.new_document()
    S.h1(doc, "Gate #1 미달 항목 2차 재평가 및 catch-up plan 추진 현황")

    # 1) 핵심 요약
    S.h2(doc, "핵심 요약")
    S.h3(doc, "미달 3건(상업화 Risk 2건, 파일럿 FpG) 12/4 2차 재평가로 Gate #1 전 항목 통과 목표 유지")
    S.h3(doc, f"상업화 Risk 2건은 기술본부 일정 제약으로 12/4 내 검증 완료 불확실, "
              f"제약 사유 및 조정 일정 정리 중 {CK}")
    S.h3(doc, f"파일럿 FpG는 nozzle 막힘 해소 후 BKM 전환 계획이나 증착기 이슈 반복 발생으로 "
              f"일정 리스크 존재 {CK}")
    S.caption(doc, "Gate #1 미달 항목별 2차 재평가 계획 및 완료 전망")
    widths(S.rnr_table(doc, T1["headers"], T1["rows"], wide_cols=[3]), [1.9, 2.4, 2.6, 4.8, 2.3, 2.0])

    # 2) 상업화 Risk
    S.h2(doc, "상업화 Risk 2건 검증 완료를 위한 2차 재평가 계획 및 제약 사항")
    S.h3(doc, "CTM 재현성, 모듈 설비 배치 사양 모두 기술본부 주관 검증으로, "
              "2차 재평가에서 상업화 적용 가능 여부 판단 예정")
    S.h3(doc, f"기술본부 측 12/4 내 완료 어려움 통보, 두 항목 공통 제약 사유는 {CK}")
    S.h3(doc, f"11/16 완료 예정이던 계획 수정 결과 {CK}, "
              f"조정 완료 시점은 기술본부와 협의 후 확정 {NG}")
    S.caption(doc, "상업화 Risk 항목별 검증 내용, 일정 및 제약 사유")
    widths(S.rnr_table(doc, T2["headers"], T2["rows"], wide_cols=[1, 4]), [2.6, 4.3, 2.2, 1.9, 3.0, 2.0])

    # 3) 파일럿 FpG
    S.h2(doc, "파일럿 FpG 목표 달성을 위한 catch-up plan 및 증착기 이슈")
    S.h3(doc, "Wet station 후면 패드 반영 이후 남은 미달 원인은 공증착 nozzle 막힘으로, "
              "해소 후 BKM 전환 계획")
    S.h3(doc, f"nozzle 막힘 외 증착기 이슈 반복 발생으로 BKM 전환 지연 우려 {CK}")
    S.h4(doc, f"이슈 내용 및 발생 빈도 {CK}")
    S.h4(doc, f"12/4 재평가 일정 영향 {CK}")
    S.h3(doc, "12/4 2차 재평가 판정 [12/4 결과: 충족 / 미충족]")
    S.caption(doc, "증착기 이슈별 발생 현황 및 조치")
    widths(S.rnr_table(doc, T3["headers"], T3["rows"], wide_cols=[3]), [3.2, 2.2, 2.6, 5.0, 2.0])

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build()
