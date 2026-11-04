# -*- coding: utf-8 -*-
"""11월 대표이사 월간보고 ③항(Track 2 1GW capa 배치 검토) docx 생성기.

다른 보고서 파일에 붙여 넣을 조각용. 번호는 글자가 아니라 Word 자동 번호 서식으로 넣는다.
③부터 시작하도록 레벨 2(①) 시작값을 3으로 둔다.
원문: Claude Docs '11월 대표이사 월간보고 - Track 2 1GW capa 가정 검증을 위한 설비 배치 검토'
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
                                   "11월_월간보고_3_Track2_배치검토.docx"))

TABLE1 = {
    "headers": ["구분 (대수)", "T01 : Cleaning", "T02~T04 : HTL 및 PVK 형성",
                "T05~T07 : ETL 형성", "T08~T10 : 전극 형성 및 Sorter"],
    "rows": [
        ["Track 1", "Wet Station(2)",
         "Sputter(2), SAM Inkjet(2), PVK Inkjet(3), Oven(3)",
         "Evaporator(2), Batch ALD(3), TCO Sputter(2)",
         "Screen Printer, Laser, Curing Oven, Light Soaker, Sorter(각 3)"],
        ["Track 2", "HF mist(2)",
         "Sputter(2), Evaporator(2), PVK Inkjet(4), Multi-step Oven(4)",
         "Evaporator, Spatial ALD, TCO Sputter(2)",
         "Screen Printer(PTP) 외 동일(각 3)"],
    ],
}

TABLE2 = {
    "headers": ["항목", "목표(RFP 수준)", "현황(RFI 수준)", "레이아웃 검토 반영"],
    "rows": [
        ["T02 Sputter\nT03~T04 Evaporator", "T02~T04 통합 레이아웃",
         "T02 Sputter, T03~T04 Evaporator 개별\n이미지 (통합 레이아웃 설계 미완료)",
         "Box Footprint 기반 검토\n(10 m × 70 m)"],
        ["T04 Inkjet", "T04 Inkjet, 90분 열처리 오븐 레이아웃", "필요 설비 수 4대",
         "자체 개발 설비 Inkjet, Cassette 오븐 사양 차용 산정\n(/w 탠덤장비개발팀)"],
        ["T05 Evaporator\nT06 ALD\nT07 Sputter", "T05~T07 통합 레이아웃",
         "T05 ~ T07 설비 사이즈 정보",
         "Box Footprint 기반 검토\n(10 m × 100 m)"],
    ],
}


def build():
    doc = S.new_document()
    L = S.native_list(doc, starts={2: 3})     # ③부터 시작
    N = lambda lvl, text, **kw: S.nitem(doc, L, lvl, text, **kw)

    N(2, "경제성 분석의 1GW capa 가정 검증을 위한 진천공장 셀 설비 배치 검토",
      bold=True, size=11.5, space_before=6)

    N(3, "1GW capa 가정의 배치 검증 배경", bold=True, space_before=4)
    N(4, "기존 기술본부 Track 2 경제성 분석 시 capa 1GW로 가정하여 반영하였으나, "
         "공정 특성상 필요 면적 증가가 예상되어 리스크 검토 추진(배치 불가 시 예상 capa 0.6GW)")
    N(4, "Track 2 주요 공정 및 설비를 기준으로 진천공장 2동 1층 배치 진행, "
         "설비 레이아웃 설계도 미확정으로 업체(Von Ardenne社) 제시 필요면적 수준의 설비 크기 반영")
    S.caption(doc, "[표. 1] 1GW capa 배치에 반영한 공정 구간별 설비 비교(Track 1 및 Track 2)")
    S.grid_table(doc, TABLE1["headers"], TABLE1["rows"])

    N(3, "1GW capa 배치 가능성 검토 결과", bold=True, space_before=8)
    N(4, "진천공장 2동 1층 내 1GW capa 설비 배치 가능 확인, 경제성 검토의 1GW capa 가정은 "
         "반영 가능하며 상세 레이아웃 추가 검토가 필요한 수준으로 판단")
    N(5, "설비 Footprint를 기준으로 상부 셀 라인 설비 배치 초안 작성 및 기존 공장 공간 활용 가능성 검토")
    N(5, "설비 배치뿐 아니라 물류 이송 동선(AMR), 자재 대기 공간(CST Buffer), 유지보수 공간 등을 "
         "고려하여 양산 운영 관점의 Layout 적정성 검토")
    p = doc.add_paragraph()
    p.alignment = S.WD_ALIGN_PARAGRAPH.CENTER
    S._set_run_font(p.add_run("(그림 삽입 위치)"))
    S.caption(doc, "[그림. 1] Track 2 공정 진천공장 레이아웃 배치 결과", space_before=0)

    N(4, "현재 설비 사양은 업체가 공정 단위로 제시한 RFI 수준으로, FEED 준비에 필요한 RFP 수준은 추가 검토 필요",
      space_before=4)
    N(5, "RFI : 당사 필요 공정에 대한 업체에서 제안하는 설비 사양")
    N(5, "RFP : 당사 필요 설비 주요 사양 협의 (Layout, UT matrix, 투자비 관련 요구사항)")
    N(4, "업체를 통해 통합 설비 Drawing 정보 확인 후 상세 레이아웃 검토 예정")
    S.caption(doc, "[표. 2] 레이아웃 검토 반영 설비 사양 목표 대비 현황")
    S.grid_table(doc, TABLE2["headers"], TABLE2["rows"])

    doc.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
