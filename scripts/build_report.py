# -*- coding: utf-8 -*-
"""제조본부장 현장 경영지시사항 보고서 (우주태양광) 생성기.

구조: 표제 + 4단계 계층 본문(1./1)/①/a)) + 표 2종 + 로드맵 이미지.
번호는 report_style 엔진이 자동 부여·리셋한다(호출부에서 번호를 넘기지 않음).
로드맵 이미지는 scripts/build_roadmap.py 로 먼저 생성해 둔다.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S

HERE = os.path.dirname(__file__)
ROADMAP = os.path.abspath(os.path.join(HERE, "..", "assets", "roadmap.png"))

# ── 표1: PVA 소싱 → 내재화 전환 ──────────────────────────────────
PVA_SOURCING = {
    "headers": ["구분", "~'28", "'29~"],
    "rows": [
        ["PVA 제작",   "외주 소싱",                 "내재화 전환"],
        ["설계 소유권", "소싱 방식 검토중",           "자사 보유"],
        ["연동 사유",   "실증 단계·조달 리드타임 확보", "flight heritage 진입·캡티브 공급 개시"],
        ["관리 필요",   "외주 사양·표준 정합",         "델타 인증 부담 관리"],
    ],
}

# ── 표2: 1차 발사 슬롯 및 PVA 협력 현황 ─────────────────────────
PVA_PARTNER = {
    "headers": ["발사 슬롯", "일정", "패키징(PVA) 수행기관", "국가"],
    "rows": [
        ["1차", "'27년 2월", "University of Potsdam", "독일"],
    ],
}


def build(path):
    doc = S.new_document()
    S.title(doc, "제조본부장 현장 경영지시사항 보고")

    # 1. PVA 및 셀 설계 전략 구체화
    S.h1(doc, "PVA 및 셀 설계 전략 구체화")

    # 1) PVA 확보 전략
    S.h2(doc, "PVA 확보 전략")

    S.h3(doc, "PVA 소싱 및 내재화 계획")
    S.h4(doc, "국내 PVA 국산화 인프라 한계로 초기 단계인 '28년까지 BOM 및 외주 소싱을 기반으로 "
              "사양·표준 정합성을 확보하고, '29년 flight heritage 단계 진입 시점에 맞춰 "
              "내재화 전환을 추진한다.")
    S.shaded_table(doc, PVA_SOURCING["headers"], PVA_SOURCING["rows"], emphasis_rows=[0])

    S.h3(doc, "PVA 협력 업체 현황")
    S.h4(doc, "'27년 2월 1차 발사 계획 및 PVA 협력 현황")
    S.dash(doc, "'27년 2월 예정된 발사 슬롯을 확보하였으며, PVA 패키징 수행기관으로 "
                "University of Potsdam(독일)을 통해 진행 예정.")
    S.shaded_table(doc, PVA_PARTNER["headers"], PVA_PARTNER["rows"])
    S.dash(doc, "국내 업체 솔랩과 협의 진행 계획 중이며, 한서대학교 이훈희 교수와 협력.")

    # 2) BOM 구성 방향
    S.h2(doc, "BOM 구성 방향")
    S.h3(doc, "우주태양광 BOM 구조")
    S.h4(doc, "CIC 5단계 공정 BOM: IV 분류 → 인터커넥터 용접 → 접착 코팅 → "
              "커버글라스 접합 → 경화·검사")
    S.h4(doc, "탠덤 특화 리스크 4개: 봉지재 화학 / 방습 배리어 / 공정온도 상한 / 인터커넥션 방식")
    S.h4(doc, "제조 함의: 항목별 내재화 가능 vs 외주 필요 구분")

    # 3) 셀 설계 방향 및 방사선 차폐 전략
    S.h2(doc, "셀 설계 방향 및 방사선 차폐 전략")
    S.h3(doc, "셀 설계 방향: 사이즈·매수는 인증 캐스케이드의 핵심 병목이며 현재 미확정.")
    S.h3(doc, "방사선 차폐 가설: 커버글라스·배리어와 연계하여 수립.")
    S.h3(doc, "검증 계획: 시험 항목·순서 및 확정 시점 정의.")

    # 4) 종합 로드맵 및 후속 조치
    S.h2(doc, "종합 로드맵 및 후속 조치")
    S.h3(doc, "확정 마일스톤: '27.12 지상 우주환경 인증을 앵커 기준으로 역산하여 관리.")
    S.h3(doc, "의사결정 필요사항: Make/Buy 전환 시점, 셀 사양 확정.")
    S.h3(doc, "정합성 관리 항목: 외주 사양·표준 정합 및 델타 인증 부담 관리.")
    if os.path.exists(ROADMAP):
        S.image(doc, ROADMAP, width_cm=16.5)

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build("/home/user/Report/제조본부장_현장경영지시사항_보고서.docx")
