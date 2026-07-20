# -*- coding: utf-8 -*-
"""
제조본부장 현장 경영지시사항 보고서 (우주태양광) 생성기.
DATA만 채우면 최종 Word가 나온다.  구조: 개요표 + 지시사항 본문(라벨형) + 종합계획.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import report_style as S

# ── 1) 상단 개요 ────────────────────────────────────────────────
OVERVIEW = [
    ("일   시", "[방문 일시 기입]"),
    ("장   소", "[라인/공장 기입]"),
    ("지 시 자", "제조본부장"),
    ("참 석 자", "[참석자 기입]"),
    ("작   성", "[작성자] / [작성일]"),
]

# ── 2) 지시사항 본문 (라벨형) ───────────────────────────────────
# 각 항목: title + (지시내용)/(현황)/(조치계획: 리스트)/(일정·담당)
DIRECTIVES = [
    {
        "title": "PVA 및 셀 설계 전략 구체화",
        "지시내용": "PVA 및 셀 설계 전략을 구체화할 것",
        "현황": "[현재 진행 현황 기입]",
        "조치계획": [
            "PVA 관련 협력업체 및 BOM 구성 방향 구체화",
            "셀 디자인 방향 및 방사선 차폐에 대한 가설 수립, 전략 및 검증 계획 마련",
        ],
        "일정담당": "[목표 일정 / 담당]",
    },
    # ← 나머지 지시사항(1~N)을 같은 형식으로 추가
]

# ── 3) 종합/향후계획 ────────────────────────────────────────────
CLOSING = [
    "[핵심 지시사항별 목표 일정·마일스톤 요약 기입]",
]


def build(path):
    doc = S.new_document()

    S.title(doc, "제조본부장 현장 경영지시사항 보고")
    S.info_table(doc, OVERVIEW)
    doc.add_paragraph()  # 개요표와 본문 사이 간격

    # 지시사항 본문
    for i, d in enumerate(DIRECTIVES, 1):
        S.h1(doc, i, d["title"])
        if d.get("지시내용"):
            S.dash_label(doc, "지시내용", d["지시내용"])
        if d.get("현황"):
            S.dash_label(doc, "현황", d["현황"])
        plans = d.get("조치계획") or []
        if plans:
            S.dash_label(doc, "조치계획", "")
            for p in plans:
                S.subbullet(doc, p)
        if d.get("일정담당"):
            S.dash_label(doc, "일정·담당", d["일정담당"])

    # 종합계획
    S.h1(doc, len(DIRECTIVES) + 1, "종합 및 향후계획")
    for c in CLOSING:
        S.dash(doc, c)

    doc.save(path)
    print("saved:", path)


if __name__ == "__main__":
    build("/home/user/Report/제조본부장_현장경영지시사항_보고서.docx")
