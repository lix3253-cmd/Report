# -*- coding: utf-8 -*-
"""우주태양광 개발 계획 로드맵(2026~2030) 간트 이미지 생성.

산출물: assets/roadmap.png  (build_report.py가 본문에 삽입)
한글 폰트가 없는 환경이므로 WenQuanYi Zen Hei(Hangul 커버)를 등록해 사용한다.
간트 세부 값은 초안이며 사용자 확정 후 조정한다.
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import Patch

# ── 한글 폰트 등록 ───────────────────────────────────────────────
_FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
if os.path.exists(_FONT_PATH):
    fm.fontManager.addfont(_FONT_PATH)
    plt.rcParams["font.family"] = fm.FontProperties(fname=_FONT_PATH).get_name()
plt.rcParams["axes.unicode_minus"] = False

# ── 색상 ─────────────────────────────────────────────────────────
C_SOL = "#2E75B6"   # 한화솔루션 트랙
C_SYS = "#C55A11"   # 한화시스템 트랙
C_PH = "#BFBFBF"    # Phase 배경

# ── 데이터 (초안) ────────────────────────────────────────────────
# (트랙, 작업명, 시작(연.분기), 종료, 색)   x축은 연도(2026.0 = 26년 초)
TASKS = [
    # 한화솔루션: 셀·CIC·PVA
    ("한화솔루션", "셀 설계/사양 확정",        2026.0, 2027.25, C_SOL),
    ("한화솔루션", "CIC·BOM 구성",            2026.5, 2028.0,  C_SOL),
    ("한화솔루션", "PVA 외주 소싱",            2026.0, 2029.0,  C_SOL),
    ("한화솔루션", "PVA 내재화 전환",          2029.0, 2030.5,  C_SOL),
    # 한화시스템: 우주실증·인증·패키징
    ("한화시스템", "PVA 패키징(U.Potsdam)",    2026.5, 2027.1,  C_SYS),
    ("한화시스템", "1차 발사·궤도 실증",        2027.1, 2028.0,  C_SYS),
    ("한화시스템", "지상 우주환경 인증",         2027.0, 2027.95, C_SYS),
    ("한화시스템", "Flight Heritage 확보",     2028.0, 2029.5,  C_SYS),
]

# 마일스톤: (연.분기, 라벨)
MILES = [
    (2027.1,  "'27.2\n1차 발사"),
    (2027.95, "'27.12\n지상 우주환경 인증"),
    (2029.0,  "'29~\nflight heritage·내재화"),
]

# Phase 구간(배경 음영): (시작, 종료, 라벨)
PHASES = [
    (2026.0, 2027.0, "Phase 0\n사업모델 구체화"),
    (2027.0, 2028.0, "Phase 1\n조건 확정"),
    (2028.0, 2029.0, "Phase 2\n인증·검증"),
    (2029.0, 2030.5, "Phase 3\nCaptive 연동"),
]

X0, X1 = 2026.0, 2030.5


def build(out_path):
    labels = [t[1] for t in TASKS]
    y = list(range(len(TASKS)))[::-1]   # 위→아래 순서

    fig, ax = plt.subplots(figsize=(11.5, 4.6))

    # Phase 배경
    for i, (s, e, lab) in enumerate(PHASES):
        ax.axvspan(s, e, color=C_PH, alpha=0.10 + 0.05 * (i % 2), zorder=0)
        ax.text((s + e) / 2, len(TASKS) - 0.35, lab, ha="center", va="bottom",
                fontsize=8, color="#595959")

    # 작업 막대
    for yi, (track, lab, s, e, col) in zip(y, TASKS):
        ax.barh(yi, e - s, left=s, height=0.5, color=col, alpha=0.9,
                edgecolor="white", zorder=3)

    # 마일스톤 (세로선 + ◆ + 라벨)
    for x, lab in MILES:
        ax.axvline(x, color="#C00000", ls="--", lw=1.1, zorder=2)
        ax.scatter([x], [len(TASKS) - 0.9], marker="D", s=55, color="#C00000",
                   zorder=4)
        ax.text(x, -1.15, lab, ha="center", va="top", fontsize=7.5,
                color="#C00000")

    # 축 서식
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_ylim(-1.6, len(TASKS) + 0.3)
    ax.set_xlim(X0, X1)
    ax.set_xticks([2026, 2027, 2028, 2029, 2030])
    ax.set_xticklabels(["2026", "2027", "2028", "2029", "2030"], fontsize=9)
    ax.set_axisbelow(True)
    ax.xaxis.grid(True, color="#E0E0E0", lw=0.6)
    for spine in ("top", "right", "left"):
        ax.spines[spine].set_visible(False)

    # 범례
    ax.legend(handles=[Patch(color=C_SOL, label="한화솔루션 (셀·CIC·PVA)"),
                       Patch(color=C_SYS, label="한화시스템 (우주실증·인증)")],
              loc="lower right", fontsize=8, framealpha=0.95, ncol=1)

    ax.set_title("우주태양광 개발 계획 로드맵 ('26~'30)", fontsize=12,
                 fontweight="bold", pad=10)
    fig.tight_layout()
    fig.savefig(out_path, dpi=180, bbox_inches="tight")
    print("saved:", out_path)


if __name__ == "__main__":
    here = os.path.dirname(__file__)
    assets = os.path.join(here, "..", "assets")
    os.makedirs(assets, exist_ok=True)
    build(os.path.abspath(os.path.join(assets, "roadmap.png")))
