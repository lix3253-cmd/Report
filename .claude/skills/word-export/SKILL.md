---
name: word-export
description: 보고서(주간보고, 월간보고)를 다른 Word 파일에 붙여넣을 docx로 만든다. Claude Docs 문서나 보고서 본문을 Word로 내보내 달라는 요청, "word export", "붙여넣기용 docx", 번호나 표 서식이 붙여넣기에서 깨진다는 요청에 사용.
---

# Word 붙여넣기용 docx 만들기

Claude Docs의 기본 Export(Word)는 번호를 글자로 내보내서 다른 Word 파일에 붙여넣으면 자동 번호가 아니다.
이 Skill은 번호를 Word 다단계 자동 번호로, 표를 붙여넣어도 눌리지 않는 서식으로 만든다.

## 도구
| 파일 | 역할 |
|---|---|
| `docx_report.py` | 입력(md 또는 HTML) → docx. 번호, 표, 서식 모두 처리 |
| `paste_test.py` | 줄 간격 고정 8pt 파일에 붙여넣은 상황을 만들어 PNG로 렌더링 |
| `example_weekly.md` | md 입력 예시 (11/16 주간보고) |

## 서식 기준 (주간, 월간 공통)
- 본문: 바탕 12pt, 줄 간격 1줄 / 표: 바탕 11pt, 줄 간격 1줄 / 여백 2cm
- 번호: `1.` → `1)` → `①` → `a)` → `i)`를 Word 자동 번호로 넣음 (스타일 "보고 1단계"\~"보고 5단계"). `-` 줄은 바로 위 번호 아래 불릿
- 표: 직접 테두리, 머리행 음영, 셀 세로 가운데, 행 높이 최소 0.7cm, 문서 격자 맞춤 해제. 표 제목은 다음 표와 같은 페이지에 둠
- 기준을 바꿔야 하면 `docx_report.py` 맨 위 상수(FONT, BODY_SIZE, TABLE_SIZE, LEVELS)만 고친다.

## 절차
1. **입력 만들기** (둘 중 하나)
   - 사용자가 Claude Docs에서 Export → HTML 파일을 저장소에 올린 경우: 그 파일을 그대로 입력으로 쓴다.
   - 그 외: Claude Docs 문서를 `read`로 읽고, 본문을 scratchpad에 md로 **글자 그대로** 옮긴다. 문장을 다듬지 않는다.
     md 규칙은 `docx_report.py` 머리말 참고 (`^^` 세로 병합, `<<` 가로 병합, `<br>` 셀 안 줄바꿈, `**굵게**`).
   - export 도구의 base64 결과를 손으로 옮겨 적지 않는다 (한 글자만 틀려도 파일이 깨짐).
2. **변환**
   ```bash
   python3 .claude/skills/word-export/docx_report.py <입력> <출력.docx> --widths "<표1 열 너비cm>;<표2 ...>"
   ```
   열 너비 합은 17cm 이내. 생략하면 균등 분할.
3. **확인**
   ```bash
   python3 .claude/skills/word-export/paste_test.py <출력.docx> <scratchpad 폴더>
   ```
   PNG를 열어 번호 단계, 소절이 바뀔 때 하위 번호 재시작, 표 행 높이와 글자 겹침을 확인한다.
4. **전달**: 파일을 보내고 아래 안내를 함께 준다. 파일은 저장소 루트에 `<보고서명>_<MMDD>.docx`로 두고 커밋한다.

## 사용자에게 줄 붙여넣기 안내
- 붙여넣기 옵션은 "원본 서식 유지"를 고른다. "텍스트만 유지"는 번호와 표 서식이 빠진다.
- 그래도 표가 눌리면 표 전체 선택 → 단락 → 줄 간격 "1줄", "문서 격자가 설정되어 있으면 격자에 맞춤" 해제.
