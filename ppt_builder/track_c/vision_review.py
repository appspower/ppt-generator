"""Vision 리뷰 인터페이스.

첫 마일스톤에서는 Claude Code가 직접 PNG를 보고 피드백을 작성하는 방식이라,
이 모듈은 "리뷰 요청 파일 생성"과 "피드백 JSON 검증/적재" 두 가지만 제공한다.

향후 확장 시 이 자리에 Anthropic API 호출 코드를 붙일 수 있다.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


REVIEW_PROMPT_TEMPLATE = """# Track C 비주얼 리뷰 요청

**작성**: {timestamp}
**원본 스키마**: `{schema_path}`
**렌더 결과**: `{pptx_path}`
**evaluate.py 점수**: {score}/100 ({status})
**슬라이드 수**: {slide_count}

---

## 리뷰 지시 (Claude Code 전용)

아래 PNG들을 **슬라이드 1번부터 N번까지 순차적으로** 모두 확인하고, 각 슬라이드에 대해 다음 항목을 빠짐없이 평가하세요:

### 체크리스트 (슬라이드별)
1. **Overflow / 잘림**: 텍스트가 박스/카드/슬라이드 경계를 넘는가?
2. **가독성**: 폰트가 너무 작거나 배경과 색이 겹치는가? (흰 배경 + 흰 글씨, 회색 배경 + 회색 글씨 등)
3. **정렬**: 컬럼·행 정렬이 어긋나는가? 카드 높이가 들쭉날쭉한가?
4. **여백**: 하단/우측에 빈 공간이 1.5인치 이상 비어 있는가?
5. **시각 위계**: 제목·본문·라벨의 폰트 위계가 충분한가?
6. **차트/표**: 라벨 겹침, 축 잘림, 행 높이 어긋남이 있는가?
7. **액션 타이틀**: 슬라이드 제목이 인사이트 문장형인가, 라벨형(짧은 단어)인가?
8. **컬러**: accent 색(#FD5108) 남용이 있는가?

### 리뷰 결과 출력 형식

다음 경로에 JSON으로 저장: `{review_response_path}`

**중요**: `patches` 필드를 반드시 포함할 것. 자동 재실행(`apply_and_rerun`)이 이걸 읽어 스키마를 자동 수정한다.

```json
{{
  "schema_path": "{schema_path}",
  "iter_num": {iter_num},
  "overall_score": 0,
  "overall_summary": "한 문단 종합 평가",
  "slides": [
    {{
      "slide_index": 1,
      "score": 0,
      "must_fix": true,
      "issues": [
        {{
          "severity": "critical|high|medium|low",
          "category": "overflow|readability|alignment|whitespace|hierarchy|chart|title|color",
          "description": "구체적으로 어떤 문제인가",
          "fix_suggestion": "JSON에서 어디를 어떻게 고쳐야 하는가"
        }}
      ],
      "strengths": ["잘 된 점들"]
    }}
  ],
  "patches": [
    {{"op": "set", "path": "slides/0/title", "value": "단축된 새 타이틀"}},
    {{"op": "set", "path": "slides/0/sections/1/height_ratio", "value": 0.74}},
    {{"op": "delete", "path": "slides/0/sections/1/elements/0/items/0/bullets/3"}},
    {{"op": "append", "path": "slides/0/sections/1/elements/0/items/1/bullets", "value": "추가 bullet 텍스트"}}
  ],
  "next_actions": [
    "patches 외 사용자 손이 필요한 정성적 액션"
  ]
}}
```

**Patch 문법**:
- `op`: `set` (값 교체) / `delete` (요소 제거) / `append` (배열 끝에 추가) / `insert` (배열 특정 위치 삽입)
- `path`: 슬래시로 분리된 경로. 숫자는 배열 인덱스. 예) `slides/0/sections/1/elements/0/items/0/bullets/3`
- `value`: set/append/insert 시 적용할 값. delete는 생략.

---

## PNG 파일 목록

{png_list}

---

## evaluate.py 자동 점수 이슈

{auto_issues}
"""


def generate_review_request(
    work_dir: Path,
    schema_path: Path,
    pptx_path: Path,
    png_paths: list[Path],
    eval_report: dict[str, Any],
    iter_num: int,
) -> Path:
    """리뷰 요청 마크다운 파일을 생성한다.

    Returns: review_request.md 경로
    """
    work_dir = Path(work_dir)
    review_path = work_dir / "review_request.md"
    response_path = work_dir / "review_response.json"

    png_list = "\n".join(
        f"{i+1}. [{p.name}]({p.as_posix()})"
        for i, p in enumerate(png_paths)
    )

    issues = eval_report.get("issues", [])
    if issues:
        auto_issues = "\n".join(f"- {iss}" for iss in issues)
    else:
        auto_issues = "_(자동 점수 이슈 없음)_"

    status = "PASS" if eval_report.get("pass") else "FAIL"

    content = REVIEW_PROMPT_TEMPLATE.format(
        timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        schema_path=Path(schema_path).as_posix(),
        pptx_path=Path(pptx_path).as_posix(),
        score=eval_report.get("score", 0),
        status=status,
        slide_count=eval_report.get("slide_count", len(png_paths)),
        review_response_path=response_path.as_posix(),
        iter_num=iter_num,
        png_list=png_list,
        auto_issues=auto_issues,
    )

    review_path.write_text(content, encoding="utf-8")
    return review_path


def load_review_response(response_path: Path) -> dict[str, Any]:
    """리뷰 응답 JSON을 로드하고 기본 구조를 검증한다."""
    response_path = Path(response_path)
    if not response_path.exists():
        raise FileNotFoundError(
            f"리뷰 응답 파일이 없습니다: {response_path}\n"
            "Claude Code가 PNG를 검토한 뒤 review_response.json을 생성해야 합니다."
        )

    data = json.loads(response_path.read_text(encoding="utf-8"))

    required = {"schema_path", "iter_num", "overall_score", "slides"}
    missing = required - set(data.keys())
    if missing:
        raise ValueError(f"리뷰 응답에 필수 키 누락: {missing}")

    # patches는 선택 (auto-refine 시 필수, 수동 모드 시 선택)
    if "patches" in data and not isinstance(data["patches"], list):
        raise ValueError("patches 필드는 list 여야 함")

    return data
