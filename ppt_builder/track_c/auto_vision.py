"""Claude API Vision 자동 평가 모듈.

슬라이드 PNG를 Claude API에 전송하고, 8항목 체크리스트 점수 + JSON patches를
structured output(tool_use 강제)으로 수신한다.

리서치 기반 설계 결정:
- 슬라이드 1장씩 개별 호출 (다장 한꺼번에 보내면 후반 품질 저하)
- tool_use + tool_choice forced → JSON 스키마 100% 보장
- 해상도: 1568×1176 (Claude 내부 최대, 추가 축소 없음)
- 체크리스트 점수 방식 (자유 서술 대비 재현성↑, 종료 조건 명확)
- 비용: Sonnet 기준 슬라이드당 ~$0.006

환경 변수:
    ANTHROPIC_API_KEY: Anthropic API 키 (필수)

API 키가 없으면 mock 모드로 동작 (개발/테스트용).
"""

from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any

# ============================================================
# tool_use 스키마 — Claude가 이 형식으로 응답을 강제당한다
# ============================================================

SLIDE_REVIEW_TOOL = {
    "name": "slide_review",
    "description": (
        "슬라이드 1장의 시각 품질을 8항목 체크리스트로 평가하고, "
        "자동 적용 가능한 JSON patches를 생성한다."
    ),
    "input_schema": {
        "type": "object",
        "required": [
            "slide_index",
            "checklist",
            "overall_score",
            "must_fix",
            "issues",
            "strengths",
            "patches",
        ],
        "properties": {
            "slide_index": {
                "type": "integer",
                "description": "슬라이드 번호 (1-based)",
            },
            "checklist": {
                "type": "object",
                "description": "8항목 체크리스트, 각 1~10점",
                "required": [
                    "overflow",
                    "alignment",
                    "whitespace",
                    "font_hierarchy",
                    "color_contrast",
                    "chart_proportion",
                    "info_density",
                    "visual_balance",
                ],
                "properties": {
                    "overflow": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "텍스트 overflow/잘림 없음 (10=완벽)",
                    },
                    "alignment": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "요소 간 정렬 일관성",
                    },
                    "whitespace": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "여백 균형 (사방)",
                    },
                    "font_hierarchy": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "폰트 크기 위계 명확",
                    },
                    "color_contrast": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "색상 대비/가독성",
                    },
                    "chart_proportion": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "차트/도형 비례 적절",
                    },
                    "info_density": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "정보 밀도 (과밀/과소 없음)",
                    },
                    "visual_balance": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 10,
                        "description": "전체 시각적 균형",
                    },
                },
            },
            "overall_score": {
                "type": "integer",
                "minimum": 0,
                "maximum": 100,
                "description": "종합 시각 점수 (0~100)",
            },
            "must_fix": {
                "type": "boolean",
                "description": "반드시 수정해야 하는 이슈가 있는가",
            },
            "issues": {
                "type": "array",
                "items": {
                    "type": "object",
                    "required": ["severity", "category", "description"],
                    "properties": {
                        "severity": {
                            "type": "string",
                            "enum": ["critical", "high", "medium", "low"],
                        },
                        "category": {
                            "type": "string",
                            "enum": [
                                "overflow", "readability", "alignment",
                                "whitespace", "hierarchy", "chart",
                                "title", "color",
                            ],
                        },
                        "description": {"type": "string"},
                    },
                },
            },
            "strengths": {
                "type": "array",
                "items": {"type": "string"},
            },
            "patches": {
                "type": "array",
                "description": (
                    "자동 적용 가능한 JSON patches. "
                    "op: set/delete/append. "
                    "path: 슬래시 분리 (slides/0/title). "
                    "디자인 속성만 수정, 콘텐츠 텍스트 변경 금지."
                ),
                "items": {
                    "type": "object",
                    "required": ["op", "path"],
                    "properties": {
                        "op": {
                            "type": "string",
                            "enum": ["set", "delete", "append", "insert"],
                        },
                        "path": {"type": "string"},
                        "value": {},
                    },
                },
            },
        },
    },
}


# ============================================================
# 시스템 프롬프트
# ============================================================

SYSTEM_PROMPT = """당신은 맥킨지/BCG 수준의 컨설팅 PPT 시각 품질 심사관입니다.

## 역할
슬라이드 PNG 1장을 받아 8항목 체크리스트로 채점하고, 자동 수정 가능한 JSON patches를 작성합니다.

## 평가 기준 (각 1~10점)
1. **overflow**: 텍스트가 박스/카드/슬라이드 경계를 넘지 않는가
2. **alignment**: 컬럼·행·카드 정렬이 일관적인가
3. **whitespace**: 여백이 균형잡힌가 (과다/과소 모두 감점)
4. **font_hierarchy**: 제목·본문·라벨의 폰트 위계가 명확한가 (최소 3단계)
5. **color_contrast**: 배경과 텍스트 대비가 충분한가, accent 남용 없는가
6. **chart_proportion**: 차트/표/도형이 적절한 비율인가
7. **info_density**: 정보가 너무 빽빽하거나 너무 비어있지 않은가
8. **visual_balance**: 전체적으로 시각적 무게가 고르게 분배되는가

## overall_score 산출
8항목 평균 × 10 (반올림). 예: 평균 7.5 → 75점.

## 패치 규칙
- **디자인 속성만 수정** (height_ratio, n_cols, style, size 등)
- **콘텐츠 텍스트 절대 변경 금지** (title, bullets 내용은 건드리지 말 것)
- 텍스트가 넘치면 → height_ratio/font_size 조정 또는 bullet 삭제 권고
- 패치가 불필요하면 빈 배열 []

## must_fix 기준
체크리스트 항목 중 하나라도 ≤4점이면 must_fix=true.

slide_review 도구를 반드시 호출하세요.
"""


# ============================================================
# API 호출
# ============================================================

def _encode_png(png_path: Path) -> str:
    """PNG 파일을 base64 인코딩."""
    return base64.standard_b64encode(png_path.read_bytes()).decode("ascii")


def review_slide(
    png_path: Path,
    slide_index: int,
    schema_context: str = "",
    model: str = "claude-sonnet-4-20250514",
    mock: bool = False,
) -> dict[str, Any]:
    """슬라이드 1장을 Claude Vision으로 평가.

    Args:
        png_path: 슬라이드 PNG 경로
        slide_index: 슬라이드 번호 (1-based)
        schema_context: 원본 스키마 JSON (drift 방지용 앵커)
        model: Claude 모델 ID
        mock: True이면 API 호출 없이 mock 결과 반환

    Returns:
        slide_review tool의 input (dict)
    """
    if mock:
        return _mock_review(slide_index)

    import anthropic

    client = anthropic.Anthropic()  # ANTHROPIC_API_KEY 환경변수 사용

    user_content: list[dict[str, Any]] = [
        {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": _encode_png(Path(png_path)),
            },
        },
        {
            "type": "text",
            "text": (
                f"슬라이드 {slide_index}번을 평가하세요.\n\n"
                + (
                    f"## 원본 스키마 (참고용 — 콘텐츠 변경 금지)\n```json\n{schema_context}\n```\n\n"
                    if schema_context
                    else ""
                )
                + "slide_review 도구를 호출하여 평가 결과를 반환하세요."
            ),
        },
    ]

    response = client.messages.create(
        model=model,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        tools=[SLIDE_REVIEW_TOOL],
        tool_choice={"type": "tool", "name": "slide_review"},
        messages=[{"role": "user", "content": user_content}],
    )

    # tool_use 블록에서 input 추출
    for block in response.content:
        if block.type == "tool_use" and block.name == "slide_review":
            return block.input

    raise RuntimeError(
        f"Claude API가 slide_review tool을 호출하지 않았습니다.\n"
        f"Response: {response.content}"
    )


def review_all_slides(
    png_paths: list[Path],
    schema_path: Path | None = None,
    model: str = "claude-sonnet-4-20250514",
    mock: bool = False,
) -> dict[str, Any]:
    """전체 슬라이드를 순차 평가하고 종합 review_response를 반환.

    Returns:
        vision_review.py의 review_response.json 형식과 호환되는 dict
    """
    schema_context = ""
    if schema_path and Path(schema_path).exists():
        schema_context = Path(schema_path).read_text(encoding="utf-8")

    slides: list[dict[str, Any]] = []
    all_patches: list[dict[str, Any]] = []
    total_score = 0

    for i, png_path in enumerate(png_paths):
        slide_index = i + 1
        result = review_slide(
            png_path=png_path,
            slide_index=slide_index,
            schema_context=schema_context,
            model=model,
            mock=mock,
        )
        slides.append(result)
        all_patches.extend(result.get("patches", []))
        total_score += result.get("overall_score", 0)

    n = len(slides) or 1
    overall_score = round(total_score / n)
    must_fix_any = any(s.get("must_fix", False) for s in slides)

    # 종합 요약
    all_issues = []
    for s in slides:
        for iss in s.get("issues", []):
            iss_copy = dict(iss)
            iss_copy["slide_index"] = s.get("slide_index", 0)
            all_issues.append(iss_copy)

    return {
        "schema_path": str(schema_path) if schema_path else "",
        "iter_num": 0,  # caller가 설정
        "overall_score": overall_score,
        "overall_summary": (
            f"{n}장 슬라이드 자동 Vision 평가 완료. "
            f"평균 {overall_score}/100. "
            f"{'수정 필요' if must_fix_any else '수정 불요'}."
        ),
        "slides": slides,
        "patches": all_patches,
        "next_actions": [
            f"critical/high 이슈 {sum(1 for i in all_issues if i.get('severity') in ('critical','high'))}건 우선 수정"
        ] if all_issues else [],
    }


# ============================================================
# Mock 모드 (API 키 없이 테스트용)
# ============================================================

def _mock_review(slide_index: int) -> dict[str, Any]:
    """API 호출 없이 합리적인 mock 결과를 반환."""
    return {
        "slide_index": slide_index,
        "checklist": {
            "overflow": 8,
            "alignment": 9,
            "whitespace": 8,
            "font_hierarchy": 8,
            "color_contrast": 9,
            "chart_proportion": 8,
            "info_density": 7,
            "visual_balance": 8,
        },
        "overall_score": 81,
        "must_fix": False,
        "issues": [
            {
                "severity": "medium",
                "category": "hierarchy",
                "description": f"[MOCK] 슬라이드 {slide_index}: 본문 폰트 10pt 미만 의심",
            }
        ],
        "strengths": [
            f"[MOCK] 슬라이드 {slide_index}: 액션 타이틀 형식 양호",
            f"[MOCK] 슬라이드 {slide_index}: 색상 배분 절제됨",
        ],
        "patches": [],
    }
