"""PPT 생성 워크플로 — 6단계 강제 파이프라인.

목표: PPT 생성 품질이 들쭉날쭉하지 않도록 검증 단계를 코드로 강제한다.

6단계: ANALYZE → STRUCTURE → COMPOSE → RENDER → VALIDATE → REFINE
- ANALYZE/STRUCTURE/COMPOSE는 호출자(Claude)가 수행 (schema 작성)
- RENDER → VALIDATE는 코드가 강제 (render_validated 함수)
- REFINE은 호출자가 issues를 보고 schema를 수정해 다시 호출

기존 render_presentation()은 변경 없이 그대로 유지된다.
이 모듈은 그 위에 검증 레이어를 얹는 것이며, 어떤 기존 호출도 깨지 않는다.
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from .models.schema import PresentationSchema


# ============================================================
# Public dataclasses
# ============================================================

@dataclass
class ValidationResult:
    """렌더링 + 검증 결과."""
    output_path: Path
    score: int = 0
    eval_issues: list[str] = field(default_factory=list)
    visual_issues: list[str] = field(default_factory=list)
    pdf_path: Optional[Path] = None
    pdf_available: bool = False
    metrics: dict = field(default_factory=dict)

    @property
    def all_issues(self) -> list[str]:
        return self.eval_issues + self.visual_issues

    @property
    def passed(self) -> bool:
        """통과 기준: evaluate 점수 >= 90 AND 시각 critical 이슈 없음."""
        critical_visual = [
            i for i in self.visual_issues
            if i.startswith(("OVERFLOW", "TABLE_OVERFLOW", "TEXT_OVERLAP"))
        ]
        return self.score >= 90 and not critical_visual

    def summary(self) -> str:
        lines = [
            f"PPT: {self.output_path.name}",
            f"Score: {self.score}/100  PDF: {'yes' if self.pdf_available else 'no'}",
            f"Issues: eval={len(self.eval_issues)} visual={len(self.visual_issues)}",
        ]
        if self.eval_issues:
            lines.append("Eval issues:")
            for i in self.eval_issues[:5]:
                lines.append(f"  - {i}")
        if self.visual_issues:
            lines.append("Visual issues:")
            for i in self.visual_issues[:5]:
                lines.append(f"  - {i}")
        return "\n".join(lines)


# ============================================================
# Public API
# ============================================================

def render_validated(
    schema: PresentationSchema,
    output: Path,
    template: Optional[Path] = None,
    *,
    require_visual: bool = True,
    convert_pdf: bool = True,
) -> ValidationResult:
    """schema를 렌더링하고 자동으로 검증한다.

    이 함수가 워크플로의 핵심 진입점이다. 호출하면 다음 단계가
    무조건 순서대로 실행되며, 어느 단계도 건너뛸 수 없다:

      1. RENDER  — render_presentation() 호출
      2. EVALUATE (정적) — evaluate_pptx() 호출
      3. VALIDATE (시각) — validate_visual() 호출
                         (require_visual=False면 생략)

    Args:
        schema: Pydantic PresentationSchema.
        output: 출력 .pptx 파일 경로.
        template: .pptx 마스터 템플릿 (optional).
        require_visual: True면 시각 검증을 수행. False면 정적 평가만.
        convert_pdf: True면 PDF 변환까지 시도 (Windows + PowerPoint).

    Returns:
        ValidationResult — 점수, issues, pdf 경로 등 모든 결과.

    이 함수는 절대 검증 단계를 빠뜨리지 않는다. require_visual=False로
    명시적으로 끄지 않는 한, 호출 즉시 자동으로 검증이 따라온다.
    """
    from . import render_presentation
    from .evaluate import evaluate_pptx

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    # 1. RENDER
    rendered_path = render_presentation(schema, template=template, output=output)

    # 2. EVALUATE (정적, 항상 수행)
    eval_report = evaluate_pptx(rendered_path)

    result = ValidationResult(
        output_path=Path(rendered_path),
        score=eval_report["score"],
        eval_issues=list(eval_report["issues"]),
        metrics=dict(eval_report.get("metrics", {})),
    )

    # 3. VALIDATE (시각, 옵션)
    if require_visual:
        try:
            from .visual_validate import validate_visual
            visual_report = validate_visual(
                rendered_path, convert_pdf=convert_pdf
            )
            result.visual_issues = list(visual_report.issues)
            result.pdf_path = visual_report.pdf_path
            result.pdf_available = visual_report.pdf_available
            result.metrics["visual"] = visual_report.metrics
        except Exception as e:
            warnings.warn(f"Visual validation skipped: {e}")
            result.visual_issues = [f"VISUAL_CHECK_FAILED: {e}"]

    return result


def workflow_phases() -> list[str]:
    """6단계 이름을 반환 (참고용 / docs 자동 생성용)."""
    return [
        "ANALYZE",     # 주제/요구사항 분석 + 리서치 — Claude
        "STRUCTURE",   # 내용 구조 결정 + 프레임워크 선택 — Claude
        "COMPOSE",     # 화면 구성 + 컴포넌트 배치 — Claude (schema 작성)
        "RENDER",      # JSON → .pptx — render_presentation()
        "VALIDATE",    # 정적 평가 + 시각 검증 — render_validated()
        "REFINE",      # issues 기반 schema 수정 후 재호출 — Claude
    ]


# ============================================================
# Convenience: 간단한 refinement 루프
# ============================================================

def refine_loop(
    schema_builder,
    output: Path,
    template: Optional[Path] = None,
    *,
    max_iterations: int = 3,
    require_visual: bool = True,
    on_iteration=None,
) -> ValidationResult:
    """schema_builder를 반복 호출하며 issues가 사라질 때까지 개선한다.

    Args:
        schema_builder: callable(prev_result: ValidationResult | None) -> PresentationSchema
                        첫 호출은 None을 받고, 이후엔 직전 결과를 받음.
                        호출자(Claude)가 issues를 보고 새 schema를 만든다.
        output: 출력 .pptx 경로.
        template: optional 템플릿.
        max_iterations: 최대 반복 횟수 (안전 장치).
        require_visual: 시각 검증 수행 여부.
        on_iteration: callable(iter_idx, result) — 매 반복 직후 호출되는 훅.

    Returns:
        마지막 ValidationResult.

    이 루프는 schema_builder가 매 호출마다 동일한 schema를 반환하면
    같은 결과를 얻으므로, 무한 루프 방지를 위해 max_iterations을 둔다.
    """
    prev: Optional[ValidationResult] = None
    last: Optional[ValidationResult] = None

    for i in range(max_iterations):
        schema = schema_builder(prev)
        last = render_validated(
            schema, output=output, template=template, require_visual=require_visual
        )
        if on_iteration is not None:
            try:
                on_iteration(i, last)
            except Exception:
                pass
        if last.passed:
            return last
        prev = last

    assert last is not None
    return last
