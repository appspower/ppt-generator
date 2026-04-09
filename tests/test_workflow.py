"""workflow + visual_validate 모듈 테스트.

목표: Layer 1+3 추가가 기존 흐름을 깨지 않으며, 새 함수가 정상 동작함을 확인.

pytest가 있으면 `pytest tests/test_workflow.py`로,
없으면 `python tests/test_workflow.py`로 직접 실행 가능.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

# pytest는 선택적 의존성 — 없으면 직접 실행 모드로 동작
try:
    import pytest  # type: ignore
    HAS_PYTEST = True
except ImportError:
    HAS_PYTEST = False

# 프로젝트 루트를 sys.path에 추가 (직접 실행 시)
_repo_root = Path(__file__).resolve().parent.parent
if str(_repo_root) not in sys.path:
    sys.path.insert(0, str(_repo_root))

from ppt_builder import render_presentation
from ppt_builder.models.schema import PresentationSchema
from ppt_builder.workflow import (
    ValidationResult,
    refine_loop,
    render_validated,
    workflow_phases,
)
from ppt_builder.visual_validate import (
    VisualCheckUnavailable,
    VisualReport,
    validate_visual,
)


# ============================================================
# Fixtures
# ============================================================

REPO_ROOT = Path(__file__).resolve().parent.parent
EXISTING_PPT_DIR = REPO_ROOT / "output" / "palantir_final"


def _minimal_schema() -> PresentationSchema:
    """4개 영역에 충분한 내용이 있는 단순한 schema."""
    data = {
        "metadata": {"title": "Workflow Test", "client": "Test"},
        "slides": [
            {
                "type": "content",
                "title": "워크플로 검증용 테스트 슬라이드 — 충분한 내용",
                "breadcrumb": "테스트",
                "header_style": "minimal",
                "layout": "columns",
                "n_cols": 2,
                "elements": [
                    {
                        "type": "card",
                        "header": "왼쪽 카드",
                        "subtitle": "테스트 부제목",
                        "style": "accent",
                        "bullets": [
                            "첫 번째 내용 항목입니다",
                            "두 번째 내용 항목입니다",
                            "세 번째 내용 항목입니다",
                            "네 번째 내용 항목입니다",
                        ],
                        "col": 0,
                    },
                    {
                        "type": "card",
                        "header": "오른쪽 카드",
                        "subtitle": "다른 부제목",
                        "style": "default",
                        "bullets": [
                            "첫 번째 비교 항목입니다",
                            "두 번째 비교 항목입니다",
                            "세 번째 비교 항목입니다",
                            "네 번째 비교 항목입니다",
                        ],
                        "col": 1,
                    },
                    {
                        "type": "takeaway_bar",
                        "message": "이것은 두 카드의 핵심 결론을 정리한 메시지입니다",
                        "style": "dark",
                    },
                ],
                "footnote": "테스트 푸트노트",
            }
        ],
    }
    return PresentationSchema.model_validate(data)


# ============================================================
# 기본 회귀: 기존 render_presentation은 그대로 동작해야 함
# ============================================================

def test_existing_render_presentation_still_works(tmp_path: Path):
    schema = _minimal_schema()
    out = tmp_path / "regression.pptx"
    result = render_presentation(schema, output=out)
    assert result.exists()
    assert result.stat().st_size > 0


# ============================================================
# Layer 3: visual_validate
# ============================================================

def test_validate_visual_static_only(tmp_path: Path):
    """정적 분석만 — PDF 변환 없이도 동작해야 함."""
    schema = _minimal_schema()
    out = tmp_path / "static_check.pptx"
    render_presentation(schema, output=out)

    report = validate_visual(out, convert_pdf=False)
    assert isinstance(report, VisualReport)
    assert report.pdf_available is False
    assert isinstance(report.issues, list)
    assert "slide_size" in report.metrics
    assert report.metrics["slide_count"] == 1


def test_validate_visual_missing_file(tmp_path: Path):
    """존재하지 않는 파일에 대해 예외 대신 issues로 반환."""
    report = validate_visual(tmp_path / "nope.pptx", convert_pdf=False)
    assert any("FILE_NOT_FOUND" in i for i in report.issues)


def test_validate_visual_severity_count(tmp_path: Path):
    schema = _minimal_schema()
    out = tmp_path / "severity.pptx"
    render_presentation(schema, output=out)
    report = validate_visual(out, convert_pdf=False)
    counts = report.severity_count()
    assert set(counts.keys()) >= {"OVERFLOW", "TABLE", "OVERLAP", "EMPTY", "OTHER"}


# ============================================================
# Layer 1: workflow.render_validated
# ============================================================

def test_render_validated_basic(tmp_path: Path):
    """render_validated가 기본 동작하고 ValidationResult를 반환해야 함."""
    schema = _minimal_schema()
    out = tmp_path / "validated.pptx"
    result = render_validated(schema, output=out, require_visual=False)

    assert isinstance(result, ValidationResult)
    assert result.output_path.exists()
    assert result.score >= 0
    assert isinstance(result.eval_issues, list)
    assert result.visual_issues == []  # require_visual=False
    assert result.pdf_available is False


def test_render_validated_with_visual_static(tmp_path: Path):
    """require_visual=True지만 convert_pdf=False — 정적 검사만 수행."""
    schema = _minimal_schema()
    out = tmp_path / "validated_static.pptx"
    result = render_validated(
        schema, output=out, require_visual=True, convert_pdf=False
    )
    assert result.output_path.exists()
    assert isinstance(result.visual_issues, list)
    assert result.pdf_available is False
    # metrics에 visual sub-dict가 들어있어야 함
    assert "visual" in result.metrics


def test_validation_result_passed_property(tmp_path: Path):
    schema = _minimal_schema()
    out = tmp_path / "passed.pptx"
    result = render_validated(schema, output=out, require_visual=False)
    # passed는 score>=90 AND critical visual 없음
    assert result.passed == (result.score >= 90)


def test_validation_result_summary_string(tmp_path: Path):
    schema = _minimal_schema()
    out = tmp_path / "summary.pptx"
    result = render_validated(schema, output=out, require_visual=False)
    s = result.summary()
    assert isinstance(s, str)
    assert "PPT:" in s
    assert "Score:" in s


# ============================================================
# refine_loop
# ============================================================

def test_refine_loop_calls_builder_and_returns(tmp_path: Path):
    """refine_loop이 schema_builder를 호출하고 결과를 반환해야 함."""
    call_count = {"n": 0}
    base = _minimal_schema()

    def builder(prev):
        call_count["n"] += 1
        return base

    out = tmp_path / "refine.pptx"
    result = refine_loop(
        builder, output=out, max_iterations=2, require_visual=False
    )
    assert call_count["n"] >= 1
    assert result.output_path.exists()


def test_refine_loop_stops_when_passed(tmp_path: Path):
    """passed=True면 max_iterations 전에 멈춰야 함."""
    base = _minimal_schema()
    call_count = {"n": 0}

    def builder(prev):
        call_count["n"] += 1
        return base

    out = tmp_path / "early_stop.pptx"
    result = refine_loop(
        builder, output=out, max_iterations=5, require_visual=False
    )
    # 첫 시도가 통과하면 1번만 호출, 아니면 5번까지
    if result.passed:
        assert call_count["n"] == 1
    else:
        assert call_count["n"] <= 5


# ============================================================
# Misc
# ============================================================

def test_workflow_phases_returns_six_stages():
    phases = workflow_phases()
    assert len(phases) == 6
    assert phases[0] == "ANALYZE"
    assert phases[-1] == "REFINE"


def test_imports_do_not_break_existing_api():
    """workflow/visual_validate import가 기존 모듈을 깨뜨리지 않아야 함."""
    import ppt_builder
    assert hasattr(ppt_builder, "render_presentation")
    # 새 export
    from ppt_builder import render_validated, ValidationResult  # noqa: F401


# ============================================================
# 직접 실행 모드 (pytest 없이)
# ============================================================

def _run_all_tests_directly() -> int:
    """pytest 없이 모든 테스트 함수를 수동 실행한다."""
    import inspect
    import tempfile
    import traceback

    # tmp_path를 인자로 받는 테스트들
    test_fns = [
        (name, fn)
        for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    ]

    passed = failed = 0
    for name, fn in test_fns:
        sig = inspect.signature(fn)
        try:
            if "tmp_path" in sig.parameters:
                with tempfile.TemporaryDirectory() as td:
                    fn(Path(td))
            else:
                fn()
            print(f"PASS  {name}")
            passed += 1
        except Exception as e:
            print(f"FAIL  {name}: {type(e).__name__}: {e}")
            traceback.print_exc()
            failed += 1

    print(f"\n{passed} passed, {failed} failed")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(_run_all_tests_directly())
