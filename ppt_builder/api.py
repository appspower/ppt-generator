"""Phase B 운영 진입점 — ScenarioInput → .pptx 빌드.

[docs/PROJECT_DIRECTION.md] §1 Mode A + N1-Lite 하이브리드의 라이브러리 인터페이스.
benchmark_5_scenarios_v6.py의 빌드 파이프라인을 ScenarioInput 기반으로 래핑한다.

핵심 함수
--------
- build_scenario(scenario: ScenarioInput, *, output: Path | None) -> BuildResult

사용 예
-------
    from ppt_builder.api import build_scenario
    from ppt_builder.models import ScenarioInput, ChartSpec, ChartSeriesSpec

    s = ScenarioInput(
        scenario_name="우리 회사 전략",
        skeleton_id="executive_strategy_40",
        content_by_role={"opening": ["전략 발표"], ...},
        chart_data={"evidence": ChartSpec(categories=[...], series=[...])},
    )
    result = build_scenario(s, output=Path("output/my_deck.pptx"))
    print(result.pptx, result.n_mode_a, result.n_n1_lite)

설계 노트
--------
benchmark_5_scenarios_v6.py는 실험 스크립트지만 v6 파이프라인의 사실상 reference
구현이다. 복사 대신 sys.path 주입으로 재사용한다 — Phase B 운영 hardening 1차 패스.
향후 Phase C 등에서 ppt_builder/build.py로 본격 이전 검토.
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

from .catalog.paragraph_query import ParagraphStore
from .models import ChartSpec, ScenarioInput

_ROOT = Path(__file__).resolve().parent.parent
_SCRIPTS = _ROOT / "scripts"
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from benchmark_5_scenarios import (  # noqa: E402
    CATALOG_PATH,
    load_skeletons,
)
from benchmark_5_scenarios_v6 import (  # noqa: E402
    OUTPUT_ROOT as _BENCHMARK_OUTPUT_ROOT,
    build_pptx_v6,
    load_components_index,
    select_deck_v6,
)


DEFAULT_OUTPUT_ROOT = _BENCHMARK_OUTPUT_ROOT.parent / "decks"


class BuildError(Exception):
    """build_scenario 단계에서 발생한 사용자 친화 에러."""


@dataclass
class BuildResult:
    """build_scenario 산출."""
    pptx: Path
    plan: list[dict]
    edits: list[dict]
    n_mode_a: int
    n_n1_lite: int
    chart_injected: dict[str, bool] = field(default_factory=dict)
    narrative: list[str] = field(default_factory=list)

    def summary_line(self) -> str:
        chart_n = sum(1 for v in self.chart_injected.values() if v)
        return (
            f"{self.pptx.name}: {len(self.plan)} slides "
            f"({self.n_mode_a} mode_a + {self.n_n1_lite} n1_lite)"
            + (f", charts: {chart_n}/{len(self.chart_injected)}"
               if self.chart_injected else "")
        )


def _safe_filename(name: str, max_len: int = 60) -> str:
    """ScenarioInput.scenario_name에서 안전한 파일명을 만든다."""
    keep = []
    for ch in name:
        if ch.isalnum() or ch in (" ", "_", "-", "."):
            keep.append(ch)
        else:
            keep.append("_")
    cleaned = "".join(keep).strip().replace(" ", "_").strip("._")
    return (cleaned[:max_len] or "scenario").strip("._") or "scenario"


def build_scenario(
    scenario: ScenarioInput,
    *,
    output: Path | None = None,
) -> BuildResult:
    """ScenarioInput → .pptx 빌드.

    Args:
        scenario: 검증된 입력 스키마
        output: .pptx 출력 경로. None이면 output/decks/{scenario_name}.pptx 자동.

    Returns:
        BuildResult: 산출 경로 + 빌드 통계 + 차트 주입 결과

    Raises:
        BuildError: skeleton 미존재, 카탈로그 누락, 빌드 실패 등
    """
    # narrative 결정 (narrative_sequence 우선)
    if scenario.narrative_sequence:
        narrative = [str(r) for r in scenario.narrative_sequence]
    else:
        try:
            skeletons = load_skeletons()
        except FileNotFoundError as e:
            raise BuildError(
                f"skeletons.json not found: {e}. catalog 빌드가 필요합니다."
            ) from e
        if scenario.skeleton_id not in skeletons:
            raise BuildError(
                f"skeleton_id '{scenario.skeleton_id}' not found. "
                f"available: {sorted(skeletons.keys())}"
            )
        narrative = list(skeletons[scenario.skeleton_id]["narrative_sequence"])

    # 데이터 로드
    if not CATALOG_PATH.exists():
        raise BuildError(f"catalog 누락: {CATALOG_PATH}")
    labels = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))["labels"]

    try:
        store = ParagraphStore.load()
    except Exception as e:  # noqa: BLE001
        raise BuildError(
            f"ParagraphStore 로드 실패: {type(e).__name__}: {e}"
        ) from e

    try:
        components_index = load_components_index()
    except FileNotFoundError as e:
        raise BuildError(
            f"components_index.json 누락: Phase 2 build_component_library.py "
            f"실행 필요. ({e})"
        ) from e

    components = components_index["components"]
    components_by_id = {c["component_id"]: c for c in components}

    # 빌드 입력 정규화
    scenario_content = {
        "scenario_name": scenario.scenario_name,
        "skeleton_id": scenario.skeleton_id,
        "content_by_role": {
            k: list(v) for k, v in scenario.content_by_role.items()
        },
    }
    chart_roles = set(scenario.chart_data.keys())
    chart_data_dict: dict[str, ChartSpec] = dict(scenario.chart_data)

    # 출력 경로
    if output is None:
        DEFAULT_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
        output = DEFAULT_OUTPUT_ROOT / f"{_safe_filename(scenario.scenario_name)}.pptx"
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    # plan + build
    try:
        plan = select_deck_v6(
            labels, narrative, scenario_content, store, components,
            chart_roles=chart_roles,
        )
    except Exception as e:  # noqa: BLE001
        raise BuildError(f"plan 생성 실패: {type(e).__name__}: {e}") from e

    try:
        build = build_pptx_v6(
            plan, scenario_content, store, components_by_id, output,
            chart_data=chart_data_dict,
        )
    except Exception as e:  # noqa: BLE001
        raise BuildError(f"빌드 실패: {type(e).__name__}: {e}") from e

    return BuildResult(
        pptx=Path(build["pptx"]),
        plan=build["plan_unique"],
        edits=build["edits"],
        n_mode_a=build["n_mode_a"],
        n_n1_lite=build["n_n1_lite"],
        chart_injected=build.get("chart_injected", {}),
        narrative=narrative,
    )


__all__ = [
    "BuildError",
    "BuildResult",
    "DEFAULT_OUTPUT_ROOT",
    "build_scenario",
]
