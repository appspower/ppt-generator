"""PPT Generator CLI — Phase B 운영 진입점.

사용
----
    # JSON 입력 파일로 빌드
    python run.py --input my_scenario.json --output out/my_deck.pptx

    # 사전 정의된 시나리오 빌드 (benchmark_5_scenarios.SCENARIO_CONTENT의 키)
    python run.py --preset consulting_proposal_30

    # 시나리오 목록
    python run.py --list-presets

JSON 입력 형식 ([docs/PROJECT_DIRECTION.md] §3 narrative role 13개)
-------------------------------------------------------------
{
  "scenario_name": "우리 회사 전략",
  "skeleton_id": "executive_strategy_40",      # OR narrative_sequence
  "narrative_sequence": ["opening", "situation", "recommendation", "closing"],
  "content_by_role": {
    "opening":        ["발표 제목"],
    "situation":      ["현재 매출 1조원 / YoY +5%"],
    "recommendation": ["3개 신사업 동시 추진"],
    "closing":        ["6월 의사결정 권고"]
  },
  "chart_data": {
    "evidence": {
      "categories": ["Q1", "Q2", "Q3", "Q4"],
      "series": [{"name": "매출", "values": [100, 120, 150, 180]}]
    }
  },
  "metadata": {"title": "보고서", "client": "ACME", "date": "2026-04-26"}
}
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Windows cp949 콘솔에서 한글/em-dash 출력 안전하게: stdout/stderr UTF-8 강제
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pydantic import ValidationError

from ppt_builder.api import BuildError, build_scenario
from ppt_builder.models import ScenarioInput


def _load_input(path: Path) -> ScenarioInput:
    if not path.exists():
        raise SystemExit(f"입력 파일 없음: {path}")
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SystemExit(f"JSON 파싱 실패 [{path}]: {e}")
    try:
        return ScenarioInput.model_validate(raw)
    except ValidationError as e:
        lines = ["입력 검증 실패:"]
        for err in e.errors():
            loc = ".".join(str(p) for p in err["loc"])
            lines.append(f"  [{loc}] {err['msg']}")
        raise SystemExit("\n".join(lines))


def _load_preset(preset_id: str) -> ScenarioInput:
    """benchmark_5_scenarios.SCENARIO_CONTENT를 ScenarioInput으로 변환."""
    sys.path.insert(0, str(ROOT / "scripts"))
    from benchmark_5_scenarios import SCENARIO_CONTENT  # noqa: E402

    if preset_id not in SCENARIO_CONTENT:
        avail = sorted(SCENARIO_CONTENT.keys())
        raise SystemExit(
            f"preset '{preset_id}' 없음. 사용 가능:\n  - "
            + "\n  - ".join(avail)
        )
    raw = SCENARIO_CONTENT[preset_id]
    return ScenarioInput.model_validate(raw)


def _list_presets() -> int:
    sys.path.insert(0, str(ROOT / "scripts"))
    from benchmark_5_scenarios import SCENARIO_CONTENT  # noqa: E402
    print("사용 가능한 preset:")
    for k, v in SCENARIO_CONTENT.items():
        n_roles = len(v.get("content_by_role", {}))
        print(f"  {k:35} ({n_roles} roles) — {v.get('scenario_name', '')}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="run.py",
        description="컨설팅 그레이드 PPT 생성 (Mode A + N1-Lite 하이브리드)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--input", "-i", type=Path,
        help="ScenarioInput JSON 파일",
    )
    parser.add_argument(
        "--preset", "-p", type=str,
        help="사전 정의된 시나리오 ID (--list-presets로 목록 확인)",
    )
    parser.add_argument(
        "--output", "-o", type=Path,
        help="출력 .pptx 경로 (미지정 시 output/decks/{name}.pptx)",
    )
    parser.add_argument(
        "--list-presets", action="store_true",
        help="사용 가능한 preset 목록 출력 후 종료",
    )
    parser.add_argument(
        "--quiet", "-q", action="store_true",
        help="진행 메시지 최소화",
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="plan 덤프 (어떤 슬라이드가 선정되었는지)",
    )
    parser.add_argument(
        "--render-pngs", action="store_true",
        help="빌드 후 PowerPoint COM으로 슬라이드 PNG 렌더 (Windows 전용)",
    )
    args = parser.parse_args(argv)

    if args.list_presets:
        return _list_presets()

    if not args.input and not args.preset:
        parser.error("--input 또는 --preset 중 하나 필요")
    if args.input and args.preset:
        parser.error("--input과 --preset은 동시 사용 불가")

    if args.input:
        scenario = _load_input(args.input)
    else:
        scenario = _load_preset(args.preset)

    if not args.quiet:
        print(f"빌드 시작: {scenario.scenario_name}")
        if scenario.skeleton_id:
            print(f"  skeleton: {scenario.skeleton_id}")
        if scenario.narrative_sequence:
            print(f"  narrative (override): {len(scenario.narrative_sequence)} steps")
        if scenario.chart_data:
            print(f"  chart_data: {sorted(scenario.chart_data.keys())}")

    try:
        result = build_scenario(scenario, output=args.output)
    except BuildError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 2

    print(f"OK {result.summary_line()}")
    print(f"   -> {result.pptx}")
    if result.chart_injected:
        ok = sum(1 for v in result.chart_injected.values() if v)
        miss = [r for r, v in result.chart_injected.items() if not v]
        if miss:
            print(f"   chart_data 미주입 role: {miss} "
                  f"(차트 슬라이드를 못 찾았거나 주입 실패)")
        else:
            print(f"   차트 {ok}건 주입 완료")

    if args.verbose:
        print()
        print("plan dump:")
        for i, p in enumerate(result.plan, 1):
            mode = p.get("mode", "?")
            role = p.get("role", "?")
            if mode == "mode_a":
                detail = (
                    f"sidx={p.get('slide_index')} "
                    f"src={p.get('source')} arch={p.get('archetype')}"
                )
            else:
                detail = (
                    f"comp={p.get('component_id')} family={p.get('family')} "
                    f"slots={p.get('n_text_slots')}"
                )
            print(f"  step {i:2d} | {role:14} | {mode:7} | {detail}")

    if args.render_pngs:
        print("PNG 렌더 중...")
        sys.path.insert(0, str(ROOT / "scripts"))
        from benchmark_5_scenarios_v6 import render_pngs  # noqa: E402
        png_dir = result.pptx.with_suffix("").with_name(result.pptx.stem + "_pngs")
        try:
            pngs = render_pngs(result.pptx.resolve(), png_dir.resolve())
            print(f"  {len(pngs)} pngs -> {png_dir}")
        except Exception as e:  # noqa: BLE001
            print(f"  [warn] PNG 렌더 실패: {type(e).__name__}: {e}",
                  file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
