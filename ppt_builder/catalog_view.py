"""Phase B 카탈로그 viewer — 사용 가능한 자산 한눈에.

[docs/PROJECT_DIRECTION.md] §3 ROLE_MODE_MAP과 N1-Lite 컴포넌트 라이브러리를
콘솔에서 빠르게 조회한다. 사용자가 어떤 role/scenario로 빌드 가능한지,
어떤 컴포넌트가 있는지를 확인하는 진입점.

서브명령
-------
- python -m ppt_builder.catalog_view roles              # 13 role + mode 표
- python -m ppt_builder.catalog_view scenarios          # preset 목록
- python -m ppt_builder.catalog_view components [family]# 33 컴포넌트 + 추출 위치
- python -m ppt_builder.catalog_view summary            # 1줄 요약 + 라이브러리 위치

미리보기 PPTX
-------------
output/component_library/{family}_family.pptx (7개) — 각 family의 후보를
한 슬라이드씩 모아 미리볼 수 있는 .pptx. PowerPoint에서 직접 열어 확인.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Windows cp949 콘솔 회피
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

# scripts/ for ROLE_MODE_MAP + SCENARIO_CONTENT
_SCRIPTS = ROOT / "scripts"
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

INDEX_PATH = ROOT / "output" / "component_library" / "components_index.json"
LIBRARY_DIR = ROOT / "output" / "component_library"


def _load_index() -> dict | None:
    if not INDEX_PATH.exists():
        return None
    return json.loads(INDEX_PATH.read_text(encoding="utf-8"))


def cmd_roles(_args) -> int:
    from benchmark_5_scenarios_v6 import ROLE_MODE_MAP  # noqa: E402

    print("13 narrative role × mode (헌법 §3 ROLE_MODE_MAP):")
    print()
    print(f"  {'role':<14} | {'mode':<8} | 설명")
    print(f"  {'-' * 14}-+-{'-' * 8}-+-{'-' * 40}")
    descriptions = {
        "opening": "표지",
        "agenda": "목차",
        "divider": "섹션 구분",
        "situation": "현황 (N1-Lite 보강)",
        "complication": "문제/난점 (N1-Lite 보강)",
        "evidence": "근거/데이터",
        "analysis": "분석/진단",
        "recommendation": "권고/솔루션",
        "roadmap": "로드맵 (N1-Lite 컴포넌트)",
        "benefit": "효과/혜택 (N1-Lite 보강)",
        "risk": "위험/리스크 (N1-Lite 보강)",
        "closing": "결론/마감",
        "appendix": "부록",
    }
    for role, mode in ROLE_MODE_MAP.items():
        print(f"  {role:<14} | {mode:<8} | {descriptions.get(role, '')}")
    print()
    print("Mode A: 마스터 슬라이드를 통째로 복제 (whole-slide reuse).")
    print("N1-Lite: 마스터의 컴포넌트(도형 그룹)를 빈 슬라이드에 단독 배치.")
    return 0


def cmd_scenarios(_args) -> int:
    from benchmark_5_scenarios import SCENARIO_CONTENT  # noqa: E402

    print(f"사용 가능한 preset 시나리오: {len(SCENARIO_CONTENT)}개")
    print()
    for sid, sc in SCENARIO_CONTENT.items():
        cb = sc.get("content_by_role", {})
        n_items = sum(len(v) for v in cb.values())
        chart_n = len(sc.get("chart_data", {}))
        chart_str = f", charts={chart_n}" if chart_n else ""
        print(f"  {sid}")
        print(f"    이름: {sc.get('scenario_name', '')}")
        print(f"    skeleton: {sc.get('skeleton_id', sid)}")
        print(f"    roles: {len(cb)}, items: {n_items}{chart_str}")
        print()
    print("빌드: python run.py --preset <id> [--output path] [--verbose]")
    return 0


def cmd_components(args) -> int:
    idx = _load_index()
    if idx is None:
        print(f"[ERROR] components_index.json 없음: {INDEX_PATH}", file=sys.stderr)
        print(
            "  → scripts/build_component_library.py 실행 필요",
            file=sys.stderr,
        )
        return 2

    summary = idx.get("summary", {})
    families = summary.get("families", {})
    print(
        f"컴포넌트 라이브러리: {summary.get('total_components', 0)}개 / "
        f"{len(families)}개 family"
    )
    print(f"라이브러리 위치: {LIBRARY_DIR}")
    print()

    components = idx.get("components", [])
    target_family = args.family
    if target_family:
        components = [c for c in components if c.get("family") == target_family]
        if not components:
            print(
                f"[ERROR] family '{target_family}' 없음. "
                f"available: {sorted(families.keys())}",
                file=sys.stderr,
            )
            return 2
        print(f"family={target_family}: {len(components)}개")
        print()

    by_family: dict[str, list[dict]] = {}
    for c in components:
        by_family.setdefault(c.get("family", "?"), []).append(c)

    for fam in sorted(by_family.keys()):
        comps = by_family[fam]
        print(f"  [{fam}] {len(comps)}개   "
              f"미리보기: {LIBRARY_DIR / f'{fam}_family.pptx'}")
        for c in comps:
            cid = c.get("component_id", "?")
            n_slots = len(c.get("slots", []))
            sidx = c.get("source", {}).get("master_slide_index", "?")
            roles = c.get("applicable_roles", [])
            print(
                f"    {cid:<28} slots={n_slots:<2} "
                f"master_sidx={sidx:<5} roles={roles}"
            )
        print()
    return 0


def cmd_summary(_args) -> int:
    from benchmark_5_scenarios_v6 import ROLE_MODE_MAP  # noqa: E402
    from benchmark_5_scenarios import SCENARIO_CONTENT  # noqa: E402

    idx = _load_index()
    n_comp = idx.get("summary", {}).get("total_components", 0) if idx else 0
    n_fam = len(idx.get("summary", {}).get("families", {})) if idx else 0

    n_mode_a = sum(1 for v in ROLE_MODE_MAP.values() if v == "mode_a")
    n_n1l = sum(1 for v in ROLE_MODE_MAP.values() if v == "n1_lite")

    print("PPT Generator — 카탈로그 요약")
    print()
    print(f"  narrative roles: {len(ROLE_MODE_MAP)}개")
    print(f"    Mode A (whole-slide reuse): {n_mode_a}")
    print(f"    N1-Lite (single-component): {n_n1l}")
    print(f"  preset 시나리오: {len(SCENARIO_CONTENT)}개")
    print(f"  N1-Lite 컴포넌트: {n_comp}개 / {n_fam}개 family"
          + ("" if idx else " (라이브러리 미빌드)"))
    print()
    print("상세:")
    print("  python -m ppt_builder.catalog_view roles")
    print("  python -m ppt_builder.catalog_view scenarios")
    print("  python -m ppt_builder.catalog_view components [family]")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="catalog_view",
        description="PPT Generator 카탈로그 viewer",
    )
    sub = parser.add_subparsers(dest="cmd", required=False)

    sub.add_parser("roles", help="13 narrative role × mode")
    sub.add_parser("scenarios", help="preset 시나리오 목록")
    pc = sub.add_parser("components", help="N1-Lite 컴포넌트 라이브러리")
    pc.add_argument(
        "family", nargs="?", default=None,
        help="필터 (chevron/card/timeline/matrix/callout/table/flow)",
    )
    sub.add_parser("summary", help="1줄 요약")

    args = parser.parse_args(argv)
    cmd = args.cmd or "summary"

    handlers = {
        "roles": cmd_roles,
        "scenarios": cmd_scenarios,
        "components": cmd_components,
        "summary": cmd_summary,
    }
    return handlers[cmd](args)


if __name__ == "__main__":
    sys.exit(main())
