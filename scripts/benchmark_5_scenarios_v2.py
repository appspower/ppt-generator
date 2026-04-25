"""Phase A3 Step 3 — paragraph-aware fill 5 벤치마크 재측정.

v1 (Mode A 단독): 슬라이드당 1개 paragraph만 채움 → fill 6.8%
v2 (paragraph fill): role별 슬롯 매핑 → 모든 fillable 슬롯 채움

Usage
-----
python scripts/benchmark_5_scenarios_v2.py [scenario_id]
"""
from __future__ import annotations

import json
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pptx import Presentation

from ppt_builder.template import edit_ops
from ppt_builder.template.editor import TemplateEditor
from ppt_builder.catalog.paragraph_query import ParagraphStore


def _replace_paragraph_universal(slide, slot, new_text: str) -> None:
    """edit_ops.replace_paragraph + table 셀 분기 처리.

    table_header / table_cell은 edit_ops가 못 함 — table API 사용.
    """
    # 일반 shape: edit_ops 사용
    if slot.shape_kind != "TABLE":
        edit_ops.replace_paragraph(slide, slot.flat_idx, slot.paragraph_id, new_text)
        return
    # TABLE 셀
    target_shape = None
    for fi, sh in edit_ops.iter_leaf_shapes(slide):
        if fi == slot.flat_idx:
            target_shape = sh
            break
    if target_shape is None or not target_shape.has_table:
        raise edit_ops.SlideEditError(f"flat_idx={slot.flat_idx} not a table")
    if slot.table_row is None or slot.table_col is None:
        raise edit_ops.SlideEditError("table slot missing row/col")
    cell = target_shape.table.cell(slot.table_row, slot.table_col)
    paras = cell.text_frame.paragraphs
    if slot.paragraph_id >= len(paras):
        raise edit_ops.SlideEditError(
            f"paragraph_id={slot.paragraph_id} out of range in cell"
        )
    para = paras[slot.paragraph_id]
    runs = list(para.runs)
    if runs:
        first, *rest = runs
        for r in rest:
            r._r.getparent().remove(r._r)
        first.text = new_text
    else:
        run = para.add_run()
        run.text = new_text

# v1과 같은 자산 재사용
sys.path.insert(0, str(ROOT / "scripts"))
from benchmark_5_scenarios import (  # noqa: E402
    SCENARIO_CONTENT,
    ARCHETYPE_FALLBACK,
    load_catalog,
    load_skeletons,
    candidates_for_role,
    select_deck,
    _reorder_sldIdLst,
)


CATALOG_PATH = ROOT / "output" / "catalog" / "final_labels.json"
TEMPLATE_PATH = ROOT / "docs" / "references" / "_master_templates" / "PPT 템플릿.pptx"
OUTPUT_ROOT = ROOT / "output" / "benchmark_v2"


# ----------------------------------------------------------------------------
# Content-by-role expansion: scenario 컨텐츠를 role별 list로 확장
# ----------------------------------------------------------------------------

def expand_content_for_slide(role: str, scenario_content: dict,
                             role_use_counter: dict[str, int],
                             slot_capacity: dict[str, int]) -> dict[str, list[str]]:
    """선택된 슬라이드 role + capacity에 맞춰 fill용 content_by_role 생성.

    Title은 항상 채우고, parallel items는 단일 best-fit 슬롯 타입에만 할당
    (slot capacity와 가장 잘 맞는 role 선택). 중복 할당 방지.
    """
    by_role = scenario_content["content_by_role"]
    use_idx = role_use_counter.get(role, 0)
    role_use_counter[role] = use_idx + 1

    role_items = by_role.get(role, [])
    if not role_items:
        role_items = [role.upper()]

    # title: 시나리오명 또는 role별 첫 line
    title = scenario_content["scenario_name"]
    if role in ("opening", "closing", "divider"):
        title = role_items[min(use_idx, len(role_items) - 1)]

    out: dict[str, list[str] | str] = {"title": title}

    # parallel content: 시각 우선순위 슬롯 1개에만 할당
    if role not in ("opening", "closing", "divider", "agenda"):
        # 시각 prominence 우선: chevron > card_header > callout > card_body
        candidates = ["chevron_label", "card_header", "callout_text", "card_body"]
        chosen = None
        for c in candidates:
            if slot_capacity.get(c, 0) > 0:
                chosen = c
                break
        if chosen:
            out[chosen] = role_items

    return out


# ----------------------------------------------------------------------------
# Build with paragraph-level fill
# ----------------------------------------------------------------------------

def build_pptx_v2(plan: list[dict], scenario_content: dict, store: ParagraphStore,
                  pptx_out: Path) -> dict:
    pptx_out.parent.mkdir(parents=True, exist_ok=True)

    seen = set()
    plan_unique = []
    for p in plan:
        if p["slide_index"] is None:
            continue
        if p["slide_index"] in seen:
            continue
        seen.add(p["slide_index"])
        plan_unique.append(p)

    desired_order = [p["slide_index"] for p in plan_unique]
    keep_unique_sorted = sorted(set(desired_order))

    editor = TemplateEditor(TEMPLATE_PATH)
    editor.keep_slides(keep_unique_sorted)
    _reorder_sldIdLst(editor.prs, keep_unique_sorted, desired_order)

    edit_results = []
    role_use_counter: dict[str, int] = {}

    for step_idx, item in enumerate(plan_unique):
        role = item["role"]
        sidx = item["slide_index"]
        step = step_idx + 1

        # 슬라이드의 슬롯 capacity (먼저)
        capacity = store.slot_capacity(sidx)
        # role별 컨텐츠 확장 (capacity-aware)
        content = expand_content_for_slide(role, scenario_content,
                                           role_use_counter, capacity)
        # None 값 제거
        content = {k: v for k, v in content.items() if v}

        # 매핑
        edits = store.match_content(sidx, content)

        # 실제 슬라이드 편집
        target_slide = editor.prs.slides[step_idx]
        n_ok = 0
        n_fail = 0
        per_edit_log = []
        edited_keys: set[tuple[int, int]] = set()
        # role -> 슬롯 인덱스 lookup (universal replace용)
        all_fillable = store.fillable_slots(sidx)
        flat_to_slot: dict[tuple[int, int], object] = {}
        for slist in all_fillable.values():
            for s_obj in slist:
                flat_to_slot[(s_obj.flat_idx, s_obj.paragraph_id)] = s_obj

        for e in edits:
            try:
                slot_obj = flat_to_slot.get((e["flat_idx"], e["paragraph_id"]))
                if slot_obj is None:
                    raise edit_ops.SlideEditError("slot lookup failed")
                _replace_paragraph_universal(target_slide, slot_obj, e["text"])
                n_ok += 1
                edited_keys.add((e["flat_idx"], e["paragraph_id"]))
                per_edit_log.append({
                    **e, "edit_ok": True, "edit_reason": "OK",
                })
            except Exception as ex:
                n_fail += 1
                per_edit_log.append({
                    **e, "edit_ok": False,
                    "edit_reason": f"{type(ex).__name__}: {ex}",
                })

        # 미매칭 fillable 슬롯의 ~~ placeholder 비우기 (시각적 깔끔함)
        n_blanked = 0
        fillable = store.fillable_slots(sidx)
        for role_name, slots in fillable.items():
            for slot in slots:
                key = (slot.flat_idx, slot.paragraph_id)
                if key in edited_keys:
                    continue
                # 원래 텍스트가 ~~ 류 placeholder일 때만 비움
                orig = (slot.text_original or "").strip()
                if orig in {"~~", "~~\x0b", "\x0b~~", "~"} or "~~" in orig:
                    try:
                        _replace_paragraph_universal(target_slide, slot, " ")
                        n_blanked += 1
                    except Exception:
                        pass

        # 통계
        fillable_total = sum(capacity.values())
        edit_results.append({
            "step": step,
            "role": role,
            "slide_index": sidx,
            "source": item["source"],
            "slot_capacity": capacity,
            "n_fillable": fillable_total,
            "n_edit_attempted": len(edits),
            "n_edit_ok": n_ok,
            "n_edit_fail": n_fail,
            "n_blanked": n_blanked,
            "fill_ratio": (n_ok / fillable_total) if fillable_total else 0,
            "visual_resolution_ratio": (
                (n_ok + n_blanked) / fillable_total if fillable_total else 0
            ),
            "n_overflow_truncated": sum(
                1 for e in per_edit_log if e.get("original_overflow")
            ),
            "edits": per_edit_log,
        })

    editor.save(pptx_out)
    editor.cleanup()

    return {"plan_unique": plan_unique, "edits": edit_results, "pptx": str(pptx_out)}


# ----------------------------------------------------------------------------
# PNG render (v1과 동일)
# ----------------------------------------------------------------------------

def render_pngs(pptx_path: Path, png_dir: Path) -> list[Path]:
    import pythoncom
    import win32com.client

    png_dir.mkdir(parents=True, exist_ok=True)
    for old in png_dir.glob("*.png"):
        old.unlink()

    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(
            str(pptx_path.resolve()), ReadOnly=True, Untitled=False, WithWindow=False,
        )
        total = presentation.Slides.Count
        out = []
        for i in range(1, total + 1):
            p = png_dir / f"step_{i:02d}.png"
            try:
                presentation.Slides(i).Export(str(p), "PNG", 1568, 1176)
                out.append(p)
            except Exception as e:
                print(f"  [err] slide {i}: {e}", flush=True)
        return out
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ----------------------------------------------------------------------------
# Metrics
# ----------------------------------------------------------------------------

def compute_v2_metrics(scenario_id: str, plan: list[dict],
                       edits: list[dict], narrative: list[str]) -> dict:
    n_total = len(narrative)
    n_role_hit = sum(1 for p in plan if p["source"] == "role")
    n_archetype = sum(1 for p in plan if p["source"] == "archetype")
    n_reuse = sum(1 for p in plan if p["source"] == "reuse")
    n_none = sum(1 for p in plan if p["source"] == "none")

    fillable_total = sum(e["n_fillable"] for e in edits)
    edited_total = sum(e["n_edit_ok"] for e in edits)
    blanked_total = sum(e["n_blanked"] for e in edits)
    fail_total = sum(e["n_edit_fail"] for e in edits)
    overflow_total = sum(e["n_overflow_truncated"] for e in edits)

    fill_ratio = edited_total / fillable_total if fillable_total else 0
    visual_resolution = (
        (edited_total + blanked_total) / fillable_total if fillable_total else 0
    )
    avg_per_slide_fillable = fillable_total / max(len(edits), 1)
    avg_per_slide_filled = edited_total / max(len(edits), 1)

    # 점수
    score_role = (n_role_hit / n_total) * 100
    score_fill = fill_ratio * 100
    # overflow 안 함 (truncate 처리해서 잘림 없으나 짧아짐 - 부분 페널티)
    score_overflow_penalty = (1 - overflow_total / max(edited_total, 1)) * 100

    score_visual = visual_resolution * 100
    composite_quant = round(
        score_role * 0.25 + score_fill * 0.30 + score_visual * 0.30
        + score_overflow_penalty * 0.15, 1
    )

    return {
        "scenario_id": scenario_id,
        "narrative_length": n_total,
        "metrics": {
            "A_role_match": {
                "role_hit": n_role_hit, "archetype_fallback": n_archetype,
                "reuse_fallback": n_reuse, "miss": n_none,
                "role_hit_pct": round(score_role, 1),
            },
            "B_slot_fill_v2": {
                "fillable_total": fillable_total,
                "filled_total": edited_total,
                "blanked_total": blanked_total,
                "edit_failed": fail_total,
                "fill_ratio": round(fill_ratio, 4),
                "fill_pct": round(score_fill, 1),
                "visual_resolution_ratio": round(visual_resolution, 4),
                "visual_resolution_pct": round(score_visual, 1),
                "avg_fillable_per_slide": round(avg_per_slide_fillable, 1),
                "avg_filled_per_slide": round(avg_per_slide_filled, 1),
            },
            "C_overflow": {
                "n_overflow_truncated": overflow_total,
                "n_edited": edited_total,
                "overflow_rate_pct": round(
                    overflow_total / max(edited_total, 1) * 100, 1
                ),
                "no_overflow_score": round(score_overflow_penalty, 1),
            },
            "composite_quant_score": composite_quant,
        },
    }


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def run_scenario(scenario_id: str, labels: list[dict], skeletons: dict,
                 store: ParagraphStore) -> dict:
    print()
    print("=" * 80)
    print(f"SCENARIO v2: {scenario_id}")
    print("=" * 80)

    sk = skeletons[scenario_id]
    narrative = sk["narrative_sequence"]
    sc = SCENARIO_CONTENT[scenario_id]

    plan = select_deck(labels, narrative)
    n_role = sum(1 for p in plan if p["source"] == "role")
    n_arch = sum(1 for p in plan if p["source"] == "archetype")
    print(f"  retrieval: role={n_role} archetype={n_arch}")

    out_dir = OUTPUT_ROOT / scenario_id
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx_out = out_dir / "deck.pptx"
    png_dir = out_dir / "pngs"

    t0 = time.time()
    build = build_pptx_v2(plan, sc, store, pptx_out)
    dt_build = time.time() - t0
    print(f"  build: {dt_build:.1f}s")

    t0 = time.time()
    pngs = render_pngs(pptx_out, png_dir)
    dt_render = time.time() - t0
    print(f"  render: {dt_render:.1f}s, {len(pngs)} pngs")

    metrics = compute_v2_metrics(scenario_id, plan, build["edits"], narrative)
    metrics["paths"] = {
        "pptx": str(pptx_out),
        "png_dir": str(png_dir),
        "pngs": [str(p) for p in pngs],
    }
    metrics["plan"] = plan
    metrics["edits_summary"] = [
        {k: v for k, v in e.items() if k != "edits"} for e in build["edits"]
    ]
    metrics["timing"] = {"build_sec": round(dt_build, 1), "render_sec": round(dt_render, 1)}

    report_path = out_dir / "report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(metrics, f, ensure_ascii=False, indent=2)
    print(f"  report: {report_path}")

    return metrics


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else None

    print("=" * 80)
    print("Phase A3 Step 3 — 5 Benchmark v2 (paragraph-aware fill)")
    print("=" * 80)

    labels = load_catalog()
    skeletons = load_skeletons()
    store = ParagraphStore.load()
    print(f"loaded: {len(labels)} labels / {len(skeletons)} skeletons "
          f"/ {store.n_slides} slides w/ paragraph store")

    if target:
        scenarios = [target]
    else:
        scenarios = list(SCENARIO_CONTENT.keys())

    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    all_results = []
    for sid in scenarios:
        res = run_scenario(sid, labels, skeletons, store)
        all_results.append(res)

    scoreboard = {
        "phase": "A3-Step3-v2",
        "mode": "Mode A + paragraph-fill",
        "n_scenarios": len(all_results),
        "results": all_results,
        "summary": {
            "avg_role_hit_pct": round(
                sum(r["metrics"]["A_role_match"]["role_hit_pct"] for r in all_results)
                / len(all_results), 1,
            ),
            "avg_fill_pct": round(
                sum(r["metrics"]["B_slot_fill_v2"]["fill_pct"] for r in all_results)
                / len(all_results), 1,
            ),
            "avg_visual_resolution_pct": round(
                sum(r["metrics"]["B_slot_fill_v2"]["visual_resolution_pct"]
                    for r in all_results) / len(all_results), 1,
            ),
            "avg_overflow_rate_pct": round(
                sum(r["metrics"]["C_overflow"]["overflow_rate_pct"] for r in all_results)
                / len(all_results), 1,
            ),
            "avg_composite_quant": round(
                sum(r["metrics"]["composite_quant_score"] for r in all_results)
                / len(all_results), 1,
            ),
        },
    }

    scoreboard_path = OUTPUT_ROOT / "scoreboard.json"
    with open(scoreboard_path, "w", encoding="utf-8") as f:
        json.dump(scoreboard, f, ensure_ascii=False, indent=2)

    print()
    print("=" * 80)
    print("SCOREBOARD v2")
    print("=" * 80)
    for r in all_results:
        m = r["metrics"]
        print(
            f"  {r['scenario_id']:>30} | role={m['A_role_match']['role_hit_pct']:5.1f}% "
            f"fill={m['B_slot_fill_v2']['fill_pct']:5.1f}% "
            f"visual={m['B_slot_fill_v2']['visual_resolution_pct']:5.1f}% "
            f"overflow={m['C_overflow']['overflow_rate_pct']:5.1f}% "
            f"comp={m['composite_quant_score']:5.1f}"
        )
    print()
    print(f"AVG composite quant = {scoreboard['summary']['avg_composite_quant']}")
    print(f"saved: {scoreboard_path}")


if __name__ == "__main__":
    main()
