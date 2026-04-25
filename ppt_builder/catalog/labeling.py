"""Track 1 Stage 2c — 클러스터 VLM 라벨링 + narrative_role anchor propagation.

전략
----
1. 각 cluster 대표 슬라이드 PNG 를 Claude (VLM)에 보냄:
   → (category_name, structure_type, intent, narrative_role_candidates) 반환
2. 사용자가 수동 annotate 한 40 anchor slide_index 를 입력으로:
   같은 cluster의 모든 slide 에 narrative_role propagate
3. 앰비규어스 클러스터 (VLM이 narrative_role 2개 이상 추천) 는 별도 표시

출력
----
  cluster_labeled.json : cluster_id → {
      category_name, structure_type, intent, narrative_role (추론),
      anchor_source (manual/propagated/vlm), confidence
  }

주의: **이 스크립트는 Claude API 호출 없이 로컬에서 프롬프트+결과 저장까지만**.
실제 API 호출은 세션 내 대화에서 Claude Code가 직접 이미지 보고 태깅하는
방식으로 처리. (엔터프라이즈 보안 + 비용 고려)
"""
from __future__ import annotations

import json
from pathlib import Path

from .schemas import NarrativeRole


VLM_PROMPT_TEMPLATE = """You are analyzing one slide from a consulting-grade PowerPoint template library.
All text is masked with `~~` placeholders — focus on the **visual/structural layout** only.

Task: Classify this slide's archetype into:
1. **category_name** (short, e.g. "3-column cards", "swimlane flow", "matrix 2x2", "timeline roadmap")
2. **structure_type** (one of: grid, flow, hierarchy, cards, chart, text_heavy, cover, divider, mixed)
3. **intent** (list from: define, compare, sequence, decompose, emphasize, showcase)
4. **narrative_role_candidates** (list from: opening, agenda, situation, complication, evidence,
   analysis, recommendation, roadmap, benefit, risk, closing, divider, appendix)
   — pick the top 1-2 most plausible roles this layout is designed for.

Return JSON only:
{
  "category_name": "...",
  "structure_type": "...",
  "intent": [...],
  "narrative_role_candidates": [...],
  "confidence": 0.0-1.0
}
"""


def build_prompt_for_cluster(cluster_id: int, representative_idx: int, size: int) -> str:
    """cluster 대표 슬라이드 라벨링용 프롬프트."""
    return (
        f"# Cluster {cluster_id} (size={size}, representative slide_index={representative_idx})\n\n"
        + VLM_PROMPT_TEMPLATE
    )


def propagate_from_anchors(
    anchor_annotations: dict[int, NarrativeRole],
    labels: list[int],
) -> dict[int, NarrativeRole]:
    """anchor slide_index → cluster majority vote → 해당 cluster 모든 멤버에 propagation.

    Args:
        anchor_annotations: { slide_index: NarrativeRole } (수동 annotate)
        labels: cluster 라벨 배열 (Stage 2b 결과)

    Returns:
        { slide_index: NarrativeRole } (1251 entries, UNKNOWN 가능)
    """
    from collections import Counter, defaultdict

    # cluster → anchor roles
    cluster_to_roles: dict[int, list[NarrativeRole]] = defaultdict(list)
    for idx, role in anchor_annotations.items():
        if 0 <= idx < len(labels):
            cluster_to_roles[labels[idx]].append(role)

    # cluster → dominant role
    cluster_role: dict[int, NarrativeRole] = {}
    for cid, roles in cluster_to_roles.items():
        dominant = Counter(roles).most_common(1)[0][0]
        cluster_role[cid] = dominant

    # propagate
    result: dict[int, NarrativeRole] = {}
    for i, lb in enumerate(labels):
        result[i] = cluster_role.get(int(lb), NarrativeRole.UNKNOWN)

    return result


def export_prompts_pack(
    clusters_path: Path,
    png_dir: Path,
    output_dir: Path,
) -> None:
    """클러스터별 대표 PNG 경로 + 프롬프트 텍스트 묶음 내보내기.

    Claude Code 세션에서 이미지를 Read 하고 응답받은 JSON을
    cluster_labeled.json 으로 수집하는 용도.
    """
    with open(clusters_path, "r", encoding="utf-8") as f:
        clusters = json.load(f)

    output_dir.mkdir(parents=True, exist_ok=True)
    pack = []
    for c in clusters:
        if c["cluster_id"] == -1:
            continue
        rep_idx = c["representative_slide_index"]
        png = png_dir / f"slide_{rep_idx:04d}.png"
        pack.append({
            "cluster_id": c["cluster_id"],
            "size": c["size"],
            "representative_slide_index": rep_idx,
            "png_path": str(png),
            "png_exists": png.exists(),
            "layout_name": c.get("representative_layout", ""),
            "structure_sig": c.get("representative_structure_sig", ""),
            "prompt": build_prompt_for_cluster(c["cluster_id"], rep_idx, c["size"]),
        })
    with open(output_dir / "labeling_pack.json", "w", encoding="utf-8") as f:
        json.dump(pack, f, ensure_ascii=False, indent=2)
    print(f"[2c] labeling pack: {len(pack)} clusters -> {output_dir}/labeling_pack.json")


if __name__ == "__main__":
    ROOT = Path(__file__).resolve().parent.parent.parent
    clusters = ROOT / "output" / "catalog" / "clusters_v3.json"
    png_dir = ROOT / "output" / "catalog" / "all_pngs"
    out = ROOT / "output" / "catalog"
    if clusters.exists():
        export_prompts_pack(clusters, png_dir, out)
    else:
        print(f"[skip] clusters_v3.json not found (Stage 2b pending)")
