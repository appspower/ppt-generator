"""Phase A3 Step 2 — paragraph 단위 결정론적 role 라벨링.

extract_paragraphs.py의 출력(paragraphs.json)을 입력으로 받아
각 paragraph에 다음을 부여:
  - role: title / subtitle / kicker / table_header / table_cell /
          chevron_label / phase_label / card_header / card_body /
          callout_text / footer / page_number / date_text /
          axis_label / data_label / kpi_value / decorative
  - role_source: "placeholder" | "table_geom" | "shape_geom" |
                 "font_size" | "position" | "default"
  - role_confidence: 0.0~1.0
  - position_in_group: int or None
  - group_signature: bbox 정렬 기반 sibling 그룹 식별자

CLI: python scripts/label_paragraphs.py
출력: output/catalog/paragraph_labels.json
"""
from __future__ import annotations

import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = ROOT / "output" / "catalog" / "paragraphs.json"
OUTPUT_PATH = ROOT / "output" / "catalog" / "paragraph_labels.json"


# ----------------------------------------------------------------------------
# Role 결정 규칙 (우선순위 순서대로)
# ----------------------------------------------------------------------------

def _role_from_placeholder(p: dict) -> tuple[str | None, float, str]:
    """placeholder_type 기반."""
    pt = p.get("placeholder_type")
    if not pt:
        return None, 0, ""
    pt_upper = pt.upper()
    mapping = {
        "TITLE": ("title", 0.95),
        "CENTER_TITLE": ("title", 0.95),
        "VERTICAL_TITLE": ("title", 0.92),
        "SUBTITLE": ("subtitle", 0.92),
        "BODY": ("body", 0.85),
        "OBJECT": ("body", 0.7),
        "PICTURE": ("decorative", 0.6),
        "FOOTER": ("footer", 0.95),
        "HEADER": ("header", 0.95),
        "DATE": ("date_text", 0.95),
        "SLIDE_NUMBER": ("page_number", 0.98),
        "PAGE_NUMBER": ("page_number", 0.98),
    }
    for key, (role, conf) in mapping.items():
        if key in pt_upper:
            return role, conf, "placeholder"
    return None, 0, ""


def _role_from_table(p: dict) -> tuple[str | None, float, str]:
    """TABLE 셀 위치 기반."""
    if p.get("shape_kind") != "TABLE":
        return None, 0, ""
    row = p.get("table_row")
    n_rows = p.get("table_n_rows", 1)
    if row is None:
        return None, 0, ""
    if row == 0:
        return "table_header", 0.95, "table_geom"
    # 마지막 행이 합계/요약일 가능성
    if n_rows >= 3 and row == n_rows - 1:
        return "table_cell_summary", 0.7, "table_geom"
    return "table_cell", 0.9, "table_geom"


def _role_from_shape_geom(p: dict) -> tuple[str | None, float, str]:
    """AUTOSHAPE 종류 기반 (chevron, callout, etc.)."""
    sk = p.get("shape_kind", "")
    if not sk.startswith("AUTOSHAPE:"):
        return None, 0, ""
    ast = sk.split(":", 1)[1]

    chevron_like = {
        "CHEVRON", "PENTAGON", "RIGHT_ARROW", "RIGHT_ARROW_CALLOUT",
        "STRIPED_RIGHT_ARROW", "NOTCHED_RIGHT_ARROW", "CIRCULAR_ARROW",
        "BENT_ARROW", "BENT_UP_ARROW",
    }
    if ast in chevron_like:
        return "chevron_label", 0.85, "shape_geom"

    callout_like = {
        "ROUNDED_RECTANGULAR_CALLOUT", "RECTANGULAR_CALLOUT",
        "OVAL_CALLOUT", "CLOUD_CALLOUT", "LINE_CALLOUT_1",
        "LINE_CALLOUT_2", "LINE_CALLOUT_3", "LINE_CALLOUT_4",
        "LINE_CALLOUT_1_NO_BORDER", "LINE_CALLOUT_2_NO_BORDER",
    }
    if ast in callout_like:
        return "callout_text", 0.85, "shape_geom"

    star_burst = {
        "STAR_4_POINT", "STAR_5_POINT", "STAR_6_POINT", "STAR_7_POINT",
        "STAR_8_POINT", "STAR_10_POINT", "STAR_12_POINT", "STAR_16_POINT",
        "STAR_24_POINT", "STAR_32_POINT", "EXPLOSION_1", "EXPLOSION_2",
    }
    if ast in star_burst:
        return "kpi_value", 0.75, "shape_geom"

    return None, 0, ""


def _role_from_position(p: dict, slide_stats: dict) -> tuple[str | None, float, str]:
    """슬라이드 내 bbox 위치 + font 크기 기반."""
    bbox = p["bbox"]
    top = bbox["top_pct"]
    left = bbox["left_pct"]
    width = bbox["width_pct"]
    height = bbox["height_pct"]
    font = p["font"]
    fs = font.get("font_size_pt") or 0
    text_len = p["text_len"]

    # 페이지 번호: 매우 작은 + 우하단 + 숫자
    if (text_len <= 4 and p["text"].strip().replace(",", "").replace(".", "").isdigit()):
        if top > 0.93 or (top > 0.85 and left > 0.85):
            return "page_number", 0.85, "position"

    # footer 후보 (bottom 5% + 작은 텍스트)
    if top + height > 0.97 and height < 0.05 and text_len < 50:
        return "footer", 0.7, "position"

    # 큰 폰트 + top 15% = title 후보
    max_fs = slide_stats["max_font_size"]
    if fs and max_fs and fs >= max_fs * 0.95 and top < 0.15:
        return "title", 0.78, "position"

    # 두번째 큰 폰트 + top 25% = subtitle/kicker
    second_fs = slide_stats["second_font_size"]
    if fs and second_fs and second_fs >= 12 and abs(fs - second_fs) < 0.5 and top < 0.25:
        return "subtitle", 0.65, "position"

    return None, 0, ""


# ----------------------------------------------------------------------------
# 슬라이드 통계 (font 크기 분포 등)
# ----------------------------------------------------------------------------

def compute_slide_stats(slide_paragraphs: list[dict]) -> dict:
    """슬라이드 내 paragraph들로부터 font size 분포 등 통계."""
    sizes = []
    for p in slide_paragraphs:
        fs = p["font"].get("font_size_pt")
        if fs:
            sizes.append(fs)
    sizes_sorted = sorted(set(sizes), reverse=True)
    return {
        "max_font_size": sizes_sorted[0] if sizes_sorted else None,
        "second_font_size": sizes_sorted[1] if len(sizes_sorted) > 1 else None,
        "n_distinct_sizes": len(sizes_sorted),
        "all_sizes": sizes_sorted,
    }


# ----------------------------------------------------------------------------
# Group signature (sibling 그룹 식별)
# ----------------------------------------------------------------------------

def assign_group_signatures(slide_paragraphs: list[dict]) -> None:
    """같은 group_path + 비슷한 폭/높이로 묶이는 sibling을 그룹화.
    position_in_group을 부여 (left to right, then top to bottom).
    """
    by_group: dict[str | None, list[int]] = defaultdict(list)
    for i, p in enumerate(slide_paragraphs):
        if p.get("paragraph_id", 0) != 0:
            continue  # 오직 paragraph 0만 (shape의 첫 줄)
        if p["shape_kind"] == "TABLE":
            continue
        gp = p.get("group_path")
        by_group[gp].append(i)

    for gp, indices in by_group.items():
        if len(indices) < 2:
            for i in indices:
                slide_paragraphs[i]["group_signature"] = None
                slide_paragraphs[i]["position_in_group"] = None
                slide_paragraphs[i]["group_size"] = 1
            continue

        # 비슷한 width/height 묶기
        records = [
            (i, slide_paragraphs[i]["bbox"]["width_pct"],
             slide_paragraphs[i]["bbox"]["height_pct"],
             slide_paragraphs[i]["bbox"]["left_pct"],
             slide_paragraphs[i]["bbox"]["top_pct"])
            for i in indices
        ]
        clusters: list[list[tuple]] = []
        for rec in records:
            i, w, h, l, t = rec
            placed = False
            for cl in clusters:
                _, w0, h0, _, _ = cl[0]
                if abs(w - w0) < 0.02 and abs(h - h0) < 0.02:
                    cl.append(rec)
                    placed = True
                    break
            if not placed:
                clusters.append([rec])

        for cl in clusters:
            if len(cl) < 2:
                idx = cl[0][0]
                slide_paragraphs[idx]["group_signature"] = None
                slide_paragraphs[idx]["position_in_group"] = None
                slide_paragraphs[idx]["group_size"] = 1
                continue
            # 좌→우, 위→아래 정렬
            cl_sorted = sorted(cl, key=lambda r: (r[4], r[3]))  # top, left
            sig_parts = (
                f"g{gp or 'none'}_w{cl_sorted[0][1]:.2f}_"
                f"h{cl_sorted[0][2]:.2f}_n{len(cl_sorted)}"
            )
            for pos, (idx, *_rest) in enumerate(cl_sorted):
                slide_paragraphs[idx]["group_signature"] = sig_parts
                slide_paragraphs[idx]["position_in_group"] = pos
                slide_paragraphs[idx]["group_size"] = len(cl_sorted)


def propagate_group_role(slide_paragraphs: list[dict]) -> None:
    """그룹의 첫 번째가 title/header이거나 chevron이면 같은 그룹은 같은 role.
    또한 group_size >= 2이고 동일 AUTOSHAPE인 sibling을 card_header/card_body로 승격.
    """
    by_sig: dict[str, list[int]] = defaultdict(list)
    for i, p in enumerate(slide_paragraphs):
        sig = p.get("group_signature")
        if sig:
            by_sig[sig].append(i)

    # 보조 인덱스: 모든 paragraph에 대해 (flat_idx, paragraph_id) 매핑
    by_flat: dict[int, list[int]] = defaultdict(list)
    for i, p in enumerate(slide_paragraphs):
        by_flat[p["flat_idx"]].append(i)

    for sig, indices in by_sig.items():
        if len(indices) < 2:
            continue
        # 그룹 시드 paragraph들의 role 빈도
        seed_roles = [slide_paragraphs[i].get("role") for i in indices]
        # chevron이 다수면 그룹 전체에 chevron_label 전파
        if seed_roles.count("chevron_label") >= 2:
            for i in indices:
                if slide_paragraphs[i].get("role") in ("decorative", None):
                    slide_paragraphs[i]["role"] = "chevron_label"
                    slide_paragraphs[i]["role_source"] = "group_propagate"
                    slide_paragraphs[i]["role_confidence"] = 0.7
            continue
        # callout 다수면 그룹 전체에 callout_text 전파
        if seed_roles.count("callout_text") >= 2:
            for i in indices:
                if slide_paragraphs[i].get("role") in ("decorative", None):
                    slide_paragraphs[i]["role"] = "callout_text"
                    slide_paragraphs[i]["role_source"] = "group_propagate"
                    slide_paragraphs[i]["role_confidence"] = 0.7
            continue
        # decorative가 대부분이고 group_size 2~12 AUTOSHAPE이면 card 후보
        # (group_size > 12는 decorative 그리드/장식일 가능성)
        n_decorative = seed_roles.count("decorative")
        group_size = len(indices)
        if n_decorative >= 2 and 2 <= group_size <= 12:
            kinds = [slide_paragraphs[i]["shape_kind"] for i in indices]
            if all(k.startswith("AUTOSHAPE:") or k == "SHAPE" for k in kinds):
                for i in indices:
                    if slide_paragraphs[i].get("role") == "decorative":
                        slide_paragraphs[i]["role"] = "card_header"
                        slide_paragraphs[i]["role_source"] = "group_card"
                        slide_paragraphs[i]["role_confidence"] = 0.7
                    # 같은 shape의 paragraph_id >= 1은 card_body
                    flat = slide_paragraphs[i]["flat_idx"]
                    for j in by_flat[flat]:
                        if j == i:
                            continue
                        if slide_paragraphs[j]["paragraph_id"] >= 1 and \
                           slide_paragraphs[j].get("role") in ("decorative", None):
                            slide_paragraphs[j]["role"] = "card_body"
                            slide_paragraphs[j]["role_source"] = "group_card"
                            slide_paragraphs[j]["role_confidence"] = 0.65


# ----------------------------------------------------------------------------
# 메인 라벨러
# ----------------------------------------------------------------------------

def label_slide(slide_paragraphs: list[dict]) -> None:
    """슬라이드 내 paragraph들에 role 부여 (in-place)."""
    stats = compute_slide_stats(slide_paragraphs)

    for p in slide_paragraphs:
        # 1. placeholder
        role, conf, src = _role_from_placeholder(p)
        if role:
            p["role"] = role
            p["role_confidence"] = conf
            p["role_source"] = src
            continue
        # 2. TABLE
        role, conf, src = _role_from_table(p)
        if role:
            p["role"] = role
            p["role_confidence"] = conf
            p["role_source"] = src
            continue
        # 3. shape geometry
        role, conf, src = _role_from_shape_geom(p)
        if role:
            p["role"] = role
            p["role_confidence"] = conf
            p["role_source"] = src
            continue
        # 4. position + font
        role, conf, src = _role_from_position(p, stats)
        if role:
            p["role"] = role
            p["role_confidence"] = conf
            p["role_source"] = src
            continue
        # default
        p["role"] = "decorative"
        p["role_confidence"] = 0.3
        p["role_source"] = "default"

    # group signature + propagate
    assign_group_signatures(slide_paragraphs)
    propagate_group_role(slide_paragraphs)


def main():
    print(f"[load] {INPUT_PATH}", flush=True)
    with open(INPUT_PATH, encoding="utf-8") as f:
        data = json.load(f)
    paragraphs = data["paragraphs"]
    print(f"  {len(paragraphs)} paragraphs / {data['summary']['n_slides']} slides")

    # slide_index로 그룹화
    by_slide: dict[int, list[dict]] = defaultdict(list)
    for p in paragraphs:
        by_slide[p["slide_index"]].append(p)

    print("[label] applying deterministic rules...", flush=True)
    for s_i, slide_paras in by_slide.items():
        label_slide(slide_paras)

    # 통계
    role_counter = Counter()
    source_counter = Counter()
    for p in paragraphs:
        role_counter[p.get("role", "?")] += 1
        source_counter[p.get("role_source", "?")] += 1

    print()
    print("=" * 60)
    print("ROLE DISTRIBUTION")
    print("=" * 60)
    for role, n in role_counter.most_common():
        pct = n / len(paragraphs) * 100
        print(f"  {role:>20}: {n:6d} ({pct:5.1f}%)")
    print()
    print("=" * 60)
    print("SOURCE DISTRIBUTION")
    print("=" * 60)
    for src, n in source_counter.most_common():
        pct = n / len(paragraphs) * 100
        print(f"  {src:>20}: {n:6d} ({pct:5.1f}%)")

    # 핵심 priority 슬롯
    high_priority = ["title", "table_header", "chevron_label",
                     "card_header", "callout_text"]
    n_hp = sum(role_counter.get(r, 0) for r in high_priority)
    print(f"\n[high priority] {n_hp} paragraphs ({n_hp/len(paragraphs)*100:.1f}%)")

    output = {
        "summary": {
            **data["summary"],
            "role_distribution": dict(role_counter.most_common()),
            "source_distribution": dict(source_counter.most_common()),
            "high_priority_count": n_hp,
        },
        "paragraphs": paragraphs,
    }

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[saved] {OUTPUT_PATH}")
    print(f"  size: {OUTPUT_PATH.stat().st_size / 1024 / 1024:.1f} MB")


if __name__ == "__main__":
    main()
