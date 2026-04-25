"""Layer 3 — 사용자 검수 HTML 페이지 생성.

입력: layer3_user_queue.json
출력: review.html (브라우저로 열어서 검수 + final_labels.json 다운로드)

기능
----
- 슬라이드별 PNG 썸네일 + 자동 라벨 + 수정 폼
- L1 macro radio / L2 archetype checkbox / L3 narrative_role checkbox
- 모든 변경 후 "Save JSON" 버튼 → final_labels.json 다운로드
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

L1_MACROS = ["chart", "table", "card", "diagram", "cover", "unknown"]
L2_ARCHETYPES = [
    "table_native", "dense_grid", "matrix_2x2", "matrix_3x3", "matrix_NxN",
    "cards_2col", "cards_3col", "cards_4col", "cards_5plus", "vertical_list",
    "orgchart", "hub_spoke", "flowchart", "roadmap", "timeline_h", "gantt",
    "swimlane", "funnel", "venn", "chart_native", "cover_divider",
    "single_block", "left_title_right_body", "unknown",
]
L3_ROLES = [
    "opening", "agenda", "situation", "complication", "evidence", "analysis",
    "recommendation", "roadmap", "benefit", "risk", "closing", "divider",
    "appendix", "unknown",
]


def build_html(queue_path: Path, png_dir: Path, output_html: Path) -> None:
    with open(queue_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    queue = data.get("queue", [])
    summary = data.get("summary", {})

    # PNG는 상대 경로로 (HTML 옆에 있다고 가정)
    rel_png_dir = png_dir.relative_to(output_html.parent.parent.parent) if png_dir.is_relative_to(output_html.parent.parent.parent) else png_dir

    cards_html: list[str] = []
    for i, q in enumerate(queue):
        idx = q["slide_index"]
        png_rel = f"all_pngs/slide_{idx:04d}.png"
        macro_radios = "".join([
            f'<label><input type="radio" name="macro_{idx}" value="{m}" '
            f'{"checked" if m == q.get("macro") else ""}> {m}</label> '
            for m in L1_MACROS
        ])
        arch_checks = "".join([
            f'<label><input type="checkbox" name="arch_{idx}" value="{a}" '
            f'{"checked" if a in q.get("archetype", []) else ""}> {a}</label> '
            for a in L2_ARCHETYPES
        ])
        role_checks = "".join([
            f'<label><input type="checkbox" name="role_{idx}" value="{r}" '
            f'{"checked" if r in q.get("narrative_role", []) else ""}> {r}</label> '
            for r in L3_ROLES
        ])
        notes = q.get("notes", "") or ""
        confidence = q.get("confidence", 0)

        cards_html.append(f"""
<div class="card" data-slide="{idx}">
  <div class="left">
    <img src="{png_rel}" loading="lazy" />
    <div class="meta">
      <strong>Slide #{idx}</strong>
      <span>conf {confidence:.2f}</span>
      <span>auto: {q.get('macro', '?')}</span>
    </div>
  </div>
  <div class="right">
    <div class="row">
      <h4>L1 Macro (정확 1개)</h4>
      {macro_radios}
    </div>
    <div class="row">
      <h4>L2 Archetype (1~3 multi-label)</h4>
      <div class="chk">{arch_checks}</div>
    </div>
    <div class="row">
      <h4>L3 Narrative Role (1~2 multi-label)</h4>
      <div class="chk">{role_checks}</div>
    </div>
    <div class="row">
      <textarea name="notes_{idx}" placeholder="notes">{notes}</textarea>
    </div>
  </div>
</div>
""")

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<title>Layer 3 — 사용자 검수 (PPT Generator)</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", sans-serif;
       margin: 20px; background: #f5f5f5; color: #222; }}
header {{ background: #fff; padding: 14px 18px; border-radius: 8px; margin-bottom: 12px;
         box-shadow: 0 1px 3px rgba(0,0,0,.06); display: flex; justify-content: space-between; align-items: center; }}
header h1 {{ font-size: 18px; margin: 0; }}
header .summary {{ font-size: 13px; color: #666; }}
header button {{ padding: 8px 14px; background: #ff6a00; color: white; border: none; border-radius: 5px;
                cursor: pointer; font-weight: 600; }}
header button:hover {{ background: #e85d00; }}
.card {{ background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.06);
         margin-bottom: 14px; padding: 14px; display: grid; grid-template-columns: 320px 1fr; gap: 16px; }}
.card .left img {{ width: 300px; height: auto; border: 1px solid #ddd; }}
.card .meta {{ font-size: 12px; color: #555; margin-top: 6px; display: flex; gap: 8px; flex-wrap: wrap; }}
.card .meta strong {{ color: #222; }}
.card .right h4 {{ font-size: 13px; margin: 4px 0; color: #333; }}
.card .row {{ margin-bottom: 8px; }}
.card label {{ font-size: 12px; margin-right: 10px; cursor: pointer; }}
.chk {{ display: flex; flex-wrap: wrap; gap: 4px; }}
.chk label {{ background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }}
.chk label input:checked + ::before {{ background: #ff6a00; }}
textarea {{ width: 100%; min-height: 36px; font-size: 12px; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>PPT Generator — Layer 3 사용자 검수</h1>
    <div class="summary">
      Total: {summary.get('total', '?')} 장 · L2 검수: {summary.get('layer2_reviewed', '?')} · L3 검수 큐: <strong>{len(queue)}</strong>
    </div>
  </div>
  <button onclick="saveLabels()">📥 Save final_labels.json</button>
</header>
<div id="cards">
{''.join(cards_html)}
</div>
<script>
function collect() {{
  const cards = document.querySelectorAll('.card');
  const labels = [];
  cards.forEach(card => {{
    const idx = parseInt(card.dataset.slide);
    const macroEl = card.querySelector(`input[name="macro_${{idx}}"]:checked`);
    const macro = macroEl ? macroEl.value : 'unknown';
    const arch = Array.from(card.querySelectorAll(`input[name="arch_${{idx}}"]:checked`)).map(e => e.value);
    const role = Array.from(card.querySelectorAll(`input[name="role_${{idx}}"]:checked`)).map(e => e.value);
    const notes = card.querySelector(`textarea[name="notes_${{idx}}"]`).value;
    labels.push({{
      slide_index: idx,
      macro: macro,
      archetype: arch,
      narrative_role: role,
      notes: notes,
      reviewed_by_user: true,
    }});
  }});
  return labels;
}}
function saveLabels() {{
  const data = collect();
  const blob = new Blob([JSON.stringify(data, null, 2)], {{ type: 'application/json' }});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'final_labels_user.json';
  a.click();
  URL.revokeObjectURL(url);
  alert(`${{data.length}} 슬라이드 라벨이 final_labels_user.json 으로 저장되었습니다.`);
}}
</script>
</body>
</html>
"""

    output_html.parent.mkdir(parents=True, exist_ok=True)
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[L3] HTML saved -> {output_html}")
    print(f"[L3] queue: {len(queue)} slides")


if __name__ == "__main__":
    queue = ROOT / "output" / "catalog" / "layer3_user_queue.json"
    png_dir = ROOT / "output" / "catalog" / "all_pngs"
    out = ROOT / "output" / "catalog" / "review.html"
    if queue.exists():
        build_html(queue, png_dir, out)
    else:
        print(f"[skip] {queue} not found yet")
