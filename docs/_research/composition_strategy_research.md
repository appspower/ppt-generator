# Composition Strategy Research: Whole-Slide vs Component-Level Reuse

**Date**: 2026-04-25
**Scope**: SOTA review for Mode A (whole-slide reuse) + N1-Lite (single-component-on-blank-slide) for the 1,251-slide PwC master deck.

---

## 1. SOTA Snapshot (2024-2026)

| Year | System | Approach | Key Finding |
|---|---|---|---|
| 2025 | **PPTAgent** (EMNLP 2025, icip-cas) | Two-stage edit-based: (1) analyze reference deck → extract functional types & schemas, (2) outline + iterative edit-API actions on a *selected reference slide*. PPTEval = Content/Design/Coherence. | Whole-slide reuse + targeted edits beats from-scratch generation across all 3 dimensions. arxiv 2501.03936; github icip-cas/PPTAgent. |
| 2025 | **AutoPresent** (CVPR 2025, Berkeley) | LLaMA-8B fine-tuned on SlidesBench (7k pairs) → emits Python that calls **SlidesLib** (high-level: `add_title`, `search_image`, …). | Matches GPT-4o on SlidesBench but execution rate of vanilla 8B/7B baselines is poor; refinement loops needed for color/size/overlap. From-scratch component composition is fragile. arxiv 2501.00912. |
| 2025 | **SlideCoder** (EMNLP 2025) | Layout-aware RAG-enhanced hierarchical decoder. | Visual design + layout shows the largest cross-model gap; closing it needs dedicated rendering pipelines, not bigger LLMs. aclanthology 2025.emnlp-main.458. |
| 2025 | **Talk to Your Slides** (arxiv 2505.11604) | LLM-driven structured edits on existing slides. | Editing > regeneration for fidelity. |
| 2025 | **Auto-Slides** (arxiv 2509.11062) | Multi-agent paper-to-deck with structural support. | Outline-first + role-tagged slides outperforms one-shot. |
| Practice | **MBB consultancies** | Internal "slide libraries" of past cases; consultants copy slides and re-fill. SCQA/Pyramid skeletons. | Reuse-then-edit is the human workflow PPTAgent explicitly emulates. theanalystacademy.com/powerpoint-storytelling. |

---

## 2. Why N1-Full (auto-composition with collision + harmonization) Failed in Literature

- **AutoPresent**: with the SlidesLib component API in hand, small open models *cannot reliably emit executable code*; even GPT-4o produces overlap, wrong colors, missing instructions until a refinement loop runs (paper §6 "auto-refinement" addresses *previously neglected* shape/bg-color/text instructions). Color & sizing are explicitly called out as weak.
- **SlideCoder / PresentBench (2026)**: Visual Design + Layout is the dimension with the largest model-to-model gap — closing it needs a dedicated visual pipeline, not a stronger backbone. Manus (53.7) still trails NotebookLM hugely on layout.
- **PPTAgent's own framing**: "Due to layout and modal complexity ... it is difficult for LLMs to directly determine which slides should be referenced" — and they only solve *selection*, not *composition from primitives*. They explicitly avoid composing because it is unsolved.
- Common failure modes across systems: element overlap, broken alignment, color drift from theme, font inconsistency, ignored instructions. All worsen as N (components on a slide) grows.

**Takeaway**: Nobody has solved multi-component auto-layout with style harmonization. PPTAgent + AutoPresent + SlideCoder all converge on "don't compose primitives — edit existing well-designed surfaces."

---

## 3. Why N1-Lite Is the Sweet Spot

- **One-insight-per-slide is canonical MBB practice** (theanalystacademy, slidescience, managementconsulted): each body slide carries one action title + one supporting visual. Many slides are effectively *one component on a frame*.
- **Empirical from our HJ scan**: of the 179 selected slides across 34 patterns, the dominant patterns (chevron-row, callout-cluster, single-card-grid, single-chart-with-takeaway, KPI-strip) are *single-component-on-frame*. Multi-pattern slides exist but are the minority and are the hardest to clone.
- **Composition risk collapses to placement-on-empty-canvas**: no collision resolution, no style negotiation between two libraries. Title bar + footer come from the master, the component is dropped into the body region. This is a solved problem (python-pptx group + anchor).
- **Style harmonization is free**: if every component is harvested *from the same master deck*, fonts/colors/sizes are already coherent. The N1-Full hard problem only appears when mixing libraries.
- N1-Lite ≈ AutoPresent's *best* config (single SlidesLib call) without the brittle code-gen step — because we keep the component as a native pptx group, not regenerated geometry.

---

## 4. Recommendations for Our 13 Narrative Roles

Heuristic: roles with **rich Mode-A inventory** and **distinctive whole-slide layouts** (cover, divider, agenda) → Mode A. Roles where the *information shape* is the message (a chevron, a 2x2, a KPI strip) → N1-Lite gives more flexibility than picking a fixed reference slide.

| Role | Inventory in 1251 deck | Recommended Mode | Why |
|---|---|---|---|
| opening | 8 (very thin) | **Mode A** + user capture | Cover/title needs whole-slide design (logo, gradient, hero). N1-Lite can't produce a hero. Need user to add 5-10 PwC opening captures. |
| agenda | moderate | **Mode A** | Agenda is itself a layout (numbered list with section headers). Whole-slide is cleaner. |
| divider | moderate | **Mode A** | Same as agenda — a divider IS a layout. |
| situation | 1 (gap) | **N1-Lite** (timeline / context block) | Inventory too thin for Mode A; situation is usually one timeline or one context callout — perfect single component. |
| complication | thin | **N1-Lite** (issue tree, callout cluster) | Often a single issue-tree or 3-card pain-points; component-friendly. |
| evidence | 452 (rich) | **Mode A primary, N1-Lite fallback** | Charts + tables come pre-laid-out; whole-slide reuse preserves data-ink ratio. |
| analysis | 1143 (richest) | **Mode A primary** | Most diverse archetype inventory; pick whole slide and refill. |
| recommendation | 640 (rich) | **Mode A** | Action-title + 3-card or matrix; well-designed reference slides exist. |
| roadmap | rich (chevrons, gantt) | **N1-Lite preferred** | Roadmap shape (chevron / gantt / phase-bar) IS the message. N1-Lite lets us pick the right shape independent of surrounding content; Mode A locks in extra elements. |
| benefit | 3 (gap) | **N1-Lite** (KPI strip, value-card grid) | Inventory too thin; benefit is canonically a KPI/value-card row. |
| risk | 1 (gap) | **N1-Lite** (risk matrix, callout) | Same — N1-Lite pulls the right component without forcing a whole-slide clone we don't have. |
| closing | 1 (gap) | **Mode A** + user capture | Like opening — needs whole-slide design. User capture required. |
| appendix | varies | **Mode A** | Backup data tables/charts are whole-slide by nature. |

**Pattern**: Mode A wins where (a) inventory is rich AND (b) the slide is more than its central component. N1-Lite wins where (a) inventory is sparse OR (b) the component itself carries the message.

---

## 5. Library Structure Recommendation

**Single .pptx per archetype family (not per component)**, plus a master metadata index.

```
ppt_builder/template/components/
  chevron_row.pptx          ← multiple chevron variants on separate slides (3/4/5/6 step)
  card_grid.pptx            ← 2x2, 3-col, 4-col card layouts
  callout_cluster.pptx      ← single callout, 2-callout, 3-callout
  kpi_strip.pptx            ← 3/4/5 KPI variants
  matrix_2x2.pptx           ← BCG/Gartner-style 2x2 frames
  timeline.pptx             ← horizontal/vertical/gantt
  issue_tree.pptx           ← MECE trees
  table_styled.pptx         ← consulting tables (ranked, RAG-status, …)
  chart_frame.pptx          ← chart-with-takeaway frames (bar/line/donut)
  ...
components_index.json       ← metadata
```

**Why family-grouped pptx, not 1-pptx-per-component**:
- python-pptx loads a `.pptx` once; reading 200 small files is slow and fragments theme/master.
- All chevrons share fonts/colors via the same master — opening one file gets coherent variants for free.
- Easier to add a new variant: drop a slide into the family file. No new file, no new metadata bootstrapping.

**Metadata to track per component (in `components_index.json`)**:
```
{
  "id": "chevron_row_5step_v1",
  "family": "chevron_row",
  "source_pptx": "components/chevron_row.pptx",
  "slide_index": 2,
  "shape_group_name": "G_CHEVRON_5",       # named group inside the slide
  "variants": {"steps": 5, "orientation": "horizontal"},
  "narrative_roles": ["roadmap", "recommendation"],
  "max_chars": {"step_label": 18, "step_detail": 60},
  "bbox_emu": [457200, 2057400, 11430000, 2400300],
  "anchors": ["body"],                     # body / full / left-half
  "harvested_from": "pwc_master_slide_0734"
}
```

The `shape_group_name` + named groups inside the source pptx are the key — N1-Lite copies the *group* (XML `<p:grpSp>`) into a blank target slide via python-pptx XML manipulation, preserving fonts/colors/sizes natively. Same trick PPTAgent uses for its edit APIs but applied at component granularity.

**Bootstrapping**: harvest components by selecting groups inside Mode-A slides we already have (the 1251 deck) — every component is provably PwC-coherent. Start with ~10 families × 3-5 variants = ~40 components covers the 6 "N1-Lite preferred" roles above.

---

## Sources

- [PPTAgent (EMNLP 2025) – paper](https://aclanthology.org/2025.emnlp-main.728.pdf)
- [PPTAgent – arxiv 2501.03936 v3](https://arxiv.org/html/2501.03936v3)
- [PPTAgent – GitHub icip-cas/PPTAgent](https://github.com/icip-cas/PPTAgent)
- [AutoPresent (CVPR 2025) – paper](https://arxiv.org/html/2501.00912v1)
- [AutoPresent – GitHub para-lost/AutoPresent](https://github.com/para-lost/AutoPresent)
- [SlideCoder (EMNLP 2025)](https://aclanthology.org/2025.emnlp-main.458.pdf)
- [Talk to Your Slides (arxiv 2505.11604)](https://arxiv.org/html/2505.11604v1)
- [Auto-Slides (arxiv 2509.11062)](https://arxiv.org/html/2509.11062)
- [PresentBench (2026)](https://arxiv.org/html/2603.07244v1)
- [SlideTailor (arxiv 2512.20292)](https://arxiv.org/html/2512.20292v1)
- [MBB SCQA storytelling – Analyst Academy](https://www.theanalystacademy.com/powerpoint-storytelling/)
- [One-insight-per-slide – Slide Science](https://slidescience.co/strategy-presentations/)
- [Consulting slide libraries – Slideworks](https://slideworks.io/resources/how-mckinsey-consultants-make-presentations)
