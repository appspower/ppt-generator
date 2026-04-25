# Component Extraction Research — python-pptx Sub-slide Reuse

Date: 2026-04-25
Scope: Extracting GroupShape / shape clusters from a 1,251-slide master deck and re-inserting them into blank slides, while preserving Korean fonts, theme colors, images, charts and tables.

---

## 1. Top 3 Recommended Techniques

### T1. lxml `copy.deepcopy` of `<p:grpSp>` + targeted relationship rewrite (RECOMMENDED)
The de-facto pattern across the python-pptx community. Steps:

1. Locate the source group element via `shape.element` (returns `CT_GroupShape`, the same XML type as `p:spTree` itself — grouping is recursive).
2. `newel = copy.deepcopy(shape.element)`.
3. Insert into destination via `dest_slide.shapes._spTree.insert_element_before(newel, "p:extLst")`.
4. **Walk the cloned tree** to (a) generate fresh `a16:creationId` GUIDs on every shape (duplicates corrupt PPTX on second insert — issue #961), and (b) rewrite every `r:embed` / `r:link` rId by calling `dest_slide.part.relate_to(source_part, RELATIONSHIP_TYPE)` and substituting the new rId in the XML.

Code-level pointer: this is exactly the pattern in the **Dasc3er gist** (GroupShape branch + recursive `copy_shapes`), plus the rId rewrite step which the gist omits.

### T2. Hybrid: deepcopy XML for vector shapes + `add_picture(blob)` for images
For Picture descendants inside the group, extract `shape.image.blob`, call `dest.shapes.add_picture(BytesIO(blob), ...)`, then carry over `crop_left/right/top/bottom` and `name`. Charts get treated similarly via `chart.part.blob` re-embedding. This is robust but loses any group-level transforms; you must replay `grpSpPr` (offset/extent/chOffset/chExtent) on the new container. See the Dasc3er gist's GroupShape branch.

### T3. Borrow PPTAgent's `shape_filter` HTML-schema layer for *indexing*, keep deepcopy for *insertion*
PPTAgent (EMNLP 2025) does **not** itself extract sub-slide components — its schema and 5 edit APIs (`del_span`, `del_image`, `clone_paragraph`, `replace_span`, `replace_image`) are paragraph/image-level only, atop whole-slide templates. But its `shape_filter(return_father=True)` and `_prs_to_html()` give a clean way to *describe* what's inside a group (category / description / content per element). Use it to label your extracted groups; do the actual XML extraction yourself.

---

## 2. Major Pitfalls and How to Avoid Them

| Pitfall | Symptom | Fix |
|---|---|---|
| **Duplicate `a16:creationId`** | "PowerPoint found unreadable content" on reopen, especially after 2nd reuse of same group | After deepcopy, walk `qn('a16:creationId')` attrs and assign fresh `{GUID}` to each |
| **Broken `r:embed` rId** | Image renders blank, chart fails to load | rId is *slide-part-scoped*; must call `dest_slide.part.relate_to(...)` for each embedded part |
| **Theme color drift** | `<a:schemeClr val="accent1">` resolves differently across master decks | When source/dest share the master, schemeClr is preserved automatically; cross-master cloning requires resolving to RGB or copying the master |
| **Korean East-Asian fonts dropped** | Hangul falls back to Calibri | `<a:ea typeface="맑은 고딕">` in run properties survives deepcopy unchanged. Already validated in this project's Phase A1 POC slide 1175 |
| **Placeholder inheritance lost** | Title font/size mysteriously changes | Don't paste placeholders into a non-matching layout; convert to ordinary `<p:sp>` by stripping `<p:nvSpPr><p:nvPr><p:ph/>` |
| **Chart embedded XLSX** | Chart shows but data is "Series 1, 2..." | Charts carry an embedded XLSX part; you must `relate_to` the XLSX too, not just the chart XML |
| **bbox collision on compose** | Two extracted groups overlap on blank canvas | See §4 below |

---

## 3. SOTA References (2024–2026)

- **PPTAgent** (Zheng et al., EMNLP 2025) — `https://aclanthology.org/2025.emnlp-main.728/`, code `https://github.com/icip-cas/PPTAgent`. Whole-slide reuse + 5 sub-slide edit APIs. **No component library.**
- **SlideCoder** (EMNLP 2025, `https://aclanthology.org/2025.emnlp-main.458.pdf`) — Layout-aware RAG, hierarchical, but generates code rather than extracting from reference decks.
- **scanny/python-pptx issues #132, #232, #533, #961, #1036** — community-tracked state of slide/shape duplication; consistent message: *no built-in support, deepcopy + rId rewrite is the answer*.
- **Dasc3er gist** (`https://gist.github.com/Dasc3er/2af5069afb728c39d54434cb28a1dbb8`) — most complete public helper for GroupShape recursive copy.
- **Aspose.Slides for Python** — commercial library that does this natively, but proprietary; not aligned with `01_VISION.md`.

---

## 4. Concrete Recommendation for Our Project

We have 1,251 master slides, Mode A (whole-slide reuse) already working, and Korean content. Recommendation:

1. **Build a `ppt_builder/template/component_ops.py` next to `edit_ops.py`** with two functions: `extract_group(source_slide, group_idx) -> ComponentBundle` and `insert_component(dest_slide, bundle, left, top)`. Bundle stores: deepcopy'd `grpSp` XML, list of `(rel_type, source_part, source_rId)` tuples, and metadata (see schema below). This keeps the existing `edit_ops.py` untouched and additive.

2. **Use Technique T1** (XML deepcopy + rId rewrite + fresh creationId). T2's `add_picture(blob)` round-trip is unnecessary because we control source masters and can re-`relate_to` parts directly. PPTAgent's `replace_image` already gives us the rId-rewrite primitive — generalize it.

3. **Component metadata schema** (store as JSON sidecar alongside `final_labels.json`):
   - `component_id`, `source_slide_idx`, `source_group_path` (e.g. `[3, 1]` = 4th top-level shape, 2nd nested child)
   - `bbox_emu`: `{left, top, width, height}`
   - `role`: one of `chevron_row | card_grid | callout_cluster | table | chart | kpi_strip | …` (extend the 7 skeleton labels)
   - `slot_count` and per-slot `{ text_path, max_chars, font_size_pt, has_korean }`
   - `dependencies`: `images[]`, `chart_part`, `theme_required`
   - `bbox_class`: bucket into 12-col grid (left-half, right-half, top-third…) for layout collision avoidance.

4. **Layout collision (N1-Lite compose)**: enforce a 12-column × 6-row grid; each component declares its grid footprint; reject placements whose footprints intersect; auto-shift Korean text components down 4pt when adjacent to dense text (Korean line-height heuristic from Phase A2 max_chars work).

5. **Skip in v1**: cross-master theme rewriting, animations/timing, smart-art. These are <5% of A-grade slides per the Phase A1 inspection report and add disproportionate complexity.

6. **Validation gate**: every extracted component must roundtrip through `evaluate.py` + COM PNG visual check (Visual Check Rule from MEMORY) on a clean blank slide before entering the library.

---

## Sources

- [python-pptx issue #1036 — move slide between decks](https://github.com/scanny/python-pptx/issues/1036)
- [python-pptx issue #533 — duplicate a shape](https://github.com/scanny/python-pptx/issues/533)
- [python-pptx issue #961 — duplicate a16:creationId corruption](https://github.com/scanny/python-pptx/issues/961)
- [python-pptx Group Shape analysis doc](https://python-pptx.readthedocs.io/en/latest/dev/analysis/shp-group-shape.html)
- [Dasc3er PPTX helper gist](https://gist.github.com/Dasc3er/2af5069afb728c39d54434cb28a1dbb8)
- [PPTAgent paper (ACL Anthology, EMNLP 2025)](https://aclanthology.org/2025.emnlp-main.728/)
- [PPTAgent arXiv HTML v3](https://arxiv.org/html/2501.03936v3)
- [PPTAgent GitHub repo](https://github.com/icip-cas/PPTAgent)
- [PPTAgent DeepWiki](https://deepwiki.com/icip-cas/PPTAgent)
- [SlideCoder — EMNLP 2025](https://aclanthology.org/2025.emnlp-main.458.pdf)
