# Phase A2 — 1,251장 마스터 템플릿 데이터 프로파일링 (종합)

> 작성일: 2026-04-25
> 목적: 자동 클러스터링 시도 이전에 자산 자체를 정밀 분석. 사용자 지시.
> 입력: 1,251장 마스터 PPT (`docs/references/_master_templates/PPT 템플릿.pptx`)

---

## 1. 분석 방법 (3 Agent 병렬 + Collage 전수 시각)

| Phase | 방법 | 출력 |
|---|---|---|
| 1A | Pillow + numpy 시각 통계 (color/split/entropy) | `phase1a_visual_stats.json` (1.9 MB) |
| 1B | python-pptx OOXML 의미구조 (group/grid/sig) | `phase1b_semantic_stats.json` (1.4 MB) |
| 1C | OOXML 휴리스틱 컴포넌트 검출 (24 type detector) | `phase1c_component_types.json` (663 KB) |
| 2 | 13장 collage (10×10 cell) 전수 시각 검토 | `collages/collage_*.png` 13장 |

---

## 2. Phase 1A — 시각 통계 (실측)

### 2.1 색상 프로필
| 측정 | 값 |
|---|---|
| 회색조 슬라이드 | **716장 (57.2%)** |
| 컬러 슬라이드 | 535장 (42.8%) |
| 강한 accent (≥10% 채도) | 69장 (5.5%) |
| 중간 accent (2~10%) | 238장 (19.0%) |
| accent_ratio mean / p90 / max | 0.020 / 0.062 / 0.318 |
| **컬러 hue 분포** | orange **399** / blue 81 / yellow 26 / green 23 / red 6 |

### 2.2 영역 점유 / 분할
| 측정 | 값 |
|---|---|
| 평균 blank_ratio | **0.82** (82% 흰 배경) |
| sparse(blank≥0.6) | 1,221장 (**97.6%**) |
| dense(blank<0.3) | **0장** |
| **분할 패턴** | balanced 765 / tb_split 201 / lr_split 183 / 4-split 59 / 3-split 31 |

### 2.3 정렬 entropy (8×8 grid)
| 분포 | 슬라이드 |
|---|---|
| high (≥0.7) | **1,208장 (96.6%)** |
| mid | 42 |
| low | 1 |

→ **96.6%가 한 점에 몰려 entropy 단일 축 변별 불가**

### 2.4 시각 element 면적
- chart-like (고분산 타일) heavy (≥40%): **42장**
- big_box (저분산 컬러) heavy (≥10%): **82장**
- text-dense (≥5%): **54장**
- 흑백 차트/표: **93장**

### 2.5 Outlier (시각 확인)
- #149 (11×5 흑백 매트릭스), #387 (다중 패턴 결합), #642 (PwC 톤 이탈), #794 (좌우 명확 대비), #864/1222/1223 (단일 큰 도형 cover)

---

## 3. Phase 1B — OOXML 의미 구조 (실측)

### 3.1 Group nesting (사람이 묶은 의미 단위)
| 측정 | 값 |
|---|---|
| Group 사용 슬라이드 | **1,115 / 1,251 (89.1%)** |
| 슬라이드당 top-level group (median/mean/max) | 2 / 2.87 / 30 |
| Group당 leaf shape (median/mean/p95) | **3** / 5.83 / 17 |
| Nesting depth | 1 (569) / 2 (403) / 3 (101) / 4 (32) / 5+ (10) |

→ **표준 group = 3 leaf** = "icon+headline+body" 또는 "number+label+bar" 의미 단위

### 3.2 Title 위치 (가장 큰 폰트)
| 위치 | 슬라이드 |
|---|---|
| Top-band (3종 합) | **956장 (76.5%)** |
| top_center | 27.6% |
| top_right | 24.5% |
| top_left | 24.4% |
| left_sidebar / right_sidebar | 4 / 3 (희귀) |

### 3.3 Shape 정렬 격자 (자동 추론)
| Archetype | 슬라이드 | % |
|---|---:|---:|
| **left_title_right_body** | **355** | **28.4%** |
| **mixed (다중 패턴)** | **321** | **25.7%** |
| **dense_grid (≥4×4)** | **277** | **22.1%** |
| 3×3 matrix | 55 | 4.4% |
| 4-col compare | 40 | 3.2% |
| 2×2 matrix | 38 | 3.0% |
| 3-col compare | 38 | 3.0% |
| single_block | 33 | 2.6% |
| 5-col compare | 30 | 2.4% |
| vertical_list | 20 | 1.6% |
| 6/2/7-col + cover/divider/accent | 44 | 3.5% |

→ **5개 archetype이 60.6% 커버 / 격자 99.5%**

### 3.4 시그니처 요소
- top_divider: **317장 (25.3%)** ← 가장 빈번한 사인
- corner_top_left_marker: 95장 (7.6%) ← 사용자 임의 입력 코드 영역
- bottom_divider: 78 / left_accent_strip: 7 / right_accent_strip: 4
- page_number: bottom-center **943장 (99% 일관)**

### 3.5 Placeholder (~~) 분포 (8×8)
- per-slide ~~ count: median=38, mean=39.84, p95=83
- 가장 빈번 셀: **top-right 1193장 / top-center 942 / mid-upper-col4 919**

---

## 4. Phase 1C — 컴포넌트 type 자동 검출

### 4.1 Multi-label 분포 (24 detector)

| Archetype | 검출 | 정확도 추정 |
|---|---:|---|
| orgchart | 729 | **과다검출** (실제 50~150) |
| table_native | 366 | **거의 정확** (slide.has_table) |
| cards_6col | 219 | medium (실제 50~80) |
| hub_spoke | 195 | medium (실제 30~80) |
| flowchart | 179 | high 42 / medium 137 |
| roadmap | 168 | **상당히 정확** |
| cards_2col | 144 | low (실제 30~50) |
| unclassified | 131 | — |
| gantt_like | 122 | high 88 / medium 34 |
| cards_3col | 71 | medium |
| timeline_h | 62 | high 38 / medium 24 |
| cards_4col | 56 | medium |
| picture_chart_like | 43 | low |
| cover_or_divider | 40 | medium (**누락 많음** 실제 60~100) |
| swimlane | 12 | high |
| chart_native | 12 | **정확** (slide.has_chart) |
| matrix_3x3 | 9 | medium |
| matrix_2x2 | 7 | high |
| venn | 7 | high |

### 4.2 Primary archetype 분포

| 카테고리 | 슬라이드 | % |
|---|---:|---:|
| C2 table-like | 468 | 37.4% |
| C3 diagram | 871 | 69.6% |
| C4 card | 512 | 40.9% |
| C1 chart-like | 55 | 4.4% |
| C5 cover/divider | 40 | 3.2% |
| unclassified | 131 | 10.5% |

### 4.3 Multi-label 정도
- 1 type 만 매칭: 472 (37.7%)
- 2 types: 388 (31.0%)
- 3 types: 258 (20.6%)
- 4+ types: 133 (10.6%)

→ **62%가 multi-label** = 한 슬라이드 = 여러 archetype 혼재. 단일 라벨 클러스터링 부적합 입증.

### 4.4 자동 검출 어려운 archetype
1. orgchart vs grid — OOXML 동일, freeform `<sp>` + `<cxnSp>`만
2. flowchart vs roadmap vs swimlane — 모두 박스+화살표
3. cards vs feature matrix — 헤더/row label 유무만 차이
4. cover/divider — placeholder가 비어있어 텍스트 길이 무용
5. picture_chart_like — 차트 PNG와 지도/스크린샷 구별 어려움
6. 수직 카드 stack — 가로 정렬 detector만 있음

---

## 5. 자동 클러스터링 실패의 정량적 진단

| 시도 | 결과 | 진단 (1A/1B/1C 통합) |
|---|---|---|
| Stage 2 v1 BGE-M3 | 12 클러스터, max=933 | 텍스트 ~~로 동질화 (placeholder 49,846개) |
| Stage 2a OOXML 68-dim | 51 클러스터, silhouette 0.034 | grid6x6 거침. **1B 14 archetype 미반영** |
| Stage 2b DiT only | 3 클러스터, max=1241 | 흑백 sparse + 오렌지 dominant 615장이 동질로 보임 |
| Stage 2b 앙상블 | 50 클러스터, 변화 미미 | 차원 imbalance + **archetype label 비활용** |

**근본 원인**:
- 96.6% highE → entropy 무력
- 615 오렌지 → vision dominant 동질
- 99.5% 격자 → 격자 자체로는 변별 불가
- **62% multi-label** → 단일 라벨 클러스터링 본질 부적합

---

## 6. 종합 — 자산의 진짜 정체

1,251장은:
1. **Process/diagram 중심 컨설팅 자산** (1C: ~80% multi-label로 process/structure 보유)
2. **PwC 시그니처 톤** (1A: 회색조 57% + 오렌지 단일 dominant 615)
3. **격자 99.5% + 표준 frame (top title 76% + bottom page# 99%)** = 외형 거의 동일
4. **Group 89%로 의미 단위 명시** (median 3 leaf = 컴포넌트 1개)
5. **Mixed 26% (1B) + Multi-label 62% (1C)** = "한 슬라이드 = N archetype" 본질
6. **Sparse 97.6% / 차트 4% / 표 37% / 카드 41% / 다이어그램 70%** = 균형 분포

**진정한 변별 차원** (1A+1B+1C):
```
gray vs color  ×  sparse vs mid  ×  layout_split (5종)
                ×  detected_archetypes (multi-label)
                ×  group_signature (의미 단위)
```

---

## 7. 분류 전략 옵션 (사용자 의사결정 필요)

이전 자동 클러스터링 (BGE-M3, DiT, OOXML grid6x6)이 모두 실패한 이유는 위 차원 중 **하나 또는 둘만** 사용했기 때문. 진짜 분류는 **multi-axis hierarchical**여야 함.

### 옵션 A — Multi-Label Detection Library (자동 70~75%)
**구조**:
- L1 Macro tag (1A 기반): `gray|color × sparse|dense × balanced|tb|lr|3split|4split` = 약 30 macro 버킷
- L2 Archetype tag (1B 14종 + 1C 24종 multi-label): 한 슬라이드 = 1~5 tag
- L3 Component group (1B group 89% + median 3 leaf): 의미 단위 retrieval

**장점**: 70~75% 자동 + multi-label 보존. evaluate.py + edit_ops 호환
**단점**: orgchart/cards/cover detector refinement 필요 (1~2주)
**예상 시간**: refinement + 검증 = 1주

### 옵션 B — 35 hand-curated templates ↔ 1,251 매핑 (Supervised, 선택적)
**구조**:
- 기존 `metadata.json`의 35 템플릿 (intent + content_types + density)을 anchor로
- 1,251장 각각을 35 템플릿 중 하나에 mapping (또는 "신규" 분류)

**장점**: 즉시 사용 가능 + 사용자 도메인 지식 직접 반영
**단점**: 1,251장 수동 매핑 시간 (사용자 검토 필수)
**예상 시간**: 사용자 시간 의존, 8시간+

### 옵션 C — Hybrid (옵션 A + B)
**구조**:
- 옵션 A의 자동 detection으로 1차 라벨링 (70%)
- 35 templates와 매핑해서 사용자 검토 우선순위 결정
- 사용자가 의문점만 검수 (~200장 추정)

**장점**: 자동 + 수동 균형. 사용자 시간 최소화
**단점**: 두 시스템 동기화 필요
**예상 시간**: 자동 1주 + 사용자 검토 4시간

### 옵션 D — 자동 클러스터링 포기 + Component Library Expansion
**구조**:
- 1,251장은 raw 자산으로 두고 retrieval은 metadata.json의 35 + 컴포넌트 라이브러리만 사용
- 1,251장은 **사용자가 즉석 선택**하는 백업 풀

**장점**: 의사결정 비용 0. evaluate/edit 인프라만 활용
**단점**: 자산 70% 사용 안 함 (낭비)

---

## 8. 산출 파일 (전체)

| 경로 | 용도 |
|---|---|
| `output/catalog/slide_meta.json` | Phase 1 메타 (1,251장) |
| `output/catalog/phase1a_visual_stats.json` | 시각 통계 |
| `output/catalog/phase1b_semantic_stats.json` | OOXML 의미 |
| `output/catalog/phase1c_component_types.json` | 컴포넌트 multi-label |
| `output/catalog/all_pngs/slide_*.png` | 1,251 PNG 렌더 |
| `output/catalog/collages/collage_*.png` | 13장 grid collage |
| `output/catalog/shape_features.npy` | OOXML 68-dim (이전 시도) |
| `output/catalog/dit_embeddings.npy` | DiT 768-dim (이전 시도) |
| `output/catalog/slot_schemas.json` | max_chars 슬롯 (Stage 2d) |
| `output/catalog/skeletons.json` | 7 narrative 스켈레톤 (Track 2) |

스크립트:
- `scripts/visual_stats_phase1a.py` (1A)
- `scripts/phase1b_semantic_stats.py` (1B)
- `scripts/detect_component_archetypes.py` (1C, 24 detector)
- `scripts/build_collage.py` (collage 생성)
- 기존: deep_layout_analysis / feature_similarity_experiment / xml_placeholder_probe
