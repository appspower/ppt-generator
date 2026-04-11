# Slide Designer Guide — Claude Code용 화면 구성 판단 가이드

> 이 문서는 Claude Code가 PPT 슬라이드의 화면 구성을 판단할 때 참조합니다.
> **핵심 원칙: 내용이 형태를 결정한다. 형태를 먼저 정하고 내용을 끼워 넣지 않는다.**

---

## §1 HARD RULES (불변 규칙)

1. **Action Title은 완전한 문장** — "매출 현황"(라벨) ❌ → "매출 20% 성장은 디지털 전환에 기인"(주장) ✓
2. **한 슬라이드 = 한 메시지** — 제목에 "and"로 2개 연결하면 분할
3. **컴포넌트 카테고리 2종+** — text+text+text ❌ → chart+text+kpi ✓
4. **같은 레이아웃 연속 사용 금지** — 직전 슬라이드와 다른 레이아웃 필수
5. **빈 존 금지** — 모든 존에 최소 1개 컴포넌트
6. **TakeawayBar 거의 항상 포함** — 핵심 시사점 1문장
7. **출처(footnote) 필수** — 데이터가 있는 슬라이드는 반드시 출처 표시

---

## §2 Exhibit-First 판단 프로세스

> **기존 방식**(메시지 유형 분류 → 매칭 테이블 조회)은 편향을 유발한다.
> **새 방식**: Action Title → 데이터 구조 분석 → Exhibit 결정 → 레이아웃 자동 결정

### Phase A: Action Title 확정

```
"이 슬라이드의 So What은 무엇인가?"

Action Title = 청중이 이 슬라이드에서 가져가야 할 단 하나의 결론.
예: "중국이 희토류 정제의 90%를 독점하고 있어, 서방의 탈중국이 시급하다"

검증:
  - 완전한 문장(주어+서술어)인가?
  - 이 문장만 읽어도 슬라이드를 안 봐도 핵심을 알 수 있는가?
  - "~현황", "~개요" 같은 라벨이 아닌가?
```

### Phase B: 데이터 구조 판단 트리

```
"이 So What을 증명하는 데이터의 구조는 무엇인가?"

Q1. 핵심 증거의 관계 유형은?

├── 크기/양을 비교 → [비교 분기]
│   ├── 전/후, AS-IS/TO-BE (2개) ──────→ comp_before_after
│   ├── 3~5개 옵션 병렬 비교 ──────────→ comp_comparison_grid
│   ├── 2축으로 분류 (영향도×확률) ────→ comp_quadrant_matrix
│   └── 다차원 점수 매트릭스 ──────────→ comp_heatmap_grid
│
├── 시간/순서/단계 → [프로세스 분기]
│   ├── 선형 3~6단계 ──────────────────→ comp_chevron_flow
│   ├── 연도별 마일스톤 ───────────────→ comp_gantt_bars 또는 comp_timeline_marker
│   ├── 순환/반복 프로세스 ────────────→ comp_cycle_arrows
│   └── 번호 붙은 N개 항목 ────────────→ grid_nxm + comp_numbered_cell
│
├── 구성/비중/분해 → [분해 분기]
│   ├── 누적·차감 (비용 Bridge) ────────→ comp_waterfall
│   ├── 단계별 축소 (전환율) ──────────→ comp_funnel
│   ├── 위계·성숙도·피라미드 ──────────→ comp_pyramid
│   └── 가치 흐름 (좌→우) ────────────→ comp_value_chain
│
├── 관계/연결/구조 → [구조 분기]
│   ├── 중심+주변 (플랫폼·생태계) ────→ comp_hub_spoke_diagram
│   ├── 계층 분해 (이슈 트리) ────────→ comp_logic_tree
│   ├── 레이어 스택 (아키텍처) ────────→ comp_architecture_stack
│   └── 4~6방향 전략 ────────────────→ center_peripheral + comp_styled_card
│
├── 단일 핵심 숫자/메시지 강조 → [Hero 분기]
│   ├── 하나의 숫자가 핵심 ────────────→ comp_kpi_card (대형) 또는 comp_hero_block
│   └── 결론·제언 강조 ───────────────→ comp_hero_block
│
└── N개 카드형 정보 병렬 나열 → [카드 분기]
    ├── 4개 카드 ──────────────────────→ grid_2x2 + comp_styled_card ×4
    ├── 3개 카드 ──────────────────────→ three_column + comp_styled_card ×3
    └── 6개 카드 ──────────────────────→ grid_nxm(2×3) + comp_numbered_cell ×6
```

### Phase C: Exhibit → 조연 → 레이아웃 결정

```
1. Phase B에서 주인공 컴포넌트(Exhibit) 결정
2. Exhibit 하나로 So What이 충분히 증명되는가?
   - YES → full 레이아웃
   - NO → 보조 증거 필요 → 조연 컴포넌트 추가

3. 조연 컴포넌트 선택 (Exhibit와 다른 카테고리)
   - 수치 보강이 필요 → comp_kpi_row 또는 comp_kpi_card
   - 맥락 설명 필요 → comp_bullet_list (**볼드** 마크업 필수)
   - 추세 보강 필요 → comp_native_chart
   - 요약 필요 → comp_mini_table

4. 컴포넌트 수와 크기가 레이아웃을 결정:
   - 주인공 1개만 → full
   - 주인공 + 조연 1개 → two_column
   - 주인공(큰) + 조연 2개(작은) → l_layout 또는 t_layout
   - 주인공 + KPI 상단 + 조연 → t_layout
   - 동일 크기 4개 → grid_2x2
   - 동일 크기 3개 → three_column
```

### Phase D: 다양성 검증

```
직전 3장의 슬라이드와 비교:
  □ 같은 주인공 컴포넌트를 사용하지 않는가?
  □ 같은 레이아웃을 사용하지 않는가?
  □ bullet_list 조연이 3장 연속이면 다른 조연으로 교체 (kpi_row, chart, mini_table)

검증 실패 시: Phase B로 돌아가서 데이터를 다른 관점으로 재해석
  예: "비교" → comparison_grid 3번 연속이면,
      같은 데이터를 "다차원 점수"로 재해석 → heatmap_grid
      또는 "프로세스"로 재해석 → chevron_flow
```

---

## §3 매칭 테이블 — 참조용 (Phase B의 보조)

> §2 판단 트리가 우선. 이 테이블은 판단 트리 결과를 교차 검증할 때 사용.

### 비교 메시지

| 내용 특성 | 레이아웃 | 주인공 컴포넌트 | 조연 컴포넌트 |
|---|---|---|---|
| A vs B (2개) | `full` | `comp_before_after` | — |
| 3안 비교 | `full` | `comp_comparison_grid` (highlight=추천) | — |
| 비교 + 수치 | `l_layout` | `comp_comparison_grid`(좌) | `comp_kpi_row`(우상) + `comp_bullet_list`(우하) |
| 2×2 포지셔닝 | `full` | `comp_quadrant_matrix` (축 라벨) | — |

### 프로세스/순서 메시지

| 내용 특성 | 레이아웃 | 주인공 컴포넌트 | 조연 컴포넌트 |
|---|---|---|---|
| 3~6단계 프로세스 | `full` | `comp_chevron_flow` (show_details=True) | — |
| 프로세스 + KPI | `t_layout`(0.22) | `comp_chevron_flow`(상) | `comp_kpi_row`(하좌) + `comp_bullet_list`(하우) |
| 타임라인 로드맵 | `full` | `comp_gantt_bars` | — |
| 타임라인 + 해설 | `timeline_band` | `comp_timeline_marker`(밴드) | `comp_bullet_list`(각 스텝) |
| 순환 프로세스 | `full` | `comp_cycle_arrows` | — |
| 순환 + 설명 | `center_peripheral_4` | `comp_cycle_arrows`(중앙) | `comp_bullet_list`(4방향) |

### 구조/관계 메시지

| 내용 특성 | 레이아웃 | 주인공 컴포넌트 | 조연 컴포넌트 |
|---|---|---|---|
| 중심+주변 관계 | `full` | `comp_hub_spoke_diagram` | — |
| 중심+주변 + 설명 | `center_peripheral_4` | `comp_hub_spoke_diagram`(중앙) | `comp_bullet_list`(4방향) |
| 기술 스택/계층 | `full` | `comp_architecture_stack` | — |
| 스택 + 설명 | `two_column` | `comp_architecture_stack`(좌) | `comp_bullet_list`(우) |
| 가치사슬 | `full` | `comp_value_chain` | — |
| 이슈 트리/분해 | `full` | `comp_logic_tree` | — |
| 전략 위계 | `full` | `comp_pyramid` | — |
| 위계 + 지표 | `two_column` | `comp_pyramid`(좌) | `comp_kpi_row`(우상) + `comp_bullet_list`(우하) |

### 성과/데이터 메시지

| 내용 특성 | 레이아웃 | 주인공 컴포넌트 | 조연 컴포넌트 |
|---|---|---|---|
| KPI 3~4개 | `t_layout`(0.30) | `comp_kpi_row`(상) | `comp_bullet_list`(하좌) + `comp_native_chart`(하우) |
| Hero 숫자 1개 | `l_layout`(0.35) | `comp_kpi_card`(좌, 대형) | `comp_native_chart`(우상) + `comp_bullet_list`(우하) |
| 비용/매출 분해 | `full` | `comp_waterfall` | — |
| 분해 + 프로세스 | `top_bottom`(0.20) | `comp_chevron_flow`(상) | `comp_waterfall`(하) |
| 리스크 맵 | `full` | `comp_heatmap_grid` | — |
| 리스크 + 대응 | `two_column` | `comp_heatmap_grid`(좌) | `comp_bullet_list`(우) |
| 시장 기회 | `full` | `comp_funnel` | — |
| 차트 + 해설 | `two_column` | `comp_native_chart`(좌) | `comp_bullet_list`(우) |

### 전략/방향 메시지

| 내용 특성 | 레이아웃 | 주인공 컴포넌트 | 조연 컴포넌트 |
|---|---|---|---|
| 핵심 메시지 강조 | `full` | `comp_hero_block` | — |
| 전략 + 근거 | `l_layout`(0.45) | `comp_hero_block`(좌) | `comp_kpi_row`(우상) + `comp_bullet_list`(우하) |
| 4대 방향 | `center_peripheral_4` | 중앙 accent box | `comp_bullet_list`(4방향) |
| 6대 역량 | `center_peripheral_6` | 중앙 도형 | `comp_bullet_list`(6방향) |
| 번호 프로세스 | `grid_nxm`(2×3) | `comp_numbered_cell` ×6 | — |

---

## §4 WORKED EXAMPLES — Exhibit-First 사고방식

### 예시 1: "중국이 희토류 정제의 90%를 독점하고 있다"

```
Phase A: So What = "90% 독점" → 하나의 숫자가 핵심
Phase B: Q1 → "단일 핵심 숫자 강조" → Hero 분기 → comp_hero_block 또는 comp_kpi_card
Phase C: Hero 하나로 충분? NO → 세부 데이터(원소별 점유율) 보강 필요
         → 조연: comp_heatmap_grid (원소×차원 매트릭스)
         → 컴포넌트 2개(대+중) → l_layout
결과: l_layout + comp_kpi_card("90%", 좌) + comp_heatmap_grid(우)
```

비교: 기존 방식이면 "비교 메시지 → comparison_grid"로 갔을 것. **데이터 구조를 먼저 보면 Hero+Heatmap 조합이 자연스럽게 나온다.**

### 예시 2: "4단계를 거쳐 클라우드를 전환하되, 비용 절감이 각 단계에서 발생한다"

```
Phase A: So What = "4단계 + 단계별 비용 절감"
Phase B: Q1 → "시간/순서/단계" → 프로세스 분기 → comp_chevron_flow
         + "비용 분해" → 분해 분기 → comp_waterfall
Phase C: 두 Exhibit가 서로 다른 관점을 증명 → 둘 다 필요
         → 컴포넌트 2개(중+중) → top_bottom
결과: top_bottom + comp_chevron_flow(상) + comp_waterfall(하)
```

### 예시 3: "플랫폼 중심으로 5개 시스템이 연결된다"

```
Phase A: So What = "중심+주변 관계"
Phase B: Q1 → "관계/연결/구조" → 구조 분기 → comp_hub_spoke_diagram
Phase C: Hub-spoke 하나로 충분? YES → full
결과: full + comp_hub_spoke_diagram
```

### 예시 4: "리스크 5개 중 기술·가격이 가장 위험하고, 각각 대응 전략이 다르다"

```
Phase A: So What = "리스크별 영향도·확률 + 대응 전략"
Phase B: Q1 → "다차원 점수" → 비교 분기 → comp_heatmap_grid
Phase C: Heatmap만으로 대응 전략은 못 보여줌 → 조연 필요
         → comp_bullet_list (리스크별 대응, **볼드** 마크업)
         → 컴포넌트 2개(중+중) → two_column
결과: two_column + comp_heatmap_grid(좌) + comp_bullet_list(우)
```

### 예시 5: "6대 전략 방향을 동시에 추진한다"

```
Phase A: So What = "6개 방향 병렬"
Phase B: Q1 → "N개 카드형 정보 병렬" → 카드 분기 → grid_nxm(2×3) + comp_numbered_cell
Phase C: 각 카드에 번호+제목+3줄 설명 → full
결과: grid_nxm(2×3) + comp_numbered_cell ×6
```

---

## §5 SELF-CHECK (코드 생성 전 자기 검증)

```
Exhibit-First 검증:
  □ "이 Exhibit가 Action Title을 가장 강력하게 증명하는가?"
  □ "다른 Exhibit로 같은 So What을 더 효과적으로 보여줄 수 있지 않은가?"
  □ "매칭 테이블을 기계적으로 조회하지 않았는가?"

HARD RULES 검증:
  □ Action Title이 완전한 문장인가?
  □ 본문의 모든 요소가 Action Title을 증명하는가?
  □ 2개+ 컴포넌트 카테고리를 사용했는가?
  □ 직전 슬라이드와 다른 레이아웃인가?
  □ 빈 존이 없는가?
  □ TakeawayBar를 포함했는가?

다양성 검증:
  □ 직전 3장과 같은 주인공 컴포넌트를 쓰지 않았는가?
  □ bullet_list 조연이 3장 연속이면 다른 조연으로 교체했는가?
```

---

## §6 밀도 규칙

- 빈 공간이 50% 이상이면 구성 재검토
- 비교 그리드 셀: **40~80자** (2~3줄). 1줄짜리 셀 금지. **볼드 키워드** + 설명 구조 권장
- 불릿 항목: **30자+** 필수. `**볼드 키워드** — 설명` 형식
- 불릿 개수: 4~6개 적정
- KPI 숫자는 20-26pt, 본문은 8~10pt, 각주 7pt
- 모든 텍스트는 한글 기준으로 밀도 계산 (한글은 영문 대비 1.5배 폭)
- 고밀도(3+ 컴포넌트) 슬라이드는 3장 연속 금지 — 중간에 Hero/단순 슬라이드 삽입

---

## APPENDIX: 사용 가능한 컴포넌트 (42개)

### Compound (16개) — 구조적 시각 요소
comp_chevron_flow, comp_hero_block, comp_hub_spoke_diagram, comp_comparison_grid,
comp_architecture_stack, comp_pyramid, comp_cycle_arrows, comp_waterfall,
comp_before_after, comp_gantt_bars, comp_value_chain, comp_logic_tree,
comp_quadrant_matrix, comp_funnel, comp_callout_annotation, comp_heatmap_grid

### Atomic (26개) — 원자 단위 요소
comp_kpi_card, comp_kpi_row, comp_mini_table, comp_bullet_list, comp_bar_chart_h,
comp_stat_row, comp_callout, comp_rag_row, comp_numbered_items, comp_section_header,
comp_progress_bar, comp_vertical_bars, comp_heat_row, comp_gauge, comp_tag_group,
comp_comparison_row, comp_metric_delta, comp_timeline_mini, comp_icon_list,
comp_data_card, comp_icon_card, comp_icon_row, comp_styled_card, comp_styled_card_row,
comp_native_chart, comp_numbered_cell, comp_timeline_marker, comp_icon_header_card

### 레이아웃 (15개)
full, two_column, top_bottom, three_column, four_column, grid_2x2,
sidebar_left, sidebar_right, center_peripheral_4, center_peripheral_6,
grid_nxm, timeline_band, asymmetric_lr, t_layout, l_layout
