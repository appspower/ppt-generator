# Slide Designer Guide — Claude Code용 화면 구성 판단 가이드

> 이 문서는 Claude Code가 PPT 슬라이드의 화면 구성을 판단할 때 참조합니다.
> 컴포넌트 42개 + 레이아웃 15개를 조합하여 슬라이드를 구성합니다.

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

## §2 판단 프로세스

```
Step 1: Action Title 작성 → 메시지 유형 분류 (비교/프로세스/구조/성과/전략)
Step 2: 콘텐츠 블록 수 파악 (1~4+)
Step 3: 아래 매칭 테이블에서 레이아웃 + 컴포넌트 선택
Step 4: HARD RULES 확인 (직전 슬라이드와 같은 레이아웃 아닌지 등)
Step 5: 코드 생성 (SlideComposer + comp_xxx)
```

---

## §3 매칭 테이블 — 콘텐츠 유형 → 레이아웃 + 컴포넌트

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

## §4 WORKED EXAMPLES (실전 예시)

### 예시 1: "매출 20% 성장은 디지털 전환 투자에 기인"

**판단**: 성과 메시지, Hero 숫자(20%) + 증거(차트) + 맥락(불릿) → 3블록 → l_layout
```python
composer = SlideComposer(slide)
composer.header(SlideHeader(title="매출 20% 성장은 디지털 전환 투자에 기인", category="Performance"))
zones = composer.layout("l_layout", left_ratio=0.35, top_ratio=0.55)
comp_kpi_card(canvas, value="20%", label="YoY 성장", trend="up", region=zones["left_full"])
comp_native_chart(canvas, chart_type="vertical_bar", ..., region=zones["right_top"])
comp_bullet_list(canvas, title="성장 동인", items=[...], region=zones["right_bottom"])
composer.takeaway("디지털 전환 투자 ROI 340% — 2027년까지 연 15% 추가 성장 전망")
```

### 예시 2: "ERP 전환은 4단계로 진행되며, 각 단계별 명확한 산출물이 정의됨"

**판단**: 프로세스 메시지, 4단계 + 산출물 상세 → 2블록 → top_bottom 또는 full(show_details)
```python
zones = composer.layout("full")
comp_chevron_flow(canvas, phases=[
    {"tag": "P1", "label": "진단", "details": ["현행 분석", "Gap 도출"]},
    {"tag": "P2", "label": "설계", "details": ["솔루션 설계", "아키텍처"]},
    {"tag": "P3", "label": "구현", "details": ["개발", "테스트"]},
    {"tag": "P4", "label": "안정화", "details": ["Go-Live", "모니터링"]},
], show_details=True, region=zones["main"])
```

### 예시 3: "클라우드 전환 3안 중 Hybrid 방식이 비용-리스크 균형 최적"

**판단**: 비교 메시지, 3안 비교 + 추천안 하이라이트 → 1블록 → full
```python
zones = composer.layout("full")
comp_comparison_grid(canvas, columns=[
    {"name": "On-Prem", "criteria": ["자체 운영", "높음", "느림"]},
    {"name": "Hybrid", "highlight": True, "criteria": ["혼합", "중간", "유연"]},
    {"name": "Full Cloud", "criteria": ["완전 위탁", "낮음", "빠름"]},
], row_labels=["운영 방식", "초기 비용", "확장성"], region=zones["main"])
```

### 예시 4: "통합 플랫폼 중심으로 5개 시스템이 연결되는 생태계 구축"

**판단**: 구조 메시지, 중심+5 주변 → 1블록(다이어그램) → full
```python
zones = composer.layout("full")
comp_hub_spoke_diagram(canvas, center="통합\n플랫폼", spokes=[
    {"title": "ERP", "detail": "재무, 구매, 생산"},
    {"title": "CRM", "detail": "고객 관리"},
    {"title": "SCM", "detail": "공급망"},
    {"title": "Analytics", "detail": "BI + AI/ML"},
    {"title": "IoT", "detail": "설비 모니터링"},
], region=zones["main"])
```

### 예시 5: "운영비용 100억에서 자동화·클라우드로 25억 절감, 최종 75억 달성"

**판단**: 성과 메시지 + 프로세스 → 쉐브론(단계) + 워터폴(금액) → 2블록 → top_bottom
```python
zones = composer.layout("top_bottom", split=0.20)
comp_chevron_flow(canvas, phases=[
    {"tag": "1", "label": "진단"}, {"tag": "2", "label": "자동화"},
    {"tag": "3", "label": "클라우드"}, {"tag": "4", "label": "최적화"},
], region=zones["top"])
comp_waterfall(canvas, start={"label": "현재", "value": 100},
    steps=[{"label": "자동화", "value": -15}, {"label": "클라우드", "value": -10}],
    end={"label": "목표", "value": 75}, unit="억", region=zones["bottom"])
```

---

## §5 SELF-CHECK (코드 생성 전 자기 검증)

```
□ Action Title이 완전한 문장인가? (라벨형 아닌가?)
□ 본문의 모든 요소가 Action Title을 증명하는가?
□ 2개+ 컴포넌트 카테고리를 사용했는가? (Structural+Data, Data+Text 등)
□ 직전 슬라이드와 다른 레이아웃인가?
□ 빈 존이 없는가?
□ TakeawayBar를 포함했는가?
```

---

## §6 밀도 규칙

- 빈 공간이 50% 이상이면 구성 재검토
- 카드 내부는 불릿 3~5개가 적정
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
