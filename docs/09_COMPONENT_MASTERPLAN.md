# 09. 컴포넌트 마스터플랜 — 패턴 해체 + 자유 조합 시스템

> "패턴"이라는 틀을 해체하고, 시각 요소를 컴포넌트 단위로 쪼갠 뒤,
> Claude Code가 내용과 상황에 맞춰 자유롭게 조합하는 시스템을 구축한다.

---

## 0. 왜 이 전환이 필요한가

### 현재 상태
```
25개 패턴 (절대좌표, 전체 슬라이드 점유)
  → 호출하면 항상 같은 모양
  → 조합 불가 (1패턴 = 1슬라이드)
  → 패턴 수를 아무리 늘려도 조합의 자유도는 0
```

### 목표 상태
```
~15개 Compound Component (Region-aware, 어디든 배치 가능)
  + 15개 레이아웃 (직사각형 + 비직사각형)
  + Claude Code 판단 가이드
  → 매 슬라이드마다 다른 조합
  → 한 슬라이드에 차트+KPI+불릿+도형이 유기적으로 공존
```

### PwC 완성본과의 차이
| 관점 | 현재 (패턴) | 목표 (컴포넌트 조합) |
|------|-----------|-------------------|
| 슬라이드당 시각 유형 | 1개 | 2~4개 |
| 같은 주제의 변주 | 패턴 이름만 바뀜 | 레이아웃×컴포넌트 조합으로 수십 가지 |
| 밀도 | 패턴이 정한 고정 밀도 | 존 개수×컴포넌트 밀도로 동적 조절 |
| 정보 계층 | 패턴 내부에 암묵적 | 존 크기+위치로 명시적 |

---

## 1. 25개 패턴의 시각 DNA 분석

### 1.1 추출할 가치가 높은 시각 요소 (HIGH VALUE)

패턴 안에 묻혀 있는 **고유 시각 구조**를 Region-aware 컴포넌트로 추출한다.

| # | 추출할 컴포넌트 | 원본 패턴 | 시각 구조 | 예상 규모 | 재사용 빈도 | 비고 |
|---|-------------|---------|---------|---------|-----------|------|
| 1 | `comp_chevron_flow` | timeline_phases, executive_summary, value_chain, chevron_timeline | 수평 화살표 체인 (3~6단계) | ~40줄 | **최고** | 4개 패턴에서 중복 사용 — 최우선 추출 |
| 2 | `comp_hero_block` | executive_summary | 대형 색상 박스 + 큰 텍스트 + 하위 불릿 | ~60줄 | 매우 높음 | 전략 선언/섹션 도입에 필수 |
| 3 | `comp_hub_spoke_diagram` | hub_spoke | 중앙 원 + 방사형 연결선 + 외곽 원 | ~70줄 | 높음 | 삼각함수 기반 자동 배치, 정사각형 region 권장 |
| 4 | `comp_comparison_grid` | comparison_matrix | N열 비교 표 (색상 헤더 + 행 데이터) | ~80줄 | 높음 | comp_mini_table과 차별화: 컬럼 강조 기능 |
| 5 | `comp_pyramid` | pyramid_layers | 층별 너비 점감 스택 | ~50줄 | 중간 | 직사각형 region에 잘 맞음 |
| 6 | `comp_before_after` | before_after | 좌우 패널 + 연결 화살표 + 색상 대비 | ~55줄 | 중간 | two_column + 화살표 조합으로 대체 가능하나, 전용 컴포넌트가 품질 높음 |
| 7 | `comp_cycle_arrows` | cycle_diagram | 4~6노드 순환 화살표 | ~75줄 | 중간 | 원형 기하학, 비정사각형 region 주의 |
| 8 | `comp_waterfall` | waterfall_bridge | 증감 브릿지 바 차트 | ~85줄 | 중간 | 복잡한 좌표 계산이지만 재무 분석 필수 |
| 9 | `comp_architecture_stack` | architecture_stack | 수직 레이어 스택 (기술 스택) | ~40줄 | 중간 | 가장 단순한 구현 |
| 10 | `comp_gantt_bars` | gantt_roadmap | 수평 바 타임라인 | ~55줄 | 중간 | 프로젝트 일정 시각화 |

### 1.2 기존 컴포넌트로 대체 가능 (MEDIUM VALUE)

새로 만들지 않고 기존 컴포넌트 조합으로 구현 가능한 패턴:

| 패턴 | 대체 방법 |
|------|---------|
| quadrant_story | `grid_2x2` 레이아웃 + `comp_bullet_list` ×4 + 축 라벨 |
| before_after | `two_column` 레이아웃 + `comp_styled_card` ×2 + 화살표 |
| kpi_dashboard | `grid_2x2` + `comp_kpi_card` ×4 (이미 있음) |
| data_narrative | `two_column` + `comp_native_chart` + `comp_bullet_list` |
| maturity_model | `comp_progress_bar` ×N (이미 있음) |
| milestone_timeline | `comp_timeline_marker` (Phase 1에서 구현됨) |
| grid_process | `grid_nxm` + `comp_numbered_cell` (Phase 1에서 구현됨) |

### 1.3 기존 컴포넌트에 통합 (CONSOLIDATE)

새로 만들지 않고, 기존 컴포넌트를 확장하여 해결:

| 패턴 | 통합 대상 | 방법 |
|------|---------|------|
| kpi_dashboard | `comp_kpi_row` (이미 있음) | grid_2x2 + comp_kpi_card ×4로 대체 |
| comparison_matrix | `comp_mini_table` 확장 | 컬럼 하이라이트 옵션 추가 |
| rag_status_table | `comp_mini_table` 확장 | RAG 색상 셀 옵션 추가 |
| quadrant_story | `grid_2x2` + `comp_bullet_list` | 2x2 코어만 추출 (~30줄), 축 라벨 추가 |

### 1.4 패턴으로 유지 (KEEP AS PATTERN)

조합보다 단독 사용이 더 자연스럽거나, 복잡한 오케스트레이션이 필요한 패턴:

| 패턴 | 유지 이유 |
|------|---------|
| executive_summary | 복합 오케스트레이션 (hero+kpi+chevron), Phase 2에서 내부를 컴포넌트 조합으로 재작성 |
| timeline_phases | 동적 높이+aux 콘텐츠, chevron 컴포넌트 추출 후 내부 위임 |
| process_flow | arrow_chain() 프리미티브 직접 사용, 추출 불필요 |
| data_narrative | 텍스트 배치 중심, 시각적 고유성 낮음 |
| bubble_chart | matplotlib 의존, Canvas 독립 추출 어려움 |
| tree_diagram | 계층 구조 렌더링 복잡, 후순위 (Phase 3+) |
| harvey_ball_matrix | comp_mini_table 확장으로 대체 가능 |
| diamond_four | Phase 2 앵커 컴포넌트로 이미 계획됨 |
| chevron_timeline | comp_chevron_flow 추출 후 내부 위임 |

---

## 2. 컴포넌트 인터페이스 표준화

### 2.1 통일 API 설계

모든 컴포넌트는 동일한 시그니처 규약을 따른다:

```python
def comp_xxx(
    c: Canvas,           # 캔버스 (필수)
    *,                   # 이후 모두 키워드 전용
    # --- 데이터 파라미터 (컴포넌트마다 다름) ---
    data_param_1: ...,
    data_param_2: ...,
    # --- 공통 스타일 파라미터 ---
    accent_color: str = "accent",  # 강조색
    text_color: str = "grey_900",  # 본문 텍스트 색
    show_border: bool = True,      # 테두리 표시
    # --- 필수 ---
    region: Region,      # 배치 영역 (필수)
) -> float:              # 사용한 높이 반환
    """컴포넌트 설명.
    
    Args: ...
    Returns: 사용한 높이 (인치). 호출자가 다음 컴포넌트 배치에 활용.
    """
```

### 2.2 컴포넌트 카테고리

| 카테고리 | 역할 | 예시 |
|---------|------|------|
| **Structural** | 정보의 구조/관계를 시각화 | chevron_flow, hub_spoke, pyramid, cycle_arrows |
| **Data** | 정량 데이터를 시각화 | native_chart, kpi_card, waterfall, gantt_bars |
| **Text** | 텍스트 정보를 구조적으로 표현 | bullet_list, comparison_grid, mini_table |
| **Anchor** | 레이아웃의 시각적 중심 (주변 존 정의) | diamond_anchor, hexagon_anchor, donut_anchor |
| **Decorative** | 보조적 시각 강조 | icon_header_card, hero_block, styled_card |

### 2.3 컴포넌트 호환성 매트릭스

같은 슬라이드에 함께 쓸 때의 궁합:

```
                    Structural  Data  Text  Anchor  Decorative
Structural            △         ◎     ◎     ◎       ○
Data                  ◎         △     ◎     ○       ◎
Text                  ◎         ◎     △     ◎       ◎
Anchor                ◎         ○     ◎     ✕       ○
Decorative            ○         ◎     ◎     ○       △

◎ = 최적 조합  ○ = 가능  △ = 같은 유형 2개는 주의  ✕ = 비권장
```

**규칙**: 한 슬라이드에 최소 2개 카테고리의 컴포넌트를 사용해야 시각 다양성이 확보된다.

---

## 3. 컨설팅 슬라이드 구성 원칙 (리서치 기반)

### 3.0 Three-Zone Anatomy (McKinsey/BCG/Bain 공통)

모든 컨설팅 슬라이드는 3개 존으로 구성된다:

```
┌─────────────────────────────────────┐
│  Header Zone (~15%)                 │  ← Action Title (결론 문장, 15단어 이내, 2줄 이내)
│  "매출 20% 성장은 디지털 전환에 기인" │
├─────────────────────────────────────┤
│                                     │
│  Body Zone (~70-75%)                │  ← 단일 Exhibit (차트/표/프레임워크)
│  [차트] [KPI] [불릿]                │     Action Title을 증명하는 근거
│                                     │
├─────────────────────────────────────┤
│  Footer Zone (~10%)                 │  ← 출처 표시 (8-9pt), 페이지 번호
└─────────────────────────────────────┘
```

**인코딩 규칙**: `title_max_words=15, body_ratio=0.70, footer_font=8pt, source_required=True`

### 3.1 "So What" 원칙 — 모든 요소는 주장을 증명해야 한다

```
정보 계층:
  1. Action Title = 결론 ("So What")
  2. 2~4개 지지 논거 ("Why?")
  3. 근거 데이터/차트 ("How do we know?")
```

**So What 테스트**: 슬라이드 위의 어떤 요소든 가리키고 "So What?"을 물었을 때, Action Title과 연결 안 되면 삭제.
**수평 흐름 테스트**: 모든 슬라이드의 Action Title만 읽었을 때 전체 논리가 이해되어야 함.

→ **코드 적용**: evaluate.py에서 제목 길이 체크 + 라벨형(짧은 명사) 제목 경고

### 3.2 One-Message Rule + 밀도 제한

| 규칙 | 값 | 근거 |
|------|---|------|
| 슬라이드당 메시지 | **1개** | 제목에 "and"로 2개 연결하면 분할 |
| 최대 요소 수 | **6개** | 인지 부하 연구: 6개 초과 시 이해도 급감 (Phillips) |
| 최대 불릿 수 | **4~5개** | Working memory 한계 (Miller's 7±2, 현대 연구 ~4) |
| 60초 규칙 | 암묵적 | 요소 수 제한으로 강제 |
| 장식 요소 | **0 (금지)** | "So What" 연결 안 되는 클립아트, 이미지 불허 |

### 3.3 Multi-Element 공간 규칙 (차트+KPI+표 공존 시)

```
규칙 1: 주인공이 60%를 차지한다
  → Action Title을 가장 직접 증명하는 요소가 Body 영역의 60%
  → 나머지 요소는 40%를 분할

규칙 2: KPI 스트립은 상단
  → KPI 카드 행은 Body의 상단 15~20%
  → 메인 차트/표는 하단 80%

규칙 3: 좌-맥락 우-상세 (Left-Right Split)
  → 맥락/요약: 좌측 30%
  → 상세 데이터: 우측 70%

규칙 4: 주석은 근접 배치
  → Callout은 참조 데이터 포인트 바로 옆
  → 떠다니는 주석 금지

규칙 5: 그리드 정렬 필수
  → 관련 요소는 공간적으로 그룹핑
  → 비관련 요소는 여백으로 분리
```

### 3.4 슬라이드 아키타입 11종 (Micro-Taxonomy)

| 아키타입 | 목적 | 전형적 조합 |
|---------|------|-----------|
| **Executive Summary** | SCR 단독 | bold-bullet 3~5개 (굵은 주장 + 들여쓴 근거) |
| **Exhibit Slide** | 단일 차트로 1개 포인트 증명 | Action Title + 차트 + 출처 + callout |
| **Waterfall/Bridge** | 가치 분해 | 폭포 차트 + 라벨링 + 시작/종료 합계 |
| **Recommendation Cascade** | 3~5개 실행 안 제시 | 번호 박스 + 영향도 지표 |
| **Scenario Comparison** | 옵션 비교 | 2~3열 + 일관된 행 기준 + 추천안 강조 |
| **Status Dashboard** | RAG 현황 | RAG 지표 + 성과/마일스톤/리스크/의사결정 |
| **Market Sizing** | TAM/SAM/SOM | 퍼널 또는 중첩 원 + 필터 |
| **Framework Slide** | 모델 적용 (Porter, BCG) | 구조 다이어그램 + 채워진 셀 |
| **Setup/Context** | 상황 설정 | 텍스트 중심 + 타임라인 또는 배경 데이터 |
| **Implication** | 데이터에서 결론 도출 | bold-bullet 종합 |
| **One-Pager** | 전체 논거 1장 | headline + 3~5포인트 + 시각 요소 + CTA |

→ **코드 적용**: 아키타입을 SEQUENCE 단계에서 태깅하고, 연속 동일 아키타입 금지

---

## 4. 조합 지능 (Composition Intelligence)

### 3.1 콘텐츠 → 조합 결정 프레임워크

Claude Code가 슬라이드를 구성할 때 따르는 4단계 판단:

```
Step 1: 메시지 유형 분류
  ┌─ 비교 (A vs B, 옵션) → 병렬 배치 레이아웃
  ├─ 프로세스 (순서, 단계) → 플로우 컴포넌트 + 설명 존
  ├─ 구조 (관계, 계층) → 다이어그램 컴포넌트 + 주변 존
  ├─ 성과 (KPI, 지표) → 데이터 컴포넌트 + 해설 존
  └─ 전략 (방향, 비전) → 앵커 컴포넌트 + 방향별 존

Step 2: 콘텐츠 블록 수 파악
  ┌─ 1블록 → full 레이아웃 + 대형 컴포넌트 1개
  ├─ 2블록 → two_column 또는 top_bottom
  ├─ 3블록 → t_layout 또는 l_layout
  └─ 4+블록 → grid 또는 center_peripheral

Step 3: 정보 계층 설정
  ┌─ 주인공 (가장 중요) → 가장 큰 존 (50%+)
  ├─ 조연 (보충 설명) → 중간 존 (25~40%)
  └─ 엑스트라 (맥락, 각주) → 작은 존 (15~25%) 또는 takeaway

Step 4: 컴포넌트 배정
  각 존에 카테고리가 다른 컴포넌트를 배치
  예: 주인공=Structural(chevron) + 조연=Data(kpi_row) + 엑스트라=Text(bullet)
```

### 3.2 콘텐츠 유형별 최적 조합 사전

| 콘텐츠 유형 | 레이아웃 | 주인공 컴포넌트 | 조연 컴포넌트 | 밀도 |
|-----------|---------|-------------|-------------|------|
| **전환 로드맵** | t_layout(0.25) | comp_chevron_flow | comp_bullet_list ×2 | 높음 |
| **성과 대시보드** | t_layout(0.30) | comp_kpi_row | comp_native_chart + comp_bullet_list | 높음 |
| **조직 구조** | center_peripheral_4 | comp_hub_spoke_diagram | comp_bullet_list ×4 | 중간 |
| **비교 분석** | two_column | comp_comparison_grid | comp_takeaway_bar | 높음 |
| **기술 아키텍처** | l_layout | comp_architecture_stack | comp_bullet_list + comp_icon_list | 높음 |
| **전략 방향** | center_peripheral_4 | comp_hero_block (중앙) | comp_bullet_list ×4 | 중간 |
| **프로세스 상세** | top_bottom(0.20) | comp_chevron_flow | comp_comparison_grid | 매우 높음 |
| **재무 분석** | two_column | comp_waterfall | comp_bullet_list + comp_kpi_card | 높음 |
| **프로젝트 일정** | full | comp_gantt_bars | comp_takeaway_bar | 중간 |
| **순환 프로세스** | center_peripheral_4 | comp_cycle_arrows | comp_icon_header_card ×4 | 중간 |
| **역량 수준** | two_column | comp_pyramid | comp_bullet_list (단계별 설명) | 중간 |
| **Cross-functional** | full | comp_swimlane | comp_takeaway_bar | 높음 |
| **Hero 숫자** | l_layout(0.35) | comp_kpi_card (대형) | comp_bullet_list + comp_native_chart | 중간 |

### 3.3 밀도 확보 규칙

**문제**: 컴포넌트를 자유 배치하면 "빈 존"이 생길 수 있다.

**해법**: 밀도 규칙을 코드와 가이드 양쪽에 적용.

```
규칙 1: 빈 존 금지
  → 모든 존에 최소 1개 컴포넌트 필수
  → evaluate.py에서 "존 대비 shape 커버리지" 체크

규칙 2: 시각 유형 다양성
  → 한 슬라이드에 최소 2개 카테고리 컴포넌트
  → text+text+text는 금지, text+chart+kpi는 권장

규칙 3: 공간 활용률 70%+
  → shape들의 면적 합 / 슬라이드 면적 ≥ 70%
  → 30% 이상 빈 공간이면 evaluate.py에서 경고

규칙 4: 텍스트 밀도 기준 유지
  → 슬라이드당 최소 300자 (차트 중심이면 완화)
  → 최대 2000자 (넘으면 분할 권장)

규칙 5: 폰트 위계 3단계
  → 제목(16-20pt) + 헤더(11-12pt) + 본문(8-10pt)
  → 위계가 2단계 이하면 경고
```

---

## 5. 덱 레벨 시각 리듬 (Step 2.5: SEQUENCE)

### 4.1 기존 워크플로우에 새 단계 삽입

```
기존: ANALYZE → PLAN → SELECT → GENERATE → EVALUATE → REFINE
신규: ANALYZE → PLAN → SEQUENCE → SELECT → GENERATE → EVALUATE → REFINE
                        ^^^^^^^^
```

**SEQUENCE 단계**가 하는 일:
- 전체 덱의 시각 리듬을 최적화
- 같은 레이아웃 연속 사용 방지
- 밀도 교차 (높음-중간-높음-낮음)
- 브레더 슬라이드 자동 삽입

### 4.2 60-25-15 규칙

10장 이상의 덱에서:
- **60% 콘텐츠** 슬라이드 (데이터, 프레임워크, 분석)
- **25% 전환/브레더** 슬라이드 (섹션 구분선, Hero KPI, 요약)
- **15% 앵커** 슬라이드 (표지, 결론, 부록 머리)

15장 덱이면: 콘텐츠 9장 + 전환 4장 + 앵커 2장

### 4.3 아키타입 태깅

모든 슬라이드에 아키타입을 태깅하고, 연속 동일 아키타입을 금지:

| 아키타입 | 설명 | 대표 레이아웃 |
|---------|------|------------|
| `grid` | 격자 배열 | grid_nxm, grid_2x2 |
| `split` | 좌우/상하 분할 | two_column, top_bottom, l_layout |
| `flow` | 흐름/순서 | chevron + bullets, timeline_band |
| `diagram` | 도형 중심 | center_peripheral, hub_spoke |
| `single_focus` | 단일 대형 요소 | full + hero_block, full + swimlane |
| `table` | 표 중심 | comparison_grid, dense_table |
| `kpi` | 숫자 중심 | kpi_row + details |

**제약**: 4장 연속 중 최소 3개 다른 아키타입 사용

### 4.4 밀도 교차 패턴

```
슬라이드:  1    2    3    4    5    6    7    8    9    10
밀도:     앵커  높음  중간  높음  브레더 높음  중간  높음  중간  앵커
아키타입: cover split flow grid  kpi   diag table split flow  end
```

---

## 6. 품질 검증 업그레이드 (evaluate.py)

### 5.1 기존 평가 항목 (유지)

- shape 수 (최소 5개)
- 텍스트 밀도 (300~2000자)
- 폰트 위계 (3단계+)
- accent 색상 비율 (25% 이하)
- 하단 빈 공간 (1.5" 이하)
- overflow 추정
- 시각 다양성 (카테고리 2종+)

### 5.2 새 평가 항목 (추가)

```python
# 10. 조합 다양성 — 컴포넌트 카테고리 다양성
# Structural, Data, Text, Anchor, Decorative 중 2종+ 필요
if component_categories < 2:
    issues.append("MEDIUM: 시각 유형이 단일 — 2종+ 권장")

# 11. 공간 활용률 — shape 면적 합 / 슬라이드 면적
total_shape_area = sum(w * h for each shape)
slide_area = 10 * 7.5  # inches
coverage = total_shape_area / slide_area
if coverage < 0.50:
    issues.append("HIGH: 공간 활용률 낮음 ({coverage:.0%})")

# 12. 덱 레벨 리듬 — 연속 동일 아키타입 감지
if consecutive_same_archetype >= 2:
    issues.append("MEDIUM: 연속 동일 아키타입 ({archetype})")

# 13. 덱 레벨 밀도 교차 — 연속 고밀도 감지
if consecutive_high_density >= 3:
    issues.append("MEDIUM: 고밀도 슬라이드 3장 연속")
```

---

## 7. slide_designer.md 전면 개편

### 6.1 현재 → 변경

| 항목 | 현재 | 변경 후 |
|------|------|--------|
| 매칭 테이블 | "내용 특성 → 패턴 이름" | "내용 특성 → 레이아웃 + 컴포넌트 조합" |
| 복합 구성 | JSON 스키마 기반 (sections) | Composer + 컴포넌트 직접 호출 코드 |
| 판단 프로세스 | 4단계 (패턴 선택 중심) | 4단계 (메시지→블록→계층→컴포넌트) |
| 밀도 규칙 | "빈 공간 30% 이하" | 5가지 밀도 규칙 (§3.3) |

### 6.2 새 판단 프로세스 (Claude Code용)

```
Step 1: 메시지 분류
  → "이 슬라이드는 무엇을 말하는가?" (비교/프로세스/구조/성과/전략)

Step 2: 블록 수 파악
  → "몇 개의 독립적 정보 블록이 있는가?" (1~4+)

Step 3: 레이아웃 선택
  → 블록 수 + 메시지 유형으로 15개 레이아웃 중 선택
  → 아키타입 제약 확인 (직전 슬라이드와 다른 아키타입)

Step 4: 컴포넌트 배정
  → 각 존에 적절한 컴포넌트 선택
  → 호환성 매트릭스 확인 (최소 2개 카테고리)
  → 밀도 규칙 확인

Step 5: 코드 생성
  → SlideComposer + comp_xxx 호출 코드 직접 작성
```

---

## 8. 구현 로드맵

### Phase 1: Compound Component 추출 (핵심 5개)

가장 사용 빈도 높은 5개부터:

| 순서 | 컴포넌트 | 원본 패턴 | 예상 | 즉시 가능해지는 조합 | 근거 |
|------|---------|---------|------|-----------------|------|
| 1 | `comp_chevron_flow` | timeline_phases 외 3개 | ~40줄 | Chevron+Bullets, Chevron+Grid, Chevron+KPI | **4개 패턴에서 중복** — 최고 ROI |
| 2 | `comp_hero_block` | executive_summary | ~60줄 | Hero+KPI, Hero+Chart, Hero+Bullets | 전략 슬라이드 필수 요소 |
| 3 | `comp_hub_spoke_diagram` | hub_spoke | ~70줄 | HubSpoke+Bullets (4~6방향) | 구조/관계 표현의 핵심 |
| 4 | `comp_comparison_grid` | comparison_matrix | ~80줄 | Comparison+Chart, Comparison+KPI | 비교 분석 필수 |
| 5 | `comp_architecture_stack` | architecture_stack | ~40줄 | Stack+Bullets, Stack+Chart | 가장 단순 구현, 즉시 완성 |

**Phase 1 완료 시**: 5개 compound (~290줄) + 기존 23+ atomic = **어떤 컨설팅 슬라이드든 조합 가능**

### Phase 2: 추가 Compound Component (5개)

| 순서 | 컴포넌트 | 원본 패턴 | 예상 |
|------|---------|---------|------|
| 6 | `comp_pyramid` | pyramid_layers | ~60줄 |
| 7 | `comp_cycle_arrows` | cycle_diagram | ~80줄 |
| 8 | `comp_waterfall` | waterfall_bridge | ~70줄 |
| 9 | `comp_swimlane` | swimlane | ~100줄 |
| 10 | `comp_gantt_bars` | gantt_roadmap | ~80줄 |

### Phase 3: 판단 시스템 업그레이드

| 순서 | 작업 | 상세 |
|------|------|------|
| 3-1 | slide_designer.md 전면 개편 | 새 판단 프로세스 + 조합 사전 |
| 3-2 | evaluate.py 업그레이드 | 조합 다양성 + 공간 활용률 + 덱 리듬 체크 |
| 3-3 | 07_WORKFLOW.md에 SEQUENCE 단계 추가 | Step 2.5 아키타입 태깅 + 리듬 최적화 |
| 3-4 | COMPOSITION_RECIPES 전면 교체 | 패턴 이름 대신 컴포넌트 조합으로 |

### Phase 4: 실증 + 검증

| 순서 | 작업 | 상세 |
|------|------|------|
| 4-1 | 넷제로 10장 덱 재생성 | 새 시스템으로 10장 전량 재구성 |
| 4-2 | PDF 시각 검증 | 10장 전수 시각 확인 |
| 4-3 | PwC 완성본 대비 검증 | 레퍼런스와 밀도/다양성 비교 |
| 4-4 | 새 주제 덱 생성 | 완전히 새로운 주제로 범용성 검증 |

---

## 9. 패턴 함수의 미래

### 즉시 삭제하지 않는다

패턴 함수 25개는 그대로 유지한다. 이유:
1. **하위 호환** — 기존 테스트와 출력물이 깨지지 않음
2. **단순 슬라이드용** — 조합이 필요 없는 단순한 1-element 슬라이드에는 여전히 유용
3. **참조 구현** — 새 컴포넌트 개발 시 좌표/스타일 참고

### 점진적 재정의

Phase 2 이후, 패턴 함수의 내부를 Composer+Component로 재작성:

```python
# patterns.py — 기존 API 유지, 내부만 변경
def executive_summary(slide, spec):
    """하위 호환 래퍼. 내부는 Composer+Component."""
    comp = SlideComposer(slide)
    comp.header(spec.header)
    zones = comp.layout("two_column", split=0.45)
    comp_hero_block(comp.canvas, ..., region=zones["left"])
    comp_kpi_row(comp.canvas, ..., region=zones["right"])
    comp.takeaway(spec.takeaway)
    comp.footer(spec.footer)
```

이렇게 하면:
- `executive_summary()` 호출은 그대로 작동
- 하지만 내부가 Region-aware → Claude Code가 직접 컴포넌트를 섞어 쓸 수도 있음

---

## 10. 리스크와 대응

| 리스크 | 확률 | 영향 | 대응 |
|-------|------|------|------|
| 컴포넌트가 존에 안 맞음 (크기 초과) | 중 | 높음 | Region 경계 체크 + 자동 축소 로직 |
| Claude Code 판단이 부정확 | 높음 | 중 | 조합 사전 + 레시피 + 시각 검증 루프 |
| 너무 많은 컴포넌트 → 선택 어려움 | 중 | 중 | 카테고리 분류 + 콘텐츠 유형별 추천 |
| 밀도가 높아지면서 가독성 저하 | 중 | 높음 | 폰트 최소 크기 제한 + overflow 체크 강화 |
| 기존 패턴 코드와 충돌 | 낮 | 낮 | 패턴은 건드리지 않고 신규 코드만 추가 |

---

## 11. 성공 기준

### 정량 기준
- [ ] 10장 덱에서 아키타입 4종+ 사용
- [ ] 슬라이드당 평균 컴포넌트 카테고리 2.5종+
- [ ] 공간 활용률 평균 65%+
- [ ] evaluate.py 점수 평균 80+
- [ ] 연속 동일 아키타입 0건

### 정성 기준
- [ ] PwC 완성본과 비교 시 밀도/다양성 동등 수준
- [ ] 같은 주제 10장이 "다 다른 느낌"
- [ ] 사용자가 "이제 컨설팅 품질 나온다" 평가

---

## 부록 A: Phase 1 컴포넌트 상세 설계

### A.1 comp_hero_block

```python
def comp_hero_block(
    c: Canvas,
    *,
    headline: str,           # 큰 메시지 (2~3줄)
    sub_points: list[str],   # 하위 불릿 (3~5개)
    bg_color: str = "grey_800",
    text_color: str = "white",
    region: Region,
) -> float:
    """대형 색상 박스 — 핵심 메시지 강조.
    
    전체 region을 bg_color로 채우고,
    상단 60%에 headline, 하단 40%에 sub_points.
    executive_summary의 좌측 Hero 영역에서 추출.
    
    사용 맥락:
      - 전략 방향 선언
      - 핵심 발견 강조
      - 섹션 도입부
    """
```

### A.2 comp_chevron_flow

```python
def comp_chevron_flow(
    c: Canvas,
    *,
    phases: list[dict],      # [{"label": "분석", "detail": "..."}, ...]
    style: str = "filled",   # "filled" | "outlined" | "gradient"
    show_numbers: bool = True,
    region: Region,
) -> float:
    """수평 쉐브론 화살표 체인.
    
    3~6개 단계를 수평 화살표 체인으로 표현.
    각 쉐브론 안에 번호+라벨, 아래에 detail.
    timeline_phases에서 추출.
    
    사용 맥락:
      - 프로세스 로드맵
      - 프로젝트 페이즈
      - 의사결정 흐름
    """
```

### A.3 comp_comparison_grid

```python
def comp_comparison_grid(
    c: Canvas,
    *,
    columns: list[dict],     # [{"header": "Option A", "items": [...], "highlight": False}, ...]
    row_labels: list[str],   # ["비용", "기간", "리스크"]
    region: Region,
) -> float:
    """N열 비교 표.
    
    컬럼별 헤더(색상 구분) + 행별 비교 데이터.
    highlight=True인 컬럼은 accent 배경.
    comparison_matrix에서 추출.
    
    사용 맥락:
      - 옵션 A/B/C 비교
      - 솔루션 벤더 비교
      - AS-IS vs TO-BE
    """
```

### A.4 comp_hub_spoke_diagram

```python
def comp_hub_spoke_diagram(
    c: Canvas,
    *,
    center: str,             # 중앙 텍스트
    spokes: list[dict],      # [{"label": "...", "detail": "..."}, ...]
    center_color: str = "accent",
    spoke_color: str = "grey_200",
    region: Region,
) -> float:
    """허브-스포크 방사형 다이어그램.
    
    중앙 원(hub) + N개 외곽 원(spoke) + 연결선.
    spoke 개수에 따라 자동 배치 (3~8개).
    hub_spoke에서 추출.
    
    사용 맥락:
      - 시스템 통합 구조
      - 핵심 역량 + 영향 영역
      - 이해관계자 맵
    """
```

### A.5 comp_architecture_stack

```python
def comp_architecture_stack(
    c: Canvas,
    *,
    layers: list[dict],      # [{"name": "Presentation", "items": [...]}, ...] (top→bottom)
    style: str = "gradient",  # "gradient" | "alternating" | "uniform"
    region: Region,
) -> float:
    """수직 레이어 스택 (기술 아키텍처).
    
    위에서 아래로 층별 박스, 각 층에 이름+항목.
    gradient면 위=연한색, 아래=진한색.
    architecture_stack에서 추출.
    
    사용 맥락:
      - 기술 스택
      - 시스템 레이어
      - 조직 계층
    """
```
