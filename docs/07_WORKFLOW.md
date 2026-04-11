# 07. PPT 생성 워크플로우 — 6단계 에이전트 프로세스

> Claude Code가 PPT를 생성할 때 반드시 따라야 하는 체계적 프로세스.
> 업계 성공 사례(PPTAgent, McKinsey Lilli, auxi) 기반 설계.

## 핵심 원칙

1. **빈 캔버스에서 생성하지 않는다** — 반드시 기존 템플릿/컴포넌트를 편집
2. **선택은 체계적으로** — 감이 아닌 메타데이터 기반 매칭
3. **만들고 끝이 아니다** — 생성 → 평가 → 수정 루프 필수

---

## 6단계 프로세스

### Step 1: ANALYZE (내용 분석)

```
입력: 사용자 요청 ("~에 대해 PPT 만들어줘")

수행:
  1. 주제 파악
  2. 웹 리서치 (필요 시)
  3. 핵심 메시지 도출 (3~5개)
  4. 데이터/수치 수집
  5. 관점 분류:
     - 비교 관점? (A vs B)
     - 순서 관점? (1→2→3)
     - 분류 관점? (영역별)
     - 분석 관점? (다차원)
     - 혼합? (2개 이상 관점)

출력: content_inventory
  {
    "topic": "...",
    "key_messages": ["...", "..."],
    "data_points": [...],
    "perspectives": ["comparison", "process"],
    "complexity": "high",
    "target_slides": 1
  }
```

### Step 2: PLAN (슬라이드 계획)

```
입력: content_inventory

수행:
  1. 슬라이드별 목적 정의
  2. 각 슬라이드의 content_type 분류
  3. 핵심 메시지 배분
  4. 복합 구성 필요성 판단 (블록 수에 따라 레이아웃 결정)
  5. 프레임 선택 (레이아웃 15개 중)
  
  ★ 덱 리듬 규칙 (필수):
  6. 같은 레이아웃 연속 사용 금지 — 직전 슬라이드와 다른 레이아웃
  7. 고밀도 슬라이드(3+ 컴포넌트) 3장 연속 금지 — 중간에 Hero/단순 삽입
  8. 10장 이상 덱에서 레이아웃 4종+ 사용

분류 기준:
  - item_count: 콘텐츠 항목 수
  - has_data: 숫자/차트 데이터 유무
  - comparison_type: 없음/2안/3안/다차원
  - process_type: 없음/순차/병렬/순환
  - density: low/medium/high

출력: slide_plan
  [{
    "slide_num": 1,
    "purpose": "현황과 문제점을 한눈에 보여준다",
    "content_type": "comparison",
    "item_count": 5,
    "has_data": true,
    "key_message": "...",
    "frame": "stacked"
  }]
```

### Step 3: SELECT (레이아웃 + 컴포넌트 선택)

```
입력: slide_plan + docs/slide_designer.md

수행:
  1. 각 슬라이드의 메시지 유형 분류 (비교/프로세스/구조/성과/전략)
  2. slide_designer.md §3 매칭 테이블에서 레이아웃 + 컴포넌트 조합 선택
  3. PLAN의 덱 리듬 규칙 확인 (연속 동일 레이아웃 아닌지)
  4. slide_designer.md §5 SELF-CHECK 실행

규칙:
  - 같은 레이아웃 연속 2회 사용 금지
  - 컴포넌트 카테고리 2종+ 사용 (text+text 금지)
  - 데이터가 있으면 시각화 컴포넌트 우선
  - 항목 5개 이상이면 grid_nxm 또는 comp_comparison_grid

출력: composition_selection
  [{
    "slide_num": 1,
    "layout": "l_layout",
    "components": {
      "left_full": "comp_hero_block",
      "right_top": "comp_kpi_row",
      "right_bottom": "comp_bullet_list"
    },
    "rationale": "성과 메시지 + Hero 숫자 + 증거 → l_layout"
  }]
```

### Step 4: GENERATE (SlideComposer 코드 생성 + 렌더링)

```
입력: slide_plan + composition_selection

수행:
  1. 각 슬라이드별 SlideComposer 초기화
  2. composer.layout() 호출로 zones 획득
  3. 각 zone에 comp_xxx() 호출로 컴포넌트 배치
  4. composer.takeaway() + composer.footer() 추가
  5. .pptx 생성

규칙:
  - Assertion Title: 핵심 인사이트를 문장으로 (라벨 금지)
  - 불릿: 3~5개 적정 (2개 이하면 내용 보충)
  - KPI: 반드시 비교 기준 포함 ("40%↓" + "전년 대비")
  - TakeawayBar: 거의 항상 포함
  - 출처(footnote): 반드시 포함
```

### Step 5: EVALUATE (자체 평가)

```
입력: 생성된 .pptx

수행 (python 스크립트):
  1. shape 수 확인
  2. 텍스트 밀도 계산 (chars/inch)
  3. 폰트 위계 확인 (2단계 이상 차이?)
  4. 카드 높이 일관성
  5. 빈 공간 비율
  6. 색상 사용률 (accent 10% 이하?)

체크리스트:
  □ 제목이 인사이트 문장인가? (라벨 아닌가?)
  □ 빈 공간이 30% 이하인가?
  □ 폰트 크기가 3단계 이상 위계를 보이는가?
  □ 오렌지 배경이 전체의 10% 이하인가?
  □ TakeawayBar가 있는가?
  □ 출처가 있는가?
  □ 텍스트가 잘리지 않았는가?

출력: evaluation_report
  {
    "score": 72,
    "issues": ["빈 공간 35%", "accent 과다 사용"],
    "pass": false
  }
```

### Step 6: REFINE (수정)

```
입력: evaluation_report

수행:
  - score >= 80: 통과, 사용자에게 전달
  - score < 80: 이슈별 수정 후 Step 4로 돌아감
  - 최대 2회 반복 (3회째는 강제 통과 + 이슈 목록 첨부)

수정 우선순위:
  1. 텍스트 오버플로 → 텍스트 축약
  2. 빈 공간 과다 → 콘텐츠 추가 또는 레이아웃 변경
  3. accent 과다 → 색상 절제 적용
  4. 제목 라벨형 → Assertion Title로 재작성
```

---

## 빠른 참조: content_type → 레이아웃 + 컴포넌트

| content_type | 레이아웃 | 주인공 컴포넌트 |
|---|---|---|
| comparison_2 | `full` | `comp_before_after` |
| comparison_3 | `full` | `comp_comparison_grid` (highlight) |
| comparison_multi | `full` | `comp_quadrant_matrix` |
| process_linear | `full` | `comp_chevron_flow` (show_details) |
| process_cycle | `full` | `comp_cycle_arrows` |
| process_grid | `grid_nxm` | `comp_numbered_cell` ×N |
| data_kpi | `t_layout` | `comp_kpi_row`(상) + `comp_native_chart`(하) |
| data_hero | `l_layout` | `comp_kpi_card`(좌) + `comp_bullet_list`(우) |
| data_trend | `full` | `comp_waterfall` |
| data_risk | `two_column` | `comp_heatmap_grid`(좌) + `comp_bullet_list`(우) |
| data_market | `full` | `comp_funnel` |
| strategy_value | `full` | `comp_value_chain` |
| strategy_hero | `l_layout` | `comp_hero_block`(좌) + `comp_kpi_row`(우) |
| structure_hub | `center_peripheral_4` | `comp_hub_spoke_diagram`(중앙) |
| structure_stack | `two_column` | `comp_architecture_stack`(좌) + `comp_bullet_list`(우) |
| structure_tree | `full` | `comp_logic_tree` |
| structure_pyramid | `two_column` | `comp_pyramid`(좌) + `comp_bullet_list`(우) |
| timeline | `full` | `comp_gantt_bars` |
| exec_summary | `l_layout` | `comp_hero_block`(좌) + `comp_kpi_row` + `comp_chevron_flow`(우) |
| status_tracking | `t_layout` | `comp_kpi_row`(상) + `comp_heatmap_grid`(하) |

> 상세 매칭 및 복합 구성 예시는 `docs/slide_designer.md` 참조
