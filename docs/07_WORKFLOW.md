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
  4. 복합 구성 필요성 판단 (stacked?)
  5. 프레임 선택 (standard/fullscreen/sidebar)

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

### Step 3: SELECT (템플릿/컴포넌트 선택)

```
입력: slide_plan + template_metadata.json

수행:
  1. content_type으로 후보 필터링 (보통 3~8개)
  2. item_count, density로 추가 필터 (2~4개)
  3. best_for / avoid_when 읽고 최종 선택
  4. 복합 구성이면 sections별로 각각 선택

규칙:
  - 같은 템플릿 연속 2회 사용 금지
  - 직전에 쓴 프레임과 다른 프레임 선택
  - 데이터가 있으면 시각화 컴포넌트 우선
  - 항목 5개 이상이면 매트릭스 or numbered_circle

출력: template_selection
  [{
    "slide_num": 1,
    "template_id": "framework_matrix" (또는 "stacked"),
    "sections": [...],  // stacked인 경우
    "rationale": "5개 항목 × 4차원 → 매트릭스 최적"
  }]
```

### Step 4: GENERATE (JSON 생성 + 렌더링)

```
입력: slide_plan + template_selection

수행:
  1. JSON 스키마 작성 (Pydantic 검증)
  2. python run.py 실행
  3. .pptx 생성

규칙:
  - Assertion Title: 핵심 인사이트를 문장으로 (라벨 금지)
  - 불릿: 4~6개 적정 (3개 이하면 내용 보충)
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

## 빠른 참조: content_type → 추천 템플릿

| content_type | 1차 추천 | 2차 추천 |
|---|---|---|
| comparison_2 | before_after, comparison | two_panel |
| comparison_3 | three_option | left_right_split |
| comparison_multi | framework_matrix | harvey_ball_matrix |
| process_linear | chevron_process | decision_flow |
| process_cycle | circular_loop | center_focus |
| process_vertical | vertical_flow | swimlane |
| classification_2_3 | columns + card | left_right_split |
| classification_4 | 4col_reference, numbered_quadrant | process_grid |
| classification_5_plus | numbered_circle | framework_matrix |
| data_kpi | kpi_dashboard | columns + card(KPI) |
| data_table | dense_table, table_with_bars | framework_matrix |
| data_trend | waterfall, mekko | three_horizons |
| data_sensitivity | tornado | harvey_ball_matrix |
| strategy_framework | porter_five_forces, bcg_matrix | value_chain |
| org_governance | org_chart, raci | swimlane |
| environment_scan | pestel, swot | prioritization_2x2 |
| exec_summary | scr | stacked(chevron + cards) |
| timeline | gantt_roadmap, timeline | left_right_split |
| architecture | hub_spoke, center_focus | circular_loop |
| status_tracking | rag_table, table_with_bars | kpi_dashboard |
| decision | decision_tree, decision_flow | comparison |
| financial | waterfall, revenue_tree | value_chain |
