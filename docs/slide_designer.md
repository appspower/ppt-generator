# Slide Designer Guide — Claude Code용 화면 구성 판단 가이드

> 이 문서는 Claude Code가 PPT 슬라이드의 화면 구성을 판단할 때 참조합니다.
> 사용자가 "~에 대해 1장 만들어줘"라고 요청하면, 이 가이드에 따라 최적의 레이아웃을 선택합니다.

## 판단 프로세스

```
Step 1: 내용 분석 → 핵심 포인트 개수, 관점(비교/순서/분류/분석) 파악
Step 2: 이 가이드의 매칭 테이블에서 최적 레이아웃 선택
Step 3: 복합 구성이 필요한지 판단 (내용이 2개 이상 관점을 포함하면)
Step 4: JSON 스키마 생성
```

## 레이아웃 선택 매칭 테이블

### 단일 관점

| 내용 특성 | 포인트 수 | 추천 레이아웃 | 예시 |
|---|---|---|---|
| **비교** (A vs B) | 2 | `comparison` 또는 `before_after` | "Big Bang vs Hybrid" |
| **분류/영역별** | 2~3개 | `columns` n_cols=2~3 + `card` | "3대 핵심 모듈" |
| **분류/영역별** | 4개 | `columns` n_cols=4 또는 `4col_reference` 템플릿 | "4대 부문 현황" |
| **순서/단계** | 3~6단계 | `chevron_process` | "4단계 로드맵" |
| **항목 나열** | 3~7개 | `numbered_circle` | "5대 추진 과제" |
| **다차원 분석** | N×M | `framework_matrix` | "5개 병목 × 4차원" |
| **강점약점** | 4사분면 | `swot` 템플릿 | "SWOT 분석" |
| **성과/지표** | 3~4개 KPI | `kpi_dashboard` 템플릿 | "4대 KPI" |
| **의사결정** | 3~5단계 | `decision_flow` 템플릿 | "도입 판단 흐름" |
| **아키텍처** | 중심+주변 | `hub_spoke` 템플릿 | "시스템 통합 구조" |
| **시간축** | 3~6시점 | `timeline` 템플릿 | "2024~2028 로드맵" |
| **변화** | 현재→미래 | `before_after` 템플릿 | "AS-IS vs TO-BE" |
| **작업 분해** | 행×열 그리드 | `task_image_activity` 템플릿 | "Task/Description/Activity" |
| **증감 분석** | 단계별 변화 | `waterfall` 템플릿 | "원가 Bridge 분석" |
| **중심+확장** | 1 중심+4 주변 | `center_focus` 템플릿 | "핵심 역량 + 4대 영향" |
| **밀도 높은 데이터** | 8×5+ 그리드 | `dense_table` 템플릿 | "상세 비교 테이블" |
| **분석+시각화** | 좌:텍스트 우:도형 | `two_panel` 템플릿 | "이슈 + 솔루션 상세" |
| **역할/책임** | RACI 그리드 | `raci` 템플릿 | "프로젝트 RACI" |
| **부서간 프로세스** | 레인×단계 | `swimlane` 템플릿 | "Cross-functional Flow" |
| **거시환경** | 6셀 분석 | `pestel` 템플릿 | "PESTEL 분석" |
| **경영진 보고** | S-C-R 3섹션 | `scr` 템플릿 | "이슈 프레이밍" |
| **맥락+상세** | 좌30%+우70% | `left_right_split` 템플릿 | "Key Point + Detail" |
| **산업 분석** | 5 Forces | `porter_five_forces` 템플릿 | "Porter 5 Forces" |
| **가치사슬** | 쉐브론+지원 | `value_chain` 템플릿 | "Value Chain 분석" |
| **포트폴리오** | 2×2 버블 | `bcg_matrix` 템플릿 | "BCG 성장-점유 매트릭스" |
| **조직도** | 3-Level 트리 | `org_chart` 템플릿 | "조직 구조" |
| **프로젝트 일정** | 트랙×시간 | `gantt_roadmap` 템플릿 | "Gantt 로드맵" |
| **우선순위** | Impact×Effort | `prioritization_2x2` 템플릿 | "Quick Win 매트릭스" |
| **민감도** | 좌우 대칭 바 | `tornado` 템플릿 | "Sensitivity 분석" |
| **분기 분석** | 트리 구조 | `decision_tree` 템플릿 | "의사결정 트리" |
| **매출 분해** | 동인 트리 | `revenue_tree` 템플릿 | "Revenue Decomposition" |
| **3안 비교** | 3열+추천 | `three_option` 템플릿 | "Option A/B/C 비교" |
| **순환 프로세스** | 4-node cycle | `circular_loop` 템플릿 | "PDCA 사이클" |
| **조직 정렬** | 7-node web | `mckinsey_7s` 템플릿 | "McKinsey 7S" |
| **혁신 포트폴리오** | 3개 S-curve | `three_horizons` 템플릿 | "3-Horizon Growth" |
| **시장 구조** | 가변폭 스택 | `mekko` 템플릿 | "Marimekko 분석" |
| **진척 현황** | 표+인라인 바 | `table_with_bars` 템플릿 | "Progress Dashboard" |

### 복합 구성 (핵심!)

| 내용 특성 | 추천 복합 구성 | sections 구성 |
|---|---|---|
| **단계 + 영역별 상세** | 상단: chevron_process + 하단: columns cards | `[{area: "top", h: 0.25}, {area: "bottom", h: 0.75}]` |
| **프레임워크 + 시사점** | 상단: framework_matrix + 하단: takeaway_bar | `[{area: "main", h: 0.85}, {area: "footer", h: 0.15}]` |
| **비교 + 근거** | 상단: comparison cards + 하단: bullet 근거 | `[{area: "top", h: 0.6}, {area: "bottom", h: 0.4}]` |
| **KPI + 상세 분석** | 좌: KPI 카드 + 우: bullet 상세 | `n_cols: 2, 좌=card(KPI), 우=bullet` |
| **프로세스 + 매트릭스** | 상단: chevron + 하단: matrix | `[{area: "top", h: 0.2}, {area: "bottom", h: 0.8}]` |
| **개요 + 4영역 상세** | 상단: 요약 텍스트 + 하단: 4컬럼 카드 | `[{area: "top", h: 0.15}, {area: "bottom", h: 0.85}]` |
| **타임라인 + 상세** | 상단: timeline + 하단: 표 또는 불릿 | `[{area: "top", h: 0.4}, {area: "bottom", h: 0.6}]` |

## 복합 구성 JSON 구조

```json
{
  "type": "content",
  "title": "...",
  "layout": "stacked",
  "sections": [
    {
      "height_ratio": 0.3,
      "layout": "full",
      "elements": [
        {"type": "chevron_process", "steps": [...]}
      ]
    },
    {
      "height_ratio": 0.7,
      "layout": "columns",
      "n_cols": 3,
      "elements": [
        {"type": "card", ...},
        {"type": "card", ...},
        {"type": "card", ...}
      ]
    }
  ]
}
```

## 판단 규칙

1. **내용이 단일 관점** → 매칭 테이블에서 1개 선택
2. **내용이 2개 관점 포함** → 복합 구성 (stacked 레이아웃)
3. **비교가 핵심** → comparison 계열 우선
4. **숫자/KPI가 핵심** → card + KPI 또는 kpi_dashboard
5. **화면이 단순하면** → 텍스트를 늘려서 밀도를 높임 (빈 공간 금지)
6. **항목 5개 이상** → numbered_circle 또는 framework_matrix (불릿 나열 금지)
7. **TakeawayBar는 거의 항상 포함** → 핵심 시사점 1문장

## 폰트/밀도 규칙

- 4:3 슬라이드 기준, 빈 공간이 30% 이상이면 구성 재검토
- 카드 내부는 불릿 4~6개가 적정 (3개 이하면 너무 비어 보임)
- KPI 숫자는 24pt, 본문은 9~10pt, 각주 7pt
- 모든 텍스트는 한글 기준으로 밀도 계산 (한글은 영문 대비 1.5배 폭)
