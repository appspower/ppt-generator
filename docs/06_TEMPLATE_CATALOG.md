# 06. Template & Component Catalog — 초정밀 템플릿/컴포넌트 마스터 리스트

> 완성본 94장 정밀 분석 + 컨설팅 업계 33개 프레임워크 리서치 기반.
> 각 항목은 구현 우선순위, 난이도, 세부 스펙을 포함하여 바로 제작에 착수할 수 있는 수준.

---

## Part 1: 완성본에서 발굴한 템플릿 패턴 (16종)

### 빈도 TOP 5 (최우선 구현)

| ID | 패턴명 | 빈도 | 난이도 | 현재 보유 | 상태 |
|---|---|---|---|---|---|
| **P5** | Left-Right Split (좌:맥락 + 우:흐름/상세) | 11회 | Medium | sidebar 유사 | 🔶 개선필요 |
| **P4** | Horizontal Linear Process (번호+화살표+상세) | 10회 | Easy-Med | chevron_process | 🔶 개선필요 |
| **P2** | Horizontal Layered Framework (행열 하이라이트) | 9회 | Medium | framework_matrix | ✅ 보유 |
| **P1** | Numbered Quadrant Info-Block (2×2 번호 박스) | 8회 | Easy | ❌ 없음 | 🔴 신규필요 |
| **P8** | Two-Column Before/After | 8회 | Easy-Med | before_after | ✅ 보유 |

### 빈도 6~7회 (2차 구현)

| ID | 패턴명 | 빈도 | 난이도 | 현재 보유 | 상태 |
|---|---|---|---|---|---|
| **P3** | Multi-Column Comparison Table (5열+) | 7회 | Medium | dense_table 유사 | 🔶 개선필요 |
| **P7** | Vertical Step-Down Process | 6회 | Medium | ❌ 없음 | 🔴 신규필요 |
| **P10** | Dense Table + Inline Charts | 6회 | Hard | dense_table | 🔶 개선필요 |
| **P12** | Central Hub + Spokes | 6회 | Hard | hub_spoke | ✅ 보유 |

### 빈도 3~5회 (3차 구현)

| ID | 패턴명 | 빈도 | 난이도 | 현재 보유 |
|---|---|---|---|---|
| **P6** | Timeline / Gantt Roadmap | 5회 | Hard | timeline |
| **P15** | Multi-Zone Dashboard | 5회 | Hard | ❌ 없음 |
| **P9** | Three-Column Option Compare | 4회 | Medium | comparison 유사 |
| **P11** | Summary Stats + Table | 4회 | Easy | kpi_dashboard 유사 |
| **P16** | Section Divider / Chapter | 4회 | Easy | section_divider |
| **P13** | Circular / Loop Diagram | 3회 | Hard | ❌ 없음 |
| **P14** | Pyramid Hierarchy | 3회 | Medium | ❌ 없음 (build_library에만) |

---

## Part 2: 컨설팅 업계 프레임워크 (33종) — 현재 미보유 패턴 중심

### A. 전략 프레임워크 (6종)

#### A1. Porter's Five Forces
```
용도: 산업 매력도, 경쟁 역학 분석
레이아웃: 중앙 박스(경쟁 강도) + 4방향 위성 박스
구성요소:
  - 중앙: "Industry Rivalry" 박스
  - 상: Threat of New Entrants
  - 하: Threat of Substitutes
  - 좌: Bargaining Power of Suppliers
  - 우: Bargaining Power of Buyers
  - 각 박스: intensity 표시 (High/Med/Low 또는 Harvey Ball)
  - 쌍방향 화살표 연결
구현 난이도: Medium (hub_spoke 변형)
```

#### A2. Value Chain (가치사슬)
```
용도: 원가/마진 분석, 경쟁우위 원천 식별
레이아웃: 
  상단 화살표 밴드: Inbound→Operations→Outbound→Marketing→Service
  하단 지원활동: Infrastructure, HR, Technology, Procurement
  우측: Margin 쐐기
구현 난이도: Medium (chevron + 하단 행)
```

#### A3. BCG Growth-Share Matrix
```
용도: 포트폴리오 우선순위, 자원 배분
레이아웃: 2×2 그리드 + 버블 차트 오버레이
  X축: 상대적 시장점유율 (좌=높음)
  Y축: 시장 성장률 (상=높음)
  사분면: Stars / Cash Cows / Question Marks / Dogs
  버블: 사업단위 (크기=매출)
구현 난이도: Medium (SWOT 변형 + 원형 오버레이)
```

#### A4. McKinsey 7S
```
용도: 조직 정렬, 변화관리
레이아웃: 7개 노드의 웹/스파이더 다이어그램
  Hard S: Strategy, Structure, Systems
  Soft S: Style, Staff, Skills
  중심: Shared Values
구현 난이도: Hard (center_focus 확장 필요, 7노드)
```

#### A5. PESTEL
```
용도: 거시환경 스캐닝
레이아웃: 6셀 그리드 (2열×3행) 또는 6-spoke
  P/E/S/T/E/L 각 셀에 불릿 포인트
구현 난이도: Easy (process_grid 변형)
```

#### A6. Three Horizons (3-Horizon Growth)
```
용도: 혁신 포트폴리오, 성장 전략 단계화
레이아웃: 3개 S-curve on 시간/수익 축
  H1: 현재 핵심사업 (단기)
  H2: 신규 성장사업 (중기)
  H3: 미래 탐색사업 (장기)
구현 난이도: Hard (곡선 shape 필요)
```

### B. 프로세스/흐름 (5종)

#### B1. Swimlane Process Map
```
용도: 부서 간 프로세스, 핸드오프 매핑
레이아웃: 수평 레인 (행=역할), 프로세스 좌→우
구성요소: 레인 헤더, 프로세스 박스, 의사결정 다이아몬드, 화살표
구현 난이도: Medium
현재: 미보유 → 신규 필요
```

#### B2. Decision Tree
```
용도: 시나리오 분석, Go/No-Go
레이아웃: 좌→우 분기 트리
  □ 의사결정 노드 → ○ 확률 노드 → △ 결과
구현 난이도: Hard (분기 연결선)
```

#### B3. Funnel (깔때기)
```
용도: 영업 파이프라인, 전환율 분석
레이아웃: 상=넓음 → 하=좁음 (테이퍼 형태)
  각 밴드: 단계명 + 수량 + 전환율%
구현 난이도: Medium (사다리꼴 shape)
```

#### B4. MECE Issue Tree
```
용도: 문제 분해, 가설 구조화
레이아웃: 좌→우 계층 트리
  Root = 문제 → L1 분기 → L2 → Leaf = 데이터/가설
구현 난이도: Medium-Hard
```

#### B5. Vertical Step-Down (P7)
```
용도: 순차 프로세스, 단계별 상세
레이아웃: 좌=라벨열, 중앙=수직 화살표 체인, 우=상세 텍스트
구현 난이도: Medium
현재: 미보유 → 신규 필요 ★
```

### C. 데이터 시각화 (7종)

#### C1. Waterfall / Bridge
```
용도: P&L Bridge, EBITDA 워크, 예산 vs 실적
이미 보유: waterfall 템플릿 (Slide 11)
개선점: 양수=그린, 음수=레드 색상 규칙 추가
```

#### C2. Harvey Ball Matrix
```
용도: 다기준 정성 평가, 벤더 비교
레이아웃: 행=옵션, 열=기준, 셀=Harvey Ball (0/25/50/75/100%)
구성요소: 원형 + 채움 비율, 범례, 가중 점수열
구현 난이도: Medium (원형 shape + 부분 채움)
현재: 미보유 → 신규 필요 ★★
```

#### C3. RAG Status Dashboard
```
용도: 프로젝트 건강성, KPI 모니터링
레이아웃: 표 or 아이콘 그리드, R/A/G 원형
구현 난이도: Easy (표 + 색상 원형)
현재: kpi_dashboard 유사하나 RAG 전용 미보유
```

#### C4. Tornado Chart (민감도)
```
용도: 핵심 가치 동인 식별, 리스크 우선순위
레이아웃: 수평 막대 좌우 대칭, 가장 넓은 것 상단
구현 난이도: Medium
현재: 미보유
```

#### C5. Mekko / Marimekko
```
용도: 시장 규모 + 점유율 동시 표현
레이아웃: 가변 폭 스택드 바
구현 난이도: Hard (가변 폭 계산)
```

#### C6. Stacked Bar (100% / Absolute)
```
용도: 구성 비율 추이, 시장 점유율 트렌드
구현 난이도: Medium (python-pptx 네이티브 차트)
```

#### C7. Scatter / Bubble Chart
```
용도: 상관관계, 포트폴리오 포지셔닝
구현 난이도: Medium (python-pptx 네이티브)
```

### D. 비교 패턴 (4종)

#### D1. 2×2 Prioritization Matrix
```
용도: 이니셔티브 우선순위, 리스크 vs 리턴
레이아웃: 2×2 + 버블 오버레이 (BCG 변형)
  축: Impact vs Effort
  사분면: Quick Wins / Big Bets / Fill-ins / Hard No
구현 난이도: Medium (SWOT + 버블)
현재: swot 변형으로 가능
```

#### D2. Feature Comparison (Harvey Ball)
```
→ C2 Harvey Ball Matrix와 동일
```

#### D3. Scenario Planning Matrix
```
용도: 전략적 불확실성, 시나리오 플래닝
레이아웃: 2×2, 축=2개 핵심 불확실성
  각 사분면: 시나리오명 + 전략 시사점
구현 난이도: Easy (SWOT 변형)
```

#### D4. Three-Column Option (P9)
```
이미 comparison 템플릿으로 유사 보유
개선점: 3열 지원 + 추천 하이라이트
```

### E. 조직 패턴 (3종)

#### E1. RACI Matrix
```
용도: 역할 명확화, 프로젝트 거버넌스
레이아웃: 행=작업, 열=역할, 셀=R/A/C/I
  색상코딩: R=다크, A=중간, C=연한, I=아웃라인
구현 난이도: Easy (framework_matrix 변형)
```

#### E2. Org Chart
```
용도: 보고체계, 팀 설계
레이아웃: 계층 트리 (CEO→직보→...)
구현 난이도: Medium-Hard (계층 연결선)
현재: 미보유 → 신규 필요 ★
```

#### E3. Governance Model
```
용도: 의사결정 권한, 위원회 구조
레이아웃: 3-Tier 박스 (Board→ExCo→WG) + 회의 주기
구현 난이도: Medium (pyramid 변형)
```

### F. 재무 패턴 (3종)

#### F1. P&L Bridge
```
→ C1 Waterfall과 동일 구조, Revenue→EBITDA 특화
```

#### F2. Cost Stack
```
용도: 원가 구조 분해, 벤치마킹
레이아웃: 단일 스택드 바 or 워터폴
구현 난이도: Medium
```

#### F3. Revenue Decomposition Tree
```
용도: 매출 동인 분석, 탑라인 성장 귀인
레이아웃: 좌=총매출 → 우측 분기 (가격×물량→지역/제품/채널)
구현 난이도: Medium-Hard (트리 구조)
```

### G. 커뮤니케이션 구조 (3종)

#### G1. SCR (Situation-Complication-Resolution)
```
용도: 경영진 보고, 이슈 프레이밍
레이아웃: 3섹션 수평 분할 or 3장 연속
  S: 현황 (1~2줄)
  C: 복잡화 ("무엇이 변했나")
  R: 해법 (60~70% 비중)
구현 난이도: Easy (stacked 3-section)
```

#### G2. Pyramid Principle Stack
```
용도: 덱 전체 논리 체크
구현: 슬라이드가 아닌 덱 구조 가이드
```

#### G3. MECE Tree
```
→ B4와 동일
```

---

## Part 3: 구현 우선순위 매트릭스

### Tier 1 — 즉시 구현 (빈도 높음 + 난이도 낮음~중간)

| # | 패턴 | 타입 | 근거 |
|---|---|---|---|
| 1 | **P1: Numbered Quadrant (2×2 번호 박스)** | 신규 템플릿 | 완성본 8회, Easy, 현재 미보유 |
| 2 | **P7: Vertical Step-Down** | 신규 템플릿 | 완성본 6회, Medium, 미보유 |
| 3 | **C2: Harvey Ball Matrix** | 신규 컴포넌트+템플릿 | 업계 표준, 벤더 비교 필수, 미보유 |
| 4 | **E1: RACI Matrix** | 신규 템플릿 | 프로젝트 거버넌스 필수, matrix 변형 |
| 5 | **B1: Swimlane** | 신규 템플릿 | 완성본 6회(P7 변형), 프로세스 필수 |
| 6 | **A5: PESTEL** | 신규 템플릿 | 전략 기본, process_grid 변형 |
| 7 | **C3: RAG Dashboard** | 신규 템플릿 | PMO 필수, kpi 변형 |
| 8 | **P5 개선: Left-Right Split** | 기존 개선 | 완성본 최빈(11회), sidebar 고도화 |
| 9 | **B3: Funnel** | 신규 템플릿 | 영업 파이프라인 필수 |
| 10 | **G1: SCR** | 신규 템플릿 | 경영진 보고 필수, stacked 변형 |

### Tier 2 — 중기 구현 (빈도 중간 or 난이도 높음)

| # | 패턴 | 타입 | 근거 |
|---|---|---|---|
| 11 | P15: Multi-Zone Dashboard | 신규 | 완성본 5회, Hard |
| 12 | A1: Porter's Five Forces | 신규 | 전략 프레임워크 표준 |
| 13 | A2: Value Chain | 신규 | 원가/프로세스 분석 표준 |
| 14 | A3: BCG Matrix | 신규 | 포트폴리오 분석 표준 |
| 15 | E2: Org Chart | 신규 | 조직 설계 필수 |
| 16 | P6 개선: Gantt Roadmap | 기존 개선 | timeline 고도화 |
| 17 | D1: 2×2 Prioritization | 신규 | SWOT 변형 + 버블 |
| 18 | C4: Tornado Chart | 신규 | 민감도 분석 |
| 19 | B2: Decision Tree | 신규 | 시나리오 분석 |
| 20 | F3: Revenue Decomposition | 신규 | 재무 분석 |

### Tier 3 — 장기/선택 (난이도 높음 or 빈도 낮음)

| # | 패턴 | 근거 |
|---|---|---|
| 21 | P13: Circular/Loop Diagram | Hard, 3회 |
| 22 | A4: McKinsey 7S | Hard (7노드 웹) |
| 23 | A6: Three Horizons | Hard (S-curve) |
| 24 | C5: Mekko/Marimekko | Hard (가변 폭) |
| 25 | P10 개선: Inline Charts in Table | Hard |

---

## Part 4: 신규 컴포넌트 필요 목록

### 현재 보유 컴포넌트
card, text_block, badge, kicker, bullet, table, chart, process_flow,
chevron_process, framework_matrix, numbered_circle, takeaway_bar, divider, image

### 신규 필요 컴포넌트

| 컴포넌트 | 용도 | Tier |
|---|---|---|
| **harvey_ball** | 0/25/50/75/100% 채움 원형 | Tier 1 |
| **rag_indicator** | R/A/G 색상 원형 + 라벨 | Tier 1 |
| **funnel_step** | 깔때기 단계 (사다리꼴) | Tier 1 |
| **numbered_quadrant** | 2×2 번호 정보 블록 | Tier 1 |
| **vertical_flow** | 수직 화살표 체인 | Tier 1 |
| **connector_arrow** | 두 shape 간 연결 화살표 | Tier 2 |
| **org_box** | 조직도 노드 (이름/직책) | Tier 2 |
| **gantt_bar** | 간트 차트 수평 바 + 마일스톤 | Tier 2 |
| **bubble_plot** | 크기 가변 원형 (BCG/Portfolio) | Tier 2 |
| **sparkline** | 미니 차트 (셀 내장) | Tier 3 |

---

## Part 5: 세부 스펙 — Tier 1 신규 항목 상세

### 5-1. Numbered Quadrant (P1)
```
구조:
  ┌──────────────────┬──────────────────┐
  │ [01] 제목        │ [02] 제목        │
  │ • 불릿1         │ • 불릿1         │
  │ • 불릿2         │ • 불릿2         │
  ├──────────────────┼──────────────────┤
  │ [03] 제목        │ [04] 제목        │
  │ • 불릿1         │ • 불릿1         │
  │ • 불릿2         │ • 불릿2         │
  └──────────────────┴──────────────────┘

JSON:
  "type": "numbered_quadrant"
  "items": [
    {"number": "01", "title": "...", "bullets": ["...", "..."]},
    {"number": "02", ...}, ...
  ]
  "colors": ["accent", "mid", "grey", "dark"]  // 4박스 색상

shape 스펙:
  - 4개 사각형: 각 (w/2 - gap/2) × (h/2 - gap/2)
  - gap: 0.12"
  - 번호 원형: d=0.35", 좌상단 배치
  - 제목: 11pt Bold, 번호 우측
  - 불릿: 9pt Regular
  - 색상: 상단 바 5px + 흰 본문 (색상 절제 규칙)
```

### 5-2. Harvey Ball Matrix (C2)
```
구조:
  ┌──────┬──────┬──────┬──────┬──────┐
  │      │기준1 │기준2 │기준3 │총점  │
  ├──────┼──────┼──────┼──────┼──────┤
  │옵션A │ ◕   │ ◑   │ ●   │ 8/12│
  │옵션B │ ◔   │ ●   │ ◑   │ 7/12│
  │옵션C │ ●   │ ◕   │ ◕   │10/12│
  └──────┴──────┴──────┴──────┴──────┘

JSON:
  "type": "harvey_ball_matrix"
  "row_headers": ["옵션A", "옵션B", "옵션C"]
  "col_headers": ["기준1", "기준2", "기준3"]
  "scores": [[75, 50, 100], [25, 100, 50], [100, 75, 75]]
  // 0=empty, 25=quarter, 50=half, 75=three-quarter, 100=full

Harvey Ball shape:
  - 배경 원: d=0.25", fill=CL_BORDER (연회색)
  - 채움 원호: d=0.25", fill=CL_ACCENT, 각도=score*3.6°
  - python-pptx: MSO_SHAPE.PIE로 구현 (start_angle, sweep_angle)
```

### 5-3. Vertical Step-Down (P7/B5)
```
구조:
  [라벨열]  [Step 1] ─────── [우측 상세]
               │
           [Step 2] ─────── [우측 상세]
               │
           [Step 3] ─────── [우측 상세]

JSON:
  "type": "vertical_flow"
  "steps": [
    {"label": "...", "detail": "...", "style": "accent"},
    ...
  ]

shape 스펙:
  - 좌: 라벨 사각형 w=1.5", 색상=Dark
  - 중앙: Step 박스 w=3.0", 상단바 5px
  - 우: 상세 텍스트 w=4.5"
  - 수직 화살표: connector, 두께 1pt, CL_GREY
```

### 5-4. Funnel (B3)
```
구조:
  ╔══════════════════════════╗  Stage 1: 1000건
  ╚════════════════════════╝    → 80%
    ╔══════════════════╗       Stage 2: 800건
    ╚════════════════╝          → 50%
      ╔══════════════╗         Stage 3: 400건
      ╚════════════╝            → 25%
        ╔══════════╗           Stage 4: 100건
        ╚════════╝

JSON:
  "type": "funnel"
  "stages": [
    {"label": "Stage 1", "value": 1000, "conversion": "80%"},
    ...
  ]

shape: 사다리꼴 (MSO_SHAPE.TRAPEZOID) 크기 순차 감소
```

### 5-5. RAG Dashboard (C3)
```
구조:
  ┌────────────┬────────┬────────┬────────┐
  │ 항목       │ 상태   │ 트렌드 │ 비고   │
  ├────────────┼────────┼────────┼────────┤
  │ KPI 1      │ 🟢    │ ↑      │ 정상   │
  │ KPI 2      │ 🟡    │ →      │ 주의   │
  │ KPI 3      │ 🔴    │ ↓      │ 조치요 │
  └────────────┴────────┴────────┴────────┘

JSON:
  "type": "rag_table"
  "items": [
    {"name": "...", "status": "green", "trend": "up", "note": "..."},
    ...
  ]

RAG 원형: d=0.2", Green=#27AE60, Amber=#F39C12, Red=#C0392B
트렌드: ↑(green) →(amber) ↓(red) Unicode 화살표
```

### 5-6. RACI Matrix (E1)
```
→ framework_matrix의 특수 변형
  셀 값: R/A/C/I (단일 문자)
  셀 색상: R=CL_ACCENT, A=CL_ACCENT_MID, C=CL_GREY_LIGHT, I=아웃라인
  추가: 가중 점수 열 (선택)
```

### 5-7. SCR (Situation-Complication-Resolution) (G1)
```
→ stacked 레이아웃의 표준 패턴
  section[0]: h=0.15, 배경=연회색, "Situation" 라벨
  section[1]: h=0.15, 배경=연오렌지, "Complication" 라벨
  section[2]: h=0.55, 배경=흰색, "Resolution" 라벨
  section[3]: h=0.15, TakeawayBar
```

---

## Part 6: 구현 로드맵

```
[즉시] Tier 1 — 10개 (1~2일)
  신규 템플릿 7개 + 기존 개선 3개
  신규 컴포넌트 5개

[1주 내] Tier 2 — 10개
  전략 프레임워크 5개 + 조직/재무 5개

[2주 내] Tier 3 — 5개
  고난이도 다이어그램

최종 목표: 템플릿 35+종, 컴포넌트 25+종
```
