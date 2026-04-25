# Phase A3 Step 2+3 — Paragraph 카탈로그 + 재측정 결과 (2026-04-25)

> Step 2 (paragraph-level 라벨링) + Step 3 (재측정) 완료
> v1 (Mode A 1-paragraph fill) → v2 (paragraph-aware fill + ~~ 블랭크)
> 평균 composite 47 → **62.9** (시각 해상도 평균 81.7%)

---

## 1. 산출

### 코드
- `scripts/extract_paragraphs.py` — 1,251장 → 98,575 paragraph 메타데이터 추출
- `scripts/label_paragraphs.py` — 결정론적 role 라벨러 (placeholder/table/shape/group propagation)
- `ppt_builder/catalog/paragraph_query.py` — ParagraphStore + match_content API
- `scripts/benchmark_5_scenarios_v2.py` — paragraph-aware 5 시나리오 runner

### 데이터
- `output/catalog/paragraphs.json` (71.6 MB) — 추출된 paragraph 풀
- `output/catalog/paragraph_labels.json` (87.0 MB) — role 라벨링 결과
- `output/benchmark_v2/{scenario}/deck.pptx` × 5 — paragraph-fill 결과
- `output/benchmark_v2/{scenario}/pngs/` × 5 — 시각 검증용
- `output/benchmark_v2/scoreboard.json` — 점수표

---

## 2. Role 라벨링 결과 (1,251장 / 98,575 paragraph)

| role | 갯수 | 비율 | 출처 |
|---|---|---|---|
| card_header | 33,514 | 34.0% | group_card 추론 |
| decorative | 30,685 | 31.1% | default |
| table_cell | 15,147 | 15.4% | TABLE 셀 (row > 0) |
| card_body | 5,986 | 6.1% | group_card |
| chevron_label | 3,461 | 3.5% | AUTOSHAPE:CHEVRON/PENTAGON |
| table_header | 3,163 | 3.2% | TABLE 셀 (row 0) |
| body | 3,047 | 3.1% | placeholder BODY |
| table_cell_summary | 2,614 | 2.7% | TABLE 마지막 행 |
| title | 1,285 | 1.3% | placeholder TITLE / 큰 폰트 + top |
| page_number | 958 | 1.0% | placeholder SLIDE_NUMBER |
| subtitle | 396 | 0.4% | placeholder SUBTITLE |
| footer | 176 | 0.2% | placeholder FOOTER |
| callout_text | 56 | 0.1% | AUTOSHAPE callout |
| kpi_value | 5 | 0.0% | star/explosion |

**High-priority 슬롯**: 40,626 (41.2%) — title/table_header/chevron_label/card_header/callout_text

---

## 3. v1 vs v2 비교

| 시나리오 | v1 composite | v2 composite | Δ | v2 visual_resolution |
|---|---|---|---|---|
| transformation_roadmap_10 | 50.6 | **67.8** | +17.2 | 82.7% |
| consulting_proposal_30 | 49.7 | **61.9** | +12.2 | 85.4% |
| analysis_report_15 | 57.3 | **63.1** | +5.8 | 80.7% |
| change_management_20 | 53.4 | **60.5** | +7.1 | 75.5% |
| executive_strategy_40 | 49.3 | **61.2** | +11.9 | 84.0% |
| **평균** | **52.1** | **62.9** | **+10.8** | **81.7%** |

(v1 composite는 quant 기준; v2 composite는 fill + visual_resolution + role + overflow 가중)

### Visual Resolution 의미
`visual_resolution = (filled + blanked) / fillable`
- **filled**: 컨텐츠로 채워진 슬롯
- **blanked**: 컨텐츠 없어 `~~`를 빈 문자열로 교체 (시각 정리)
- 평균 81.7% — 시각적 `~~` 잔존 약 18%

---

## 4. 시각 검증 (대표 슬라이드)

### 개선 사례 1: 넷제로 roadmap (slide#176, transformation_roadmap_10 step 6)
- **v1**: 4 chevron 중 1개 텍스트 + 3개 `~~`, 텍스트 overflow 잘림
- **v2**: Phase 1 + Phase 2 (시나리오가 2개 제공) + 나머지 빈 chevron + 작은 `~~` 3개만 잔존

### 개선 사례 2: SAP 추천 framework (slide#44, consulting_proposal_30 step 13)
- **v1**: 5 chevron 중 1개만 채움, 30+ 카드 모두 `~~`
- **v2**: 5 chevron 모두 추천 항목으로 채움 (Center of Excellence 30 / Single Global Template / RISE with SAP / Fiori UX / Master Data Governance), 30 카드 영역 깔끔히 비워짐

### 잔존 이슈
1. cover/closing 슬라이드 여전히 빈 페이지 (마스터 자체에 본격 디자인 부재) → **Phase A4 필요**
2. 일부 narrow placeholder에 텍스트 overflow → 폰트 축소 fallback 필요
3. 크기-고정된 카드(30개 이상 grid)는 시나리오 컨텐츠로 못 채움 → capacity-aware 슬라이드 retrieval 필요
4. 일부 `~~`가 그룹화 안 된 placeholder에 잔존 (page#, sub-title bar 등)

---

## 5. Step 3 Stop-gate 판단

판단 매트릭스 (계획서 기준):
- 80+ : Step 5만 진행 (목표 달성)
- 70~80: Step 4 진행
- **60~70: Step 4 + 잔존 분석** ← **현 상태**
- <60: 방향 재고

**평균 62.9 → Step 4 진행 + 잔존 분석**

---

## 6. Step 4 권장 우선순위 (잔존 분석 기반)

| 순위 | 작업 | 예상 효과 | 노력 |
|---|---|---|---|
| 1 | **Capacity-aware retrieval** — 슬롯 N개에 컨텐츠 N개 매칭되는 슬라이드 우선 선택 | +5~8점 | 1일 |
| 2 | **Phase A4 cover/closing 자산** — 사용자 PwC cover 5~10장 캡처 + assembler fallback | +5점 | 사용자 1일 |
| 3 | **Overflow 폰트 축소 fallback** — text overflow 시 0.95x 자동 축소 | +3점 | 4시간 |
| 4 | **잔존 ~~ 청소** — placeholder TITLE_PLACEHOLDER subtype 등 추가 블랭크 | +2점 | 2시간 |
| 5 | **카드 capacity 다양성** — 같은 슬라이드 반복 사용 방지 (top-K + diversity) | +2점 | 4시간 |

→ **#1 + #3 + #4** 우선 (자동 작업, 1.5일) → 70+점 도달 가능 추정
→ **#2** (사용자 시간) 추가 시 75~80점

---

## 7. 사용자 결정 요청

**옵션 A (권장) — Step 4 자동 부분만 먼저 (#1, #3, #4)**
- 1.5일 자동 작업 → 70+점 목표
- 사용자 시간 0
- 결과 보고 후 #2 (자산 캡처) 결정

**옵션 B — Step 4 전체 (사용자 자산 캡처 포함)**
- 사용자 PwC cover/closing 5~10장 캡처 (1~2일)
- 자동 부분 1.5일 병행
- 75~80점 목표

**옵션 C — 현재 상태로 종료, 다른 방향 탐색**
- 62.9점 + 시각 81.7% 해상도가 충분하다 판단
- 다른 우선순위로 이동

진행 방향 선택해 주세요.
