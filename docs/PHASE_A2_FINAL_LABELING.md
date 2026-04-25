# Phase A2 — 1,251장 전수 라벨링 최종 보고

> 완료일: 2026-04-25
> 입력: 1,251장 마스터 PPT
> 출력: `output/catalog/final_labels.json` (1,251장 multi-label)

---

## 1. 최종 결과 요약

**전수 검수 완료 — 1,251 / 1,251 장 (100%)**

| 항목 | 값 |
|---|---|
| 총 슬라이드 | 1,251 |
| Layer 1 (자동 휴리스틱) | 1,251 |
| Layer 2 (Agent 시각 검수) | **1,251** (전수) |
| Layer 1 자동 정확도 (agreed) | **160 / 1,251 (12.8%)** |
| 평균 confidence | 0.85 (median) |

**Layer 1 정확도가 매우 낮음을 정량 확인** — 자동 휴리스틱만 신뢰했다면 87% 라벨이 잘못됐을 것. 사용자가 전수 검수를 지시한 판단이 정확.

---

## 2. 최종 라벨 분포 (1,251장)

### L1 Macro
| | 슬라이드 | % |
|---|---:|---:|
| **diagram** | **793** | **63.4%** |
| table | 331 | 26.5% |
| card | 109 | 8.7% |
| chart | 12 | 1.0% |
| unknown | 4 | 0.3% |
| cover | 2 | 0.2% |

→ **컨설팅 자산은 다이어그램 중심 (63%)** + 표 (26%) + 카드 (9%). 차트/cover는 매우 적음.

### L2 Archetype Top 15 (multi-label, 누계)
| | 슬라이드 |
|---|---:|
| orgchart | 345 (Layer 1 728 → 345, **-53%**) |
| table_native | 339 |
| flowchart | 318 (Layer 1 183 → 318, **+74%**) |
| dense_grid | 273 |
| cards_5plus | 190 |
| left_title_right_body | 187 |
| roadmap | 131 |
| vertical_list | 130 |
| cards_3col | 126 |
| cards_2col | 125 |
| swimlane | ~110 |
| hub_spoke | ~100 |
| matrix_NxN | ~80 |
| matrix_2x2 | ~50 |
| timeline_h | ~50 |

### Layer 2 보정 패턴 (Layer 1 대비)
- **orgchart 53% 감소**: 자동 휴리스틱 과적용 정정
- **flowchart 74% 증가**: orgchart에서 재할당됨 (실제 트리 아니라 박스+화살표 흐름)
- **swimlane / vertical_list 대량 신규**: 자동이 left_title_right_body로 후퇴했던 것
- **gantt 거의 0**: 시간축 없는 chevron이 대부분 → roadmap으로 재할당
- **chart_native 누락 보강**: waterfall/box plot/histogram 신규 식별 (#986, #987, #1219, #1244)

### L3 Narrative Role
| | 슬라이드 |
|---|---:|
| analysis | 1,143 |
| recommendation | 640 |
| evidence | 452 |
| roadmap | 125 |
| appendix | 15 |
| agenda | 12 |
| divider | 9 |
| opening | 8 |
| (그 외) | 12 |

→ **analysis/recommendation/evidence가 컨설팅 덱의 80% 차지**. opening/closing/divider는 매우 적음 (마스터 템플릿 자체가 본문 슬라이드 위주).

---

## 3. 자동 vs 검수 정확도 (Layer 1 휴리스틱 한계)

| 카테고리 | Layer 1 자동 | 검수 후 정정 |
|---|---:|---:|
| orgchart | 728 (58%) | **345 (28%)** |
| flowchart | 183 (15%) | **318 (25%)** |
| swimlane | (낮음) | ~110 (9%) |
| gantt | 116 (9%) | ~30 (2%) |
| vertical_list | (낮음) | 130 (10%) |
| chart_native | 12 | ~20 (보강) |

**핵심 진단**:
- 자동 휴리스틱이 **박스+선 = orgchart**로 거의 모든 다이어그램을 분류
- 진짜 orgchart (3+ 깊이 트리, 명확 부모-자식) 는 약 50~80장 추정 (많은 변형 있어 명확치 않음)
- 자동 detector 자체에 큰 false positive

---

## 4. 검수 작업 통계

### Round별 진행
| Round | Batch | 슬라이드 | 평균 agreed | 시간 |
|---|---|---:|---:|---|
| 1 (이전) | 1~8 (needs_review) | 238 | 19% | 8 Agent 병렬 |
| 1 (이번) | 9~16 | 240 | ~20% | 8 Agent 병렬 |
| 2 | 17~24 | 240 | ~10% | 8 Agent 병렬 |
| 3 | 25~32 | 240 | ~13% | 8 Agent 병렬 |
| 4 | 33~40 | 240 | ~13% | 8 Agent 병렬 |
| 5 | 41~42 | 53 | ~13% | 2 Agent |
| **합계** | **42 batch** | **1,251** | **12.8%** | 약 1.5시간 |

### 자주 보정된 패턴
1. **orgchart 과적용**: 거의 모든 batch에서 80~95% 보정
2. **gantt 오분류**: 시간축 없는 chevron을 gantt로 → roadmap
3. **left_title_right_body 후퇴**: cards_Ncol/swimlane이 더 적합
4. **swimlane 누락**: 좌측 카테고리 라벨 + row별 컨텐츠 패턴 인식 부족
5. **chart_native 누락**: 파이/막대 외 (waterfall, box plot, histogram) 누락

---

## 5. 흥미로운 발견

1. **컨설팅 덱은 process/diagram 중심 자산** (63% diagram, +26% table) — text-heavy 슬라이드 거의 없음
2. **PwC 시그니처 패턴**: 좌측 phase chevron + 우측 swimlane (전체 slide의 ~10%)
3. **Pictogram library**: 1226~1232 (35-cell icon grid, appendix role)
4. **As-Is/To-Be 패턴**: swimlane + 신호등 신호로 구현 (다수)
5. **반복 템플릿 변형**: 여러 슬라이드가 거의 동일한 구조 변형 (488/493, 370/371, 388/389 등)
6. **narrative_role 편향**: analysis 91%, recommendation 51% — 마스터는 본문 슬라이드 위주

---

## 6. 산출 파일

| 경로 | 용도 |
|---|---|
| `output/catalog/final_labels.json` | **1,251장 최종 라벨 (multi-label)** |
| `output/catalog/auto_labels_v1.json` | Layer 1 자동 (참고) |
| `output/catalog/auto_labels_v2.json` | Layer 1+2 통합 |
| `output/catalog/layer2_results/batch_*.json` | 42 batch 검수 결과 |
| `output/catalog/full_batches/batch_*.json` | 검수 입력 batch |
| `output/catalog/all_pngs/slide_*.png` | 1,251 PNG (89MB) |
| `output/catalog/slot_schemas.json` | 73,425 슬롯 max_chars |
| `output/catalog/skeletons.json` | 7 narrative 스켈레톤 |
| `output/catalog/shape_features.npy` | OOXML 68-dim |
| `output/catalog/dit_embeddings.npy` | DiT 768-dim |

---

## 7. 다음 단계 (Phase A3)

1. **카탈로그 retrieval API 구현** (`ppt_builder/catalog/query.py`)
   - 입력: (narrative_role 시퀀스, max_chars 범위, archetype 선호도)
   - 출력: 상위 N개 slide_index + 신뢰도
2. **PPT 생성 워크플로우 통합** (`docs/07_WORKFLOW.md` Step 2~3)
3. **5개 벤치마크 시나리오 실증**:
   - 넷제로 전환 10장 (transformation_roadmap_10)
   - SAP ERP 30장 (consulting_proposal_30)
   - Q1 재무 분석 15장 (analysis_report_15)
   - 조직 개편 20장 (change_management_20)
   - 연간 전략 40장 (executive_strategy_40)
4. **evaluate_deck.py** — Track 3 deck-level coherence 하니스
