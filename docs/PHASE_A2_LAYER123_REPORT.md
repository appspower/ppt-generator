# Phase A2 — 3-Layer Multi-label 라벨링 결과

> 작성일: 2026-04-25
> 입력: 1,251장 마스터 PPT, Phase 1A/1B/1C 통계
> 출력: `output/catalog/auto_labels_v2.json`, `layer3_user_queue.json`, `review.html`

---

## 1. 3-Layer 구조 결과

| Layer | 작업 | 결과 |
|---|---|---|
| **L1** | 1B+1C 휴리스틱 통합 자동 라벨링 | 1,251장 / needs_review 238장 (19%) |
| **L2** | Agent 8개 병렬 30장씩 PNG 시각 검수 | 238장 검수 / agreed 41장 (17%) / 수정 197장 |
| **L3** | 사용자 검수 큐 (HTML) | 197장 (auto와 disagreed 또는 conf<0.6) |

---

## 2. 최종 라벨 분포 (Layer 2 후, 1,251장 전수)

### L1 Macro
| | 슬라이드 | % |
|---|---:|---:|
| diagram | 724 | 57.9% |
| table | 408 | 32.6% |
| card | 98 | 7.8% |
| chart | 16 | 1.3% |
| cover | 2 | 0.2% |
| unknown | 3 | 0.2% |

### L2 Archetype Top 10 (multi-label, 누계)
| | 슬라이드 |
|---|---:|
| orgchart | 728 |
| table_native | 373 |
| left_title_right_body | 251 |
| cards_5plus | 212 |
| dense_grid | 189 |
| flowchart | 183 |
| roadmap | 174 |
| hub_spoke | 170 |
| cards_2col | 119 |
| gantt | 116 |

### L3 Narrative Role 분포
- recommendation, analysis, evidence가 dominant (Layer 1에서 휴리스틱)
- Layer 3에서 사용자가 직접 archetype-role 매핑 보정 예정

---

## 3. 자동 vs 검수 일치율 (Layer 1 정확도)

| Batch | Agreed | 수정 | Avg conf |
|---|---:|---:|---:|
| 1 | 4/30 | 26 | — |
| 2 | 5/30 | 25 | 0.62 |
| 3 | 3/30 | 27 | 0.76 |
| 4 | 5/30 | 25 | 0.55-0.80 |
| 5 | 5/30 | 25 | ≥0.7 |
| 6 | 11/30 | 19 | 0.65-0.85 |
| 7 | 9/30 | 21 | 0.76 |
| 8 | 4/28 | 24 | high |
| **합계** | **46/238 (19%)** | **192** | ~0.74 |

**해석**:
- Layer 1만으로는 needs_review 슬라이드 중 **19%만 정확**. Layer 2 시각 검수의 가치 입증.
- 자동 휴리스틱 한계:
  - `left_title_right_body` 과다 적용 (Layer 1 default fallback)
  - 다이어그램(flowchart/swimlane/hub_spoke)을 card로 오분류
  - `unknown` 32장 → Layer 2 후 3장으로 감소
  - `cover_or_divider` 오적용 (실제는 content 슬라이드)

---

## 4. 흥미로운 패턴 발견 (Agent 보고)

1. **Pictogram/Icon library 6장 (1226~1232)** — 모두 35-cell icon grid + appendix role (high confidence)
2. **디자인 시스템 페이지 2장 (1219, 1222)** — 한국어 컴포넌트 카탈로그 reference
3. **반복 템플릿 변형**:
   - 488/493: 7×3 matrix + 3-stack
   - 370/371: 2-column comparison
   - 388/389: 미러 flowchart + cards_2col
   - 422/383: roadmap variants (vertical/horizontal chevron)
4. **As-Is/To-Be 패턴**: 1133이 swimlane + 신호등 신호 (Layer 1 cards_3col 오분류)
5. **macro 충돌**: 자동이 `card`로 잡은 슬라이드 다수가 실제 `diagram` (PwC 다이어그램이 카드 박스 형식을 빌려쓰는 경우 다수)

---

## 5. Layer 3 사용자 검수 (다음 단계)

**검수 화면**: `output/catalog/review.html` (894 KB, 197장)
- 슬라이드 PNG 썸네일 + 자동 라벨 + L1 radio / L2 checkbox / L3 checkbox
- 모든 항목 수정 후 "Save final_labels.json" 버튼 → 다운로드

**사용 방법**:
```
1. 브라우저로 file:///c:/Users/y2kbo/Apps/PPT/output/catalog/review.html 열기
2. 197장 PNG + 자동 라벨 보면서 수정
3. "📥 Save final_labels.json" 버튼 클릭 → final_labels_user.json 저장
4. 저장된 파일을 c:/Users/y2kbo/Apps/PPT/output/catalog/ 에 배치 후 알림
```

**예상 시간**: 슬라이드당 1분 × 197장 ≈ **3시간 (휴식 포함)**

**검수 우선순위**:
- conf < 0.6 슬라이드부터
- Layer 2가 Layer 1과 disagree한 슬라이드 (더 신뢰 가능하지만 사용자 도메인 지식으로 확정)

---

## 6. 산출 파일

| 경로 | 용도 |
|---|---|
| `output/catalog/auto_labels_v1.json` | Layer 1 자동 라벨 (1,251) |
| `output/catalog/layer2_batches/batch_*.json` | Layer 2 입력 batch 8개 |
| `output/catalog/layer2_results/batch_*.json` | Layer 2 검수 결과 8개 |
| `output/catalog/auto_labels_v2.json` | Layer 1+2 통합 (1,251) |
| `output/catalog/layer3_user_queue.json` | Layer 3 검수 큐 (197) |
| `output/catalog/review.html` | 사용자 검수 UI |

---

## 7. 다음 단계 (사용자 검수 후)

1. 사용자가 `final_labels_user.json` 제공
2. 통합 → `output/catalog/final_labels.json` (1,251장 전수)
3. **카탈로그 retrieval API 구현** (`ppt_builder/catalog/query.py`)
   - 입력: (narrative_role, max_chars 범위, archetype 선호) 
   - 출력: 상위 N개 slide_index
4. **PPT 생성 워크플로우 통합** (`docs/07_WORKFLOW.md` Step 2~3에서 사용)
5. 5개 벤치마크 시나리오 실증 (넷제로 / SAP ERP / 재무 / Change Mgmt / 전략 리뷰)
