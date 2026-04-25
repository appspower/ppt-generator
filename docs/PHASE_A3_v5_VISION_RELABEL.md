# Phase A3 v5 — Vision 라벨 재검수 결과 (2026-04-25)

> 1차 옵션 (P1 + 옵션 X 통합) 진행: Agent 8 병렬로 144장 마스터 PPT vision 검수 완료
> Track 2 자산 캡처 회피 + HJ 패턴 reference 활용

---

## 1. 작업 요약

### 후보 추출
- archetype에 `cover_divider` 있는 슬라이드: 21장
- fillable 슬롯 ≤ 10 (단순 메시지 = cover/closing/divider 후보): 127장
- 합집합 unique: **144장**

### 8 Agent 병렬 vision 검수
- 18장씩 8 batch로 분할
- 각 Agent: PNG vision으로 narrative_role 판정
- 결과: 144장 중 128장 override (conf≥0.6) + 7장 merge + 9장 decorative 처리

### 통합 → final_labels_v2.json
- 모든 1,251장 라벨 갱신 (vision 검수 분만 수정, 나머지 유지)

---

## 2. 발견된 결손 role 변화

| role | v4 | **v5** | Δ | 평가 |
|---|---|---|---|---|
| **complication** | 2 | **9** | **+7** | ✅ 큰 win |
| **benefit** | 3 | **5** | **+2** | ✅ |
| **closing** | 1 | **2** | **+1** | 작은 개선 |
| **appendix** | 15 | 18 | +3 | 정확도 ↑ |
| **agenda** | 12 | 11 | -1 | 미세 |
| **opening** | 8 | **2** | **-6** | ❌ false positive 정리 |
| **divider** | 9 | **4** | **-5** | ❌ false positive 정리 |
| **situation** | 1 | 1 | 0 | 변화 없음 |
| **risk** | 1 | 1 | 0 | 변화 없음 |
| analysis | 1143 | 1118 | -25 | 정확도 ↑ |
| recommendation | 640 | 630 | -10 | 정확도 ↑ |
| evidence | 452 | 434 | -18 | 정확도 ↑ |

### 핵심 발견
1. **complication +7**: 463/464/495 (red issue boxes), 896/900/901/937 (As-Is→To-Be 패턴) 등이 자동라벨러에서 missing → 발견
2. **opening/divider 감소**: 자동라벨러가 chart library page (1217/1218/1220), agenda 등을 잘못 opening/divider로 분류 → vision으로 정정. **단기적으로 풀 감소지만 정확도는 ↑**
3. **closing/situation/risk는 거의 변화 없음**: 마스터 풀에 진짜로 그런 슬라이드가 없거나, 144장 후보 외에 숨어있음

---

## 3. 5 시나리오 점수 비교

| 시나리오 | v4 (auto-label) | **v5 (vision)** | Δ |
|---|---|---|---|
| transformation_roadmap_10 | 68.8 | 67.1 | -1.7 |
| consulting_proposal_30 | 60.2 | 59.2 | -1.0 |
| analysis_report_15 | 53.3 | 53.0 | -0.3 |
| change_management_20 | 50.8 | 52.9 | +2.1 |
| executive_strategy_40 | 59.0 | 56.8 | -2.2 |
| **평균** | **58.4** | **57.8** | **-0.6** |

### 주의 — 점수 metric의 한계
- v5에서 role 매칭 정확도는 ↑ (consulting 86.7→93.3%, executive 87.5→92.5%)
- 하지만 visual_resolution은 ↓ (false positive 슬라이드가 fillable 풀에서 빠져 분포 변동)
- **점수 자체는 거의 변화 없으나 라벨 정확도는 명확히 개선**

---

## 4. 시각 검증 — v5의 진짜 가치

### 개선된 부분
- **slide_1217/1218** (chart library page) → 자동라벨이 opening으로 잘못 분류했던 것을 vision으로 decorative로 수정. opening 시나리오에서 우연히 picked되는 위험 제거.
- **complication 발굴**: 463/464가 As-Is→To-Be 문제 정의 슬라이드인데 evidence/analysis로 잘못 분류됐던 것을 정정. 시나리오에서 complication 매칭 실패율 감소.

### 변하지 않은 부분
- **transformation step 5 (recommendation)**: "전기로 전환 / 수소환원제철 / 그린전력 PPA" 3 chevron 정확 분배 (v4 그대로 유지)
- 시각적 큰 개선은 없으나 잘못된 슬라이드 선택 방지 효과

---

## 5. 정직한 평가

### v5의 효과
- **정량 점수**: -0.6 (거의 변화 없음)
- **라벨 정확도**: 명확히 개선 (false positive 정리)
- **결손 role 발굴**: complication +7, benefit +2 → 이 분야 시나리오는 부분 개선
- **opening/divider 감소**: 풀이 진짜 빈약함 노출 → Track 2 캡처 더 절실

### Vision relabel의 진짜 의의
- "자동 라벨 41% 정확도" → "vision 검수 후 추정 70-80% 정확도"
- 점수가 동일해도 **잘못된 매칭 방지** 효과 (시각적 일관성 ↑)
- **결손 role 풀 한계 데이터적으로 확인**: opening 2장 / divider 4장 / closing 2장 / situation 1 / risk 1 = 진짜로 부족

### 자동만으로 70+ 도달은 어려움 재확인
- 144장 vision 검수해도 평균 점수는 거의 동일
- closing/situation/risk가 마스터 풀에 진짜 없음 (자동라벨 미스가 아니라 자료 부재)
- **Track 2 사용자 자산 캡처 또는 Mode B 코드 fallback이 다음 lever**

---

## 6. 다음 단계 (재정리)

| 옵션 | 작업 | 예상 효과 | 시간 | 회사PC |
|---|---|---|---|---|
| **P3 (Track 2)** | PwC PPT 5~10장 캡처 → 마스터 합병 | +5~10점 | 사용자 1~2h | ✅ 필요 |
| **P2 (Mode B)** | assembler로 cover/closing/divider 코드 생성 | +5~8점 | 자동 2~3일 | ❌ |
| 옵션 Y | 1437장 vision → placeholder 슬라이드 자동 생성 | +5~8점 | 자동 2~3일 | ❌ |
| 옵션 W | 추가 vision 검수 (의심 confidence 0.5~0.7 슬라이드) | +1~3점 | 자동 30분 | ❌ |
| 현 상태 종료 | 58점으로 실 사용 테스트 | - | 즉시 | - |

### 권장
- **회사PC 접근 가능 시**: P3 (가장 ROI 높음)
- **자동만**: P2 (Mode B) 시작
- **단기 개선만**: 옵션 W (의심 슬라이드 추가 검수)

---

## 7. 산출

- `scripts/extract_relabel_candidates.py` — 144장 후보 추출
- `scripts/merge_vision_relabel.py` — 8 batch 통합
- `output/catalog/vision_relabel_batches/` — 8 batch 입출력 + manifest
- `output/catalog/final_labels_v2.json` — vision-relabeled 카탈로그
- `output/benchmark_v2/` — v5 5 시나리오 결과
- 본 보고서

---

## 8. 종합 — 이 세션 전체 요약

| 단계 | 점수 | 핵심 변경 |
|---|---|---|
| Mode A 단독 (시작) | 47.0 | baseline |
| Step 2+3 paragraph fill | 62.9 | 슬롯 단위 채움 |
| Step 4 v3 (~~ 청소) | ~57 | 시각 청결 |
| Step 4 v4 (content expand + diversity) | 58.4 | fill 8%→18% |
| **Step 4 v5 (vision relabel)** | **57.8** | **라벨 정확도 ↑** |

**자동 작업 한계 = 50대 후반**. 70+ 도달은 사용자 자산 캡처(P3) 또는 Mode B(P2) 필요.
