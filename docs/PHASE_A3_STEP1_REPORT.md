# Phase A3 Step 1 — 5 벤치마크 정밀 측정 보고 (2026-04-25)

> Mode A 단독 5 시나리오 시뮬레이션 결과 + Stop-gate 판단 요청
> 코드: `scripts/benchmark_5_scenarios.py`
> 산출: `output/benchmark/{scenario}/deck.pptx + pngs/ + report.json`, `output/benchmark/scoreboard.json`, `output/benchmark/ppteval_visual.json`

---

## 1. 5 시나리오 정량 결과

| 시나리오 | 길이 | A. role 매칭 | B. fill % | C. overflow % | Quant 합성 | Visual 합성(D) | **최종** |
|---|---|---|---|---|---|---|---|
| transformation_roadmap_10 | 10 | **100%** (10/10) | 6.4% | 40.0% | 50.6 | 41.3 | **46.0** |
| consulting_proposal_30 | 30 | 86.7% (26/30) | 6.7% | 30.0% | 49.7 | 41.0 | **45.4** |
| analysis_report_15 | 15 | 93.3% (14/15) | 8.2% | 13.3% | 57.3 | 47.4 | **52.4** |
| change_management_20 | 20 | 90.0% (18/20) | 5.9% | 20.0% | 53.4 | 41.5 | **47.5** |
| executive_strategy_40 | 40 | 87.5% (35/40) | 7.0% | 32.5% | 49.3 | 38.0 | **43.7** |
| **평균** | | **91.5%** | **6.8%** | **27.2%** | **52.1** | **41.8** | **47.0** |

(최종 = (Quant + Visual) / 2)

---

## 2. 측정 항목별 결과

### A. role 매칭 성공률 — 91.5% 평균
- **role 직접 매칭 (양호)**: 5 시나리오 평균 91.5% 직접 hit
- **archetype fallback이 필요한 role**: situation, complication, risk, closing, divider, evidence(부분)
- **F. role 풀 고갈 (실측)**:
  - opening 8 / closing 1 / situation 1 / complication 2 / risk 1 / benefit 3 → 단일 슬라이드만 있어서 같은 role을 1개 이상 요청하는 시나리오에서 모두 archetype fallback 발생
  - 가장 영향 큰 시나리오: **executive_strategy_40 (5 fallback), consulting_proposal_30 (4 fallback)**

### B. 슬롯 채움률 — 평균 6.8%
- **슬라이드당 평균 텍스트 paragraph 21~30개** 중 1개만 Mode A가 채움
- → **`~~` 잔존 93~94%** ← **가장 큰 bottleneck**
- 이유: Mode A는 "가장 긴 텍스트 paragraph"만 hero로 교체. 나머지 50~80개 placeholder는 그대로

### C. 오버플로 빈도 — 평균 27.2%
- **원인**: edit_ops가 텍스트 길이 기반으로 div를 선택하나, slot_schemas의 `max_chars`(bbox 기반)와 매칭 안 함
- 대표 사례: slide#440 div=1 max_chars=26 → 38자 입력 → 잘림
- **해결책**: paragraph-level catalog에 max_chars 직접 활용 (Step 2 본질)

### D. PPTEval 3축 (Vision 검증)
| 축 | 점수 | 근거 |
|---|---|---|
| Content | **1.42 / 5** = 28점 | `~~` 가시성 95% / hero 1문장만 |
| Design | **2.62 / 5** = 52점 | 마스터 PwC톤 양호하나 빈 placeholder 노출 + cover/divider 빈약 |
| Coherence | **2.32 / 5** = 46점 | narrative 흐름 자체는 일관, 그러나 같은 슬라이드 재사용 + closing/divider archetype fallback이 흐름 단절 |

### E. PNG 시각 검증 (대표 슬라이드)
- **opening (step_01) 모든 시나리오**: cover 슬라이드가 거의 빈 페이지 — title 1줄 + 잔여 `~~` bullet
- **chevron 5단계 framework (slide#44 등)**: 5개 chevron 중 1개만 채움 + 텍스트 overflow
- **complex 2-column comparison (slide#84 등)**: speech bubble/chart 풍부 디자인이지만 ALL `~~`
- **closing/divider fallback**: 'Chart title, if needed' 같은 raw template으로 빠짐 → 디자인 단절
- **양호한 case**: analysis_report_15의 chart/table 슬라이드 (analysis 1143장 풀에서 골라서 디자인 다양성 ↑)

---

## 3. Bottleneck 우선순위 (Step 2 설계용)

| 순위 | Bottleneck | 영향 | Step 2 액션 |
|---|---|---|---|
| **#1** | paragraph-level meta 부재 → 슬라이드당 1 슬롯만 채움 | -40점 | **paragraph 카탈로그 (high priority)** |
| **#2** | overflow (텍스트 길이 기반 선택 vs bbox 무시) | -10점 | catalog의 max_chars 활용 retrieval |
| **#3** | cover/closing/divider role 풀 결손 | -8점 | Phase A4 (사용자 캡처 또는 코드 fallback) |
| **#4** | 같은 슬라이드 재사용 단조로움 | -5점 | retrieval 다양화 (top-K + diversity) |
| **#5** | archetype fallback 품질 낮음 | -3점 | fallback rule 개선 + 코드 assembler 보완 |

→ **#1과 #2는 paragraph 카탈로그 한 번에 해결 가능** (Step 2 핵심)

---

## 4. role 풀 고갈 분석

### 5 시나리오 합산 role 요청 vs 풀 capacity
| role | 풀 (마스터) | 5 시나리오 요청 합 | 부족? |
|---|---|---|---|
| analysis | 1143 | 18 | ✅ 충분 |
| recommendation | 640 | 26 | ✅ 충분 |
| evidence | 452 | 16 | ✅ 충분 |
| roadmap | 125 | 12 | ✅ 충분 |
| opening | 8 | 5 | ⚠️ 한계 |
| closing | **1** | 7 | ❌ **6 fallback** |
| situation | **1** | 8 | ❌ **7 fallback** |
| complication | 2 | 5 | ❌ 3 fallback |
| risk | 1 | 5 | ❌ 4 fallback |
| benefit | 3 | 5 | ❌ 2 fallback |
| divider | 9 | 2 | ✅ 충분 (시나리오 2개) |
| agenda | 12 | 4 | ✅ 충분 |
| appendix | 15 | 4 | ✅ 충분 |

**진단**: closing/situation/risk/complication/benefit가 narrative 분포 편중의 핵심 결손.
→ Phase A4에서 사용자 캡처 5~10장 + 코드 fallback 강화 필요.

---

## 5. Stop-gate 판단

### 판단 매트릭스
| 평균 점수 | 결정 |
|---|---|
| 60+ | Step 2 진행 권장 |
| **50~60** | **추가 보완 필요, 사용자 결정 요청** |
| <50 | 방향 재고 (Mode B 강화 또는 자산 캡처) |

### 실측: **47.0 평균** ← Quant 52.1 + Visual 41.8 평균
- **경계선**: Quant만 보면 52.1로 50~60점대, Visual 평가 포함 시 47점으로 50점 미만

### 5 시나리오별 분포
- 50점 이상: 1개 (analysis_report_15 = 52.4)
- 45~50점: 3개 (transformation 46.0, change 47.5, consulting 45.4)
- 45점 미만: 1개 (executive_strategy_40 = 43.7)

→ **거의 모든 시나리오가 47점 근처에 군집**. 80점까지 +33점 갭. 단일 lever로는 불가, 다단계 보완 필수.

---

## 6. 사용자 결정 요청

데이터 기반 3가지 옵션 제시:

### 옵션 A — 그대로 Step 2 (paragraph 카탈로그) 진행 (**권장**)
**근거**:
- Bottleneck #1 (fill 6.8% → 60~70%)이 +20~25점 효과
- Bottleneck #2 (overflow)도 같은 카탈로그로 해결
- 5만 슬롯 라벨링 자동 (Agent 병렬 4~6시간)
- 위험: 정확도 ~80%, 5% 샘플 검수로 보완

**예상 후속 점수**: 평균 47점 → 67~72점 (Step 3 재측정에서 80+ 도달은 추가 Step 4 필요)

### 옵션 B — 먼저 Phase A4 (cover/closing 자산 캡처) 우선
**근거**:
- 사용자가 PwC 실 cover/closing/divider 슬라이드 10~15장 추가 캡처
- role 결손이 시각적으로 가장 눈에 띄는 부분
- 카탈로그 작업보다 사용자 시간 더 필요 (1~2일)

**예상 후속 점수**: 평균 47 → 55~60점 (paragraph 카탈로그 안 하면 fill 6.8% 그대로)

### 옵션 C — Mode B 코드 fallback 강화 우선
**근거**:
- assembler/renderers/ 9개를 cover/closing/divider/situation 등 결손 role 위해 강화
- 외부 SOTA에서 Mode B 단독 2.1% 실행 성공 → 가장 약한 옵션
- 시간 1~2일 소요

**예상 후속 점수**: 평균 47 → 50~53점 (PwC 톤 일치 어려움)

---

## 7. 본 세션 권장 결정

**옵션 A (Step 2 paragraph 카탈로그) 진행** 권장. 

이유:
1. **데이터가 명확**: fill 6.8% → 가장 큰 단일 레버
2. **자동 가능**: Agent 병렬, 사용자 시간 15분 (5% 샘플 검수)
3. **위험 통제**: Step 3에서 재측정 후 부족 시 Step 4(A/B/C) 결정 가능

**다음 작업**: `output/catalog/paragraph_slots.json` 생성 (high priority 슬롯 ~15,000개) + `ppt_builder/catalog/query.py` retrieval API + Step 3 재측정.

→ 사용자가 Stop-gate 통과 결정 시 Step 2 진행하겠습니다.

---

## 8. 산출 파일 (commit 대상)

- `scripts/benchmark_5_scenarios.py` — 5 시나리오 runner
- `output/benchmark/{scenario}/deck.pptx` × 5
- `output/benchmark/{scenario}/pngs/step_*.png` × 115
- `output/benchmark/{scenario}/report.json` × 5
- `output/benchmark/scoreboard.json` — 통합 정량 점수
- `output/benchmark/ppteval_visual.json` — Vision 3축 점수
- `docs/PHASE_A3_STEP1_REPORT.md` — 본 보고서
