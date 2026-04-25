# Phase A2 → A3 방향 분석 보고서 (2026-04-25)

> 입력: 3 Agent 병렬 분석 (시스템/외부 리서치/실측 시뮬레이션)
> 목적: Mode A/B/C 중 80+점 컨설팅 그레이드 PPT를 가장 안정적으로 만드는 방향 결정

---

## 1. 결정적 실측 증거 — 3 분석 모두 같은 결론에 수렴

### 1A 시스템 평가 (5점 매트릭스)
| | Mode A | Mode B | Mode C |
|---|---|---|---|
| 인프라 완성도 | 5 | 4 | 4 |
| 1,251장 자산 적합성 | 3 | 2 | **4** |
| 80+점 달성 가능성 | 4 | 2 | **5** |
| 사용자 노력 | 5 | 2 | 4 |
| 유지보수 | 4 | 2 | 3 |
| **총점** | 21/25 | 12/25 | **20/25** |

### 1B 외부 SOTA 벤치마킹
- **Mode A (PPTAgent EMNLP 2025)**: PPTEval 3.62~3.67/5, design+coherence SOTA, 95% 실행 성공
- **Mode B 단독 (AutoPresent)**: LLaMA8B raw **2.1% 실행 성공**, color fidelity 10~18 vs human 73.5 → **컨설팅 그레이드 부적합**
- **Mode B + SlidesLib**: +10~34pt 부스트지만 design gap 존재
- **컨설팅 실무**: McKinsey "library of old cases", BCG Knowledge Navigator → **재사용 우선**
- **1,251장 = PPTAgent reference pool(50장)의 25배** → 충분

### 1C Mode A 실측 시뮬레이션 (넷제로 10장)
- 인프라 작동: 10장 9.7초 / role 매칭 10/10 / edit 무오류
- **종합 점수 52/100** — 80점에 28점 부족
- 5가지 결정적 한계 측정:
  1. **슬라이드당 1개 placeholder만 채움** (실제는 7~87개) → 95% `~~` 잔존
  2. **role 풀 극단 편중**: opening 8 / closing 1 / situation 1 / risk 1 / benefit 3 — 누적 16장
  3. **텍스트 오버플로**: 5/10 슬라이드에서 잘림
  4. **cover 디자인 부재**: 마스터에 본격 cover 슬라이드 없음
  5. **catalog에 paragraph-level 메타 부재**: max_chars/component_role 없음

---

## 2. 핵심 통찰 — 왜 Mode A 단독은 52점인가

**근본 원인**: 카탈로그가 **slide-level**만 라벨링했고 **paragraph-level (각 `~~` 슬롯의 의미)** 정보가 없음.

```
현재:  슬라이드 #234 = {macro: card, archetype: cards_3col, role: recommendation}
필요:  슬라이드 #234 = {
         macro/archetype/role + 
         paragraph[0] = {role: title, max_chars: 80},
         paragraph[1] = {role: card1_header, max_chars: 30},
         paragraph[2] = {role: card1_body, max_chars: 200},
         paragraph[3] = {role: card2_header, max_chars: 30},
         ...
       }
```

→ **paragraph-level 카탈로그 확장이 80+점의 가장 큰 레버 (+20점)**

---

## 3. 3 안 최종 평가 매트릭스 (3 분석 통합)

| 평가 축 | Mode A 단독 | Mode B 단독 | **Mode C Hybrid** |
|---|---|---|---|
| 외부 SOTA 점수 | 3.62~3.67/5 | 2.1% 실행 | (PPTAgent 패러다임) |
| 우리 시스템 점수 | 21/25 | 12/25 | 20/25 |
| 실측 시뮬레이션 | **52/100** | 미실측 | (개선 시 80+ 추정) |
| 인프라 완성도 | ✅ | ⚠️ | ✅ |
| narrative 분포 결손 대응 | ❌ | ✅ | ✅ |
| 디자인 일관성 | ✅ | ❌ | ✅ |
| 사용자 노력 | 즉시 | 수개월 라이브러리 구축 | 즉시+점진 |
| **80+점 달성** | ❌ (52점) | ❌ | **✅ (예상 80~85)** |

---

## 4. 추천 — Mode C (Hybrid) 단계적 구현

### Phase A3 — paragraph-level 카탈로그 확장 (가장 큰 레버, +20점)
- 1,251장 슬라이드의 각 `~~` 슬롯을 분류:
  - `slot_role`: title / card_header / card_body / kpi_value / table_cell / chevron_label / footer 등
  - `max_chars`: bbox area 기반 정확 측정 (Phase 2d 데이터 활용)
  - `position_in_group`: group 내 순서 (의미 단위 매핑)
- **방법**: Agent 병렬 검수 (1,251장 × 평균 40 슬롯 = 5만 슬롯). batch 단위로.
- 출력: `catalog.json` (slide_index → slots[] with role + max_chars)

### Phase A4 — narrative_role 결손 보완 (+5점)
- opening/closing/situation/risk/benefit 6 role의 자산 보강:
  - **선택지 A**: 사용자가 추가 PwC 슬라이드 캡처 (10~20장)
  - **선택지 B**: 코드 fallback (assembler/renderers/) 강화로 기존 안 활용
- Layer 3 (코드 fallback) 강화: assembler 9개 renderer를 사용

### Phase A5 — 컨텐츠 정밀 매칭 + Overflow 처리 (+3점)
- 컨텐츠 → 적합 slide 검색: max_chars 매칭 + role 매칭
- 텍스트 오버플로 감지 시:
  - 자동 줄바꿈
  - 폰트 0.95x 스텝 축소 (3 step까지)
  - 그래도 넘치면 다른 max_chars 큰 슬라이드 후보로 fallback

### Phase A6 — 5 벤치마크 실증 + REFINE
- 실측한 52점 → 80+점 달성 확인
- 실패 케이스 root cause 분석 → Phase A3~A5 반복 보정

---

## 5. 시간/노력 추정

| Phase | 작업 | 시간 | 사용자 시간 |
|---|---|---|---|
| A3 | paragraph-level 카탈로그 (Agent 병렬) | 4~6시간 | 0 |
| A4 | role 결손 보완 (코드 fallback 강화) | 1~2일 | 0 |
| A5 | 컨텐츠 매칭 + overflow | 1일 | 0 |
| A6 | 5 벤치마크 실증 + REFINE | 1~2일 | 검토 1~2시간 |
| **합계** | | **3~5일** | **1~2시간** |

---

## 6. 위험 요소 (정직한 평가)

1. **paragraph-level 카탈로그 정확도**: Agent가 5만 슬롯 라벨링 시 ~85% 정확도 추정. 핵심 슬라이드만 사용자 검수하면 보완.
2. **role 결손이 코드 fallback으로 충분히 메울지**: 마스터에 부재한 디자인을 코드로 재현은 PwC 톤 일치 어려움. **선택지 A (사용자 캡처) 권장**.
3. **80+ 달성 미보장**: 시뮬레이션은 추정치. 5 벤치마크 실증에서 70대 점수 나오면 추가 Phase 필요할 수 있음.

---

## 7. 결정 요청

**3 분석 모두 Mode C 추천.** 단계별 진행 권장:

### 옵션 1 (권장) — Mode C 전체 단계 진행
- Phase A3 (paragraph 카탈로그) 즉시 시작
- A4~A6 순차 진행
- 3~5일 + 사용자 검토 1~2시간

### 옵션 2 — 단계별 stop-gate
- Phase A3만 먼저 → 결과 보고 → 사용자 결정 후 A4 진행
- 안전하지만 시간 ↑

### 옵션 3 — 빠른 실증 우선
- Phase A3 skip하고 5 벤치마크 시뮬레이션 먼저 (1일)
- 실측 52점 → 어디가 부족한지 정확히 측정 → 그 다음 A3 우선순위 결정
- 가장 정확하지만 우회로

---

## 산출 파일
- `c:/Users/y2kbo/Apps/PPT/output/simulation/netzero_modeA.pptx` (실측 시뮬레이션 결과)
- `c:/Users/y2kbo/Apps/PPT/output/simulation/netzero_modeA_pngs/step_*.png`
- `c:/Users/y2kbo/Apps/PPT/output/test_modeB_extract.pptx` (Mode B group 추출 검증)
- `c:/Users/y2kbo/Apps/PPT/scripts/simulate_mode_a.py`
- 본 보고서
