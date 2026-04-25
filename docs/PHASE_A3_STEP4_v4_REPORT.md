# Phase A3 Step 4 v4 — Content Expand + Slide Diversity (2026-04-25)

> 사용자 요청 #1 (컨텐츠 expand) + #5 (slide diversity) 적용 결과
> 점수보다 시각 품질 변화에 초점

---

## 1. 적용한 변경

### #1 — 컨텐츠 auto-split + slot fill
- 입력 1개 항목 (예: "전기로 전환 + 수소환원제철 + 그린전력 PPA") → 분리 (` / `, ` + `, `, ` 기준) → 슬롯 수에 맞춤
- 슬롯 수보다 입력 항목이 적으면 split 시도, 없으면 그대로 (가짜 컨텐츠 추가 안 함)
- 적용 위치: `_expand_to_capacity()` in `benchmark_5_scenarios_v2.py`

### #5 — Slide diversity (diversity-aware retrieval)
- top-30 후보 중 종합 점수 평가:
  - 0.4 × role confidence
  - 0.4 × capacity fitness (target_n vs primary slot capacity)
  - 0.2 × narrow penalty (max_chars < 20 회피)
  - 0.1 × signature 다양성 보너스 (이미 쓴 macro:archetype 회피)
- 같은 시나리오 내에서 같은 디자인 반복 회피
- 적용 위치: `select_deck_diverse()`

---

## 2. 시각 검증 결과 (대표 슬라이드)

### ✅ 큰 win — transformation_roadmap step 5 (recommendation)
- v3: 5 chevron 중 1개만 채움 ("전기로 전환 + 수소환원제철 + 그린전력 PPA" 한 박스 wrap)
- **v4: 3 chevron 정확 분배** (전기로 전환 / 수소환원제철 / 그린전력 PPA), 4-5번 chevron 빈 dark
- → split 동작 명확히 확인

### ✅ Diversity win — transformation_roadmap step 6 (roadmap)
- v3: 항상 chevron 슬라이드 #176 사용
- v4: 4-column phase roadmap 다른 슬라이드 사용 (Phase 1, 2 헤더 + 빈 phase 3, 4)
- → 시나리오 내 디자인 다양화

### ✅ Diversity win — executive_strategy step 16 (recommendation)
- 8개 pillar가 PwC+Palantir 협업 슬라이드의 카드들로 분산
- Vision + Pillar 1~7 깔끔히 채움 + 미국/영국/인도 국기 데코 유지

### ⚠️ 일부 실패 — consulting_proposal step 13
- diversity가 narrow row 슬라이드 선택 → SAP 6 recommendation이 horizontal wrap으로 깨짐
- 폰트 축소 + truncate가 narrow 슬롯에서 부작용
- narrow 페널티는 적용했으나 충분히 강하지 않음

### ⚠️ 일부 실패 — executive_strategy step 18
- 8 pillar가 chevron 17개 슬라이드에 들어갔는데 텍스트 overlap
- 같은 슬라이드 디자인의 다중 group이 같은 pillar를 받음

---

## 3. 점수 (정량 — 직접 비교 한계 있음)

| 시나리오 | v3 (Step 4 baseline) | **v4 (현재)** | Δ | v4 fill % |
|---|---|---|---|---|
| transformation_roadmap_10 | 64.4 | **68.8** | +4.4 | 16.1% (v3 9.9%) |
| consulting_proposal_30 | 48.7 | **60.2** | +11.5 | 22.6% (v3 8.1%) |
| analysis_report_15 | 50.0 | **53.3** | +3.3 | 13.4% (v3 6.1%) |
| change_management_20 | 54.9 | **50.8** | -4.1 | 14.0% (v3 9.5%) |
| executive_strategy_40 | 46.2 | **59.0** | +12.8 | 23.3% (v3 8.4%) |
| **평균** | **52.8** | **58.4** | **+5.6** | **17.9% (v3 8.4%)** |

### 주의: fill % 두 배로 증가 (8.4% → 17.9%)
- 컨텐츠 expand 작동 명확히 확인
- visual_resolution이 떨어진 케이스는 diversity가 새 슬라이드 선택해서 fillable 분포 다양화 (분모 변동)

---

## 4. 정직한 평가

### 시각 품질 — 명확히 개선
- 일부 슬라이드는 v3보다 훨씬 깨끗 (transformation step 5/6, executive step 16)
- Slide diversity로 같은 디자인 반복 줄어들어 시나리오 내 시각 다양성 ↑
- 컨텐츠 expand로 chevron framework이 제 역할 함 (1개만 채우는 어색함 사라짐)

### 점수 — 작은 개선 + 시나리오 편차
- 평균 +5.6점 (52.8 → 58.4)
- 큰 win: consulting_proposal +11.5, executive_strategy +12.8 (8 pillar 같은 풍부 컨텐츠가 expand 효과 봄)
- 중립/부정: change_management -4.1 (diversity가 부적합 슬라이드 선택)

### 자동만으로 70+ 가능?
**아직은 No.** 컨텐츠 expand 효과 큰 시나리오는 60대 도달, 적은 시나리오는 50대.
70+점 도달하려면 다음 중 하나 필요:
1. **사용자 자산 캡처** (Track 2) — cover/closing/divider 자산 보강
2. **Vision-based label 검증** — narrow 슬롯 함정 자동 회피 (1~2일)
3. **Mode B fallback** — assembler로 결손 role 코드 생성 (2~3일)

---

## 5. 다음 단계 옵션

### 옵션 P1 — Vision-based 라벨 검증 (자동, 1~2일)
- Agent 8개 병렬로 1,251장 PNG vision 검수
- 슬롯의 실제 모양/사용 가능성 판단 (narrow 함정 자동 감지)
- 결과: capacity-aware retrieval 안전하게 활성화 가능
- 예상 +5점

### 옵션 P2 — Mode B fallback (자동, 2~3일)
- assembler/renderers/ 9개 활용해 cover/closing/divider 코드 생성
- PwC톤 흉내, 70%+ 품질 추정
- 예상 +5~8점

### 옵션 P3 — Track 2 자산 캡처 (사용자 1~2시간)
- PwC 슬라이드 5~10장 추가 캡처 → 마스터 풀 보강
- 가이드: docs/PHASE_A4_USER_CAPTURE_GUIDE.md
- 예상 +5~10점

### 옵션 P4 — 현 상태 + 실 사용 테스트
- 점수보다 실제 컨설팅 업무에 사용해보고 평가
- 사용자가 직접 컨텐츠 입력 → 생성 → 결과 검수
- 미진한 부분 case-by-case로 보완

---

## 6. 권장

데이터가 보여주는 것:
- **자동 작업 한계**: 컨텐츠 expand + diversity 조합으로 평균 58점, 일부 시나리오 60대
- **80+ 도달**: 사용자 자산 캡처 + 추가 자동 작업 둘 다 필요
- **단계적 접근**: P1 (Vision) → P3 (캡처) → 측정

**가장 효율적**: P1 (Vision 검증, 1~2일) 또는 P3 (캡처, 사용자 1~2시간) 중 선택

어느 쪽으로 진행할지 결정 요청드립니다.

---

## 7. 산출

- `scripts/benchmark_5_scenarios_v2.py` — content expand + diversity 추가
- `output/benchmark_v2/{scenario}/deck.pptx` × 5 — 새 결과
- `output/benchmark_v2/{scenario}/pngs/` × 5 — 시각 검증용
- `docs/PHASE_A3_STEP4_v4_REPORT.md` — 본 보고서
