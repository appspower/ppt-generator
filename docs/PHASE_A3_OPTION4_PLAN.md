# Phase A3 옵션 4 상세 실행 계획 (2026-04-25 작성)

> 목적: 80+점 컨설팅 그레이드 PPT 자동 생성 달성
> 방법: 측정 기반 단계별 진행 (옵션 4) + 다중 위험 상쇄
> 작성: 현재 세션 (분석 컨텍스트 살아있는 상태)
> 실행: 새 세션 권장

---

## 1. 데이터/상황 재점검 (실행 시작 전 확인)

### 1.1 우리가 가진 자산
| 자산 | 상태 | 용도 |
|---|---|---|
| 1,251장 마스터 PPT (`docs/references/_master_templates/PPT 템플릿.pptx`) | ✅ | 원본 |
| `final_labels.json` (1,251 multi-label, slide-level) | ✅ | retrieval 키 |
| `slot_schemas.json` (73,425 슬롯 max_chars) | ✅ | overflow 방지 |
| `skeletons.json` (7 narrative skeleton) | ✅ | 덱 구조 |
| `all_pngs/` (1,251 PNG, 89MB) | ✅ | 시각 검증 |
| Phase D edit_ops (5 API) | ✅ | 슬라이드 편집 |
| `cloner.py` whole-slide clone | ✅ | Mode A 코어 |
| `assembler/` 9 renderer | ✅ | Mode B fallback |
| 시뮬레이션 인프라 (`simulate_mode_a.py`) | ✅ | 자동 측정 |

### 1.2 알려진 한계 (실측)
1. **slide-level만 라벨, paragraph-level 메타 없음** ← Step 2 핵심 해결 대상
2. **narrative_role 분포 편중**:
   ```
   풍부:  analysis 1143 / recommendation 640 / evidence 452 (91%)
   결손:  opening 8 / closing 1 / situation 1 / risk 1 / benefit 3
          → 누적 16장 (2~3 덱 만들면 풀 고갈)
   ```
3. **Mode A 단독 시뮬레이션 점수: 52/100** (Mode A 인프라 검증 시뮬, 넷제로 10장)
4. **357 group 시그니처 분산** → Mode B 자동 라이브러리화 ROI 낮음
5. **마스터에 cover 디자인 부재** (있는 8장은 거의 빈 슬라이드)

### 1.3 우리가 만들 것 (Mode C Hybrid)
```
사용자 input → Outline-First Planning (스켈레톤 매칭)
            → Mode A retrieval (final_labels + slot_schemas)
            → 적합 시 whole-slide clone + paragraph 채움
            → 부적합 시 Mode B (assembler) fallback
            → PPTEval 3축 + visual loop 검증
            → 80+점 달성
```

---

## 2. 옵션 4 상세 단계 (5 Step + Stop-gates)

### Step 1 — 5 벤치마크 정밀 측정 (1일, 사용자 시간 30분)

**목적**: 데이터 기반으로 진짜 부족한 부분 정량화

**5 시나리오** (메모리 `project_phase_a2_plan`과 일치):
1. 넷제로 전환 로드맵 10장 — `transformation_roadmap_10`
2. SAP ERP 전환 제안 30장 — `consulting_proposal_30`
3. 2026 Q1 재무 분석 15장 — `analysis_report_15`
4. 조직 개편 Change Mgmt 20장 — `change_management_20`
5. 연간 전략 리뷰 40장 — `executive_strategy_40`

**측정 항목** (각 시나리오마다):
- **A. role 매칭 성공률** — 스켈레톤 N개 role 중 매칭된 슬라이드 수
- **B. 슬롯 채움률** — 슬라이드당 ~~ 중 의미있게 채워진 비율 (현재 ~5%)
- **C. 오버플로 빈도** — 텍스트 잘림 발생한 슬라이드/전체
- **D. PPTEval 3축 점수** — Content / Design / Coherence (Claude judge)
- **E. 시각 검증** — 사용자 또는 Claude가 PNG 검수
- **F. role 풀 고갈** — 같은 슬라이드 N회 이상 사용한 role 추적

**Stop-gate**: 5 시나리오 평균 점수
- 60점 이상: 정상 진행 (Step 2)
- 50~60점: A3 + 추가 보완 필요 (전체 옵션 4)
- 50점 미만: **방향 재고** (Mode B 강화 or 추가 자산 캡처 결정)

**산출물**:
- `output/benchmark/scenario_*.pptx` × 5
- `output/benchmark/scenario_*_pngs/` × 5
- `output/benchmark/scoreboard.json` (정량 측정 통합)
- 보고서: 어디가 가장 큰 bottleneck인지

**시간**: 자동 6시간 + 사용자 검토 30분

---

### Step 2 — A3 paragraph 카탈로그 (Bottleneck 우선, 4~8시간)

**목적**: 슬라이드의 각 ~~를 의미 단위로 라벨링

**전제**: Step 1에서 어느 슬롯 타입이 가장 자주 fail하는지 확인 후 우선순위 결정

**확장 스키마** (per slide):
```python
class ParagraphSlot:
    slide_index: int
    flat_idx: int  # iter_leaf_shapes 인덱스
    paragraph_id: int
    
    role: Literal[
        "title", "subtitle", "kicker",
        "card_header", "card_body", "card_kpi",
        "table_header", "table_cell",
        "chevron_label", "phase_label",
        "callout_text", "footer", "page_number",
        "axis_label", "data_label",
        "icon_caption", "decorative",
    ]
    max_chars: int  # bbox area × 200 chars/in² × 0.95 (이미 있음)
    position_in_group: int  # 의미 단위 매핑
    alignment: "left" | "center" | "right"
    font_size_estimate: float
    is_critical: bool  # title/header 등 필수 vs 보조
```

**우선순위 (Step 1 데이터 기반 조정)**:
- **High priority** (모든 슬라이드 필수): title, card_header, table_header, chevron_label
- **Medium**: card_body, table_cell, callout_text
- **Low**: footer, decorative

**작업 방식**:
- 1,251장 × 평균 40 슬롯 = ~50,000 슬롯
- 그 중 high priority만 ~15,000 슬롯 (Step 2a)
- medium은 다음 라운드 (Step 2b, 필요 시)
- Agent 병렬 (8 Agent × 약 25 라운드 = 약 5~6시간)

**Stop-gate**: 정확도 검수
- High priority 슬롯 5% 무작위 샘플 검수
- 정확도 < 80% 시 → 프롬프트 개선 후 재실행

**산출물**:
- `output/catalog/paragraph_slots.json` (50,000 슬롯)
- `ppt_builder/catalog/slot_schema.py` (Pydantic 스키마)

---

### Step 3 — Step 2 후 5 벤치마크 재측정 (1시간)

**목적**: A3 효과 정량 측정

**작업**:
- Step 1과 동일한 5 시나리오 재실행
- paragraph 카탈로그 활용해 슬롯 정밀 매칭
- 점수 비교 (Step 1 → Step 3)

**Stop-gate**: 평균 점수 변화
- 80점 이상 도달: **목표 달성! Step 5만 진행** (Step 4 skip)
- 70~80점: Step 4 (A4/A5) 진행
- 60~70점: Step 4 + 추가 측정으로 잔존 bottleneck 분석
- 60점 미만: **방향 재고** (Mode B 강화 또는 자산 추가 캡처)

---

### Step 4 — 잔존 부족분 보완 (조건부, 1~2일)

**A4 — Role 결손 보완** (Step 3에서 role 매칭 실패율 ≥20% 시)
- 선택지 A: 사용자가 PwC 실제 opening/closing/cover 슬라이드 5~10장 추가 캡처
- 선택지 B: assembler 9 renderer 강화 (코드 fallback)
- 선택지 C: 양쪽 병행

**A5 — Overflow 처리** (Step 3에서 오버플로 ≥10% 시)
- 동적 폰트 축소 (0.95x 스텝)
- 자동 줄바꿈
- max_chars 초과 시 다른 슬라이드 후보로 fallback
- 이미 `evaluate.py`에 overflow 검사 있으니 활용

**A6 — Visual Loop** (Step 3에서 design 점수 < 4.0 시)
- 사용자 메모리 `feedback_visual_check` 원칙 자동화
- 각 슬라이드 PNG → Claude vision 검증 → 결함 자동 보고

**선택적 진행**: Step 3 데이터에 따라 A4/A5/A6 중 필요한 것만

---

### Step 5 — 최종 5 벤치마크 + 사용자 검토 (0.5일)

**작업**:
- 5 시나리오 최종 실행
- 각 시나리오별 대표 PNG 5장씩 = 25장 사용자 시각 검수
- PPTEval 3축 최종 점수
- 5 시나리오 모두 80+ 확인

**Stop-gate**: 80+점 달성 여부
- Yes: Phase A3 완료, 다음 Phase로
- No: REFINE 라운드 (가장 점수 낮은 시나리오부터)
- 사용자 시각 검수에서 "안 됨" 판정 시: 추가 분석 필요

---

## 3. 위험 상쇄 방안 (옵션 4 약점 보완)

### 위험 1: Step 1에서 평균 50점 미만 → A3+α도 80점 미달
**상쇄**:
- Step 1 측정 시 **bottleneck 정량화**까지 함께 (어디가 부족한지 정확히)
- Step 1 후 Stop-gate에서 **방향 재고** 명시 — 단순 진행 금지
- 50점 미만 시: Mode B (assembler 강화) 또는 사용자 자산 추가 캡처 결정

### 위험 2: A3 paragraph 라벨링 정확도 < 80%
**상쇄**:
- 5% 무작위 샘플 사용자 검수 (15분)
- Agent 프롬프트 개선 후 재실행 가능
- High priority 슬롯만 우선 (전수 라벨링은 결정 후)

### 위험 3: role 결손 슬라이드 코드 fallback 품질 낮음
**상쇄**:
- Step 4 선택지 A (사용자 캡처) 우선 권장
- 코드 fallback은 Tier C (참고용)으로만 사용
- assembler 9 renderer는 이미 있으나 PwC 톤 100% 일치 어려움 — 명시

### 위험 4: 5 시나리오 동시 80+ 어려움 (시나리오마다 부족 영역 다름)
**상쇄**:
- 시나리오별 **독립 점수** 추적
- 80점 미달 시나리오만 추가 REFINE
- 일부 시나리오는 80+이고 일부는 70대인 경우, 평균이 아니라 최저 시나리오 기준 판단

### 위험 5: 시간 초과 (3~5일 → 1주+)
**상쇄**:
- 각 Step 후 시간 체크포인트
- Step 4 중 시간 부족 시: A4/A5 중 효과 큰 하나만 진행
- 사용자에게 매 단계 보고하여 중단/계속 결정 가능

### 위험 6: 사용자가 시각 검수에서 "디자인 별로"라 판정
**상쇄**:
- Step 1 시뮬레이션 결과를 **Step 2 시작 전 사용자 검토** 필수
- 디자인 만족도 낮으면 → Mode B 코드 fallback 비중 증가 결정
- 매 단계마다 사용자 sign-off

### 위험 7: paragraph 라벨링 후 retrieval API 미구현
**상쇄**:
- Step 2에서 카탈로그뿐 아니라 `ppt_builder/catalog/query.py` 도 함께 작성
- Step 3 재측정 시 query.py 직접 사용

---

## 4. 새 세션에서 시작하는 방법

### 새 세션 시작 시 필수 로드
1. **자동 로드되는 메모리**:
   - `MEMORY.md` 인덱스
   - 핵심 메모리 4종:
     - `project_phase_a2_final.md` (라벨링 결과)
     - `project_phase_a2_plan.md` (3-Track 계획)
     - `project_netzero_56_diagnosis.md` (실패 모드)
     - `feedback_visual_check.md` (시각 검증 원칙)

2. **수동 읽어야 할 파일**:
   - `docs/PHASE_A3_OPTION4_PLAN.md` (이 파일)
   - `docs/PHASE_A2_DIRECTION_ANALYSIS.md` (3 안 분석)
   - `docs/PHASE_A2_FINAL_LABELING.md` (라벨링 결과)
   - `output/simulation/netzero_modeA_report.json` (실측 시뮬)
   - `scripts/simulate_mode_a.py` (시뮬 코드, Step 1에서 5 시나리오로 확장)

3. **첫 명령** (사용자가 새 세션에서 입력 권장):
   ```
   docs/PHASE_A3_OPTION4_PLAN.md 의 옵션 4를 진행. 
   먼저 Step 1 (5 벤치마크 시뮬레이션) 시작. 
   완료 후 결과 보고 + Stop-gate 판단 요청.
   ```

### Git 상태 (실행 시작 전)
- 마지막 커밋: [dbe8bfb](https://github.com/appspower/ppt-generator/commit/dbe8bfb) "방향 분석"
- 다음 commit 예정: 본 계획서 + 메모리 업데이트

---

## 5. 시간/노력 최종 견적

| Step | 작업 | 자동 시간 | 사용자 시간 |
|---|---|---|---|
| 1 | 5 벤치마크 정밀 측정 | 6h | 30분 (검토) |
| 2 | A3 paragraph 카탈로그 | 4~8h | 15분 (샘플 검수) |
| 3 | Step 2 후 재측정 | 1h | 15분 (점수 보고) |
| 4 | 잔존 보완 (조건부) | 8~16h | 30분 (선택지 결정) |
| 5 | 최종 검증 | 4h | 1~2h (시각 검수) |
| **합계** | | **23~35h (3~5일)** | **2.5~3.5시간** |

---

## 6. 80+점 달성 가능성 (정량 추정)

| 시나리오 | 현재 (Mode A) | A3 후 | A4 후 | 최종 |
|---|---|---|---|---|
| 넷제로 10장 | 52 | 70~75 | 78~82 | **80+** |
| SAP 30장 | 추정 60 | 75~80 | 82~85 | **85+** |
| 재무 15장 | 추정 65 | 78~82 | 85~88 | **85+** |
| Change 20장 | 추정 55 | 70~75 | 78~82 | **80+** |
| 전략 40장 | 추정 50 | 65~70 | 75~80 | **78+** |

→ **5 시나리오 평균 80+ 달성 가능**, 단 전략 40장은 가장 어려움 (자산 부족)

---

## 7. 실행 시작 체크리스트

새 세션에서 이 계획대로 진행 시:
- [ ] 메모리 자동 로드 확인
- [ ] `docs/PHASE_A3_OPTION4_PLAN.md` 읽기
- [ ] Step 1 시뮬레이션 코드 확장 (`simulate_mode_a.py` → 5 시나리오)
- [ ] PPTEval 자동 평가 모듈 추가 (선택)
- [ ] 5 시나리오 실행 + 결과 통합
- [ ] **Stop-gate 1: 평균 점수 확인**
- [ ] 사용자에게 Step 1 결과 보고
- [ ] Step 2 진행 결정
- [ ] (반복)

---

## 8. 결정

**옵션 4 (이 계획서)대로 새 세션에서 진행 권장**.

이 세션에서는:
1. 본 계획서 commit + push
2. 메모리 업데이트 (새 세션 시작 시 컨텍스트 회복)
3. 종료 → 새 세션 시작 안내

새 세션에서:
1. 메모리 + 보고서 자동 로드
2. Step 1 즉시 시작
3. Step별 사용자 검토 + 진행
