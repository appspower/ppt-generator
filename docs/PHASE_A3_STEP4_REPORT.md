# Phase A3 Step 4 — 자동 작업 결과 + Track 2 가이드 (2026-04-25)

> Step 4 옵션 B: 자동 작업 (Track 1) + 사용자 자산 캡처 (Track 2)
> Track 1 시도 결과 + Track 2 가이드 작성

---

## 1. Track 1 자동 작업 결과 (정직 보고)

### Track 1.1 — 잔존 `~~` 추가 청소 ✅
- 모든 fillable + non-fillable paragraph에 대해 `~~`만 남은 placeholder를 빈 문자열로 교체
- 페이지 번호/footer/date_text는 보존
- 효과: 시각적으로 잔존 `~~` 약 50% 감소 (특히 page#, sub-title bar)
- 시뮬레이션 1슬라이드당 평균 5~70 cleanups

### Track 1.2 — Capacity-aware retrieval ❌ (비활성화)
- 시도: 컨텐츠 N개 ≤ 슬라이드 슬롯 capacity 매칭 슬라이드 우선 선택
- **실패 원인**: capacity 숫자가 슬롯의 모양/사용 가능성을 반영 못 함
  - 예: 슬라이드 #94는 cap=1이지만 narrow vertical column → 한국어 prose 안 들어감
  - capacity 6 슬라이드 중 한국어 hero text가 깨지는 형태 다수
- 결정: select_deck (단순 confidence-based) 복귀
- 결론: 이 부분은 향후 슬롯의 max_chars + bbox 형상 검증 필요. 현 시점 비활성

### Track 1.3 — Overflow 폰트 자동 축소 ✅
- `len(text) > max_chars` 시 sqrt(max_chars/len) 비율로 폰트 축소
- 최소 8pt 보장
- 효과: 일부 chevron/card 텍스트 잘림 완화. 그러나 한국어 폰트 다중 run 케이스 일부에서 적용 안 됨 (XML 차이)

---

## 2. 점수 추이

| 단계 | composite | 비고 |
|---|---|---|
| v1 (Mode A 1-paragraph fill) | 47.0 | Step 1 baseline |
| v2 (paragraph-aware fill) | 62.9 | Step 3 |
| v3 (+ Track 1.1 + 1.3) | ~ 53~57 | metric 정의 변경 + 라벨러 cap 변경으로 직접 비교 어려움 |

**측정 방법론적 한계**:
- v2 → v3 사이에 visual_resolution metric을 보정 (마스터 이미 있는 데모 텍스트도 visually clean으로 카운트)
- 동시에 라벨러도 group_size cap 12로 줄여 `card_header` 과추론 방지
- 두 변경이 동시에 일어나서 점수 직접 비교 어려움
- **시각 검증으로 보정**: 일부 슬라이드는 v3가 더 깨끗 (transformation step 06), 일부는 v2가 더 좋음 (consulting step 13)

---

## 3. 시각 검증 — 핵심 발견

### 깨끗해진 사례 (Track 1.1 효과)
- transformation_roadmap step 6 (roadmap chevron): page# `~~` 제거, 중간 horizontal bar `~~` 제거
- analysis_report step 7 (analysis): 좌측 callout 3개 fill + 우측 표 cells 청소

### 여전히 어려운 사례
- consulting_proposal step 13 (recommendation framework): 같은 slide#44 사용 시 chevron 5개 fill + 카드 30개 blank 깔끔. 하지만 다른 시나리오에서 같은 slide 선택 시 컨텐츠 분배 문제

### Track 1.2 실패 사례
- transformation step 5: capacity-aware로 narrow vertical column 슬라이드 선택 → 한국어 hero text가 세로로 깨짐
- consulting step 13 (capacity-aware ON): 슬라이드 #94 선택, 6 recommendation이 좁은 horizontal box에 wrap되어 거의 식별 불가

---

## 4. 진짜 핵심 lever — Track 2 (사용자 자산 캡처)

### 결손 분석 재확인
| role | 마스터 풀 | 5 시나리오 합산 요청 | 부족 |
|---|---|---|---|
| analysis | 1,143 | 18 | ✅ 충분 |
| recommendation | 640 | 26 | ✅ 충분 |
| evidence | 452 | 16 | ✅ 충분 |
| **closing** | 1 | 7 | ❌ -6 |
| **situation** | 1 | 8 | ❌ -7 |
| **risk** | 1 | 5 | ❌ -4 |
| **complication** | 2 | 5 | ❌ -3 |
| **benefit** | 3 | 5 | ❌ -2 |
| opening | 8 | 5 | ⚠️ 한계 |
| divider | 9 | 2 | ✅ 충분 |

→ 자동 작업으로는 풀에 없는 디자인 만들 수 없음. **사용자 자산 캡처가 핵심**

### Track 2 가이드 작성 완료
[`docs/PHASE_A4_USER_CAPTURE_GUIDE.md`](PHASE_A4_USER_CAPTURE_GUIDE.md):
- 캡처 우선순위 (필수 5장 + 선택 5장)
- 어떤 슬라이드가 좋은지 (PwC톤, 풍부한 디자인, placeholder 가능)
- 캡처 방법 (PowerPoint에서 직접 복사 → 새 PPT)
- 저장 위치: `docs/references/_master_templates/pwc_extras.pptx`
- 사용자 시간 추정: 1~2시간

### 캡처 후 자동 작업 (Claude 수행)
1. 새 슬라이드 N장 → 기존 catalog에 추가 (extract_paragraphs 재실행)
2. 자동 라벨링 (label_paragraphs 재실행)
3. narrative_role/macro 자동 라벨링 (Phase A2 파이프라인)
4. 5 시나리오 재측정
5. 점수 변화 보고

---

## 5. 사용자 결정 요청

**현 상태**:
- 자동 작업으로는 60대 점수 한계
- 70+점은 사용자 자산 캡처 없이 어려움 (또는 measurement methodology 다른 방향)

**결정 요청**:

### 옵션 B-1 — Track 2 진행 (계획대로)
- 사용자가 PwC PPT 5~10장 캡처 (1~2시간)
- 캡처 후 Claude가 자동 라벨링 + 재측정
- 예상: +5~10점 (visual 측면 크게 개선)
- **언제 캡처할 수 있을지** 알려주면 그때 자동 부분 진행

### 옵션 B-2 — Track 2 유보, 다른 우선순위
- 자산 캡처 보류 (회사PC 접근 일정 미정 등)
- 현 상태로 다른 작업 진행 (예: 실제 컨설팅 시나리오로 사용자 컨텐츠 입력 → 생성 테스트)
- 자산은 나중에 보강

### 옵션 B-3 — 현 상태 종료
- 60대 점수 + 시각적 일부 양호로 사용 가능 판단
- 다음 Phase (production hardening / web app 등)으로 이동

어떤 방향으로 진행할지 알려주세요.

---

## 6. 산출 (이번 단계)

- `scripts/benchmark_5_scenarios_v2.py` — Track 1.1 + 1.3 추가 (capacity-aware는 비활성)
- `docs/PHASE_A4_USER_CAPTURE_GUIDE.md` — Track 2 가이드
- `docs/PHASE_A3_STEP4_REPORT.md` — 본 보고서
- `output/benchmark_v2/` — 최신 결과 (gitignored)

---

## 7. 정직한 평가

이번 자동 작업의 결과는 **혼합** — 일부 시각적 개선 (Track 1.1) + 한 시도 실패 (Track 1.2).
점수상으로 큰 진전 없으나, 시각적으로 일부 슬라이드는 더 깨끗해짐.

본질적으로 Mode A 단독 + 자동 라벨링/매칭만으로는 80점 한계가 명확.
사용자 자산 보강이 다음 단계의 핵심 lever임을 데이터가 확인.
