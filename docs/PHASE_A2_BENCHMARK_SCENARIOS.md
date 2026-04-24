# Phase A2 벤치마크 시나리오 5종

> 작성일: 2026-04-24
> 목적: Phase A2 완료 후 A/B 검증. 대조군(Track 1만) vs 실험군(Track 1+2+3) 비교.

## 측정 지표

각 시나리오는 아래 3개 축으로 평가:

| 축 | 지표 | 현재 도구 | 목표 |
|---|---|---|---|
| Per-slide | `evaluate.py` 점수 | 기존 | 평균 80+ (PASS threshold) |
| Deck coherence | `evaluate_deck.py` (Track 3) | 신규 | 규칙 기반 ≥0.8, VLM signal ≥3.5/5 |
| Qualitative | 사용자(SAP 컨설턴트) 검수 | 수동 | 5점 척도 ≥4/5 |

A/B 대조: 대조군은 Track 1만, 실험군은 Track 1+2+3. 같은 input brief로 생성.

---

## 시나리오 1: 넷제로 전환 로드맵 (재도전)

**Why**: 기존 56점 정체의 직접 재도전. 가장 중요한 benchmark.

**Input brief**:
> HD현대가 2050 넷제로 달성을 위한 3단계 로드맵을 이사회에 보고. 현재 배출량 현황 → 2030/2040/2050 마일스톤 → 주요 투자 영역 → 기대 감축 효과 → 리스크 → 실행 조직.

- **목표 슬라이드 수**: 10장
- **narrative_sequence**: `[opening, agenda, situation, complication, roadmap×3, benefit, risk, closing]`
- **필요 템플릿 태그**: 
  - `situation` × 1 (현재 배출량 dense table/chart)
  - `roadmap` × 3 (단계별 타임라인)
  - `benefit` × 1 (감축 효과 KPI)
  - `risk` × 1
- **합격선**: evaluate.py 평균 80+ (현재 56), coherence ≥0.8

**Validation**: 기존 코드 기반 생성본과 직접 비교. 같은 브리프.

---

## 시나리오 2: SAP ERP 전환 제안서 (Proposal, 30장)

**Why**: 사용자 실제 업무 도메인. 가장 긴 덱 → 덱 리듬 중요.

**Input brief**:
> 제조업 고객사의 SAP S/4HANA 전환 제안. 현행 Legacy/Oracle 시스템 분석 → 전환 배경 → SAP S/4 솔루션 구조 → 모듈별 roadmap (FI/CO/MM/SD/PP/QM) → 12개월 일정 → 예상 ROI → 유사 사례 → Change Management → 실행 조직 → 다음 단계.

- **목표 슬라이드 수**: 30장
- **narrative_sequence**: 
  ```
  opening → agenda → situation×2 → complication×2 → 
  analysis×5 → recommendation×6 → roadmap×4 → 
  benefit×3 → risk×2 → evidence×2 (유사 사례) → closing×2
  ```
- **필요 템플릿 태그**: proposal 스켈레톤 전체 커버
- **합격선**: evaluate.py 평균 80+, coherence ≥0.85, 사용자 검수 4/5+

**Validation**: 사용자가 실제 PwC 형식으로 납품 가능한 품질인지 판단.

---

## 시나리오 3: 2026 Q1 재무 분석 보고 (Analysis, 15장)

**Why**: Analysis-heavy — 그리드/차트 dense 슬라이드 주력. overflow 실패 모드 테스트.

**Input brief**:
> HD현대 2026 Q1 매출/영업이익 실적. 전년 동기 대비 성장률, 부문별 기여도, 주요 드라이버, 지역별 분포, 원가 구조 변화, 차기 분기 전망.

- **목표 슬라이드 수**: 15장
- **narrative_sequence**: 
  ```
  opening → agenda → situation → evidence×5 → 
  analysis×4 → complication → recommendation → closing
  ```
- **핵심 테스트**: dense table/chart 슬라이드 (evidence×5)의 overflow 방지 + density 컨트롤
- **합격선**: evaluate.py 평균 85+ (per-slide 높게), coherence ≥0.8

**Validation**: #1 실패 모드(overflow) 해결 여부 검증.

---

## 시나리오 4: 조직 개편 Change Management (Recommendation-heavy, 20장)

**Why**: recommendation × 다수 슬롯 — narrative_role 태그 정밀도 검증.

**Input brief**:
> 디지털 전환을 위한 조직 개편 제안. As-Is 조직도 → 변화 필요성 → To-Be 조직 구조 → 3개 핵심 권고안 (각각 2장) → 일정 → 인력 이동 계획 → 의사소통 전략 → 저항 관리 → 성공 지표 → 결론.

- **목표 슬라이드 수**: 20장
- **narrative_sequence**: 
  ```
  opening → agenda → situation×2 → complication×2 → 
  recommendation×6 → roadmap×3 → benefit×2 → 
  risk×2 → closing×2
  ```
- **핵심 테스트**: 연속 `recommendation` 슬롯에 대해 **layout_variety**가 작동하는가 (3장 연속 동일 레이아웃 금지)
- **합격선**: coherence.layout_variety ≥0.8 (evaluate.py:237 STUB 구현 확인)

**Validation**: #5 실패 모드(덱 리듬) 해결 여부 검증.

---

## 시나리오 5: 연 1회 전략 리뷰 (Executive, 40장)

**Why**: 가장 긴 덱 + 최대 complexity. 스켈레톤 활용 + context flow 검증.

**Input brief**:
> HD현대 2026 전사 전략 리뷰. 경영 환경 → 지난해 성과 (부문별) → 도전 과제 → 3개 전략 pillar (각각 성장/운영/혁신, pillar별 4장) → 주요 투자 계획 → 재무 목표 → Risk → 실행 조직 → 다음 단계 → Q&A.

- **목표 슬라이드 수**: 40장
- **narrative_sequence**:
  ```
  opening → agenda → situation×3 → evidence×4 (성과) → 
  complication×2 → recommendation×3 → roadmap×4 (pillar 1) → 
  roadmap×4 (pillar 2) → roadmap×4 (pillar 3) → 
  benefit×3 → risk×2 → closing×3
  ```
- **핵심 테스트**: 
  - 40장에 걸친 narrative arc 일관성 (title_chain 유지)
  - 스켈레톤 matching: `skeletons.json`의 "executive review" 스켈레톤 vs 직접 생성 비교
- **합격선**: coherence ≥0.85, pass@80 rate ≥90% 슬라이드

**Validation**: 전체 파이프라인의 scalability 확인.

---

## 실행 절차

### 단계 1: 환경 준비 (Phase A2 완료 후)
- `catalog.json` 완성 (Track 1 Stage 1-5)
- `skeletons.json` 완성 (Track 2 B1-B3)
- `evaluate_deck.py` 완성 (Track 3)

### 단계 2: A/B 병렬 실행
- 대조군 (Track 1만): 5 시나리오 × 1 run = 5 .pptx
- 실험군 (Track 1+2+3): 5 시나리오 × 1 run = 5 .pptx
- **병렬 실행**: 5 시나리오를 Agent 5개로 동시 실행

### 단계 3: 측정 + 분석
- 각 .pptx → evaluate.py + evaluate_deck.py 자동 실행
- PowerPoint COM으로 PNG 변환 → VLM 리뷰 signal
- 사용자 검수 (시나리오 2, 4 우선)

### 단계 4: 결과 보고
- 시나리오별 점수 테이블
- 합격/미달 분석
- 실패 시나리오 → 설계 파라미터 튜닝 (narrative_role 세분화, skeleton 추가 등)

---

## 합격 판정 기준

- **부분 성공**: 3/5 시나리오가 합격선 통과
- **전체 성공**: 5/5 시나리오가 합격선 통과
- **실패**: 2/5 이하 → Phase A2 설계 근본 재검토

**최소 보증**: 시나리오 1 (넷제로)이 56 → 80+로 올라가야 함. 이것도 실패 시 전체 롤백 대상.
