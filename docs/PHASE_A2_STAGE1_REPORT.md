# Phase A2 Stage 1 + Stage B1 실행 보고

> 작성일: 2026-04-24
> 실행: `ppt_builder/catalog/extract_meta.py` + `detect_decks.py`
> 산출: `output/catalog/slide_meta.json`, `output/catalog/deck_boundaries.json`

---

## 1. Stage 1 결과 — 슬라이드 메타데이터 추출

**완료: 1,251장 전부 파싱 성공**. Phase A1 측정치와 **완벽 일치**.

| 지표 | Phase A1 | Stage 1 측정 | 일치 |
|---|---|---|---|
| 총 슬라이드 | 1,251 | 1,251 | ✅ |
| `~~` 보유 슬라이드 | 1,235 (98.7%) | 1,235 (98.7%) | ✅ |
| `~~` 총 개수 | 49,925 | 49,846 | ✅ (99.8%) |
| 평균 `~~`/슬라이드 | 39.9 | 39.8 | ✅ |

### 추가 측정치 (Phase A1 이후 신규)

| 지표 | 값 |
|---|---|
| SimHash unique groups | 3 중복 group, 7 슬라이드 (0.56%) |
| 구조 시그니처 unique | 937 (1,251 중) |
| Leaf shape 평균 | 63.3개 (max 2,676 — #189 아웃라이어) |
| 페이지 번호 보유 | 943 / 1,251 (75.4%) |
| 제목(상단 큰 폰트) 추출 | 977 / 1,251 (78.1%) |
| KR-heavy (한글 5자+) | 21 |
| EN-heavy (영문 20자+, 한글 ≤5) | 11 |
| Empty text (`~~`만) | 1,121 — 템플릿 특성상 정상 |

### 판단
- **텍스트 중복 극히 적음** — 1,251장 대부분 unique. dedup 통해 제거할 여지 미미
- **구조 시그니처 937 unique** — 클러스터링으로 60-120 layout으로 압축될 여지 큼
- **페이지 번호 보유율 75%** — Stage B1 경계 감지 신뢰도에 직접 영향

---

## 2. Stage B1 결과 — 덱 경계 감지

**117개 경계 감지 → 117 원본 덱으로 분할 (min=3, max=275, avg=8.6장)**

### 문제점 (튜닝 필요)

1. **경계 과다**: 117개는 너무 많음
   - PwC 제안서는 보통 20-40장 → 예상 덱 수는 30-50개
   - 현재 117개 = 잡음 경계 다수 (60~80개 잘못 감지 추정)

2. **장기 검출 실패 구간 존재**: max=275장
   - 연속 275장에 대해 경계 신호가 전혀 잡히지 않음
   - 원인 추정: 페이지 번호 누락 구간 + layout name 변화 없음

3. **평균 8.6장/덱**: 실제 제안서보다 짧음
   - PwC 제안서 표준 30장 기준 → 3.5배 과다 분할

### 근본 이슈 (재검토 필요)

**가정 재평가**: "1,251장 = 여러 완성 덱의 합본"이라는 전제가 부분적으로만 맞을 가능성.

| 시나리오 | 확률 | 영향 |
|---|---|---|
| 1,251장이 복수 덱 합본 | 중 | Stage B1~B3 작동 가능 |
| 1,251장이 컴포넌트 카탈로그 성격 | 중 | Stage B1~B3 무의미 → 재설계 필요 |
| 혼합 (일부 덱 + 일부 카탈로그) | 높음 | Stage B1 threshold 상향 + 부분만 사용 |

### 튜닝 액션 (다음 세션)

1. **Threshold 상향**: 0.5 → 0.7 (잡음 경계 제거)
2. **페이지 번호 reset 강화**: `curr_pn==1 AND prev_pn>=5` 필수 조건화
3. **opening layout 신호 가중 상향**: 0.3 → 0.5
4. **title 유사도 신호 제거**: 제목이 placeholder(`~~`)인 슬라이드 다수 → 신뢰 불가
5. **수동 검증**: 상위 20개 경계 → PNG 확인으로 정확도 측정

### 대안 (Track 2 재설계 옵션)

만약 1,251장이 실제로는 카탈로그 성격이면:
- **Option A**: `docs/references/완성본/` (115MB) 안의 실제 완성 덱들을 Track 2 소스로 사용
- **Option B**: HJ 제안서 179장 선별본 (기존 `project_hj_proposal_selection`) 을 덱 단위로 재구성
- **Option C**: 스켈레톤을 PwC 표준 컨설팅 프레임워크(SCQA, Minto Pyramid)로 직접 encoding — 데이터 기반 추출 skip

**추천**: Option C + 검증용으로 A/B 병행. 이론적 스켈레톤 5~10개를 먼저 하드코딩하고, 데이터 기반은 보조.

---

## 3. 인프라 검증 상태

| 항목 | 상태 |
|---|---|
| `ppt_builder/catalog/__init__.py` | ✅ |
| `ppt_builder/catalog/schemas.py` (Pydantic 6 models) | ✅ |
| `ppt_builder/catalog/extract_meta.py` (Stage 1) | ✅ 실행 검증 |
| `ppt_builder/catalog/detect_decks.py` (Stage B1) | ✅ 실행됨, 튜닝 필요 |
| `output/catalog/slide_meta.json` (1,251 entries) | ✅ |
| `output/catalog/deck_boundaries.json` (117 entries) | ✅ (v1) |

---

## 4. 다음 단계

### 즉시 (다음 세션 시작 시)
1. **Stage B1 튜닝** (2-3시간)
   - Threshold + 신호 조정
   - PNG 확인으로 정확도 측정
   - 목표: 30-50 덱 (현재 117 → 정리)

2. **Track 2 방향 최종 결정**
   - Data-driven(B1-B3) vs Theory-driven(하드코딩) vs 하이브리드
   - 1시간 이내 결정

### Stage 2 (Track 1, 임베딩 + 클러스터링)
- **병렬 시작 가능** (Stage B1 튜닝과 독립)
- BGE-M3 로드 → 1,251장 임베딩
- HDBSCAN 클러스터링
- 예상: 60-120 클러스터
- 소요: 3-4시간

### Stage 3 (Track 1, 태깅 20 Agent 병렬)
- Stage 2 완료 후 시작
- 클러스터별 대표 + 태그 → `clusters_tagged.json`
- 소요: 2-3일 runtime

---

## 5. 결정 이력

본 Stage 1 보고를 기반으로:
- Stage 1 설계가 적절함을 **데이터로 검증** (Phase A1 수치 99.8% 일치)
- Stage B1은 **1차 설계에서 튜닝 필요** (예상된 반복 단계)
- Track 2 Data-driven 접근의 실효성 재검토 필요 (1시간 이내 결정)
