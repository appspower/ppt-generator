# Phase A2 Stage 1 + Stage B1 실행 보고

> 작성일: 2026-04-24 (v2, Track 2 방향 최종 확정)
> 실행: `ppt_builder/catalog/extract_meta.py` + `detect_decks.py` (폐기)
> 산출: `output/catalog/slide_meta.json`, `output/catalog/skeletons.json` (신규)

---

## 1. Stage 1 결과 — 슬라이드 메타데이터 추출

**완료: 1,251장 전부 파싱 성공**. Phase A1 측정치와 **완벽 일치**.

| 지표 | Phase A1 | Stage 1 측정 | 일치 |
|---|---|---|---|
| 총 슬라이드 | 1,251 | 1,251 | ✅ |
| `~~` 보유 슬라이드 | 1,235 (98.7%) | 1,235 (98.7%) | ✅ |
| `~~` 총 개수 | 49,925 | 49,846 | ✅ (99.8%) |
| 평균 `~~`/슬라이드 | 39.9 | 39.8 | ✅ |

### 추가 측정치

| 지표 | 값 |
|---|---|
| SimHash unique groups | 3 중복 group, 7 슬라이드 (0.56%) |
| 구조 시그니처 unique | 937 / 1,251 |
| Leaf shape 평균 | 63.3개 (max 2,676 — #189 아웃라이어) |
| 페이지 번호 보유 | 943 / 1,251 (75.4%) |
| 제목(상단 큰 폰트) 추출 | 977 / 1,251 (78.1%) |
| KR-heavy (한글 5자+) | 21 |
| EN-heavy (영문 20자+, 한글 ≤5) | 11 |
| Empty text (`~~`만) | 1,121 — 템플릿 특성상 정상 |

### 판단
- 구조 시그니처 937 unique → 클러스터링으로 60~120 layout으로 압축 여지 큼
- 텍스트 중복 극히 적음, dedup 효과 미미

---

## 2. Stage B1 결과 — **폐기** (가정 틀림)

### 2.1 초기 실행 결과
- 117개 경계 감지, min=3, max=275, avg=8.6장
- 목표(30~50 덱) 대비 과다

### 2.2 가정 재검증 → 완전 기각

**결정적 증거 3가지** (2026-04-24 데이터 분석):

| 증거 | 값 | 해석 |
|---|---|---|
| `page_number_text == "1"` 슬라이드 수 | **1** (전체 1,251 중) | 페이지 번호 리셋 없음 → 덱 시작 점 없음 |
| 유일한 페이지 번호 (1~942 연속) | 942개 고유 값 | **단일 연속 덱**의 증거 |
| 단일 layout `1_Title and Full Content` 차지 | **953 / 1,251 (76%)** | 덱별 layout 변화 없음 |
| Stage B1 경계 중 layout 힌트 매칭 | **116/117** | 전부 동일 layout의 부분 문자열 오탐 |

**결론**: 마스터 템플릿 1,251장은 **여러 완성 덱의 합본이 아니라 단일 카탈로그**.
따라서 Stage B1 (경계 감지) + Stage B2-B3 (LCS 스켈레톤 추출) **전부 전제 실패**.

### 2.3 산출물 처리
- `output/catalog/deck_boundaries.json` — 무효 처리, 파일만 유지 (역사 기록)
- `ppt_builder/catalog/detect_decks.py` — 파일 상단에 **DEPRECATED** 주석 추가
- 다음 실행에서는 호출하지 않음

---

## 3. Track 2 방향 — **Option C 확정 (이론 기반 하드코딩)**

### 3.1 결정 근거

| 옵션 | 평가 | 채택? |
|---|---|---|
| A — `docs/references/완성본/` JPG 스캔 | 원본 PPTX 없음, narrative role 추출 불가 | ❌ |
| B — HJ 제안서 선별 179장 | 패턴별 카탈로그, 덱 단위 아님 | ❌ |
| **C — SCQA/Minto 이론 기반 하드코딩** | 2025 SOTA rule-based gate 권고와 정합 | ✅ |

### 3.2 산출물: `ppt_builder/catalog/skeletons.py`

7개 표준 컨설팅 스켈레톤 encoding:

| skeleton_id | 슬라이드 | 유스케이스 |
|---|---|---|
| `consulting_proposal_30` | 30 (25-40) | 제안서, 프로젝트 RFP 응답 |
| `analysis_report_15` | 15 (12-20) | 시장 분석, 재무 리뷰, Q1/Q2 리뷰 |
| `transformation_roadmap_10` | 10 (8-14) | 넷제로 로드맵, 디지털 전환 |
| `executive_strategy_40` | 40 (35-50) | 연간 전략, 이사회 자료 |
| `change_management_20` | 20 (16-25) | 조직 개편, M&A 통합 |
| `progress_update_10` | 10 (6-14) | 월간/분기 업데이트 |
| `short_pitch_8` | 8 (5-10) | 임원 요약, 투자 설명 |

API:
```python
from ppt_builder.catalog import SKELETONS, get_skeleton, recommend_skeleton

sk = recommend_skeleton("넷제로 전환 로드맵", target_slides=10)
# → transformation_roadmap_10
```

`narrative_sequence: list[NarrativeRole]`는 PPT 생성 Step 1 (Outline)에서
각 슬라이드의 역할을 확정하는 **Outline-First Planning의 spine**.

### 3.3 5개 벤치마크 ↔ 스켈레톤 매핑 (확정)

| # | 벤치마크 시나리오 | 스켈레톤 |
|---|---|---|
| 1 | 넷제로 전환 로드맵 10장 | `transformation_roadmap_10` |
| 2 | SAP ERP 전환 제안 30장 | `consulting_proposal_30` |
| 3 | 2026 Q1 재무 분석 15장 | `analysis_report_15` |
| 4 | 조직 개편 Change Mgmt 20장 | `change_management_20` |
| 5 | 연 1회 전략 리뷰 40장 | `executive_strategy_40` |

---

## 4. 인프라 상태

| 항목 | 상태 |
|---|---|
| `ppt_builder/catalog/schemas.py` (Pydantic 6 models) | ✅ |
| `ppt_builder/catalog/extract_meta.py` (Stage 1) | ✅ 실행 검증 |
| `ppt_builder/catalog/detect_decks.py` (Stage B1) | ⚠️ DEPRECATED |
| `ppt_builder/catalog/skeletons.py` (Track 2) | ✅ **신규** — 7개 스켈레톤 |
| `ppt_builder/catalog/embed_cluster.py` (Stage 2) | ✅ **신규** — 실행 중 |
| `output/catalog/slide_meta.json` (1,251 entries) | ✅ |
| `output/catalog/deck_boundaries.json` (117 entries) | ⚠️ 무효 (역사 기록) |
| `output/catalog/skeletons.json` (7 skeletons) | ✅ **신규** |
| `output/catalog/clusters.json` | 🔄 Stage 2 진행 중 |

---

## 5. 다음 단계

### 진행 중 — Stage 2 (Track 1)
- BGE-M3 로드 (`BAAI/bge-m3`, ~2GB)
- 1,251 슬라이드 임베딩 (text + layout + structure_sig)
- HDBSCAN min_cluster_size=5
- 목표 클러스터 수: 60~120
- 출력: `clusters.json`, `embeddings.npy`, `cluster_labels.json`

### Stage 3 (Track 1) — 다음 세션 이후
- 클러스터 대표 슬라이드 PNG 추출 (PowerPoint COM)
- 20 Agent 병렬 태깅 (`narrative_role`, `intent`, `visual`)
- 출력: `clusters_tagged.json`

### Stage 4 (Track 1)
- 클러스터별 max_chars 95 percentile 측정
- `content_schema` 생성 (Pydantic `SlotSchema`)
- 56점 실패 #1 (overflow) 해결

### Track 3 — `evaluate_deck.py`
- `scqa_structure` + `narrative_arc` + `layout_variety` + `title_chain` + `keyword_persistence` 규칙 기반
- `vlm_flow_score` 1-5 (signal만)

---

## 6. 결정 이력

| 날짜 | 결정 | 근거 |
|---|---|---|
| 2026-04-24 | Stage B1 폐기 | 페이지 번호/레이아웃 분포 증거 3종 |
| 2026-04-24 | Track 2 Option C 채택 | SOTA rule-based 권고 + 데이터 전제 실패 |
| 2026-04-24 | 7개 표준 스켈레톤 하드코딩 | SCQA + Minto Pyramid 기반 |
| 2026-04-24 | Stage 2 병렬 시작 | 클러스터링은 덱 구조와 독립 |
