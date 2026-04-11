# 10. Phase 3: 판단 시스템 설계서

> 42개 컴포넌트(재료)를 Claude Code가 **언제, 어떻게, 왜** 조합할지 결정하는 시스템.
> 3개 리서치 에이전트 + 5개 파일 정밀 갭 분석 기반.

---

## 0. 왜 판단 시스템이 핵심인가

```
컴포넌트 42개 × 레이아웃 15개 = 이론적 조합 630가지
→ 이 중 "컨설팅 품질"이 나오는 조합은 50~80가지
→ Claude Code가 콘텐츠를 보고 그 50~80가지 중 최적을 골라야 한다
→ 판단이 잘못되면 재료가 아무리 좋아도 결과물이 안 좋다
```

---

## 1. 리서치 결과 종합

### 1.1 AI 도구의 레이아웃 결정 로직 (벤치마킹)

| 도구 | 접근법 | 핵심 인사이트 |
|------|-------|------------|
| **Gamma** | 20+ AI 모델 병렬 — 텍스트/이미지/레이아웃/브랜드 각각 별도 모델 | 레이아웃 공간을 이산 카탈로그로 제한, 자유 생성 아님 |
| **Beautiful.ai** | 300+ Smart Slide 셸 — 사용자가 유형 선택, AI가 간격/정렬 적응 | **선택은 사람, 적응은 AI** — 우리는 선택도 AI(Claude) |
| **PPTAgent** (EMNLP 2025) | 2단계: ①레퍼런스 클러스터링 → ②최적 매칭+LLM 편집 | 레이아웃 선택 = **레퍼런스 검색 문제** |
| **Presenton** | 파일명 기반 시맨틱 매칭 (.tsx 파일명이 콘텐츠 슬롯 설명) | 소규모 카탈로그에서 단순하고 효과적 |
| **SlideCoder** (EMNLP 2025) | LLM이 코드 액션을 생성하여 슬라이드 편집 | 우리 Composer 코드 생성과 동일 구조 |

**우리 시스템에 적용할 접근법**: **2단계 하이브리드**
```
Phase A: 규칙 기반 필터링 → 후보 3~5개 좁히기
Phase B: 5요소 스코어링 → 최적 1개 선택
```

### 1.2 컨설턴트의 슬라이드 판단 규칙 7가지

| # | 규칙 | 적용 시점 |
|---|------|---------|
| R1 | **Action Title → Exhibit 카테고리** (주장 유형이 시각 형태 결정) | Step 3 SELECT |
| R2 | **데이터 성격 → 정량 vs 정성** (숫자 3개+ → 차트, 3개 이하 → KPI) | Step 3 SELECT |
| R3 | **논거 강도 → 밀도** (자명 → Hero, 1개 증거 → Standard, 삼각검증 → Dense) | Step 3 SELECT |
| R4 | **단일 vs 복합** (차트 단독으로 제목 증명 → 단일, 맥락 필요 → 복합) | Step 3 SELECT |
| R5 | **제목↔본문 정합** (본문의 모든 요소가 제목의 단어를 증명해야 함) | Step 5 EVALUATE |
| R6 | **메시지→시각 형태 매핑** (성장 → 라인/워터폴, 추천 → 비교표) | Step 3 SELECT |
| R7 | **청중 → 밀도 오버라이드** (C-suite → Hero/Standard, 실무진 → Dense 허용) | Step 2 PLAN |

### 1.3 5요소 스코어링 공식

```python
def score_template(template, content_profile, recent_slides):
    """각 후보 템플릿/레시피에 대해 0~100 점수 산출."""
    return (
        match_type(template.content_types, content.relationship_type) * 40
        + fit_items(template.max_items, content.item_count) * 25
        + density_match(template.density, content.density) * 15
        + variety_bonus(template, recent_slides) * 10
        + avoid_penalty(template.avoid_when, content) * 10
    )
```

---

## 2. 현재 파일별 갭 분석

### 2.1 slide_designer.md — 전면 개편 필요

| 영역 | 현재 | 변경 |
|------|------|------|
| 매칭 테이블 (30+행) | 패턴 이름 참조 (`comparison`, `chevron_process`) | **레이아웃 + 컴포넌트 조합** 참조 |
| 복합 구성 섹션 | JSON 스키마 예시 | **SlideComposer 코드** 예시 |
| 판단 규칙 (7개) | 패턴 선택 중심 | **7가지 컨설팅 판단 규칙 + 5요소 스코어링** |
| 레이아웃 목록 | 8개 (직사각형만) | **15개** (비직사각형 7개 추가) |
| 컴포넌트 호환성 | 없음 | **5카테고리 호환성 매트릭스** |
| 아키타입 제약 | 없음 | **7종 아키타입 + 연속 금지 규칙** |

### 2.2 07_WORKFLOW.md — 구조 업데이트

| 단계 | 현재 | 변경 |
|------|------|------|
| Step 2→3 사이 | 없음 | **Step 2.5 SEQUENCE 삽입** |
| Step 3 SELECT | "template_id 선택" | "layout + component 조합 선택 (스코어링)" |
| Step 4 GENERATE | "JSON 스키마 작성" | "SlideComposer + comp_xxx 코드 생성" |
| Step 5 EVALUATE | 8개 체크 | **11개 체크** (조합 다양성, 공간 활용, 덱 리듬 추가) |

### 2.3 evaluate.py — 3개 체크 추가

| 체크 | 현재 | 추가 |
|------|------|------|
| 조합 다양성 | 시각 카테고리 2종+ (기존) | **컴포넌트 카테고리 5종 중 2종+ 필수** 강화 |
| 공간 활용률 | 하단 빈 공간만 체크 | **전체 shape 면적 / 슬라이드 면적 ≥ 50%** |
| 덱 리듬 | 없음 | **연속 동일 아키타입 감지 + 밀도 교차** |
| 텍스트 밀도 | 최소 800자 | **300~2000자** (차트 중심 완화) |

### 2.4 composer.py RECIPES — 정비

| 영역 | 현재 | 변경 |
|------|------|------|
| 레시피 수 | 12개 | **~20개** (Compound 16개 반영) |
| 컴포넌트 참조 | "kpi_row OR data_card_row" (모호) | **정확한 comp_xxx 이름 + alternatives** |
| 메타데이터 | when만 | **archetype, density, expected_categories 추가** |

---

## 3. 비판적 검증 결과 — 계획 수정

### 3.0 검증에서 나온 5가지 수정 사항

2차 리서치(비판적 검증 + LLM 가이드 포맷)에서 원안의 문제점이 드러남:

| 원안 | 문제 | 수정 |
|------|------|------|
| 5요소 스코어링 공식 | Claude Code는 Python 함수를 실행하지 않고 markdown을 읽음. 공식은 무의미 | **구체적 IF-THEN 규칙 + 5개 예시**로 대체 |
| Step 2.5 SEQUENCE 별도 단계 | PLAN에서 이미 전체 슬라이드를 보고 있음. 별도 단계는 중복 | **PLAN 단계에 리듬 규칙 3개 통합** |
| 별도 COMPOSITION_RECIPES 문서 | 조회 포인트가 2개(guide + recipes)면 혼란 | **매칭 테이블에 조합 열 추가**로 통합 |
| evaluate.py 실행 시점 변경 | python-pptx는 저장 후에야 검증 가능. 중간 실행 불가 | **기존 post-hoc 유지**, 덱 레벨 체크 2개만 추가 |
| 피드백 루프 코드 구현 | 이미 memory/ 폴더에 feedback_*.md 메커니즘이 있음 | **기존 메모리 시스템 활용**, 코드 불필요 |

### 3.1 slide_designer.md의 최적 포맷 (LLM 에이전트 리서치 기반)

**핵심 발견**: LLM은 플랫 테이블 40행보다 **계층형 의사결정 트리 + 구체적 예시**를 훨씬 잘 따른다.

```
최적 구조 (총 ~200줄, 핵심은 800토큰 이내):

§1 HARD RULES — 7개 불변 규칙 (최상단, ~30줄)
   "MUST: Action Title은 완전한 문장"
   "NEVER: 같은 레이아웃 연속 사용"
   "MUST: 한 슬라이드에 2개+ 컴포넌트 카테고리"
   
§2 DECISION TREE — 계층형 분기 (~60줄)
   ## IF 비교 메시지 →
   ### IF 2개 항목 → comp_before_after (full)
   ### IF 3개 항목 → comp_comparison_grid (full)
   ### IF 4+ 항목 → comp_comparison_grid + comp_kpi_row (t_layout)
   
§3 WORKED EXAMPLES — 5개 실전 예시 + 판단 근거 (~80줄)
   "매출 20% 성장을 보여줘야 함"
   → 판단: 변화 메시지, 숫자 1개 핵심, 맥락 필요
   → 선택: l_layout + comp_kpi_card(좌) + comp_native_chart(우상) + comp_bullet_list(우하)
   → 이유: Hero 숫자가 주인공, 차트가 증거, 불릿이 맥락

§4 SELF-CHECK — 4개 검증 질문 (~15줄)
   "□ 본문의 모든 요소가 Action Title을 증명하는가?"
   "□ 2개+ 컴포넌트 카테고리를 사용했는가?"
   "□ 직전 슬라이드와 다른 레이아웃인가?"
   "□ 빈 존이 없는가?"

APPENDIX: 전체 컴포넌트 카탈로그 (참조용)
```

---

## 4. 최종 세부 실행 계획 (수정본)

### 4-1. slide_designer.md 전면 개편 (~200줄)

| 섹션 | 내용 | 근거 |
|------|------|------|
| §1 HARD RULES | 7개 불변 규칙 (최상단 배치) | LLM은 문서 앞 20%와 뒤 10%에 가장 높은 주의 |
| §2 DECISION TREE | 메시지 유형(5종) → 항목수 → 레이아웃+컴포넌트 | 계층형 좁히기가 플랫 테이블보다 LLM 정확도 높음 |
| §3 WORKED EXAMPLES | 5개 실전 예시 (입력→판단→선택→이유) | 5~7개 구체 예시 > 20개 추상 규칙 (LLM 연구 결과) |
| §4 SELF-CHECK | 4개 검증 질문 | 자기 검증으로 오류율 감소 |
| APPENDIX | 42개 컴포넌트 + 15개 레이아웃 카탈로그 | 상세 참조용, 핵심 판단에는 영향 없음 |

### 4-2. 07_WORKFLOW.md 업데이트 (~50줄)

**SEQUENCE 별도 단계 취소** → PLAN에 리듬 규칙 통합:

```
Step 2: PLAN (기존 + 리듬 규칙 추가)
  기존 수행 사항 유지 +
  추가 규칙 3개:
    1. 같은 레이아웃 연속 사용 금지
    2. 고밀도 슬라이드 3장 연속 금지 (중간에 Hero/브레더 삽입)
    3. 10장 중 레이아웃 4종+ 사용

Step 3: SELECT (패턴→컴포넌트 전환)
  기존: "template_id 선택"
  변경: "slide_designer.md의 DECISION TREE 따라 레이아웃 + 컴포넌트 선택"

Step 4: GENERATE (Composer 코드)
  기존: "JSON 스키마 작성"
  변경: "SlideComposer + comp_xxx() 호출 코드 작성"
```

### 4-3. evaluate.py 업그레이드 (~60줄 코드)

**기존 post-hoc 유지**, 덱 레벨 체크 2개 + 임계값 조정:

```python
# 수정:
# - 텍스트 밀도 임계값: 800→300 최소 (차트 중심 완화)

# 추가:
# 10. 공간 활용률 (shape 면적 합 / 슬라이드 면적 ≥ 50%)
# 11. 덱 레벨: 인접 슬라이드 유사도 감지 (같은 template type 연속 경고)
```

### 4-4. composer.py RECIPES 정비 (~60줄)

**별도 문서 안 만듦** — 기존 RECIPES에 메타데이터만 추가:

```python
"kpi_summary_detail": {
    ...기존...
    "archetype": "kpi",          # 아키타입 태그 추가
    "density": "high",           # 밀도 레벨
    "primary_component": "comp_kpi_row",  # 명확한 컴포넌트 참조
},
```

---

## 5. 미해결 질문 (보강 완료)

| 원래 질문 | 답변 |
|---------|------|
| 스코어링 가중치 최적화 | **폐기** — 스코어링 공식 대신 IF-THEN 규칙 + 예시 |
| 아키타입 자동 태깅 | **RECIPES 메타데이터**로 해결 (각 레시피에 archetype 필드) |
| 청중 파라미터 | PLAN 단계에서 Claude가 자연어로 판단 (코드화 불필요) |
| 레퍼런스 검색 방식 | PwC 완성본 대량 입수 후 별도 검토 (지금은 범위 밖) |
| 피드백 루프 | **기존 memory/ 시스템 활용** (feedback_*.md), 코드 불필요 |

---

## 5. 성공 기준

- [ ] slide_designer.md가 패턴 이름 0개, 컴포넌트 조합만으로 구성
- [ ] 워크플로우 7단계 (SEQUENCE 포함) 문서화 완료
- [ ] evaluate.py 신규 체크 3개 통과하는 테스트 덱 생성
- [ ] 넷제로 10장 덱을 새 판단 시스템으로 생성 시 evaluate 80+ 점수
- [ ] 10장 중 아키타입 4종+ 사용, 연속 동일 0건
