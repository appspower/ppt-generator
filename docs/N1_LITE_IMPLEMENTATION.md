# N1-Lite 구현 사양 (2026-04-25)

> Mode A + N1-Lite 하이브리드의 N1-Lite 부분 세부 구현
> [PROJECT_DIRECTION.md](PROJECT_DIRECTION.md) 헌법 기반
> 리서치 근거: [_research/component_extraction_research.md](_research/component_extraction_research.md), [_research/composition_strategy_research.md](_research/composition_strategy_research.md)

---

## 1. N1-Lite 정의 (재확인)

**한 문장**: 1,251장 마스터에서 도형 그룹(GroupShape) 1개를 추출해, 빈 슬라이드에 단독 배치 + placeholder 채움.

**적용 role** (5개): situation, complication, roadmap, benefit, risk

**NOT 적용**:
- 여러 컴포넌트 자동 조합 (= N1-Full, 미해결 hard problem)
- 컴포넌트 간 충돌 해결, style harmonization
- 컴포넌트 transformation (resize, recolor 등)

**적용 정책**: situation/complication/risk 등 **풀 결손 role**에서, Mode A로 못 채우는 자리에 N1-Lite로 단일 컴포넌트 슬라이드 사용.

---

## 2. 시스템 아키텍처

```
사용자 input
    ↓
Outline-First Planning (스켈레톤 + role 매칭)
    ↓
role별 모드 결정 (Mode A or N1-Lite)
    ↓                   ↓
[Mode A]           [N1-Lite]
whole-slide       blank slide
clone (cloner.py)  + insert_component (component_ops.py)
    ↓                   ↓
paragraph fill (paragraph_query.match_content)
    ↓
저장 + PNG 렌더 + 평가
```

---

## 3. 신규 모듈 — `ppt_builder/template/component_ops.py`

`edit_ops.py` 옆에 추가. 기존 인프라 재사용 + 확장.

### 3.1 핵심 API (4개)

```python
# 컴포넌트 추출 (마스터에서)
def extract_group(
    src_slide: Slide,
    group_indices: list[int],  # iter_leaf_shapes 평탄 인덱스
) -> ComponentXML

# 컴포넌트 삽입 (대상 슬라이드에)
def insert_component(
    dst_slide: Slide,
    component: ComponentXML,
    position: tuple[int, int] | None = None,  # (left_emu, top_emu), None=원본 위치
    rel_remap: dict[str, str] | None = None,
) -> int  # inserted shape의 dst_slide flat_idx

# 컴포넌트 라이브러리 빌드 (1251 마스터 → library)
def build_component_library(
    master_path: Path,
    library_path: Path,
    selection_rules: dict,
) -> ComponentIndex

# 컴포넌트 단독 슬라이드 생성 (N1-Lite)
def create_single_component_slide(
    target_pptx: Path,
    component_id: str,
    library: ComponentIndex,
    title_text: str | None = None,
) -> Slide
```

### 3.2 추출 알고리즘 (Technique T1, 리서치 권장)

```python
# Pseudo-code
def extract_group(src_slide, group_indices):
    # 1. iter_leaf_shapes로 대상 shapes 식별
    # 2. _spTree에서 해당 element 찾기
    # 3. copy.deepcopy로 element 복제
    # 4. a16:creationId 재생성 (UUID4) — 중복 방지
    # 5. rId list 추출 (이미지/차트 참조)
    # 6. ComponentXML(xml_element, rids, bbox) 반환
```

### 3.3 삽입 알고리즘

```python
def insert_component(dst_slide, component, position=None):
    # 1. dst_slide.part에 component.rids re-relate (이미지/차트)
    # 2. component.xml의 r:embed/r:link rId rewrite
    # 3. position 지정 시 _grpSpPr/transform/offset 갱신
    # 4. dst_slide._spTree.insert_element_before(...) 추가
    # 5. 새 flat_idx 반환
```

### 3.4 함정 + 회피책 (리서치 §2)

| 함정 | 회피 |
|---|---|
| `a16:creationId` 중복 → 2회째 사용 시 PPT 손상 | deepcopy 후 모든 creationId UUID4로 재생성 |
| `r:embed` rId가 slide-part scope → 다른 슬라이드에 그대로 붙이면 이미지 깨짐 | `dst_slide.part.relate_to()` + XML rId rewrite |
| 한글 폰트 (`<a:ea>`) | deepcopy로 보존됨 (Phase A1 검증 완료) |
| 차트 part 참조 | XML deepcopy로 그래프 데이터 단절 — chart part도 복사 필요 (v1: chart 컴포넌트 제외) |
| Cross-master theme | v1 범위 외 (같은 master 내 추출만) |

---

## 4. 컴포넌트 라이브러리 구조

### 4.1 파일 배치

```
output/component_library/
├── components_index.json         # 메타데이터 (전체 인덱스)
├── chevron_family.pptx           # 5-chevron, 4-chevron, 3-chevron 등
├── card_family.pptx              # cards_3col, cards_4col, cards_5plus
├── timeline_family.pptx          # timeline_h, roadmap, gantt
├── matrix_family.pptx            # 2x2_matrix, comparison_table
├── callout_family.pptx           # speech_bubbles, side_callouts
├── kpi_family.pptx               # KPI metric tiles, 큰 숫자
├── icon_grid_family.pptx         # 4-icon, 6-icon, 8-icon grids
├── table_family.pptx             # standard tables (small)
├── flow_family.pptx              # flowchart, swimlane
└── divider_family.pptx           # section dividers
```

각 파일 안에는 그 family의 변형(variant) 슬라이드들 보관.
- 파일당 평균 3~5 variants
- 총 약 30~50 컴포넌트 (5 N1-Lite role 커버 충분)

### 4.2 components_index.json 스키마

```json
{
  "version": "1.0",
  "generated_at": "2026-04-25",
  "components": [
    {
      "component_id": "chevron_5col_v1",
      "family": "chevron",
      "library_path": "chevron_family.pptx",
      "library_slide_index": 0,
      "source": {
        "master_slide_index": 44,
        "group_indices": [11, 12, 13, 14, 15],
        "macro": "diagram",
        "archetype": ["roadmap", "cards_5plus"]
      },
      "geometry": {
        "bbox_emu": {"left": 457200, "top": 914400, "width": 8534400, "height": 685800},
        "bbox_pct": {"left": 0.046, "top": 0.133, "width": 0.862, "height": 0.1},
        "footprint_grid_12x6": [[0,1], [11,1]]
      },
      "slots": [
        {"flat_idx": 0, "paragraph_id": 0, "role": "chevron_label", "max_chars": 22, "position_in_group": 0},
        {"flat_idx": 1, "paragraph_id": 0, "role": "chevron_label", "max_chars": 22, "position_in_group": 1},
        ...
      ],
      "applicable_roles": ["roadmap", "recommendation", "complication"],
      "narrative_hints": ["병렬 5단계", "수평 흐름"]
    }
  ]
}
```

### 4.3 컴포넌트 후보 추출 규칙

`build_component_library()`가 1,251 마스터 순회 시:

```python
selection_rules = {
    "min_group_size": 2,      # 그룹 최소 멤버 수
    "max_group_size": 12,     # 그룹 최대 (decorative grid 회피)
    "min_text_slots": 1,      # 최소 채울 슬롯 1개 이상
    "exclude_archetypes": ["dense_grid", "dense_table"],  # 컴포넌트 부적합
    "exclude_decorative_only": True,  # 모든 슬롯이 decorative만이면 skip
    "prefer_archetypes": [
        "chevron_flow", "cards_3col", "cards_5plus",
        "roadmap", "timeline_h", "matrix_2x2",
        "left_title_right_body", "vertical_list",
    ],
}
```

규칙 통과한 그룹 → 같은 archetype family로 묶어 family.pptx에 저장.

---

## 5. 하이브리드 Retrieval + Assembly

### 5.1 모드 결정 (PROJECT_DIRECTION §3 매핑 적용)

```python
ROLE_MODE_MAP = {
    "opening": "mode_a",
    "agenda": "mode_a",
    "divider": "mode_a",
    "situation": "n1_lite",
    "complication": "n1_lite",
    "evidence": "mode_a",
    "analysis": "mode_a",
    "recommendation": "mode_a",
    "roadmap": "n1_lite",
    "benefit": "n1_lite",
    "risk": "n1_lite",
    "closing": "mode_a",
    "appendix": "mode_a",
}
```

사용자 override 가능: `select_deck_hybrid(narrative, scenario_content, override={"roadmap": "mode_a"})`.

### 5.2 N1-Lite 슬라이드 생성 흐름

```
role = "situation"
↓
mode = "n1_lite"
↓
컴포넌트 라이브러리에서 후보 검색 (applicable_roles 매칭)
↓
컨텐츠 N개 vs 컴포넌트 슬롯 capacity 매칭 (best fit)
↓
빈 슬라이드 생성 (master 같은 theme 적용, blank layout)
↓
component_ops.insert_component(blank, chosen_component)
↓
title 추가 (옵션) — 마스터의 title placeholder 가져오기
↓
paragraph_query.match_content로 슬롯 채움
↓
완료
```

### 5.3 빈 슬라이드 생성 (master theme 보존)

```python
def create_blank_slide_with_master_theme(target_prs, master_prs):
    """마스터의 blank layout을 재사용해 같은 theme의 빈 슬라이드 생성."""
    # 1. master에서 blank layout 찾기 (보통 마지막 layout)
    # 2. target에 같은 layout 추가 (이미 같은 master 사용 시 자동)
    # 3. blank layout으로 새 슬라이드 추가
```

이미 우리는 같은 마스터(`PPT 템플릿.pptx`) 기반이라 theme/font 자동 일치.

---

## 6. 5 시나리오 적용 시뮬레이션

### transformation_roadmap_10 (10 슬라이드)

| step | role | 모드 | 자산 |
|---|---|---|---|
| 1 | opening | Mode A | master slide#783 (cover) |
| 2 | situation | **N1-Lite** | callout_family / cards_2col 컴포넌트 |
| 3 | complication | **N1-Lite** | matrix_2x2 / cards_3col |
| 4 | analysis | Mode A | master slide#8 (table+chart) |
| 5 | recommendation | Mode A | master slide#44 (chevron framework) |
| 6 | roadmap | **N1-Lite** | timeline_family / chevron_5col |
| 7 | roadmap | **N1-Lite** | (다른 timeline variant) |
| 8 | benefit | **N1-Lite** | callout_family / kpi_grid |
| 9 | risk | **N1-Lite** | matrix_2x2 / cards_3col |
| 10 | closing | Mode A | master slide#80 |

→ 10 슬라이드 중 6개 N1-Lite, 4개 Mode A. 결손 role 모두 N1-Lite로 해결.

---

## 7. 구현 단계 (실행 가이드)

### Phase 1 — 컴포넌트 추출 인프라 (1일)
- [ ] `ppt_builder/template/component_ops.py` 작성
  - extract_group, insert_component 핵심 2개
  - rId rewrite + creationId 재생성 헬퍼
- [ ] 단위 테스트: 마스터 slide#44에서 5-chevron 추출 → 빈 슬라이드 삽입 → PNG 검증

### Phase 2 — 컴포넌트 라이브러리 구축 (0.5일)
- [ ] `scripts/build_component_library.py` 작성
- [ ] 1,251장 순회 → selection_rules 통과 그룹 추출 → 10 family.pptx 생성
- [ ] components_index.json 생성
- [ ] 사용자 시각 검수 (10 family에서 대표 1~2장씩 = 10~20장 PNG)

### Phase 3 — 하이브리드 retrieval + assembly (0.5일)
- [ ] `scripts/benchmark_5_scenarios_v6.py` 작성 (v2 base)
- [ ] ROLE_MODE_MAP 적용 + N1-Lite 분기
- [ ] create_single_component_slide 호출
- [ ] paragraph fill 통합

### Phase 4 — 5 시나리오 측정 + 시각 검증 (0.5일)
- [ ] 5 시나리오 v6 실행
- [ ] 점수: v5 (57.8) → v6 비교
- [ ] 시각 검수 25장 (5 시나리오 × 5 대표)
- [ ] 70+ 도달 여부 판정

### Phase 5 — 결정 (사용자)
- 70+ 도달: 운영 hardening (Phase B 전환)
- 60~70: REFINE 반복
- 60 미만: 헌법 §6 변경 조건 검토

**총 자동 시간**: 2~3일 (사용자 검수 시간 1~2시간)

---

## 8. 위험 + 대응

| 위험 | 확률 | 대응 |
|---|---|---|
| insert_component 후 PPT 손상 (creationId 등) | 중 | 단위 테스트로 사전 검증, validate_pptx 헬퍼 추가 |
| N1-Lite 슬라이드가 빈약/허전해 보임 | 중 | title placeholder 추가, decorative line/divider 추가 옵션 |
| 컴포넌트 포지션이 슬라이드 가운데 안 맞음 | 중 | bbox 자동 centering 옵션 |
| 5 family.pptx 파일 크기 누적 | 낮 | 파일별 50~150MB 예상 — 문제 없음 |
| 차트/이미지 의존성 손실 | 중 | v1은 차트 컴포넌트 제외 (selection_rules에서 chart_native 빼기) |

---

## 9. 측정 지표 (v6 평가)

기존 v5 metric + N1-Lite 전용 추가:

```python
metrics_v6 = {
    "A_role_match_pct": ...,
    "B_fill_ratio": ...,
    "C_overflow_rate": ...,
    # N1-Lite 신규
    "N_n1_lite_used": int,           # N1-Lite 사용 슬라이드 수
    "N_component_match_rate": float, # N1-Lite role 컨텐츠가 컴포넌트 슬롯에 fit
    "N_visual_continuity": float,    # vision 검수 (선택적)
    "composite_v6": ...,
}
```

목표:
- 평균 composite **70+**
- 시각 검수 5 시나리오 모두 "수용 가능" 판정

---

## 10. 산출 (이 사양 구현 후)

### 코드
- `ppt_builder/template/component_ops.py` — 컴포넌트 추출/삽입 API
- `scripts/build_component_library.py` — 라이브러리 빌드
- `scripts/benchmark_5_scenarios_v6.py` — 하이브리드 평가

### 데이터
- `output/component_library/*.pptx` (10 family)
- `output/component_library/components_index.json`
- `output/benchmark_v6/{scenario}/deck.pptx + pngs/ + report.json` × 5

### 문서
- 본 사양 (이 파일) 업데이트 시
- 측정 결과 보고서: `docs/PHASE_A3_v6_HYBRID_REPORT.md`

---

## 11. 새 세션에서 시작 시

```
1. PROJECT_DIRECTION.md 읽기 (헌법)
2. 본 N1_LITE_IMPLEMENTATION.md 읽기 (사양)
3. _research/component_extraction_research.md 읽기 (기술 근거)
4. _research/composition_strategy_research.md 읽기 (전략 근거)
5. 진행 위치 확인 → Phase 1~5 중 어디인지
6. 다음 단계 진행
```

---

## 12. 변경 이력

| 일자 | 변경 |
|---|---|
| 2026-04-25 | 초안 + PROJECT_DIRECTION 헌법 기반 사양 작성 |
