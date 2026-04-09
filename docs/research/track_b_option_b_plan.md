# Track B - Option B 구현 계획서

> **목표**: 외부 .pptx 템플릿에서 컴퍼넌트(SWOT 박스, KPI 카드, 타임라인 등)를 단위로 추출하고, 코드로 새 슬라이드에 자유롭게 재조립한다.
>
> **북극성**: "원본 디자인 무결성 100% 보존 + 임의 조합 자유"

---

## 1. 핵심 도전 과제 (왜 어려운가)

| 영역 | 난이도 | 이유 |
|---|---|---|
| Group shape 추출 | 낮음 | python-pptx에서 이미 지원 |
| 그룹화 안된 shape의 자동 클러스터링 | 중간 | 휴리스틱 (DBSCAN, alignment) 필요 |
| Cross-slide shape XML 복사 | 중간 | `cNvPr/@id` unique 재할당 필요 |
| **이미지 rId remap** | **중간** | 현재 cloner.py의 깨진 부분. relate_to() 패턴으로 해결 |
| **차트 cross-slide 복사** | **높음** | python-pptx 한계. matplotlib 자체 생성으로 우회 |
| SmartArt 복사 | 매우 높음 | python-pptx 미지원. ungroup된 템플릿만 사용 |
| Theme 색상 호환성 | 중간 | accent1 등이 대상 테마에서 다르게 렌더. RGB 하드코딩 |
| Collision detection / 자동 배치 | 중간 | 단순 grid + rectpack으로 충분 |
| 색상/폰트 통일 (회사 컬러 강제 적용) | 중간 | shape 트리 순회하며 원본 색을 회사 컬러로 매핑 |

---

## 2. 벤치마크 (오픈소스 사례)

| 프로젝트 | 시사점 |
|---|---|
| [scanny/python-pptx](https://github.com/scanny/python-pptx) | 본체. `oxml/` 패키지가 XML helper 보고. |
| [APryor6/pptx_tools - copy_slide_from_external_prs](https://github.com/APryor6/pptx_tools) | 외부 pptx → 현재 pptx 슬라이드 임포트, ImagePart 이식 정석 구현 |
| [mrtj/pptx-tools](https://github.com/mrtj/pptx-tools) | 슬라이드 merge/split, duplicate_slide의 rels 재매핑 |
| [secnot/rectpack](https://github.com/secnot/rectpack) | 자동 배치 (shelf packing) |
| [Aspose.Slides for Python](https://products.aspose.com/slides/python-net/) | 상용 — SmartArt/차트/애니메이션 완벽. python-pptx로 안 되는 영역 확인용 |
| [Stack Overflow #56074651 (scanny)](https://stackoverflow.com/a/56074651) | duplicate_slide canonical 답변 |
| [Instrumenta](https://github.com/iappyx/Instrumenta) | 슬라이드 라이브러리 UX 패턴 (VBA지만 개념 참고) |

**가장 중요한 참고**: APryor6의 `copy_slide_from_external_prs` — 이미지 Part를 외부 prs에서 현재 prs로 복사하는 로직이 우리 프로젝트의 핵심 기술 부채(현재 cloner.py의 `_copy_image_rels`가 `pass`만 있음)를 정확히 해결합니다.

---

## 3. 아키텍처 설계

### 3.1 새로운 모듈 구조

```
ppt_builder/
├── component_lib/                  ← NEW: 컴퍼넌트 라이브러리 시스템
│   ├── __init__.py
│   ├── extractor.py                ← 외부 .pptx에서 컴퍼넌트 추출
│   ├── catalog.py                  ← 추출된 컴퍼넌트 메타데이터 (JSON)
│   ├── injector.py                 ← 빈 슬라이드에 컴퍼넌트 삽입
│   ├── grouper.py                  ← 자동 그룹핑 (DBSCAN, alignment)
│   ├── color_normalizer.py         ← 회사 컬러로 정규화
│   └── packer.py                   ← 자동 배치 (rectpack)
├── template/
│   ├── editor.py                   ← (기존, Track B 슬라이드 단위)
│   └── ...
└── ...

templates/
├── component_library/              ← NEW: 추출된 컴퍼넌트 저장소
│   ├── catalog.json                ← 모든 컴퍼넌트 메타데이터
│   ├── components/
│   │   ├── swot_2x2_smartsheet.xml ← 추출된 shape XML 직렬화
│   │   ├── kpi_card_3col_general.xml
│   │   └── ...
│   └── thumbnails/                 ← 미리보기 PNG (선택)
```

### 3.2 데이터 모델

**Component (추출 단위)**:
```python
@dataclass
class Component:
    id: str                    # "swot_2x2_smartsheet_general"
    category: str              # "framework" | "kpi" | "timeline" | "chart" | "callout"
    subcategory: str           # "swot" | "bcg_matrix" | "value_chain" ...
    source_file: str           # "smartsheet_general.pptx"
    source_slide: int          # 4
    
    # Geometry (원본 좌표, EMU)
    bbox: tuple[int, int, int, int]  # (left, top, width, height)
    aspect_ratio: float                # width/height
    
    # XML 데이터
    sp_tree_xml: str           # 직렬화된 shape XML 묶음
    image_rels: dict           # {old_rId: image_blob_path}
    
    # Slot (텍스트 치환 placeholder)
    slots: list[Slot]          # 각 텍스트 위치와 의미 정보
    
    # 메타데이터 (자동 분석)
    color_palette: list[str]   # 사용된 색상 (RGB hex)
    font_families: list[str]
    text_density: int          # 총 글자 수
    has_images: bool
    has_charts: bool
    has_smartart: bool         # True면 사용 비추천 (복사 깨짐)


@dataclass
class Slot:
    """컴퍼넌트 내 텍스트 치환 위치."""
    slot_id: str               # "title" | "quadrant_1_label" | "kpi_value_1"
    semantic_role: str         # "title" | "label" | "value" | "description"
    original_text: str         # 원본 텍스트 (치환 전)
    bbox_in_component: tuple[int, int, int, int]
    max_chars: int             # 안 넘치는 한도 (휴리스틱)
    font_size: int             # 원본 폰트 크기
```

**ComponentRequest (사용 시)**:
```python
@dataclass
class ComponentRequest:
    component_id: str          # "swot_2x2_smartsheet_general"
    target_bbox: tuple[float, float, float, float]  # 대상 위치 (인치)
    content: dict[str, str]    # {slot_id: text}
    color_override: dict[str, str] | None  # 원본 색 → 새 색 매핑 (선택)
```

### 3.3 핵심 API

```python
# Phase 1: 추출
from ppt_builder.component_lib.extractor import ComponentExtractor

extractor = ComponentExtractor()
components = extractor.extract_from_pptx("templates/external/umbrex.pptx")
# → 자동 그룹핑으로 N개 컴퍼넌트 식별
extractor.save_to_library(components, "templates/component_library/")

# Phase 2: 카탈로그 조회
from ppt_builder.component_lib.catalog import ComponentCatalog

catalog = ComponentCatalog("templates/component_library/")
swot_components = catalog.find(category="framework", subcategory="swot")
# → 여러 SWOT 디자인 후보 반환

# Phase 3: 조립
from ppt_builder.component_lib.injector import ComponentInjector
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

injector = ComponentInjector(catalog)
injector.inject(
    slide,
    ComponentRequest(
        component_id="swot_2x2_smartsheet_general",
        target_bbox=(0.3, 1.0, 6.0, 5.5),  # 인치
        content={
            "title": "SAP S/4HANA 전환 SWOT",
            "quadrant_1_label": "강점",
            "quadrant_1_text": "글로벌 표준화",
            ...
        },
    )
)

injector.inject(
    slide,
    ComponentRequest(
        component_id="kpi_card_3col_general",
        target_bbox=(6.5, 1.0, 3.2, 5.5),
        content={...}
    )
)

prs.save("output/composed.pptx")
```

---

## 4. 단계별 구현 로드맵

### Phase 1: 기반 — Group Shape 추출 + 단일 컴퍼넌트 복사 (가장 안전)

**범위**: 명시적 GroupShape만 다룸. 자동 그룹핑 없음.

| 작업 | 파일 | 출력 |
|---|---|---|
| 1.1 | `component_lib/__init__.py`, `models.py` | dataclass 정의 (Component, Slot, ComponentRequest) |
| 1.2 | `component_lib/extractor.py` | `extract_groups(pptx_path) -> list[Component]` — 모든 GroupShape를 컴퍼넌트로 추출 |
| 1.3 | `component_lib/catalog.py` | `Catalog.save() / load()` — JSON + XML 직렬화 |
| 1.4 | `component_lib/injector.py` | `inject(slide, request)` — bbox 위치에 sp_tree 복사 + 좌표 변환 |
| 1.5 | **이미지 rId remap** (`_remap_image_rels`) | APryor6 패턴 적용. cloner.py의 깨진 함수도 함께 수정 |
| 1.6 | `tests/test_extract_inject.py` | smartsheet_general.pptx에서 그룹 1개 추출 → 빈 슬라이드 삽입 → 시각 검증 |

**검증 기준**:
- 추출 → 삽입 후 원본과 비교했을 때 색/폰트/위치가 일치한다.
- 이미지가 포함된 그룹도 정상 복사된다 (현재 깨진 부분 해결).

### Phase 2: 자동 그룹핑 — Ungrouped Shape 클러스터링

**범위**: 그룹화 안 된 shape들도 시각적으로 묶어서 컴퍼넌트로 식별.

| 작업 | 파일 | 알고리즘 |
|---|---|---|
| 2.1 | `component_lib/grouper.py` | (a) Centroid DBSCAN, (b) bbox overlap union-find, (c) alignment grouping |
| 2.2 | `extractor.py` 확장 | `extract_components(pptx_path)` — group + auto-grouped 모두 |
| 2.3 | `tests/test_grouper.py` | smartsheet_strategy의 cash flow 슬라이드에서 자동 그룹 N개 식별 |

**검증 기준**:
- 자동 그룹핑 결과가 사람이 보기에 합리적이다 (수동 검증 10케이스).

### Phase 3: 슬롯 인식 — 의미 있는 텍스트 치환

**범위**: 컴퍼넌트 안의 텍스트를 의미 단위로 식별.

| 작업 | 파일 | 방법 |
|---|---|---|
| 3.1 | `component_lib/slot_detector.py` | 폰트 크기/위치/색으로 title/label/value 분류 |
| 3.2 | `Component.slots` 자동 채우기 | 추출 단계에서 슬롯 추론 |
| 3.3 | `injector.py` 텍스트 치환 강화 | TextSubstitutor 활용 + 길이 검증 (overflow 경고) |

**검증 기준**:
- SWOT 컴퍼넌트에서 slot이 `title, q1_label, q1_text, q2_label, q2_text, q3_label, q3_text, q4_label, q4_text` 자동 추론된다.

### Phase 4: 색상/폰트 정규화 — 회사 컬러 강제 적용

**범위**: 추출된 컴퍼넌트의 색을 회사 팔레트(Orange #FD5108 + Grey)로 자동 변환.

| 작업 | 파일 | 방법 |
|---|---|---|
| 4.1 | `component_lib/color_normalizer.py` | 원본 색을 K-means로 클러스터링 → 회사 팔레트의 가장 가까운 색에 매핑 |
| 4.2 | XML 트리 순회하며 `<a:srgbClr val="...">` 치환 | accent 추론도 RGB로 변환 |
| 4.3 | 폰트 통일 | `<a:rFont typeface="...">`를 회사 폰트로 |
| 4.4 | `injector.inject(...)`에 `normalize_colors=True` 옵션 | 디폴트 ON |

**검증 기준**:
- 빨강/파랑/녹색 SWOT을 추출 → 회사 컬러로 자동 변환 → 모든 강조색이 #FD5108

### Phase 5: 자동 배치 — Multiple Components on One Slide

**범위**: 한 슬라이드에 여러 컴퍼넌트를 자동 위치 배치.

| 작업 | 파일 | 방법 |
|---|---|---|
| 5.1 | `component_lib/packer.py` | 12-col grid + rectpack shelf-fit |
| 5.2 | `SlideComposer` 클래스 | 여러 ComponentRequest를 받아 자동 배치 |
| 5.3 | Collision detection | AABB overlap 검사, 충돌 시 경고 + 재배치 시도 |
| 5.4 | `tests/test_compose.py` | SWOT + KPI + 콜아웃을 한 슬라이드에 자동 배치 |

**검증 기준**:
- 4개 컴퍼넌트를 자동 배치 → 겹침 없음 + 안전 영역(margin) 준수

### Phase 6: 카탈로그 빌드 + 통합 워크플로우

| 작업 | 결과 |
|---|---|
| 6.1 | 외부 6개 + Umbrex 200 + 24Slides 등 모든 .pptx에서 컴퍼넌트 추출 → catalog.json |
| 6.2 | 카탈로그 통계 (예: SWOT 디자인 12개, KPI 디자인 8개) |
| 6.3 | Claude Code용 컴퍼넌트 선택 가이드 (`docs/component_selector.md`) |
| 6.4 | `render_presentation()` API에 컴퍼넌트 조립 모드 추가 |

---

## 5. 우선순위 결정 매트릭스

| Phase | 가치 | 난이도 | 우선순위 |
|---|---|---|---|
| Phase 1 (Group 추출 + rId remap) | 매우 높음 | 중간 | **즉시** |
| Phase 4 (색상 정규화) | 매우 높음 | 중간 | **2순위** — 회사 컬러 통일은 컨설팅 품질의 핵심 |
| Phase 3 (슬롯 인식) | 높음 | 낮음 | 3순위 |
| Phase 5 (자동 배치) | 높음 | 중간 | 4순위 |
| Phase 6 (카탈로그 빌드) | 높음 | 낮음 (반복 작업) | 5순위 |
| Phase 2 (자동 그룹핑) | 중간 | 높음 | 6순위 — 일단 group된 컴퍼넌트만 다뤄도 충분 |

**Phase 2를 후순위로 미루는 이유**: 외부 템플릿 대부분의 핵심 다이어그램은 이미 GroupShape로 묶여 있을 가능성이 높다. 먼저 group만 추출해서 카탈로그를 만들고, 부족하면 그때 자동 그룹핑을 도입한다.

---

## 6. 알려진 함정과 대응

| 함정 | 대응 |
|---|---|
| 차트(`graphicFrame`) 복사 시 embedded xlsx가 깨짐 | 차트 포함 컴퍼넌트는 추출 거부. 차트는 matplotlib 자체 생성 |
| SmartArt(`<dgm:>`) 복사 시 일반 shape으로 깨짐 | 추출 단계에서 SmartArt 감지 → 컴퍼넌트로 등록하지 않음 |
| Theme accent1 색 차이로 색이 달라 보임 | XML 추출 시 모든 색을 RGB로 하드코딩 |
| 폰트 미설치로 viewer 머신에서 폰트 대체됨 | 회사 폰트(Arial/Georgia/맑은 고딕)로 강제 통일 |
| Image blob 중복 → 파일 크기 폭증 | 동일 image hash는 한 번만 임베드 (Phase 6 최적화) |
| `cNvPr/@id` 충돌 → PowerPoint 열기 실패 | 삽입 시 max_id+1 부터 재할당 |
| 그룹 좌표가 chOff/chExt 변환을 거쳐 inject 후 위치가 어긋남 | 부모 `<a:off>`만 수정. 자식 좌표는 건드리지 않음 |
| 일부 .pptx가 손상되어 못 열림 (marknold 사례) | extractor는 try/except로 skip + 로그 |

---

## 7. 성공 측정 지표

Phase 1 완료 시점:
- [ ] smartsheet_general.pptx에서 GroupShape N개 추출 성공
- [ ] 추출된 그룹을 빈 슬라이드에 삽입했을 때 원본과 시각적으로 일치
- [ ] 이미지 포함 그룹도 정상 복사 (현재 cloner.py 깨진 부분 해결)
- [ ] cloner.py의 `_copy_image_rels` 정상 동작 보너스

Phase 4 완료 시점:
- [ ] 빨강 SWOT → 회사 오렌지로 자동 변환
- [ ] 폰트 자동 통일

Phase 5 완료 시점:
- [ ] 한 슬라이드에 3+ 컴퍼넌트 자동 배치, 충돌 0

전체 완료 시점:
- [ ] 카탈로그에 50+ 컴퍼넌트 등록
- [ ] Claude Code가 카탈로그를 보고 컴퍼넌트 선택 → 조립 → .pptx 생성하는 end-to-end 워크플로우 동작

---

## 8. 즉시 시작 가능한 첫 단계

**Phase 1.1 + 1.2**부터 시작:

1. `ppt_builder/component_lib/__init__.py`, `models.py`, `extractor.py` 생성
2. `extract_groups("templates/external/smartsheet_general.pptx")` → GroupShape 인벤토리 출력
3. 추출된 그룹의 XML을 파일로 저장
4. 빈 슬라이드에 다시 삽입 → 시각 검증

이걸 먼저 해보면 **현실적인 어려움이 어디에 있는지** 빠르게 파악할 수 있습니다.

---

## 9. 사용자 결정 필요 사항

다음 중 어떤 .pptx부터 추출을 시작할까요?

A. **smartsheet_general.pptx** (이미 있음, 검증 용이)
B. **Umbrex 200 슬라이드** (다운로드 필요, 가장 풍부)
C. **회사 템플릿** (최우선이지만 아직 안 받음)

권장: A로 첫 검증 → 성공 후 B/C로 확장

---

## 10. 참고 — 옵션 A vs 옵션 B 비교

| 항목 | 옵션 A (현재 editor.py) | 옵션 B (이 계획) |
|---|---|---|
| 단위 | 슬라이드 | 컴퍼넌트 |
| 디자인 보존 | 100% | 95% (색 정규화로 약간 변형) |
| 자유도 | 낮음 | 높음 |
| 한 슬라이드에 여러 요소 조합 | X | O |
| 회사 컬러 강제 적용 | X | O (Phase 4) |
| 구현 난이도 | 완료 | Phase 1만 1주 정도 |
| 기존 slidemodel.com/24slides 등 활용도 | 슬라이드 단위 | **컴퍼넌트 단위 → 활용도 폭발** |

옵션 B 완성 시 Track A의 코드 컴퍼넌트(현재 18종)를 외부 컴퍼넌트(50+)로 보강할 수 있어, 컨설팅 품질의 다양성이 크게 확장됩니다.
