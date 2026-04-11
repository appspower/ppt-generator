# 08. 복합 조합 시스템 설계서

> "패턴 1개 = 슬라이드 1개" 탈피 → 복수 패턴+컴포넌트 유기적 조합

## 1. 현황 진단

### 1.1 완성본 역설계 (27장 전수 분석)

PwC Component Library 사진에서 추출한 **6가지 복합 조합 패턴(CP)**:

| CP# | 패턴명 | 구조 | 대표 슬라이드 |
|-----|-------|------|------------|
| CP1 | **Central Diagram + Peripheral Text** | 중앙 도형(헥사곤/다이아몬드/도넛) + 주변 4~6개 텍스트 블록 | B03, B04, B06, B07 |
| CP2 | **Numbered Grid** | N×M 그리드, 각 셀에 번호+헤더+본문, 색상 그라데이션 | B00, B01 |
| CP3 | **Timeline Band + Content Zones** | 중앙 타임라인 밴드 + 상하 교차 텍스트 | B08, B09, B10 |
| CP4 | **Heterogeneous Panels** | 2~3개 서로 다른 유형(차트+KPI+표)이 한 슬라이드에 공존 | A00, A01, A03 |
| CP5 | **Diagram + Rich Annotation** | 차트/도넛 다이어그램 + 주변 상세 텍스트/불릿 | B11, B12 |
| CP6 | **Shape-Enhanced Cards** | 동일 레이아웃에 아이콘/도형 오버레이 추가 | B05 vs B06, B07 |

### 1.2 핵심 발견

1. **시각 요소가 레이아웃의 구조적 중심** — 헥사곤, 다이아몬드, 도넛이 장식이 아니라 슬라이드 구조 자체를 결정
2. **비대칭 존 배치** — 좌우/상하가 동일 컴포넌트 반복이 아니라, 완전히 다른 유형이 공존
3. **텍스트는 도형의 위성** — 텍스트 블록이 중앙 도형 주변에 방사형으로 배치
4. **색상이 정보 계층을 구분** — 그리드 셀별 색상 그라데이션으로 단계/우선순위 표현
5. **아이콘이 구조적 역할** — 단순 장식이 아니라 카드/섹션의 시각적 앵커

### 1.3 현재 시스템 격차

| 격차 | 원인 (코드 레벨) | 영향 |
|------|--------------|------|
| **G1: 패턴이 전체 슬라이드 독점** | `patterns.py` 25개 함수 모두 절대좌표 하드코딩 (HX=0.3, HY=1.25 등) | 패턴 중첩/조합 불가 |
| **G2: Composer가 직사각형만 지원** | `composer.py` 8개 레이아웃 전부 rectangular (two_column, grid_2x2 등) | CP1 (중앙+방사형) 불가 |
| **G3: 시각 요소가 레이아웃 역할 못함** | 도형이 Canvas 드로잉일 뿐, 주변 존을 정의하지 않음 | CP1, CP5, CP6 불가 |
| **G4: 패턴 스펙에 컴포넌트 참조 없음** | XSpec 데이터클래스가 데이터만 받고 컴포넌트 참조 불가 | CP4 (이종 패널) 불가 |
| **G5: fill_pattern()이 미작동** | push_region() 후에도 패턴 내부 좌표가 절대값 | 패턴→존 배치 실패 |

**근본 원인**: 두 설계 시대(Era 1: 패턴=풀슬라이드 vs Era 2: 컴포넌트+존)가 통합되지 않음.

---

## 2. AI PPT 도구 벤치마킹

### 2.1 주요 도구의 복합 구성 방식

| 도구 | 레이아웃 방식 | 복합 구성 | 우리 시스템 적용 가능성 |
|------|-----------|---------|-------------------|
| **Gamma** | Card-Block 기반 — 카드 내 중첩 블록(텍스트+차트+표+임베드) | 카드 안에 카드 가능, 컬럼 기반 그리드, 수직 확장 | **중첩 구성(Nested Card)** 개념 → Zone 안에 Sub-zone |
| **Beautiful.ai** | Constraint-based Smart Slide (300+종) — 콘텐츠 추가/삭제 시 자동 리플로우 | 존 기반 적응형 크기 조정, 폰트/간격 자동 스케일 | **Zone min/max 제약** + 콘텐츠 양 → 자동 조정 |
| **Presenton** | Schema-first 템플릿 — .tsx에 JSON 스키마 + React 렌더러 | density 파라미터(concise/standard/dense), AI가 템플릿 선택 | **Schema 선언 → 레이아웃 선택** 우리 Pydantic과 유사 |
| **Tome** | 내러티브 기반 스토리보드 — AI 이미지 인라인 생성 | 자체 형식, PPT 호환 약함 | 참고 가치 낮음 |
| **Canva** | 템플릿 라이브러리 + Magic Design — 기존 레이아웃 변형 제안 | 브랜드 기반 변형 추천, 자유 배치 | 우리 RECIPES와 유사 |

### 2.2 핵심 교훈 (채택 가능한 아키텍처 패턴)

1. **Constraint-based Zone Reflow** (Beautiful.ai) — 존에 min/max 크기 제약을 두고, 콘텐츠 양에 따라 존 크기가 제약 범위 내에서 적응. 하드코딩된 절대좌표의 대안
2. **Schema-first Template Declaration** (Presenton) — 각 레이아웃이 "어떤 콘텐츠 슬롯을 받는지" 타입으로 선언. AI가 콘텐츠 형태를 보고 레이아웃 선택. 우리 Pydantic 모델이 이미 부분적으로 구현
3. **Nested Card Composition** (Gamma) — 조합 안에 하위 조합이 들어갈 수 있음 (차트카드 → 그리드 안). Region.sub()로 구현 가능
4. **Content Density Parameter** (Presenton) — 단일 density 파라미터로 폰트/간격/요소 수를 전역 조정. 우리 P1 Dynamic Sizing과 맥락 동일
5. **우리의 차별점**: Claude Code가 "두뇌" 역할이므로, Gamma/Beautiful.ai의 자동 레이아웃 엔진을 **Claude의 판단으로 대체**. 코드에는 **충분한 레이아웃 옵션**만 제공하면 됨

### 2.3 참고 소스

- Gamma: card 기반 중첩 구성, 2026년 "Generative Layout" 도입 (슬라이드별 커스텀 레이아웃 생성)
- Beautiful.ai: 특허 기반 constraint 엔진, 300+ Smart Slide, 자유 배치 의도적 차단
- Presenton: 오픈소스 (GitHub), React+Tailwind 렌더링, 스키마 기반 레이아웃 선택

---

## 3. 설계: Blueprint 시스템

### 3.1 아키텍처

기존 패턴을 전면 리팩토링하지 않고, **새 레이어를 추가**한다:

```
현재:
  Pattern(full-slide, 절대좌표) ──→ Slide (1:1)
  Component(region-aware) ──→ Region
  Composer(직사각형 8종) ──→ Layout

추가:
  Blueprint(복합 레이아웃)
    ├─ Anchor Component (시각 중심 — 도형/차트/타임라인)
    ├─ Layout Function (앵커 기반 비직사각형 존 분할)
    └─ Zone Fillers (각 존에 기존 component 배치)
```

**핵심 원칙**: Blueprint는 기존 Composer를 **확장**한다. 새 레이아웃 함수 + 새 앵커 컴포넌트를 추가하되, 기존 코드는 건드리지 않는다.

### 3.2 구현 계획

#### Phase A: Composer 레이아웃 확장 (`composer.py`)

기존 8개 레이아웃에 **7개 비직사각형 레이아웃** 추가:

```python
# 새 레이아웃 함수들
def _center_peripheral_4(r: Region, center_ratio=0.35, gap=0.15) -> dict[str, Region]:
    """중앙 + 상하좌우 4개 존.
    
    Returns: {"center", "top", "right", "bottom", "left"}
    
    대응 CP: CP1 (다이아몬드/도넛 + 주변 텍스트)
    """
    cw = r.w * center_ratio
    ch = r.h * center_ratio
    cx = r.x + (r.w - cw) / 2
    cy = r.y + (r.h - ch) / 2
    side_w = (r.w - cw) / 2 - gap
    side_h = (r.h - ch) / 2 - gap
    return {
        "center": Region(cx, cy, cw, ch),
        "top":    Region(r.x, r.y, r.w, side_h),
        "bottom": Region(r.x, cy + ch + gap, r.w, side_h),
        "left":   Region(r.x, cy, side_w, ch),
        "right":  Region(cx + cw + gap, cy, side_w, ch),
    }


def _center_peripheral_6(r: Region, center_ratio=0.30, gap=0.1) -> dict[str, Region]:
    """중앙 + 6개 주변 존 (좌3 + 우3).
    
    Returns: {"center", "tl", "ml", "bl", "tr", "mr", "br"}
    
    대응 CP: CP1 (헥사곤 6단계 + 주변 설명)
    """
    cw = r.w * center_ratio
    ch = r.h * 0.7
    cx = r.x + (r.w - cw) / 2
    cy = r.y + (r.h - ch) / 2
    side_w = (r.w - cw) / 2 - gap
    row_h = (r.h - gap * 2) / 3
    return {
        "center": Region(cx, cy, cw, ch),
        "tl": Region(r.x, r.y, side_w, row_h),
        "ml": Region(r.x, r.y + row_h + gap, side_w, row_h),
        "bl": Region(r.x, r.y + (row_h + gap) * 2, side_w, row_h),
        "tr": Region(cx + cw + gap, r.y, side_w, row_h),
        "mr": Region(cx + cw + gap, r.y + row_h + gap, side_w, row_h),
        "br": Region(cx + cw + gap, r.y + (row_h + gap) * 2, side_w, row_h),
    }


def _grid_nxm(r: Region, rows=2, cols=3, gap=0.1) -> dict[str, Region]:
    """N×M 균등 그리드.
    
    Returns: {"r0c0", "r0c1", ..., "r{n}c{m}"}
    
    대응 CP: CP2 (번호+색상 그리드)
    """
    cell_w = (r.w - gap * (cols - 1)) / cols
    cell_h = (r.h - gap * (rows - 1)) / rows
    return {
        f"r{ri}c{ci}": Region(
            r.x + ci * (cell_w + gap),
            r.y + ri * (cell_h + gap),
            cell_w, cell_h
        )
        for ri in range(rows) for ci in range(cols)
    }


def _timeline_band(r: Region, steps=5, band_ratio=0.08, gap=0.08) -> dict[str, Region]:
    """중앙 타임라인 밴드 + 상하 교차 콘텐츠 존.
    
    Returns: {"band", "above_0", "below_1", "above_2", ...}
    
    대응 CP: CP3 (타임라인 + 상하 교차)
    """
    band_h = r.h * band_ratio
    band_y = r.y + (r.h - band_h) / 2
    step_w = (r.w - gap * (steps - 1)) / steps
    above_h = band_y - r.y - gap
    below_h = r.y + r.h - (band_y + band_h) - gap
    
    zones = {"band": Region(r.x, band_y, r.w, band_h)}
    for i in range(steps):
        sx = r.x + i * (step_w + gap)
        if i % 2 == 0:  # 짝수: 아래
            zones[f"step_{i}"] = Region(sx, band_y + band_h + gap, step_w, below_h)
        else:  # 홀수: 위
            zones[f"step_{i}"] = Region(sx, r.y, step_w, above_h)
    return zones


def _asymmetric_lr(r: Region, left_ratio=0.45, gap=0.15) -> dict[str, Region]:
    """비대칭 좌우 분할 — 좌측에 도표/차트, 우측에 해설.
    
    left를 다시 sub-zone으로 나눌 수 있도록 Region.sub() 활용.
    
    대응 CP: CP5 (차트 + 상세 해설)
    """
    lw = (r.w - gap) * left_ratio
    rw = r.w - lw - gap
    return {
        "diagram": Region(r.x, r.y, lw, r.h),
        "annotation": Region(r.x + lw + gap, r.y, rw, r.h),
    }


def _t_layout(r: Region, top_ratio=0.35, right_ratio=0.4, gap=0.12) -> dict[str, Region]:
    """T자 레이아웃 — 상단 전폭 + 하단 좌우 분할.
    
    대응 CP: CP3 변형 (타임라인 밴드 위 + 아래 좌우 분석)
    """
    top_h = (r.h - gap) * top_ratio
    bottom_h = r.h - top_h - gap
    lw = (r.w - gap) * (1 - right_ratio)
    rw = r.w - lw - gap
    return {
        "top": Region(r.x, r.y, r.w, top_h),
        "bottom_left": Region(r.x, r.y + top_h + gap, lw, bottom_h),
        "bottom_right": Region(r.x + lw + gap, r.y + top_h + gap, rw, bottom_h),
    }


def _l_layout(r: Region, left_ratio=0.35, top_ratio=0.5, gap=0.12) -> dict[str, Region]:
    """L자 레이아웃 — 좌측 전높이 + 우측 상하 분할.
    
    대응 CP: CP4 (사이드바 KPI + 우측 차트/해설)
    """
    lw = (r.w - gap) * left_ratio
    rw = r.w - lw - gap
    top_h = (r.h - gap) * top_ratio
    bottom_h = r.h - top_h - gap
    return {
        "left_full": Region(r.x, r.y, lw, r.h),
        "right_top": Region(r.x + lw + gap, r.y, rw, top_h),
        "right_bottom": Region(r.x + lw + gap, r.y + top_h + gap, rw, bottom_h),
    }
```

#### Phase B: 앵커 컴포넌트 (`components.py` 확장)

기존 컴포넌트에 **앵커 역할 컴포넌트** 6개 추가:

```python
def comp_diamond_anchor(
    canvas: Canvas,
    *,
    labels: list[str],        # ["01","02","03","04"]
    section_titles: list[str], # 각 변에 붙는 라벨
    center_text: str = "",
    region: Region,
) -> None:
    """다이아몬드 4분할 중앙 도형.
    
    대응 CP: CP1 (B05, B06)
    도형 자체만 렌더하고, 주변 텍스트는 Composer가 zone에 배치.
    """
    ...

def comp_hexagon_anchor(
    canvas: Canvas,
    *,
    labels: list[str],        # ["01"~"06"]
    center_text: str = "",
    style: str = "filled",    # "filled" | "outlined"  
    region: Region,
) -> None:
    """헥사곤 6분할 중앙 도형.
    
    대응 CP: CP1 (B03, B04)
    """
    ...

def comp_donut_anchor(
    canvas: Canvas,
    *,
    segments: list[dict],     # [{"label": "PwC", "value": 40, "color": "accent"}, ...]
    center_text: str = "",
    region: Region,
) -> None:
    """도넛 차트 앵커 — 중앙 라벨 + 세그먼트.
    
    대응 CP: CP5 (B11 Alliance 다이어그램)
    """
    ...

def comp_numbered_cell(
    canvas: Canvas,
    *,
    number: str,              # "01", "02", ...
    header: str,
    body: str,
    bg_color: str = "white",  # 그라데이션용 색상
    number_size: int = 36,
    region: Region,
) -> None:
    """번호+색상 코딩된 그리드 셀.
    
    대응 CP: CP2 (B00, B01)
    """
    ...

def comp_timeline_marker(
    canvas: Canvas,
    *,
    labels: list[str],        # ["2021", "2022", ...]
    style: str = "arrow",     # "arrow" | "dots" | "chevron"
    highlight_idx: int = -1,  # 강조할 단계 인덱스
    region: Region,
) -> None:
    """타임라인 밴드 마커.
    
    대응 CP: CP3 (B08, B10)
    zone에 배치되며, 상하 콘텐츠 존은 별도 컴포넌트가 채움.
    """
    ...

def comp_icon_header_card(
    canvas: Canvas,
    *,
    icon: str,                # 아이콘 이름
    header: str,
    body: str,
    icon_size: float = 0.5,
    region: Region,
) -> float:
    """아이콘 + 헤더 + 본문 카드.
    
    대응 CP: CP6 (B07, B09)
    기존 comp_icon_card와 유사하나 region 안에서 더 유연한 배치.
    """
    ...
```

#### Phase C: Blueprint 사용 예시 (Claude Code가 생성할 코드)

```python
# 예시 1: CP1 — 다이아몬드 + 4방향 텍스트 (B06 재현)
composer = SlideComposer(slide)
composer.header(SlideHeader(title="Four segments", category="Strategy"))

zones = composer.layout("center_peripheral_4", center_ratio=0.38)

comp_diamond_anchor(composer.canvas,
    labels=["01","02","03","04"],
    section_titles=["Section title"]*4,
    center_text="Our goal\nlorem ipsum",
    region=zones["center"])

for pos in ["top", "right", "bottom", "left"]:
    comp_bullet_list(composer.canvas,
        header="Header",
        items=["Lorem ipsum dolor sit amet..."],
        region=zones[pos])


# 예시 2: CP2 — 6칸 번호 그리드 (B00 재현)
composer = SlideComposer(slide)
composer.header(SlideHeader(title="Process", category="Operations"))

zones = composer.layout("grid_nxm", rows=2, cols=3, gap=0.0)

colors = ["grey_200", "grey_100", "accent_light", 
          "accent_mid", "accent", "accent"]
for i in range(6):
    comp_numbered_cell(composer.canvas,
        number=f"{i:02d}",
        header="Header",
        body="Lorem ipsum dolor sit amet...",
        bg_color=colors[i],
        region=zones[f"r{i//3}c{i%3}"])


# 예시 3: CP3 — 타임라인 + 교차 콘텐츠 (B08 재현)
composer = SlideComposer(slide)
composer.header(SlideHeader(title="Timeline", category="Roadmap"))

zones = composer.layout("timeline_band", steps=5)

comp_timeline_marker(composer.canvas,
    labels=["01","02","03","04","05"],
    style="dots",
    region=zones["band"])

for i in range(5):
    comp_bullet_list(composer.canvas,
        header="Header",
        items=["Lorem ipsum dolor sit amet..."],
        region=zones[f"step_{i}"])


# 예시 4: CP4 — 이종 패널 (A01 재현: 파이+KPI+카드)
composer = SlideComposer(slide)

zones = composer.layout("l_layout", left_ratio=0.45, top_ratio=0.55)

comp_native_chart(composer.canvas,
    chart_type="pie",
    data={"2025": [40, 25, 20, 15]},
    region=zones["left_full"])

comp_kpi_card(composer.canvas,
    value="20x", label="Text with Key Figure",
    region=zones["right_top"])

comp_styled_card(composer.canvas,
    header="6,176", body="비교 데이터 : 수치",
    style="accent_block",
    region=zones["right_bottom"])


# 예시 5: CP5 — 도넛 + 상세 해설 (B11 재현)
composer = SlideComposer(slide)
composer.header(SlideHeader(title="Alliance", category="Partnership"))

zones = composer.layout("center_peripheral_6")

comp_donut_anchor(composer.canvas,
    segments=[
        {"label": "PwC Alliance", "value": 45, "color": "dark"},
        {"label": "Consortium", "value": 30, "color": "accent"},
        {"label": "Synergy", "value": 25, "color": "accent_mid"},
    ],
    center_text="PwC Alliance\nConsortium\nSynergy",
    region=zones["center"])

# 좌측 3존, 우측 3존에 각각 다른 내용
for i, pos in enumerate(["tl","ml","bl","tr","mr","br"]):
    comp_bullet_list(composer.canvas,
        header=partner_headers[i],
        items=partner_details[i],
        region=zones[pos])
```

### 3.3 파일별 변경 상세

| 파일 | 변경 유형 | 상세 |
|------|---------|------|
| `ppt_builder/composer.py` | **확장** | 7개 새 레이아웃 함수 + LAYOUTS 레지스트리 추가. 기존 코드 수정 없음 |
| `ppt_builder/components.py` | **확장** | 6개 앵커 컴포넌트 추가 (comp_diamond_anchor 등). 기존 코드 수정 없음 |
| `ppt_builder/patterns.py` | **수정 없음** | 기존 25개 패턴은 그대로 유지. 단순 슬라이드에 계속 사용 |
| `ppt_builder/primitives.py` | **소규모 확장** | Region.center(), Region.margin() 편의 메서드 추가 가능 |
| `docs/slide_designer.md` | **업데이트** | 6개 CP 패턴 + 사용 가이드 추가 |
| `ppt_builder/composer.py` RECIPES | **확장** | 6개 새 레시피 추가 (CP1~CP6 대응) |

### 3.4 새 COMPOSITION_RECIPES (composer.py에 추가)

```python
NEW_RECIPES = {
    "central_diagram_4": {
        "description": "중앙 다이아몬드/도넛 + 4방향 텍스트 블록",
        "layout": "center_peripheral_4",
        "layout_params": {"center_ratio": 0.38},
        "zones": {
            "center": {"component": "diamond_anchor OR donut_anchor", "tone": "accent"},
            "top/right/bottom/left": {"component": "bullet_list OR icon_header_card", "tone": "light"},
        },
        "when": "4대 전략, 4분면 분석, 핵심 가치 등 중앙 집중형 메시지",
    },
    "central_diagram_6": {
        "description": "중앙 헥사곤/원형 + 6방향 텍스트 블록",
        "layout": "center_peripheral_6",
        "layout_params": {"center_ratio": 0.30},
        "zones": {
            "center": {"component": "hexagon_anchor", "tone": "accent"},
            "tl/ml/bl/tr/mr/br": {"component": "bullet_list", "tone": "light"},
        },
        "when": "6단계 프로세스, 6대 역량, 순환 구조 등",
    },
    "numbered_process_grid": {
        "description": "N×M 번호+색상 코딩 그리드",
        "layout": "grid_nxm",
        "layout_params": {"rows": 2, "cols": 3, "gap": 0.0},
        "zones": {
            "r{i}c{j}": {"component": "numbered_cell", "tone": "gradient"},
        },
        "when": "다단계 프로세스, 방법론 개요, 프레임워크 소개",
    },
    "timeline_zigzag": {
        "description": "타임라인 밴드 + 상하 교차 콘텐츠",
        "layout": "timeline_band",
        "layout_params": {"steps": 5, "band_ratio": 0.08},
        "zones": {
            "band": {"component": "timeline_marker", "tone": "accent"},
            "step_{i}": {"component": "bullet_list OR icon_header_card", "tone": "light"},
        },
        "when": "로드맵, 마일스톤, 연도별 계획 등 시간축 중심 메시지",
    },
    "heterogeneous_panels": {
        "description": "이종 패널 — 차트+KPI+표 등 서로 다른 유형 조합",
        "layout": "l_layout OR t_layout",
        "layout_params": {"varies": True},
        "zones": {
            "각 zone": {"component": "ANY — 차트, KPI, 표, 불릿 자유 배치", "tone": "varies"},
        },
        "when": "데이터 대시보드, 종합 분석, 여러 관점 동시 제시",
    },
    "diagram_annotated": {
        "description": "대형 도표 + 상세 해설 텍스트",
        "layout": "asymmetric_lr",
        "layout_params": {"left_ratio": 0.45},
        "zones": {
            "diagram": {"component": "donut_anchor OR native_chart", "tone": "light"},
            "annotation": {"component": "bullet_list_stacked", "tone": "light"},
        },
        "when": "컨소시엄 구조, 조직도+역할 설명, 아키텍처+상세",
    },
}
```

---

## 4. 구현 우선순위

### Phase 1 (즉시 — 가장 높은 ROI)

| 순서 | 작업 | 예상 규모 | 효과 |
|------|------|---------|------|
| 1-1 | `composer.py`에 7개 새 레이아웃 추가 | ~120줄 | 모든 CP의 존 분할 가능 |
| 1-2 | `comp_numbered_cell` 구현 | ~40줄 | CP2 (번호 그리드) 즉시 가능 |
| 1-3 | `comp_timeline_marker` 구현 | ~50줄 | CP3 (타임라인 밴드) 즉시 가능 |
| 1-4 | `comp_icon_header_card` 구현 | ~35줄 | CP6 (아이콘 강화 카드) 즉시 가능 |

### Phase 2 (다음 — 도형 앵커)

| 순서 | 작업 | 예상 규모 | 효과 |
|------|------|---------|------|
| 2-1 | `comp_diamond_anchor` 구현 | ~70줄 | CP1 다이아몬드 변형 |
| 2-2 | `comp_hexagon_anchor` 구현 | ~80줄 | CP1 헥사곤 변형 |
| 2-3 | `comp_donut_anchor` 구현 | ~60줄 | CP5 도넛+해설 |
| 2-4 | 새 COMPOSITION_RECIPES 추가 | ~60줄 | Claude Code 가이드 |

### Phase 3 (후속 — 품질 향상)

| 순서 | 작업 | 예상 규모 | 효과 |
|------|------|---------|------|
| 3-1 | Region.center(), .margin(), .inset() 편의 메서드 | ~20줄 | 코드 가독성 |
| 3-2 | slide_designer.md에 CP1~CP6 가이드 추가 | 문서 | Claude Code 판단력 향상 |
| 3-3 | 색상 그라데이션 유틸 (그리드 셀 색상 자동 계산) | ~30줄 | CP2 색상 품질 |
| 3-4 | 전체 덱 테스트 (넷제로 10장 재생성) | 실행 | 실증 검증 |

---

## 5. 설계 원칙

1. **기존 코드 파괴 금지** — 25개 패턴, 20+개 컴포넌트, 8개 레이아웃 모두 그대로 작동
2. **Composer 확장, 대체 아님** — 새 레이아웃을 LAYOUTS 레지스트리에 추가만 함
3. **앵커 컴포넌트는 순수 컴포넌트** — Region-aware, Canvas만 사용, 패턴과 독립
4. **Claude Code가 판단** — 코드에는 옵션만 제공, "어떤 CP를 쓸지"는 Claude가 콘텐츠 보고 결정
5. **점진적 가치 제공** — Phase 1만으로도 CP2, CP3, CP4, CP6 구현 가능

---

## 6. 기대 효과

### Before (현재)
```
슬라이드 1: executive_summary (단일 패턴)
슬라이드 2: timeline_phases (단일 패턴)  
슬라이드 3: kpi_dashboard (단일 패턴)
→ 10장이 다 "깔끔한 단일 패턴", 밀도/다양성 부족
```

### After (Blueprint 적용)
```
슬라이드 1: center_peripheral_4 + diamond_anchor + bullet_lists
슬라이드 2: timeline_band + timeline_marker + icon_header_cards
슬라이드 3: l_layout + native_chart(pie) + kpi_card + styled_card
→ 10장 각각이 복수 요소의 유기적 조합, 컨설팅 밀도 달성
```

### 정량 목표
- **레이아웃 다양성**: 8종 → 15종 (+87%)
- **구현 가능 CP**: 0/6 → 6/6 (100%)
- **슬라이드당 컴포넌트 수**: 평균 1~2개 → 3~5개
- **코드 변경 규모**: ~500줄 추가 (기존 ~3,800줄 대비 13%)
