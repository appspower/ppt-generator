# 03. Tech Spec

> `02_PRD.md`의 기능을 코드로 구현하기 위한 기술 가이드입니다.

## Architecture Overview

### 핵심 원칙: Claude Code가 두뇌, 코드는 도구

이 프로젝트는 별도의 LLM API 레이어를 두지 않습니다.
**Claude Code 자체가 리서치, 구조화, 품질검증을 수행**하고,
python-pptx 렌더링 라이브러리를 **도구로 실행**합니다.

```
사용자 ↔ Claude Code (대화)
              │
              │  1. 리서치 (웹 검색, 파일 분석)
              │  2. 슬라이드 구조화 (JSON 스키마 생성)
              │  3. 렌더링 실행 (python 코드 실행)
              │  4. 품질 검증 & 수정 반복
              │
              ▼
    ┌─────────────────────────────────┐
    │   ppt_builder (렌더링 라이브러리)  │
    │                                 │
    │  ┌───────────┐  ┌───────────┐  │
    │  │ Assembler │  │  Charts   │  │
    │  │ (렌더러들) │  │ (차트생성) │  │
    │  └─────┬─────┘  └─────┬─────┘  │
    │        │              │         │
    │  ┌─────▼──────────────▼──────┐  │
    │  │  Slide Schema (Pydantic)  │  │
    │  └───────────────────────────┘  │
    │        │                        │
    │  ┌─────▼──────────────────────┐ │
    │  │  python-pptx Engine        │ │
    │  │  + .pptx Master Template   │ │
    │  │  + matplotlib chart images │ │
    │  └────────────────────────────┘ │
    └─────────────────┬───────────────┘
                      │
                      ▼
                 output.pptx
```

### 왜 이 구조인가
- Claude Code가 이미 웹 검색, 파일 분석, 멀티스텝 추론을 내장하고 있음
- 별도 LLM API 호출 불필요 → 비용 절감, 복잡도 감소
- 대화형 반복 수정이 가능 → 컨설팅 품질 달성에 유리
- 렌더링 라이브러리는 독립적 → 나중에 웹앱에서도 재사용 가능

## Tech Stack Detail

| Component | Technology | Version | Rationale |
|---|---|---|---|
| Runtime | Python | 3.11+ | 생태계, python-pptx 호환 |
| PPT Engine | python-pptx | 1.0.x | .pptx 네이티브, 템플릿 기반 |
| Charts | matplotlib | 3.8+ | 워터폴, 메코 등 커스텀 차트 |
| Validation | Pydantic | 2.x | 슬라이드 스키마 검증 |
| Orchestrator | Claude Code | - | 리서치, 구조화, 품질검증, 실행 |
| Image Analysis | Claude Code (Vision) | - | 레퍼런스 이미지 분석 (내장) |

> **제거된 것**: FastAPI, Streamlit, anthropic SDK, litellm
> 이것들은 나중에 웹앱이 필요할 때 추가합니다.

## Folder Structure

```
PPT/
├── CLAUDE.md                       # 프로젝트 네비게이션 허브
├── docs/                           # 프로젝트 문서
│   ├── 01_VISION.md
│   ├── 02_PRD.md
│   ├── 03_TECH_SPEC.md (이 파일)
│   ├── 04_DECISION_LOG.md
│   ├── 05_SLIDE_TYPES.md
│   └── references/                 # 레퍼런스 캡처 이미지
├── ppt_builder/                    # 핵심 렌더링 라이브러리 (독립적)
│   ├── __init__.py                 # render_presentation() 공개 API
│   ├── models/
│   │   ├── __init__.py
│   │   ├── schema.py               # Pydantic: 슬라이드 JSON 스키마
│   │   └── enums.py                # 슬라이드 타입, 차트 타입 Enum
│   ├── assembler/
│   │   ├── __init__.py
│   │   ├── engine.py               # 메인 디스패처
│   │   ├── renderers/
│   │   │   ├── __init__.py
│   │   │   ├── base.py             # 렌더러 기본 클래스
│   │   │   ├── title.py
│   │   │   ├── executive_summary.py
│   │   │   ├── chart.py
│   │   │   ├── table.py
│   │   │   ├── two_column.py
│   │   │   ├── matrix_2x2.py
│   │   │   ├── process.py
│   │   │   ├── waterfall.py
│   │   │   ├── section_divider.py
│   │   │   ├── image.py
│   │   │   └── appendix.py
│   │   └── styles.py               # 공통 스타일 상수
│   └── charts/
│       ├── __init__.py
│       ├── bar.py
│       ├── line.py
│       ├── pie.py
│       ├── waterfall.py             # matplotlib 기반
│       └── matrix.py                # shape 기반 2x2
├── templates/                       # .pptx 마스터 템플릿
│   └── default.pptx
├── output/                          # 생성된 PPT 출력
├── tests/
│   ├── test_schema.py
│   ├── test_renderers.py
│   └── fixtures/                    # 테스트용 JSON 스키마
├── requirements.txt
└── run.py                           # CLI 진입점 (스키마 JSON → PPT)
```

## Public API (진입점)

```python
# ppt_builder/__init__.py
from pathlib import Path
from .models.schema import PresentationSchema

def render_presentation(
    schema: PresentationSchema,
    template: Path = Path("templates/default.pptx"),
    output: Path = Path("output/presentation.pptx"),
) -> Path:
    """
    슬라이드 스키마를 받아서 .pptx 파일을 생성하는 순수 함수.
    
    - Claude Code가 호출할 수도 있고
    - 나중에 FastAPI/Streamlit이 호출할 수도 있음
    - 이 함수는 LLM에 의존하지 않음
    """
```

## 실제 작업 흐름 (Claude Code 중심)

```
1. 사용자: "디지털 전환 보고서 만들어줘"
   
2. Claude Code:
   - 웹 리서치로 트렌드/데이터 수집
   - 컨설팅 프레임워크 선택
   - JSON 슬라이드 스키마 생성 (05_SLIDE_TYPES.md 참조)
   - schema.json 파일로 저장

3. Claude Code:
   - python run.py schema.json 실행
   - 또는 Python 코드 직접 실행
   - output/ 폴더에 .pptx 생성

4. 사용자: "3번 슬라이드 수정해줘"
   
5. Claude Code:
   - schema.json 수정
   - 재렌더링
   - 변경 사항 설명
```

## Key Design Decisions

### 1. 슬라이드 타입 디스패처 패턴
```python
# ppt_builder/assembler/engine.py
RENDERERS = {
    SlideType.TITLE: TitleRenderer,
    SlideType.EXECUTIVE_SUMMARY: ExecSummaryRenderer,
    SlideType.CHART: ChartRenderer,
    SlideType.TABLE: TableRenderer,
    # ...
}

def assemble(prs, slide_def):
    renderer = RENDERERS[slide_def.type]
    return renderer.render(prs, slide_def)
```

### 2. 템플릿 기반 렌더링
- 항상 `.pptx` 마스터 템플릿의 슬라이드 레이아웃 사용
- 직접 텍스트박스 생성 최소화 → 플레이스홀더 활용
- 폰트/색상은 템플릿에서 상속

### 3. 차트 렌더링 전략
- 기본 차트 (막대, 선, 파이): python-pptx 네이티브 차트 객체
- 복잡 차트 (워터폴, 메코, 2x2): matplotlib → 고해상도 PNG → `add_picture()`

### 4. 렌더링 라이브러리 독립성
- `ppt_builder/`는 Claude Code, FastAPI, CLI 어디서든 import 가능
- LLM 의존성 없음 — 순수 입력(스키마) → 출력(파일) 함수
- 테스트 가능 — JSON fixture로 단위 테스트

## Naming Conventions
- 파일명: `snake_case.py`
- 클래스: `PascalCase`
- 함수/변수: `snake_case`
- 상수: `UPPER_SNAKE_CASE`
- 슬라이드 타입 키: `snake_case` (JSON, Enum)
