# PPT Generator — Consulting Grade

컨설팅 품질의 PowerPoint를 생성하는 렌더링 라이브러리. Claude Code가 두뇌
(리서치, 구조화, 품질 검증), 코드는 도구(렌더링).

> **현재 상태**: Phase A3 완료 (2026-04-26) — 5 시나리오 평균 composite **88.7**
> (legacy 78.7), role 100%, visual 100%. 80+ 4/5 시나리오.
> 동결 보고서: [docs/PHASE_A3_FINAL_REPORT.md](docs/PHASE_A3_FINAL_REPORT.md)

> **방향**: Mode A + N1-Lite 하이브리드 (헌법 [docs/PROJECT_DIRECTION.md](docs/PROJECT_DIRECTION.md)).
> Mode A = 마스터 슬라이드 통째 복제, N1-Lite = 컴포넌트 단독 재사용.
> 자동 컴포넌트 조합(N1-Full)은 **거부됨** (외부 SOTA 미해결).

## 빠른 시작

```bash
# 사전 정의된 시나리오 빌드
python run.py --preset analysis_report_15 --output out/q1.pptx

# JSON 입력 파일로 빌드
python run.py --input my_scenario.json --output out/my_deck.pptx

# 시나리오 목록
python run.py --list-presets

# 어떤 role/컴포넌트가 사용 가능한지 확인
python -m ppt_builder.catalog_view summary
python -m ppt_builder.catalog_view roles
python -m ppt_builder.catalog_view components [family]
```

자세한 사용법: [USAGE.md](USAGE.md).

## 프로젝트 구조

```
PPT/
├── CLAUDE.md                    Claude Code 진입점 + 현재 상태
├── README.md / USAGE.md         (이 문서)
├── run.py                       CLI 진입점
├── docs/
│   ├── PROJECT_DIRECTION.md     ⭐ 방향 헌법 (SSoT)
│   ├── N1_LITE_IMPLEMENTATION.md  N1-Lite 세부 사양
│   ├── PHASE_A3_FINAL_REPORT.md  A3 동결 결과
│   ├── 01_VISION.md ~ 07_WORKFLOW.md  프로세스/스펙
│   └── _research/               외부 SOTA 보고서
├── ppt_builder/                 핵심 렌더링 라이브러리
│   ├── api.py                   build_scenario(scenario) — 운영 진입점
│   ├── models/scenario.py       ScenarioInput Pydantic 스키마
│   ├── catalog_view.py          카탈로그 CLI viewer
│   ├── template/
│   │   ├── component_ops.py     extract_group / insert_component / replace_chart_data
│   │   ├── edit_ops.py          5 편집 API (PPTAgent 이식)
│   │   └── editor.py            keep_slides / 마스터 조작
│   └── catalog/                 라벨링/추출/검색 인프라
├── scripts/
│   ├── benchmark_5_scenarios_v6.py  v6 빌드 파이프라인 (운영 reference)
│   └── build_component_library.py   컴포넌트 라이브러리 빌더
├── output/
│   ├── component_library/       33 컴포넌트 / 7 family + 미리보기 .pptx
│   ├── benchmark_v6/            5 시나리오 산출 + 시각 검수
│   ├── catalog/                 final_labels.json / skeletons.json / paragraph_labels.json
│   └── decks/                   run.py 기본 출력 위치
└── docs/references/_master_templates/
    └── PPT 템플릿.pptx           1,251장 마스터 (편집 자산)
```

## 핵심 API

### Python 라이브러리

```python
from pathlib import Path
from ppt_builder.api import build_scenario
from ppt_builder.models import ScenarioInput, ChartSpec, ChartSeriesSpec

scenario = ScenarioInput(
    scenario_name="우리 회사 Q1 검토",
    skeleton_id="analysis_report_15",
    content_by_role={
        "opening": ["Q1 실적 검토"],
        "evidence": ["매출 4.2조 (+6.3% YoY)", "원가율 73.8%"],
        "recommendation": ["원자재 장기계약 비중 35%→50%"],
        "closing": ["Q2 회복 시나리오"],
    },
    chart_data={
        "evidence": ChartSpec(
            categories=["Q1", "Q2", "Q3", "Q4"],
            series=[
                ChartSeriesSpec(
                    name="매출(조)", values=[4.0, 4.2, 4.4, 4.5],
                    color="#D04A02",  # PwC accent (옵션)
                )
            ],
        )
    },
)
result = build_scenario(scenario, output=Path("output/q1.pptx"))
print(result.summary_line())
# → q1.pptx: 15 slides (12 mode_a + 3 n1_lite), charts: 1/1
```

### narrative role 13개 (헌법 §3)

| role           | mode     | 설명                             |
|----------------|----------|----------------------------------|
| opening        | mode_a   | 표지                             |
| agenda         | mode_a   | 목차                             |
| divider        | mode_a   | 섹션 구분                        |
| situation      | n1_lite  | 현황 (풀 결손 → 컴포넌트 보강)    |
| complication   | n1_lite  | 문제/난점                         |
| evidence       | mode_a   | 근거/데이터                       |
| analysis       | mode_a   | 분석/진단                         |
| recommendation | mode_a   | 권고/솔루션                       |
| roadmap        | n1_lite  | 로드맵 (timeline/chevron 컴포넌트) |
| benefit        | n1_lite  | 효과/혜택                         |
| risk           | n1_lite  | 위험/리스크                       |
| closing        | mode_a   | 결론/마감                         |
| appendix       | mode_a   | 부록                             |

## 설치

```bash
pip install -r requirements.txt
```

핵심 의존성: `python-pptx>=1.0`, `pydantic>=2.0`, `matplotlib>=3.8`,
`Pillow>=10.0`, `lxml`. PNG 시각 검수는 Windows + Microsoft PowerPoint
COM (선택).

## 데이터 사전 빌드

운영 빌드는 다음 데이터 산출에 의존한다 (이미 리포지토리에 포함):

| 경로 | 내용 |
|---|---|
| `output/catalog/final_labels.json` | 1,251장 마스터의 multi-label 카탈로그 |
| `output/catalog/skeletons.json` | 시나리오별 narrative 시퀀스 |
| `output/catalog/paragraph_labels.json` | 슬롯 단위 paragraph store |
| `output/component_library/components_index.json` | 33 N1-Lite 컴포넌트 메타 |
| `output/component_library/{family}_family.pptx` × 7 | 컴포넌트 미리보기 |

새 마스터를 추가했을 때 재빌드 명령은 [USAGE.md](USAGE.md) 참고.

## 테스트

```bash
python -m pytest tests/ -v
```

핵심 단위테스트:
- `tests/test_scenario_schema.py` — Pydantic 입력 검증 (11)
- `tests/test_api_errors.py` — BuildError 경로 + 파일명 정규화 (4)
- `tests/test_component_ops.py` — extract_group / insert_component (12)
- `tests/test_edit_ops.py` — 5 편집 API (14)

## 헌법 (방향 SSoT)

다음 두 문서를 새 세션 첫 작업에 반드시 로드:

1. **[docs/PROJECT_DIRECTION.md](docs/PROJECT_DIRECTION.md)** — 방향 확정 헌법.
   §6 변경 조건 충족 + 사용자 명시적 결정 시 외에는 절대 변경 금지.
2. **[docs/N1_LITE_IMPLEMENTATION.md](docs/N1_LITE_IMPLEMENTATION.md)** — 세부 구현 사양.

## 라이센스

Internal (PwC SAP 컨설팅 자산 활용).
