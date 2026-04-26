# PPT Generator - Consulting Grade

## Project Summary
컨설팅 품질의 PowerPoint 프레젠테이션을 생성하는 렌더링 라이브러리.
**Claude Code가 두뇌(리서치, 구조화, 품질검증)**, 코드는 도구(렌더링).

## Architecture (핵심)
```
사용자 ↔ Claude Code (리서치/구조화/검증) → ppt_builder 라이브러리 → .pptx
```
- 별도 LLM API 없음. Claude Code 자체가 오케스트레이터
- `ppt_builder/`는 독립 라이브러리 — 나중에 웹앱에서도 재사용 가능

## Tech Stack
- **Language**: Python 3.11+
- **PPT Engine**: python-pptx
- **Charts**: matplotlib (PNG export)
- **Validation**: Pydantic 2.x
- **Orchestrator**: Claude Code (리서치, 구조화, 실행, 검증)

## Current Phase
**Phase B 완료 (2026-04-26)** — 운영 hardening 6 작업 완료.
- B1 Pydantic 스키마: `ppt_builder/models/scenario.py` (ScenarioInput / ChartSpec)
- B2 차트 자동 통합: `chart_data` → `replace_chart_data` 자동 호출 + chart_penalty 해제
- B3 라이브러리 진입점 + CLI: `ppt_builder/api.py::build_scenario` + `run.py`
- B4 에러 핸들링 + verbose: BuildError 메시지 + plan dump + `--render-pngs`
- B5 카탈로그 viewer: `python -m ppt_builder.catalog_view {summary,roles,scenarios,components}`
- B6 README + USAGE.md: 빠른 시작 + JSON 입력 가이드 + 데이터 재빌드 절차
- 단위테스트 27/27 (scenario_schema 11 + api_errors 4 + component_ops 12)

**Phase A3 (2026-04-26 동결)** — composite 88.7 (new) / 78.7 (legacy), 80+ 4/5, role 100%, visual 100%.
보고서: **[docs/PHASE_A3_FINAL_REPORT.md](docs/PHASE_A3_FINAL_REPORT.md)**.

**다음 후보** (사용자 결정 필요):
- 컴포넌트 라이브러리 확장 (kpi / icon_grid / divider 추가 추출/디자인)
- 차트 series style 정리 (현재 마스터 디폴트 회색이라 시각적으로 흐림)
- 새 마스터 템플릿 추가 (사용자 자산 입수 시) → 헌법 §6 trigger
- N1-Full 일부 도입 검토 — 현재 §6 미트리거

## ⚠️ 방향 확정 (2026-04-25) — 흔들지 않음
**[docs/PROJECT_DIRECTION.md](docs/PROJECT_DIRECTION.md)** = 단일 진실 소스 (SSoT).
- **Mode A** (whole-slide reuse) 백본 + **N1-Lite** (single-component reuse) 보조
- 거부된 옵션: N1-Full (자동 컴포넌트 조합), P3 (회사PC 추가 캡처), Mode B 단독
- 변경은 헌법 §6 변경 조건 충족 + 사용자 명시적 결정 시에만

**[docs/N1_LITE_IMPLEMENTATION.md](docs/N1_LITE_IMPLEMENTATION.md)** = 세부 구현 사양.

새 세션 시작 시 위 두 문서 + `docs/_research/` 리서치 보고서 2개 필독.

## Document Hierarchy

| 문서 | 역할 | 변경 빈도 |
|---|---|---|
| `CLAUDE.md` (이 파일) | 현재 상태 + 네비게이션 허브 | 매 세션 |
| **`docs/PROJECT_DIRECTION.md`** | **방향 확정 헌법 (SSoT)** | **거의 불변** |
| **`docs/N1_LITE_IMPLEMENTATION.md`** | **N1-Lite 세부 구현 사양** | **사양 변경 시** |
| `docs/_research/component_extraction_research.md` | python-pptx 컴포넌트 추출 SOTA | 거의 불변 |
| `docs/_research/composition_strategy_research.md` | 조합 전략 SOTA + role 매핑 | 거의 불변 |
| `docs/01_VISION.md` | 왜 만드는가 - 핵심 원칙과 북극성 | 거의 불변 |
| `docs/02_PRD.md` | 무엇을 만드는가 - 기능 범위, 슬라이드 타입 | 드물게 변경 |
| `docs/03_TECH_SPEC.md` | 어떻게 만드는가 - 아키텍처, 폴더 구조 | 가끔 변경 |
| `docs/04_DECISION_LOG.md` | 결정 이력 | 추가만 (append-only) |
| `docs/05_SLIDE_TYPES.md` | 슬라이드 타입 카탈로그 상세 정의 | 필요시 확장 |
| `docs/06_TEMPLATE_CATALOG.md` | 템플릿+컴포넌트 마스터 리스트 (50+항목) | 구현 시 업데이트 |
| `docs/07_WORKFLOW.md` | **[바이블]** PPT 생성 유일 프로세스. GATE 미통과 시 다음 단계 진행 금지 | 사용자 합의 시만 변경 |
| `docs/slide_designer.md` | Claude Code용 화면 구성 판단 가이드 | 패턴 추가 시 |
| `ppt_builder/template/metadata.json` | 35개 템플릿 메타데이터 (태깅) | 템플릿 추가 시 |
| `ppt_builder/evaluate.py` | 자동 품질 평가 (Step 5) | 평가 기준 변경 시 |

## Document Maintenance Rules
- **PPT 생성 시** → `docs/07_WORKFLOW.md`(바이블)의 6+1단계를 **반드시** 순서대로 따름. GATE 미통과 시 다음 단계 진행 금지
- **07_WORKFLOW.md 변경** → 사용자와 명시적 합의 후에만 변경. Claude Code가 임의로 예외/완화를 추가하지 않음
- **기능 추가/변경** → `02_PRD.md` 업데이트 + `04_DECISION_LOG.md` 추가
- **기술 결정** → `03_TECH_SPEC.md` 업데이트 + `04_DECISION_LOG.md` 추가
- **코드 작성 완료** → `CLAUDE.md` 현재 상태 섹션 업데이트
- **01_VISION.md** → 명시적 피벗 결정 없이 절대 수정하지 않음
- **04_DECISION_LOG.md** → 추가만 가능, 삭제/수정 금지

## Folder Structure
```
PPT/
├── CLAUDE.md
├── docs/
│   ├── 01_VISION.md
│   ├── 02_PRD.md
│   ├── 03_TECH_SPEC.md
│   ├── 04_DECISION_LOG.md
│   ├── 05_SLIDE_TYPES.md
│   └── references/              ← 레퍼런스 캡처 이미지
├── ppt_builder/                 ← 핵심 렌더링 라이브러리 (독립적)
│   ├── __init__.py              ← render_presentation() 공개 API
│   ├── models/                  ← Pydantic 스키마
│   ├── assembler/               ← 슬라이드 타입별 렌더러
│   └── charts/                  ← matplotlib 차트 생성
├── templates/                   ← .pptx 마스터 템플릿
├── output/                      ← 생성된 PPT
├── tests/
├── requirements.txt
└── run.py                       ← CLI 진입점
```

## Conventions
- 모든 슬라이드 타이틀은 **액션 타이틀** (인사이트 문장형, 라벨 아님)
- 슬라이드 스키마는 JSON으로 정의, Pydantic 모델로 검증
- 한글/영문 혼용 가능, 코드와 변수명은 영문
- `ppt_builder/`는 Claude Code에 종속되지 않는 순수 라이브러리로 유지
