# USAGE — PPT Generator 사용 가이드

> 운영 빌드 (Phase B). [README.md](README.md)와 함께 본다.
> 헌법: [docs/PROJECT_DIRECTION.md](docs/PROJECT_DIRECTION.md).

## 목차

1. [CLI 사용](#cli-사용)
2. [JSON 입력 파일 작성](#json-입력-파일-작성)
3. [Python 라이브러리 직접 호출](#python-라이브러리-직접-호출)
4. [차트 데이터 주입](#차트-데이터-주입)
5. [카탈로그 조회](#카탈로그-조회)
6. [PNG 시각 검수](#png-시각-검수)
7. [에러 처리](#에러-처리)
8. [데이터 재빌드](#데이터-재빌드-마스터-변경-시)

---

## CLI 사용

```bash
# preset 시나리오 빌드 (5개 사전 정의)
python run.py --preset analysis_report_15
python run.py --preset analysis_report_15 --output output/q1.pptx --verbose

# JSON 파일로 직접 빌드
python run.py --input my_scenario.json
python run.py -i my_scenario.json -o out/my.pptx --render-pngs

# 사용 가능한 preset 목록
python run.py --list-presets
```

### CLI 플래그

| 플래그 | 설명 |
|---|---|
| `--input` / `-i` | ScenarioInput JSON 파일 |
| `--preset` / `-p` | preset 시나리오 ID |
| `--output` / `-o` | 출력 .pptx 경로 (기본: `output/decks/{name}.pptx`) |
| `--list-presets` | preset 목록 출력 후 종료 |
| `--quiet` / `-q` | 진행 메시지 최소화 |
| `--verbose` / `-v` | plan 덤프 (어떤 슬라이드/컴포넌트 선정) |
| `--render-pngs` | 빌드 후 PNG 렌더 (Windows + PowerPoint 필요) |

---

## JSON 입력 파일 작성

```json
{
  "scenario_name": "우리 회사 Q1 검토",
  "skeleton_id": "analysis_report_15",

  "content_by_role": {
    "opening":        ["Q1 실적 검토 — CFO Office"],
    "agenda":         ["1. 손익 / 2. 분석 / 3. 권고"],
    "situation":      ["매출 4.2조원 / 영업이익 3,800억원"],
    "evidence":       ["원가 구성: 원자재 58% / 인건비 22%"],
    "analysis":       ["원자재 +14% / 환율 +210억원"],
    "recommendation": ["원자재 장기계약 35→50%"],
    "risk":           ["철광석 가격 변동성"],
    "closing":        ["Q2 회복 시나리오 권고"]
  },

  "chart_data": {
    "evidence": {
      "categories": ["Q1", "Q2", "Q3", "Q4"],
      "series": [
        {"name": "매출(조)", "values": [4.0, 4.2, 4.4, 4.5]}
      ]
    }
  }
}
```

### 필드

| 필드 | 타입 | 설명 |
|---|---|---|
| `scenario_name` | string (required) | 출력 파일명 + 표지 후보 |
| `skeleton_id` | string | 사전 정의된 narrative 시퀀스 (`narrative_sequence`와 둘 중 하나 필수) |
| `narrative_sequence` | list[role] | 직접 narrative 지정. skeleton보다 우선 |
| `content_by_role` | dict[role, list[str]] | role별 컨텐츠 |
| `chart_data` | dict[role, ChartSpec] | 차트 데이터 (선택) |
| `metadata` | object | `title` / `client` / `author` / `date` (선택) |

### narrative role

`opening`, `agenda`, `divider`, `situation`, `complication`, `evidence`,
`analysis`, `recommendation`, `roadmap`, `benefit`, `risk`, `closing`,
`appendix` (헌법 §3, 13개). 다른 문자열은 거부됨.

### narrative_sequence vs skeleton_id

- `skeleton_id`만: skeleton의 narrative 시퀀스 사용
- `narrative_sequence`만: 직접 지정한 시퀀스 사용
- 둘 다: `narrative_sequence`가 override
- 둘 다 없음: 검증 실패

---

## Python 라이브러리 직접 호출

```python
from pathlib import Path
from ppt_builder.api import build_scenario, BuildError
from ppt_builder.models import ScenarioInput, ChartSpec, ChartSeriesSpec

scenario = ScenarioInput(
    scenario_name="Q1 분석",
    narrative_sequence=["opening", "evidence", "analysis", "recommendation", "closing"],
    content_by_role={
        "opening": ["Q1 실적 검토"],
        "evidence": ["매출 4.2조 (+6.3% YoY)"],
        "analysis": ["원자재 +14% 영향"],
        "recommendation": ["장기계약 비중 ↑"],
        "closing": ["Q2 가이던스"],
    },
    chart_data={
        "evidence": ChartSpec(
            categories=["Q1", "Q2", "Q3", "Q4"],
            series=[
                ChartSeriesSpec(name="매출(조)", values=[4.0, 4.2, 4.4, 4.5]),
                ChartSeriesSpec(name="이익(억)", values=[380, 410, 450, 480]),
            ],
        )
    },
)

try:
    result = build_scenario(scenario, output=Path("output/my.pptx"))
except BuildError as e:
    print(f"빌드 실패: {e}")
else:
    print(result.summary_line())     # "my.pptx: 5 slides (4 mode_a + 1 n1_lite), charts: 1/1"
    print(result.pptx)               # Path("output/my.pptx")
    print(result.plan)               # 슬라이드별 메타데이터
    print(result.chart_injected)     # {"evidence": True}
```

### BuildResult 필드

| 필드 | 타입 | 의미 |
|---|---|---|
| `pptx` | Path | 출력 .pptx 경로 |
| `plan` | list[dict] | 슬라이드별 plan (mode/role/slide_index/component_id 등) |
| `edits` | list[dict] | 슬라이드별 채움 통계 (filled/blanked/overflow) |
| `n_mode_a` / `n_n1_lite` | int | 모드별 슬라이드 수 |
| `chart_injected` | dict[role, bool] | 차트 주입 결과 |
| `narrative` | list[str] | 실제 사용된 narrative 시퀀스 |

---

## 차트 데이터 주입

`chart_data` 키는 `narrative_role` 그대로. 빌드는 다음을 수행:

1. **선정**: 해당 role의 첫 슬라이드에 차트 슬라이드 우선 선택
   (chart_penalty 해제 → 차트 슬라이드 자체에 보너스).
2. **주입**: 빌드 후 첫 chart shape에 `replace_chart_data`로 categories + series 적용.

같은 role이 여러 step에 나오면 첫 step에만 주입. 차트 슬라이드를 못 찾으면
`chart_injected[role] = False`로 보고 (빌드는 성공).

```json
"chart_data": {
  "evidence": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [
      {"name": "매출(조)", "values": [4.0, 4.2, 4.4, 4.5]},
      {"name": "이익(억)", "values": [380, 410, 450, 480]}
    ]
  }
}
```

검증: `series[*].values` 길이가 `categories` 길이와 다르면 거부됨.

---

## 카탈로그 조회

```bash
python -m ppt_builder.catalog_view summary       # 1줄 요약
python -m ppt_builder.catalog_view roles         # 13 role × mode
python -m ppt_builder.catalog_view scenarios     # preset 목록
python -m ppt_builder.catalog_view components    # 33 컴포넌트 전체
python -m ppt_builder.catalog_view components chevron  # family 필터
```

컴포넌트 family 7종: `chevron`, `card`, `timeline`, `matrix`, `callout`,
`table`, `flow`. 각 family는 `output/component_library/{family}_family.pptx`로
미리보기. PowerPoint에서 직접 열어 확인.

---

## PNG 시각 검수

운영 검수는 **반드시** PDF/PNG 변환 후 시각 확인 (메모리: visual_check_rule).
evaluate.py 점수만 믿지 말 것.

```bash
# 빌드와 동시에 PNG 렌더
python run.py --preset analysis_report_15 --render-pngs

# 또는 Python에서
python -c "
import sys; sys.path.insert(0, 'scripts')
from benchmark_5_scenarios_v6 import render_pngs
from pathlib import Path
render_pngs(Path('output/decks/_smoke.pptx').resolve(),
            Path('output/decks/_smoke_pngs').resolve())
"
```

PNG 렌더는 Windows + Microsoft PowerPoint COM 사용 (`pywin32` 필요).
다른 OS에선 LibreOffice unoconv 등으로 대체 가능 (별도 작업).

---

## 에러 처리

### Pydantic 검증 실패

```text
입력 검증 실패:
  [content_by_role.bogus] Input should be 'opening', 'agenda', ... (literal)
  [chart_data.evidence.series.0.values] List should have at least 1 item
```

→ JSON에서 잘못된 role 이름이나 series 길이 mismatch.

### BuildError

```text
[ERROR] skeleton_id 'foo' not found. available: ['analysis_report_15', ...]
[ERROR] components_index.json 누락: Phase 2 build_component_library.py 실행 필요
[ERROR] catalog 누락: c:\Users\...\final_labels.json
```

→ 데이터 누락. [데이터 재빌드](#데이터-재빌드-마스터-변경-시) 참고.

---

## 데이터 재빌드 (마스터 변경 시)

마스터 `.pptx`를 교체하거나 추가했을 때:

```bash
# 1. label 카탈로그 재빌드 (Phase A2 작업, 시간 소요)
#    output/catalog/final_labels.json + paragraph_labels.json
#    (구체 명령은 docs/_progress 참고)

# 2. skeletons 재생성
#    output/catalog/skeletons.json

# 3. 컴포넌트 라이브러리 재빌드 (Phase 2 산출)
python scripts/build_component_library.py
#    → output/component_library/{family}_family.pptx × 7
#    → output/component_library/components_index.json

# 4. 5 시나리오 재측정 (수치 변화 확인)
python scripts/benchmark_5_scenarios_v6.py
```

마스터 변경 없이 단순 빌드만이면 1~3은 불필요.

---

## 헌법 준수

[docs/PROJECT_DIRECTION.md](docs/PROJECT_DIRECTION.md) §6의 **변경 조건**
4가지 외에는 방향 변경 금지. 코드 수정 시 헌법 §3 ROLE_MODE_MAP 변경은
사용자 명시적 결정 필요.

[docs/07_WORKFLOW.md](docs/07_WORKFLOW.md)의 6+1단계 GATE는 PPT 생성 시 반드시
순서대로 따른다 (PPT 생성 워크플로우 — 빌드 자동화는 별개).
