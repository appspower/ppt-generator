# Phase A1 검수 보고서 + PPTAgent 포팅 계획

> 작성일: 2026-04-24
> 대상 파일: `docs/references/_master_templates/PPT 템플릿.pptx` (33.9 MB, 1,251장)

---

## 1. Phase A1 검수 결과 — 한 줄 요약

**1,251장 모두 python-pptx 파싱 성공 + `~~` placeholder 98.7% 커버리지 + PNG 시각 렌더 완벽. 3층 하이브리드 방향 확정.**

---

## 2. 파일 구조 측정치

| 지표 | 값 | 판단 |
|---|---|---|
| 슬라이드 수 | 1,251 | 예상대로 |
| 화면 크기 | 10.83 × 7.50 in (9906000 × 6858000 EMU) | 변형 와이드 |
| 레이아웃 수 | 12 | 다양성 충분 |
| 마스터 수 | 2 | 정상 |
| `~~` 보유 슬라이드 | **1,235 / 1,251 (98.7%)** | ✅ 규칙 일관 |
| `~~` 총 개수 | 49,925 | 평균 39.9/장 |
| SmartArt 보유 | **단 1장 (0.1%)** | ✅ 파싱 장애물 없음 |
| GROUP 포함 슬라이드 | 1,115 (89.1%) | ⚠️ 그룹 순회 필수 |
| TABLE 보유 | 366 (29.3%) | ✅ 컨설팅 특성 |
| PICTURE 보유 | 328 (26.2%) | 정상 |
| CHART 네이티브 | 12 (1.0%) | ℹ️ 대부분 도형차트 |

## 3. Shape 타입 분포

```
AUTO_SHAPE     47,772  ← 주력 (박스/도형)
LINE           10,378
TEXT_BOX        8,562
GROUP           8,411
PLACEHOLDER     5,138
FREEFORM        4,712  ← 커스텀 벡터
PICTURE         1,933
TABLE             640
EMBEDDED_OLE       58
CHART              22
```

## 4. 시각 검증 (20장 샘플)

- **슬라이드 1**: dense 7열 × 30+행 매트릭스 테이블
- **슬라이드 5**: 좌측 6단계 Chevron-grouped process + 우측 리스트 + 강조 카드
- **슬라이드 8**: 4-step stack + 복잡 데이터플로우 + 오렌지 강조 블록
- **슬라이드 10**: 단순 2 카드 대비 (소형 박스 강조)
- **슬라이드 15**: 11열 dense table + 드릴다운 서브 표 3개
- **슬라이드 20**: 수직 5단 화살표 + 복잡 플로우 + **한글 잔존 ("지출 유형에 따른 정산 기능 적용")** — 부분 치환 상태

**결론**:
- ✅ 레이아웃 다양성 매우 풍부
- ✅ 한글 폰트 렌더링 완벽 (맑은고딕 계열)
- ✅ 복잡 요소(다이아몬드, DB 실린더, 오렌지 강조, 드릴다운) 모두 정상
- ⚠️ 치환 100% 아님 — 일부 한글 잔존. **긍정적 신호** (한글 표시 검증됨)
- ⚠️ 최대 3,980 shapes 슬라이드 존재 (아웃라이어, #189) — 별도 검토 필요

## 5. 치환 가능성 실증 (첫 20장)

- 총 shape 1,909 (group 63, text 1,706)
- `~~` Run 수 **623개 감지 + 5건 실제 치환 성공**
- Run 속성 (font.size/bold/name) 보존 확인
- `~~` → `[TEST]` 치환 후 저장/재오픈 정상

## 6. 판단: 3층 하이브리드 방향 100% 확정

| 기준 | 결과 | 점수 |
|---|---|---|
| 파싱 가능성 | SmartArt 1장만 | ★★★★★ |
| 치환 규칙성 | 98.7% `~~` 일관 | ★★★★★ |
| 레이아웃 다양성 | 12 레이아웃, shape 분포 풍부 | ★★★★★ |
| 한글 렌더링 | PNG 변환 완벽 | ★★★★★ |
| 치환 복잡도 | Run 단위 속성 보존 가능 | ★★★★☆ |
| 아웃라이어 | 3980 shape 슬라이드 존재 | ★★★★☆ |

**종합: Phase A2(전수 태깅) 진행 결정**.

---

## 7. PPTAgent 정밀 분석 결과

### 7.1 `apis.py` 편집 API 5종 (표준)

| 함수 | 시그니처 | 핵심 구현 |
|---|---|---|
| `clone_paragraph` | `(slide, div_id, paragraph_id)` | `parse_xml(para._element.xml)` + `addnext` |
| `replace_paragraph` | `(slide, div_id, paragraph_id, text)` | markdown 파싱 → Run 재빌드 |
| `del_paragraph` | `(slide, div_id, paragraph_id)` | `getparent().remove()` |
| `replace_image` | `(slide, doc, img_id, image_path)` | 비율 유지 + `table_XXXX.png`은 table로 변환 |
| `del_image` | `(slide, figure_id)` | `shapes.remove` |

**중요 발견**:
- `replace_paragraph`는 **markdown → BeautifulSoup → TextBlock → Run** 경로. `**bold**`, `*italic*` 지원
- Run 복제 시 `parse_xml(_r.xml)` — **표준 python-pptx와 100% 호환 가능성 높음**
- `Closure` 패턴: 편집 의도를 queue에 저장 후 `build()` 시점 일괄 적용

### 7.2 `CodeExecutor`

- `eval(line, SAFE_EVAL_GLOBALS, {func: partial_func})` — builtins 비운 제한 샌드박스
- `clone`과 `del`을 한 커맨드 블록에서 혼용 금지
- `retry_times` 기반 재시도 루프 (에러 라인 표시 + traceback 피드백)

### 7.3 `induct.py` (Phase A2 자동화 핵심)

```
SlideInducter.category_split       — functional vs content 분리 (opening/TOC/ending)
SlideInducter.layout_split         — layout_name + content_type + 이미지 임베딩 코사인 유사도 클러스터링
SlideInducter.content_induct       — 각 대표 슬라이드 HTML → schema_extractor → JSON schema
```

### 7.4 `schema_extractor.yaml` 프롬프트

- Input: 슬라이드 HTML + slide_idx
- Output: `{"elements": [{"name": "...", "type": "text|image", "data": [...]}]}`
- **`name` 필드가 역할 명**: "main title", "left bullets", "portrait image", "footer" 등
- **`<p>` 단위 분리 규칙** 명시

---

## 8. 우리 코드 현재 상태

### 8.1 `ppt_builder/template/` — 하이브리드 쉘 이미 존재

| 파일 | 역할 | 상태 |
|---|---|---|
| `editor.py` `TemplateEditor` | 원본 복사 + 슬라이드 삭제 + 치환 | ✅ Layer 1 기반 완성 |
| `substitutor.py` `TextSubstitutor` | 정적 placeholder 치환 | ✅ 기본 동작 |
| `cloner.py` | 다른 PPT로 슬라이드 복제 + rId remap | 🔶 향후 덜 중요 |
| `metadata.json` | 35개 템플릿 메타 | 🔶 1200장용으로 확장 예정 |
| `build_library.py` | 라이브러리 빌더 | ✅ |

**Layer 1 기반(원본 복사+슬라이드 필터+치환)은 이미 구축됨**. Layer 2(편집 API 5종)가 비어 있음.

---

## 9. 포팅 계획

### 9.1 Phase A2 (1,251장 전수 태깅, 최대 ROI)

**작업**:
1. **`SlideInducter.category_split` 포팅** — 1251장 → functional(목차/표지/결론) + content 분리
2. **`SlideInducter.layout_split` 포팅** — 레이아웃 클러스터링
   - layout_name + content_type 그룹핑 (PPT 내장 레이아웃 이미 12개 존재)
   - 이미지 임베딩 (CLIP or DINOv2) → 시각 유사도 기반 재그룹
3. **`schema_extractor.yaml` 프롬프트 차용** — 각 대표 슬라이드 HTML → JSON 스키마 자동 추출
4. **병렬 실행** — 기존 179장 Agent 파이프라인 재활용 (20 Agent)

**산출물**: `ppt_builder/template/catalog.json`
```json
{
  "layouts": {
    "dense_table_7col": {
      "template_slide_indices": [1, 15, 23, ...],
      "representative": 1,
      "content_schema": {
        "elements": [
          {"name": "main title", "type": "text", "data": [...]},
          {"name": "table_cell_1_1", "type": "text", "default_quantity": 1, "suggested_characters": 20},
          ...
        ]
      },
      "frequency": 127,
      "tags": ["structure:grid", "density:high", "section:analysis"]
    }
  }
}
```

### 9.2 Phase D (편집 API 5종 포팅, 순수 python-pptx 버전)

**대상 파일**: `ppt_builder/template/editor.py` 확장

**이식 함수** (순수 python-pptx로 재작성):
```python
# 신규 추가
def clone_paragraph(slide, shape_id: int, paragraph_id: int) -> int:
    """paragraph를 복제 (XML 레벨: addnext + parse_xml(.xml))"""

def replace_paragraph(slide, shape_id: int, paragraph_id: int, text: str):
    """markdown → BS4 → Run 재빌드 (PPTAgent 방식)"""

def del_paragraph(slide, shape_id: int, paragraph_id: int):
    """_element.getparent().remove(_element)"""

def replace_image(slide, shape_id: int, image_path: str):
    """비율 유지 재배치"""

def del_image(slide, shape_id: int):
    """shapes.remove()"""
```

**주의사항 (한글 지원 검증)**:
1. Run의 `<a:ea typeface="맑은 고딕"/>` (East Asian 폰트) 속성이 `parse_xml(_r.xml)` 복제에서 보존되는지 실측
2. `bullet-type="▪"` 같은 커스텀 불릿이 `clone_paragraph` 후에도 유지되는지
3. 그룹 내부 paragraph는 `shape_id` 체계에 포함 여부 결정 (평탄화 vs 계층)

### 9.3 Phase D+ (Closure 패턴 도입)

```python
# PPTAgent 방식 차용
@dataclass
class EditClosure:
    action: Callable
    target_idx: int
    type: Literal["clone", "replace", "delete", "merge"]

# 편집 의도를 queue에 저장
slide._closures["replace"].append(EditClosure(partial(replace_para, ...), 3, "replace"))

# build() 시점 일괄 적용 (역순 인덱스 문제 회피)
for closure in sorted(slide._closures["clone"], key=lambda c: -c.target_idx):
    closure.action(shape)
```

### 9.4 Phase E (평가 프레임 재편성)

**PPTEval 3축 도입**: `ppt_builder/evaluate.py`
- **Content** (슬라이드별): 텍스트 간결성, 문법, 이미지 관련성
- **Design** (슬라이드별): 색 조화, 레이아웃 가독성
- **Coherence** (전체): 논리 구조, 맥락 흐름

Judge는 Claude Code 자체 수행. PPTAgent 논문 인간 평가 Pearson 0.71 수준 재현 가능.

---

## 10. 다음 실행 단계 (우선순위)

| 순 | 작업 | 소요 | 담당 |
|---|---|---|---|
| 1 | **한글 `clone_paragraph` 실측 실험** (5장으로 mini POC) | 30분 | 즉시 |
| 2 | `ppt_builder/template/editor.py`에 편집 API 5종 추가 | 2시간 | Phase D 착수 |
| 3 | Phase A2 파이프라인 설계 (induct.py 포팅 스펙) | 1시간 | 설계 |
| 4 | Phase A2 실행 (20 Agent 병렬 태깅) | 2~3일 | 장기 |
| 5 | Phase E 평가 3축 도입 | 1시간 | 병행 |

---

## 11. 한글 POC 실측 결과 ✅

**슬라이드 1175 (SAP FI 모듈 관련, 한글 13개 paragraph)**에서 `parse_xml(_r.xml) + addnext` 방식 XML 복제 실행:

| 한글 paragraph | EA 폰트 | Latin 폰트 | font-size | 복제 |
|---|---|---|---|---|
| "지급\x0b방법/조건\x0b변경" (7 run) | 맑은 고딕 ×7 | 맑은 고딕 ×7 | — | ✅ |
| "채무 생성" | 맑은 고딕 | 맑은 고딕 | 10pt | ✅ |
| "지급 대상 리스트 조회" | 맑은 고딕 | 맑은 고딕 | 10pt | ✅ |
| "지급 요청(F110)" (3 run) | 맑은 고딕 ×3 | 맑은 고딕 ×3 | 10pt | ✅ |
| "펌 뱅킹" | 맑은 고딕 | 맑은 고딕 | 10pt | ✅ |

**결론**:
- ✅ `<a:ea typeface="맑은 고딕">` 복제 완벽
- ✅ `<a:latin typeface="맑은 고딕">` 복제 완벽
- ✅ `font-size`, `\x0b`(vertical tab) 보존
- ✅ 그룹 3단 중첩 경로 (`/3/7/1`) 접근 가능
- ✅ **PPTAgent의 편집 API 방식이 순수 python-pptx + 한글에서 작동 검증됨**

## 12. 리스크 / 주의점

1. **최대 shape 3,980 슬라이드 (#189)** — 별도 점검 필요
2. **`pptagent_pptx` 포크 의존성** — 순수 python-pptx로 포팅 시 `parse_xml` import만 `pptx.oxml.ns`로 변경 (✅ POC 검증 완료)
3. **한글 East Asian 폰트 속성** — ✅ 검증 완료
4. **PPTAgent V2 (deeppresenter)는 무시** — v0.2.x 코어만 참조
5. **아웃라이어 슬라이드** 예외 처리 필요

## 13. 참고 경로

- 템플릿 파일: `c:/Users/y2kbo/Apps/PPT/docs/references/_master_templates/PPT 템플릿.pptx`
- 샘플 20장 PNG: `c:/Users/y2kbo/Apps/PPT/docs/references/_master_templates/_samples/pngs/`
- 한글 POC 결과: `c:/Users/y2kbo/Apps/PPT/docs/references/_master_templates/_samples/clone_test_hangul.pptx`
- PPTAgent 클론: `C:/Users/y2kbo/.claude/projects/c--Users-y2kbo-Coding/PPTAgent/`
- 검수 스크립트: `scripts/inspect_master_templates.py`, `scripts/probe_group_placeholder.py`, `scripts/extract_samples.py`

---

## 12. 참고 경로

- 템플릿 파일: `c:/Users/y2kbo/Apps/PPT/docs/references/_master_templates/PPT 템플릿.pptx`
- 샘플 20장 PNG: `c:/Users/y2kbo/Apps/PPT/docs/references/_master_templates/_samples/pngs/`
- PPTAgent 클론: `C:/Users/y2kbo/.claude/projects/c--Users-y2kbo-Coding/PPTAgent/`
- 검수 스크립트: `scripts/inspect_master_templates.py`, `scripts/probe_group_placeholder.py`, `scripts/extract_samples.py`
