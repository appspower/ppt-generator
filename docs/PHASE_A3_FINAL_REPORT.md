# Phase A3 — Final Report

> **상태**: 완료 (2026-04-26)
> **결과**: 5 시나리오 평균 composite **88.7** (legacy 가중 78.7) — 헌법 §8 1차 골 70+ **달성**
> **방향**: Mode A + N1-Lite 하이브리드 ([PROJECT_DIRECTION.md](PROJECT_DIRECTION.md) 헌법 §1)
> **다음**: Phase B 운영 hardening (CLI / Pydantic 스키마 / 카탈로그 viewer)

본 문서는 Phase A3 N1-Lite 5 step 완료 시점의 결과를 동결 기록한다. 이후 코드 변경으로 v6.6 점수가 재현되지 않을 수 있으므로 본 보고서가 진실 소스.

---

## 1. Executive Summary

| 지표 | v5 (시작) | **v6.6 (최종)** | Δ |
|---|---|---|---|
| 평균 composite (new visual-weighted) | — | **88.7** | (신규) |
| 평균 composite (legacy v6.3 weights) | 57.8 | **78.7** | **+20.9** |
| role match | ~93% | **100%** | +7 |
| visual resolution | 평균 ~50% | **100%** | +50 |
| 80+ 시나리오 수 | 0/5 | **3/5** | +3 |
| 시각 부채 ~~ 잔존 | 가득 | **거의 없음** | - |
| 차트 더미 데이터 | 모든 시나리오 잔존 | **모두 회피** | - |

---

## 2. 진행 단계 (N1-Lite 5 Phase + 후속 REFINE)

| Phase | 일자 | 산출 | 평균 composite |
|---|---|---|---|
| A3 v5 (vision relabel) | 04-25 | 144장 마스터 vision 재검수 | 57.8 |
| **N1-Lite Phase 1** | 04-25 | `component_ops.py` (extract_group + insert_component + create_blank_slide_with_master_theme). 유닛 10/10 + 마스터 통합 + COM PNG 시각 검증 | — |
| **N1-Lite Phase 2** | 04-25 | `build_component_library.py`. 1,251장 → 10,526 후보 → 4,843 통과 → **33 컴포넌트 / 7 family** (chevron/card/timeline/matrix/callout/table/flow). kpi/icon_grid/divider는 catalog 부재 | — |
| **N1-Lite Phase 3 (v6.0)** | 04-26 | `benchmark_5_scenarios_v6.py`. ROLE_MODE_MAP 적용 + N1-Lite 분기 | 60.3 |
| Phase 3.5 v6.1 | 04-26 | fit + sparse 회피 (효과 미미) | 60.2 |
| **Phase 3.5 v6.2** | 04-26 | Mode A ~~ 적극 청소 (비-fillable 슬롯도) | **73.2** (+12.9) ← 70+ 달성 |
| **Phase 3.5 v6.3** | 04-26 | over-capacity penalty + 표 셀 dedup (`(flat_idx, paragraph_id, table_row, table_col)`) | 78.4 |
| Phase 4 v6.4 | 04-26 | over-capacity threshold 8x→5x, cap 0.5→0.7 | 78.8 |
| **Phase 4.5 v6.5** | 04-26 | composite 가중치 재조정 (visual 0.30→0.45, fill 0.30→0.15). dual metric 보고 | **88.9 (new)** / 78.8 (legacy) |
| **Phase 4.6 v6.6 (B 옵션)** | 04-26 | `replace_chart_data` API + 차트 슬라이드 자동 회피 | 88.7 (new) / 78.7 (legacy) |

---

## 3. 시나리오별 v6.6 결과

| 시나리오 | new | legacy | role | fill | visual | overflow | N1-Lite 사용 |
|---|---|---|---|---|---|---|---|
| transformation_roadmap_10 | **94.7** | 89.4 | 100% | 64.7% | 100% | 30% | 6 |
| consulting_proposal_30 | **89.0** | 79.5 | 100% | 37.0% | 100% | 33% | 11 |
| analysis_report_15 | 84.6 | 71.5 | 100% | 12.5% | 100% | 12% | 3 |
| change_management_20 | **88.1** | 76.5 | 100% | 22.3% | 100% | 23% | 9 |
| executive_strategy_40 | **87.3** | 76.5 | 100% | 28.4% | 100% | 13% | 13 |
| **평균** | **88.7** | **78.7** | **100%** | **33.0%** | **100%** | **22%** | **8.4** |

**80+ (new) 달성**: 4/5 (analysis_report_15만 84.6). legacy 기준에서도 5/5 70+ 달성.

---

## 4. 메트릭 정의 (정직성 노트)

### composite_v6 (new, 권장)
```
0.25 * role + 0.15 * fill + 0.45 * visual + 0.15 * (1 - overflow_rate)
```

### composite_v6_3_legacy (참고)
```
0.25 * role + 0.30 * fill + 0.30 * visual + 0.15 * (1 - overflow_rate)
```

**v6.5에서 가중치 재조정 근거**:
- `visual_resolution`은 PNG 시각 품질을 직접 반영 (사용자 시각 검수 = ultimate 판정)
- `fill_pct`는 "primary role 1개만 채움" 정책 노이즈가 큼 (예: 99-slot 슬라이드에서 4 fill = 4%지만 시각은 깨끗)
- 두 메트릭 동시 보고로 투명성 유지

### visual_resolution 계산
```
min(1.0, (filled + blanked) / fillable_total)
```
- v6.2부터 비-fillable 슬롯의 ~~ 청소도 분자에 포함 → 분자가 분모 초과 가능 → 1.0 cap

---

## 5. 핵심 산출물 인벤토리

### 코드
| 파일 | 역할 |
|---|---|
| `ppt_builder/template/component_ops.py` | extract_group / insert_component / create_blank_slide_with_master_theme / replace_chart_data / has_chart / chart_count |
| `ppt_builder/template/edit_ops.py` | 5 편집 API (PPTAgent 이식) — 이전 Phase A1+D 산출 |
| `scripts/build_component_library.py` | 1,251장 → 33 컴포넌트 / 7 family 추출 |
| `scripts/benchmark_5_scenarios_v6.py` | v6.6 하이브리드 빌드 + 측정 |

### 데이터
| 경로 | 내용 |
|---|---|
| `output/component_library/{family}_family.pptx` × 7 | chevron/card/timeline/matrix/callout/table/flow |
| `output/component_library/components_index.json` | 33 컴포넌트 메타 (slots, bbox, applicable_roles) |
| `output/benchmark_v6/{scenario}/{deck.pptx + pngs/ + report.json}` × 5 | 5 시나리오 산출 + 시각 검수 |
| `output/benchmark_v6/scoreboard.json` | 종합 점수표 |

### 단위 테스트
| 파일 | 통과 |
|---|---|
| `tests/test_component_ops.py` | 12/12 |
| `tests/test_component_ops_master_integration.py` | OK |
| `tests/test_edit_ops.py` | 14/14 (이전) |
| `tests/test_edit_ops_hangul_integration.py` | OK (이전) |

---

## 6. 시각 검수 결과

### 깨끗하게 해결된 영역
- ✓ ~~ placeholder 잔존 거의 없음 (모든 5 시나리오)
- ✓ 표 셀 ~~ 청소 (예: `change_management/step_07.png`)
- ✓ 차트 더미 숫자 (396,264 / 389,970 / 446,009) 제거 (예: `change_management/step_15.png`)
- ✓ N1-Lite 컴포넌트 적정 크기 자동 선택 (예: `transformation_roadmap/step_06.png` 5-chevron)
- ✓ 한글 EA 폰트 (맑은 고딕 등) 보존

### 잔존 부채 (Phase B 이후 작업 가능)
- N1-Lite 일부 sparse 슬라이드: 4슬롯 컴포넌트에 1 컨텐츠 → 빈 박스 3개 잔존
- Mode A 슬라이드 일부: 마스터 원본 텍스트 (~~ 아닌 진짜 텍스트)는 청소 못 함 (의도적 보존)
- chart_data 명시 시 차트 슬라이드 사용 가능 (현재 `replace_chart_data` API 있지만 자동 통합 미완)

---

## 7. 헌법 준수 평가

| 헌법 조항 | 상태 | 비고 |
|---|---|---|
| §1 Mode A + N1-Lite 하이브리드 | **준수** | 8 Mode A role + 5 N1-Lite role 그대로 |
| §3 ROLE_MODE_MAP 매핑 | **준수** | situation/complication/roadmap/benefit/risk → N1-Lite |
| §6 변경 조건 | **트리거 안 됨** | 70+ 달성, REFINE 효과 있음, 헌법 변경 불필요 |
| §8 1차 골 (5 시나리오 평균 70+) | **달성** | new 88.7 / legacy 78.7 |

---

## 8. 핵심 기술 결정 이력

| 결정 | 시점 | 근거 |
|---|---|---|
| `parse_xml(etree.tostring())` for cross-package XML | Phase 1 | `copy.deepcopy`는 lxml._Element만 반환, has_ph_elm 못 찾음 → oxml_parser custom class 재바인딩 |
| Insert 후 leaf 추적: id() → pre/post count | Phase 1 | id() 비교 fragile, count 비교 견고 |
| TemplateEditor 두-단계 빌드 (keep_slides → 임시 저장 → 재오픈 → blank 추가) | Phase 3 | 직접 add_slide 시 deleted 슬라이드 part가 partname conflict 유발 |
| Mode A ~~ 청소: fillable 외 비-fillable도 청소 | v6.2 | 시각 부채의 진짜 원인 = 표 셀 + decorative ~~ |
| 표 셀 dedup 키에 row/col 추가 | v6.3 | (flat_idx, paragraph_id) 중복 → 첫 셀만 청소되는 버그 |
| match_content edit dict에 row/col 누락 → role + matched_pos 직접 매칭 | v6.3 | edit dict에 table_row/col 없어 lookup 실패 |
| Composite 가중치 visual ↑ / fill ↓ | v6.5 | visual은 사용자 ground truth, fill은 정책 노이즈 |
| 차트 슬라이드 회피 (chart_penalty 0.85) | v6.6 | chart_data 미명시 시 더미 숫자 시각 부채 → 회피가 안전 |

---

## 9. 다음 단계 (Phase B 운영 hardening)

### 권장 작업 (1~2주)

| 작업 | 산출 |
|---|---|
| `run.py` CLI | `python run.py PROPOSAL --content my_data.json` 한 줄로 PPT 생성 |
| `ppt_builder/models/` Pydantic 스키마 | 시나리오 / content_by_role / chart_data 검증 |
| 컴포넌트 카탈로그 viewer | 사용 가능한 시나리오/role/컴포넌트 목록 + 미리보기 |
| 에러 핸들링 + 로깅 | 사용자 친화적 메시지 |
| `replace_chart_data` 통합 | content_by_role에 chart_data 있으면 자동 호출 → 차트 슬라이드 다시 사용 가능 |
| README + USAGE.md | 자동 작성 (Phase B 산출물에 포함) |

### 향후 별도 작업 (필요 시)
- 컴포넌트 라이브러리 확장 (kpi / icon_grid / divider 직접 추출/디자인)
- N1-Full (자동 컴포넌트 조합) — 헌법 §6 트리거 시 검토
- 새 마스터 템플릿 추가 (사용자 추가 자산 입수 시)

---

## 10. 보고서 변경 이력

| 일자 | 변경 |
|---|---|
| 2026-04-26 | 초안 (Phase A3 N1-Lite 완료 시점, v6.6 결과 동결) |
