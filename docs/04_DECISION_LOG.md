# 04. Decision Log

> 추가만 가능합니다. 기존 항목을 삭제하거나 수정하지 않습니다.

| 날짜 | 결정 | 이유 | 고려한 대안 |
|---|---|---|---|
| 2026-04-08 | python-pptx를 PPT 엔진으로 선택 | .pptx 네이티브, 템플릿 기반, 가장 성숙한 Python 라이브러리 | PptxGenJS (Node.js), aspose-slides (상용), LibreOffice UNO |
| 2026-04-08 | 계층적 문서 구조 (CLAUDE.md + docs/) 채택 | ai-company-builder 프로젝트에서 검증된 방법론, Claude Code와의 호환성 | 단일 README, 위키 기반 |
| 2026-04-08 | 슬라이드 스키마를 JSON으로 정의 | LLM 출력 파싱 용이, Pydantic 검증 가능, 렌더러와 분리 | YAML, XML, 직접 python-pptx 호출 |
| 2026-04-08 | ~~MVP는 Streamlit, 프로덕션은 Next.js~~ → 아래 항목으로 대체됨 | ~~빠른 프로토타이핑 → 점진적 확장 전략~~ | ~~Next.js만, Gradio, Flask~~ |
| 2026-04-08 | 복잡 차트는 matplotlib PNG 임베드 | python-pptx가 워터폴/메코 미지원, matplotlib로 자유도 확보 | plotly only, 직접 shape 그리기 |
| 2026-04-08 | Claude Code를 오케스트레이터로 채택, 별도 LLM API/웹앱 레이어 제거 | Claude Code가 리서치/구조화/검증을 이미 내장, 대화형 반복 수정이 컨설팅 품질에 유리, API 비용 절감 | 별도 Claude API + FastAPI + Streamlit 웹앱 (Approach B), 하이브리드 (Approach C) |
| 2026-04-08 | 렌더링 코드를 독립 라이브러리(ppt_builder/)로 설계 | Claude Code에 종속되지 않으면 나중에 웹앱 전환 시 재사용 가능, 테스트 용이 | src/ 하위에 모든 코드 혼합 |
| 2026-04-08 | 슬라이드 타입 → 컴포넌트+레이아웃 엔진 아키텍처로 전환 | 레퍼런스(pwc-ppt) 분석 결과 컴포넌트 조합이 고정 슬라이드 타입보다 유연, Edge-to-Edge 배치와 calc_columns 패턴 차용 | 기존 슬라이드 타입별 고정 렌더러 방식 유지 |
| 2026-04-08 | 컬러 정책: 모노크롬(White-Grey-Black) + 강조색 1개 | 레퍼런스 SKILL.md의 컬러 정책 차용, 유채색 남용 방지, 전문적 톤 유지 | 다색 팔레트 (Blue/Orange/Green 등) |
| 2026-04-08 | Assertion Title 규칙 채택 (핵심어 중심, 문장형 금지) | 컨설팅 업계 표준, 레퍼런스 SKILL.md 규칙 반영 | 일반 문장형 제목 |
| 2026-04-08 | 강조색 #FD5108 (253,81,8)로 변경, Orange 3단계 + Grey 3단계 팔레트 | 회사 실제 PPT 컬러 정책 사진에서 추출한 정확한 RGB 값 반영 | #D04A02 (레퍼런스 pwc-ppt 기본값) |
| 2026-04-08 | MARGIN 0.4"로 축소, CONTENT_Y 1.3"으로 상향 | spec.md 분석 결과 레퍼런스가 0.4" 사용, 콘텐츠 영역 최대화 | 기존 0.6" 마진 |
| 2026-04-08 | 본문 슬라이드에 오렌지 헤더 바 (h=0.9") 추가, 제목 흰색 | spec.md + Template.pptx 분석, 회사 표준 슬라이드 형식 | 헤더 바 없이 텍스트만 |
| 2026-04-08 | 표지에 오렌지 평행사변형 장식, 종료에 피치 그라데이션 배경 | Template.pptx + Cover and End 에셋 분석, 회사 브랜드 아이덴티티 재현 | 단순 텍스트 표지 |
| 2026-04-08 | Badge를 pill 형태(roundRect adj=0.5)로, Kicker에 SubMarker(오렌지 수직 바) 추가 | components.md 분석: HandBadge=roundRect adj=30000, SubMarker=accent1 수직 바 | 사각형 뱃지, 텍스트만 키커 |
| 2026-04-08 | 프로세스 플로우에 Orange→Grey 그라데이션 단계별 색상 | 템플릿 Process 슬라이드 분석: 번호 박스가 Orange→Medium→Light→Grey 순으로 변화 | 단색 프로세스 |
| 2026-04-08 | 템플릿 인젝션 시스템 도입 (SlideCloner + TextSubstitutor + template_library) | 완성본 94장 분석 결과 40% 이상이 코드로 직접 그리기 어려운 복잡 다이어그램, 사전 제작 슬라이드 복제+치환이 현실적 | 모든 슬라이드를 코드로 렌더링 |
| 2026-04-08 | template_library.pptx 10종 구축 (hub_spoke, timeline, comparison, swimlane, kpi_dashboard, pyramid, sidebar_nav, before_after, swot, value_chain) | 완성본에서 가장 빈번한 복잡 패턴 10가지 선정, JSON으로 데이터만 주입하면 복잡 슬라이드 생성 가능 | 필요시마다 개별 렌더러 구현 |
| 2026-04-24 | 3층 하이브리드 방향 확정 (Layer1 복제+치환 / Layer2 편집API / Layer3 코드fallback) | PwC 1200+장 placeholder 마스터템플릿 확보, 코드 재현 방식의 넷제로 56점 한계 극복. PPTAgent(EMNLP 2025)와 동일 원칙 | 코드 컴포넌트 주력 유지, 완전 LLM 생성 방식 |
| 2026-04-24 | PPTAgent 편집 API 5종 이식 (clone/replace/del paragraph + replace/del image) to `ppt_builder/template/edit_ops.py` | 복제된 슬라이드에 요소 추가/제거/교체 필요 케이스(15% 추정) 커버. PPTAgent MIT 라이선스 + 한글 POC 검증 완료 | 자체 설계, Talk-to-Your-Slides의 `exec()` 방식 |
| 2026-04-24 | `div_id` / `paragraph_id` 매핑을 **평탄화(flat)** 로 결정 (`iter_leaf_shapes` 제너레이터) | PPTAgent `schema_extractor.yaml` + `induct.py` 파이프라인과 호환. Phase A2에서 프롬프트 그대로 재활용 가능 | 계층형 tuple 경로, `shape.shape_id` 기반 |
| 2026-04-24 | `exec()` 기반 코드 실행 방식 불채택 | Talk-to-Your-Slides의 LLM→Python 코드 생성 방식은 엔터프라이즈 보안 부적합. PPTAgent sandbox도 같은 우려 — 본 이식은 API 호출만 허용 | exec+샌드박스, ast.literal_eval |
| 2026-04-24 | Closure 큐잉 패턴 도입 보류 (Phase D+ 후속) | 편집 API 5종은 즉시 적용 버전으로도 유닛 14/14 + 통합 테스트 통과. 역순 인덱스 문제는 호출자 관리로 충분 | 초기부터 Closure 패턴 적용 |
| 2026-04-26 | Phase B v6.7/v6.8 REFINE: density_penalty 강화 (cap 0.7→0.95, slope 0.07→0.10, 가중치 0.15→0.30) | analysis_report_15 (84.6) PNG 시각 검수에서 99-slot 거대 빈 표 부채 확인. 평균 88.7→90.2 (+1.5), 80+ 4/5→**5/5**, 회귀 없음 | hard cutoff (continue skip), 시나리오 컨텐츠 expand, 가중치 미변경 |
| 2026-04-26 | 컴포넌트 라이브러리 확장 (kpi/icon_grid/divider) **abort** | paragraph_labels 카탈로그에 `kpi_value` role 5개(3 슬라이드)뿐, `icon_label` 0개. 라벨 인프라 한계로 빌드 자연 스킵. 헌법 §6 trigger 미충족 (5/5 80+ 평균 90.2 + visual 100% 달성). 자연 해결 = §6 trigger #3 (새 마스터 자산) 또는 별도 Phase에서 카탈로그 재라벨링 | B 재라벨링 (1,251장 reproc, ROI 낮음), C 휴리스틱 추출 (라이브러리 품질 저하 위험) |
