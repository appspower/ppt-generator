# PROJECT DIRECTION — 확정 헌법 (2026-04-25)

> **본 문서는 프로젝트 방향의 단일 진실 소스(SSoT).**
> 사용자와 Claude Code 모두 이 방향을 흔들지 않음.
> 변경은 §6 변경 조건 충족 + 사용자 명시적 결정 시에만.

---

## 1. 확정 방향: **Mode A + N1-Lite 하이브리드**

### Mode A (Whole-slide reuse) — 백본
- 1,251장 마스터 PPT의 슬라이드 1장을 **통째로 복제**
- 슬라이드의 placeholder (`~~`) 텍스트만 교체
- 패턴 템플릿을 그대로 사용하는 방식
- 외부 SOTA (PPTAgent EMNLP 2025)와 일치

### N1-Lite (Single-component reuse) — 보조
- 1,251장에서 도형 그룹(GroupShape) 단위 컴포넌트 추출
- **빈 슬라이드에 컴포넌트 1개만 배치 + placeholder 채움**
- 자동 조합(여러 컴포넌트 + 충돌 해결)은 **하지 않음** (N1-Full = 미해결 hard problem)
- 컴포넌트는 **같은 마스터에서 추출** → 스타일 자동 일치 보장

### 하이브리드 — 선택 로직
- 시나리오의 narrative_role에 따라 Mode A 또는 N1-Lite 선택
- 사용자가 명시적 override 가능

---

## 2. 거부된 옵션 (왜 안 하는지 명시)

| 옵션 | 거부 이유 |
|---|---|
| **N1-Full** (자동 컴포넌트 조합 + 충돌 해결) | 외부 SOTA(AutoPresent 2.1% 실행 성공)도 미해결. style 일치 + bbox 충돌 hard problem. 시간 4~10일, 결과 보장 X |
| **P3 (회사PC PwC 자산 추가 캡처)** | 1,251장이 이미 placeholder. 추가 캡처해도 같은 작업 반복. 자산 가치 중복 |
| **Mode B 단독 (코드 생성)** | AutoPresent 단독 2.1% 성공률. 컨설팅 그레이드 부적합. fallback에만 의미 있음 |
| **Whole-slide만** (N1-Lite 빼고) | situation/complication/roadmap/benefit/risk 풀 결손 → 시각 다양성 부족 |
| **컴포넌트 자동 조합** | (= N1-Full) 위와 같음 |

---

## 3. role별 모드 할당 (확정)

리서치 기반 매핑 (composition_strategy_research.md §4):

| narrative_role | 모드 | 이유 |
|---|---|---|
| opening | Mode A | 풀 부족하지만 cover 디자인은 슬라이드 통째 |
| agenda | Mode A | 목차는 풀 슬라이드 |
| divider | Mode A | 섹션 구분은 슬라이드 통째 |
| **situation** | **N1-Lite** | 풀 1장 결손 — 컴포넌트 조합으로 보강 |
| **complication** | **N1-Lite** | 풀 9장 결손 — 컴포넌트 |
| evidence | Mode A | 풀 충분 (452) |
| analysis | Mode A | 풀 충분 (1118) |
| recommendation | Mode A | 풀 충분 (630) |
| **roadmap** | **N1-Lite** | timeline/chevron 컴포넌트 단독 사용이 깔끔 |
| **benefit** | **N1-Lite** | 풀 5장 결손 |
| **risk** | **N1-Lite** | 풀 1장 결손 |
| closing | Mode A | 풀 부족하지만 슬라이드 통째 |
| appendix | Mode A | 풀 충분 |

→ **5 role(situation/complication/roadmap/benefit/risk)이 N1-Lite, 나머지 8 role이 Mode A**

---

## 4. 외부 SOTA 정합성 (왜 이 방향이 옳은지)

| 시스템 | 출처 | 결론 |
|---|---|---|
| **PPTAgent** | EMNLP 2025 | whole-slide select + edit API. 컴포넌트 조합 안 함 |
| **AutoPresent** | CVPR 2025 | 코드 생성, 2.1% 실행 성공. color fidelity 10~18 vs human 73.5 |
| **SlideCoder / Talk-to-Your-Slides** | 2025 | edit existing > compose primitives |
| **PresentBench** | 2026 | Visual Design + Layout 가장 큰 cross-model 갭. dedicated rendering 필요, LLM만으론 불가 |
| **McKinsey/BCG/PwC 실무** | - | "library of old cases" — 슬라이드 reuse 우선 |

**모두 같은 결론**: edit existing slides, do not compose from primitives.

우리 방향(Mode A 백본 + N1-Lite 단일 컴포넌트)은 이 SOTA와 **완전 일치**.

---

## 5. 자산 활용 정책 (확정)

### 핵심 자산
- **편집 가능 자산**: `PPT 템플릿.pptx` (1,251장 placeholder PPT) — Mode A + N1-Lite 둘 다 여기서 추출
- **참조 자산**: HJ 1437장 사진 (user PC) + 179장 선별본 — **컴포넌트 우선순위 결정**의 reference, 직접 PPT 자산은 아님

### 추가 자산 캡처 정책
- **회사PC P3 보류**: 1,251장이 이미 placeholder. 추가 캡처 무의미
- **HJ 사진**: vision reference로만 활용. PPT 자산화 안 함
- **새 자산 필요 시**: 사용자 명시적 결정 필요 (현재 안 함)

---

## 6. 변경 조건 (이 헌법을 흔들 수 있는 유일한 근거)

다음 중 하나가 **데이터로 입증**될 때만 방향 재검토:

1. **70+점 도달 후 ceiling 확인**: Mode A + N1-Lite 구현 완료 후 5 시나리오 평균 80+ 미달성 + REFINE으로도 안 풀림 → N1-Full 일부 도입 검토
2. **N1-Lite 구현 후 시각 단절 발생**: 컴포넌트 단독 슬라이드가 어색 → Mode A 비중 ↑
3. **사용자가 새로운 .pptx 자산 추가 가능**: 풀 결손 해결 가능 → role별 모드 재매핑
4. **외부 SOTA에서 N1-Full 돌파구**: 학계 새 논문에서 자동 조합 품질 80%+ 달성 → 도입 검토

위 4가지 외에는 **임의 변경 금지**.

---

## 7. 새 세션 시작 시 필수 로드

새 세션의 첫 작업은 항상 본 문서 + N1_LITE_IMPLEMENTATION.md 확인:

```
1. CLAUDE.md (프로젝트 규칙) → docs/PROJECT_DIRECTION.md (본 문서)
2. docs/N1_LITE_IMPLEMENTATION.md (세부 구현 사양)
3. 메모리: project_phase_a3_v5_vision (현재 점수) +
           project_direction_v1 (본 문서 인덱스)
4. 현재 진행 위치 확인: docs/_progress/CURRENT_STATE.md
```

---

## 8. 1차 골 (이 헌법 기반)

**Mode A + N1-Lite 하이브리드 구현 + 5 시나리오 평균 70+ 도달**

진행 단계 (N1_LITE_IMPLEMENTATION.md 상세):
1. **컴포넌트 추출 인프라** (`component_ops.py`)
2. **컴포넌트 라이브러리 구축** (5 N1-Lite role)
3. **하이브리드 retrieval + assembly** (Mode A vs N1-Lite 선택 로직)
4. **5 시나리오 재측정**

성공 기준: **5 시나리오 평균 70+ 점, 사용자 시각 검수 통과**

---

## 9. 변경 이력

| 일자 | 변경 | 결정자 |
|---|---|---|
| 2026-04-25 | 초안 작성, Mode A + N1-Lite 하이브리드 확정 | 사용자 + Claude |

(이후 변경은 §6 조건 충족 + 사용자 결정 시에만 추가)
