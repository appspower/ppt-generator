# Palantir × SAP S/4HANA — 기술적 심층 방법론 (보강)

> 구체적 기술 단계, SAP 테이블, 도구 비교까지 포함

---

## 1. 테스트 관리 — AIP 기술 상세

### 비즈니스 규칙 추출 (Interpretation AI)
1. Blueprint PDF/Word를 Foundry Pipeline Builder LLM Node에 수집
2. AIP Logic 프롬프트 체인이 문서 청크별로 구조화된 검증 술어 추출
   예: "자재는 MRP Run 전 플랜트 배정이 필수"
3. Ontology 객체 속성(validation rule)으로 타입화하여 Rule Registry에 저장
4. 비기술 사용자가 No-code Rule Engine에서 규칙 검증/보완

### 테스트케이스 형식
- ECATT 스크립트 아님, Excel 아님, Jira 아님
- **Foundry Ontology 객체** (입력/기대출력 속성 쌍)로 저장
- Jira/Excel 연동은 별도 파이프라인 구축 필요
- Tester에게는 Workshop 또는 Quiver 앱으로 표출

### LLM-as-Judge 검증 (단계별)
1. 대상 AIP Logic 함수가 테스트 입력 → 출력 생성
2. 별도 "Judge" AIP Logic 함수가 (입력, 실제출력, 채점기준) 수신
3. 채점기준 예: "사용자 스토리의 모든 인수기준을 커버하는가?"
4. Judge LLM이 0~10 점수 부여, 최소 통과 임계값(예: 9) 설정
5. Rubric Grader가 Pass/Fail + 점수 반환, 케이스별 추론 토큰 디버그
6. 전체 Pass % 집계 → 테스트 품질 대시보드

### SAP 연결 — Foundry Connector 2.0
- RFC, SLT (CDC), BEx, Function Module, T-Code 리포트, HANA View
- 주요 테이블: CDHDR, BKPF, MKPF, AUFK, VBAK, E070(Transport)
- Add-On: SP30+ (2024.02~), SP32 (2024.10~)

### vs SAP 테스트 도구
| | ECATT | TDMS | Tricentis | Palantir AIP |
|---|---|---|---|---|
| 역할 | 스크립트 실행 | 테스트 데이터 | UI 자동화+커버리지 | **시나리오 생성+검증** |
| SAP 연동 | 네이티브 | 네이티브 | 인증 | Connector Add-On |
| AI 기능 | 없음 | 없음 | Risk-based | LLM 규칙추출+Judge |
| 포지셔닝 | 실행 | 데이터 | 실행+분석 | **생성+우선순위화** |

→ **Palantir = 테스트 생성/우선순위, Tricentis = 실행 레이어** 보완 관계

---

## 2. Cutover 오케스트레이션 — Ontology 모델

### CutoverTask 객체 속성
```
task_id, owner, dependency_ids[], 
planned_start, planned_end, actual_start, 
status (NOT_STARTED/IN_PROGRESS/COMPLETE/BLOCKED),
system, category
```

### 카테고리 예시 (실제 SAP Cutover)
- MRP_RUN: 최종 MRP 실행
- BALANCE_MIGRATION: 잔액 이관 (FI 마이그레이션)
- MASTER_DATA_FREEZE: 마스터 데이터 동결
- OPEN_ITEM_CARRY_FORWARD: 미결항목 이월
- IDOC_QUEUE_CLEAR: IDoc 큐 정리
- BATCH_JOB_DISABLE: 배치 잡 비활성화
- CUTOFF_CONFIRMATION: 마감 확인

### 실시간 Cutover 대시보드 (Workshop)
- 워크스트림별 Task 상태 스윔레인 (Basis/FI/MM/SD/Master Data)
- 의존성 그래프 (노드 색상=상태)
- 미해결 Blocker 피드 (Task 객체 링크)
- 계획 타임라인 대비 진행률 (% 완료 vs 경과 시간)
- 데이터 로드 진행 바 (타겟 적재 행 수 vs 예상)
- AIP Assist: "FI 태스크 중 blocked인 것은?" 즉시 응답

### SAP CALM 연동
- 공식 네이티브 커넥터 없음 (2026.04 기준)
- REST connector로 CALM OData API 호출하여 구축 가능
- SAP-Palantir 파트너십은 BDC(Business Data Cloud) 중심

---

## 3. 프로세스 마이닝 — SAP 이벤트 로그

### 소스 테이블 매핑
| Foundry 필드 | SAP 소스 | 컬럼 |
|---|---|---|
| object_id | CDHDR.OBJECTID | 변경문서 객체 |
| activity | CDHDR.TCODE + CDPOS.FNAME | 트랜잭션+변경필드 |
| timestamp | CDHDR.UDATE + UTIME | 일시 |
| actor | CDHDR.USERNAME | 사용자 |
| case_id | BKPF.BELNR(FI), VBAK.VBELN(SD), MKPF.MBLNR(MM) | 문서번호 |

- SLT 스트리밍으로 연속 수집
- Z-Transaction: CDHDR.TCODE에 Z* 포함 → **자동 탐지 가능**

### vs Celonis/Signavio
| | Celonis | Palantir Machinery | Signavio |
|---|---|---|---|
| SAP 커넥터 | CEBT 인증 추출기 | RFC/SLT | 모델 기반 (로그 아님) |
| 프로세스 템플릿 | P2P/O2C 사전 구축 | 없음 (커스텀) | BPMN 모델링 |
| 프로세스 발견 | Alpha+DFG | DFG/상태머신 | 없음 |
| 액션 엔진 | Celonis Action → SAP 푸시 | Ontology Action → SAP Writeback | 없음 |
| Z-Transaction | 커스텀 추출기 필요 | **CDHDR 네이티브 캡처** | N/A |

→ **Palantir = 탐지→시뮬→실행 단일 플랫폼** (Celonis는 탐지+분석 중심)

---

## 4. 추가 구체적 방법론

### Sprint Planning AI Assist
- AIP Agent: Jira 백로그 + 과거 Velocity + 미결 결함 수 조회
- 용량 조정된 Sprint Story Point 추천 생성

### Config 문서 자동화
- RFC로 SAP Config 테이블 추출 (T-code SPRO) → Dataset
- AIP Logic로 사람이 읽을 수 있는 Config 문서 자동 생성

### Regression 테스트 우선순위
- Transport Request(E070) 변경 영향도 점수
- 기능 영역별 과거 결함 밀도
- 인터페이스 의존성 수
→ 가중 우선순위 큐 → Workshop 테스트 관리 앱

### Hypercare 모니터링
- SM21(시스템 로그), EDIDS(IDoc 에러), TBTCO(배치 잡) SLT 수집
- 이상 탐지: 에러율 기준선 초과, 배치 런타임 2σ 이탈, 특정 IDoc 누적
- Workshop 대시보드 + AIP Agent 근본원인 질의

---

## 핵심 주의사항
- Palantir는 SAP Activate와 같은 인증 구현 방법론을 제공하지 않음
- 위 패턴은 Foundry/AIP 공식 문서 + SAP Community + 파트너 사례에서 조합
- Cutover/거버넌스 앱은 상당한 커스텀 빌드 필요 (사전 패키지 아님)

출처: Palantir Foundry Docs, AIP Evals, Connector 2.0 for SAP, SAP Community, Unit8, Springer Process Mining
