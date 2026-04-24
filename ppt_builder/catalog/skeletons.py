"""Track 2 Stage B — 이론 기반 덱 스켈레톤 (Option C, 2026-04-24 확정).

배경
----
Stage B1 (데이터 기반 경계 감지) 실패: 마스터 템플릿 1,251장이 **연속 카탈로그**
(페이지번호 1~942 연속, 합본 아님)임이 드러남. 데이터 기반 LCS 추출 불가.

대안: SCQA/Minto Pyramid 기반 7개 표준 컨설팅 서사를 직접 encoding.
이 방식은 VLM-SlideEval 2025 권고 — "rule-based coherence gate" — 와 정합.

사용
----
PPT 생성 워크플로우 Step 1 (Outline)에서:
  1. 시나리오 분류 → skeleton 선택
  2. `DeckOutline.narrative_sequence` 생성 (역할 시퀀스)
  3. Step 2-3에서 각 역할 슬라이드를 카탈로그에서 retrieval
  4. Step 5에서 evaluate_deck_coherence로 검증

참조
----
- McKinsey Pyramid Principle (Barbara Minto, 1987)
- SCQA framework (Situation-Complication-Question-Answer)
- ArcDeck 2025: Outline-First Planning (RST + global commitment)
- memory/reference_deck_level_research.md
"""
from __future__ import annotations

from .schemas import DeckSkeleton, NarrativeRole

R = NarrativeRole  # alias


# ---------------------------------------------------------------------------
# 7개 표준 스켈레톤 (컨설팅 덱 유형별)
# ---------------------------------------------------------------------------

_PROPOSAL_30 = DeckSkeleton(
    skeleton_id="consulting_proposal_30",
    use_cases=["제안서", "프로젝트 RFP 응답", "사업 제안"],
    narrative_sequence=[
        R.OPENING,                           # 1: 표지
        R.AGENDA,                            # 2: 목차
        R.SITUATION, R.SITUATION,            # 3-4: 고객 현황, 비즈니스 배경
        R.COMPLICATION, R.COMPLICATION,      # 5-6: 문제 진단, 갭 분석
        R.EVIDENCE, R.EVIDENCE,              # 7-8: 근거 데이터
        R.ANALYSIS, R.ANALYSIS, R.ANALYSIS,  # 9-11: 원인 분석
        R.DIVIDER,                           # 12: 해결 방안 구분
        R.RECOMMENDATION, R.RECOMMENDATION,  # 13-14: 접근 전략 개요
        R.RECOMMENDATION, R.RECOMMENDATION,  # 15-16: 세부 방안 1-2
        R.RECOMMENDATION, R.RECOMMENDATION,  # 17-18: 세부 방안 3-4
        R.ROADMAP, R.ROADMAP, R.ROADMAP,     # 19-21: 단계별 실행 계획
        R.BENEFIT, R.BENEFIT,                # 22-23: 기대 효과 (정량/정성)
        R.RISK, R.RISK,                      # 24-25: 리스크 + 완화책
        R.EVIDENCE,                          # 26: 수행 실적 (레퍼런스)
        R.EVIDENCE,                          # 27: 팀 역량
        R.CLOSING,                           # 28: 결론
        R.CLOSING,                           # 29: Q&A / 감사
        R.APPENDIX,                          # 30: 부록
    ],
    slide_count_range=(25, 40),
    frequency=0,  # 이론 기반 — 추출 아님
    example_deck_ids=[],
)

_ANALYSIS_15 = DeckSkeleton(
    skeleton_id="analysis_report_15",
    use_cases=["시장 분석", "재무 리뷰", "운영 진단", "Q1/Q2 리뷰"],
    narrative_sequence=[
        R.OPENING,                           # 1
        R.AGENDA,                            # 2
        R.SITUATION, R.SITUATION,            # 3-4: 분석 대상 정의, 기간
        R.EVIDENCE, R.EVIDENCE,              # 5-6: 핵심 데이터 제시
        R.ANALYSIS, R.ANALYSIS, R.ANALYSIS,  # 7-9: Finding 1-3
        R.ANALYSIS, R.ANALYSIS,              # 10-11: Finding 4-5
        R.RECOMMENDATION, R.RECOMMENDATION,  # 12-13: 시사점 + Next Step
        R.RISK,                              # 14: 불확실 요소
        R.CLOSING,                           # 15: 요약
    ],
    slide_count_range=(12, 20),
    frequency=0,
    example_deck_ids=[],
)

_ROADMAP_10 = DeckSkeleton(
    skeleton_id="transformation_roadmap_10",
    use_cases=["전환 로드맵", "넷제로 실증", "디지털 전환 계획"],
    narrative_sequence=[
        R.OPENING,                           # 1: 표지 + governing thought
        R.SITUATION,                         # 2: 현재 상태
        R.COMPLICATION,                      # 3: 전환 이유 (규제/경쟁/내부 driver)
        R.ANALYSIS,                          # 4: 현재 gap / 목표와의 차이
        R.RECOMMENDATION,                    # 5: 전략 pillars
        R.ROADMAP, R.ROADMAP,                # 6-7: Phase 1 / Phase 2
        R.BENEFIT,                           # 8: 도달 시 효과
        R.RISK,                              # 9: 전환 리스크
        R.CLOSING,                           # 10: Next Step + Call to action
    ],
    slide_count_range=(8, 14),
    frequency=0,
    example_deck_ids=[],
)

_STRATEGY_40 = DeckSkeleton(
    skeleton_id="executive_strategy_40",
    use_cases=["연간 전략 리뷰", "이사회 자료", "중장기 비전"],
    narrative_sequence=[
        R.OPENING,                              # 1
        R.AGENDA,                               # 2
        R.SITUATION, R.SITUATION, R.SITUATION,  # 3-5: 시장/경쟁/내부
        R.COMPLICATION, R.COMPLICATION,         # 6-7: 구조적 이슈
        R.EVIDENCE, R.EVIDENCE, R.EVIDENCE,     # 8-10: 벤치마크, 트렌드
        R.ANALYSIS, R.ANALYSIS,                 # 11-12: 전략 옵션 비교
        R.DIVIDER,                              # 13: 제안 전략
        R.RECOMMENDATION, R.RECOMMENDATION,     # 14-15: 비전 + 미션
        R.RECOMMENDATION, R.RECOMMENDATION,     # 16-17: Pillar 1
        R.RECOMMENDATION, R.RECOMMENDATION,     # 18-19: Pillar 2
        R.RECOMMENDATION, R.RECOMMENDATION,     # 20-21: Pillar 3
        R.RECOMMENDATION,                       # 22: Pillar 4 (옵션)
        R.ROADMAP, R.ROADMAP, R.ROADMAP,        # 23-25: 3-year / 1-year / Quick win
        R.ROADMAP,                              # 26: 거버넌스 / 조직
        R.BENEFIT, R.BENEFIT,                   # 27-28: 정량/정성 효과
        R.EVIDENCE,                             # 29: 투자 계획
        R.RISK, R.RISK,                         # 30-31: 전략적/운영적 리스크
        R.ANALYSIS,                             # 32: 민감도 분석
        R.EVIDENCE, R.EVIDENCE,                 # 33-34: KPI 대시보드
        R.RECOMMENDATION,                       # 35: 의사결정 요청
        R.CLOSING, R.CLOSING,                   # 36-37: 요약 + Q&A
        R.APPENDIX, R.APPENDIX, R.APPENDIX,     # 38-40: 부록
    ],
    slide_count_range=(35, 50),
    frequency=0,
    example_deck_ids=[],
)

_CHANGE_MGMT_20 = DeckSkeleton(
    skeleton_id="change_management_20",
    use_cases=["조직 개편", "M&A 통합", "Change Management 실행"],
    narrative_sequence=[
        R.OPENING,                              # 1
        R.AGENDA,                               # 2
        R.SITUATION, R.SITUATION,               # 3-4: 변화 배경
        R.COMPLICATION,                         # 5: 변화 필요성
        R.ANALYSIS, R.ANALYSIS,                 # 6-7: 이해관계자 / 영향도
        R.RECOMMENDATION, R.RECOMMENDATION,     # 8-9: To-Be 조직/프로세스
        R.RECOMMENDATION, R.RECOMMENDATION,     # 10-11: Change 프레임워크
        R.ROADMAP, R.ROADMAP, R.ROADMAP,        # 12-14: 3단계 실행
        R.EVIDENCE,                             # 15: 커뮤니케이션 전략
        R.EVIDENCE,                             # 16: 교육 / 온보딩
        R.BENEFIT,                              # 17: 기대 효과
        R.RISK, R.RISK,                         # 18-19: 저항 관리 + Contingency
        R.CLOSING,                              # 20
    ],
    slide_count_range=(16, 25),
    frequency=0,
    example_deck_ids=[],
)

_UPDATE_10 = DeckSkeleton(
    skeleton_id="progress_update_10",
    use_cases=["프로젝트 진척 보고", "월간/분기 업데이트", "스티어링 위원회"],
    narrative_sequence=[
        R.OPENING,
        R.EVIDENCE,                             # 2: 이번 기간 하이라이트
        R.EVIDENCE, R.EVIDENCE,                 # 3-4: KPI 현황
        R.ANALYSIS,                             # 5: 진척 분석 (Plan vs Actual)
        R.COMPLICATION,                         # 6: 이슈 / 블로커
        R.RECOMMENDATION,                       # 7: 대응 방안
        R.ROADMAP,                              # 8: 다음 기간 계획
        R.RISK,                                 # 9: 예상 리스크
        R.CLOSING,                              # 10: 의사결정 요청
    ],
    slide_count_range=(6, 14),
    frequency=0,
    example_deck_ids=[],
)

_PITCH_8 = DeckSkeleton(
    skeleton_id="short_pitch_8",
    use_cases=["임원 요약", "투자 설명", "8분 pitch"],
    narrative_sequence=[
        R.OPENING,              # 1
        R.SITUATION,            # 2: 시장/고객
        R.COMPLICATION,         # 3: Pain point
        R.RECOMMENDATION,       # 4: 솔루션
        R.EVIDENCE,             # 5: 차별화 / 증거
        R.BENEFIT,              # 6: Traction / 효과
        R.ROADMAP,              # 7: 실행 계획
        R.CLOSING,              # 8: Ask / Call to action
    ],
    slide_count_range=(5, 10),
    frequency=0,
    example_deck_ids=[],
)


SKELETONS: dict[str, DeckSkeleton] = {
    s.skeleton_id: s for s in [
        _PROPOSAL_30, _ANALYSIS_15, _ROADMAP_10,
        _STRATEGY_40, _CHANGE_MGMT_20, _UPDATE_10, _PITCH_8,
    ]
}


def get_skeleton(skeleton_id: str) -> DeckSkeleton:
    """스켈레톤 ID로 조회."""
    if skeleton_id not in SKELETONS:
        raise KeyError(
            f"Unknown skeleton: {skeleton_id}. "
            f"Available: {list(SKELETONS.keys())}"
        )
    return SKELETONS[skeleton_id]


def recommend_skeleton(use_case: str, target_slides: int | None = None) -> DeckSkeleton:
    """유스케이스 키워드 + 목표 슬라이드 수로 가장 가까운 스켈레톤 추천.

    단순 키워드 매칭 + 슬라이드 수 근접도. 2025-04-24 기준 rule-based.
    """
    uc = (use_case or "").lower()
    candidates = []
    for sk in SKELETONS.values():
        hits = sum(1 for kw in sk.use_cases if any(t in uc for t in kw.lower().split()))
        if hits > 0:
            candidates.append((sk, hits))

    if not candidates:
        # fallback — 슬라이드 수로만
        if target_slides is None:
            return _PROPOSAL_30
        best = min(
            SKELETONS.values(),
            key=lambda s: abs(((s.slide_count_range[0] + s.slide_count_range[1]) / 2) - target_slides),
        )
        return best

    # 키워드 hit 많은 것 우선, 동률이면 슬라이드 수 근접
    candidates.sort(
        key=lambda x: (
            -x[1],
            abs(((x[0].slide_count_range[0] + x[0].slide_count_range[1]) / 2) - (target_slides or 20)),
        ),
    )
    return candidates[0][0]


def export_skeletons_json(output_path) -> None:
    """`skeletons.json` 으로 저장 — 외부 툴/테스트 용."""
    import json
    from pathlib import Path
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(
            {sid: s.model_dump() for sid, s in SKELETONS.items()},
            f, ensure_ascii=False, indent=2, default=str,
        )


if __name__ == "__main__":
    from pathlib import Path
    ROOT = Path(__file__).resolve().parent.parent.parent
    out = ROOT / "output" / "catalog" / "skeletons.json"
    export_skeletons_json(out)
    print(f"[OK] {len(SKELETONS)} skeletons -> {out}")
    for sid, sk in SKELETONS.items():
        seq_summary = ", ".join(r.value for r in sk.narrative_sequence[:5]) + "..."
        print(f"  {sid}: {len(sk.narrative_sequence)}장 [{seq_summary}]")
