"""Track C - Vision Feedback Loop for ppt_builder.

Track A의 렌더 엔진을 그대로 사용하되, Step 5(EVALUATE)를 비주얼 검증 루프로
강화하는 실험 트랙. Track A 코드는 한 줄도 수정하지 않고 wrapping만 한다.

Workflow:
    1. Track A 렌더러로 .pptx 생성
    2. PowerPoint COM으로 슬라이드별 PNG 추출
    3. evaluate.py 점수 + 시각 리뷰 요청 파일 생성
    4. (Claude Code가 PNG를 직접 보고 피드백 JSON 작성)
    5. JSON 수정 → 재실행 → 비교

Public API:
    run_iteration(schema_path, work_dir, iter_num) -> dict
"""

from .pipeline import run_iteration

__all__ = ["run_iteration"]
