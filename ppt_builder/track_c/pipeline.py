"""Track C 오케스트레이션 파이프라인.

한 번의 iteration:
    1. Track A 렌더러로 schema_path → pptx
    2. PNG 익스포트 (PowerPoint COM)
    3. evaluate.py 자동 점수
    4. 리뷰 요청 마크다운 생성
    5. (Claude Code가 PNG를 직접 보고 review_response.json 작성)

여러 iteration을 돌리려면 schema를 수정해서 다시 호출하면 된다.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ppt_builder import render_presentation
from ppt_builder.evaluate import evaluate_pptx
from ppt_builder.models.schema import PresentationSchema

from .png_export import pptx_to_pngs
from .refine import refine_schema_file
from .vision_review import generate_review_request


def run_iteration(
    schema_path: Path,
    work_dir: Path,
    iter_num: int = 1,
    template: Path | None = None,
) -> dict[str, Any]:
    """Track C 한 회차 실행.

    Args:
        schema_path: 입력 슬라이드 스키마 JSON
        work_dir: iteration 작업 디렉토리 (예: output/track_c/main_r1/)
        iter_num: 회차 번호 (1, 2, 3...)
        template: pptx 마스터 템플릿 (None이면 빈 프레젠테이션)

    Returns:
        {
            "iter_num": int,
            "pptx_path": Path,
            "png_paths": list[Path],
            "eval_report": dict,
            "review_request_path": Path,
            "iter_dir": Path,
        }
    """
    schema_path = Path(schema_path).resolve()
    work_dir = Path(work_dir).resolve()

    iter_dir = work_dir / f"iter_{iter_num:03d}"
    iter_dir.mkdir(parents=True, exist_ok=True)
    pngs_dir = iter_dir / "pngs"

    # 1. Schema 로드 + iter_dir에 사본 저장 (체이닝/추적용)
    with open(schema_path, encoding="utf-8") as f:
        data = json.load(f)
    schema = PresentationSchema.model_validate(data)

    saved_schema_path = iter_dir / f"{schema_path.stem}.json"
    saved_schema_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # 2. Track A 엔진으로 렌더
    pptx_path = iter_dir / f"{schema_path.stem}.pptx"
    render_presentation(schema, template=template, output=pptx_path)

    # 3. PNG 익스포트
    png_paths = pptx_to_pngs(pptx_path, pngs_dir)

    # 4. evaluate.py
    eval_report = evaluate_pptx(pptx_path)

    # 5. 리뷰 요청 마크다운
    review_request_path = generate_review_request(
        work_dir=iter_dir,
        schema_path=schema_path,
        pptx_path=pptx_path,
        png_paths=png_paths,
        eval_report=eval_report,
        iter_num=iter_num,
    )

    # 6. iteration 메타데이터 저장
    meta = {
        "iter_num": iter_num,
        "source_schema_path": str(schema_path),
        "saved_schema_path": str(saved_schema_path),
        "pptx_path": str(pptx_path),
        "png_paths": [str(p) for p in png_paths],
        "eval_score": eval_report["score"],
        "eval_pass": eval_report["pass"],
        "eval_issues": eval_report["issues"],
        "review_request_path": str(review_request_path),
    }
    (iter_dir / "iteration.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return {
        "iter_num": iter_num,
        "pptx_path": pptx_path,
        "png_paths": png_paths,
        "eval_report": eval_report,
        "review_request_path": review_request_path,
        "iter_dir": iter_dir,
    }


def apply_and_rerun(
    work_dir: Path,
    source_iter: int,
    template: Path | None = None,
) -> dict[str, Any]:
    """source_iter의 review_response.json을 읽어 patches를 적용하고
    iter_(source_iter+1)로 재렌더한다.

    체이닝 가능: iter_001 → iter_002 → iter_003 ...

    Args:
        work_dir: Track C 작업 디렉토리 (예: output/track_c/main_r1/)
        source_iter: patches를 가져올 회차 번호
        template: pptx 마스터 템플릿

    Returns:
        run_iteration의 결과 dict + "refine_stats" 키
    """
    work_dir = Path(work_dir).resolve()
    source_dir = work_dir / f"iter_{source_iter:03d}"
    target_iter = source_iter + 1
    target_dir = work_dir / f"iter_{target_iter:03d}"
    target_dir.mkdir(parents=True, exist_ok=True)

    # source iter의 메타데이터 로드 → saved_schema_path 추출
    iter_meta_path = source_dir / "iteration.json"
    if not iter_meta_path.exists():
        raise FileNotFoundError(
            f"source iter 메타데이터 없음: {iter_meta_path}\n"
            f"먼저 run_iteration({source_iter})을 실행해야 함."
        )
    iter_meta = json.loads(iter_meta_path.read_text(encoding="utf-8"))
    source_schema_path = Path(iter_meta["saved_schema_path"])

    # review_response.json 경로
    review_response_path = source_dir / "review_response.json"
    if not review_response_path.exists():
        raise FileNotFoundError(
            f"review_response.json 없음: {review_response_path}\n"
            "Claude Code가 PNG 검토 후 patches를 포함한 review_response.json을 작성해야 함."
        )

    # 새 스키마를 target_dir 에 생성
    target_schema_path = target_dir / source_schema_path.name
    refine_stats = refine_schema_file(
        source_schema_path=source_schema_path,
        review_response_path=review_response_path,
        output_schema_path=target_schema_path,
    )

    # target_iter로 run_iteration 실행
    result = run_iteration(
        schema_path=target_schema_path,
        work_dir=work_dir,
        iter_num=target_iter,
        template=template,
    )
    result["refine_stats"] = refine_stats
    return result


def _print_summary(result: dict[str, Any]) -> None:
    import io
    import sys

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    print()
    print("=" * 60)
    print(f"Track C Iteration {result['iter_num']:03d}")
    print("=" * 60)
    print(f"PPTX:           {result['pptx_path']}")
    print(f"PNG count:      {len(result['png_paths'])}")
    print(f"Auto score:     {result['eval_report']['score']}/100")
    print(f"Auto pass:      {result['eval_report']['pass']}")
    print(f"Review request: {result['review_request_path']}")
    print()
    print("다음 단계: Claude Code가 PNG를 검토하고 review_response.json을 작성")


def main() -> None:
    """CLI 진입점.

    Usage:
        # 신규 iteration 실행
        python -m ppt_builder.track_c.pipeline run <schema.json> [work_dir] [iter_num] [template.pptx]

        # 기존 review_response.json의 patches로 자동 재실행
        python -m ppt_builder.track_c.pipeline refine <work_dir> <source_iter> [template.pptx]

        # 호환 모드 (run 생략 가능)
        python -m ppt_builder.track_c.pipeline <schema.json>
    """
    import sys

    if len(sys.argv) < 2:
        print(__doc__)
        print(main.__doc__)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "refine":
        if len(sys.argv) < 4:
            print("Usage: python -m ppt_builder.track_c.pipeline refine <work_dir> <source_iter> [template.pptx]")
            sys.exit(1)
        work_dir = Path(sys.argv[2])
        source_iter = int(sys.argv[3])
        template = Path(sys.argv[4]) if len(sys.argv) > 4 else None
        result = apply_and_rerun(work_dir=work_dir, source_iter=source_iter, template=template)
        _print_summary(result)
        print(f"Refine stats:   {result['refine_stats']}")
        return

    # run 모드 (기본)
    if cmd == "run":
        args = sys.argv[2:]
    else:
        args = sys.argv[1:]

    if not args:
        print("Usage: python -m ppt_builder.track_c.pipeline run <schema.json> [work_dir] [iter_num] [template.pptx]")
        sys.exit(1)

    schema_path = Path(args[0])
    work_dir = (
        Path(args[1])
        if len(args) > 1
        else Path("output/track_c") / schema_path.stem
    )
    iter_num = int(args[2]) if len(args) > 2 else 1
    template = Path(args[3]) if len(args) > 3 else None

    result = run_iteration(
        schema_path=schema_path,
        work_dir=work_dir,
        iter_num=iter_num,
        template=template,
    )
    _print_summary(result)


if __name__ == "__main__":
    main()
