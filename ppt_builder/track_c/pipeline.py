"""Track C 오케스트레이션 파이프라인.

모드 3종:
    1. run — 1회 iteration (렌더→PNG→evaluate→review_request 생성)
    2. refine — review_response.json의 patches로 자동 재실행
    3. auto — 무인 루프: render→PNG→Claude Vision API→patches→re-render (max 3회)
    4. batch — 여러 스키마 일괄 처리

auto 모드가 P0의 핵심. ANTHROPIC_API_KEY 없으면 mock 모드로 동작.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ppt_builder import render_presentation
from ppt_builder.evaluate import evaluate_pptx
from ppt_builder.models.schema import PresentationSchema

from .auto_vision import review_all_slides
from .png_export import pptx_to_pngs
from .refine import apply_patches, refine_schema_file
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


def run_auto_loop(
    schema_path: Path,
    work_dir: Path,
    max_iterations: int = 3,
    pass_threshold: int = 7,
    template: Path | None = None,
    model: str = "claude-sonnet-4-20250514",
    mock: bool = False,
) -> dict[str, Any]:
    """무인 Vision 피드백 루프.

    render → PNG → Claude Vision API → patches → re-render를 자동 반복한다.
    ANTHROPIC_API_KEY가 없으면 mock=True로 자동 전환.

    종료 조건:
        1. 모든 체크리스트 항목 ≥ pass_threshold → PASS, 루프 종료
        2. max_iterations 도달 → 강제 종료 + 현재 상태 반환
        3. 점수 회귀 감지 → 이전 iteration으로 롤백 + 종료

    Args:
        schema_path: 입력 슬라이드 스키마 JSON
        work_dir: 작업 디렉토리
        max_iterations: 최대 반복 횟수 (기본 3)
        pass_threshold: 체크리스트 각 항목 최소 점수 (기본 7)
        template: pptx 마스터 템플릿
        model: Claude 모델 ID
        mock: True이면 API 호출 없이 mock 결과 사용

    Returns:
        {
            "final_iter": int,
            "final_score": int,
            "passed": bool,
            "stop_reason": str,  # "all_pass" | "max_iterations" | "score_regression" | "no_patches"
            "iterations": [각 iter의 run_iteration 결과 + vision_result],
        }
    """
    import os

    schema_path = Path(schema_path).resolve()
    work_dir = Path(work_dir).resolve()

    # API 키 확인 → 없으면 자동 mock 전환
    if not mock and not os.environ.get("ANTHROPIC_API_KEY"):
        print("[auto_loop] ANTHROPIC_API_KEY 미설정 → mock 모드로 전환")
        mock = True

    iterations: list[dict[str, Any]] = []
    current_schema_path = schema_path
    prev_score = 0

    for iter_num in range(1, max_iterations + 1):
        print(f"\n{'='*50}")
        print(f"Auto Loop — Iteration {iter_num}/{max_iterations}")
        print(f"{'='*50}")

        # 1. render + png + evaluate
        result = run_iteration(
            schema_path=current_schema_path,
            work_dir=work_dir,
            iter_num=iter_num,
            template=template,
        )

        # 2. Claude Vision 자동 평가
        print(f"  Vision 평가 중... ({'mock' if mock else model})")
        vision_result = review_all_slides(
            png_paths=result["png_paths"],
            schema_path=current_schema_path,
            model=model,
            mock=mock,
        )
        vision_result["iter_num"] = iter_num
        current_score = vision_result.get("overall_score", 0)

        # vision_result를 review_response.json으로 저장
        response_path = result["iter_dir"] / "review_response.json"
        response_path.write_text(
            json.dumps(vision_result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        result["vision_result"] = vision_result
        iterations.append(result)

        print(f"  evaluate.py: {result['eval_report']['score']}/100")
        print(f"  Vision:      {current_score}/100")
        print(f"  must_fix:    {vision_result.get('slides', [{}])[0].get('must_fix', False) if vision_result.get('slides') else 'N/A'}")

        # 3. 종료 조건 체크

        # 3a. 모든 체크리스트 항목 ≥ threshold?
        all_pass = True
        for slide in vision_result.get("slides", []):
            checklist = slide.get("checklist", {})
            for key, val in checklist.items():
                if val < pass_threshold:
                    all_pass = False
                    break
            if not all_pass:
                break

        if all_pass:
            print(f"  → ALL PASS (모든 항목 ≥ {pass_threshold}점). 루프 종료.")
            return {
                "final_iter": iter_num,
                "final_score": current_score,
                "passed": True,
                "stop_reason": "all_pass",
                "iterations": iterations,
            }

        # 3b. 점수 회귀?
        if iter_num > 1 and current_score < prev_score:
            print(f"  → 점수 회귀 감지 ({prev_score} → {current_score}). 이전 버전 유지, 루프 종료.")
            return {
                "final_iter": iter_num - 1,
                "final_score": prev_score,
                "passed": False,
                "stop_reason": "score_regression",
                "iterations": iterations,
            }

        # 3c. patches가 없으면?
        patches = vision_result.get("patches", [])
        if not patches:
            print(f"  → patches 없음. 수정 불가. 루프 종료.")
            return {
                "final_iter": iter_num,
                "final_score": current_score,
                "passed": False,
                "stop_reason": "no_patches",
                "iterations": iterations,
            }

        # 3d. 마지막 iteration이면?
        if iter_num >= max_iterations:
            print(f"  → max_iterations={max_iterations} 도달. 강제 종료.")
            return {
                "final_iter": iter_num,
                "final_score": current_score,
                "passed": False,
                "stop_reason": "max_iterations",
                "iterations": iterations,
            }

        # 4. patches 적용 → 다음 iteration용 스키마 생성
        print(f"  → {len(patches)}개 patches 적용 중...")
        prev_score = current_score
        saved_schema = result["iter_dir"] / f"{schema_path.stem}.json"
        current_data = json.loads(saved_schema.read_text(encoding="utf-8"))

        try:
            refined = apply_patches(current_data, patches)
            next_schema_path = work_dir / f"iter_{iter_num+1:03d}" / f"{schema_path.stem}.json"
            next_schema_path.parent.mkdir(parents=True, exist_ok=True)
            next_schema_path.write_text(
                json.dumps(refined, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            current_schema_path = next_schema_path
            print(f"  → 다음 스키마: {next_schema_path}")
        except Exception as e:
            print(f"  → patches 적용 실패: {e}. 루프 종료.")
            return {
                "final_iter": iter_num,
                "final_score": current_score,
                "passed": False,
                "stop_reason": f"patch_error: {e}",
                "iterations": iterations,
            }

    # 여기에 올 일은 없지만 안전망
    return {
        "final_iter": max_iterations,
        "final_score": prev_score,
        "passed": False,
        "stop_reason": "unknown",
        "iterations": iterations,
    }


def run_batch(
    schemas: list[Path],
    work_root: Path = Path("output/track_c"),
    iter_num: int = 1,
    template: Path | None = None,
) -> dict[str, Any]:
    """여러 스키마 파일을 한 번에 처리한다.

    각 스키마는 자체 work_dir(work_root/<stem>)에 격리되어 실행된다.
    PowerPoint COM은 한 번의 호출로 한 파일씩 순차 처리한다 (병렬 불가).

    Args:
        schemas: 처리할 schema JSON 경로 리스트
        work_root: 모든 work_dir의 부모 (예: output/track_c/)
        iter_num: 각 스키마의 회차 번호 (보통 1)
        template: pptx 마스터 템플릿

    Returns:
        {
            "results": [run_iteration 결과 dict, ...],
            "summary": [{"name", "slides", "score", "pass", "issue_count"}, ...],
            "report_path": batch 보고서 마크다운 경로,
        }
    """
    work_root = Path(work_root).resolve()
    work_root.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    summary: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []

    for schema_path in schemas:
        schema_path = Path(schema_path).resolve()
        name = schema_path.stem
        try:
            result = run_iteration(
                schema_path=schema_path,
                work_dir=work_root / name,
                iter_num=iter_num,
                template=template,
            )
            results.append(result)
            report = result["eval_report"]
            summary.append({
                "name": name,
                "slides": report["slide_count"],
                "score": report["score"],
                "pass": report["pass"],
                "issue_count": len(report["issues"]),
                "pptx": str(result["pptx_path"]),
                "review_request": str(result["review_request_path"]),
            })
        except Exception as e:
            failed.append({
                "name": name,
                "schema_path": str(schema_path),
                "error": f"{type(e).__name__}: {e}",
            })

    # 배치 보고서 생성
    report_path = _write_batch_report(work_root, summary, failed, iter_num)

    return {
        "results": results,
        "summary": summary,
        "failed": failed,
        "report_path": report_path,
    }


def _write_batch_report(
    work_root: Path,
    summary: list[dict[str, Any]],
    failed: list[dict[str, Any]],
    iter_num: int,
) -> Path:
    from datetime import datetime

    report_dir = work_root / "_batch_reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = report_dir / f"batch_iter{iter_num:03d}_{timestamp}.md"

    lines = [
        f"# Track C Batch Report — iter_{iter_num:03d}",
        "",
        f"**작성**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"**총 스키마**: {len(summary) + len(failed)}",
        f"**성공**: {len(summary)}",
        f"**실패**: {len(failed)}",
        "",
        "## 스키마별 결과 (evaluate.py 자동 점수)",
        "",
        "| 스키마 | 슬라이드 | evaluate 점수 | PASS | 이슈 수 |",
        "|---|---|---|---|---|",
    ]
    for item in summary:
        status = "PASS" if item["pass"] else "**FAIL**"
        lines.append(
            f"| {item['name']} | {item['slides']} | {item['score']}/100 | {status} | {item['issue_count']} |"
        )

    if summary:
        scores = [item["score"] for item in summary]
        avg = sum(scores) / len(scores)
        lines.extend([
            "",
            f"**evaluate.py 평균**: {avg:.1f}/100",
            "",
        ])

    if failed:
        lines.extend([
            "## 실패한 스키마",
            "",
        ])
        for item in failed:
            lines.append(f"- **{item['name']}** (`{item['schema_path']}`): {item['error']}")
        lines.append("")

    lines.extend([
        "## 다음 단계 — Vision 리뷰",
        "",
        "각 스키마의 review_request.md를 열어 PNG를 확인하고 review_response.json을 작성하세요:",
        "",
    ])
    for item in summary:
        lines.append(f"- {item['name']}: `{item['review_request']}`")

    lines.append("")
    lines.append("Vision 리뷰 후 patches를 작성했다면 자동 재실행:")
    lines.append("```bash")
    for item in summary:
        work_dir = work_root / item["name"]
        lines.append(f"python -m ppt_builder.track_c.pipeline refine {work_dir} {iter_num}")
    lines.append("```")
    lines.append("")

    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


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

        # 여러 스키마 batch 처리 (glob 패턴)
        python -m ppt_builder.track_c.pipeline batch <glob_pattern> [work_root] [iter_num]

        # 호환 모드 (run 생략 가능)
        python -m ppt_builder.track_c.pipeline <schema.json>
    """
    import glob
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

    if cmd == "review":
        # pptx 직접 평가 (schema 없이 — Phase D 등 Canvas 빌드 파일용)
        if len(sys.argv) < 3:
            print("Usage: python -m ppt_builder.track_c.pipeline review <file.pptx> [--mock]")
            sys.exit(1)
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        pptx_path = Path(sys.argv[2])
        mock = "--mock" in sys.argv
        work_dir = Path("output/track_c/_review") / pptx_path.stem
        work_dir.mkdir(parents=True, exist_ok=True)
        pngs = pptx_to_pngs(pptx_path, work_dir / "pngs")
        eval_report = evaluate_pptx(pptx_path)
        vision = review_all_slides(pngs, mock=mock)
        (work_dir / "review_response.json").write_text(
            json.dumps(vision, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"evaluate.py:  {eval_report['score']}/100")
        print(f"Vision:       {vision['overall_score']}/100")
        print(f"must_fix:     {any(s.get('must_fix') for s in vision.get('slides', []))}")
        for s in vision.get("slides", []):
            cl = s.get("checklist", {})
            items = " | ".join(f"{k}:{v}" for k, v in cl.items())
            print(f"  Slide {s.get('slide_index',0)}: [{items}]")
        print(f"결과 저장: {work_dir / 'review_response.json'}")
        return

    if cmd == "auto":
        if len(sys.argv) < 3:
            print("Usage: python -m ppt_builder.track_c.pipeline auto <schema.json> [work_dir] [max_iter] [--mock]")
            sys.exit(1)
        args = sys.argv[2:]
        schema_path = Path(args[0])
        work_dir = Path(args[1]) if len(args) > 1 and not args[1].startswith("-") else Path("output/track_c") / schema_path.stem
        max_iter_idx = 2 if len(args) > 2 and not args[2].startswith("-") else None
        max_iter = int(args[max_iter_idx]) if max_iter_idx and len(args) > max_iter_idx else 3
        mock = "--mock" in args

        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        result = run_auto_loop(
            schema_path=schema_path,
            work_dir=work_dir,
            max_iterations=max_iter,
            mock=mock,
        )
        print()
        print("=" * 60)
        print(f"Auto Loop 완료: {result['stop_reason']}")
        print(f"최종 iteration: {result['final_iter']}")
        print(f"최종 Vision 점수: {result['final_score']}/100")
        print(f"PASSED: {result['passed']}")
        print("=" * 60)
        return

    if cmd == "batch":
        if len(sys.argv) < 3:
            print("Usage: python -m ppt_builder.track_c.pipeline batch <glob_pattern> [work_root] [iter_num]")
            sys.exit(1)
        pattern = sys.argv[2]
        work_root = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("output/track_c")
        iter_num = int(sys.argv[4]) if len(sys.argv) > 4 else 1

        schemas = sorted(Path(p) for p in glob.glob(pattern))
        if not schemas:
            print(f"패턴에 매칭되는 파일 없음: {pattern}")
            sys.exit(1)

        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        print(f"Batch: {len(schemas)}개 스키마 처리 시작")
        for s in schemas:
            print(f"  - {s}")
        print()

        result = run_batch(
            schemas=schemas,
            work_root=work_root,
            iter_num=iter_num,
        )

        print()
        print("=" * 60)
        print(f"Batch 완료: 성공 {len(result['summary'])} / 실패 {len(result['failed'])}")
        print("=" * 60)
        for item in result["summary"]:
            status = "PASS" if item["pass"] else "FAIL"
            print(f"  [{status}] {item['name']:<20} {item['score']:>3}/100  ({item['slides']} slides, {item['issue_count']} issues)")
        for item in result["failed"]:
            print(f"  [ERR ] {item['name']:<20} {item['error']}")
        print()
        print(f"보고서: {result['report_path']}")
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
