"""PPT 자동 품질 평가 — Step 5: EVALUATE.

사용법:
  from ppt_builder.evaluate import evaluate_pptx
  report = evaluate_pptx("output/my_slide.pptx")
  print(report)
"""

from pathlib import Path
from pptx import Presentation


def evaluate_pptx(path: str | Path) -> dict:
    """생성된 PPT의 품질을 자동 평가한다.

    Returns:
        {
            "score": int (0~100),
            "pass": bool,
            "issues": list[str],
            "metrics": dict,
        }
    """
    prs = Presentation(str(path))
    issues = []
    metrics = {}

    for si, slide in enumerate(prs.slides):
        slide_issues = _evaluate_slide(slide, si)
        issues.extend(slide_issues)

    # 점수 계산 (100점 - 이슈당 감점)
    deductions = {
        "CRITICAL": 15,
        "HIGH": 8,
        "MEDIUM": 4,
        "LOW": 2,
    }
    total_deduction = sum(deductions.get(iss.split(":")[0], 3) for iss in issues)
    score = max(0, 100 - total_deduction)

    return {
        "score": score,
        "pass": score >= 75,
        "issues": issues,
        "metrics": metrics,
        "slide_count": len(prs.slides),
    }


def _evaluate_slide(slide, slide_idx: int) -> list[str]:
    """단일 슬라이드 평가."""
    issues = []
    shapes = slide.shapes

    # 1. Shape 수
    n_shapes = len(shapes)
    if n_shapes < 5:
        issues.append(f"LOW: Slide {slide_idx+1}: shape 수 부족 ({n_shapes}개)")

    # 2. 텍스트 밀도
    total_chars = 0
    font_sizes = {}
    accent_shapes = 0
    total_colored = 0

    for s in shapes:
        # 색상 분석
        try:
            if s.fill.type is not None:
                rgb = str(s.fill.fore_color.rgb)
                total_colored += 1
                if rgb in ("FD5108", "FE7C39"):  # accent colors
                    accent_shapes += 1
        except:
            pass

        # 텍스트 분석
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                chars = len(p.text)
                total_chars += chars
                if p.font.size:
                    sz = round(p.font.size / 12700, 0)
                    font_sizes[sz] = font_sizes.get(sz, 0) + chars

    # 3. 텍스트 밀도 체크
    if total_chars < 100:
        issues.append(f"HIGH: Slide {slide_idx+1}: 텍스트 부족 ({total_chars}자)")
    elif total_chars > 2000:
        issues.append(f"MEDIUM: Slide {slide_idx+1}: 텍스트 과다 ({total_chars}자)")

    # 4. 폰트 위계 체크
    if font_sizes:
        size_range = max(font_sizes.keys()) - min(font_sizes.keys())
        if size_range < 4:
            issues.append(f"MEDIUM: Slide {slide_idx+1}: 폰트 위계 부족 (범위 {size_range}pt)")

    # 5. 색상 절제 체크
    if total_colored > 0:
        accent_ratio = accent_shapes / total_colored
        if accent_ratio > 0.25:
            issues.append(f"MEDIUM: Slide {slide_idx+1}: accent 과다 ({accent_ratio:.0%})")

    # 6. 빈 공간 체크
    max_bottom = 0
    for s in shapes:
        if s.top and s.height:
            bottom = (s.top + s.height) / 914400
            if bottom > max_bottom:
                max_bottom = bottom
    empty_bottom = 7.5 - max_bottom
    if empty_bottom > 1.5:
        issues.append(f"HIGH: Slide {slide_idx+1}: 하단 빈 공간 과다 ({empty_bottom:.1f}\")")

    # 7. 제목 체크 (첫 번째 텍스트가 인사이트 문장인지)
    title_text = ""
    for s in shapes:
        if s.has_text_frame and s.text_frame.text.strip():
            title_text = s.text_frame.text.strip()
            break
    if title_text and len(title_text) < 10:
        issues.append(f"LOW: Slide {slide_idx+1}: 제목이 너무 짧음 (라벨형?)")

    return issues


def print_report(report: dict):
    """평가 보고서를 출력한다."""
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    status = "PASS" if report["pass"] else "FAIL"
    print(f"\n{'='*50}")
    print(f"PPT Quality: {report['score']}/100 [{status}]")
    print(f"Slides: {report['slide_count']}")
    print(f"{'='*50}")

    if report["issues"]:
        print(f"\nIssues ({len(report['issues'])}):")
        for iss in report["issues"]:
            print(f"  - {iss}")
    else:
        print("\nNo issues")

    return report


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python -m ppt_builder.evaluate <pptx_file>")
        sys.exit(1)
    report = evaluate_pptx(sys.argv[1])
    print_report(report)
