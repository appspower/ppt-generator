"""Phase B — 디자인 결정 체크리스트.

목적: Claude가 매번 같은 디테일 수준으로 슬라이드를 만들도록 강제한다.
패턴 라이브러리만으로는 "01 번호를 원형으로 할지 사각으로 할지" 같은
결정이 들쭉날쭉해진다. 이 모듈이 그런 결정을 명시화하고,
빌드 후 디자인 품질을 점검한다.

두 종류의 검사:
1. PRE-BUILD (decide_*) — 콘텐츠 spec을 받아 디자인 결정을 추천
   - "이 spec에 어울리는 도형은?" 같은 질문에 명확한 답을 내림
   - Claude는 이걸 호출해서 자기가 빌드할 때 결정 근거를 가짐
2. POST-BUILD (inspect_*) — 빌드된 캔버스를 분석해 문제 검출
   - validate_visual()이 잡지 못하는 디자인 품질 문제
   - 시각 위계 부족, 영역별 빈 공간, 색상 균형, 정보 밀도

기존 검증과의 분담:
- evaluate.py: 스코어 (텍스트 길이, 폰트 크기 등)
- visual_validate.py: 시각 잘림 / 오버플로 / 겹침
- design_check.py: 디자인 품질 (위계, 균형, 결정 일관성)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Optional

from pptx import Presentation


# ============================================================
# Public dataclasses
# ============================================================


@dataclass
class DesignDecision:
    """디자인 결정 결과 — Claude가 빌드 전에 참조."""
    rationale: str
    recommendation: dict
    alternatives: list[dict] = field(default_factory=list)


@dataclass
class DesignIssue:
    severity: Literal["high", "medium", "low"]
    category: str  # hierarchy / balance / density / decision / contrast
    message: str
    suggestion: str = ""

    def __str__(self) -> str:
        return f"[{self.severity.upper()}/{self.category}] {self.message}"


@dataclass
class DesignReport:
    issues: list[DesignIssue] = field(default_factory=list)
    metrics: dict = field(default_factory=dict)

    @property
    def passed(self) -> bool:
        return not any(i.severity == "high" for i in self.issues)

    def summary(self) -> str:
        lines = [f"Design issues: {len(self.issues)}"]
        for i in self.issues[:10]:
            lines.append(f"  - {i}")
            if i.suggestion:
                lines.append(f"    → {i.suggestion}")
        return "\n".join(lines)


# ============================================================
# PRE-BUILD: 콘텐츠 → 디자인 결정 추천
# ============================================================


def decide_number_marker(
    *,
    item_count: int,
    has_sequence: bool,
    space_h: float,
) -> DesignDecision:
    """01/02/03 같은 번호 시퀀스에 어떤 도형을 쓸지 결정.

    Args:
        item_count: 항목 개수
        has_sequence: 순서가 의미를 갖는가? (시간/단계/우선순위)
        space_h: 항목당 사용 가능한 세로 공간 (인치)

    Returns:
        DesignDecision with shape: 'circle' | 'square' | 'chevron'
    """
    # 순서가 명확한 단계 (4개 이하, 가로로 펼쳐짐) → chevron
    if has_sequence and item_count <= 5 and space_h < 0.7:
        return DesignDecision(
            rationale=(
                f"{item_count}개 단계 항목, 단계 의미 명확 → "
                "chevron 화살표가 시퀀스를 시각적으로 가장 잘 표현"
            ),
            recommendation={
                "shape": "chevron",
                "primitive": "Canvas.chevron",
                "size_hint": "h=0.4~0.5, w=균등 분배",
            },
            alternatives=[
                {"shape": "arrow_chain", "reason": "단계가 6개 이상이면"},
            ],
        )

    # 세로 리스트, 항목별 디테일이 있음 → circle
    if space_h >= 0.5:
        return DesignDecision(
            rationale=(
                f"{item_count}개 항목, 항목당 {space_h:.1f}\" 공간 — "
                "원형 번호가 텍스트 옆에 작게 들어가 깔끔한 위계 형성"
            ),
            recommendation={
                "shape": "circle",
                "primitive": "Canvas.circle or Canvas.numbered_list",
                "size_hint": "d=0.32~0.40",
            },
            alternatives=[
                {"shape": "square", "reason": "더 격식 있는 톤이 필요하면"},
            ],
        )

    # 작은 공간 + 다수 항목 → 사각 박스
    return DesignDecision(
        rationale=(
            f"{item_count}개 항목, 좁은 공간 ({space_h:.1f}\") — "
            "사각 박스가 공간 효율 최고"
        ),
        recommendation={
            "shape": "square",
            "primitive": "Canvas.box with text",
            "size_hint": "h=0.3~0.4 정사각",
        },
        alternatives=[
            {"shape": "circle", "reason": "디자인 다양성을 위해"},
        ],
    )


def decide_emphasis_color(
    *,
    palette: Literal["grey", "grey_orange", "monochrome"] = "grey",
    needs_strong_emphasis: bool = False,
) -> DesignDecision:
    """강조 색상 결정. 사용자 정책: 회색 톤 우선, 진한 오렌지 회피."""
    if palette == "grey":
        primary = "grey_900" if needs_strong_emphasis else "grey_800"
        return DesignDecision(
            rationale="회색 위계 정책 — 강조는 다크 그레이로",
            recommendation={
                "primary": primary,
                "secondary": "grey_700",
                "tertiary": "grey_400",
                "background": "grey_100",
            },
        )
    if palette == "monochrome":
        return DesignDecision(
            rationale="순수 모노크롬",
            recommendation={
                "primary": "black",
                "secondary": "grey_700",
                "tertiary": "grey_400",
                "background": "white",
            },
        )
    # grey_orange — 오렌지 사용 허용 (현재 사용자는 비선호)
    return DesignDecision(
        rationale="회색+오렌지 (사용자 정책상 비선호 — 명시 요청 시만)",
        recommendation={
            "primary": "accent",
            "secondary": "grey_800",
            "tertiary": "grey_400",
            "background": "white",
        },
    )


def decide_density(
    *,
    available_area: float,  # sq-in
    intended_chars: int,
) -> DesignDecision:
    """공간 대비 텍스트 밀도가 적절한지 판단.

    가이드:
    - 너무 빈약: 100자 / sq-in 미만 → 콘텐츠 추가 필요
    - 적절: 100~300자 / sq-in
    - 과밀: 300자 / sq-in 초과 → 폰트 줄이거나 박스 키우기
    """
    if available_area <= 0:
        return DesignDecision(
            rationale="영역 없음",
            recommendation={"action": "n/a"},
        )
    density = intended_chars / available_area
    if density < 100:
        return DesignDecision(
            rationale=(
                f"밀도 {density:.0f} chars/sq-in — 너무 빈약. "
                "sub-bullet, mini-list, 또는 보조 정보로 채워야 함"
            ),
            recommendation={
                "action": "enrich",
                "suggestions": [
                    "sub-bullet 2~3개 추가",
                    "보조 KPI 또는 stat_block 추가",
                    "mini-callout으로 디테일 보강",
                ],
            },
        )
    if density > 300:
        return DesignDecision(
            rationale=(
                f"밀도 {density:.0f} chars/sq-in — 과밀. "
                "텍스트 압축 또는 영역 확장 필요"
            ),
            recommendation={
                "action": "reduce_or_expand",
                "suggestions": [
                    "텍스트 압축 (불필요 단어 제거)",
                    "영역 확장",
                    "두 영역으로 분할",
                ],
            },
        )
    return DesignDecision(
        rationale=f"밀도 {density:.0f} chars/sq-in — 적절",
        recommendation={"action": "ok"},
    )


def decide_layout_archetype(
    *,
    intent: Literal[
        "executive", "timeline", "comparison", "process", "quadrant", "data"
    ],
    item_count: int,
) -> DesignDecision:
    """슬라이드 의도(intent)와 항목 개수에 따라 적합한 레이아웃 패턴 추천."""
    rules = {
        "executive": {
            "pattern": "hero_with_metrics",
            "layout": "left hero (35-45%) + right metric grid + bottom takeaway",
            "rationale": "C-level은 결론+근거 숫자가 즉시 보여야 함",
        },
        "timeline": {
            "pattern": "horizontal_phases_with_deliverables",
            "layout": "top chevron sequence + bottom per-phase deliverable lists",
            "rationale": "시간 흐름은 가로축이 자연스럽고, 각 단계 산출물이 함께 보여야 신뢰",
        },
        "comparison": {
            "pattern": "side_by_side_or_matrix",
            "layout": (
                "2개: 좌우 callout / "
                "3-4개: 균등 columns / "
                "5+: framework matrix"
            ),
            "rationale": "비교는 동일 차원에서 나란히 보여야 함",
        },
        "process": {
            "pattern": "horizontal_arrow_chain_with_callouts",
            "layout": "arrow_chain 위 + 각 단계 옆 callout",
            "rationale": "프로세스는 흐름 방향이 명확해야 함",
        },
        "quadrant": {
            "pattern": "2x2_with_center_insight",
            "layout": "2x2 grid + 중앙 또는 하단 인사이트 박스",
            "rationale": "사분면은 두 축의 의미가 강조되어야 함",
        },
        "data": {
            "pattern": "chart_with_callouts",
            "layout": "큰 차트/표 + 옆 인사이트 stat_blocks",
            "rationale": "데이터는 숫자만 있으면 의미 전달 안 됨, 인사이트 동반",
        },
    }
    rec = rules[intent]
    return DesignDecision(
        rationale=rec["rationale"],
        recommendation=rec,
    )


# ============================================================
# POST-BUILD: 캔버스 분석으로 디자인 품질 점검
# ============================================================


def inspect_design(
    pptx_path: str,
    *,
    expect_dense: bool = True,
    forbidden_colors: Optional[list[str]] = None,
    pattern_kind: Optional[str] = None,
) -> DesignReport:
    """렌더된 .pptx를 디자인 관점에서 점검한다.

    검사 항목:
    1. 색상 균형 — 전체에서 어떤 색이 얼마나 쓰였는가
    2. 영역별 빈 공간 — 슬라이드를 4분면으로 나눠 콘텐츠 없는 영역 검출
    3. 시각 위계 — 폰트 크기 분포가 충분한 단계를 가지는가
    4. 정보 밀도 — 슬라이드 전체 밀도 (expect_dense=True일 때)
       pattern_kind로 패턴별 임계값 다르게 적용 가능
    5. 금지색 사용 — forbidden_colors에 들어 있는 색이 쓰였는가

    Args:
        pattern_kind: 'executive' | 'timeline' | 'comparison' | 'process' |
                      'quadrant' | 'data' — 패턴별 density 임계값 조정.
                      comparison/quadrant는 짧은 셀+여백이 본질이라 낮은
                      밀도가 정상.
    """
    # 패턴별 density 임계값
    # 기본: 12 (그 미만이면 너무 빈약)
    # comparison/quadrant: 6 (셀이 짧고 여백이 의미)
    # data: 8 (차트가 면적 대부분 차지)
    density_floors = {
        "executive": 12,
        "timeline": 12,
        "comparison": 6,
        "quadrant": 6,
        "process": 10,
        "data": 8,
    }
    density_floor = density_floors.get(pattern_kind or "", 12)
    issues: list[DesignIssue] = []
    metrics: dict = {}

    if forbidden_colors is None:
        # 사용자 정책: 진한 오렌지 회피
        forbidden_colors = ["FD5108", "FE7C39"]

    prs = Presentation(pptx_path)

    for si, slide in enumerate(prs.slides):
        sn = si + 1
        slide_w = prs.slide_width / 914400
        slide_h = prs.slide_height / 914400

        # ---------- 1. 색상 분포 ----------
        color_use: dict[str, float] = {}  # color → total area
        forbidden_hits: list[tuple[str, str]] = []  # (color, where)
        font_sizes: list[float] = []
        total_chars = 0

        # ---------- 2. 영역별 점유 ----------
        # 4분면 (TL, TR, BL, BR) 각 영역의 사용 면적
        quadrants = {"TL": 0.0, "TR": 0.0, "BL": 0.0, "BR": 0.0}
        slide_area = slide_w * slide_h

        for shape in slide.shapes:
            if shape.left is None or shape.top is None:
                continue
            l = shape.left / 914400
            t = shape.top / 914400
            w = (shape.width or 0) / 914400
            h = (shape.height or 0) / 914400
            if w <= 0 or h <= 0:
                continue
            area = w * h

            # 색상
            try:
                if shape.fill.type is not None:
                    rgb = str(shape.fill.fore_color.rgb)
                    color_use[rgb] = color_use.get(rgb, 0) + area
                    if rgb.upper() in (c.upper() for c in forbidden_colors):
                        forbidden_hits.append((rgb, f"shape@{l:.1f},{t:.1f}"))
            except Exception:
                pass

            # 4분면 — shape 중심점이 어디에 속하는가
            cx = l + w / 2
            cy = t + h / 2
            row = "T" if cy < slide_h / 2 else "B"
            col = "L" if cx < slide_w / 2 else "R"
            quadrants[row + col] += area

            # 텍스트 분석
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                total_chars += len(txt)
                for p in shape.text_frame.paragraphs:
                    if p.font.size:
                        font_sizes.append(p.font.size / 12700)

        # ---------- 1. 금지색 ----------
        if forbidden_hits:
            issues.append(
                DesignIssue(
                    severity="high",
                    category="contrast",
                    message=f"Slide {sn}: 금지색 사용 ({len(forbidden_hits)}건) — {forbidden_hits[0][0]}",
                    suggestion="회색 위계로 교체 (grey_700/800/900)",
                )
            )

        # ---------- 2. 색상 균형 ----------
        # 가장 큰 비중 색이 70% 초과면 모노톤 과다
        if color_use:
            total_colored = sum(color_use.values())
            if total_colored > 0:
                top_color, top_area = max(color_use.items(), key=lambda kv: kv[1])
                top_ratio = top_area / total_colored
                metrics[f"slide_{sn}_top_color"] = (top_color, round(top_ratio, 2))
                if top_ratio > 0.85 and top_color.upper() not in ("FFFFFF",):
                    issues.append(
                        DesignIssue(
                            severity="medium",
                            category="balance",
                            message=f"Slide {sn}: 단일 색이 {top_ratio:.0%} 점유 — 시각 위계 약함",
                            suggestion="2~3 단계 회색 위계 추가 (grey_200/grey_700 등)",
                        )
                    )

        # ---------- 3. 영역 균형 ----------
        # 4분면 점유율 계산 — 임계값 3% (작은 컴포넌트도 있을 수 있음)
        for q, area in quadrants.items():
            ratio = area / slide_area if slide_area > 0 else 0
            metrics[f"slide_{sn}_{q}"] = round(ratio, 2)
            if ratio < 0.03:
                issues.append(
                    DesignIssue(
                        severity="medium",
                        category="balance",
                        message=f"Slide {sn}: {q} 영역이 거의 비어 있음 ({ratio:.0%})",
                        suggestion=f"{q}에 보조 콘텐츠 (stat_block, callout, badge) 추가 검토",
                    )
                )

        # ---------- 4. 시각 위계 ----------
        if font_sizes:
            unique_sizes = sorted(set(round(s) for s in font_sizes))
            metrics[f"slide_{sn}_font_steps"] = unique_sizes
            if len(unique_sizes) < 3:
                issues.append(
                    DesignIssue(
                        severity="medium",
                        category="hierarchy",
                        message=f"Slide {sn}: 폰트 위계 단계 {len(unique_sizes)}개 — 위계 부족",
                        suggestion="제목/부제/본문/디테일 4단계 권장",
                    )
                )
            elif max(unique_sizes) - min(unique_sizes) < 5:
                issues.append(
                    DesignIssue(
                        severity="low",
                        category="hierarchy",
                        message=f"Slide {sn}: 폰트 크기 범위 좁음 ({min(unique_sizes)}~{max(unique_sizes)}pt)",
                        suggestion="제목과 본문 차이를 6pt 이상으로 키워야 위계 강화",
                    )
                )

        # ---------- 5. 정보 밀도 (패턴별 임계값) ----------
        # 임계값 근거: v3 (사용자 OK한 밀도)가 ~19 chars/sq-in, v1 (빈약)이 ~10
        # comparison/quadrant는 짧은 셀이 정상 → density_floor=6
        if expect_dense:
            content_area = slide_area * 0.7  # 헤더/푸터 제외 추정
            density = total_chars / content_area if content_area > 0 else 0
            metrics[f"slide_{sn}_density"] = round(density, 1)
            if density < density_floor:
                issues.append(
                    DesignIssue(
                        severity="high",
                        category="density",
                        message=f"Slide {sn}: 정보 밀도 매우 낮음 ({density:.0f} chars/sq-in, 패턴 floor {density_floor})",
                        suggestion="콘텐츠 대폭 보강 필요 — sub-bullet, stat_block, mini-list",
                    )
                )
            elif density > 80:
                issues.append(
                    DesignIssue(
                        severity="medium",
                        category="density",
                        message=f"Slide {sn}: 정보 밀도 과다 ({density:.0f} chars/sq-in)",
                        suggestion="텍스트 압축 또는 영역 확장",
                    )
                )

    return DesignReport(issues=issues, metrics=metrics)


# ============================================================
# 통합 워크플로 헬퍼
# ============================================================


def design_check_pipeline(
    pptx_path: str,
    *,
    expect_dense: bool = True,
    forbidden_colors: Optional[list[str]] = None,
) -> DesignReport:
    """convenience: render_validated 후 호출하여 디자인 점검까지 한 번에."""
    return inspect_design(
        pptx_path,
        expect_dense=expect_dense,
        forbidden_colors=forbidden_colors,
    )
