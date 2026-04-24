"""Track 2 Stage B1 — 1,251장에서 원본 덱 경계 감지.

가정: PwC 마스터 템플릿은 여러 완성 덱이 합본된 상태.
따라서 "opening 슬라이드 → 목차 → 본문 → closing" 패턴이 반복됨.

감지 신호 (복합):
  1. **페이지 번호 리셋**: 하단 page_number가 1로 돌아감 → 새 덱 시작
  2. **opening layout 변화**: layout_name이 cover/title 계열로 변함
  3. **제목 급변**: 연속 title 텍스트 유사도 ↓
  4. **shape 수 급변**: 2배 이상 변화 시 경계 후보

boundary score = 신호 각각의 가중합. threshold 넘으면 경계로 판정.
"""
from __future__ import annotations

import json
from pathlib import Path

from .schemas import DeckBoundary, SlideMeta

# layout 키워드 (대소문자 무관)
OPENING_LAYOUT_HINTS = [
    "title", "cover", "표지", "divider", "section", "chapter",
]
CLOSING_LAYOUT_HINTS = [
    "thank", "end", "closing", "감사", "결론", "마침",
]


def _layout_hints_score(layout_name: str, hints: list[str]) -> float:
    ln = (layout_name or "").lower()
    return 1.0 if any(h.lower() in ln for h in hints) else 0.0


def _page_num_to_int(text: str) -> int | None:
    t = (text or "").strip()
    if not t:
        return None
    # "1" / "P.1" / "1/30" 등에서 앞쪽 숫자 추출
    import re
    m = re.search(r"(\d{1,3})", t)
    return int(m.group(1)) if m else None


def _title_similarity(a: str, b: str) -> float:
    """간단한 Jaccard on tokens (ngram 2)."""
    if not a or not b:
        return 0.0

    def _ngrams(s, n=2):
        s = s.strip()
        return {s[i:i + n] for i in range(len(s) - n + 1)} if len(s) >= n else {s}

    A, B = _ngrams(a), _ngrams(b)
    if not A or not B:
        return 0.0
    inter = len(A & B)
    union = len(A | B)
    return inter / union if union else 0.0


def detect_boundaries(metas: list[SlideMeta], min_deck_size: int = 3) -> list[DeckBoundary]:
    """순차 스캔으로 덱 경계 감지.

    페이지 번호 리셋 + opening layout + 제목 유사도 + shape 급변 신호 복합.
    """
    if not metas:
        return []

    boundaries: list[tuple[int, float, str]] = [(0, 1.0, "start")]

    prev_page = None

    for i in range(1, len(metas)):
        m_prev = metas[i - 1]
        m_curr = metas[i]

        signals: list[str] = []
        score = 0.0

        # 1. page number reset
        curr_pn = _page_num_to_int(m_curr.page_number_text)
        prev_pn = _page_num_to_int(m_prev.page_number_text)
        if curr_pn == 1 and (prev_pn is None or prev_pn >= 2):
            score += 0.5
            signals.append("page_reset")
        elif prev_pn is not None and curr_pn is not None:
            if curr_pn <= prev_pn - 3:  # 페이지가 대폭 역행 (이상)
                score += 0.3
                signals.append("page_backward")

        # 2. opening layout
        if _layout_hints_score(m_curr.layout_name, OPENING_LAYOUT_HINTS) > 0:
            score += 0.3
            signals.append(f"layout:{m_curr.layout_name}")

        # 3. 이전이 closing layout이었다면
        if _layout_hints_score(m_prev.layout_name, CLOSING_LAYOUT_HINTS) > 0:
            score += 0.3
            signals.append("prev_closing")

        # 4. 제목 유사도 급락 (낮을수록 경계)
        title_sim = _title_similarity(m_prev.title_text, m_curr.title_text)
        if title_sim < 0.1 and m_prev.title_text and m_curr.title_text:
            score += 0.2
            signals.append(f"title_sim={title_sim:.2f}")

        # 5. shape 수 급변
        if m_prev.shape_count_total > 0 and m_curr.shape_count_total > 0:
            ratio = m_curr.shape_count_total / m_prev.shape_count_total
            if ratio >= 3 or ratio <= 1 / 3:
                score += 0.2
                signals.append(f"shape_ratio={ratio:.1f}")

        # threshold 0.5 이상이면 경계
        if score >= 0.5:
            boundaries.append((i, min(score, 1.0), ";".join(signals)))

    # 경계 → DeckBoundary 변환
    result: list[DeckBoundary] = []
    for idx, (start_i, conf, signal) in enumerate(boundaries):
        next_start = boundaries[idx + 1][0] if idx + 1 < len(boundaries) else len(metas)
        slide_count = next_start - start_i
        if slide_count < min_deck_size:
            continue  # 너무 짧은 덱은 제외 (잡음)
        result.append(DeckBoundary(
            deck_id=f"deck_{idx + 1:03d}",
            start_index=start_i,
            end_index=next_start - 1,
            slide_count=slide_count,
            detection_signal=signal,
            confidence=conf,
        ))

    return result


def run_detect(meta_path: Path, output_path: Path) -> list[DeckBoundary]:
    """메타데이터 로드 + 경계 감지 + 저장."""
    with open(meta_path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    metas = [SlideMeta.model_validate(r) for r in raw]
    boundaries = detect_boundaries(metas)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(
            [b.model_dump() for b in boundaries],
            f, ensure_ascii=False, indent=2,
        )
    print(f"  detected {len(boundaries)} deck boundaries -> {output_path}")

    # 요약
    sizes = [b.slide_count for b in boundaries]
    if sizes:
        print(f"  deck size: min={min(sizes)}, max={max(sizes)}, avg={sum(sizes)/len(sizes):.1f}")
    return boundaries


if __name__ == "__main__":
    ROOT = Path(__file__).resolve().parent.parent.parent
    meta = ROOT / "output" / "catalog" / "slide_meta.json"
    out = ROOT / "output" / "catalog" / "deck_boundaries.json"
    run_detect(meta, out)
