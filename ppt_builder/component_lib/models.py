"""Component Library 데이터 모델.

핵심 개념:
- Component: 외부 .pptx에서 추출한 재사용 가능한 단위 (GroupShape 또는 자동 그룹)
- Slot: Component 내부의 텍스트 치환 위치
- ComponentRequest: Component를 슬라이드에 삽입할 때의 사용자 요청
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any


# EMU per inch (PowerPoint 표준)
EMU_PER_INCH = 914400


@dataclass
class Slot:
    """Component 내부의 텍스트 치환 위치."""

    slot_id: str                                  # "title", "q1_label", "kpi_value_1"
    semantic_role: str = "text"                   # "title" | "label" | "value" | "description" | "text"
    original_text: str = ""                       # 추출 당시 텍스트 (참조용)
    bbox_emu: tuple[int, int, int, int] = (0, 0, 0, 0)  # (left, top, width, height) in EMU
    font_size_pt: float = 0.0
    font_bold: bool = False
    max_chars: int = 200                          # 안전 한도 (휴리스틱)


@dataclass
class Component:
    """외부 .pptx에서 추출한 재사용 가능한 컴퍼넌트.

    XML 직렬화는 sp_tree_xml에 저장 (여러 shape의 묶음).
    이미지가 포함된 경우 image_blobs에 blob 데이터 별도 저장.
    """

    # 식별자
    id: str                                       # "swot_2x2_strategy_s12"
    category: str = "unknown"                     # "framework" | "kpi" | "timeline" | "callout" | "chart"
    subcategory: str = ""                         # "swot" | "bcg_matrix" | "value_chain"
    name: str = ""                                # human-readable name

    # 출처
    source_file: str = ""                         # "smartsheet_strategy.pptx"
    source_slide_index: int = -1                  # 0-based

    # Geometry (원본 좌표, EMU)
    bbox_emu: tuple[int, int, int, int] = (0, 0, 0, 0)  # (left, top, width, height)

    # XML 페이로드 (직렬화된 shape XML 묶음 — bytes로 저장)
    sp_xml_bytes: bytes = b""

    # 이미지 의존성: {old_rId: image_blob_bytes}
    image_blobs: dict[str, bytes] = field(default_factory=dict)
    # 각 image의 content_type (예: "image/png")
    image_content_types: dict[str, str] = field(default_factory=dict)

    # 슬롯 (텍스트 치환 위치)
    slots: list[Slot] = field(default_factory=list)

    # 메타데이터 (자동 분석 결과)
    color_palette: list[str] = field(default_factory=list)   # ["#FD5108", "#404040", ...]
    font_families: list[str] = field(default_factory=list)
    text_density: int = 0                         # 총 글자 수
    shape_count: int = 0                          # 자식 shape 개수
    has_images: bool = False
    has_charts: bool = False
    has_smartart: bool = False
    has_table: bool = False

    # 추출 시점의 EMU 단위 — Phase 5에서 비율 계산용
    @property
    def aspect_ratio(self) -> float:
        _, _, w, h = self.bbox_emu
        return (w / h) if h > 0 else 1.0

    @property
    def width_inches(self) -> float:
        return self.bbox_emu[2] / EMU_PER_INCH

    @property
    def height_inches(self) -> float:
        return self.bbox_emu[3] / EMU_PER_INCH

    def to_metadata_dict(self) -> dict[str, Any]:
        """JSON 직렬화 가능한 메타데이터만 반환 (sp_xml_bytes/image_blobs 제외)."""
        d = asdict(self)
        d.pop("sp_xml_bytes", None)
        d.pop("image_blobs", None)
        d["bbox_emu"] = list(self.bbox_emu)
        d["slots"] = [
            {**asdict(s), "bbox_emu": list(s.bbox_emu)} for s in self.slots
        ]
        return d


@dataclass
class ComponentRequest:
    """Component를 슬라이드에 삽입할 때의 사용자 요청."""

    component: Component
    target_bbox_inches: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    # (left, top, width, height) in inches
    # width/height = 0이면 원본 크기 유지

    content: dict[str, str] = field(default_factory=dict)
    # {slot_id: replacement_text}

    normalize_colors: bool = False
    # True면 회사 컬러 팔레트로 변환 (Phase 4)

    color_override: dict[str, str] = field(default_factory=dict)
    # {original_hex: new_hex} — 명시적 색상 매핑
