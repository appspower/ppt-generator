"""ComponentCatalog — 추출된 컴퍼넌트의 영구 저장소.

디렉토리 구조:
    component_library/
    ├── catalog.json              # 모든 컴퍼넌트 메타데이터
    ├── components/
    │   ├── <component_id>.xml    # shape XML 페이로드
    │   └── <component_id>/       # (이미지가 있는 경우) blob 디렉토리
    │       ├── meta.json         # {old_rid: filename, content_type}
    │       ├── img1.png
    │       └── img2.jpg

저장: extractor의 결과를 영구화
조회: id, category, subcategory로 검색
로드: 저장된 컴퍼넌트를 Component 객체로 복원
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Iterator

from .models import Component, Slot


_EXT_MAP = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
    "image/x-emf": "emf",
    "image/x-wmf": "wmf",
}


def _ext_for(content_type: str) -> str:
    return _EXT_MAP.get(content_type.lower(), "bin")


def _safe_filename(rid: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]", "_", rid)


class ComponentCatalog:
    """컴퍼넌트 라이브러리 영구 저장소."""

    def __init__(self, root_dir: str | Path):
        self.root = Path(root_dir)
        self.components_dir = self.root / "components"
        self.catalog_path = self.root / "catalog.json"
        self._index: dict[str, Component] = {}

    # ------------------------------------------------------------
    # 저장
    # ------------------------------------------------------------
    def save(self, components: list[Component], merge: bool = True) -> None:
        """컴퍼넌트 리스트를 디스크에 저장한다.

        Args:
            components: 저장할 컴퍼넌트
            merge: True면 기존 카탈로그에 추가, False면 덮어쓰기
        """
        self.root.mkdir(parents=True, exist_ok=True)
        self.components_dir.mkdir(parents=True, exist_ok=True)

        if merge and self.catalog_path.exists():
            self.load()
        else:
            self._index = {}

        for comp in components:
            self._save_one(comp)
            self._index[comp.id] = comp

        self._write_catalog_json()

    def _save_one(self, comp: Component) -> None:
        """개별 컴퍼넌트의 XML + 이미지 blob을 디스크에 쓴다."""
        # XML
        xml_path = self.components_dir / f"{comp.id}.xml"
        xml_path.write_bytes(comp.sp_xml_bytes)

        # 이미지 blob (있는 경우)
        if comp.image_blobs:
            img_dir = self.components_dir / comp.id
            img_dir.mkdir(parents=True, exist_ok=True)

            meta = {}
            for old_rid, blob in comp.image_blobs.items():
                content_type = comp.image_content_types.get(old_rid, "image/png")
                ext = _ext_for(content_type)
                filename = f"{_safe_filename(old_rid)}.{ext}"
                (img_dir / filename).write_bytes(blob)
                meta[old_rid] = {
                    "filename": filename,
                    "content_type": content_type,
                }
            (img_dir / "meta.json").write_text(
                json.dumps(meta, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

    def _write_catalog_json(self) -> None:
        catalog_data = {
            "version": 1,
            "count": len(self._index),
            "components": [comp.to_metadata_dict() for comp in self._index.values()],
        }
        self.catalog_path.write_text(
            json.dumps(catalog_data, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    # ------------------------------------------------------------
    # 로드
    # ------------------------------------------------------------
    def load(self) -> int:
        """카탈로그 + 모든 컴퍼넌트를 메모리로 로드한다.

        Returns:
            로드된 컴퍼넌트 개수
        """
        if not self.catalog_path.exists():
            self._index = {}
            return 0

        data = json.loads(self.catalog_path.read_text(encoding="utf-8"))
        self._index = {}

        for meta in data.get("components", []):
            comp_id = meta["id"]
            comp = self._load_one(comp_id, meta)
            if comp is not None:
                self._index[comp_id] = comp

        return len(self._index)

    def _load_one(self, comp_id: str, meta: dict) -> Component | None:
        """단일 컴퍼넌트를 메모리로 복원한다."""
        xml_path = self.components_dir / f"{comp_id}.xml"
        if not xml_path.exists():
            return None

        sp_xml_bytes = xml_path.read_bytes()

        # 이미지 blob 로드
        image_blobs: dict[str, bytes] = {}
        image_content_types: dict[str, str] = {}
        img_dir = self.components_dir / comp_id
        if img_dir.is_dir() and (img_dir / "meta.json").exists():
            img_meta = json.loads((img_dir / "meta.json").read_text(encoding="utf-8"))
            for old_rid, info in img_meta.items():
                blob_path = img_dir / info["filename"]
                if blob_path.exists():
                    image_blobs[old_rid] = blob_path.read_bytes()
                    image_content_types[old_rid] = info["content_type"]

        # Slot 복원
        slots = []
        for s_meta in meta.get("slots", []):
            slots.append(
                Slot(
                    slot_id=s_meta["slot_id"],
                    semantic_role=s_meta.get("semantic_role", "text"),
                    original_text=s_meta.get("original_text", ""),
                    bbox_emu=tuple(s_meta.get("bbox_emu", [0, 0, 0, 0])),
                    font_size_pt=s_meta.get("font_size_pt", 0.0),
                    font_bold=s_meta.get("font_bold", False),
                    max_chars=s_meta.get("max_chars", 200),
                )
            )

        return Component(
            id=meta["id"],
            category=meta.get("category", "unknown"),
            subcategory=meta.get("subcategory", ""),
            name=meta.get("name", ""),
            source_file=meta.get("source_file", ""),
            source_slide_index=meta.get("source_slide_index", -1),
            bbox_emu=tuple(meta.get("bbox_emu", [0, 0, 0, 0])),
            sp_xml_bytes=sp_xml_bytes,
            image_blobs=image_blobs,
            image_content_types=image_content_types,
            slots=slots,
            color_palette=meta.get("color_palette", []),
            font_families=meta.get("font_families", []),
            text_density=meta.get("text_density", 0),
            shape_count=meta.get("shape_count", 0),
            has_images=meta.get("has_images", False),
            has_charts=meta.get("has_charts", False),
            has_smartart=meta.get("has_smartart", False),
            has_table=meta.get("has_table", False),
        )

    # ------------------------------------------------------------
    # 조회
    # ------------------------------------------------------------
    def get(self, component_id: str) -> Component | None:
        return self._index.get(component_id)

    def find(
        self,
        category: str | None = None,
        subcategory: str | None = None,
        source_file: str | None = None,
    ) -> list[Component]:
        """필터링된 컴퍼넌트 리스트를 반환한다."""
        result = list(self._index.values())
        if category is not None:
            result = [c for c in result if c.category == category]
        if subcategory is not None:
            result = [c for c in result if c.subcategory == subcategory]
        if source_file is not None:
            result = [c for c in result if c.source_file == source_file]
        return result

    def __len__(self) -> int:
        return len(self._index)

    def __iter__(self) -> Iterator[Component]:
        return iter(self._index.values())

    def list_categories(self) -> dict[str, int]:
        """{category: count} 통계."""
        stats: dict[str, int] = {}
        for c in self._index.values():
            key = f"{c.category}/{c.subcategory}" if c.subcategory else c.category
            stats[key] = stats.get(key, 0) + 1
        return stats
