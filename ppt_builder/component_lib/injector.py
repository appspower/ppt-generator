"""ComponentInjector — 추출된 Component를 슬라이드에 삽입한다.

핵심 책임:
1. shape XML deep copy → 대상 슬라이드 sp_tree에 append
2. 좌표 변환 — 원본 bbox → 대상 bbox로 scale + translate
3. cNvPr/@id 충돌 회피 (슬라이드 내 unique 보장)
4. 이미지 rId remap — APryor6 패턴
   - Component.image_blobs에서 blob을 꺼내 대상 슬라이드의 part에 새 ImagePart 생성
   - 새 rId 할당 후 XML 내 r:embed 속성 치환
5. 텍스트 슬롯 치환

Phase 4 (color normalize), Phase 5 (auto layout)는 별도 단계.
"""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO

from lxml import etree
from pptx.slide import Slide

from ..template.cloner import _add_image_part
from .models import Component, ComponentRequest, EMU_PER_INCH

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

A = lambda tag: f"{{{A_NS}}}{tag}"
P = lambda tag: f"{{{P_NS}}}{tag}"
R = lambda tag: f"{{{R_NS}}}{tag}"


class ComponentInjector:
    """Component를 Slide에 삽입한다."""

    def inject(
        self,
        slide: Slide,
        request: ComponentRequest,
    ) -> etree._Element:
        """Component를 슬라이드에 삽입한다.

        Args:
            slide: 대상 슬라이드
            request: 삽입 요청 (Component, target_bbox, content)

        Returns:
            삽입된 sp_tree element (필요 시 추가 조작용)
        """
        component = request.component

        # 1. XML deep copy (원본 보존)
        new_elm = etree.fromstring(component.sp_xml_bytes)

        # 2. cNvPr ID 재할당 (슬라이드 내 unique 보장)
        self._reassign_shape_ids(slide, new_elm)

        # 3. 이미지 rId remap (APryor6 패턴)
        self._remap_image_rels(slide, new_elm, component)

        # 4. 좌표 변환 (target bbox로 scale + translate)
        if request.target_bbox_inches and any(request.target_bbox_inches[2:]):
            self._transform_geometry(new_elm, component, request.target_bbox_inches)

        # 5. 텍스트 슬롯 치환
        if request.content:
            self._substitute_slots(new_elm, component, request.content)

        # 6. 색상 매핑 (선택)
        if request.color_override:
            self._override_colors(new_elm, request.color_override)

        # 7. 슬라이드 sp_tree에 삽입
        sp_tree = slide.shapes._spTree
        sp_tree.append(new_elm)

        return new_elm

    # ------------------------------------------------------------
    # cNvPr ID 재할당
    # ------------------------------------------------------------
    def _reassign_shape_ids(self, slide: Slide, new_elm: etree._Element) -> None:
        """슬라이드 내 cNvPr/@id가 unique하도록 재할당한다."""
        existing_ids = set()
        sp_tree = slide.shapes._spTree
        for cnv in sp_tree.iter(P("cNvPr")):
            id_attr = cnv.get("id")
            if id_attr and id_attr.isdigit():
                existing_ids.add(int(id_attr))

        next_id = max(existing_ids, default=1) + 1

        for cnv in new_elm.iter(P("cNvPr")):
            cnv.set("id", str(next_id))
            next_id += 1

    # ------------------------------------------------------------
    # 이미지 rId remap
    # ------------------------------------------------------------
    def _remap_image_rels(
        self, slide: Slide, new_elm: etree._Element, component: Component
    ) -> None:
        """그룹 내부의 모든 image rId를 대상 슬라이드의 새 rId로 치환한다.

        모던 PPTX는 한 이미지에 두 개의 rId를 가질 수 있다 (PNG + SVG fallback).
        모든 element의 r:embed/r:link 속성을 검사한다.
        """
        if not component.image_blobs:
            return

        slide_part = slide.part
        rid_map: dict[str, str] = {}

        for old_rid, blob in component.image_blobs.items():
            content_type = component.image_content_types.get(old_rid, "image/png")
            try:
                _, new_rid = _add_image_part(slide_part, blob, content_type)
                rid_map[old_rid] = new_rid
            except Exception as e:
                print(f"[WARN] component image rId remap failed for {old_rid}: {e}")

        embed_attr = R("embed")
        link_attr = R("link")

        # 모든 element 순회하면서 r:embed/r:link 치환
        for el in new_elm.iter():
            old_embed = el.get(embed_attr)
            if old_embed and old_embed in rid_map:
                el.set(embed_attr, rid_map[old_embed])
            old_link = el.get(link_attr)
            if old_link and old_link in rid_map:
                el.set(link_attr, rid_map[old_link])

    # ------------------------------------------------------------
    # 좌표 변환
    # ------------------------------------------------------------
    def _transform_geometry(
        self,
        grp_elm: etree._Element,
        component: Component,
        target_bbox_inches: tuple[float, float, float, float],
    ) -> None:
        """그룹의 위치/크기를 target bbox에 맞춘다.

        그룹의 <p:grpSpPr><a:xfrm>의 off/ext만 수정.
        chOff/chExt는 그대로 두고, off/ext만 변경하면 자식들이 비례 스케일된다.
        """
        target_left, target_top, target_w, target_h = target_bbox_inches

        # EMU 변환
        new_off_x = int(target_left * EMU_PER_INCH)
        new_off_y = int(target_top * EMU_PER_INCH)
        new_ext_cx = int(target_w * EMU_PER_INCH)
        new_ext_cy = int(target_h * EMU_PER_INCH)

        # <p:grpSpPr><a:xfrm> 찾기
        grp_sp_pr = grp_elm.find(P("grpSpPr"))
        if grp_sp_pr is None:
            return

        xfrm = grp_sp_pr.find(A("xfrm"))
        if xfrm is None:
            # xfrm 없으면 새로 만듦
            xfrm = etree.SubElement(grp_sp_pr, A("xfrm"))
            etree.SubElement(xfrm, A("off"))
            etree.SubElement(xfrm, A("ext"))
            etree.SubElement(xfrm, A("chOff"))
            etree.SubElement(xfrm, A("chExt"))

        off = xfrm.find(A("off"))
        ext = xfrm.find(A("ext"))
        ch_off = xfrm.find(A("chOff"))
        ch_ext = xfrm.find(A("chExt"))

        if off is None:
            off = etree.SubElement(xfrm, A("off"))
        if ext is None:
            ext = etree.SubElement(xfrm, A("ext"))

        off.set("x", str(new_off_x))
        off.set("y", str(new_off_y))
        ext.set("cx", str(new_ext_cx))
        ext.set("cy", str(new_ext_cy))

        # chOff/chExt가 있으면 그대로 둠 (자식 좌표는 chOff/chExt 기준)
        # chOff/chExt가 ext와 같으면 1:1 스케일, 다르면 비례 스케일
        # 원본 chOff/chExt를 보존하면 ext가 바뀌어도 자식이 자동 비례
        if ch_off is None:
            ch_off = etree.SubElement(xfrm, A("chOff"))
            ch_off.set("x", str(component.bbox_emu[0]))
            ch_off.set("y", str(component.bbox_emu[1]))
        if ch_ext is None:
            ch_ext = etree.SubElement(xfrm, A("chExt"))
            ch_ext.set("cx", str(component.bbox_emu[2]))
            ch_ext.set("cy", str(component.bbox_emu[3]))

    # ------------------------------------------------------------
    # 텍스트 슬롯 치환
    # ------------------------------------------------------------
    def _substitute_slots(
        self,
        grp_elm: etree._Element,
        component: Component,
        content: dict[str, str],
    ) -> None:
        """슬롯 ID 기반으로 텍스트를 치환한다.

        Phase 1: slot_0, slot_1 순서대로 paragraph에 매칭.
        Phase 3: 의미 기반 매칭으로 고도화 예정.
        """
        # 슬롯 순서 = paragraph 순서 (extractor와 동일 로직)
        slot_idx = 0

        for sp in grp_elm.iter(P("sp")):
            tx_body = sp.find(P("txBody"))
            if tx_body is None:
                continue

            for p in tx_body.iter(A("p")):
                runs = list(p.iter(A("r")))
                if not runs:
                    continue

                # 빈 텍스트 paragraph는 건너뜀
                has_text = any(
                    r.find(A("t")) is not None and (r.find(A("t")).text or "").strip()
                    for r in runs
                )
                if not has_text:
                    continue

                slot_id = f"slot_{slot_idx}"
                slot_idx += 1

                if slot_id not in content:
                    continue

                new_text = content[slot_id]

                # 첫 run에 새 텍스트 통째로 넣고, 나머지 run의 텍스트는 비움
                first_t = runs[0].find(A("t"))
                if first_t is not None:
                    first_t.text = new_text

                for r in runs[1:]:
                    t = r.find(A("t"))
                    if t is not None:
                        t.text = ""

    # ------------------------------------------------------------
    # 색상 매핑 (Phase 4의 단순 버전)
    # ------------------------------------------------------------
    def _override_colors(
        self, grp_elm: etree._Element, color_override: dict[str, str]
    ) -> None:
        """{원본 hex: 새 hex} 매핑으로 srgbClr를 치환한다."""
        # 정규화: # 제거 + 대문자
        norm = {
            k.lstrip("#").upper(): v.lstrip("#").upper()
            for k, v in color_override.items()
        }

        for clr in grp_elm.iter(A("srgbClr")):
            val = (clr.get("val") or "").upper()
            if val in norm:
                clr.set("val", norm[val])
