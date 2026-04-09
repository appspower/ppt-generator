"""Component Library — 외부 .pptx에서 컴퍼넌트 단위로 추출하고 재조립한다.

Phase 1: GroupShape 추출 + 단일 컴퍼넌트 복사
Phase 2: 자동 그룹핑 (DBSCAN, alignment)
Phase 3: 슬롯 인식 (텍스트 치환)
Phase 4: 색상 정규화 (회사 컬러 강제)
Phase 5: 자동 배치 (multi-component)
Phase 6: 카탈로그 빌드

기본 사용:
    from ppt_builder.component_lib import ComponentExtractor, ComponentInjector

    extractor = ComponentExtractor()
    components = extractor.extract_groups("templates/external/strategy.pptx")

    injector = ComponentInjector()
    injector.inject(slide, components[0], target_bbox_inches=(1, 1, 6, 4))
"""

from .models import Component, Slot, ComponentRequest
from .extractor import ComponentExtractor
from .injector import ComponentInjector
from .catalog import ComponentCatalog

__all__ = [
    "Component",
    "Slot",
    "ComponentRequest",
    "ComponentExtractor",
    "ComponentInjector",
    "ComponentCatalog",
]
