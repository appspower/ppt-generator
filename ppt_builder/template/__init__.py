from .cloner import SlideCloner
from .component_ops import (
    ComponentXML,
    create_blank_slide_with_master_theme,
    extract_group,
    insert_component,
)
from .edit_ops import (
    SlideEditError,
    clone_paragraph,
    del_image,
    del_paragraph,
    iter_leaf_shapes,
    replace_image,
    replace_paragraph,
)
from .editor import TemplateEditor
from .substitutor import TextSubstitutor
