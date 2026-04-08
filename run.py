"""CLI 진입점 - JSON 스키마 파일로부터 PPT를 생성한다."""

import json
import sys
from pathlib import Path

from ppt_builder import render_presentation
from ppt_builder.models.schema import PresentationSchema


def main():
    if len(sys.argv) < 2:
        print("Usage: python run.py <schema.json> [output.pptx] [template.pptx]")
        sys.exit(1)

    schema_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("output/presentation.pptx")
    template_path = Path(sys.argv[3]) if len(sys.argv) > 3 else None

    with open(schema_path, encoding="utf-8") as f:
        data = json.load(f)

    schema = PresentationSchema.model_validate(data)
    result = render_presentation(schema, template=template_path, output=output_path)
    print(f"PPT 생성 완료: {result}")


if __name__ == "__main__":
    main()
