"""스키마 자동 수정 (Vision 피드백 → JSON patch 적용).

review_response.json의 `patches` 필드를 읽어 원본 스키마에 적용한다.
JSON Patch (RFC 6902)의 단순화된 서브셋 — 외부 의존성 없음.

Patch 연산:
    set    : 값 교체 또는 새 키 추가 (배열 인덱스도 가능)
    delete : 키/배열 요소 제거
    append : 배열 끝에 추가 (path는 배열을 가리킴)
    insert : 배열 특정 인덱스에 삽입 (path는 인덱스로 끝남)

Path 문법:
    슬래시로 분리. 예) "slides/0/sections/1/elements/0/items/0/bullets/3"
    숫자는 배열 인덱스, 그 외는 dict 키.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Any

from ppt_builder.models.schema import PresentationSchema


class PatchError(ValueError):
    """Patch 적용 실패."""


def apply_patches(schema: dict[str, Any], patches: list[dict[str, Any]]) -> dict[str, Any]:
    """원본 스키마(dict)에 patch 리스트를 적용한 새 dict를 반환.

    원본은 변경하지 않는다 (deepcopy).
    적용 후 PresentationSchema로 검증하여 구조 무결성을 확인한다.
    """
    result = copy.deepcopy(schema)

    for i, patch in enumerate(patches):
        try:
            _apply_one(result, patch)
        except (KeyError, IndexError, TypeError) as e:
            raise PatchError(
                f"Patch #{i+1} 적용 실패: {patch}\n  원인: {type(e).__name__}: {e}"
            ) from e

    # Pydantic 검증 — 구조가 깨졌으면 여기서 잡힘
    try:
        PresentationSchema.model_validate(result)
    except Exception as e:
        raise PatchError(
            f"Patch 적용 후 PresentationSchema 검증 실패:\n{e}"
        ) from e

    return result


def _apply_one(obj: Any, patch: dict[str, Any]) -> None:
    """단일 patch를 obj에 in-place 적용."""
    op = patch.get("op")
    path = patch.get("path", "")
    value = patch.get("value")

    if op not in {"set", "delete", "append", "insert"}:
        raise PatchError(f"미지원 op: {op}")
    if not path:
        raise PatchError("path가 비어 있음")

    parts = [p for p in path.split("/") if p != ""]
    if not parts:
        raise PatchError(f"path가 유효하지 않음: {path}")

    # append는 path가 배열을 가리키므로 마지막까지 따라가서 append
    if op == "append":
        target = _navigate(obj, parts)
        if not isinstance(target, list):
            raise PatchError(f"append 대상이 list가 아님: {path}")
        target.append(value)
        return

    # set/delete/insert는 부모를 찾고 마지막 요소로 작업
    parent = _navigate(obj, parts[:-1])
    last = parts[-1]

    if op == "set":
        if isinstance(parent, list):
            parent[int(last)] = value
        else:
            parent[last] = value
    elif op == "delete":
        if isinstance(parent, list):
            del parent[int(last)]
        else:
            del parent[last]
    elif op == "insert":
        if not isinstance(parent, list):
            raise PatchError(f"insert 부모가 list가 아님: {path}")
        parent.insert(int(last), value)


def _navigate(obj: Any, parts: list[str]) -> Any:
    """parts 경로를 따라 obj 내부 노드를 반환."""
    cursor = obj
    for part in parts:
        if isinstance(cursor, list):
            cursor = cursor[int(part)]
        elif isinstance(cursor, dict):
            cursor = cursor[part]
        else:
            raise PatchError(
                f"경로 탐색 중 list/dict가 아닌 노드 만남: {part} (현재 타입 {type(cursor).__name__})"
            )
    return cursor


def refine_schema_file(
    source_schema_path: Path,
    review_response_path: Path,
    output_schema_path: Path,
) -> dict[str, Any]:
    """파일 단위 헬퍼.

    1. source_schema_path 에서 JSON 로드
    2. review_response_path 에서 patches 추출
    3. apply_patches 실행
    4. 결과를 output_schema_path 로 저장

    Returns: 적용 통계 dict
    """
    source_schema_path = Path(source_schema_path)
    review_response_path = Path(review_response_path)
    output_schema_path = Path(output_schema_path)

    if not source_schema_path.exists():
        raise FileNotFoundError(f"source schema 없음: {source_schema_path}")
    if not review_response_path.exists():
        raise FileNotFoundError(f"review response 없음: {review_response_path}")

    schema = json.loads(source_schema_path.read_text(encoding="utf-8"))
    review = json.loads(review_response_path.read_text(encoding="utf-8"))

    patches = review.get("patches", [])
    if not patches:
        raise PatchError(
            f"review_response.json 에 patches 필드가 없거나 비어 있음: {review_response_path}\n"
            "Vision 리뷰어가 자동 적용 가능한 patch를 제공해야 함."
        )

    refined = apply_patches(schema, patches)

    output_schema_path.parent.mkdir(parents=True, exist_ok=True)
    output_schema_path.write_text(
        json.dumps(refined, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return {
        "source_schema": str(source_schema_path),
        "review_response": str(review_response_path),
        "output_schema": str(output_schema_path),
        "patches_applied": len(patches),
        "patch_ops": [p.get("op") for p in patches],
    }
