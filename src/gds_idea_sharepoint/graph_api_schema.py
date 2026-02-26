import logging
from datetime import datetime
from enum import Enum
from typing import Any, Union, get_args, get_origin

from pydantic import BaseModel

logger = logging.getLogger(__name__)


def _unwrap_optional(tp: Any) -> Any:
    """Optional[T] -> T"""
    origin = get_origin(tp)
    if origin is Union:
        args = [a for a in get_args(tp) if a is not type(None)]
        if len(args) == 1:
            return args[0]
    return tp


def _is_enum_type(tp: Any) -> bool:
    try:
        return isinstance(tp, type) and issubclass(tp, Enum)
    except TypeError:
        return False


def generate_graph_schema(model: type[BaseModel], list_name: str) -> dict[str, Any]:
    """
    Generate a Microsoft Graph list schema from a Pydantic model.

    This version skips 'title' and 'uin' fields during creation because
    SharePoint creates a mandatory 'Title' column by default.

    Iterates over the model's fields and converts them into Graph-compatible
    column definitions, mapping Python types to appropriate column types

    The field named "title" (case-insensitive) is treated specially and mapped
    to the required "Title" column.

    Args:
        model: A Pydantic BaseModel subclass describing the schema.
        list_name: The display name of the resulting list.

    Returns:
        A dictionary representing the Graph list schema payload.

    """
    columns: list[dict[str, Any]] = []

    built_in_redirects = {"title", "uin"}

    for name, field in model.model_fields.items():
        # 1. Skip the redirect fields
        if name.lower() in built_in_redirects:
            logger.info(f"Field '{name}' will be handled by the built-in 'Title' column. Skipping schema entry.")
            continue

        tp = _unwrap_optional(field.annotation)
        formatted_display_name = name.replace("_", " ").title()

        column: dict[str, Any] = {
            "name": name,
            "displayName": formatted_display_name,
            "description": (field.description or ""),
        }

        if _is_enum_type(tp):
            column["choice"] = {
                "choices": [e.value for e in tp],
                "allowTextEntry": False,
                "displayAs": "dropDown",
            }
        elif tp is datetime:
            column["dateTime"] = {}
        elif tp is bool:
            column["boolean"] = {}
        else:
            column["text"] = {"allowMultipleLines": True, "richText": False}

        if field.is_required():
            column["required"] = True

        columns.append(column)

    return {
        "displayName": list_name,
        "columns": columns,
        "list": {"template": "genericList"},
    }
