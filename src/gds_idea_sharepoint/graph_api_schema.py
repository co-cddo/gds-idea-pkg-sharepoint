from datetime import datetime
from enum import Enum
from typing import Any, Union, get_args, get_origin

from pydantic import BaseModel


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
    columns: list[dict[str, Any]] = []

    for name, field in model.model_fields.items():
        tp = _unwrap_optional(field.annotation)

        formatted_display_name = name.replace("_", " ").title()
        is_title = name.lower() == "title"

        column: dict[str, Any] = {
            "name": "Title" if is_title else name,
            "displayName": "Title" if is_title else formatted_display_name,
            "description": (field.description or ""),
        }

        # Type mapping (order matters)
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
