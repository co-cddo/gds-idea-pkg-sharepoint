from typing import Any

from pydantic import BaseModel


def generate_graph_schema(model: type[BaseModel], list_name: str) -> dict[str, Any]:
    """
    Generates a Graph API schema where:
    1. Display names are Title Cased (e.g., event_type -> Event Type).
    2. Text fields are set to Multiple Lines by default.
    """
    pydantic_schema = model.model_json_schema()
    properties = pydantic_schema.get("properties", {})
    required_fields = pydantic_schema.get("required", [])

    columns = []

    for name, info in properties.items():
        formatted_display_name = name.replace("_", " ").title()

        is_title = name.lower() == "title"

        column = {
            "name": "Title" if is_title else name,
            "displayName": "Title" if is_title else formatted_display_name,
            "description": info.get("description", ""),
        }

        if "enum" in info:
            column["choice"] = {"choices": info["enum"], "allowTextEntry": False, "displayAs": "dropDown"}
        elif info.get("format") == "date-time":
            column["dateTime"] = {}
        elif info.get("type") == "boolean":
            column["boolean"] = {}
        else:
            column["text"] = {"allowMultipleLines": True, "richText": False}

        if name in required_fields:
            column["required"] = True

        columns.append(column)

    return {"displayName": list_name, "columns": columns, "list": {"template": "genericList"}}