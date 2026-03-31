from datetime import datetime
from enum import Enum
from pprint import pprint

from pydantic import AnyHttpUrl, BaseModel, Field

from src.box2.sharepoint.graph_api_schema import generate_graph_schema

# --- Mock Models for Testing ---


class MockStatus(str, Enum):
    OPEN = "open"
    CLOSED = "closed"


class MockModel(BaseModel):
    title: str = Field(description="The main title")
    event_date: datetime = Field(description="When it happens")
    is_urgent: bool = Field(default=False)
    status: MockStatus = Field(default=MockStatus.OPEN)
    long_description: str = Field(description="A very long explanation")
    participants: list[str] = Field(default_factory=list)
    links: list[AnyHttpUrl] = Field(default_factory=list, description="Reference URLs")


def test_title_field_mapping():
    """Verify 'title' is skipped because SharePoint creates a built-in Title column."""
    schema = generate_graph_schema(MockModel, "Test List")
    column_names = [c["name"] for c in schema["columns"]]

    assert "title" not in column_names
    assert "Title" not in column_names


def test_title_case_formatting():
    """Verify snake_case fields are converted to Title Case in displayName."""
    schema = generate_graph_schema(MockModel, "Test List")
    event_col = next(c for c in schema["columns"] if c["name"] == "event_date")

    # name stays snake_case for data mapping, but displayName is pretty
    assert event_col["name"] == "event_date"
    assert event_col["displayName"] == "Event Date"


def test_text_fields_default_to_multiline():
    """Verify standard string fields are set to multiple lines by default."""
    schema = generate_graph_schema(MockModel, "Test List")
    desc_col = next(c for c in schema["columns"] if c["name"] == "long_description")

    assert "text" in desc_col
    assert desc_col["text"]["allowMultipleLines"] is True
    assert desc_col["text"]["textType"] == "plain"


def test_choice_column_mapping():
    """Verify Enums are correctly mapped to SharePoint choice columns."""
    schema = generate_graph_schema(MockModel, "Test List")
    status_col = next(c for c in schema["columns"] if c["name"] == "status")

    assert "choice" in status_col
    assert status_col["choice"]["choices"] == ["open", "closed"]


def test_datetime_column_mapping():
    """Verify datetime types are mapped to dateTime columns."""
    schema = generate_graph_schema(MockModel, "Test List")
    date_col = next(c for c in schema["columns"] if c["name"] == "event_date")

    assert "dateTime" in date_col
    assert "text" not in date_col


def test_list_to_multiline_text():
    """Verify list[str] (arrays) are handled as multiline plain text."""
    schema = generate_graph_schema(MockModel, "Test List")
    part_col = next(c for c in schema["columns"] if c["name"] == "participants")

    assert part_col["text"]["allowMultipleLines"] is True
    assert part_col["text"]["textType"] == "plain"


def test_url_list_to_rich_text():
    """Verify list[AnyHttpUrl] fields produce a rich-text column."""
    schema = generate_graph_schema(MockModel, "Test List")
    links_col = next(c for c in schema["columns"] if c["name"] == "links")

    assert "text" in links_col
    assert links_col["text"]["allowMultipleLines"] is True
    assert links_col["text"]["textType"] == "richText"


def test_debug_schema_output():
    schema = generate_graph_schema(MockModel, "Test List")
    pprint(schema)
