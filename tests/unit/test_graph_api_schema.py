from datetime import datetime
from enum import StrEnum
from pprint import pprint
from typing import Optional, Union

import pytest
from pydantic import AnyHttpUrl, AnyUrl, BaseModel, Field, HttpUrl

from gds_idea_sharepoint.graph_api_schema import contains_url_type, generate_graph_schema, unwrap_optional

# --- Mock Models for Testing ---


class MockStatus(StrEnum):
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


# ===== Public type helpers =====


def test_unwrap_optional_strips_typing_optional():
    """Optional[T] is unwrapped to T."""
    assert unwrap_optional(Optional[str]) is str  # noqa: UP045 - deliberately testing the legacy form


@pytest.mark.parametrize("tp", [str, int, list[str], str | int, int | str | None])
def test_unwrap_optional_leaves_other_types_unchanged(tp):
    """Non-optional types and multi-type unions are returned as-is."""
    assert unwrap_optional(tp) == tp


@pytest.mark.parametrize(
    ("tp", "expected"),
    [
        (str | None, str),
        (None | str, str),
        (Union[str, None], str),  # noqa: UP007 - deliberately testing the legacy form
        (list[str] | None, list[str]),
        (list[AnyHttpUrl] | None, list[AnyHttpUrl]),
        (MockStatus | None, MockStatus),
        (datetime | None, datetime),
    ],
)
def test_unwrap_optional_unwraps_every_spelling_of_optional(tp, expected):
    """``T | None``, ``None | T`` and ``Union[T, None]`` unwrap identically on every Python version."""
    assert unwrap_optional(tp) == expected


def test_contains_url_type_sees_through_optional_list_of_urls():
    """An optional list of URLs is still recognised as a URL field once unwrapped."""
    assert contains_url_type(unwrap_optional(list[AnyHttpUrl] | None)) is True


# ===== Optional fields in generated schemas =====


class OptionalFieldsModel(BaseModel):
    """Every optional (``T | None``) column type the generator special-cases."""

    status: MockStatus | None = Field(default=None, description="optional enum")
    when: datetime | None = Field(default=None)
    flag: bool | None = Field(default=None)
    links: list[AnyHttpUrl] | None = Field(default=None)
    note: str | None = Field(default=None)


def _columns(model: type[BaseModel]) -> dict[str, dict]:
    return {c["name"]: c for c in generate_graph_schema(model, "L")["columns"]}


def test_optional_enum_is_a_dropdown():
    """``Enum | None`` produces a choice column, the same as a required enum."""
    column = _columns(OptionalFieldsModel)["status"]
    assert column["choice"]["choices"] == ["open", "closed"]
    assert "required" not in column


def test_optional_datetime_and_bool_get_typed_columns():
    """``datetime | None`` and ``bool | None`` produce dateTime and boolean columns."""
    columns = _columns(OptionalFieldsModel)
    assert "dateTime" in columns["when"]
    assert "boolean" in columns["flag"]


def test_optional_url_list_is_rich_text():
    """``list[AnyHttpUrl] | None`` is a rich-text column so links render."""
    assert _columns(OptionalFieldsModel)["links"]["text"]["textType"] == "richText"


def test_optional_plain_text_is_unchanged():
    """``str | None`` stays a plain multi-line text column."""
    assert _columns(OptionalFieldsModel)["note"]["text"] == {"allowMultipleLines": True, "textType": "plain"}


def test_optional_and_required_fields_produce_the_same_column_shape():
    """Making a field optional changes only whether it is required, never its column type."""

    class Required(BaseModel):
        status: MockStatus
        when: datetime
        flag: bool

    class Optional_(BaseModel):  # noqa: N801
        status: MockStatus | None = None
        when: datetime | None = None
        flag: bool | None = None

    for name, required_column in _columns(Required).items():
        optional_column = _columns(Optional_)[name]
        assert {k: v for k, v in required_column.items() if k != "required"} == optional_column


@pytest.mark.parametrize("tp", [AnyHttpUrl, HttpUrl, AnyUrl, list[AnyHttpUrl], list[HttpUrl]])
def test_contains_url_type_true_for_url_types(tp):
    """Bare URL types and generic aliases wrapping them are detected."""
    assert contains_url_type(tp) is True


@pytest.mark.parametrize("tp", [str, int, list[str], datetime])
def test_contains_url_type_false_for_other_types(tp):
    """Non-URL types are not detected as URL types."""
    assert contains_url_type(tp) is False


def test_helpers_exported_from_package():
    """unwrap_optional and contains_url_type are part of the public sharepoint API."""
    import gds_idea_sharepoint as sp

    assert sp.unwrap_optional is unwrap_optional
    assert sp.contains_url_type is contains_url_type
