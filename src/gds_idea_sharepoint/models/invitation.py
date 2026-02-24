from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field


class EventType(str, Enum):
    """Type of ministerial event."""

    MEETING = "meeting"
    SPEECH = "speech"
    PANEL = "panel"
    RECEPTION = "reception"
    SITE_VISIT = "site_visit"
    CONFERENCE = "conference"
    OTHER = "other"


class SharepointInvitation(BaseModel):
    """Data schema for the Sharepoint Invitation List"""

    title: str = Field(description="Invitation title")

    document_id: str = Field(description="Unique identifier linking back to the source document")

    event_type: EventType = Field(description="Type of event being invited to")

    host_organisation: str = Field(description="Organization or individual hosting the event")

    purpose: str = Field(description="Stated purpose or description of the event")

    event_summary: str = Field(
        min_length=20,
        description="Concise summary of the event for quick review",
    )

    topics: list[str] = Field(description="Relevant policy topics from the minister's portfolio (DSIT/DESNZ taxonomy)")

    proposed_times: list[str] = Field(
        description="Proposed date/time options as raw text (e.g., '15th February 2026, 6:00 PM', 'week of March 4')"
    )

    is_time_flexible: bool = Field(description="Whether multiple time options are offered or flexibility is mentioned")

    location: str = Field(
        description="Event location as stated in email (e.g., 'The Royal Society, London' or 'Virtual via Zoom')"
    )

    deadline_to_respond: str | None = Field(
        default=None,
        description="Deadline for responding if mentioned (as raw text, e.g., '5th February 2026')",
    )

    model_decision: Literal["accept", "decline", "delegate", "request_more_info", "defer"] = Field(
        description="Recommended action for this invitation"
    )

    priority: Literal["high", "medium", "low"] = Field(description="Priority level of this invitation for the minister")

    reason: str = Field(
        description="Short explanation for the decision (1-2 sentences)",
    )

    draft_response: str = Field(description="Draft email response from LLM")

    affected_events: list[str] = Field(
        default_factory=list,
        description="Titles of calendar events that conflict or are relevant to this decision",
    )

    minister_comment: str | None = Field(
        default=None,
        description="Minister's feedback on the invitation",
    )

    minister_decision: Literal["accept", "decline"] | None = Field(
        default=None,
        description="Minister's decision on whether to accept or decline the invitation",
    )
