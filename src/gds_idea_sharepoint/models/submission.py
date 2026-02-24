from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field


class SharepointSubmission(BaseModel):
    """Schema for the Sharepoint submission list"""

    # Metadata
    document_id: str = Field(description="Unique identifier linking back to the source document ")

    policy_area: str = Field(
        description="Primary policy area covered by this submission "
        "(e.g., 'AI Safety and International Collaboration', 'Horizon Europe', 'Quantum Research')",
    )
    responsible_deputy_director: str = Field(
        description="Name and title of the Deputy Director or official submitting this document "
        "(e.g., 'Jane Smith, Deputy Director - AI Policy')",
    )
    summary: str = Field(
        description="One-paragraph summary extracted from the submission document, "
        "including key context, background, and main points",
    )

    submission_date: datetime = Field(
        default_factory=datetime.now,
        description="Date the submission was created or received",
    )
    decision_deadline: str | None = Field(
        default=None,
        description="Deadline for minister's decision as raw text from document "
        "(e.g., '7 February 2026 (Treasury deadline)', 'By end of week')",
    )
    key_dates: list[str] = Field(
        default_factory=list,
        description="List of important dates mentioned in submission with context "
        "(e.g., '10 Feb: International announcement', '15 Feb: Contracts must be signed')",
    )

    required_decisions: list[str] = Field(
        default_factory=list,
        description="Specific decisions that the minister must make "
        "(e.g., 'Approve £3M from contingency reserve', 'Choose between Option A and B')",
    )
    official_recommendation: str = Field(
        min_length=5,
        description="The official recommendation from the submission author "
        "(e.g., 'Approve £3M from contingency reserve', 'Note the report; maintain current approach')",
    )

    urgency: Literal["urgent", "routine"] = Field(
        description="Urgency level indicated in the document",
    )
    related_items: list[str] = Field(
        default_factory=list,
        description="References to similar previous submissions or related policy items "
        "(e.g., 'Previous contingency draw-down for quantum centre (Oct 2025)')",
    )

    overall_confidence: float | None = Field(
        default=None,
        ge=0.0,
        le=1.0,
        description="LLM's confidence in extraction quality (0.0-1.0). "
        "1.0 = very clear submission format, 0.5 = some ambiguity, 0.2 = unclear or incomplete",
    )

    minister_comment: str | None = Field(
        default=None, description="Ministers recommendation on how submission should be handled"
    )
