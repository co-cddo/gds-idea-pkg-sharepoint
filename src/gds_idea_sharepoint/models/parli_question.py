from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field


class SharepointPQs(BaseModel):
    """Schema for the Sharepoint parliamentary questions list"""

    uin: str = Field(description="Unique identifier for PQ")

    questiontext: str = Field(description="Parliamentary question")

    house: Literal["commons", "lords"] = Field(description="Member to submit the PQ")

    datetabled: datetime = Field(description="The date the invitation was tabled")

    date_for_answer: datetime = Field(description="The date the invitation was tabled")

    asking_member_id: str = Field(description="ID of member of parliament asking the PQ")

    answering_body_name: str = Field(description="Name of department answering the PQ")

    asking_member_name: str = Field(description="name of member of parliament asking the PQ")

    ai_expansive_answer: str = Field(min_length=55, description="Detailed LLM draft response")

    ai_generic_answer: str = Field(min_length=55, description="Generic LLM draft response")

    url: list[str] = Field(
        min_length=15, default_factory=list, description="List of links used by LLM for draft response"
    )

    ai_predicted_directorate: str = Field(description="Parliamentary question")

    urgency: Literal["urgent", "not urgent"] = Field(
        description="Urgency for the minister to review draft response. Urgent responses <=2 days before dateforAnswer"
    )

    minister_comment: str = Field(
        min_length=55, description="Minister's additional feedback for his decision to approve/request redraft"
    )

    minister_decision: Literal["approve", "request redraft"] | None = Field(
        default=None, description="Minister decision on drafted response quality"
    )
