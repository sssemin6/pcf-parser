from dataclasses import dataclass, field


METADATA_FIELDS = ("hull_no", "ship_type", "plan_title", "plan_no", "comment_party", "approval_status")


@dataclass
class Turn:
    round_label: str = ""
    buyer_comment: str = ""
    builder_reply: str = ""
    has_image: bool = False
    attachments: tuple[str, ...] = ()


@dataclass
class Thread:
    comment_no: str
    source: str
    raw_status: str = ""
    turns: list[Turn] = field(default_factory=list)
    reviewer_initial: str = ""


@dataclass
class Document:
    metadata: dict[str, str]
    threads: list[Thread]
    warnings: list[str]
    raw_preview: list[dict] = field(default_factory=list)
