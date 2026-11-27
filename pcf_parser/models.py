from dataclasses import dataclass, field


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


@dataclass
class Document:
    metadata: dict[str, str]
    threads: list[Thread]
    warnings: list[str]
