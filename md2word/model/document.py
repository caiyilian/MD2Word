from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional, Union


@dataclass
class TextRun:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    underline: bool = False
    strikethrough: bool = False
    superscript: bool = False
    subscript: bool = False
    font_name: Optional[str] = None
    font_size: Optional[int] = None


@dataclass
class Image:
    src: str
    alt: str = ""
    width: Optional[str] = None
    height: Optional[str] = None
    align: Optional[str] = None
    ocr_text: Optional[str] = None


@dataclass
class Hyperlink:
    url: str
    runs: List[TextRun] = field(default_factory=list)


InlineElement = Union[TextRun, Image, Hyperlink]


@dataclass
class Heading:
    level: int
    runs: List[InlineElement] = field(default_factory=list)


@dataclass
class Paragraph:
    runs: List[InlineElement] = field(default_factory=list)
    alignment: Optional[str] = None


@dataclass
class CodeBlock:
    code: str
    language: str = ""


@dataclass
class ListItem:
    elements: List[Union[Paragraph, CodeBlock, Image]] = field(default_factory=list)


@dataclass
class ListBlock:
    ordered: bool
    items: List[ListItem] = field(default_factory=list)
    tight: bool = True


@dataclass
class Table:
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    align: List[Optional[str]] = field(default_factory=list)


@dataclass
class HorizontalRule:
    pass


@dataclass
class Formula:
    latex: str
    display: bool = False
    numbering: Optional[str] = None


@dataclass
class PageBreak:
    pass


@dataclass
class Footnote:
    footnote_id: str
    text: str


@dataclass
class Comment:
    author: str
    text: str
    date: Optional[str] = None
    target: Optional[str] = None


BlockElement = Union[Heading, Paragraph, CodeBlock, ListBlock, Table, Image, HorizontalRule, Formula, PageBreak, Footnote, Comment]


@dataclass
class Document:
    metadata: dict = field(default_factory=dict)
    elements: List[BlockElement] = field(default_factory=list)
    headers: List[str] = field(default_factory=list)
    footers: List[str] = field(default_factory=list)
    sections: List[dict] = field(default_factory=list)
