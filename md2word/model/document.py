from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional, Union


@dataclass
class TextRun:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass
class Heading:
    level: int
    runs: List[TextRun] = field(default_factory=list)


@dataclass
class Paragraph:
    runs: List[TextRun] = field(default_factory=list)


@dataclass
class CodeBlock:
    code: str
    language: str = ""


@dataclass
class ListItem:
    elements: List[Union[Paragraph, CodeBlock]] = field(default_factory=list)


@dataclass
class ListBlock:
    ordered: bool
    items: List[ListItem] = field(default_factory=list)
    tight: bool = True


@dataclass
class HorizontalRule:
    pass


BlockElement = Union[Heading, Paragraph, CodeBlock, ListBlock, HorizontalRule]


@dataclass
class Document:
    metadata: dict = field(default_factory=dict)
    elements: List[BlockElement] = field(default_factory=list)
