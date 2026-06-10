from __future__ import annotations
import re
from typing import List, Optional

import yaml
import mistune


class _ASTRenderer:
    """Custom renderer that returns raw token list."""
    def __call__(self, tokens, state):
        return list(tokens)

from md2word.model.document import (
    TextRun, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, HorizontalRule, Document,
    BlockElement,
)
from md2word.exceptions import ParseError


def parse_frontmatter(text: str):
    match = re.match(r'^---\s*\n(.*?)\n---\s*\n', text, re.DOTALL)
    if match:
        try:
            metadata = yaml.safe_load(match.group(1))
            rest = text[match.end():]
            return metadata or {}, rest
        except yaml.YAMLError:
            pass
    return {}, text


class MarkdownParser:
    def __init__(self):
        self.md = mistune.Markdown(renderer=_ASTRenderer())

    def parse(self, text: str) -> Document:
        metadata, body = parse_frontmatter(text)
        try:
            ast = self.md(body)
        except Exception as e:
            raise ParseError(f"Failed to parse markdown: {e}") from e
        elements = self._parse_blocks(ast)
        return Document(metadata=metadata, elements=elements)

    def _parse_blocks(self, tokens: list) -> List[BlockElement]:
        elements: List[BlockElement] = []
        for token in tokens:
            element = self._parse_block(token)
            if element is not None:
                elements.append(element)
        return elements

    def _parse_block(self, token: dict) -> Optional[BlockElement]:
        t = token.get("type", "")

        if t == "heading":
            return self._parse_heading(token)
        elif t == "paragraph":
            return self._parse_paragraph(token)
        elif t == "block_text":
            return self._parse_paragraph(token)
        elif t == "block_code":
            return self._parse_code_block(token)
        elif t == "list":
            return self._parse_list(token)
        elif t == "thematic_break":
            return HorizontalRule()
        elif t == "blank_line":
            return None
        elif t == "block_quote":
            return self._parse_block_quote(token)
        else:
            return None

    def _parse_heading(self, token: dict) -> Heading:
        level = token.get("attrs", {}).get("level", 1)
        runs = self._parse_inline(token.get("children", []))
        return Heading(level=level, runs=runs)

    def _parse_paragraph(self, token: dict) -> Paragraph:
        runs = self._parse_inline(token.get("children", []))
        return Paragraph(runs=runs)

    def _parse_code_block(self, token: dict) -> CodeBlock:
        attrs = token.get("attrs", {})
        language = attrs.get("info", "")
        raw = token.get("raw", "")
        children = token.get("children", [])
        if raw:
            code = raw
        elif children:
            code = "".join(c.get("raw", "") for c in children)
        else:
            code = ""
        return CodeBlock(code=code.rstrip("\n"), language=language)

    def _parse_block_quote(self, token: dict) -> BlockElement:
        children = token.get("children", [])
        quote_text = ""
        for child in children:
            ct = child.get("type", "")
            if ct == "paragraph":
                runs = self._parse_inline(child.get("children", []))
                quote_text += "".join(r.text for r in runs) + "\n"
            elif ct == "block_text":
                runs = self._parse_inline(child.get("children", []))
                quote_text += "".join(r.text for r in runs) + "\n"
        # Render blockquote as a normal paragraph for now
        return Paragraph(runs=[TextRun(text=quote_text.strip(), italic=True)])

    def _parse_list(self, token: dict) -> ListBlock:
        attrs = token.get("attrs", {})
        ordered = attrs.get("ordered", False)
        tight = attrs.get("tight", True)
        items: List[ListItem] = []
        for child in token.get("children", []):
            if child.get("type") == "list_item":
                item = self._parse_list_item(child)
                if item is not None:
                    items.append(item)
        return ListBlock(ordered=ordered, items=items, tight=tight)

    def _parse_list_item(self, token: dict) -> Optional[ListItem]:
        children = token.get("children", [])
        elements: List[BlockElement] = []
        for child in children:
            ct = child.get("type", "")
            if ct == "paragraph":
                elements.append(self._parse_paragraph(child))
            elif ct == "block_text":
                elements.append(self._parse_paragraph(child))
            elif ct == "list":
                elements.append(self._parse_list(child))
            elif ct == "block_code":
                elements.append(self._parse_code_block(child))
        if not elements:
            return None
        return ListItem(elements=elements)

    def _parse_inline(self, tokens: list, bold: bool = False,
                      italic: bool = False, code: bool = False) -> List[TextRun]:
        runs: List[TextRun] = []
        for token in tokens:
            t = token.get("type", "")
            if t == "text":
                runs.append(TextRun(
                    text=token.get("raw", ""),
                    bold=bold, italic=italic, code=code,
                ))
            elif t == "strong":
                children = token.get("children", [])
                runs.extend(self._parse_inline(children, bold=True, italic=italic, code=code))
            elif t == "emphasis":
                children = token.get("children", [])
                runs.extend(self._parse_inline(children, bold=bold, italic=True, code=code))
            elif t == "codespan":
                runs.append(TextRun(
                    text=token.get("raw", ""),
                    bold=bold, italic=italic, code=True,
                ))
            elif t == "softbreak" or t == "linebreak":
                runs.append(TextRun(text=" ", bold=bold, italic=italic, code=code))
            elif t == "image":
                alt = token.get("attrs", {}).get("alt", "")
                runs.append(TextRun(text=f"[图片: {alt}]"))
            elif t == "link":
                children = token.get("children", [])
                runs.extend(self._parse_inline(children, bold, italic, code))
            else:
                raw = token.get("raw", "")
                if raw:
                    runs.append(TextRun(text=raw, bold=bold, italic=italic, code=code))
        return runs
