from __future__ import annotations
import re
from typing import List, Optional, Tuple

import yaml
import mistune
from mistune import import_plugin

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, Table, HorizontalRule, Document,
    InlineElement, BlockElement,
)
from md2word.exceptions import ParseError


class _ASTRenderer:
    NAME = "ast"
    def __call__(self, tokens, state):
        return list(tokens)


def parse_frontmatter(text: str) -> Tuple[dict, str]:
    match = re.match(r'^---\s*\n(.*?)\n---\s*\n', text, re.DOTALL)
    if match:
        try:
            metadata = yaml.safe_load(match.group(1))
            rest = text[match.end():]
            return metadata or {}, rest
        except yaml.YAMLError:
            pass
    return {}, text


def _parse_img_attrs(attrs_str: str) -> dict:
    attrs = {}
    for part in attrs_str.split():
        if "=" in part:
            k, v = part.split("=", 1)
            attrs[k.strip()] = v.strip()
        else:
            attrs[part.strip()] = True
    return attrs


def preprocess_image_attributes(text: str) -> Tuple[str, dict]:
    """Transform ![alt](path){:attrs} into ![alt](path) and return attrs lookup.

    Returns (processed_text, {(src, alt): attrs_dict}).
    """
    attrs_map: dict = {}

    def _replacer(m):
        alt = m.group(1) or ""
        src = m.group(2) or ""
        attrs = _parse_img_attrs(m.group(3))
        key = (src, alt)
        attrs_map[key] = attrs
        return f"![{alt}]({src})"

    text = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)\{([^}]*)\}', _replacer, text)
    return text, attrs_map


class MarkdownParser:
    def __init__(self):
        self.md = mistune.Markdown(
            renderer=_ASTRenderer(),
            plugins=[import_plugin("table")],
        )
        self._img_attrs_map: dict = {}

    def parse(self, text: str) -> Document:
        metadata, body = parse_frontmatter(text)
        body, self._img_attrs_map = preprocess_image_attributes(body)
        try:
            ast = self.md(body)
        except Exception as e:
            raise ParseError(f"Failed to parse markdown: {e}") from e
        elements = self._parse_blocks(ast)
        return Document(metadata=metadata, elements=elements)

    # ---- block parsing ----

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
        elif t == "table":
            return self._parse_table(token)
        elif t == "thematic_break":
            return HorizontalRule()
        elif t == "blank_line":
            return None
        elif t == "block_quote":
            return self._parse_block_quote(token)
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
        lines: List[str] = []
        for child in children:
            ct = child.get("type", "")
            if ct in ("paragraph", "block_text"):
                runs = self._parse_inline(child.get("children", []))
                text = "".join(
                    r.text for r in runs if isinstance(r, TextRun)
                )
                lines.append(text)
        return Paragraph(runs=[TextRun(text="\n".join(lines), italic=True)])

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
            if ct in ("paragraph", "block_text"):
                elements.append(self._parse_paragraph(child))
            elif ct == "list":
                elements.append(self._parse_list(child))
            elif ct == "block_code":
                elements.append(self._parse_code_block(child))
        if not elements:
            return None
        return ListItem(elements=elements)

    def _parse_table(self, token: dict) -> Table:
        headers: List[str] = []
        rows: List[List[str]] = []
        align: List[Optional[str]] = []

        children = token.get("children", [])

        # Extract column alignment from head cells
        def _cell_text(cell_token: dict) -> str:
            parts: List[str] = []
            for c in cell_token.get("children", []):
                if c.get("type") == "text":
                    parts.append(c.get("raw", ""))
                elif c.get("type") == "codespan":
                    parts.append(c.get("raw", ""))
                elif c.get("type") in ("strong", "emphasis"):
                    for cc in c.get("children", []):
                        if cc.get("type") == "text":
                            parts.append(cc.get("raw", ""))
            return "".join(parts)

        for child in children:
            ct = child.get("type", "")
            if ct == "table_head":
                for cell in child.get("children", []):
                    if cell.get("type") == "table_cell":
                        cell_align = cell.get("attrs", {}).get("align")
                        align.append(cell_align)
                        headers.append(_cell_text(cell))
            elif ct == "table_body":
                for row in child.get("children", []):
                    if row.get("type") == "table_row":
                        row_cells: List[str] = []
                        for cell in row.get("children", []):
                            if cell.get("type") == "table_cell":
                                row_cells.append(_cell_text(cell))
                        if row_cells:
                            rows.append(row_cells)

        return Table(headers=headers, rows=rows, align=align)

    # ---- inline parsing ----

    def _parse_inline(
        self, tokens: list, bold: bool = False,
        italic: bool = False, code: bool = False,
    ) -> List[InlineElement]:
        runs: List[InlineElement] = []
        for token in tokens:
            t = token.get("type", "")
            if t == "text":
                runs.append(TextRun(
                    text=token.get("raw", ""),
                    bold=bold, italic=italic, code=code,
                ))
            elif t == "strong":
                children = token.get("children", [])
                runs.extend(self._parse_inline(children, True, italic, code))
            elif t == "emphasis":
                children = token.get("children", [])
                runs.extend(self._parse_inline(children, bold, True, code))
            elif t == "codespan":
                runs.append(TextRun(
                    text=token.get("raw", ""),
                    bold=bold, italic=italic, code=True,
                ))
            elif t in ("softbreak", "linebreak"):
                runs.append(TextRun(text=" ", bold=bold, italic=italic, code=code))
            elif t == "image":
                src = token.get("attrs", {}).get("url", "")
                alt = ""
                for c in token.get("children", []):
                    if c.get("type") == "text":
                        alt = c.get("raw", "")
                # Look up extra attributes from pre-processing
                extra = self._img_attrs_map.get((src, alt), {})
                img = Image(
                    src=src,
                    alt=alt,
                    width=extra.get("width"),
                    height=extra.get("height"),
                    align=extra.get("align"),
                )
                runs.append(img)
            elif t == "link":
                children = token.get("children", [])
                runs.extend(self._parse_inline(children, bold, italic, code))
            else:
                raw = token.get("raw", "")
                if raw:
                    runs.append(TextRun(
                        text=raw, bold=bold, italic=italic, code=code,
                    ))
        return runs
