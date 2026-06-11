from __future__ import annotations
import re
from typing import List, Optional, Tuple

import yaml
import mistune
from mistune import import_plugin

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, Table, HorizontalRule, Formula,
    Document, InlineElement, BlockElement,
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
        key = part.lstrip(":")
        if "=" in key:
            k, v = key.split("=", 1)
            attrs[k.strip()] = v.strip()
        else:
            attrs[key.strip()] = True
    return attrs


def preprocess_image_attributes(text: str) -> Tuple[str, dict]:
    attrs_map: dict = {}
    counter = 0

    def _replacer(m):
        nonlocal counter
        alt = m.group(1) or ""
        src = m.group(2) or ""
        raw_attrs = m.group(3) or ""
        attrs = _parse_img_attrs(raw_attrs) if raw_attrs else {}
        attrs_map[counter] = attrs
        counter += 1
        return f"![{alt}]({src})"

    text = re.sub(
        r'!\[([^\]]*)\]\(([^)]+)\)(?:{([^}]*)})?',
        _replacer, text,
    )
    return text, attrs_map


def _parse_span_style(raw: str) -> dict:
    props = {}
    m = re.search(r'font-family\s*:\s*([^;"]+)', raw, re.IGNORECASE)
    if m:
        props["font_name"] = m.group(1).strip().strip("'\"")
    m = re.search(r'font-size\s*:\s*(\d+(\.\d+)?)\s*pt', raw, re.IGNORECASE)
    if m:
        props["font_size"] = int(float(m.group(1)))
    m = re.search(r'font-size\s*:\s*(\d+(\.\d+)?)\s*px', raw, re.IGNORECASE)
    if m:
        props["font_size"] = int(float(m.group(1)) * 72 / 96)
    return props


_FORMULA_PLACEHOLDER_RE = re.compile(
    r'\x00FORMULA_(INLINE|BLOCK|EQ)_(\d+)\x00'
)


def preprocess_formulas(text: str) -> Tuple[str, dict]:
    formula_map: dict = {}
    counter = [0]

    def _replace_block(m):
        idx = counter[0]
        counter[0] += 1
        latex = m.group(1).strip()
        formula_map[f"BLOCK_{idx}"] = {"latex": latex, "display": True}
        return f"\x00FORMULA_BLOCK_{idx}\x00"

    def _replace_eq(m):
        idx = counter[0]
        counter[0] += 1
        latex = m.group(1).strip()
        formula_map[f"EQ_{idx}"] = {"latex": latex, "display": True, "numbering": ""}
        return f"\x00FORMULA_EQ_{idx}\x00"

    def _replace_inline(m):
        idx = counter[0]
        counter[0] += 1
        latex = m.group(1).strip()
        formula_map[f"INLINE_{idx}"] = {"latex": latex, "display": False}
        return f"\x00FORMULA_INLINE_{idx}\x00"

    def _clean_latex(content: str) -> str:
        content = re.sub(r'\\label\{[^}]*\}', '', content)
        return content.strip()

    # Block formulas: $$...$$
    text = re.sub(r'\$\$(.*?)\$\$', lambda m: _replace_block(
        type("m", (), {"group": lambda self, i: m.group(i)})()
    ), text, flags=re.DOTALL)

    # Align environments: \begin{align}...\end{align} / \begin{align*}...\end{align*}
    def _replace_align(m):
        idx = counter[0]
        counter[0] += 1
        latex = _clean_latex(m.group(1))
        formula_map[f"BLOCK_{idx}"] = {"latex": latex, "display": True}
        return f"\n\n\x00FORMULA_BLOCK_{idx}\x00\n\n"
    text = re.sub(
        r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        _replace_align, text, flags=re.DOTALL,
    )

    # Numbered equations: \begin{equation}...\end{equation} / \begin{equation*}...\end{equation*}
    def _replace_eq(m):
        idx = counter[0]
        counter[0] += 1
        latex = _clean_latex(m.group(1))
        formula_map[f"EQ_{idx}"] = {"latex": latex, "display": True, "numbering": ""}
        return f"\n\n\x00FORMULA_EQ_{idx}\x00\n\n"
    text = re.sub(
        r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        _replace_eq, text, flags=re.DOTALL,
    )
    # Inline formulas: $...$ (must be after $$)
    text = re.sub(r'\$(.+?)\$', _replace_inline, text)

    return text, formula_map


class MarkdownParser:
    def __init__(self):
        self.md = mistune.Markdown(
            renderer=_ASTRenderer(),
            plugins=[
                import_plugin("table"),
                import_plugin("strikethrough"),
            ],
        )
        self._img_attrs_map: dict = {}
        self._img_index: int = 0
        self._formula_map: dict = {}

    def parse(self, text: str) -> Document:
        metadata, body = parse_frontmatter(text)
        body, self._img_attrs_map = preprocess_image_attributes(body)
        body, self._formula_map = preprocess_formulas(body)
        self._img_index = 0
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

    def _parse_paragraph(self, token: dict) -> BlockElement:
        runs = self._parse_inline(token.get("children", []))
        if len(runs) == 1 and isinstance(runs[0], Image):
            img = runs[0]
            return Image(
                src=img.src, alt=img.alt,
                width=img.width, height=img.height, align=img.align,
            )
        # Check for standalone block formula
        if len(runs) == 1 and isinstance(runs[0], TextRun):
            m = _FORMULA_PLACEHOLDER_RE.match(runs[0].text)
            if m:
                ftype, fidx = m.group(1), m.group(2)
                key = f"{ftype}_{fidx}"
                info = self._formula_map.get(key)
                if info:
                    return Formula(
                        latex=info["latex"],
                        display=info["display"],
                        numbering=info.get("numbering"),
                    )
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

        for child in token.get("children", []):
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
        self, tokens: list,
        bold: bool = False, italic: bool = False, code: bool = False,
        underline: bool = False, strikethrough: bool = False,
        superscript: bool = False, subscript: bool = False,
        font_name: Optional[str] = None, font_size: Optional[int] = None,
    ) -> List[InlineElement]:
        runs: List[InlineElement] = []
        for token in tokens:
            t = token.get("type", "")
            if t == "text":
                raw = token.get("raw", "")
                # Check for inline formula placeholders in text
                if _FORMULA_PLACEHOLDER_RE.search(raw):
                    parts = _FORMULA_PLACEHOLDER_RE.split(raw)
                    # parts alternates: [text, type_idx, text, type_idx, ...]
                    placeholders = list(_FORMULA_PLACEHOLDER_RE.finditer(raw))
                    pi = 0
                    for i, part in enumerate(parts):
                        if pi < len(placeholders) and _FORMULA_PLACEHOLDER_RE.fullmatch(placeholders[pi].group()):
                            # Reconstruct the full match to check alignment
                            pass
                    # Simpler approach: iterate through parts and placeholders
                    text_parts = [p for i, p in enumerate(parts) if i % 3 == 0]
                    type_parts = [p for i, p in enumerate(parts) if i % 3 == 1]
                    idx_parts = [p for i, p in enumerate(parts) if i % 3 == 2]
                    for i, txt in enumerate(text_parts):
                        if txt:
                            runs.append(TextRun(
                                text=txt,
                                bold=bold, italic=italic, code=code,
                                underline=underline, strikethrough=strikethrough,
                                superscript=superscript, subscript=subscript,
                                font_name=font_name, font_size=font_size,
                            ))
                        if i < len(type_parts):
                            key = f"{type_parts[i]}_{idx_parts[i]}"
                            info = self._formula_map.get(key)
                            if info:
                                runs.append(Formula(
                                    latex=info["latex"],
                                    display=info["display"],
                                ))
                else:
                    runs.append(TextRun(
                        text=raw,
                        bold=bold, italic=italic, code=code,
                        underline=underline, strikethrough=strikethrough,
                        superscript=superscript, subscript=subscript,
                        font_name=font_name, font_size=font_size,
                    ))
            elif t == "strong":
                children = token.get("children", [])
                runs.extend(self._parse_inline(
                    children, True, italic, code,
                    underline, strikethrough, superscript, subscript,
                    font_name, font_size,
                ))
            elif t == "emphasis":
                children = token.get("children", [])
                runs.extend(self._parse_inline(
                    children, bold, True, code,
                    underline, strikethrough, superscript, subscript,
                    font_name, font_size,
                ))
            elif t == "codespan":
                runs.append(TextRun(
                    text=token.get("raw", ""),
                    bold=bold, italic=italic, code=True,
                    font_name=font_name, font_size=font_size,
                ))
            elif t == "strikethrough":
                children = token.get("children", [])
                runs.extend(self._parse_inline(
                    children, bold, italic, code,
                    underline, True, superscript, subscript,
                    font_name, font_size,
                ))
            elif t == "inline_html":
                raw = token.get("raw", "")
                if raw == "<u>" or raw == "<ins>":
                    underline = True
                elif raw == "</u>" or raw == "</ins>":
                    underline = False
                elif raw in ("<s>", "<strike>", "<del>"):
                    strikethrough = True
                elif raw in ("</s>", "</strike>", "</del>"):
                    strikethrough = False
                elif raw == "<sup>":
                    superscript = True
                elif raw == "</sup>":
                    superscript = False
                elif raw == "<sub>":
                    subscript = True
                elif raw == "</sub>":
                    subscript = False
                elif raw.startswith("<span"):
                    span_props = _parse_span_style(raw)
                    font_name = span_props.get("font_name", font_name)
                    font_size = span_props.get("font_size", font_size)
                elif raw == "</span>":
                    font_name = None
                    font_size = None
            elif t in ("softbreak", "linebreak"):
                runs.append(TextRun(
                    text=" ",
                    bold=bold, italic=italic, code=code,
                    underline=underline, strikethrough=strikethrough,
                ))
            elif t == "image":
                src = token.get("attrs", {}).get("url", "")
                alt = ""
                for c in token.get("children", []):
                    if c.get("type") == "text":
                        alt = c.get("raw", "")
                extra = self._img_attrs_map.get(self._img_index, {})
                self._img_index += 1
                img = Image(
                    src=src, alt=alt,
                    width=extra.get("width"),
                    height=extra.get("height"),
                    align=extra.get("align"),
                )
                runs.append(img)
            elif t == "link":
                children = token.get("children", [])
                runs.extend(self._parse_inline(
                    children, bold, italic, code,
                    underline, strikethrough, superscript, subscript,
                    font_name, font_size,
                ))
            else:
                raw = token.get("raw", "")
                if raw:
                    runs.append(TextRun(text=raw, bold=bold, italic=italic, code=code))
        return runs
