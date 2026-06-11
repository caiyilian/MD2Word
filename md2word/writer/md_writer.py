from __future__ import annotations
from typing import List

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, Table, HorizontalRule, Formula,
    Document, InlineElement, BlockElement,
)


class MdWriter:
    def __init__(self, default_font_name: str = "\u7b49\u7ebf",
                 default_font_size: int = 12):
        self.default_font_name = default_font_name
        self.default_font_size = default_font_size

    def write(self, document: Document) -> str:
        lines: List[str] = []
        for element in document.elements:
            text = self._write_block(element)
            if text:
                lines.append(text)
        return "\n".join(lines) + "\n"

    def _write_block(self, element: BlockElement) -> str:
        if isinstance(element, Heading):
            return self._write_heading(element)
        if isinstance(element, Paragraph):
            return self._write_paragraph(element)
        if isinstance(element, CodeBlock):
            return self._write_code_block(element)
        if isinstance(element, ListBlock):
            return self._write_list(element)
        if isinstance(element, Table):
            return self._write_table(element)
        if isinstance(element, Image):
            return self._write_image(element)
        if isinstance(element, HorizontalRule):
            return "---"
        if isinstance(element, Formula):
            return self._write_formula(element)
        return ""

    def _write_heading(self, heading: Heading) -> str:
        prefix = "#" * heading.level
        text = self._write_inline(heading.runs)
        return f"{prefix} {text}\n"

    def _write_paragraph(self, paragraph: Paragraph) -> str:
        text = self._write_inline(paragraph.runs)
        return f"{text}\n"

    def _write_code_block(self, code_block: CodeBlock) -> str:
        lang = code_block.language or ""
        return f"```{lang}\n{code_block.code}\n```\n"

    def _write_list(self, list_block: ListBlock) -> str:
        lines: List[str] = []
        for idx, item in enumerate(list_block.items, start=1):
            lines.extend(self._write_list_item(item, list_block.ordered, idx))
        return "\n".join(lines) + "\n" if lines else ""

    def _write_list_item(self, item: ListItem, ordered: bool, idx: int, depth: int = 0) -> List[str]:
        lines: List[str] = []
        indent = "  " * depth
        prefix = f"{indent}{idx}. " if ordered else f"{indent}- "
        for element in item.elements:
            if isinstance(element, Paragraph):
                text = self._write_inline(element.runs)
                if text:
                    lines.append(f"{prefix}{text}")
                else:
                    lines.append(prefix.rstrip())
            elif isinstance(element, CodeBlock):
                for line in element.code.split("\n"):
                    lines.append(f"  {indent}  {line}")
            elif isinstance(element, ListBlock):
                lines.extend(self._write_list(element))
        return lines

    def _write_table(self, table: Table) -> str:
        if not table.headers and not table.rows:
            return ""
        num_cols = max(len(table.headers), max((len(r) for r in table.rows), default=0))
        if num_cols == 0:
            return ""

        def _escape_cell(text: str) -> str:
            return text.replace("|", "\\|").replace("\n", "<br>")

        lines: List[str] = []
        if table.headers:
            row = "| " + " | ".join(_escape_cell(h) for h in table.headers[:num_cols]) + " |"
            lines.append(row)
            sep_parts: List[str] = []
            for i in range(num_cols):
                a = table.align[i] if i < len(table.align) else None
                if a == "center":
                    sep_parts.append(":---:")
                elif a == "right":
                    sep_parts.append("---:")
                elif a == "left":
                    sep_parts.append(":---")
                else:
                    sep_parts.append("-----")
            lines.append("| " + " | ".join(sep_parts) + " |")
        for row_data in table.rows:
            row = "| " + " | ".join(_escape_cell(c) for c in row_data[:num_cols]) + " |"
            lines.append(row)
        return "\n".join(lines) + "\n"

    def _write_image(self, image: Image) -> str:
        alt = image.alt or ""
        src = image.src or ""
        attrs = ""
        if image.width or image.height or image.align:
            parts: List[str] = []
            if image.width:
                parts.append(f"width={image.width}")
            if image.height:
                parts.append(f"height={image.height}")
            if image.align:
                parts.append(f"align={image.align}")
            if parts:
                attrs = "{" + " ".join(parts) + "}"
        return f"![{alt}]({src}){attrs}\n"

    def _write_formula(self, formula: Formula) -> str:
        latex = formula.latex
        if formula.display:
            return f"$$\n{latex}\n$$\n"
        else:
            return f"${latex}$\n"

    def _write_inline(self, runs: List[InlineElement]) -> str:
        parts: List[str] = []
        for run in runs:
            if isinstance(run, Image):
                parts.append(self._write_image(run).strip())
            elif isinstance(run, TextRun):
                parts.append(self._format_text_run(run))
            elif isinstance(run, Formula):
                parts.append(self._write_formula(run).strip())
        return "".join(parts)

    def _format_text_run(self, run: TextRun) -> str:
        text = run.text
        if run.code:
            return f"`{text}`"

        if run.strikethrough:
            text = f"~~{text}~~"
        if run.superscript:
            text = f"<sup>{text}</sup>"
        if run.subscript:
            text = f"<sub>{text}</sub>"
        if run.underline:
            text = f"<u>{text}</u>"
        if run.italic:
            text = f"*{text}*"
        if run.bold:
            text = f"**{text}**"

        font_name = run.font_name
        font_size = run.font_size
        if font_name and font_name.lower() == self.default_font_name.lower():
            font_name = None
        if font_size == self.default_font_size:
            font_size = None
        if font_name or font_size:
            style = ""
            if font_name:
                style += f"font-family:{font_name};"
            if font_size:
                style += f"font-size:{font_size}pt;"
            if style:
                text = f'<span style="{style}">{text}</span>'

        return text
