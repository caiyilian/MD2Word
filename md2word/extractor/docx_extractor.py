from __future__ import annotations
import os
from typing import List, Optional, Set

from docx import Document as DocxDocument
from docx.oxml.ns import qn

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, Table, HorizontalRule, Formula,
    Document, InlineElement, BlockElement,
)


_MONO_FONTS: Set[str] = {
    "consolas", "courier new", "courier", "monospace",
    "source code pro", "fira code", "menlo", "monaco",
    "droid sans mono", "dejavu sans mono", "liberation mono",
}


def _is_mono_font(name: Optional[str]) -> bool:
    return name is not None and name.lower().strip() in _MONO_FONTS


def _alignment_to_str(align) -> Optional[str]:
    if align is None:
        return None
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    if align == WD_ALIGN_PARAGRAPH.LEFT:
        return "left"
    if align == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    if align == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    if align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "justify"
    return None


class DocxExtractor:
    def extract(self, docx_path: str) -> Document:
        doc = DocxDocument(docx_path)
        elements: List[BlockElement] = []
        list_buffer: List[ListItem] = []
        list_ordered: bool = False
        list_tight: bool = True

        for para in doc.paragraphs:
            style_name = para.style.name if para.style else ""
            element = self._extract_paragraph(para, style_name)

            if isinstance(element, ListBlock):
                items = element.items
                if items:
                    if list_buffer and list_ordered == element.ordered:
                        list_buffer.extend(items)
                    else:
                        if list_buffer:
                            elements.append(ListBlock(
                                ordered=list_ordered, items=list_buffer,
                                tight=list_tight,
                            ))
                        list_buffer = list(items)
                        list_ordered = element.ordered
                        list_tight = element.tight
                continue

            is_hr = self._is_horizontal_rule(para)
            if is_hr:
                element = HorizontalRule()

            if list_buffer:
                elements.append(ListBlock(
                    ordered=list_ordered, items=list_buffer, tight=list_tight,
                ))
                list_buffer = []
                list_tight = True

            if element is not None:
                elements.append(element)

        if list_buffer:
            elements.append(ListBlock(
                ordered=list_ordered, items=list_buffer, tight=list_tight,
            ))

        return Document(elements=elements)

    def _extract_paragraph(self, para, style_name: str) -> Optional[BlockElement]:
        # Heading
        if style_name.startswith("Heading"):
            try:
                level = int(style_name.split()[-1])
            except (ValueError, IndexError):
                level = 1
            runs = self._extract_runs(para)
            return Heading(level=level, runs=runs)

        # List
        if style_name in ("List Bullet", "List Number") or style_name.startswith("List Bullet ") or style_name.startswith("List Number "):
            ordered = "Number" in style_name
            item = ListItem(elements=[Paragraph(runs=self._extract_runs(para))])
            tight = not para.paragraph_format.space_before and not para.paragraph_format.space_after
            return ListBlock(ordered=ordered, items=[item], tight=tight)

        # Check for numbering via XML
        numPr = para._p.find(qn("w:pPr"))
        if numPr is not None:
            numPr = numPr.find(qn("w:numPr"))
        if numPr is not None:
            ordered = True
            item = ListItem(elements=[Paragraph(runs=self._extract_runs(para))])
            return ListBlock(ordered=ordered, items=[item], tight=True)

        # Regular paragraph
        runs = self._extract_runs(para)
        alignment = _alignment_to_str(para.alignment)
        return Paragraph(runs=runs, alignment=alignment)

    def _extract_runs(self, para) -> List[InlineElement]:
        runs: List[InlineElement] = []
        for run in para.runs:
            text = run.text
            if not text:
                continue
            font = run.font
            font_name = font.name
            font_size = font.size
            code = _is_mono_font(font_name)

            runs.append(TextRun(
                text=text,
                bold=run.bold or False,
                italic=run.italic or False,
                code=code,
                underline=run.underline or False,
                strikethrough=font.strike or False,
                superscript=font.superscript or False,
                subscript=font.subscript or False,
                font_name=font_name,
                font_size=int(font_size.pt) if font_size else None,
            ))
        return runs

    def _is_horizontal_rule(self, para) -> bool:
        try:
            pPr = para._p.find(qn("w:pPr"))
            if pPr is not None:
                pBdr = pPr.find(qn("w:pBdr"))
                if pBdr is not None:
                    bottom = pBdr.find(qn("w:bottom"))
                    if bottom is not None and bottom.get(qn("w:val")) == "single":
                        return True
        except Exception:
            pass
        return False
