from __future__ import annotations
from typing import List

from docx import Document as DocxDocument
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from md2word.model.document import (
    TextRun, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, HorizontalRule, Document,
    BlockElement,
)


class DocxRenderer:
    def __init__(self, font_name: str = "等线", font_size: int = 12):
        self.font_name = font_name
        self.font_size = font_size

    def render(self, document: Document, output_path: str):
        doc = DocxDocument()
        self._set_default_style(doc)

        for element in document.elements:
            self._render_element(doc, element)

        doc.save(output_path)
        return output_path

    def _set_default_style(self, doc: DocxDocument):
        style = doc.styles["Normal"]
        font = style.font
        font.name = self.font_name
        font.size = Pt(self.font_size)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.15

    def _render_element(self, doc: DocxDocument, element: BlockElement):
        if isinstance(element, Heading):
            self._render_heading(doc, element)
        elif isinstance(element, Paragraph):
            self._render_paragraph(doc, element)
        elif isinstance(element, CodeBlock):
            self._render_code_block(doc, element)
        elif isinstance(element, ListBlock):
            self._render_list(doc, element, indent_level=0)
        elif isinstance(element, HorizontalRule):
            self._render_horizontal_rule(doc)

    def _render_heading(self, doc: DocxDocument, heading: Heading):
        style_name = f"Heading {heading.level}"
        p = doc.add_paragraph(style=style_name)
        self._apply_runs(p, heading.runs)

    def _render_paragraph(self, doc: DocxDocument, paragraph: Paragraph):
        p = doc.add_paragraph()
        self._apply_runs(p, paragraph.runs)

    def _render_code_block(self, doc: DocxDocument, code_block: CodeBlock):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Inches(0.3)

        # Add background shading
        pPr = p._p.get_or_add_pPr()
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "F2F2F2")
        shading.set(qn("w:val"), "clear")
        pPr.append(shading)

        run = p.add_run(code_block.code)
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    def _render_list(self, doc: DocxDocument, list_block: ListBlock,
                     indent_level: int = 0):
        for idx, item in enumerate(list_block.items, start=1):
            self._render_list_item(doc, item, list_block.ordered, idx, indent_level)

    def _render_list_item(self, doc: DocxDocument, item: ListItem,
                          ordered: bool, idx: int, indent_level: int):
        for element in item.elements:
            if isinstance(element, Paragraph):
                p = doc.add_paragraph()
                indent = Inches(0.5 + indent_level * 0.4)
                p.paragraph_format.left_indent = indent
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)

                prefix = f"{idx}. " if ordered else "\u2022 "
                prefix_run = p.add_run(prefix)
                prefix_run.font.name = self.font_name
                prefix_run.font.size = Pt(self.font_size)

                self._apply_runs(p, element.runs)

            elif isinstance(element, CodeBlock):
                self._render_code_block(doc, element)
            elif isinstance(element, ListBlock):
                self._render_list(doc, element, indent_level + 1)

    def _render_horizontal_rule(self, doc: DocxDocument):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "999999")
        pBdr.append(bottom)
        pPr.append(pBdr)

    def _apply_runs(self, paragraph, runs: List[TextRun]):
        for run_data in runs:
            if not run_data.text:
                continue
            r = paragraph.add_run(run_data.text)
            r.bold = run_data.bold
            r.italic = run_data.italic

            if run_data.code:
                r.font.name = "Consolas"
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            else:
                r.font.name = self.font_name
                r.font.size = Pt(self.font_size)
