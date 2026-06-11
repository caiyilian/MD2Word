from __future__ import annotations
import os
from typing import List, Optional

from docx import Document as DocxDocument
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock,
    ListBlock, ListItem, Table, HorizontalRule, Document,
    InlineElement, BlockElement,
)
from md2word.utils.unit_converter import parse_size


class DocxRenderer:
    def __init__(self, font_name: str = "等线", font_size: int = 12, base_dir: str = ""):
        self.font_name = font_name
        self.font_size = font_size
        self.base_dir = base_dir

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
        elif isinstance(element, Table):
            self._render_table(doc, element)
        elif isinstance(element, Image):
            self._render_image(doc, element)
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

        pPr = p._p.get_or_add_pPr()
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "F2F2F2")
        shading.set(qn("w:val"), "clear")
        pPr.append(shading)

        run = p.add_run(code_block.code)
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    def _render_image(self, doc: DocxDocument, image: Image):
        p = doc.add_paragraph()
        self._apply_image_alignment(p, image.align)
        if not self._add_image_to_paragraph(p, image):
            run = p.add_run(f"[图片未找到: {image.src}]")
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            run.font.size = Pt(9)

    def _render_table(self, doc: DocxDocument, table: Table):
        if not table.headers and not table.rows:
            return

        num_cols = max(len(table.headers), max((len(r) for r in table.rows), default=0))
        if num_cols == 0:
            return

        t = doc.add_table(rows=1 + len(table.rows), cols=num_cols)
        t.style = "Table Grid"
        t.autofit = True

        # Headers
        if table.headers:
            for i, header in enumerate(table.headers):
                if i >= num_cols:
                    break
                cell = t.cell(0, i)
                cell.text = ""
                run = cell.paragraphs[0].add_run(header)
                run.bold = True
                run.font.name = self.font_name
                run.font.size = Pt(self.font_size)
                # Header shading
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), "E8E8E8")
                shading.set(qn("w:val"), "clear")
                cell._tc.get_or_add_tcPr().append(shading)

        # Rows
        for row_idx, row_data in enumerate(table.rows):
            for col_idx, cell_text in enumerate(row_data):
                if col_idx >= num_cols:
                    break
                cell = t.cell(1 + row_idx, col_idx)
                cell.text = ""
                run = cell.paragraphs[0].add_run(cell_text)
                run.font.name = self.font_name
                run.font.size = Pt(self.font_size)

        # Column alignment
        for i, align_val in enumerate(table.align):
            if i >= num_cols:
                break
            if align_val == "left":
                alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif align_val == "center":
                alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align_val == "right":
                alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                continue
            # Set alignment for header
            if table.headers:
                t.cell(0, i).paragraphs[0].alignment = alignment
            # Set alignment for body cells
            for row_idx in range(len(table.rows)):
                t.cell(1 + row_idx, i).paragraphs[0].alignment = alignment

        doc.add_paragraph()  # spacing after table

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

    # ---- helpers ----

    def _apply_runs(self, paragraph, runs: List[InlineElement]):
        for run_data in runs:
            if isinstance(run_data, Image):
                self._add_image_to_paragraph(paragraph, run_data)
            elif isinstance(run_data, TextRun):
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

    def _add_image_to_paragraph(self, paragraph, image: Image):
        img_path = self._resolve_image_path(image.src)
        if img_path is None:
            return False

        try:
            page_width = paragraph.part.document.element.body.sectPr.xpath(
                "./w:pgSz/@w:w"
            )
            if page_width:
                pw = Emu(int(page_width[0]))
            else:
                pw = None
        except Exception:
            pw = None

        width = parse_size(image.width, pw)
        height = parse_size(image.height)

        run = paragraph.add_run()
        run.add_picture(img_path, width=width, height=height)
        return True

    def _resolve_image_path(self, src: str) -> Optional[str]:
        if os.path.isabs(src):
            return src if os.path.exists(src) else None
        if self.base_dir:
            resolved = os.path.join(self.base_dir, src)
            if os.path.exists(resolved):
                return resolved
        return src if os.path.exists(src) else None

    def _apply_image_alignment(self, paragraph, align: Optional[str]):
        if align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == "left":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
