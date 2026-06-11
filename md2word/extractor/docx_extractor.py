from __future__ import annotations
import os
import re
from typing import List, Optional, Set, Tuple
from xml.etree.ElementTree import Element

from docx import Document as DocxDocument
from docx.oxml.ns import qn
from lxml import etree

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


# XML namespaces (lxml format)
_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"


class DocxExtractor:
    def __init__(self, output_dir: str = ""):
        self.output_dir = output_dir
        self._img_counter: int = 0
        self._saved_images: Set[str] = set()

    def extract(self, docx_path: str) -> Document:
        doc = DocxDocument(docx_path)
        elements: List[BlockElement] = []
        list_buffer: List[ListItem] = []
        list_ordered: bool = False
        list_tight: bool = True

        for child in doc.element.body:
            tag = child.tag
            if tag.endswith("}p"):
                para = self._find_paragraph_in_doc(doc, child)
                if para is None:
                    continue
                element = self._extract_paragraph(para)

                # Check for images in raw XML
                images = self._extract_images_from_xml(child, doc, para)
                if images:
                    if isinstance(element, Paragraph):
                        combined = element.runs + images if element.runs else images
                        element = Paragraph(runs=combined, alignment=element.alignment)

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

                if element is not None:
                    elements.append(element)

            elif tag.endswith("}tbl"):
                table = self._extract_table_xml(child, doc)
                if table is not None:
                    if list_buffer:
                        elements.append(ListBlock(
                            ordered=list_ordered, items=list_buffer, tight=list_tight,
                        ))
                        list_buffer = []
                    elements.append(table)

        if list_buffer:
            elements.append(ListBlock(
                ordered=list_ordered, items=list_buffer, tight=list_tight,
            ))

        return Document(elements=elements)

    def _find_paragraph_in_doc(self, doc: DocxDocument, xml_elem) -> Optional:
        for para in doc.paragraphs:
            if para._p is xml_elem:
                return para
        return None

    def _get_paragraph_align(self, p_elem) -> Optional[str]:
        try:
            pPr = p_elem.find(f"{{{_NS_W}}}pPr")
            if pPr is not None:
                jc = pPr.find(f"{{{_NS_W}}}jc")
                if jc is not None:
                    val = jc.get(f"{{{_NS_W}}}val")
                    align_map = {"left": "left", "center": "center",
                                 "right": "right", "both": "justify"}
                    return align_map.get(val)
        except Exception:
            pass
        return None

    def _extract_images_from_xml(self, p_elem, doc: DocxDocument, para=None) -> List[Image]:
        images: List[Image] = []
        try:
            tree = etree.fromstring(etree.tostring(p_elem))
        except Exception:
            return images

        para_align = self._get_paragraph_align(tree)

        # Inline images
        for drawing in tree.iter(f"{{{_NS_WP}}}inline"):
            img = self._extract_single_image(drawing, doc)
            if img:
                images.append(img)

        # Anchored images (floating)
        for drawing in tree.iter(f"{{{_NS_WP14}}}anchor"):
            img = self._extract_single_image(drawing, doc)
            if img:
                images.append(img)

        return images

    def _extract_single_image(self, drawing, doc: DocxDocument) -> Optional[Image]:
        # Size
        extent = drawing.find(f"{{{_NS_WP}}}extent")
        cx = int(extent.get("cx")) if extent is not None and extent.get("cx") else None
        cy = int(extent.get("cy")) if extent is not None and extent.get("cy") else None

        # Alt text
        docPr = drawing.find(f"{{{_NS_WP}}}docPr")
        alt_text = ""
        if docPr is not None:
            alt_text = docPr.get("descr") or docPr.get("name") or ""

        # Image data
        blip = drawing.find(f".//{{{_NS_A}}}blip")
        if blip is None:
            return None
        embed = blip.get(f"{{{_NS_R}}}embed")
        if not embed or embed not in doc.part.related_parts:
            return None

        image_part = doc.part.related_parts[embed]
        img_name = self._save_image(image_part)
        rel_path = "images/" + img_name

        width_str = f"{cx / 914400:.2f}in" if cx else None
        height_str = f"{cy / 914400:.2f}in" if cy else None

        return Image(src=rel_path, alt=alt_text, width=width_str, height=height_str)

    def _save_image(self, image_part) -> str:
        ext = self._get_image_ext(image_part)
        self._img_counter += 1
        filename = f"image_{self._img_counter:03d}{ext}"
        self._saved_images.add(filename)
        if self.output_dir:
            filepath = os.path.join(self.output_dir, filename)
            with open(filepath, "wb") as f:
                f.write(image_part.blob)
        return filename

    @staticmethod
    def _get_image_ext(image_part) -> str:
        content_type = image_part.content_type or ""
        mapping = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tiff",
            "image/svg+xml": ".svg",
        }
        for ct, ext in mapping.items():
            if ct in content_type:
                return ext
        return ".png"

    def _extract_table_xml(self, tbl_elem, doc: DocxDocument) -> Optional[Table]:
        try:
            tree = etree.fromstring(etree.tostring(tbl_elem))
        except Exception:
            return None

        rows_xml = tree.findall(f".//{{{_NS_W}}}tr")
        if not rows_xml:
            return None

        all_rows: List[List[str]] = []
        aligns: List[Optional[str]] = []

        for row_idx, row_xml in enumerate(rows_xml):
            cells = row_xml.findall(f".//{{{_NS_W}}}tc")
            row_data: List[str] = []
            for cell_idx, cell in enumerate(cells):
                cell_text = ""
                texts: List[str] = []
                for p in cell.findall(f".//{{{_NS_W}}}p"):
                    images_in_p = self._extract_images_from_cell_xml(p, doc)
                    if images_in_p:
                        for img in images_in_p:
                            texts.append(self._format_cell_image(img))
                    for t in p.findall(f".//{{{_NS_W}}}t"):
                        if t.text:
                            texts.append(t.text)
                    texts.append("\n")
                cell_text = "".join(texts).strip()
                row_data.append(cell_text)

                if row_idx == 0:
                    p = cell.find(f".//{{{_NS_W}}}p")
                    align_val = None
                    if p is not None:
                        pPr = p.find(f"{{{_NS_W}}}pPr")
                        if pPr is not None:
                            jc = pPr.find(f"{{{_NS_W}}}jc")
                            if jc is not None:
                                val = jc.get(f"{{{_NS_W}}}val")
                                align_map = {"left": "left", "center": "center",
                                             "right": "right", "both": "justify"}
                                align_val = align_map.get(val)
                    while len(aligns) <= cell_idx:
                        aligns.append(None)
                    aligns[cell_idx] = align_val

            all_rows.append(row_data)

        if not all_rows:
            return None

        num_cols = max(len(r) for r in all_rows)
        headers = all_rows[0] if all_rows else []
        body_rows = all_rows[1:] if len(all_rows) > 1 else []

        aligns = aligns[:num_cols] if aligns else []
        while len(aligns) < num_cols:
            aligns.append(None)

        return Table(headers=headers, rows=body_rows, align=aligns)

    def _extract_images_from_cell_xml(self, p_elem, doc: DocxDocument) -> List[Image]:
        images: List[Image] = []
        for drawing in p_elem.iter(f"{{{_NS_WP}}}inline"):
            img = self._extract_single_image(drawing, doc)
            if img:
                images.append(img)
        for drawing in p_elem.iter(f"{{{_NS_WP14}}}anchor"):
            img = self._extract_single_image(drawing, doc)
            if img:
                images.append(img)
        return images

    def _format_cell_image(self, img: Image) -> str:
        parts = []
        if img.width:
            parts.append(f"width={img.width}")
        if img.height:
            parts.append(f"height={img.height}")
        attrs = "{" + " ".join(parts) + "}" if parts else ""
        return f"![{img.alt}]({img.src}){attrs}"

    def _extract_paragraph(self, para) -> Optional[BlockElement]:
        style_name = para.style.name if para.style else ""

        if style_name.startswith("Heading"):
            try:
                level = int(style_name.split()[-1])
            except (ValueError, IndexError):
                level = 1
            runs = self._extract_runs(para)
            return Heading(level=level, runs=runs)

        if style_name in ("List Bullet", "List Number") or \
           style_name.startswith("List Bullet ") or \
           style_name.startswith("List Number "):
            ordered = "Number" in style_name
            item = ListItem(elements=[Paragraph(runs=self._extract_runs(para))])
            tight = not para.paragraph_format.space_before and not para.paragraph_format.space_after
            return ListBlock(ordered=ordered, items=[item], tight=tight)

        numPr = para._p.find(qn("w:pPr"))
        if numPr is not None:
            numPr = numPr.find(qn("w:numPr"))
        if numPr is not None:
            item = ListItem(elements=[Paragraph(runs=self._extract_runs(para))])
            return ListBlock(ordered=True, items=[item], tight=True)

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

    @staticmethod
    def _is_horizontal_rule(para) -> bool:
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
