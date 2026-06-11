from __future__ import annotations
import os
from typing import Dict, List, Optional, Set

from docx import Document as DocxDocument
from docx.oxml.ns import qn
from lxml import etree

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock, Hyperlink,
    ListBlock, ListItem, Table, HorizontalRule, Formula, PageBreak,
    Footnote, Comment,
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


# XML namespaces
_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
_NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _qname(tag: str, ns: str = _NS_W) -> str:
    return f"{{{ns}}}{tag}"


class DocxExtractor:
    def __init__(self, output_dir: str = "", ocr: bool = False, max_ocr_images: int = 0):
        self.output_dir = output_dir
        self._img_counter: int = 0
        self._saved_images: Set[str] = set()
        self._ocr = ocr
        self._ocr_engine = None
        self._ocr_count = 0
        self._max_ocr_images = max_ocr_images
        if ocr:
            try:
                from paddleocr import PaddleOCR
                import warnings
                warnings.filterwarnings("ignore")
                self._ocr_engine = PaddleOCR(use_textline_orientation=True, lang="ch")
            except ImportError:
                pass

    def extract(self, docx_path: str) -> Document:
        doc = DocxDocument(docx_path)
        elements: List[BlockElement] = []
        list_buffer: List[ListItem] = []
        list_ordered: bool = False
        list_tight: bool = True
        code_buffer: List[str] = []

        # Extract metadata
        metadata = self._extract_metadata(doc)

        # Extract headers and footers
        headers = self._extract_headers(doc)
        footers = self._extract_footers(doc)

        # Extract section properties
        sections = self._extract_sections(doc)

        footnotes = self._load_footnotes(doc)
        comments = self._load_comments(doc)

        for child in doc.element.body:
            tag = child.tag

            if tag.endswith("}p"):
                para = self._find_paragraph_in_doc(doc, child)
                if para is None:
                    # Still extract images from orphan paragraphs
                    _, orphan_images = self._extract_runs_and_images(child, doc)
                    if orphan_images:
                        self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
                        list_buffer = []
                        list_tight = True
                        code_buffer = []
                        for img in orphan_images:
                            elements.append(img)
                    continue

                # Page break
                if self._has_page_break(child):
                    self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
                    elements.append(PageBreak())
                    list_buffer = []
                    list_tight = True
                    code_buffer = []
                    continue

                # Horizontal rule
                if self._is_horizontal_rule(para):
                    self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
                    elements.append(HorizontalRule())
                    list_buffer = []
                    list_tight = True
                    code_buffer = []
                    continue

                # Extract runs (text, hyperlinks, images) from XML
                runs, para_images = self._extract_runs_and_images(child, doc)

                # Check for standalone display formula (centered paragraph with only formula)
                text_runs = [r for r in runs if isinstance(r, TextRun)]
                formula_runs = [r for r in runs if isinstance(r, Formula)]
                alignment = _alignment_to_str(para.alignment)
                if not text_runs and formula_runs and alignment == "center" and len(formula_runs) == 1:
                    self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
                    list_buffer = []
                    list_tight = True
                    code_buffer = []
                    formula_runs[0].display = True
                    elements.append(formula_runs[0])
                    elements.extend(self._get_para_footnotes(child, footnotes))
                    elements.extend(self._get_para_comments(child, comments))
                    continue

                # Check monospace → code block
                flat_text = "".join(r.text for r in runs if isinstance(r, TextRun))
                is_code = runs and all(isinstance(r, TextRun) and r.code for r in runs) and flat_text.strip()
                if is_code:
                    code_buffer.append(flat_text)
                    continue
                elif code_buffer:
                    self._flush_code_buffer(elements, code_buffer)
                    code_buffer = []

                # Merge images into runs if no runs
                if not runs and para_images:
                    runs = para_images

                element = self._extract_paragraph_element(para, runs)

                if isinstance(element, ListBlock):
                    items = element.items
                    if items:
                        if list_buffer and list_ordered == element.ordered:
                            list_buffer.extend(items)
                        else:
                            self._flush_list_buffer(elements, list_buffer, list_ordered, list_tight)
                            list_buffer = list(items)
                            list_ordered = element.ordered
                            list_tight = element.tight
                    continue

                self._flush_list_buffer(elements, list_buffer, list_ordered, list_tight)
                list_buffer = []
                list_tight = True

                if element is not None:
                    elements.append(element)

                # Footnotes & comments
                elements.extend(self._get_para_footnotes(child, footnotes))
                elements.extend(self._get_para_comments(child, comments))

            elif tag.endswith("}tbl"):
                self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
                list_buffer = []
                code_buffer = []
                table_elem = self._extract_table_xml(child, doc)
                if table_elem is not None:
                    elements.append(table_elem)

        self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
        return Document(metadata=metadata, elements=elements,
                       headers=headers, footers=footers, sections=sections)

    def _flush(self, elements, code_buffer, list_buffer, list_ordered, list_tight):
        self._flush_code_buffer(elements, code_buffer)
        self._flush_list_buffer(elements, list_buffer, list_ordered, list_tight)

    def _flush_code_buffer(self, elements, code_buffer):
        if code_buffer:
            elements.append(CodeBlock(code="\n".join(code_buffer)))

    def _flush_list_buffer(self, elements, buffer, ordered, tight):
        if buffer:
            elements.append(ListBlock(ordered=ordered, items=buffer, tight=tight))

    def _find_paragraph_in_doc(self, doc: DocxDocument, xml_elem):
        for para in doc.paragraphs:
            if para._p is xml_elem:
                return para
        return None

    def _has_page_break(self, p_elem) -> bool:
        try:
            for br in p_elem.iter(_qname("br")):
                if br.get(_qname("type")) == "page":
                    return True
        except Exception:
            pass
        return False

    def _extract_paragraph_element(self, para, runs) -> Optional[BlockElement]:
        style_name = para.style.name if para.style else ""

        if style_name.startswith("Heading"):
            try:
                level = int(style_name.split()[-1])
            except (ValueError, IndexError):
                level = 1
            return Heading(level=level, runs=runs)

        if style_name in ("List Bullet", "List Number") or \
           style_name.startswith("List Bullet ") or \
           style_name.startswith("List Number "):
            ordered = "Number" in style_name
            item = ListItem(elements=[Paragraph(runs=runs)])
            tight = not para.paragraph_format.space_before \
                    and not para.paragraph_format.space_after
            return ListBlock(ordered=ordered, items=[item], tight=tight)

        numPr = para._p.find(qn("w:pPr"))
        if numPr is not None:
            numPr = numPr.find(qn("w:numPr"))
        if numPr is not None:
            item = ListItem(elements=[Paragraph(runs=runs)])
            return ListBlock(ordered=True, items=[item], tight=True)

        alignment = _alignment_to_str(para.alignment)
        return Paragraph(runs=runs, alignment=alignment)

    def _extract_runs_and_images(self, p_elem, doc) -> (List[InlineElement], List[Image]):
        runs: List[InlineElement] = []
        images: List[Image] = []

        for child in p_elem:
            tag = child.tag

            if tag == _qname("r"):
                text = self._get_run_text(child)
                rPr = child.find(_qname("rPr"))

                if text:
                    runs.append(self._build_text_run(text, rPr))

                # Check for image in drawing inside this run
                drawing_elem = child.find(_qname("drawing"))
                if drawing_elem is not None:
                    img_elem = drawing_elem.find(f"{{{_NS_WP}}}inline")
                    if img_elem is None:
                        img_elem = drawing_elem.find(f"{{{_NS_WP14}}}anchor")
                    if img_elem is not None:
                        img = self._extract_single_image(img_elem, doc)
                        if img:
                            images.append(img)

            elif tag == _qname("hyperlink"):
                hyperlink = self._extract_hyperlink(child, doc)
                if hyperlink:
                    runs.append(hyperlink)

            elif tag == f"{{{_NS_M}}}oMath":
                latex = self._omml_to_latex(child)
                if latex:
                    runs.append(Formula(latex=latex, display=False))

            elif tag == f"{{{_NS_M}}}oMathPara":
                for omath in child.iter(f"{{{_NS_M}}}oMath"):
                    latex = self._omml_to_latex(omath)
                    if latex:
                        runs.append(Formula(latex=latex, display=True))

        return runs, images

    def _extract_hyperlink(self, hyperlink_elem, doc) -> Optional[Hyperlink]:
        r_id = hyperlink_elem.get(_qname("id", _NS_R))
        url = ""
        if r_id:
            try:
                rel = doc.part.rels[r_id]
                url = rel.target_ref if rel else ""
            except (KeyError, AttributeError):
                pass
            if not url:
                # Fallback: read rels XML directly from the package
                try:
                    from docx.opc.constants import RELATIONSHIP_TYPE as RT
                    rels_xml = doc.part.rels._rels
                    for rel_key, rel_obj in rels_xml.items():
                        if hasattr(rel_obj, 'rId') and rel_obj.rId == r_id:
                            url = rel_obj.target_ref
                            break
                except Exception:
                    pass
            if not url:
                # Last fallback: parse word/_rels/document.xml.rels from zip
                try:
                    import io
                    with zipfile.ZipFile(doc.part.package._blob) if hasattr(doc.part.package, '_blob') else contextlib.nullcontext() as z:
                        pass
                except Exception:
                    pass

        hyper_runs: List[TextRun] = []
        for r in hyperlink_elem.iter(_qname("r")):
            text = self._get_run_text(r)
            if not text:
                continue
            rPr = r.find(_qname("rPr"))
            hyper_runs.append(self._build_text_run(text, rPr))

        if not hyper_runs:
            return None
        return Hyperlink(url=url, runs=hyper_runs)

    def _get_run_text(self, run_elem) -> str:
        texts: List[str] = []
        for child in run_elem:
            tag = child.tag
            if tag == _qname("t"):
                if child.text:
                    texts.append(child.text)
            elif tag == _qname("br"):
                texts.append("\n")
        return "".join(texts)

    def _build_text_run(self, text: str, rPr) -> TextRun:
        bold = False
        italic = False
        underline = False
        strike = False
        superscript = False
        subscript = False
        font_name = None
        font_size_pt = None

        if rPr is not None:
            b_elem = rPr.find(_qname("b"))
            if b_elem is not None:
                val = b_elem.get(_qname("val"))
                bold = val not in ("false", "0")

            i_elem = rPr.find(_qname("i"))
            if i_elem is not None:
                val = i_elem.get(_qname("val"))
                italic = val not in ("false", "0")

            u_elem = rPr.find(_qname("u"))
            if u_elem is not None:
                u_val = u_elem.get(_qname("val"))
                underline = u_val and u_val not in ("none", "0")

            s_elem = rPr.find(_qname("strike"))
            if s_elem is not None:
                val = s_elem.get(_qname("val"))
                strike = val not in ("false", "0")

            va = rPr.find(_qname("vertAlign"))
            if va is not None:
                val = va.get(_qname("val"))
                if val == "superscript":
                    superscript = True
                elif val == "subscript":
                    subscript = True

            rFonts = rPr.find(_qname("rFonts"))
            if rFonts is not None:
                font_name = rFonts.get(_qname("ascii")) or \
                            rFonts.get(_qname("hAnsi"))

            sz = rPr.find(_qname("sz"))
            if sz is not None:
                sz_val = sz.get(_qname("val"))
                if sz_val:
                    font_size_pt = int(sz_val) // 2

        code = _is_mono_font(font_name)
        return TextRun(
            text=text, bold=bold, italic=italic, code=code,
            underline=underline, strikethrough=strike,
            superscript=superscript, subscript=subscript,
            font_name=font_name, font_size=font_size_pt,
        )

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

    # --- Images ---
    def _extract_single_image(self, drawing, doc) -> Optional[Image]:
        extent = drawing.find(_qname("extent", _NS_WP))
        cx = int(extent.get("cx")) if extent is not None and extent.get("cx") else None
        cy = int(extent.get("cy")) if extent is not None and extent.get("cy") else None

        docPr = drawing.find(_qname("docPr", _NS_WP))
        alt_text = ""
        if docPr is not None:
            alt_text = docPr.get("descr") or docPr.get("name") or ""

        blip = drawing.find(f".//{_qname('blip', _NS_A)}")
        if blip is None:
            return None
        embed = blip.get(_qname("embed", _NS_R))
        if not embed:
            return None

        try:
            image_part = doc.part.related_parts[embed]
        except KeyError:
            return None
        except AttributeError:
            return None

        img_name = self._save_image(image_part)
        rel_path = "images/" + img_name

        width_str = f"{cx / 914400:.2f}in" if cx else None
        height_str = f"{cy / 914400:.2f}in" if cy else None

        # Run OCR if enabled
        ocr_text = None
        if self._ocr and self._ocr_engine and self.output_dir:
            self._ocr_count += 1
            if self._max_ocr_images == 0 or self._ocr_count <= self._max_ocr_images:
                img_path = os.path.join(self.output_dir, img_name)
                try:
                    result = self._ocr_engine.ocr(img_path)
                    if result and isinstance(result, list) and len(result) > 0:
                        r = result[0]
                        if hasattr(r, 'get'):
                            # New PaddleOCR API: result is a dict-like object
                            texts = r.get('rec_texts', [])
                            if texts:
                                ocr_text = "\n".join(texts)
                        elif isinstance(r, (list, tuple)):
                            # Old PaddleOCR API: result is list of [box, (text, score)]
                            lines = []
                            for item in r:
                                if isinstance(item, (list, tuple)) and len(item) >= 2:
                                    text_info = item[1]
                                    if isinstance(text_info, (list, tuple)):
                                        lines.append(text_info[0])
                                    else:
                                        lines.append(str(text_info))
                            if lines:
                                ocr_text = "\n".join(lines)
                except Exception:
                    pass

        return Image(src=rel_path, alt=alt_text, width=width_str, height=height_str,
                     ocr_text=ocr_text)

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
        ct = image_part.content_type or ""
        mapping = {
            "image/png": ".png", "image/jpeg": ".jpg",
            "image/gif": ".gif", "image/bmp": ".bmp",
            "image/tiff": ".tiff", "image/svg+xml": ".svg",
        }
        for c, ext in mapping.items():
            if c in ct:
                return ext
        return ".png"

    # --- OMML to LaTeX ---
    def _omml_to_latex(self, omml_elem) -> str:
        """Basic OMML to LaTeX converter for common patterns."""
        ns_m = _NS_M
        parts: List[str] = []

        for child in omml_elem:
            tag = child.tag
            if tag == f"{{{ns_m}}}r":
                # Run element - contains text
                for t in child.iter(f"{{{ns_m}}}t"):
                    if t.text:
                        parts.append(t.text)
            elif tag == f"{{{ns_m}}}frac":
                # Fraction: <m:num>...</m:num><m:den>...</m:den>
                num = child.find(f"{{{ns_m}}}num")
                den = child.find(f"{{{ns_m}}}den")
                num_latex = self._omml_to_latex(num) if num is not None else ""
                den_latex = self._omml_to_latex(den) if den is not None else ""
                parts.append(f"\\frac{{{num_latex}}}{{{den_latex}}}")
            elif tag == f"{{{ns_m}}}rad":
                # Radical: <m:deg>...</m:deg><m:e>...</m:e>
                deg = child.find(f"{{{ns_m}}}deg")
                e = child.find(f"{{{ns_m}}}e")
                deg_latex = self._omml_to_latex(deg) if deg is not None else ""
                e_latex = self._omml_to_latex(e) if e is not None else ""
                if deg_latex and deg_latex != "2":
                    parts.append(f"\\sqrt[{deg_latex}]{{{e_latex}}}")
                else:
                    parts.append(f"\\sqrt{{{e_latex}}}")
            elif tag == f"{{{ns_m}}}sSup":
                # Superscript: <m:e>...</m:e><m:sup>...</m:sup>
                e = child.find(f"{{{ns_m}}}e")
                sup = child.find(f"{{{ns_m}}}sup")
                e_latex = self._omml_to_latex(e) if e is not None else ""
                sup_latex = self._omml_to_latex(sup) if sup is not None else ""
                parts.append(f"{{{e_latex}}}^{{{sup_latex}}}")
            elif tag == f"{{{ns_m}}}sSub":
                # Subscript: <m:e>...</m:e><m:sub>...</m:sub>
                e = child.find(f"{{{ns_m}}}e")
                sub = child.find(f"{{{ns_m}}}sub")
                e_latex = self._omml_to_latex(e) if e is not None else ""
                sub_latex = self._omml_to_latex(sub) if sub is not None else ""
                parts.append(f"{{{e_latex}}}_{{{sub_latex}}}")
            elif tag == f"{{{ns_m}}}sSubSup":
                # Sub-superscript
                e = child.find(f"{{{ns_m}}}e")
                sub = child.find(f"{{{ns_m}}}sub")
                sup = child.find(f"{{{ns_m}}}sup")
                e_latex = self._omml_to_latex(e) if e is not None else ""
                sub_latex = self._omml_to_latex(sub) if sub is not None else ""
                sup_latex = self._omml_to_latex(sup) if sup is not None else ""
                parts.append(f"{{{e_latex}}}_{{{sub_latex}}}^{{{sup_latex}}}")
            elif tag == f"{{{ns_m}}}nary":
                # N-ary operator (sum, integral, etc.)
                nary_pr = child.find(f"{{{ns_m}}}naryPr")
                chr_val = ""
                if nary_pr is not None:
                    chr_elem = nary_pr.find(f"{{{ns_m}}}chr")
                    if chr_elem is not None:
                        chr_val = chr_elem.get(f"{{{ns_m}}}val", "")
                e = child.find(f"{{{ns_m}}}e")
                e_latex = self._omml_to_latex(e) if e is not None else ""
                chr_map = {"∑": "\\sum", "∫": "\\int", "∏": "\\prod", "∬": "\\iint", "∭": "\\iiint"}
                latex_cmd = chr_map.get(chr_val, f"\\operatorname{{{chr_val}}}")
                parts.append(f"{latex_cmd} {e_latex}")
            elif tag == f"{{{ns_m}}}func":
                # Function: <m:fName>...</m:fName><m:e>...</m:e>
                fname = child.find(f"{{{ns_m}}}fName")
                e = child.find(f"{{{ns_m}}}e")
                fname_latex = self._omml_to_latex(fname) if fname is not None else ""
                e_latex = self._omml_to_latex(e) if e is not None else ""
                parts.append(f"{fname_latex}({e_latex})")
            elif tag == f"{{{ns_m}}}d":
                # Delimiter (parentheses, brackets, etc.)
                inner_parts = []
                for sub in child:
                    sub_tag = sub.tag
                    if sub_tag == f"{{{ns_m}}}e":
                        inner_parts.append(self._omml_to_latex(sub))
                parts.append("\\left( " + " ".join(inner_parts) + " \\right)")
            elif tag == f"{{{ns_m}}}eqArr":
                # Equation array
                rows = []
                for sub in child:
                    if sub.tag == f"{{{ns_m}}}e":
                        rows.append(self._omml_to_latex(sub))
                parts.append("\\begin{aligned} " + " \\\\ ".join(rows) + " \\end{aligned}")

        return "".join(parts)
    def _extract_table_xml(self, tbl_elem, doc: DocxDocument) -> Optional[Table]:
        try:
            tree = etree.fromstring(etree.tostring(tbl_elem))
        except Exception:
            return None

        rows_xml = tree.findall(f".//{_qname('tr')}")
        if not rows_xml:
            return None

        all_rows: List[List[str]] = []
        aligns: List[Optional[str]] = []

        for row_idx, row_xml in enumerate(rows_xml):
            cells = row_xml.findall(f".//{_qname('tc')}")
            row_data: List[str] = []
            for cell_idx, cell in enumerate(cells):
                texts: List[str] = []
                for p in cell.findall(f".//{_qname('p')}"):
                    for t in p.findall(f".//{_qname('t')}"):
                        if t.text:
                            texts.append(t.text)
                    texts.append("\n")
                cell_text = "".join(texts).strip()
                row_data.append(cell_text)

                if row_idx == 0:
                    p = cell.find(f".//{_qname('p')}")
                    align_val = None
                    if p is not None:
                        pPr = p.find(_qname("pPr"))
                        if pPr is not None:
                            jc = pPr.find(_qname("jc"))
                            if jc is not None:
                                val = jc.get(_qname("val"))
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

    # --- Footnotes & Comments ---
    def _load_footnotes(self, doc) -> dict:
        fn_map = {}
        try:
            part = doc.part.package.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            )
            tree = etree.fromstring(part.blob)
            for fn in tree.findall(f".//{_qname('footnote')}"):
                fn_id = fn.get(_qname("id"))
                texts: List[str] = []
                for t in fn.iter(_qname("t")):
                    if t.text:
                        texts.append(t.text)
                fn_map[fn_id] = "".join(texts)
        except Exception:
            pass
        return fn_map

    def _get_para_footnotes(self, p_elem, fn_map: dict) -> List[Footnote]:
        results = []
        try:
            for ref in p_elem.iter(_qname("footnoteReference")):
                fn_id = ref.get(_qname("id"))
                if fn_id and fn_id in fn_map:
                    results.append(Footnote(footnote_id=fn_id, text=fn_map[fn_id]))
        except Exception:
            pass
        return results

    def _load_comments(self, doc) -> dict:
        cm_map = {}
        try:
            part = doc.part.package.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            )
            tree = etree.fromstring(part.blob)
            for cm in tree.findall(f".//{_qname('comment')}"):
                cm_id = cm.get(_qname("id"), "")
                author = cm.get(_qname("author"), "")
                date_str = cm.get(_qname("date"), "")
                texts: List[str] = []
                for t in cm.iter(_qname("t")):
                    if t.text:
                        texts.append(t.text)
                cm_map[cm_id] = {
                    "author": author,
                    "text": "".join(texts),
                    "date": date_str or None,
                }
        except Exception:
            pass
        return cm_map

    def _get_para_comments(self, p_elem, cm_map: dict) -> List[Comment]:
        results = []
        try:
            for ref in p_elem.iter(_qname("commentReference")):
                cm_id = ref.get(_qname("id"))
                if cm_id and cm_id in cm_map:
                    info = cm_map[cm_id]
                    results.append(Comment(author=info["author"], text=info["text"],
                                           date=info["date"]))
        except Exception:
            pass
        return results

    # --- Metadata, Headers, Footers, Sections ---

    def _extract_metadata(self, doc: DocxDocument) -> dict:
        """Extract document metadata (title, author, created, modified)."""
        metadata = {}
        props = doc.core_properties
        if props.title:
            metadata["title"] = props.title
        if props.author:
            metadata["author"] = props.author
        if props.created:
            metadata["created"] = props.created.isoformat()
        if props.modified:
            metadata["modified"] = props.modified.isoformat()
        if props.subject:
            metadata["subject"] = props.subject
        if props.keywords:
            metadata["keywords"] = props.keywords
        return metadata

    def _extract_headers(self, doc: DocxDocument) -> List[str]:
        """Extract header text from all sections."""
        headers = []
        try:
            for section in doc.sections:
                header = section.header
                if header and not header.is_linked_to_previous:
                    texts = []
                    for para in header.paragraphs:
                        text = para.text.strip()
                        if text:
                            texts.append(text)
                    if texts:
                        headers.append("\n".join(texts))
        except Exception:
            pass
        return headers

    def _extract_footers(self, doc: DocxDocument) -> List[str]:
        """Extract footer text from all sections."""
        footers = []
        try:
            for section in doc.sections:
                footer = section.footer
                if footer and not footer.is_linked_to_previous:
                    texts = []
                    for para in footer.paragraphs:
                        text = para.text.strip()
                        if text:
                            texts.append(text)
                    if texts:
                        footers.append("\n".join(texts))
        except Exception:
            pass
        return footers

    def _extract_sections(self, doc: DocxDocument) -> List[dict]:
        """Extract section properties (page size, margins, orientation)."""
        sections = []
        try:
            for section in doc.sections:
                sec_info = {}
                # Page size
                if section.page_width:
                    sec_info["page_width"] = f"{section.page_width / 914400:.2f}in"
                if section.page_height:
                    sec_info["page_height"] = f"{section.page_height / 914400:.2f}in"
                # Margins
                if section.top_margin:
                    sec_info["top_margin"] = f"{section.top_margin / 914400:.2f}in"
                if section.bottom_margin:
                    sec_info["bottom_margin"] = f"{section.bottom_margin / 914400:.2f}in"
                if section.left_margin:
                    sec_info["left_margin"] = f"{section.left_margin / 914400:.2f}in"
                if section.right_margin:
                    sec_info["right_margin"] = f"{section.right_margin / 914400:.2f}in"
                # Orientation
                from docx.enum.section import WD_ORIENT
                if section.orientation == WD_ORIENT.LANDSCAPE:
                    sec_info["orientation"] = "landscape"
                else:
                    sec_info["orientation"] = "portrait"
                sections.append(sec_info)
        except Exception:
            pass
        return sections
