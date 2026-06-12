from __future__ import annotations
import os
from typing import Dict, List, Optional, Set

from docx import Document as DocxDocument
from docx.oxml.ns import qn
from lxml import etree

from md2word.model.document import (
    TextRun, Image, Heading, Paragraph, CodeBlock, Hyperlink,
    ListBlock, ListItem, Table, HorizontalRule, Formula, PageBreak,
    Footnote, Comment, Blockquote,
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
    def __init__(self, output_dir: str = "", ocr: bool = False, max_ocr_images: int = 0,
                 style_mappings: Optional[Dict[str, str]] = None, ocr_engine: str = "glm-ocr"):
        self.output_dir = output_dir
        self._img_counter: int = 0
        self._saved_images: Set[str] = set()
        self._ocr = ocr
        self._ocr_engine = None
        self._ocr_engine_name = ocr_engine
        self._ocr_count = 0
        self._max_ocr_images = max_ocr_images
        self._style_mappings = style_mappings or {}
        self._numbering_cache: Dict[int, dict] = {}
        self._number_counters: Dict[int, int] = {}
        
        if ocr:
            self._init_ocr_engine(ocr_engine)

    def _init_ocr_engine(self, engine_name: str):
        """Initialize OCR engine with fallback."""
        if engine_name == "glm-ocr":
            # Try glm-ocr (ollama) first
            try:
                import subprocess
                result = subprocess.run(
                    ["ollama", "list"],
                    capture_output=True, timeout=10
                )
                if result.returncode == 0 and "glm-ocr" in result.stdout.decode("utf-8", errors="replace"):
                    self._ocr_engine = "glm-ocr"
                    return
            except (ImportError, FileNotFoundError, subprocess.TimeoutExpired):
                pass
            # Fallback to PaddleOCR
            print("glm-ocr not available, falling back to PaddleOCR...")
            self._init_paddleocr()
        elif engine_name == "paddleocr":
            self._init_paddleocr()
        else:
            print(f"Unknown OCR engine: {engine_name}")

    def _init_paddleocr(self):
        """Initialize PaddleOCR engine."""
        try:
            from paddleocr import PaddleOCR
            import warnings
            warnings.filterwarnings("ignore")
            self._ocr_engine = PaddleOCR(use_textline_orientation=True, lang="ch")
            self._ocr_engine_name = "paddleocr"
        except ImportError:
            print("PaddleOCR not available. Install with: pip install paddleocr paddlepaddle")

    def extract(self, docx_path: str) -> Document:
        doc = DocxDocument(docx_path)
        elements: List[BlockElement] = []
        list_buffer: List[ListItem] = []
        list_ordered: bool = False
        list_tight: bool = True
        list_level: int = 0
        code_buffer: List[str] = []

        # Multi-level numbering counters
        self._number_counters: Dict[int, int] = {}

        # Extract metadata
        metadata = self._extract_metadata(doc)

        # Extract headers and footers
        headers = self._extract_headers(doc)
        footers = self._extract_footers(doc)

        # Extract section properties
        sections = self._extract_sections(doc)

        # Load numbering definitions
        self._load_numbering(doc)

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

                # Blockquote
                bq_level = self._is_blockquote(para)
                if bq_level > 0:
                    self._flush(elements, code_buffer, list_buffer, list_ordered, list_tight)
                    list_buffer = []
                    list_tight = True
                    code_buffer = []
                    runs, _ = self._extract_runs_and_images(child, doc)
                    # Check if previous element is also a blockquote at same level
                    if elements and isinstance(elements[-1], Blockquote) and elements[-1].level == bq_level:
                        # Merge with previous blockquote
                        elements[-1].runs.append(TextRun(text="\n"))
                        elements[-1].runs.extend(runs)
                    else:
                        elements.append(Blockquote(runs=runs, level=bq_level))
                    elements.extend(self._get_para_footnotes(child, footnotes))
                    elements.extend(self._get_para_comments(child, comments))
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

                element = self._extract_paragraph_element(para, runs, child)

                if isinstance(element, ListBlock):
                    items = element.items
                    if items:
                        if list_buffer and list_ordered == element.ordered and list_level == element.level:
                            # Same level, merge items
                            list_buffer.extend(items)
                        else:
                            # Different level or type, flush and start new
                            self._flush_list_buffer(elements, list_buffer, list_ordered, list_tight, list_level)
                            list_buffer = list(items)
                            list_ordered = element.ordered
                            list_tight = element.tight
                            list_level = element.level
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

    def _flush_list_buffer(self, elements, buffer, ordered, tight, level=0):
        if buffer:
            elements.append(ListBlock(ordered=ordered, items=buffer, tight=tight, level=level))

    def _find_paragraph_in_doc(self, doc: DocxDocument, xml_elem):
        for para in doc.paragraphs:
            if para._p is xml_elem:
                return para
        return None

    def _detect_toc(self, doc: DocxDocument) -> bool:
        """Detect if document has a Table of Contents field."""
        try:
            for para in doc.paragraphs:
                for run in para.runs:
                    # Check for TOC field codes
                    for child in run._r:
                        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                        if tag == "fldChar":
                            fldCharType = child.get(qn("w:fldCharType"))
                            if fldCharType == "begin":
                                # Look for instrText with TOC
                                next_elem = child.getnext()
                                while next_elem is not None:
                                    next_tag = next_elem.tag.split("}")[-1] if "}" in next_elem.tag else next_elem.tag
                                    if next_tag == "instrText":
                                        if "TOC" in (next_elem.text or ""):
                                            return True
                                    elif next_tag == "fldChar":
                                        break
                                    next_elem = next_elem.getnext()
        except Exception:
            pass
        return False

    def _detect_bookmarks(self, doc: DocxDocument) -> List[str]:
        """Detect bookmarks in the document."""
        bookmarks = []
        try:
            for para in doc.paragraphs:
                for child in para._p:
                    tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    if tag == "bookmarkStart":
                        name = child.get(qn("w:name"))
                        if name and name.startswith("heading_"):
                            bookmarks.append(name)
        except Exception:
            pass
        return bookmarks

    def _has_page_break(self, p_elem) -> bool:
        try:
            for br in p_elem.iter(_qname("br")):
                if br.get(_qname("type")) == "page":
                    return True
        except Exception:
            pass
        return False

    def _extract_paragraph_element(self, para, runs, p_elem=None) -> Optional[BlockElement]:
        style_name = para.style.name if para.style else ""

        # Check for custom style mapping
        if style_name in self._style_mappings:
            mapping = self._style_mappings[style_name]
            # Return as special paragraph with style info
            return Paragraph(runs=runs, alignment=None)

        if style_name.startswith("Heading"):
            try:
                level = int(style_name.split()[-1])
            except (ValueError, IndexError):
                level = 1
            return Heading(level=level, runs=runs)

        # Detect list style with nesting level
        if style_name.startswith("List Bullet"):
            ordered = False
            # Extract level from style name: "List Bullet" = 0, "List Bullet 2" = 1, etc.
            level = 0
            if style_name != "List Bullet":
                try:
                    level = int(style_name.split()[-1]) - 1
                except (ValueError, IndexError):
                    level = 0
            # Check for task list checkboxes in text
            checked = None
            flat_text = "".join(r.text for r in runs if isinstance(r, TextRun))
            if flat_text.startswith("\u2610 ") or flat_text.startswith("[ ] "):  # ☐ or [ ]
                checked = False
            elif flat_text.startswith("\u2611 ") or flat_text.startswith("[x] ") or flat_text.startswith("[X] "):  # ☑ or [x]
                checked = True
            # Remove checkbox prefix from text runs
            if checked is not None:
                for r in runs:
                    if isinstance(r, TextRun):
                        if r.text.startswith("\u2610 ") or r.text.startswith("[ ] "):
                            r.text = r.text[2:] if r.text.startswith("\u2610 ") else r.text[4:]
                        elif r.text.startswith("\u2611 ") or r.text.startswith("[x] ") or r.text.startswith("[X] "):
                            r.text = r.text[2:] if r.text.startswith("\u2611 ") else r.text[4:]
            item = ListItem(elements=[Paragraph(runs=runs)], checked=checked)
            tight = not para.paragraph_format.space_before \
                    and not para.paragraph_format.space_after
            return ListBlock(ordered=ordered, items=[item], tight=tight, level=level)

        if style_name.startswith("List Number"):
            ordered = True
            # Extract level from style name
            level = 0
            if style_name != "List Number":
                try:
                    level = int(style_name.split()[-1]) - 1
                except (ValueError, IndexError):
                    level = 0
            item = ListItem(elements=[Paragraph(runs=runs)])
            tight = not para.paragraph_format.space_before \
                    and not para.paragraph_format.space_after
            return ListBlock(ordered=ordered, items=[item], tight=tight, level=level)

        # Check for numPr with ilvl
        if p_elem is not None:
            pPr = p_elem.find(_qname("pPr"))
            if pPr is not None:
                numPr = pPr.find(_qname("numPr"))
                if numPr is not None:
                    ilvl_elem = numPr.find(_qname("ilvl"))
                    numId_elem = numPr.find(_qname("numId"))
                    ilvl = int(ilvl_elem.get(_qname("val"))) if ilvl_elem is not None else 0
                    numId = int(numId_elem.get(_qname("val"))) if numId_elem is not None else 0

                    # Determine if ordered based on numbering format
                    ordered = True
                    text_fmt = "%1."
                    if numId in self._numbering_cache:
                        num_info = self._numbering_cache[numId]
                        fmt = num_info.get("fmt", "decimal")
                        ordered = fmt != "bullet"
                        # Get format for this level
                        levels = num_info.get("levels", {})
                        level_info = levels.get(ilvl, {"fmt": "decimal", "text": "%1."})
                        text_fmt = level_info.get("text", "%1.")

                    # Track numbering for multi-level lists
                    self._number_counters[ilvl] = self._number_counters.get(ilvl, 0) + 1
                    # Reset counters for deeper levels
                    for deeper_level in list(self._number_counters.keys()):
                        if deeper_level > ilvl:
                            self._number_counters[deeper_level] = 0

                    # Generate numbering prefix based on format
                    prefix = self._generate_numbering_prefix(ilvl, text_fmt)

                    item = ListItem(elements=[Paragraph(runs=runs)])
                    return ListBlock(ordered=ordered, items=[item], tight=True,
                                    level=ilvl, numbering_prefix=prefix)

                # Check for indentation-based list detection
                ind = pPr.find(_qname("ind"))
                if ind is not None:
                    left_val = ind.get(_qname("left"))
                    if left_val:
                        left_twips = int(left_val)
                        # Calculate level based on indentation (720 twips = 0.5 inch = 1 level)
                        level = max(0, (left_twips - 720) // 720) if left_twips > 720 else 0
                        # Check if text starts with bullet or number pattern
                        flat_text = "".join(r.text for r in runs if isinstance(r, TextRun))
                        if flat_text and level > 0:
                            # Detect bullet or number pattern
                            is_bullet = flat_text.startswith(("•", "◦", "▪", "-", "–"))
                            import re
                            is_number = bool(re.match(r'^\d+\.?\s', flat_text))
                            if is_bullet or is_number:
                                item = ListItem(elements=[Paragraph(runs=runs)])
                                return ListBlock(ordered=is_number, items=[item], tight=True,
                                                level=level)

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

    def _is_blockquote(self, para) -> int:
        """Check if paragraph is a blockquote. Returns blockquote level (0 if not)."""
        try:
            pPr = para._p.find(qn("w:pPr"))
            if pPr is None:
                return 0
            # Check for left border (blockquote indicator)
            pBdr = pPr.find(qn("w:pBdr"))
            if pBdr is not None:
                left = pBdr.find(qn("w:left"))
                if left is not None and left.get(qn("w:val")) == "single":
                    # Determine level from left indent
                    ind = pPr.find(qn("w:ind"))
                    if ind is not None:
                        left_val = ind.get(qn("w:left"))
                        if left_val:
                            left_twips = int(left_val)
                            level = max(1, left_twips // 720)  # 720 twips = 0.5 inch
                            return min(level, 3)  # Max level 3
                    return 1
            # Check for Quote style
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_val = pStyle.get(qn("w:val"))
                if style_val and ("quote" in style_val.lower() or "blockquote" in style_val.lower()):
                    return 1
        except Exception:
            pass
        return 0

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
        if self._ocr and self.output_dir:
            self._ocr_count += 1
            if self._max_ocr_images == 0 or self._ocr_count <= self._max_ocr_images:
                img_path = os.path.join(self.output_dir, img_name)
                try:
                    if self._ocr_engine_name == "glm-ocr":
                        ocr_text = self._run_glm_ocr(img_path)
                    elif self._ocr_engine_name == "paddleocr" and self._ocr_engine:
                        ocr_text = self._run_paddleocr(img_path)
                except Exception:
                    pass

        return Image(src=rel_path, alt=alt_text, width=width_str, height=height_str,
                     ocr_text=ocr_text)

    def _run_glm_ocr(self, img_path: str) -> Optional[str]:
        """Run glm-ocr (ollama) on an image."""
        import subprocess
        import re
        try:
            abs_path = os.path.abspath(img_path)
            result = subprocess.run(
                ["ollama", "run", "glm-ocr", f"Text Recognition: {abs_path}"],
                capture_output=True, timeout=60
            )
            text = result.stdout.decode("utf-8", errors="replace").strip()
            # Clean ANSI escape codes
            text = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', text)
            text = re.sub(r'\[\?25[hl]', '', text)
            text = re.sub(r'\[\?2026[hl]', '', text)
            text = re.sub(r'\[K', '', text)
            text = re.sub(r'\[2K', '', text)
            text = re.sub(r'\[1G', '', text)
            # Clean up markdown formatting
            if text.startswith("```markdown"):
                text = text[len("```markdown"):].strip()
            if text.endswith("```"):
                text = text[:-3].strip()
            return text if text else None
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return None

    def _run_paddleocr(self, img_path: str) -> Optional[str]:
        """Run PaddleOCR on an image."""
        try:
            result = self._ocr_engine.ocr(img_path)
            if result and isinstance(result, list) and len(result) > 0:
                r = result[0]
                if hasattr(r, 'get'):
                    # New PaddleOCR API: result is a dict-like object
                    texts = r.get('rec_texts', [])
                    if texts:
                        return "\n".join(texts)
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
                        return "\n".join(lines)
        except Exception:
            pass
        return None

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

        # Detect TOC
        if self._detect_toc(doc):
            metadata["toc"] = True

        # Detect bookmarks
        bookmarks = self._detect_bookmarks(doc)
        if bookmarks:
            metadata["bookmarks"] = bookmarks

        # Detect RTL
        if self._detect_rtl(doc):
            metadata["rtl"] = True

        return metadata

    def _detect_rtl(self, doc: DocxDocument) -> bool:
        """Detect if document has Right-to-Left text direction."""
        try:
            for section in doc.sections:
                sectPr = section._sectPr
                pPr = sectPr.find(qn("w:pPr"))
                if pPr is not None:
                    bidi = pPr.find(qn("w:bidi"))
                    if bidi is not None:
                        return True
        except Exception:
            pass
        return False

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
        """Extract section properties (page size, margins, orientation, columns)."""
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
                # Columns
                try:
                    if section.column_count > 1:
                        sec_info["columns"] = section.column_count
                except AttributeError:
                    pass
                sections.append(sec_info)
        except Exception:
            pass
        return sections

    def _generate_numbering_prefix(self, level: int, text_fmt: str) -> str:
        """Generate numbering prefix based on format and level."""
        import re
        # Replace %N with counter values
        counters = {}
        for i in range(level + 1):
            counters[i + 1] = self._number_counters.get(i, 1)

        def replace_counter(m):
            n = int(m.group(1))
            return str(counters.get(n, 1))

        prefix = re.sub(r'%(\d+)', replace_counter, text_fmt)
        return prefix

    def _load_numbering(self, doc: DocxDocument):
        """Load numbering definitions from word/numbering.xml."""
        try:
            numbering_part = doc.part.numbering_part
            if numbering_part is None:
                return
            tree = etree.fromstring(numbering_part.blob)
            ns_w = _NS_W

            # Build abstract num lookup
            abstract_nums = {}
            for abstract in tree.findall(f".//{{{ns_w}}}abstractNum"):
                abstract_id = int(abstract.get(f"{{{ns_w}}}abstractNumId"))
                levels = {}
                for lvl in abstract.findall(f"{{{ns_w}}}lvl"):
                    ilvl = int(lvl.get(f"{{{ns_w}}}ilvl"))
                    num_fmt = lvl.find(f"{{{ns_w}}}numFmt")
                    fmt = num_fmt.get(f"{{{ns_w}}}val") if num_fmt is not None else "decimal"
                    lvl_text = lvl.find(f"{{{ns_w}}}lvlText")
                    text_fmt = lvl_text.get(f"{{{ns_w}}}val") if lvl_text is not None else "%1."
                    levels[ilvl] = {"fmt": fmt, "text": text_fmt}
                abstract_nums[abstract_id] = levels

            # Map numId to abstract num
            for num in tree.findall(f".//{{{ns_w}}}num"):
                num_id = int(num.get(f"{{{ns_w}}}numId"))
                abstract_ref = num.find(f"{{{ns_w}}}abstractNumId")
                if abstract_ref is not None:
                    abstract_id = int(abstract_ref.get(f"{{{ns_w}}}val"))
                    if abstract_id in abstract_nums:
                        levels = abstract_nums[abstract_id]
                        # Get format for level 0 (or first available)
                        level_0 = levels.get(0, {"fmt": "decimal", "text": "%1."})
                        self._numbering_cache[num_id] = {
                            "abstract_id": abstract_id,
                            "levels": levels,
                            "fmt": level_0["fmt"],
                            "text": level_0["text"],
                        }
        except Exception:
            pass
