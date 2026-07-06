from __future__ import annotations

import base64
import html
import json
import mimetypes
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from md2word.meta import extract_docx_metadata


META_FILENAME = "document.json"


def render_docx_to_html(docx_path: str, output_path: str) -> str:
    return DocxHtmlRenderer().render_docx(docx_path, output_path)


def render_metadata_to_html(meta_path: str, output_path: str) -> str:
    return DocxHtmlRenderer().render_metadata(meta_path, output_path)


class DocxHtmlRenderer:
    """Render the structured DOCX metadata index to a standalone HTML file."""

    def render_docx(self, docx_path: str, output_path: str) -> str:
        out = Path(output_path)
        meta_dir = out.with_suffix("").with_name(f"{out.stem}_html_assets") / "meta"
        metadata = extract_docx_metadata(docx_path, str(meta_dir))
        return self.render(metadata, meta_dir, output_path)

    def render_metadata(self, meta_path: str, output_path: str) -> str:
        metadata, base_dir = _load_metadata(meta_path)
        return self.render(metadata, base_dir, output_path)

    def render(self, metadata: Dict[str, Any], base_dir: Path, output_path: str) -> str:
        document = metadata.get("document", {})
        out = Path(output_path)
        if out.parent:
            out.parent.mkdir(parents=True, exist_ok=True)

        body = self._render_body(document.get("body", []), document)
        header = self._render_header_footer(document.get("headers_footers", {}).get("headers", {}), "header")
        footer = self._render_header_footer(document.get("headers_footers", {}).get("footers", {}), "footer")
        notes = self._render_notes(document)

        title = _html_escape(metadata.get("source", {}).get("filename") or out.stem)
        page_style = self._page_style(document)
        html_text = "\n".join([
            "<!doctype html>",
            '<html lang="zh-CN">',
            "<head>",
            '  <meta charset="utf-8">',
            '  <meta name="viewport" content="width=device-width, initial-scale=1">',
            f"  <title>{title}</title>",
            "  <script>",
            "    window.MathJax = {tex: {inlineMath: [['\\\\(', '\\\\)']], displayMath: [['\\\\[', '\\\\]']]}};",
            "  </script>",
            '  <script defer src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js"></script>',
            "  <style>",
            self._stylesheet(page_style),
            "  </style>",
            "</head>",
            "<body>",
            '<main class="docx-page">',
            header,
            '<section class="docx-body">',
            body,
            "</section>",
            footer,
            notes,
            "</main>",
            "</body>",
            "</html>",
            "",
        ])
        html_text = self._inline_resources(html_text, base_dir)
        out.write_text(html_text, encoding="utf-8")
        return str(out)

    def _render_body(self, elements: List[Dict[str, Any]], document: Dict[str, Any]) -> str:
        parts: List[str] = []
        index = 0
        while index < len(elements):
            element = elements[index]
            if _is_numbered_paragraph(element):
                list_html, index = self._render_list(elements, index, document)
                parts.append(list_html)
                continue
            parts.append(self._render_block(element, document))
            index += 1
        return "\n".join(part for part in parts if part)

    def _render_block(self, element: Dict[str, Any], document: Dict[str, Any]) -> str:
        element_type = element.get("type")
        if element_type == "paragraph":
            return self._render_paragraph(element, document)
        if element_type == "table":
            return self._render_table(element, document)
        if element_type == "structured_document_tag":
            body = element.get("body")
            if body:
                return f'<div class="sdt">{self._render_body(body, document)}</div>'
            return f'<span class="sdt">{_html_escape(element.get("text", ""))}</span>'
        if element_type == "revision":
            body = element.get("body")
            if body:
                return f'<div class="revision revision-{_html_attr(element.get("kind", ""))}">{self._render_body(body, document)}</div>'
            return self._render_runs(element.get("runs", []), document)
        if element_type == "section_properties":
            return '<div class="section-break"></div>'
        return ""

    def _render_paragraph(self, paragraph: Dict[str, Any], document: Dict[str, Any]) -> str:
        props = paragraph.get("properties", {})
        style_id = props.get("style", "")
        tag = _heading_tag(style_id)
        css = self._paragraph_css(props)
        classes = ["paragraph"]
        if paragraph.get("fields"):
            classes.append("has-fields")
        content = self._render_runs(paragraph.get("runs", []), document)
        if not content and paragraph.get("text"):
            content = _html_escape(paragraph.get("text", ""))
        content = self._render_markers(paragraph.get("markers", [])) + content
        if not content:
            content = "&nbsp;"
        class_attr = " ".join(classes)
        return f'<{tag} class="{class_attr}" style="{css}">{content}</{tag}>'

    def _render_markers(self, markers: List[Dict[str, Any]]) -> str:
        rendered = []
        for marker in markers:
            if marker.get("type") != "bookmarkStart":
                continue
            name = marker.get("attributes", {}).get("w:name")
            if name:
                rendered.append(f'<a id="{_html_attr(name)}" class="bookmark-anchor"></a>')
        return "".join(rendered)

    def _render_runs(self, runs: List[Dict[str, Any]], document: Dict[str, Any]) -> str:
        return "".join(self._render_inline(run, document) for run in runs)

    def _render_inline(self, item: Dict[str, Any], document: Dict[str, Any]) -> str:
        item_type = item.get("type")
        if item_type == "hyperlink":
            rel = item.get("relationship", {})
            href = rel.get("target") or f"#{item.get('anchor', '')}"
            label = self._render_runs(item.get("runs", []), document) or _html_escape(item.get("text", ""))
            return f'<a href="{_html_attr(href)}">{label}</a>'
        if item_type == "structured_document_tag":
            return self._render_runs(item.get("runs", []), document) or _html_escape(item.get("text", ""))
        if item_type == "revision":
            tag = "del" if item.get("kind") in {"del", "moveFrom"} else "ins"
            body = self._render_runs(item.get("runs", []), document) or _html_escape(item.get("text", ""))
            return f'<{tag} class="revision revision-{_html_attr(item.get("kind", ""))}">{body}</{tag}>'

        pieces: List[str] = []
        for drawing in item.get("drawings", []):
            pieces.append(self._render_drawing(drawing))
        for vml in item.get("vml", []):
            pieces.append(self._render_vml(vml, document))
        for math_item in item.get("math", []):
            pieces.append(self._render_math(math_item))

        text = item.get("text") or item.get("deleted_text") or ""
        if text:
            text_html = _html_escape(text).replace("\t", '<span class="tab"></span>').replace("\n", "<br>")
            pieces.append(self._style_text(text_html, item.get("properties", {})))

        for br in item.get("breaks", []):
            if br.get("w:type") == "page":
                pieces.append('<div class="page-break"></div>')
            else:
                pieces.append("<br>")
        for ref in item.get("footnote_references", []):
            note_id = ref.get("w:id", "")
            pieces.append(f'<sup><a href="#footnote-{_html_attr(note_id)}">[{_html_escape(note_id)}]</a></sup>')
        for ref in item.get("endnote_references", []):
            note_id = ref.get("w:id", "")
            pieces.append(f'<sup><a href="#endnote-{_html_attr(note_id)}">[{_html_escape(note_id)}]</a></sup>')
        for ref in item.get("comment_references", []):
            comment_id = ref.get("w:id", "")
            pieces.append(f'<sup><a href="#comment-{_html_attr(comment_id)}">[注{_html_escape(comment_id)}]</a></sup>')
        return "".join(pieces)

    def _style_text(self, text_html: str, props: Dict[str, Any]) -> str:
        css = self._run_css(props)
        wrappers: List[Tuple[str, str]] = []
        if props.get("bold"):
            wrappers.append(("strong", ""))
        if props.get("italic"):
            wrappers.append(("em", ""))
        if props.get("strike") or props.get("double_strike"):
            wrappers.append(("del", ""))
        if props.get("vertical_align") == "superscript":
            wrappers.append(("sup", ""))
        elif props.get("vertical_align") == "subscript":
            wrappers.append(("sub", ""))

        body = f'<span style="{css}">{text_html}</span>' if css else text_html
        for tag, attrs in wrappers:
            body = f"<{tag}{attrs}>{body}</{tag}>"
        return body

    def _render_drawing(self, drawing: Dict[str, Any]) -> str:
        resource = drawing.get("resource")
        if not resource:
            return ""
        extent = drawing.get("extent", {})
        style_parts = []
        if extent.get("cx"):
            style_parts.append(f"width:{_emu_to_px(extent['cx']):.2f}px")
        if extent.get("cy"):
            style_parts.append(f"height:{_emu_to_px(extent['cy']):.2f}px")
        wrap = drawing.get("wrap", {})
        if drawing.get("type") == "anchor":
            style_parts.append("position:relative")
            if wrap.get("type"):
                style_parts.append("margin:0.35rem")
        doc_props = drawing.get("doc_properties", {})
        alt = doc_props.get("descr") or doc_props.get("title") or doc_props.get("name") or ""
        src = f"md2word-resource:{resource}"
        return (
            f'<img class="docx-image" src="{_html_attr(src)}" '
            f'alt="{_html_attr(alt)}" style="{";".join(style_parts)}">'
        )

    def _render_vml(self, vml: Dict[str, Any], document: Dict[str, Any]) -> str:
        textbox = vml.get("textbox")
        text = textbox.get("text") if textbox else vml.get("text", "")
        if not text:
            return ""
        style = vml.get("style", {})
        css = []
        if style.get("width"):
            css.append(f"width:{style['width']}")
        if style.get("height"):
            css.append(f"min-height:{style['height']}")
        return f'<span class="vml-shape" style="{";".join(css)}">{_html_escape(text)}</span>'

    def _render_math(self, math_item: Dict[str, Any]) -> str:
        text = math_item.get("text", "").strip()
        if not text:
            return '<span class="math math-omml">[OMML]</span>'
        return f'<span class="math">\\({_html_escape(text)}\\)</span>'

    def _render_table(self, table: Dict[str, Any], document: Dict[str, Any]) -> str:
        rows = table.get("rows", [])
        if not rows:
            return ""
        table_css = self._table_css(table.get("properties", {}))
        html_rows = []
        for row_index, row in enumerate(rows):
            cells_html = []
            cells = row.get("cells", [])
            for cell in cells:
                props = cell.get("properties", {})
                vmerge = props.get("vertical_merge", {})
                if vmerge and vmerge.get("w:val") == "continue":
                    continue
                attrs = []
                grid_span = props.get("grid_span")
                if grid_span:
                    attrs.append(f'colspan="{_html_attr(grid_span)}"')
                rowspan = self._vertical_rowspan(rows, row_index, cell.get("index"))
                if rowspan > 1:
                    attrs.append(f'rowspan="{rowspan}"')
                css = self._cell_css(props)
                if css:
                    attrs.append(f'style="{css}"')
                body = self._render_body(cell.get("body", []), document) or _html_escape(cell.get("text", ""))
                tag = "th" if row_index == 0 and table.get("properties", {}).get("style") else "td"
                cells_html.append(f"<{tag} {' '.join(attrs)}>{body}</{tag}>")
            html_rows.append(f"<tr>{''.join(cells_html)}</tr>")
        return f'<table class="docx-table" style="{table_css}"><tbody>{"".join(html_rows)}</tbody></table>'

    def _vertical_rowspan(self, rows: List[Dict[str, Any]], row_index: int, cell_index: Optional[int]) -> int:
        if cell_index is None:
            return 1
        cell = _cell_at(rows[row_index], cell_index)
        if not cell:
            return 1
        vmerge = cell.get("properties", {}).get("vertical_merge", {})
        if vmerge and vmerge.get("w:val") not in {"restart", None}:
            return 1
        count = 1
        for next_row in rows[row_index + 1:]:
            next_cell = _cell_at(next_row, cell_index)
            next_merge = (next_cell or {}).get("properties", {}).get("vertical_merge", {})
            if next_merge and next_merge.get("w:val") == "continue":
                count += 1
                continue
            break
        return count

    def _render_list(self, elements: List[Dict[str, Any]], start: int, document: Dict[str, Any]) -> Tuple[str, int]:
        items = []
        index = start
        while index < len(elements) and _is_numbered_paragraph(elements[index]):
            para = elements[index]
            num = para.get("properties", {}).get("numbering", {})
            level = int(num.get("level", "0"))
            fmt = self._numbering_format(document, num.get("num_id"), level)
            ordered = fmt not in {"bullet", "none"}
            content = self._render_runs(para.get("runs", []), document) or _html_escape(para.get("text", ""))
            items.append((level, ordered, content, fmt))
            index += 1

        html_parts: List[str] = []
        stack: List[bool] = []
        previous_level = -1
        for level, ordered, content, fmt in items:
            while len(stack) > level + 1:
                tag = "ol" if stack.pop() else "ul"
                html_parts.append(f"</li></{tag}>")
            if len(stack) == level + 1 and previous_level >= level:
                html_parts.append("</li>")
            while len(stack) < level + 1:
                tag = "ol" if ordered else "ul"
                html_parts.append(f'<{tag} class="list-level-{len(stack)} list-format-{_html_attr(fmt)}">')
                stack.append(ordered)
            html_parts.append(f"<li>{content}")
            previous_level = level
        while stack:
            tag = "ol" if stack.pop() else "ul"
            html_parts.append(f"</li></{tag}>")
        return "".join(html_parts), index

    def _numbering_format(self, document: Dict[str, Any], num_id: Optional[str], level: int) -> str:
        numbering = document.get("numbering", {})
        abstract_by_id = {
            item.get("abstract_num_id"): item
            for item in numbering.get("abstract_numbers", [])
        }
        for number in numbering.get("numbers", []):
            if number.get("num_id") != num_id:
                continue
            abstract = abstract_by_id.get(number.get("abstract_num_id"))
            if not abstract:
                return "decimal"
            for lvl in abstract.get("levels", []):
                if str(lvl.get("level")) == str(level):
                    return lvl.get("format") or "decimal"
        return "decimal"

    def _render_header_footer(self, parts: Dict[str, Any], role: str) -> str:
        if not parts:
            return ""
        rendered = []
        for path, part in sorted(parts.items()):
            body = self._render_body(part.get("body", []), {})
            if body:
                rendered.append(f'<section class="docx-{role}" data-part="{_html_attr(path)}">{body}</section>')
        return "\n".join(rendered)

    def _render_notes(self, document: Dict[str, Any]) -> str:
        sections = []
        for key, label in (("footnotes", "脚注"), ("endnotes", "尾注"), ("comments", "批注")):
            items = document.get(key, [])
            if not items:
                continue
            rendered_items = []
            for item in items:
                item_id = item.get("id", "")
                anchor = key[:-1] if key.endswith("s") else key
                body = self._render_body(item.get("body", []), document) or _html_escape(item.get("text", ""))
                meta = ""
                if item.get("author"):
                    meta = f'<span class="note-author">{_html_escape(item["author"])}</span> '
                rendered_items.append(f'<li id="{anchor}-{_html_attr(item_id)}">{meta}{body}</li>')
            sections.append(f'<section class="notes notes-{key}"><h2>{label}</h2><ol>{"".join(rendered_items)}</ol></section>')
        return "\n".join(sections)

    def _paragraph_css(self, props: Dict[str, Any]) -> str:
        css = []
        alignment = props.get("alignment")
        if alignment:
            css.append(f"text-align:{_css_align(alignment)}")
        indent = props.get("indent", {})
        if indent.get("w:left"):
            css.append(f"margin-left:{_twips_to_pt(indent['w:left']):.2f}pt")
        if indent.get("w:right"):
            css.append(f"margin-right:{_twips_to_pt(indent['w:right']):.2f}pt")
        if indent.get("w:firstLine"):
            css.append(f"text-indent:{_twips_to_pt(indent['w:firstLine']):.2f}pt")
        spacing = props.get("spacing", {})
        if spacing.get("w:before"):
            css.append(f"margin-top:{_twips_to_pt(spacing['w:before']):.2f}pt")
        if spacing.get("w:after"):
            css.append(f"margin-bottom:{_twips_to_pt(spacing['w:after']):.2f}pt")
        if spacing.get("w:line"):
            css.append(f"line-height:{max(_twips_to_pt(spacing['w:line']) / 12, 1):.3f}")
        shading = props.get("shading", {})
        if shading.get("w:fill"):
            css.append(f"background-color:#{shading['w:fill']}")
        return ";".join(css)

    def _run_css(self, props: Dict[str, Any]) -> str:
        css = []
        fonts = props.get("fonts", {})
        font_name = fonts.get("w:ascii") or fonts.get("w:hAnsi") or fonts.get("w:eastAsia") or fonts.get("w:cs")
        if font_name:
            css.append(f"font-family:{_css_string(font_name)}")
        size = props.get("size_half_points")
        if size:
            css.append(f"font-size:{float(size) / 2:.2f}pt")
        color = props.get("color", {})
        if color.get("w:val") and color["w:val"] != "auto":
            css.append(f"color:#{color['w:val']}")
        elif color.get("w:themeColor"):
            css.append(f"--docx-theme-color:{_css_string(color['w:themeColor'])}")
        highlight = props.get("highlight", {}).get("w:val")
        if highlight:
            css.append(f"background-color:{_highlight_color(highlight)}")
        shading = props.get("shading", {})
        if shading.get("w:fill"):
            css.append(f"background-color:#{shading['w:fill']}")
        underline = props.get("underline", {})
        if underline:
            style = underline.get("w:val", "single")
            color = underline.get("w:color")
            decoration = "underline"
            if props.get("double_strike"):
                decoration += " line-through"
            css.append(f"text-decoration-line:{decoration}")
            if style != "single":
                css.append(f"text-decoration-style:{_underline_style(style)}")
            if color and color != "auto":
                css.append(f"text-decoration-color:#{color}")
        spacing = props.get("character_spacing", {}).get("w:val")
        if spacing:
            css.append(f"letter-spacing:{_twips_to_pt(spacing):.2f}pt")
        scale = props.get("character_scale", {}).get("w:val")
        if scale and scale != "100":
            css.append(f"display:inline-block;transform:scaleX({float(scale) / 100:.3f});transform-origin:left center")
        position = props.get("position", {}).get("w:val")
        if position:
            css.append(f"position:relative;top:{-_twips_to_pt(position):.2f}pt")
        border = props.get("border", {})
        if border:
            css.append(f"border:{_border_css(border)}")
        if props.get("hidden"):
            css.append("visibility:hidden")
        return ";".join(css)

    def _table_css(self, props: Dict[str, Any]) -> str:
        css = []
        width = props.get("width", {})
        if width.get("w:w") and width.get("w:type") == "pct":
            css.append(f"width:{float(width['w:w']) / 50:.2f}%")
        elif width.get("w:w"):
            css.append(f"width:{_twips_to_pt(width['w:w']):.2f}pt")
        alignment = props.get("alignment")
        if alignment == "center":
            css.append("margin-left:auto;margin-right:auto")
        elif alignment == "right":
            css.append("margin-left:auto")
        shading = props.get("shading", {})
        if shading.get("w:fill"):
            css.append(f"background-color:#{shading['w:fill']}")
        return ";".join(css)

    def _cell_css(self, props: Dict[str, Any]) -> str:
        css = []
        width = props.get("width", {})
        if width.get("w:w") and width.get("w:type") != "auto":
            css.append(f"width:{_twips_to_pt(width['w:w']):.2f}pt")
        shading = props.get("shading", {})
        if shading.get("w:fill"):
            css.append(f"background-color:#{shading['w:fill']}")
        valign = props.get("vertical_align")
        if valign:
            css.append(f"vertical-align:{_css_vertical_align(valign)}")
        return ";".join(css)

    def _page_style(self, document: Dict[str, Any]) -> Dict[str, str]:
        section = (document.get("sections") or [{}])[0].get("properties", {})
        size = section.get("page_size", {})
        margins = section.get("page_margins", {})
        style = {
            "width": f"{_twips_to_pt(size.get('w:w', '12240')):.2f}pt",
            "min_height": f"{_twips_to_pt(size.get('w:h', '15840')):.2f}pt",
            "padding_top": f"{_twips_to_pt(margins.get('w:top', '1440')):.2f}pt",
            "padding_right": f"{_twips_to_pt(margins.get('w:right', '1440')):.2f}pt",
            "padding_bottom": f"{_twips_to_pt(margins.get('w:bottom', '1440')):.2f}pt",
            "padding_left": f"{_twips_to_pt(margins.get('w:left', '1440')):.2f}pt",
        }
        if size.get("w:orient") == "landscape":
            style["orientation"] = "landscape"
        return style

    def _stylesheet(self, page_style: Dict[str, str]) -> str:
        return f"""
    :root {{ color-scheme: light; }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: #e5e7eb;
      color: #111827;
      font-family: "Calibri", "Microsoft YaHei", "SimSun", sans-serif;
      line-height: 1.35;
    }}
    .docx-page {{
      width: {page_style["width"]};
      min-height: {page_style["min_height"]};
      margin: 24px auto;
      padding: {page_style["padding_top"]} {page_style["padding_right"]} {page_style["padding_bottom"]} {page_style["padding_left"]};
      background: #fff;
      box-shadow: 0 8px 28px rgba(15, 23, 42, 0.18);
    }}
    .docx-header {{ border-bottom: 1px solid #d1d5db; margin-bottom: 18pt; color: #4b5563; }}
    .docx-footer {{ border-top: 1px solid #d1d5db; margin-top: 18pt; color: #4b5563; }}
    .paragraph {{ margin: 0 0 6pt; min-height: 1em; }}
    h1, h2, h3, h4, h5, h6 {{ margin: 12pt 0 6pt; line-height: 1.2; }}
    .docx-table {{ border-collapse: collapse; margin: 8pt 0; max-width: 100%; }}
    .docx-table td, .docx-table th {{ border: 1px solid #9ca3af; padding: 4pt 6pt; vertical-align: top; }}
    .docx-image {{ max-width: 100%; height: auto; vertical-align: middle; }}
    .math {{ font-family: "Cambria Math", "Times New Roman", serif; }}
    .vml-shape {{ display: inline-block; border: 1px solid #9ca3af; padding: 3pt; }}
    .tab {{ display: inline-block; width: 2em; }}
    .page-break {{ break-after: page; border-top: 1px dashed #9ca3af; margin: 14pt 0; }}
    .section-break {{ border-top: 2px solid #d1d5db; margin: 16pt 0; }}
    .notes {{ border-top: 1px solid #d1d5db; margin-top: 18pt; font-size: 9pt; }}
    .notes h2 {{ font-size: 12pt; }}
    @media print {{
      body {{ background: #fff; }}
      .docx-page {{ margin: 0; box-shadow: none; }}
    }}
""".strip()

    def _inline_resources(self, html_text: str, base_dir: Path) -> str:
        def replace(match: re.Match[str]) -> str:
            resource = match.group(1)
            path = _safe_join(base_dir, resource)
            if not path.exists():
                return ""
            mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            return f"data:{mime};base64,{encoded}"

        return re.sub(r"md2word-resource:([^\"']+)", replace, html_text)


def _load_metadata(path: str) -> Tuple[Dict[str, Any], Path]:
    p = Path(path)
    if p.is_dir():
        meta_file = p / META_FILENAME
        base_dir = p
    else:
        meta_file = p
        base_dir = p.parent
    return json.loads(meta_file.read_text(encoding="utf-8")), base_dir


def _is_numbered_paragraph(element: Dict[str, Any]) -> bool:
    return element.get("type") == "paragraph" and bool(element.get("properties", {}).get("numbering"))


def _heading_tag(style_id: str) -> str:
    match = re.search(r"heading\s*([1-6])", style_id, re.IGNORECASE)
    if not match:
        match = re.search(r"Heading([1-6])", style_id)
    return f"h{match.group(1)}" if match else "p"


def _cell_at(row: Dict[str, Any], index: int) -> Optional[Dict[str, Any]]:
    for cell in row.get("cells", []):
        if cell.get("index") == index:
            return cell
    return None


def _safe_join(root: Path, relative: str) -> Path:
    parts = Path(*Path(relative).parts)
    if parts.is_absolute() or ".." in parts.parts:
        raise ValueError(f"Unsafe resource path: {relative}")
    return root.joinpath(parts)


def _twips_to_pt(value: Any) -> float:
    try:
        return float(value) / 20.0
    except (TypeError, ValueError):
        return 0.0


def _emu_to_px(value: Any) -> float:
    try:
        return float(value) / 9525.0
    except (TypeError, ValueError):
        return 0.0


def _html_escape(value: Any) -> str:
    return html.escape(str(value or ""), quote=False)


def _html_attr(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def _css_string(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def _css_align(value: str) -> str:
    return {"both": "justify"}.get(value, value)


def _css_vertical_align(value: str) -> str:
    return {"center": "middle"}.get(value, value)


def _highlight_color(value: str) -> str:
    return {
        "black": "#000000",
        "blue": "#0000ff",
        "cyan": "#00ffff",
        "green": "#00ff00",
        "magenta": "#ff00ff",
        "red": "#ff0000",
        "yellow": "#ffff00",
        "white": "#ffffff",
        "darkBlue": "#00008b",
        "darkCyan": "#008b8b",
        "darkGreen": "#006400",
        "darkMagenta": "#8b008b",
        "darkRed": "#8b0000",
        "darkYellow": "#808000",
        "darkGray": "#a9a9a9",
        "lightGray": "#d3d3d3",
    }.get(value, value)


def _underline_style(value: str) -> str:
    if "wave" in value.lower():
        return "wavy"
    if value.lower() in {"double"}:
        return "double"
    if value.lower() in {"dotted", "dash", "dotDash", "dashDotDotHeavy"}:
        return "dotted" if "dot" in value.lower() else "dashed"
    return "solid"


def _border_css(border: Dict[str, Any]) -> str:
    width = _twips_to_pt(border.get("w:sz", "8"))
    color = border.get("w:color", "000000")
    style = "solid" if border.get("w:val", "single") != "none" else "none"
    return f"{max(width / 8, 0.5):.2f}pt {style} #{color}"
