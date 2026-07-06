from __future__ import annotations

import posixpath
from io import BytesIO
from typing import Any, Dict, List, Optional

from lxml import etree


NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_W10 = "urn:schemas-microsoft-com:office:word"

_PREFIXES = {
    NS_W: "w",
    NS_R: "r",
    NS_REL: "rel",
    NS_WP: "wp",
    NS_A: "a",
    NS_PIC: "pic",
    NS_M: "m",
    NS_CP: "cp",
    NS_DC: "dc",
    NS_DCTERMS: "dcterms",
    NS_EP: "ep",
    NS_VT: "vt",
    NS_C: "c",
    NS_DGM: "dgm",
    NS_V: "v",
    NS_O: "o",
    NS_W10: "w10",
}


def build_semantic_document(
    xml_parts: Dict[str, bytes],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    relationships = _relationship_index(xml_parts)
    document: Dict[str, Any] = {
        "resources": [
            {"package_path": path, "resource": resource}
            for path, resource in sorted(resource_map.items())
        ],
        "relationships": relationships,
    }

    properties = _document_properties(xml_parts)
    if properties:
        document["metadata"] = properties

    settings = _settings_metadata(xml_parts)
    if settings:
        document["settings"] = settings

    styles = _styles_metadata(xml_parts)
    if styles:
        document["styles"] = styles

    numbering = _numbering_metadata(xml_parts)
    if numbering:
        document["numbering"] = numbering

    fonts = _fonts_metadata(xml_parts)
    if fonts:
        document["fonts"] = fonts

    theme = _theme_metadata(xml_parts)
    if theme:
        document["theme"] = theme

    advanced_parts = _advanced_parts_metadata(xml_parts, resource_map, relationships)
    if advanced_parts:
        document.update(advanced_parts)

    root = _parse_part(xml_parts.get("word/document.xml"))
    if root is not None:
        body = root.find(_w("body"))
        if body is not None:
            document["body"] = _body_elements(body, "word/document.xml", relationships, resource_map)
            document["sections"] = [
                _section_properties(sect, "word/document.xml", relationships)
                for sect in body.findall(f".//{_w('sectPr')}")
            ]

    headers: Dict[str, Any] = {}
    footers: Dict[str, Any] = {}
    for path, data in sorted(xml_parts.items()):
        if path.startswith("word/header") and path.endswith(".xml"):
            part_root = _parse_part(data)
            if part_root is not None:
                headers[path] = {
                    "body": _body_elements(part_root, path, relationships, resource_map),
                }
        elif path.startswith("word/footer") and path.endswith(".xml"):
            part_root = _parse_part(data)
            if part_root is not None:
                footers[path] = {
                    "body": _body_elements(part_root, path, relationships, resource_map),
                }
    if headers or footers:
        document["headers_footers"] = _clean({"headers": headers, "footers": footers})

    footnotes = _notes_metadata(xml_parts, "word/footnotes.xml", "footnote", relationships, resource_map)
    if footnotes:
        document["footnotes"] = footnotes
    endnotes = _notes_metadata(xml_parts, "word/endnotes.xml", "endnote", relationships, resource_map)
    if endnotes:
        document["endnotes"] = endnotes
    comments = _comments_metadata(xml_parts, relationships, resource_map)
    if comments:
        document["comments"] = comments

    return _clean(document)


def _document_properties(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    result = {
        "core": _simple_children_metadata(_parse_part(xml_parts.get("docProps/core.xml"))),
        "app": _simple_children_metadata(_parse_part(xml_parts.get("docProps/app.xml"))),
        "custom": _custom_properties_metadata(_parse_part(xml_parts.get("docProps/custom.xml"))),
    }
    return _clean(result)


def _settings_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    root = _parse_part(xml_parts.get("word/settings.xml"))
    if root is None:
        return {}
    return _clean({
        "items": [_simple_element_metadata(child) for child in root],
        "document_protection": _child_attrs(root, "documentProtection"),
        "view": _child_attrs(root, "view"),
        "zoom": _child_attrs(root, "zoom"),
        "default_tab_stop": _child_val(root, "defaultTabStop"),
        "track_revisions": _on_off(root, "trackRevisions"),
        "even_and_odd_headers": _on_off(root, "evenAndOddHeaders"),
        "embed_true_type_fonts": _on_off(root, "embedTrueTypeFonts"),
        "math": _simple_element_metadata(root.find(_w("mathPr"))),
    })


def _styles_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    root = _parse_part(xml_parts.get("word/styles.xml"))
    if root is None:
        return {}

    doc_defaults = root.find(_w("docDefaults"))
    rPr_default = None
    pPr_default = None
    if doc_defaults is not None:
        rPr_default = doc_defaults.find(f"{_w('rPrDefault')}/{_w('rPr')}")
        pPr_default = doc_defaults.find(f"{_w('pPrDefault')}/{_w('pPr')}")

    return _clean({
        "defaults": {
            "run": _run_properties(rPr_default),
            "paragraph": _paragraph_properties(pPr_default),
        },
        "latent_styles": _simple_element_metadata(root.find(_w("latentStyles"))),
        "styles": [
            _style_metadata(style)
            for style in root.findall(_w("style"))
        ],
    })


def _style_metadata(style: etree._Element) -> Dict[str, Any]:
    return _clean({
        "style_id": style.get(_w("styleId")),
        "type": style.get(_w("type")),
        "default": style.get(_w("default")),
        "custom": style.get(_w("customStyle")),
        "name": _child_val(style, "name"),
        "aliases": _child_val(style, "aliases"),
        "based_on": _child_val(style, "basedOn"),
        "next": _child_val(style, "next"),
        "link": _child_val(style, "link"),
        "ui_priority": _child_val(style, "uiPriority"),
        "q_format": _on_off(style, "qFormat"),
        "hidden": _on_off(style, "hidden"),
        "semi_hidden": _on_off(style, "semiHidden"),
        "unhide_when_used": _on_off(style, "unhideWhenUsed"),
        "run_properties": _run_properties(style.find(_w("rPr"))),
        "paragraph_properties": _paragraph_properties(style.find(_w("pPr"))),
        "table_properties": _table_properties(style.find(_w("tblPr"))),
    })


def _numbering_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    root = _parse_part(xml_parts.get("word/numbering.xml"))
    if root is None:
        return {}
    return _clean({
        "abstract_numbers": [
            _abstract_number_metadata(abstract)
            for abstract in root.findall(_w("abstractNum"))
        ],
        "numbers": [
            _number_metadata(number)
            for number in root.findall(_w("num"))
        ],
    })


def _abstract_number_metadata(abstract: etree._Element) -> Dict[str, Any]:
    return _clean({
        "abstract_num_id": abstract.get(_w("abstractNumId")),
        "namespace_id": _child_val(abstract, "nsid"),
        "multi_level_type": _child_val(abstract, "multiLevelType"),
        "template_code": _child_val(abstract, "tmpl"),
        "style_link": _child_val(abstract, "styleLink"),
        "numbering_style_link": _child_val(abstract, "numStyleLink"),
        "levels": [
            _numbering_level_metadata(level)
            for level in abstract.findall(_w("lvl"))
        ],
    })


def _numbering_level_metadata(level: etree._Element) -> Dict[str, Any]:
    return _clean({
        "level": level.get(_w("ilvl")),
        "start": _child_val(level, "start"),
        "format": _child_val(level, "numFmt"),
        "text": _child_val(level, "lvlText"),
        "justification": _child_val(level, "lvlJc"),
        "paragraph_properties": _paragraph_properties(level.find(_w("pPr"))),
        "run_properties": _run_properties(level.find(_w("rPr"))),
    })


def _number_metadata(number: etree._Element) -> Dict[str, Any]:
    return _clean({
        "num_id": number.get(_w("numId")),
        "abstract_num_id": _child_val(number, "abstractNumId"),
        "level_overrides": [
            _clean({
                "level": override.get(_w("ilvl")),
                "start_override": _child_val(override, "startOverride"),
                "level_definition": _numbering_level_metadata(override.find(_w("lvl")))
                if override.find(_w("lvl")) is not None else None,
            })
            for override in number.findall(_w("lvlOverride"))
        ],
    })


def _fonts_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    root = _parse_part(xml_parts.get("word/fontTable.xml"))
    if root is None:
        return {}
    return _clean({
        "fonts": [
            _clean({
                "name": font.get(_w("name")),
                "alternate_name": _child_val(font, "altName"),
                "charset": _child_attrs(font, "charset"),
                "family": _child_val(font, "family"),
                "pitch": _child_val(font, "pitch"),
                "panose": _child_attrs(font, "panose1"),
                "embed_regular": _child_attrs(font, "embedRegular"),
                "embed_bold": _child_attrs(font, "embedBold"),
                "embed_italic": _child_attrs(font, "embedItalic"),
                "embed_bold_italic": _child_attrs(font, "embedBoldItalic"),
            })
            for font in root.findall(_w("font"))
        ],
    })


def _theme_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    theme_path = next((path for path in sorted(xml_parts) if path.startswith("word/theme/") and path.endswith(".xml")), None)
    root = _parse_part(xml_parts.get(theme_path) if theme_path else None)
    if root is None:
        return {}
    elements = root.find(_a("themeElements"))
    return _clean({
        "path": theme_path,
        "name": root.get("name"),
        "colors": _theme_colors(elements.find(_a("clrScheme")) if elements is not None else None),
        "fonts": _theme_fonts(elements.find(_a("fontScheme")) if elements is not None else None),
        "format_scheme": _simple_element_metadata(elements.find(_a("fmtScheme")) if elements is not None else None),
    })


def _theme_colors(color_scheme: Optional[etree._Element]) -> Dict[str, Any]:
    if color_scheme is None:
        return {}
    colors: Dict[str, Any] = {}
    for color in color_scheme:
        color_value = {}
        for child in color:
            color_value = {
                "type": _local_name(child.tag),
                "attributes": _attrs(child),
            }
            break
        colors[_local_name(color.tag)] = _clean(color_value)
    return colors


def _theme_fonts(font_scheme: Optional[etree._Element]) -> Dict[str, Any]:
    if font_scheme is None:
        return {}
    return _clean({
        "name": font_scheme.get("name"),
        "major": _theme_font_collection(font_scheme.find(_a("majorFont"))),
        "minor": _theme_font_collection(font_scheme.find(_a("minorFont"))),
    })


def _theme_font_collection(font_parent: Optional[etree._Element]) -> Dict[str, Any]:
    if font_parent is None:
        return {}
    return _clean({
        "latin": _child_attrs(font_parent, "latin", NS_A),
        "east_asian": _child_attrs(font_parent, "ea", NS_A),
        "complex_script": _child_attrs(font_parent, "cs", NS_A),
    })


def _advanced_parts_metadata(
    xml_parts: Dict[str, bytes],
    resource_map: Dict[str, str],
    relationships: Dict[str, List[Dict[str, Any]]],
) -> Dict[str, Any]:
    return _clean({
        "media": _resource_family(resource_map, "word/media/"),
        "embeddings": _resource_family(resource_map, "word/embeddings/"),
        "vba_project": _resource_item(resource_map, "word/vbaProject.bin"),
        "active_x": _active_x_metadata(xml_parts, resource_map),
        "charts": _charts_metadata(xml_parts, relationships),
        "diagrams": _diagram_metadata(xml_parts),
        "people": _people_metadata(xml_parts),
        "glossary": _glossary_metadata(xml_parts, relationships, resource_map),
        "custom_xml": _custom_xml_metadata(xml_parts),
        "mail_merge": _mail_merge_metadata(xml_parts),
    })


def _resource_family(resource_map: Dict[str, str], prefix: str) -> List[Dict[str, Any]]:
    return [
        _clean({
            "package_path": path,
            "resource": resource,
            "extension": posixpath.splitext(path)[1].lstrip(".").lower(),
        })
        for path, resource in sorted(resource_map.items())
        if path.startswith(prefix)
    ]


def _resource_item(resource_map: Dict[str, str], path: str) -> Dict[str, Any]:
    resource = resource_map.get(path)
    if not resource:
        return {}
    return {
        "package_path": path,
        "resource": resource,
        "extension": posixpath.splitext(path)[1].lstrip(".").lower(),
    }


def _active_x_metadata(
    xml_parts: Dict[str, bytes],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    controls = []
    for path in sorted(xml_parts):
        if path.startswith("word/activeX/") and path.endswith(".xml"):
            root = _parse_part(xml_parts[path])
            controls.append(_clean({
                "package_path": path,
                "root": _simple_element_metadata(root),
            }))
    binaries = _resource_family(resource_map, "word/activeX/")
    return _clean({"controls": controls, "binaries": binaries})


def _charts_metadata(
    xml_parts: Dict[str, bytes],
    relationships: Dict[str, List[Dict[str, Any]]],
) -> List[Dict[str, Any]]:
    charts = []
    for path in sorted(xml_parts):
        if not (path.startswith("word/charts/") and path.endswith(".xml")):
            continue
        root = _parse_part(xml_parts[path])
        if root is None:
            continue
        chart = root.find(f".//{_c('chart')}")
        title = root.find(f".//{_c('title')}")
        plot_area = chart.find(_c("plotArea")) if chart is not None else root.find(f".//{_c('plotArea')}")
        plot_area_tag = _c("plotArea")
        charts.append(_clean({
            "package_path": path,
            "chart_types": sorted({
                _local_name(elem.tag)
                for elem in root.findall(f".//{plot_area_tag}/*")
                if _local_name(elem.tag).endswith("Chart")
            }),
            "title": _text_content(title) if title is not None else None,
            "series_count": len(root.findall(f".//{_c('ser')}")),
            "chart_groups": [
                _chart_group_metadata(child)
                for child in plot_area
                if _local_name(child.tag).endswith("Chart")
            ] if plot_area is not None else [],
            "axes": [
                _chart_axis_metadata(child)
                for child in plot_area
                if _local_name(child.tag) in {"catAx", "dateAx", "valAx", "serAx"}
            ] if plot_area is not None else [],
            "legend": _legend_metadata(chart.find(_c("legend"))) if chart is not None else {},
            "plot_area_format": _simple_element_metadata(plot_area.find(_c("spPr"))) if plot_area is not None else {},
            "relationships": relationships.get(path),
        }))
    return charts


def _chart_group_metadata(group: etree._Element) -> Dict[str, Any]:
    return _clean({
        "type": _local_name(group.tag),
        "grouping": _child_val_ns(group, "grouping", NS_C),
        "bar_direction": _child_val_ns(group, "barDir", NS_C),
        "vary_colors": _child_val_ns(group, "varyColors", NS_C),
        "shape": _child_val_ns(group, "shape", NS_C),
        "style": _child_val_ns(group, "style", NS_C),
        "axis_ids": [
            attrs.get("val")
            for attrs in (_attrs(axis_id) for axis_id in group.findall(_c("axId")))
            if attrs.get("val")
        ],
        "series": [
            _chart_series_metadata(series)
            for series in group.findall(_c("ser"))
        ],
    })


def _chart_series_metadata(series: etree._Element) -> Dict[str, Any]:
    return _clean({
        "index": _child_val_ns(series, "idx", NS_C),
        "order": _child_val_ns(series, "order", NS_C),
        "title": _chart_text_metadata(series.find(_c("tx"))),
        "categories": _chart_data_source_metadata(series.find(_c("cat"))),
        "values": _chart_data_source_metadata(series.find(_c("val"))),
        "x_values": _chart_data_source_metadata(series.find(_c("xVal"))),
        "y_values": _chart_data_source_metadata(series.find(_c("yVal"))),
        "bubble_size": _chart_data_source_metadata(series.find(_c("bubbleSize"))),
        "marker": _simple_element_metadata(series.find(_c("marker"))),
        "data_labels": _simple_element_metadata(series.find(_c("dLbls"))),
        "shape_properties": _simple_element_metadata(series.find(_c("spPr"))),
    })


def _chart_text_metadata(text_parent: Optional[etree._Element]) -> Dict[str, Any]:
    if text_parent is None:
        return {}
    str_ref = text_parent.find(_c("strRef"))
    if str_ref is not None:
        return _clean({
            "formula": _child_text(str_ref, "f", NS_C),
            "points": _chart_cache_points(str_ref.find(_c("strCache"))),
        })
    value = text_parent.find(_c("v"))
    if value is not None:
        return {"value": value.text}
    rich = text_parent.find(_c("rich"))
    if rich is not None:
        return {"text": _text_content(rich)}
    return {}


def _chart_data_source_metadata(data_parent: Optional[etree._Element]) -> Dict[str, Any]:
    if data_parent is None:
        return {}
    for ref_name, cache_name in [
        ("strRef", "strCache"),
        ("numRef", "numCache"),
        ("multiLvlStrRef", "multiLvlStrCache"),
    ]:
        ref = data_parent.find(_c(ref_name))
        if ref is not None:
            return _clean({
                "reference_type": ref_name,
                "formula": _child_text(ref, "f", NS_C),
                "points": _chart_cache_points(ref.find(_c(cache_name))),
            })
    for lit_name in ("strLit", "numLit"):
        literal = data_parent.find(_c(lit_name))
        if literal is not None:
            return _clean({
                "reference_type": lit_name,
                "points": _chart_cache_points(literal),
            })
    return {}


def _chart_cache_points(cache: Optional[etree._Element]) -> List[Dict[str, Any]]:
    if cache is None:
        return []
    points = []
    for point in cache.findall(_c("pt")):
        value = point.find(_c("v"))
        points.append(_clean({
            "index": point.get("idx"),
            "value": value.text if value is not None else None,
        }))
    return points


def _chart_axis_metadata(axis: etree._Element) -> Dict[str, Any]:
    title = axis.find(_c("title"))
    return _clean({
        "type": _local_name(axis.tag),
        "axis_id": _child_val_ns(axis, "axId", NS_C),
        "position": _child_val_ns(axis, "axPos", NS_C),
        "title": _text_content(title) if title is not None else None,
        "scaling": _simple_element_metadata(axis.find(_c("scaling"))),
        "number_format": _child_attrs(axis, "numFmt", NS_C),
        "major_tick_mark": _child_val_ns(axis, "majorTickMark", NS_C),
        "minor_tick_mark": _child_val_ns(axis, "minorTickMark", NS_C),
        "tick_label_position": _child_val_ns(axis, "tickLblPos", NS_C),
        "cross_axis": _child_val_ns(axis, "crossAx", NS_C),
        "crosses": _child_val_ns(axis, "crosses", NS_C),
    })


def _legend_metadata(legend: Optional[etree._Element]) -> Dict[str, Any]:
    if legend is None:
        return {}
    return _clean({
        "position": _child_val_ns(legend, "legendPos", NS_C),
        "overlay": _child_val_ns(legend, "overlay", NS_C),
        "layout": _simple_element_metadata(legend.find(_c("layout"))),
    })


def _diagram_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    diagrams: Dict[str, Any] = {}
    for path in sorted(xml_parts):
        if not (path.startswith("word/diagrams/") and path.endswith(".xml")):
            continue
        root = _parse_part(xml_parts[path])
        if root is None:
            continue
        if path.startswith("word/diagrams/data"):
            diagrams.setdefault("data", []).append(_diagram_data_metadata(path, root))
        elif path.startswith("word/diagrams/layout"):
            diagrams.setdefault("layouts", []).append(_diagram_layout_metadata(path, root))
        elif path.startswith("word/diagrams/quickStyle"):
            diagrams.setdefault("quick_styles", []).append(_diagram_style_metadata(path, root))
        elif path.startswith("word/diagrams/colors"):
            diagrams.setdefault("colors", []).append(_diagram_colors_metadata(path, root))
    return diagrams


def _diagram_data_metadata(path: str, root: etree._Element) -> Dict[str, Any]:
    return _clean({
        "package_path": path,
        "root": _local_name(root.tag),
        "points": [
            _diagram_point_metadata(point)
            for point in root.findall(f".//{_dgm('pt')}")
        ],
        "connections": [
            _clean({
                "model_id": connection.get("modelId"),
                "source_id": connection.get("srcId"),
                "destination_id": connection.get("destId"),
                "source_order": connection.get("srcOrd"),
                "destination_order": connection.get("destOrd"),
                "type": connection.get("type"),
                "attributes": _attrs(connection),
            })
            for connection in root.findall(f".//{_dgm('cxn')}")
        ],
        "background": _simple_element_metadata(root.find(_dgm("bg"))),
        "whole": _simple_element_metadata(root.find(_dgm("whole"))),
    })


def _diagram_point_metadata(point: etree._Element) -> Dict[str, Any]:
    property_set = point.find(_dgm("prSet"))
    text_body = point.find(f"{_dgm('t')}/{_a('txBody')}")
    return _clean({
        "model_id": point.get("modelId"),
        "type": point.get("type"),
        "connection_id": point.get("cxnId"),
        "property_set": _attrs(property_set),
        "text": _text_content(text_body) if text_body is not None else None,
        "shape_properties": _simple_element_metadata(point.find(f"{_dgm('spPr')}")),
        "text_body": _simple_element_metadata(text_body),
    })


def _diagram_layout_metadata(path: str, root: etree._Element) -> Dict[str, Any]:
    return _clean({
        "package_path": path,
        "root": _local_name(root.tag),
        "unique_id": root.get("uniqueId"),
        "min_version": root.get("minVer"),
        "title": _diagram_title(root),
        "description": _diagram_description(root),
        "categories": _diagram_categories(root),
        "layout_nodes": [
            _diagram_layout_node_metadata(node)
            for node in root.findall(f".//{_dgm('layoutNode')}")
        ],
    })


def _diagram_layout_node_metadata(node: etree._Element) -> Dict[str, Any]:
    return _clean({
        "name": node.get("name"),
        "style_label": node.get("styleLbl"),
        "move_with": node.get("moveWith"),
        "children_order": node.get("chOrder"),
        "algorithm": _simple_element_metadata(node.find(_dgm("alg"))),
        "shape": _simple_element_metadata(node.find(_dgm("shape"))),
        "presentation_of": _simple_element_metadata(node.find(_dgm("presOf"))),
        "constraints": [_simple_element_metadata(constraint) for constraint in node.findall(f"{_dgm('constrLst')}/{_dgm('constr')}")],
        "rules": [_simple_element_metadata(rule) for rule in node.findall(f"{_dgm('ruleLst')}/{_dgm('rule')}")],
        "variables": [_simple_element_metadata(variable) for variable in node.findall(f"{_dgm('varLst')}/*")],
    })


def _diagram_style_metadata(path: str, root: etree._Element) -> Dict[str, Any]:
    return _clean({
        "package_path": path,
        "root": _local_name(root.tag),
        "unique_id": root.get("uniqueId"),
        "min_version": root.get("minVer"),
        "title": _diagram_title(root),
        "description": _diagram_description(root),
        "categories": _diagram_categories(root),
        "style_labels": [
            _clean({
                "name": style_label.get("name"),
                "scene_3d": _simple_element_metadata(style_label.find(_dgm("scene3d"))),
                "shape_3d": _simple_element_metadata(style_label.find(_dgm("sp3d"))),
                "text_properties": _simple_element_metadata(style_label.find(_dgm("txPr"))),
                "style": _simple_element_metadata(style_label.find(_dgm("style"))),
            })
            for style_label in root.findall(f".//{_dgm('styleLbl')}")
        ],
    })


def _diagram_colors_metadata(path: str, root: etree._Element) -> Dict[str, Any]:
    color_labels = list(root.findall(_dgm("styleLbl"))) + list(root.findall(f".//{_dgm('colorsDef')}/{_dgm('styleLbl')}"))
    return _clean({
        "package_path": path,
        "root": _local_name(root.tag),
        "unique_id": root.get("uniqueId"),
        "min_version": root.get("minVer"),
        "title": _diagram_title(root),
        "description": _diagram_description(root),
        "categories": _diagram_categories(root),
        "color_labels": [
            _clean({
                "name": color_label.get("name"),
                "fill_colors": [
                    _fill_metadata(fill)
                    for fill in color_label.findall(f".//{_dgm('fillClrLst')}/*")
                ],
                "line_colors": [
                    _fill_metadata(fill)
                    for fill in color_label.findall(f".//{_dgm('linClrLst')}/*")
                ],
                "effect_colors": [
                    _fill_metadata(fill)
                    for fill in color_label.findall(f".//{_dgm('effectClrLst')}/*")
                ],
                "text_line_colors": [
                    _fill_metadata(fill)
                    for fill in color_label.findall(f".//{_dgm('txLinClrLst')}/*")
                ],
                "text_fill_colors": [
                    _fill_metadata(fill)
                    for fill in color_label.findall(f".//{_dgm('txFillClrLst')}/*")
                ],
            })
            for color_label in color_labels
        ],
    })


def _diagram_title(root: etree._Element) -> Optional[str]:
    title = root.find(f"{_dgm('title')}")
    return title.get("val") if title is not None else None


def _diagram_description(root: etree._Element) -> Optional[str]:
    desc = root.find(f"{_dgm('desc')}")
    return desc.get("val") if desc is not None else None


def _diagram_categories(root: etree._Element) -> List[Dict[str, Any]]:
    return [
        _clean({
            "type": category.get("type"),
            "priority": category.get("pri"),
        })
        for category in root.findall(f"{_dgm('catLst')}/{_dgm('cat')}")
    ]


def _people_metadata(xml_parts: Dict[str, bytes]) -> List[Dict[str, Any]]:
    root = _parse_part(xml_parts.get("word/people.xml"))
    if root is None:
        return []
    people = []
    for person in root:
        people.append(_clean({
            "name": person.get(_w("name")) or person.get("name"),
            "id": person.get(_w("id")) or person.get("id"),
            "attributes": _attrs(person),
            "children": [_simple_element_metadata(child) for child in person],
        }))
    return people


def _glossary_metadata(
    xml_parts: Dict[str, bytes],
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    path = "word/glossary/document.xml"
    root = _parse_part(xml_parts.get(path))
    if root is None:
        return {}
    body = root.find(_w("body"))
    doc_parts = []
    for doc_part in root.findall(f".//{_w('docPart')}"):
        doc_parts.append(_clean({
            "name": _child_val(doc_part.find(_w("docPartPr")), "name") if doc_part.find(_w("docPartPr")) is not None else None,
            "category": _simple_element_metadata(doc_part.find(f"{_w('docPartPr')}/{_w('category')}")),
            "types": [_simple_element_metadata(part_type) for part_type in doc_part.findall(f"{_w('docPartPr')}/{_w('types')}/{_w('type')}")],
            "body": _body_elements(doc_part.find(_w("docPartBody")), path, relationships, resource_map)
            if doc_part.find(_w("docPartBody")) is not None else [],
        }))
    return _clean({
        "body": _body_elements(body, path, relationships, resource_map) if body is not None else [],
        "doc_parts": doc_parts,
    })


def _custom_xml_metadata(xml_parts: Dict[str, bytes]) -> List[Dict[str, Any]]:
    items = []
    for path in sorted(xml_parts):
        if path.startswith("customXml/") and path.endswith(".xml"):
            root = _parse_part(xml_parts[path])
            items.append(_clean({
                "package_path": path,
                "root": _local_name(root.tag) if root is not None else None,
                "attributes": _attrs(root),
                "text": _text_content(root) if root is not None else None,
            }))
    return items


def _mail_merge_metadata(xml_parts: Dict[str, bytes]) -> Dict[str, Any]:
    root = _parse_part(xml_parts.get("word/settings.xml"))
    if root is None:
        return {}
    mail_merge = root.find(_w("mailMerge"))
    return _simple_element_metadata(mail_merge)


def _parse_part(data: Optional[bytes]) -> Optional[etree._Element]:
    if data is None:
        return None
    parser = etree.XMLParser(
        resolve_entities=False,
        remove_blank_text=False,
        remove_comments=False,
        huge_tree=True,
    )
    return etree.parse(BytesIO(data), parser).getroot()


def _simple_children_metadata(root: Optional[etree._Element]) -> Dict[str, Any]:
    if root is None:
        return {}
    result: Dict[str, Any] = {}
    for child in root:
        key = _local_name(child.tag)
        value = _simple_element_metadata(child)
        if key in result:
            existing = result[key]
            if not isinstance(existing, list):
                result[key] = [existing]
            result[key].append(value)
        else:
            result[key] = value
    return _clean(result)


def _simple_element_metadata(element: Optional[etree._Element]) -> Dict[str, Any]:
    if element is None:
        return {}
    children = [_simple_element_metadata(child) for child in element]
    return _clean({
        "name": _local_name(element.tag),
        "attributes": _attrs(element),
        "text": element.text.strip() if element.text and element.text.strip() else None,
        "children": children,
    })


def _custom_properties_metadata(root: Optional[etree._Element]) -> List[Dict[str, Any]]:
    if root is None:
        return []
    properties = []
    for prop in root:
        value = None
        value_type = None
        for child in prop:
            value_type = _local_name(child.tag)
            value = child.text
            break
        properties.append(_clean({
            "name": prop.get("name"),
            "pid": prop.get("pid"),
            "fmtid": prop.get("fmtid"),
            "type": value_type,
            "value": value,
        }))
    return properties


def _relationship_index(xml_parts: Dict[str, bytes]) -> Dict[str, List[Dict[str, Any]]]:
    result: Dict[str, List[Dict[str, Any]]] = {}
    for path, data in xml_parts.items():
        if not path.endswith(".rels"):
            continue
        source = _relationship_source_part(path)
        if source is None:
            continue
        root = _parse_part(data)
        if root is None:
            continue
        relationships: List[Dict[str, Any]] = []
        for rel in root.findall(_rel("Relationship")):
            rel_id = rel.get("Id")
            target = rel.get("Target")
            target_mode = rel.get("TargetMode")
            package_path = None
            if target and target_mode != "External":
                package_path = _relationship_target_path(source, target)
            relationships.append(_clean({
                "id": rel_id,
                "type": rel.get("Type"),
                "target": target,
                "target_mode": target_mode,
                "package_path": package_path,
            }))
        if relationships:
            result[source] = relationships
    return result


def _relationship_source_part(rels_path: str) -> Optional[str]:
    if rels_path == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in rels_path or not rels_path.endswith(".rels"):
        return None
    base, rel_name = rels_path.split(marker, 1)
    return posixpath.join(base, rel_name[:-5])


def _relationship_target_path(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    source_dir = posixpath.dirname(source_part)
    if not source_dir:
        return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join(source_dir, target))


def _body_elements(
    container: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    elements: List[Dict[str, Any]] = []
    for child in container:
        if child.tag == _w("p"):
            elements.append(_paragraph_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("tbl"):
            elements.append(_table_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("sectPr"):
            elements.append({
                "type": "section_properties",
                "properties": _section_properties(child, part_path, relationships),
            })
        elif child.tag == _w("sdt"):
            elements.append(_sdt_metadata(child, part_path, relationships, resource_map))
        elif _is_revision_tag(child.tag):
            elements.append(_revision_metadata(child, part_path, relationships, resource_map))
    return elements


def _paragraph_metadata(
    paragraph: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    runs: List[Dict[str, Any]] = []
    markers: List[Dict[str, Any]] = []
    for child in paragraph:
        if child.tag == _w("r"):
            runs.append(_run_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("hyperlink"):
            runs.append(_hyperlink_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("sdt"):
            runs.append(_sdt_metadata(child, part_path, relationships, resource_map))
        elif _is_revision_tag(child.tag):
            runs.append(_revision_metadata(child, part_path, relationships, resource_map))
        elif child.tag in {
            _w("commentRangeStart"), _w("commentRangeEnd"),
            _w("bookmarkStart"), _w("bookmarkEnd"),
        }:
            markers.append({
                "type": _local_name(child.tag),
                "attributes": _attrs(child),
            })
    return _clean({
        "type": "paragraph",
        "text": _text_content(paragraph),
        "properties": _paragraph_properties(paragraph.find(_w("pPr"))),
        "runs": runs,
        "markers": markers,
        "fields": _paragraph_fields(paragraph),
    })


def _run_metadata(
    run: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    breaks = [_attrs(br) for br in run.findall(_w("br"))]
    drawings = [
        drawing
        for drawing_elem in run.findall(_w("drawing"))
        for drawing in _drawing_metadata(drawing_elem, part_path, relationships, resource_map)
    ]
    footnote_refs = [_attrs(ref) for ref in run.findall(_w("footnoteReference"))]
    endnote_refs = [_attrs(ref) for ref in run.findall(_w("endnoteReference"))]
    comment_refs = [_attrs(ref) for ref in run.findall(_w("commentReference"))]
    math = [_math_metadata(math_elem) for math_elem in run.findall(_m("oMath"))]
    field_chars = [_attrs(field) for field in run.findall(_w("fldChar"))]
    instr_text = "".join(text.text or "" for text in run.findall(_w("instrText")))
    deleted_text = "".join(text.text or "" for text in run.findall(_w("delText")))
    vml = [
        vml_item
        for container in list(run.findall(_w("pict"))) + list(run.findall(_w("object")))
        for vml_item in _vml_container_metadata(container, part_path, relationships, resource_map)
    ]
    return _clean({
        "type": "run",
        "text": _run_text(run),
        "deleted_text": deleted_text,
        "properties": _run_properties(run.find(_w("rPr"))),
        "breaks": breaks,
        "tabs": len(run.findall(_w("tab"))) or None,
        "drawings": drawings,
        "footnote_references": footnote_refs,
        "endnote_references": endnote_refs,
        "comment_references": comment_refs,
        "math": math,
        "field_chars": field_chars,
        "field_instruction": instr_text,
        "vml": vml,
    })


def _hyperlink_metadata(
    hyperlink: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    rel_id = hyperlink.get(_r("id"))
    return _clean({
        "type": "hyperlink",
        "text": _text_content(hyperlink),
        "anchor": hyperlink.get(_w("anchor")),
        "tooltip": hyperlink.get(_w("tooltip")),
        "relationship": _relationship_lookup(relationships, part_path, rel_id),
        "runs": [
            _run_metadata(run, part_path, relationships, resource_map)
            for run in hyperlink.findall(_w("r"))
        ],
    })


def _sdt_metadata(
    sdt: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    content = sdt.find(_w("sdtContent"))
    body = _body_elements(content, part_path, relationships, resource_map) if content is not None else []
    runs = []
    if content is not None:
        for child in content:
            if child.tag == _w("r"):
                runs.append(_run_metadata(child, part_path, relationships, resource_map))
            elif child.tag == _w("hyperlink"):
                runs.append(_hyperlink_metadata(child, part_path, relationships, resource_map))
            elif _is_revision_tag(child.tag):
                runs.append(_revision_metadata(child, part_path, relationships, resource_map))
    return _clean({
        "type": "structured_document_tag",
        "text": _text_content(sdt),
        "properties": _sdt_properties(sdt.find(_w("sdtPr"))),
        "body": body,
        "runs": runs,
    })


def _sdt_properties(sdtPr: Optional[etree._Element]) -> Dict[str, Any]:
    if sdtPr is None:
        return {}
    control = _sdt_control_metadata(sdtPr)
    return _clean({
        "alias": _child_val(sdtPr, "alias"),
        "tag": _child_val(sdtPr, "tag"),
        "id": _child_val(sdtPr, "id"),
        "lock": _child_attrs(sdtPr, "lock"),
        "placeholder": _simple_element_metadata(sdtPr.find(_w("placeholder"))),
        "data_binding": _child_attrs(sdtPr, "dataBinding"),
        "temporary": _on_off(sdtPr, "temporary"),
        "showing_placeholder": _on_off(sdtPr, "showingPlcHdr"),
        "control": control,
    })


def _sdt_control_metadata(sdtPr: etree._Element) -> Dict[str, Any]:
    control_names = {
        "text", "richText", "picture", "date", "dropDownList", "comboBox",
        "checkBox", "repeatingSection", "repeatingSectionItem", "group",
        "citation", "bibliography", "docPartObj", "docPartList",
    }
    for child in sdtPr:
        local = _local_name(child.tag)
        if local in control_names:
            return _clean({
                "type": local,
                "attributes": _attrs(child),
                "children": [_simple_element_metadata(grandchild) for grandchild in child],
            })
    return {}


def _revision_metadata(
    revision: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    runs = []
    body = []
    for child in revision:
        if child.tag == _w("r"):
            runs.append(_run_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("p"):
            body.append(_paragraph_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("tbl"):
            body.append(_table_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("sdt"):
            body.append(_sdt_metadata(child, part_path, relationships, resource_map))
        elif _is_revision_tag(child.tag):
            body.append(_revision_metadata(child, part_path, relationships, resource_map))
    return _clean({
        "type": "revision",
        "kind": _local_name(revision.tag),
        "id": revision.get(_w("id")),
        "author": revision.get(_w("author")),
        "date": revision.get(_w("date")),
        "attributes": _attrs(revision),
        "text": _text_content(revision),
        "runs": runs,
        "body": body,
    })


def _paragraph_fields(paragraph: etree._Element) -> List[Dict[str, Any]]:
    fields = []
    current: Optional[Dict[str, Any]] = None
    in_result = False
    for run in paragraph.findall(_w("r")):
        for child in run:
            if child.tag == _w("fldChar"):
                fld_type = child.get(_w("fldCharType"))
                if fld_type == "begin":
                    current = {"type": "field", "instruction": "", "result": "", "chars": [_attrs(child)]}
                    in_result = False
                elif current is not None:
                    current.setdefault("chars", []).append(_attrs(child))
                    if fld_type == "separate":
                        in_result = True
                    elif fld_type == "end":
                        fields.append(_clean(current))
                        current = None
                        in_result = False
            elif current is not None and child.tag == _w("instrText") and child.text:
                current["instruction"] += child.text
            elif current is not None and in_result and child.tag == _w("t") and child.text:
                current["result"] += child.text
    return fields


def _paragraph_properties(pPr: Optional[etree._Element]) -> Dict[str, Any]:
    if pPr is None:
        return {}
    return _clean({
        "style": _child_val(pPr, "pStyle"),
        "alignment": _child_val(pPr, "jc"),
        "numbering": _numbering_properties(pPr.find(_w("numPr"))),
        "indent": _child_attrs(pPr, "ind"),
        "spacing": _child_attrs(pPr, "spacing"),
        "shading": _child_attrs(pPr, "shd"),
        "borders": _border_properties(pPr.find(_w("pBdr"))),
        "tabs": [_attrs(tab) for tab in pPr.findall(f"{_w('tabs')}/{_w('tab')}")],
        "keep_next": _on_off(pPr, "keepNext"),
        "keep_lines": _on_off(pPr, "keepLines"),
        "page_break_before": _on_off(pPr, "pageBreakBefore"),
        "widow_control": _on_off(pPr, "widowControl"),
        "outline_level": _child_val(pPr, "outlineLvl"),
        "bidi": _on_off(pPr, "bidi"),
        "text_direction": _child_val(pPr, "textDirection"),
    })


def _run_properties(rPr: Optional[etree._Element]) -> Dict[str, Any]:
    if rPr is None:
        return {}
    return _clean({
        "style": _child_val(rPr, "rStyle"),
        "bold": _on_off(rPr, "b"),
        "italic": _on_off(rPr, "i"),
        "underline": _underline_properties(rPr.find(_w("u"))),
        "strike": _on_off(rPr, "strike"),
        "double_strike": _on_off(rPr, "dstrike"),
        "vertical_align": _child_val(rPr, "vertAlign"),
        "fonts": _child_attrs(rPr, "rFonts"),
        "size_half_points": _child_val(rPr, "sz"),
        "size_cs_half_points": _child_val(rPr, "szCs"),
        "color": _child_attrs(rPr, "color"),
        "highlight": _child_attrs(rPr, "highlight"),
        "shading": _child_attrs(rPr, "shd"),
        "character_spacing": _child_attrs(rPr, "spacing"),
        "character_scale": _child_attrs(rPr, "w"),
        "position": _child_attrs(rPr, "position"),
        "effect": _child_val(rPr, "effect"),
        "border": _child_attrs(rPr, "bdr"),
        "outline": _on_off(rPr, "outline"),
        "shadow": _on_off(rPr, "shadow"),
        "imprint": _on_off(rPr, "imprint"),
        "emboss": _on_off(rPr, "emboss"),
        "caps": _on_off(rPr, "caps"),
        "small_caps": _on_off(rPr, "smallCaps"),
        "hidden": _on_off(rPr, "vanish"),
        "language": _child_attrs(rPr, "lang"),
    })


def _underline_properties(underline: Optional[etree._Element]) -> Dict[str, Any]:
    if underline is None:
        return {}
    attrs = _attrs(underline)
    if "w:val" not in attrs:
        attrs["w:val"] = "single"
    return attrs


def _numbering_properties(numPr: Optional[etree._Element]) -> Dict[str, Any]:
    if numPr is None:
        return {}
    return _clean({
        "level": _child_val(numPr, "ilvl"),
        "num_id": _child_val(numPr, "numId"),
    })


def _table_metadata(
    table: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    rows = []
    for row_idx, row in enumerate(table.findall(_w("tr"))):
        cells = []
        for cell_idx, cell in enumerate(row.findall(_w("tc"))):
            cells.append(_clean({
                "index": cell_idx,
                "text": _text_content(cell),
                "properties": _cell_properties(cell.find(_w("tcPr"))),
                "body": _body_elements(cell, part_path, relationships, resource_map),
            }))
        rows.append(_clean({
            "index": row_idx,
            "properties": _row_properties(row.find(_w("trPr"))),
            "cells": cells,
        }))
    return _clean({
        "type": "table",
        "properties": _table_properties(table.find(_w("tblPr"))),
        "grid": [_attrs(col) for col in table.findall(f"{_w('tblGrid')}/{_w('gridCol')}")],
        "rows": rows,
    })


def _table_properties(tblPr: Optional[etree._Element]) -> Dict[str, Any]:
    if tblPr is None:
        return {}
    return _clean({
        "style": _child_val(tblPr, "tblStyle"),
        "width": _child_attrs(tblPr, "tblW"),
        "alignment": _child_val(tblPr, "jc"),
        "indent": _child_attrs(tblPr, "tblInd"),
        "borders": _border_properties(tblPr.find(_w("tblBorders"))),
        "shading": _child_attrs(tblPr, "shd"),
        "layout": _child_attrs(tblPr, "tblLayout"),
        "look": _child_attrs(tblPr, "tblLook"),
        "cell_margins": _margins_properties(tblPr.find(_w("tblCellMar"))),
    })


def _row_properties(trPr: Optional[etree._Element]) -> Dict[str, Any]:
    if trPr is None:
        return {}
    return _clean({
        "height": _child_attrs(trPr, "trHeight"),
        "cant_split": _on_off(trPr, "cantSplit"),
        "table_header": _on_off(trPr, "tblHeader"),
        "grid_before": _child_val(trPr, "gridBefore"),
        "grid_after": _child_val(trPr, "gridAfter"),
    })


def _cell_properties(tcPr: Optional[etree._Element]) -> Dict[str, Any]:
    if tcPr is None:
        return {}
    vmerge = tcPr.find(_w("vMerge"))
    return _clean({
        "width": _child_attrs(tcPr, "tcW"),
        "grid_span": _child_val(tcPr, "gridSpan"),
        "vertical_merge": (_attrs(vmerge) or {"w:val": "continue"}) if vmerge is not None else None,
        "shading": _child_attrs(tcPr, "shd"),
        "borders": _border_properties(tcPr.find(_w("tcBorders"))),
        "margins": _margins_properties(tcPr.find(_w("tcMar"))),
        "vertical_align": _child_val(tcPr, "vAlign"),
        "text_direction": _child_val(tcPr, "textDirection"),
    })


def _drawing_metadata(
    drawing: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    result = []
    for container in list(drawing.findall(_wp("inline"))) + list(drawing.findall(_wp("anchor"))):
        blip = container.find(f".//{_a('blip')}")
        embed_id = blip.get(_r("embed")) if blip is not None else None
        link_id = blip.get(_r("link")) if blip is not None else None
        rel_id = embed_id or link_id
        relationship = _relationship_lookup(relationships, part_path, rel_id)
        package_path = relationship.get("package_path") if relationship else None
        wrap = _wrap_metadata(container)
        picture = _picture_metadata(container, part_path, relationships)
        result.append(_clean({
            "type": _local_name(container.tag),
            "relationship": relationship,
            "resource": resource_map.get(package_path) if package_path else None,
            "extent": _child_attrs(container, "extent", NS_WP),
            "effect_extent": _child_attrs(container, "effectExtent", NS_WP),
            "doc_properties": _child_attrs(container, "docPr", NS_WP),
            "distances": {
                key: value
                for key, value in _attrs(container).items()
                if key in {"distT", "distB", "distL", "distR", "wp:distT", "wp:distB", "wp:distL", "wp:distR"}
            },
            "wrap": wrap,
            "position": {
                "horizontal": _position_metadata(container.find(_wp("positionH"))),
                "vertical": _position_metadata(container.find(_wp("positionV"))),
                "simple": _child_attrs(container, "simplePos", NS_WP),
            },
            "picture": picture,
            "transform": picture.get("transform"),
            "crop": picture.get("crop"),
            "effects": picture.get("effects"),
        }))
    return result


def _picture_metadata(
    container: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
) -> Dict[str, Any]:
    pic = container.find(f".//{_pic('pic')}")
    if pic is None:
        return {}

    nv_pic_pr = pic.find(_pic("nvPicPr"))
    c_nv_pr = nv_pic_pr.find(_pic("cNvPr")) if nv_pic_pr is not None else None
    blip_fill = pic.find(_pic("blipFill"))
    sp_pr = pic.find(_pic("spPr"))
    transform = sp_pr.find(_a("xfrm")) if sp_pr is not None else None
    line = sp_pr.find(_a("ln")) if sp_pr is not None else None
    preset_geometry = sp_pr.find(_a("prstGeom")) if sp_pr is not None else None

    hyperlink = None
    if c_nv_pr is not None:
        hlink = c_nv_pr.find(_a("hlinkClick"))
        if hlink is not None:
            rel_id = hlink.get(_r("id"))
            hyperlink = _clean({
                "relationship": _relationship_lookup(relationships, part_path, rel_id),
                "tooltip": hlink.get("tooltip"),
                "action": hlink.get("action"),
                "attributes": _attrs(hlink),
            })

    return _clean({
        "non_visual_properties": _attrs(c_nv_pr),
        "hyperlink": hyperlink,
        "crop": _child_attrs(blip_fill, "srcRect", NS_A),
        "transform": _transform_metadata(transform),
        "geometry": _attrs(preset_geometry),
        "line": _line_metadata(line),
        "effects": _effects_metadata(sp_pr),
    })


def _transform_metadata(transform: Optional[etree._Element]) -> Dict[str, Any]:
    if transform is None:
        return {}
    return _clean({
        "attributes": _attrs(transform),
        "offset": _child_attrs(transform, "off", NS_A),
        "extent": _child_attrs(transform, "ext", NS_A),
        "rotation": transform.get("rot"),
        "flip_horizontal": _bool_attr(transform.get("flipH")),
        "flip_vertical": _bool_attr(transform.get("flipV")),
    })


def _line_metadata(line: Optional[etree._Element]) -> Dict[str, Any]:
    if line is None:
        return {}
    return _clean({
        "attributes": _attrs(line),
        "solid_fill": _fill_metadata(line.find(_a("solidFill"))),
        "dash": _child_attrs(line, "prstDash", NS_A),
        "head_end": _child_attrs(line, "headEnd", NS_A),
        "tail_end": _child_attrs(line, "tailEnd", NS_A),
    })


def _effects_metadata(parent: Optional[etree._Element]) -> Dict[str, Any]:
    if parent is None:
        return {}
    effects: Dict[str, Any] = {}
    for container_name in ("effectLst", "effectDag"):
        container = parent.find(_a(container_name))
        if container is None:
            continue
        effects[container_name] = [
            _clean({
                "type": _local_name(effect.tag),
                "attributes": _attrs(effect),
                "fill": _first_fill_metadata(effect),
                "children": [_simple_element_metadata(child) for child in effect],
            })
            for effect in container
        ]
    return _clean(effects)


def _first_fill_metadata(parent: etree._Element) -> Dict[str, Any]:
    for child in parent:
        if _local_name(child.tag).endswith("Fill") or _local_name(child.tag) in {"srgbClr", "schemeClr", "prstClr"}:
            return _fill_metadata(child)
    return {}


def _fill_metadata(fill: Optional[etree._Element]) -> Dict[str, Any]:
    if fill is None:
        return {}
    colors = []
    for child in fill:
        if _local_name(child.tag).endswith("Clr"):
            colors.append(_clean({
                "type": _local_name(child.tag),
                "attributes": _attrs(child),
                "children": [_simple_element_metadata(grandchild) for grandchild in child],
            }))
    return _clean({
        "type": _local_name(fill.tag),
        "attributes": _attrs(fill),
        "colors": colors,
    })


def _bool_attr(value: Optional[str]) -> Optional[bool]:
    if value is None:
        return None
    return value.lower() in {"1", "true", "on"}


def _vml_container_metadata(
    container: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    items = []
    for child in container:
        if _is_vml_shape_tag(child.tag):
            items.append(_vml_shape_metadata(child, part_path, relationships, resource_map))
        elif child.tag == _w("drawing"):
            items.extend(_drawing_metadata(child, part_path, relationships, resource_map))
    return items


def _vml_shape_metadata(
    shape: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> Dict[str, Any]:
    textbox = shape.find(_v("textbox"))
    txbx_content = textbox.find(_w("txbxContent")) if textbox is not None else None
    nested_shapes = [
        _vml_shape_metadata(child, part_path, relationships, resource_map)
        for child in shape
        if _is_vml_shape_tag(child.tag)
    ]
    image_data = shape.find(_v("imagedata"))
    rel_id = image_data.get(_r("id")) if image_data is not None else None
    relationship = _relationship_lookup(relationships, part_path, rel_id)
    package_path = relationship.get("package_path") if relationship else None
    return _clean({
        "type": "vml_shape",
        "shape_type": _local_name(shape.tag),
        "attributes": _attrs(shape),
        "style": _style_string_to_dict(shape.get("style")),
        "fill": _simple_element_metadata(shape.find(_v("fill"))),
        "stroke": _simple_element_metadata(shape.find(_v("stroke"))),
        "shadow": _simple_element_metadata(shape.find(_v("shadow"))),
        "path": _simple_element_metadata(shape.find(_v("path"))),
        "textbox": {
            "attributes": _attrs(textbox),
            "text": _text_content(txbx_content) if txbx_content is not None else None,
            "body": _body_elements(txbx_content, part_path, relationships, resource_map)
            if txbx_content is not None else [],
        },
        "image": {
            "relationship": relationship,
            "resource": resource_map.get(package_path) if package_path else None,
            "attributes": _attrs(image_data),
        },
        "children": nested_shapes,
    })


def _is_vml_shape_tag(qname: str) -> bool:
    return qname in {
        _v("shape"),
        _v("rect"),
        _v("roundrect"),
        _v("oval"),
        _v("line"),
        _v("polyline"),
        _v("curve"),
        _v("group"),
        _v("textbox"),
    }


def _style_string_to_dict(style: Optional[str]) -> Dict[str, str]:
    if not style:
        return {}
    result = {}
    for item in style.split(";"):
        if ":" not in item:
            continue
        key, value = item.split(":", 1)
        key = key.strip()
        if key:
            result[key] = value.strip()
    return result


def _wrap_metadata(container: etree._Element) -> Dict[str, Any]:
    for child in container:
        name = _local_name(child.tag)
        if name.startswith("wrap"):
            return _clean({"type": name, "attributes": _attrs(child)})
    return {}


def _position_metadata(position: Optional[etree._Element]) -> Dict[str, Any]:
    if position is None:
        return {}
    return _clean({
        "attributes": _attrs(position),
        "align": _child_text(position, "align", NS_WP),
        "offset": _child_text(position, "posOffset", NS_WP),
    })


def _section_properties(
    sectPr: etree._Element,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
) -> Dict[str, Any]:
    return _clean({
        "type": _child_val(sectPr, "type"),
        "page_size": _child_attrs(sectPr, "pgSz"),
        "page_margins": _child_attrs(sectPr, "pgMar"),
        "columns": _child_attrs(sectPr, "cols"),
        "page_numbers": _child_attrs(sectPr, "pgNumType"),
        "title_page": _on_off(sectPr, "titlePg"),
        "headers": _section_refs(sectPr, "headerReference", part_path, relationships),
        "footers": _section_refs(sectPr, "footerReference", part_path, relationships),
    })


def _section_refs(
    sectPr: etree._Element,
    tag_name: str,
    part_path: str,
    relationships: Dict[str, List[Dict[str, Any]]],
) -> List[Dict[str, Any]]:
    refs = []
    for ref in sectPr.findall(_w(tag_name)):
        rel_id = ref.get(_r("id"))
        refs.append(_clean({
            "type": ref.get(_w("type")),
            "relationship": _relationship_lookup(relationships, part_path, rel_id),
        }))
    return refs


def _notes_metadata(
    xml_parts: Dict[str, bytes],
    path: str,
    tag_name: str,
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    root = _parse_part(xml_parts.get(path))
    if root is None:
        return []
    notes = []
    for note in root.findall(_w(tag_name)):
        notes.append(_clean({
            "id": note.get(_w("id")),
            "type": note.get(_w("type")),
            "text": _text_content(note),
            "body": _body_elements(note, path, relationships, resource_map),
        }))
    return notes


def _comments_metadata(
    xml_parts: Dict[str, bytes],
    relationships: Dict[str, List[Dict[str, Any]]],
    resource_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    path = "word/comments.xml"
    root = _parse_part(xml_parts.get(path))
    if root is None:
        return []
    comments = []
    for comment in root.findall(_w("comment")):
        comments.append(_clean({
            "id": comment.get(_w("id")),
            "author": comment.get(_w("author")),
            "initials": comment.get(_w("initials")),
            "date": comment.get(_w("date")),
            "attributes": _attrs(comment),
            "text": _text_content(comment),
            "body": _body_elements(comment, path, relationships, resource_map),
        }))
    return comments


def _math_metadata(math_elem: etree._Element) -> Dict[str, Any]:
    return _clean({
        "type": _local_name(math_elem.tag),
        "text": _text_content(math_elem),
    })


def _relationship_lookup(
    relationships: Dict[str, List[Dict[str, Any]]],
    part_path: str,
    rel_id: Optional[str],
) -> Dict[str, Any]:
    if not rel_id:
        return {}
    for relationship in relationships.get(part_path, []):
        if relationship.get("id") == rel_id:
            return relationship
    return {"id": rel_id}


def _border_properties(parent: Optional[etree._Element]) -> Dict[str, Any]:
    if parent is None:
        return {}
    return _clean({
        _local_name(child.tag): _attrs(child)
        for child in parent
    })


def _margins_properties(parent: Optional[etree._Element]) -> Dict[str, Any]:
    if parent is None:
        return {}
    return _clean({
        _local_name(child.tag): _attrs(child)
        for child in parent
    })


def _child_attrs(parent: Optional[etree._Element], tag_name: str, namespace: str = NS_W) -> Dict[str, str]:
    if parent is None:
        return {}
    child = parent.find(_qn(namespace, tag_name))
    if child is None:
        return {}
    return _attrs(child)


def _child_val(parent: Optional[etree._Element], tag_name: str) -> Optional[str]:
    attrs = _child_attrs(parent, tag_name)
    return attrs.get("w:val") or attrs.get("val")


def _child_val_ns(parent: Optional[etree._Element], tag_name: str, namespace: str) -> Optional[str]:
    attrs = _child_attrs(parent, tag_name, namespace)
    return attrs.get("val") or attrs.get(f"{_PREFIXES.get(namespace)}:val")


def _child_text(parent: etree._Element, tag_name: str, namespace: str) -> Optional[str]:
    child = parent.find(_qn(namespace, tag_name))
    return child.text if child is not None else None


def _on_off(parent: Optional[etree._Element], tag_name: str) -> Optional[bool]:
    if parent is None:
        return None
    child = parent.find(_w(tag_name))
    if child is None:
        return None
    value = child.get(_w("val")) or child.get("val")
    if value is None:
        return True
    return value.lower() not in {"0", "false", "off", "none"}


def _run_text(run: etree._Element) -> str:
    parts = []
    for child in run:
        if child.tag == _w("t") and child.text:
            parts.append(child.text)
        elif child.tag == _w("delText") and child.text:
            parts.append(child.text)
        elif child.tag == _w("tab"):
            parts.append("\t")
        elif child.tag == _w("br"):
            parts.append("\n")
    return "".join(parts)


def _text_content(element: etree._Element) -> str:
    return "".join(text for text in element.itertext() if text)


def _attrs(element: Optional[etree._Element]) -> Dict[str, str]:
    if element is None:
        return {}
    return {
        _display_qname(key): value
        for key, value in element.attrib.items()
    }


def _display_qname(qname: str) -> str:
    if qname.startswith("{"):
        namespace, local = qname[1:].split("}", 1)
        prefix = _PREFIXES.get(namespace)
        return f"{prefix}:{local}" if prefix else f"{{{namespace}}}{local}"
    return qname


def _local_name(qname: str) -> str:
    if qname.startswith("{"):
        return qname.split("}", 1)[1]
    return qname


def _is_revision_tag(qname: str) -> bool:
    return qname in {
        _w("ins"),
        _w("del"),
        _w("moveFrom"),
        _w("moveTo"),
    }


def _clean(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            key: cleaned
            for key, item in value.items()
            if (cleaned := _clean(item)) not in (None, {}, [])
        }
    if isinstance(value, list):
        return [cleaned for item in value if (cleaned := _clean(item)) not in (None, {}, [])]
    return value


def _qn(namespace: str, tag_name: str) -> str:
    return f"{{{namespace}}}{tag_name}"


def _w(tag_name: str) -> str:
    return _qn(NS_W, tag_name)


def _r(tag_name: str) -> str:
    return _qn(NS_R, tag_name)


def _rel(tag_name: str) -> str:
    return _qn(NS_REL, tag_name)


def _wp(tag_name: str) -> str:
    return _qn(NS_WP, tag_name)


def _a(tag_name: str) -> str:
    return _qn(NS_A, tag_name)


def _m(tag_name: str) -> str:
    return _qn(NS_M, tag_name)


def _c(tag_name: str) -> str:
    return _qn(NS_C, tag_name)


def _v(tag_name: str) -> str:
    return _qn(NS_V, tag_name)


def _pic(tag_name: str) -> str:
    return _qn(NS_PIC, tag_name)


def _dgm(tag_name: str) -> str:
    return _qn(NS_DGM, tag_name)
