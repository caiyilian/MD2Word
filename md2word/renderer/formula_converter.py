from __future__ import annotations
import os
import re
from typing import Optional

from lxml import etree

try:
    import latex2mathml.converter as _latex2mathml_converter
except ModuleNotFoundError:
    _latex2mathml_converter = None


_NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_XSLT_PATH = os.path.join(
    os.path.dirname(__file__), "resources", "MML2OMML.XSL",
)
_XSLT_TRANSFORM: Optional[etree.XSLT] = None


def _get_xslt() -> etree.XSLT:
    global _XSLT_TRANSFORM
    if _XSLT_TRANSFORM is None:
        xslt_tree = etree.parse(_XSLT_PATH)
        _XSLT_TRANSFORM = etree.XSLT(xslt_tree)
    return _XSLT_TRANSFORM


def _fix_empty_base(omml_str: str) -> str:
    try:
        root = etree.fromstring(
            f'<root xmlns:m="{_NS_M}">{omml_str}</root>'.encode("utf-8"),
        )
        changed = False
        for nary in root.xpath("//m:nary", namespaces={"m": _NS_M}):
            e_elem = nary.find(f"{{{_NS_M}}}e")
            if e_elem is not None and len(e_elem) == 0:
                # Empty <m:e/> — grab next sibling of nary as base expression
                parent = nary.getparent()
                if parent is not None:
                    siblings = list(parent)
                    idx = siblings.index(nary)
                    if idx + 1 < len(siblings):
                        next_sib = siblings[idx + 1]
                        parent.remove(next_sib)
                        e_elem.append(next_sib)
                        changed = True
        if changed:
            result = etree.tostring(root, encoding="unicode")
            match = re.search(
                r'<m:oMath\b[^>]*>.*?</m:oMath>',
                result, re.DOTALL,
            )
            if match:
                return match.group(0)
        return omml_str
    except Exception:
        return omml_str


def latex_to_omml(latex: str, display: bool = False) -> Optional[str]:
    if _latex2mathml_converter is None:
        return None

    try:
        mathml_str = _latex2mathml_converter.convert(
            latex,
            display="block" if display else "inline",
        )
    except Exception:
        return None

    try:
        transform = _get_xslt()
        wrapper = (
            f'<root xmlns:mml="http://www.w3.org/1998/Math/MathML">'
            f'{mathml_str}</root>'
        )
        tree = etree.fromstring(wrapper.encode("utf-8"))
        result = transform(tree)
        result_str = str(result)
        match = re.search(
            r'(<m:oMath\b[^>]*>.*?</m:oMath>)',
            result_str, re.DOTALL,
        )
        if match:
            omml = match.group(1)
            return _fix_empty_base(omml)
        return None
    except Exception:
        return None
