from __future__ import annotations
from typing import Optional

import yaml
from docx.shared import Pt, RGBColor, Inches, Emu


DEFAULT_STYLE_CONFIG = {
    "body": {
        "font": "等线",
        "size": 12,
        "color": None,
        "spacing": 1.15,
    },
    "headings": {
        "h1": {"size": 22, "bold": True, "color": "1F3864"},
        "h2": {"size": 18, "bold": True, "color": "2F5496"},
        "h3": {"size": 15, "bold": True, "color": "2F5496"},
        "h4": {"size": 13, "bold": True, "color": None},
        "h5": {"size": 12, "bold": True, "color": None},
        "h6": {"size": 12, "bold": True, "italic": True, "color": None},
    },
    "code": {
        "font": "Consolas",
        "size": 9,
        "color": "333333",
        "bg_color": "F2F2F2",
    },
    "table": {
        "header_bg": "E8E8E8",
        "border": True,
    },
    "page": {
        "margin_top": None,
        "margin_bottom": None,
        "margin_left": None,
        "margin_right": None,
    },
}


def load_style_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        user = yaml.safe_load(f) or {}
    merged = _merge_config(DEFAULT_STYLE_CONFIG, user)
    return merged


def _merge_config(base: dict, override: dict) -> dict:
    result = base.copy()
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _merge_config(result[k], v)
        else:
            result[k] = v
    return result


def parse_color(hex_str: Optional[str]) -> Optional[RGBColor]:
    if not hex_str:
        return None
    hex_str = hex_str.lstrip("#")
    try:
        return RGBColor(
            int(hex_str[0:2], 16),
            int(hex_str[2:4], 16),
            int(hex_str[4:6], 16),
        )
    except (ValueError, IndexError):
        return None
