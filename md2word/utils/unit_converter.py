from __future__ import annotations
from typing import Optional

from docx.shared import Emu, Inches, Cm


def parse_size(value: Optional[str], page_width: Optional[Emu] = None) -> Optional[Emu]:
    if value is None:
        return None

    value = value.strip()

    if value.endswith("px"):
        try:
            pixels = float(value[:-2])
            return Emu(int(pixels * 9525))
        except ValueError:
            return None

    if value.endswith("cm"):
        try:
            return Cm(float(value[:-2]))
        except ValueError:
            return None

    if value.endswith("%"):
        if page_width is not None:
            try:
                return Emu(int(page_width * float(value[:-1]) / 100))
            except ValueError:
                return None
        return None

    if value.endswith("in"):
        try:
            return Inches(float(value[:-2]))
        except ValueError:
            return None

    # Try as inches by default
    try:
        return Inches(float(value))
    except ValueError:
        return None
