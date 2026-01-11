from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any
import xml.etree.ElementTree as ET


def norm_ppd(value: Any) -> str:
    """261-0020 -> 2610020"""
    if value is None:
        return ""
    return re.sub(r"\D+", "", str(value).strip())


def fmt_date(value: Any) -> str:
    """Excel date/datetime -> dd.mm.yyyy, inak string."""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()


def to_str(value: Any) -> str:
    return "" if value is None else str(value).strip()


def to_num_str(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    return str(value).strip()


def indent(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent(child, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i
