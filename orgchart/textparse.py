"""Heuristics that turn the free text inside a chart box into structured fields."""

from __future__ import annotations

import re
from typing import Dict, List

# Words that make a line look like a job title rather than a person's name.
TITLE_KEYWORDS = [
    "ceo", "cfo", "coo", "cto", "cio", "chief", "president", "vp",
    "vice president", "director", "manager", "supervisor", "head",
    "officer", "lead", "coordinator", "specialist", "analyst",
    "engineer", "assistant", "secretary", "accountant", "administrator",
    "board", "chairman", "consultant", "team leader", "intern",
    "مدير", "المدير", "رئيس", "الرئيس", "نائب", "مشرف", "المشرف", "قائد",
    "مسؤول", "مساعد", "أمين", "محاسب", "مهندس", "أخصائي", "منسق", "محلل",
    "موظف", "سكرتير", "عضو", "مجلس",
]

DEPARTMENT_KEYWORDS = [
    "department", "dept", "division", "unit", "section", "team", "branch",
    "قسم", "القسم", "إدارة", "الإدارة", "ادارة", "الادارة", "وحدة", "شعبة", "فرع", "فريق",
]

# Separators used when a whole record is squeezed into a single line.
INLINE_SEPARATORS = [" - ", " – ", " — ", " | ", " / ", " :: ", " , ", "، ", "؛ "]

_WS = re.compile(r"[ \t ]+")


def clean_text(value: str) -> str:
    """Normalise whitespace but keep line breaks."""
    if not value:
        return ""
    lines = [_WS.sub(" ", ln).strip(" \r\t") for ln in value.replace("\v", "\n").splitlines()]
    return "\n".join(ln for ln in lines if ln.strip())


def split_lines(value: str) -> List[str]:
    return [ln.strip() for ln in clean_text(value).splitlines() if ln.strip()]


def _looks_like_title(line: str) -> bool:
    low = line.strip().lower()
    return any(kw in low for kw in TITLE_KEYWORDS)


def _looks_like_department(line: str) -> bool:
    low = line.strip().lower()
    return any(kw in low for kw in DEPARTMENT_KEYWORDS)


def _split_inline(line: str):
    for sep in INLINE_SEPARATORS:
        if sep in line:
            left, right = line.split(sep, 1)
            left, right = left.strip(" -–—|/,،؛:"), right.strip(" -–—|/,،؛:")
            if left and right:
                return left, right
    # "Name (Title)" form
    m = re.match(r"^(.{2,}?)\s*[\(﴾](.+?)[\)﴿]\s*$", line)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return None


def parse_box_text(raw: str, smart: bool = True) -> Dict[str, str]:
    """Split the text of one box into name / title / department / extra.

    Rules (in order):
      * several lines  -> line1 = name, line2 = title, a department-looking
        line goes to `department`, the rest is kept in `extra`.
      * one line with a separator ("Ali Ahmed - Sales Manager") -> name + title.
      * one line only  -> treated as a name, unless it reads like a job title
        (``smart=True``), in which case it is stored as the title.
    """
    lines = split_lines(raw)
    out = {"name": "", "title": "", "department": "", "extra": ""}
    if not lines:
        return out

    if len(lines) == 1:
        pair = _split_inline(lines[0])
        if pair:
            out["name"], out["title"] = pair
        elif smart and _looks_like_title(lines[0]) and not _looks_like_department(lines[0]):
            out["title"] = lines[0]
        elif smart and _looks_like_department(lines[0]):
            out["department"] = lines[0]
        else:
            out["name"] = lines[0]
        return out

    rest: List[str] = []
    out["name"] = lines[0]
    for line in lines[1:]:
        if not out["department"] and _looks_like_department(line):
            out["department"] = line
        elif not out["title"]:
            out["title"] = line
        else:
            rest.append(line)

    # "Sales Manager / Ali" style first line: swap when line 1 is clearly a title
    if smart and not out["title"] and _looks_like_title(out["name"]):
        out["title"], out["name"] = out["name"], ""

    out["extra"] = " | ".join(rest)
    return out


def normalise_key(value: str) -> str:
    """Key used when matching a manager name written by hand in a table."""
    value = (value or "").strip().lower()
    value = re.sub(r"[ً-ْ]", "", value)          # Arabic diacritics
    value = value.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    value = value.replace("ة", "ه").replace("ى", "ي")
    return _WS.sub(" ", re.sub(r"[^\w؀-ۿ ]+", " ", value)).strip()
