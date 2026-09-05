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

# خلية الدرجة الوظيفية: M5 / M4 / G12 / 39 / 38 ...
GRADE_RE = re.compile(r"^(?:[A-Za-z\u0621-\u064a]{0,3}[ .\-]?\d{1,3}|[A-Z]{1,3})$")


def is_grade(text: str) -> bool:
    """هل هذا النص خلية درجة وظيفية (رمز قصير) وليس مسمى وظيفيًا؟"""
    value = (text or "").strip()
    if not value or len(value) > 6 or "\n" in value:
        return False
    return bool(GRADE_RE.match(value))



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


def parse_position_text(raw: str, smart: bool = True) -> Dict[str, str]:
    """وضع الوظائف: المربع يمثّل وظيفة لا شخصًا.

    السطر الأول هو المسمى الوظيفي، والسطر الذي يشبه اسم قسم يذهب إلى
    `department`، وأي اسم شخص (إن وُجد) يوضع في `name` كـ «شاغل الوظيفة».
    """
    lines = split_lines(raw)
    out = {"name": "", "title": "", "department": "", "extra": ""}
    if not lines:
        return out

    if len(lines) == 1:
        pair = _split_inline(lines[0])
        if pair:
            left, right = pair
            if smart and _looks_like_title(right) and not _looks_like_title(left):
                title, other = right, left
            else:
                title, other = left, right
            out["title"] = title
            if _looks_like_department(other):
                out["department"] = other
            else:
                out["name"] = other
        else:
            out["title"] = lines[0]
        return out

    # عدة أسطر: ابحث عن السطر الذي يشبه مسمى وظيفي ليكون الوظيفة
    title_idx = 0
    if smart and not _looks_like_title(lines[0]):
        for idx, line in enumerate(lines[1:], start=1):
            if _looks_like_title(line) and not _looks_like_department(line):
                title_idx = idx
                break

    out["title"] = lines[title_idx]
    rest: List[str] = []
    for idx, line in enumerate(lines):
        if idx == title_idx:
            continue
        if not out["department"] and _looks_like_department(line):
            out["department"] = line
        elif not out["name"] and not _looks_like_title(line):
            out["name"] = line              # شاغل الوظيفة إن كُتب داخل المربع
        else:
            rest.append(line)
    out["extra"] = " | ".join(rest)
    return out
