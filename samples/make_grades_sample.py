"""ملف تجريبي على شكل هيكل وظائف بدرجات وظيفية.

كل وظيفة = خليتان: خلية صغيرة للدرجة (M5 / M4 / 39 ...) وخلية للمسمى،
وأسفل كل مدير عمود من صفوف الدرجات، بعضها بلا مسمى (وظائف غير مسمّاة).

    python samples/make_grades_sample.py        -> samples/sample_grades.pptx
"""

from __future__ import annotations

import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm, Pt

COLUMNS = [
    (1.5, "M4", "Quality Assurance and Human Resources System Management Director",
     ["39", "38", "37", "36", "35"]),
    (11.0, "M4", "Payroll and Benefits Director", ["39", "38", "37", "36", "35"]),
    (20.5, "M3", "Employees Services Director", ["38", "37", "36", "35"]),
]


def _cell(slide, text: str, left: float, top: float, width: float, height: float = 0.9):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(left), Cm(top),
                                   Cm(width), Cm(height))
    frame = shape.text_frame
    frame.word_wrap = True
    frame.text = text
    frame.paragraphs[0].runs[0].font.size = Pt(9)
    return shape


def _row(slide, grade: str, title: str, left: float, top: float,
         grade_w: float = 1.2, title_w: float = 7.0) -> None:
    _cell(slide, grade, left, top, grade_w)
    if title:
        _cell(slide, title, left + grade_w, top, title_w)


def build(path: str) -> str:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.87), Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    _row(slide, "M5", "General Manager of Human Capital Services", 22.0, 0.6)

    for left, grade, title, grades in COLUMNS:
        _row(slide, grade, title, left, 2.6)
        top = 4.0
        for value in grades:
            _row(slide, value, "", left, top)      # صف درجة بلا مسمى
            top += 1.3

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    prs.save(path)
    return path


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "sample_grades.pptx")
    print("تم إنشاء الملف التجريبي:", build(target))
