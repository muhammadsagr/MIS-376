"""Writes the extracted hierarchy into a formatted Excel workbook."""

from __future__ import annotations

from datetime import datetime
from typing import Dict, List, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .models import ExtractionResult, OrgNode

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
LEVEL_FILLS = ["DDEBF7", "E2EFDA", "FFF2CC", "FCE4D6", "EDEDED", "F2F2F2"]
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# أعمدة وضع الوظائف (الافتراضي): المربع = وظيفة
COLUMNS_POSITIONS: Sequence[tuple] = (
    ("م", 6),
    ("الشريحة", 9),
    ("عنوان الشريحة", 22),
    ("المستوى", 9),
    ("الدرجة الوظيفية", 12),
    ("المسمى الوظيفي", 30),
    ("شاغل الوظيفة (إن وُجد)", 22),
    ("القسم / الإدارة", 22),
    ("الوظيفة الأعلى", 30),
    ("وظائف تابعة مباشرة", 16),
    ("إجمالي الوظائف التابعة", 18),
    ("المسار الوظيفي", 46),
    ("المصدر", 12),
    ("المعرّف", 12),
    ("النص الأصلي", 40),
)

# أعمدة وضع الموظفين: المربع = شخص
COLUMNS_EMPLOYEES: Sequence[tuple] = (
    ("م", 6),
    ("الشريحة", 9),
    ("عنوان الشريحة", 22),
    ("المستوى", 9),
    ("الدرجة الوظيفية", 12),
    ("الاسم", 26),
    ("المسمى الوظيفي", 26),
    ("القسم / الإدارة", 22),
    ("المدير المباشر", 26),
    ("مرؤوسون مباشرون", 16),
    ("إجمالي المرؤوسين", 16),
    ("المسار الوظيفي", 46),
    ("المصدر", 12),
    ("المعرّف", 12),
    ("النص الأصلي", 40),
)

SOURCE_LABEL = {"smartart": "SmartArt", "shape": "أشكال", "table": "جدول"}


def _style_header(ws: Worksheet, widths: Sequence[tuple], rtl: bool) -> None:
    for idx, (title, width) in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=idx, value=title)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = rtl


def _write_main_sheet(ws: Worksheet, nodes: List[OrgNode], by_id: Dict[str, OrgNode],
                      rtl: bool, positions: bool = True) -> None:
    columns = COLUMNS_POSITIONS if positions else COLUMNS_EMPLOYEES
    _style_header(ws, columns, rtl)
    for row_idx, node in enumerate(nodes, start=2):
        parent = by_id.get(node.parent_id) if node.parent_id else None
        values = [
            row_idx - 1,
            node.slide_index,
            node.slide_title,
            node.level,
            node.grade,
            node.title if positions else node.name,
            node.name if positions else node.title,
            node.department,
            parent.display_name if parent else "",
            node.direct_reports,
            node.total_reports,
            node.path,
            SOURCE_LABEL.get(node.source, node.source),
            node.node_id,
            node.raw_text.replace("\n", " / "),
        ]
        fill = PatternFill("solid", fgColor=LEVEL_FILLS[min(node.level - 1, len(LEVEL_FILLS) - 1)])
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = BORDER
            cell.fill = fill
            cell.alignment = Alignment(
                vertical="center", wrap_text=col_idx in (3, 6, 7, 8, 9, 12, 15),
                horizontal="center" if col_idx in (1, 2, 4, 5, 10, 11, 13) else "right" if rtl else "left",
            )
            if col_idx in (6, 7) and node.level == 1:
                cell.font = Font(bold=True)
    if len(nodes):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(nodes) + 1}"


def _write_tree_sheet(ws: Worksheet, nodes: List[OrgNode], rtl: bool,
                      positions: bool = True) -> None:
    second = ("شاغل الوظيفة", 24) if positions else ("المسمى الوظيفي", 28)
    cols = (("المستوى", 9), ("الدرجة", 10), ("الهيكل الشجري", 60), second,
            ("القسم / الإدارة", 22),
            ("وظائف تابعة" if positions else "عدد المرؤوسين", 14))
    _style_header(ws, cols, rtl)
    for row_idx, node in enumerate(nodes, start=2):
        indent = max(node.level - 1, 0)
        label = ("└─ " if indent else "") + (node.display_name or "-")
        ws.cell(row=row_idx, column=1, value=node.level).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=node.grade).alignment = Alignment(horizontal="center")
        cell = ws.cell(row=row_idx, column=3, value=label)
        cell.alignment = Alignment(indent=indent * 2, horizontal="right" if rtl else "left")
        if node.level == 1:
            cell.font = Font(bold=True)
        ws.cell(row=row_idx, column=4, value=node.name if positions else node.title)
        ws.cell(row=row_idx, column=5, value=node.department)
        ws.cell(row=row_idx, column=6, value=node.direct_reports).alignment = \
            Alignment(horizontal="center")
        for col in range(1, 7):
            ws.cell(row=row_idx, column=col).border = BORDER
        if node.level > 1:
            ws.row_dimensions[row_idx].outlineLevel = min(node.level - 1, 7)
    ws.sheet_properties.outlinePr.summaryBelow = False


def _write_summary_sheet(ws: Worksheet, nodes: List[OrgNode], result: ExtractionResult,
                         rtl: bool, positions: bool = True) -> None:
    _style_header(ws, (("البيان", 34), ("القيمة", 58)), rtl)
    levels = {}
    departments = {}
    for node in nodes:
        levels[node.level] = levels.get(node.level, 0) + 1
        if node.department:
            departments[node.department] = departments.get(node.department, 0) + 1

    rows = [
        ("ملف العرض التقديمي", result.source_file),
        ("تاريخ التحويل", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("عدد الشرائح المستخدمة", ", ".join(str(s) for s in result.slides_used()) or "-"),
        ("نوع الهيكل", "هيكل وظائف" if positions else "هيكل موظفين"),
        ("إجمالي عدد الوظائف" if positions else "إجمالي عدد الموظفين", len(nodes)),
        ("عدد المستويات الإدارية", max(levels) if levels else 0),
        ("عدد الوظائف في القمة (بلا وظيفة أعلى)", sum(1 for n in nodes if n.parent_id is None)),
        ("عدد الوظائف بلا وظائف تابعة", sum(1 for n in nodes if n.direct_reports == 0)),
    ]
    for level in sorted(levels):
        rows.append((f"عدد الوظائف في المستوى {level}", levels[level]))
    grades = {}
    for node in nodes:
        if node.grade:
            grades[node.grade] = grades.get(node.grade, 0) + 1
    for grade, count in sorted(grades.items()):
        rows.append((f"عدد الوظائف بالدرجة {grade}", count))
    for dept, count in sorted(departments.items(), key=lambda kv: -kv[1]):
        rows.append((f"القسم: {dept}", count))
    for warning in result.warnings:
        rows.append(("ملاحظة", warning))

    for row_idx, (label, value) in enumerate(rows, start=2):
        c1 = ws.cell(row=row_idx, column=1, value=label)
        c2 = ws.cell(row=row_idx, column=2, value=value)
        c1.font = Font(bold=True)
        for cell in (c1, c2):
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="right" if rtl else "left")


def export_to_excel(nodes: List[OrgNode], result: ExtractionResult, output_path: str,
                    rtl: bool = True, positions: bool = True) -> str:
    """Create the workbook (3 sheets) and return the path it was saved to."""
    wb = Workbook()
    by_id = {n.node_id: n for n in nodes}

    main = wb.active
    main.title = "الهيكل الوظيفي"
    _write_main_sheet(main, nodes, by_id, rtl, positions)

    _write_tree_sheet(wb.create_sheet("العرض الشجري"), nodes, rtl, positions)
    _write_summary_sheet(wb.create_sheet("ملخص وتقرير"), nodes, result, rtl, positions)

    wb.save(output_path)
    return output_path
