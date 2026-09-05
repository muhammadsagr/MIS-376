"""Command line interface: python main.py chart.pptx -o result.xlsx"""

from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional

from .converter import convert
from .models import OrgNode

BANNER = r"""
==========================================================
   تحويل الهيكل الوظيفي من PowerPoint إلى Excel
   OrgChart PPTX  ->  XLSX Converter
==========================================================
"""


def _parse_slides(value: Optional[str]) -> Optional[List[int]]:
    """'1,3,5-7' -> [1, 3, 5, 6, 7]"""
    if not value:
        return None
    pages: List[int] = []
    for chunk in value.replace(" ", "").split(","):
        if not chunk:
            continue
        if "-" in chunk:
            start, end = chunk.split("-", 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(chunk))
    return sorted(set(pages))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="orgchart",
        description="تحويل معلومات الهيكل الوظيفي من ملف PowerPoint إلى ملف Excel.",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="مثال:\n  python main.py company.pptx -o company.xlsx --slides 1-3",
    )
    parser.add_argument("input", nargs="?", help="مسار ملف العرض التقديمي (.pptx)")
    parser.add_argument("-o", "--output", help="مسار ملف الإكسل الناتج (.xlsx)")
    parser.add_argument("-s", "--slides", help="أرقام الشرائح المطلوبة، مثال: 1,3,5-7")
    parser.add_argument("--include-titles", action="store_true",
                        help="إدراج عناوين الشرائح كعناصر في الهيكل")
    parser.add_argument("--no-tables", action="store_true", help="تجاهل الجداول داخل الشرائح")
    parser.add_argument("--no-smart-text", action="store_true",
                        help="عدم استخدام الذكاء في فصل الاسم عن المسمى الوظيفي")
    parser.add_argument("--no-geometry", action="store_true",
                        help="عدم استنتاج التسلسل من مواقع المربعات عند غياب خطوط الربط")
    parser.add_argument("--employees", action="store_true",
                        help="اعتبار المربعات أشخاصًا (موظفين) بدل الوظائف")
    parser.add_argument("--ltr", action="store_true",
                        help="إخراج ملف إكسل باتجاه من اليسار لليمين")
    parser.add_argument("--tree", action="store_true", help="طباعة الهيكل في الشاشة بعد التحويل")
    parser.add_argument("-q", "--quiet", action="store_true", help="إخفاء الرسائل")
    parser.add_argument("--gui", action="store_true", help="فتح الواجهة الرسومية")
    return parser


def print_tree(nodes: List[OrgNode]) -> None:
    for node in nodes:
        prefix = "    " * (node.level - 1) + ("└── " if node.level > 1 else "")
        # في وضع الوظائف يظهر شاغل الوظيفة بين قوسين، وفي وضع الموظفين يظهر المسمى
        secondary = node.name if node.prefer_title else node.title
        extra = f" ({secondary})" if secondary and secondary != node.display_name else ""
        print(f"{prefix}{node.display_name}{extra}")


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if args.gui or not args.input:
        try:
            from .gui import run_gui
        except Exception as exc:                                   # pragma: no cover
            print(f"تعذّر فتح الواجهة الرسومية ({exc}). استخدم: python main.py file.pptx")
            return 2
        return run_gui(args.input)

    if not args.quiet:
        print(BANNER)

    try:
        output, nodes, result = convert(
            pptx_path=args.input,
            excel_path=args.output,
            slides=_parse_slides(args.slides),
            include_titles=args.include_titles,
            use_tables=not args.no_tables,
            smart_text=not args.no_smart_text,
            infer_geometry=not args.no_geometry,
            rtl=not args.ltr,
            positions=not args.employees,
        )
    except Exception as exc:
        print(f"[خطأ] {exc}", file=sys.stderr)
        return 1

    if args.tree:
        print_tree(nodes)

    if not args.quiet:
        levels = max((n.level for n in nodes), default=0)
        label = "عدد الموظفين المستخرجين" if args.employees else "عدد الوظائف المستخرجة"
        print(f"{label} : {len(nodes)}")
        print(f"عدد المستويات الإدارية: {levels}")
        for warning in result.warnings:
            print(f"  - تنبيه: {warning}")
        print(f"تم حفظ الملف في      : {os.path.abspath(output)}")
    return 0
