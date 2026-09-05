"""One-call façade: pptx file in, xlsx file out."""

from __future__ import annotations

import os
from typing import Iterable, List, Optional, Tuple

from .exporter import export_to_excel
from .extractor import ExtractOptions, extract_from_presentation
from .hierarchy import build_hierarchy
from .models import ExtractionResult, OrgNode


def default_output_path(pptx_path: str) -> str:
    base, _ = os.path.splitext(pptx_path)
    return base + "_OrgChart.xlsx"


def convert(pptx_path: str,
            excel_path: Optional[str] = None,
            slides: Optional[Iterable[int]] = None,
            include_titles: bool = False,
            use_tables: bool = True,
            smart_text: bool = True,
            infer_geometry: bool = True,
            rtl: bool = True) -> Tuple[str, List[OrgNode], ExtractionResult]:
    """Convert one presentation and return (excel path, nodes, extraction result)."""
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"لم يتم العثور على الملف: {pptx_path}")
    if not pptx_path.lower().endswith((".pptx", ".pptm")):
        raise ValueError("الملف يجب أن يكون بصيغة .pptx (وليس .ppt القديمة).")

    options = ExtractOptions(
        include_titles=include_titles,
        use_tables=use_tables,
        smart_text=smart_text,
        infer_geometry=infer_geometry,
        slides=slides,
    )
    result = extract_from_presentation(pptx_path, options)
    nodes = build_hierarchy(result)

    if not nodes:
        raise ValueError(
            "لم يتم العثور على أي عناصر في العرض التقديمي. "
            "تأكد من أن الشرائح تحتوي على مخطط SmartArt أو مربعات نصية أو جدول."
        )

    output = excel_path or default_output_path(pptx_path)
    if os.path.isdir(output):
        output = os.path.join(output, os.path.basename(default_output_path(pptx_path)))
    directory = os.path.dirname(os.path.abspath(output))
    if directory:
        os.makedirs(directory, exist_ok=True)

    export_to_excel(nodes, result, output, rtl=rtl)
    return output, nodes, result
