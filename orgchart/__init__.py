"""OrgChart PPT -> Excel converter package."""

from .models import OrgNode, ExtractionResult
from .extractor import extract_from_presentation
from .hierarchy import build_hierarchy
from .exporter import export_to_excel
from .converter import convert

__all__ = [
    "OrgNode",
    "ExtractionResult",
    "extract_from_presentation",
    "build_hierarchy",
    "export_to_excel",
    "convert",
]

__version__ = "1.0.0"
