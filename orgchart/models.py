"""Data models used across the converter."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

EMU_PER_CM = 360000.0


@dataclass
class OrgNode:
    """A single box (position / employee) taken from the org chart."""

    node_id: str
    slide_index: int
    source: str                      # "smartart" | "shape" | "table"
    raw_text: str = ""

    name: str = ""
    title: str = ""
    department: str = ""
    extra: str = ""

    parent_id: Optional[str] = None
    level: int = 0
    path: str = ""
    direct_reports: int = 0
    total_reports: int = 0

    # geometry in EMU (None for table rows / SmartArt without layout info)
    left: Optional[int] = None
    top: Optional[int] = None
    width: Optional[int] = None
    height: Optional[int] = None

    slide_title: str = ""

    @property
    def center(self) -> Optional[Tuple[float, float]]:
        if self.left is None or self.top is None:
            return None
        return (
            self.left + (self.width or 0) / 2.0,
            self.top + (self.height or 0) / 2.0,
        )

    def cm(self, value: Optional[int]) -> Optional[float]:
        return None if value is None else round(value / EMU_PER_CM, 2)

    @property
    def display_name(self) -> str:
        return self.name or self.title or self.raw_text.replace("\n", " ").strip()


@dataclass
class ExtractionResult:
    """Everything pulled out of one presentation before hierarchy is resolved."""

    nodes: List[OrgNode] = field(default_factory=list)
    edges: List[Tuple[str, str]] = field(default_factory=list)   # (parent_id, child_id)
    warnings: List[str] = field(default_factory=list)
    source_file: str = ""

    @property
    def by_id(self) -> Dict[str, OrgNode]:
        return {n.node_id: n for n in self.nodes}

    def slides_used(self) -> List[int]:
        return sorted({n.slide_index for n in self.nodes})
