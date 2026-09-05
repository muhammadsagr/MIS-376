"""تحويل الهيكل الوظيفي من PowerPoint إلى Excel - برنامج كامل في ملف واحد.
OrgChart PPTX -> XLSX converter (single file edition).

المتطلبات:
    pip install python-pptx openpyxl

الاستخدام:
    python orgchart_converter.py                     -> الواجهة الرسومية
    python orgchart_converter.py chart.pptx          -> تحويل مباشر
    python orgchart_converter.py chart.pptx -o out.xlsx --slides 1-3 --tree

أو كمكتبة داخل كود آخر:
    from orgchart_converter import convert
    excel_path, nodes, result = convert("chart.pptx")

يقرأ البرنامج أربعة أشكال للهياكل داخل الشرائح:
    1. مخططات SmartArt (يُقرأ التسلسل من بيانات المخطط نفسها)
    2. مربعات مرتبطة بخطوط ربط Connectors
    3. مربعات بدون خطوط ربط (يُستنتج التسلسل من مواقعها على الشريحة)
    4. جداول فيها أعمدة: الاسم / المسمى الوظيفي / القسم / المدير المباشر
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import threading
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from pptx import Presentation
from pptx.util import Emu

try:                                    # الواجهة الرسومية اختيارية
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    HAS_TKINTER = True
except Exception:                       # pragma: no cover
    HAS_TKINTER = False

__version__ = "1.0.0"


# ==========================================================================
#  1) نماذج البيانات  -  Data models
# ==========================================================================

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


# ==========================================================================
#  2) تحليل نص المربع  -  Text parsing
# ==========================================================================

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


# ==========================================================================
#  3) قراءة العرض التقديمي  -  Extraction (SmartArt / shapes / tables)
# ==========================================================================

# --------------------------------------------------------------------------- #
# XML namespaces
# --------------------------------------------------------------------------- #
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
R_EMBED = "{%s}dm" % NS["r"]

SKIP_PLACEHOLDERS = {13, 12, 7, 15, 16, 11, 10}  # slide number, footer, date, ...


# --------------------------------------------------------------------------- #
# geometry helpers
# --------------------------------------------------------------------------- #
class Transform:
    """Maps coordinates of shapes nested inside groups back to slide space."""

    def __init__(self, dx: float = 0.0, dy: float = 0.0, sx: float = 1.0, sy: float = 1.0):
        self.dx, self.dy, self.sx, self.sy = dx, dy, sx, sy

    def point(self, x: float, y: float) -> Tuple[float, float]:
        return (self.dx + x * self.sx, self.dy + y * self.sy)

    def box(self, left, top, width, height):
        if left is None or top is None:
            return None, None, None, None
        x, y = self.point(left, top)
        return int(x), int(y), int((width or 0) * self.sx), int((height or 0) * self.sy)

    def child_of_group(self, group) -> "Transform":
        """Combine this transform with the group's own child-offset transform."""
        try:
            xfrm = group._element.find(".//" + _q("a:xfrm"))
            off = xfrm.find(_q("a:off"))
            ext = xfrm.find(_q("a:ext"))
            ch_off = xfrm.find(_q("a:chOff"))
            ch_ext = xfrm.find(_q("a:chExt"))
            ox, oy = int(off.get("x")), int(off.get("y"))
            ex, ey = int(ext.get("cx")), int(ext.get("cy"))
            cox, coy = int(ch_off.get("x")), int(ch_off.get("y"))
            cex, cey = int(ch_ext.get("cx")), int(ch_ext.get("cy"))
            sx = (ex / cex) if cex else 1.0
            sy = (ey / cey) if cey else 1.0
        except Exception:
            return Transform(self.dx, self.dy, self.sx, self.sy)

        gx, gy = self.point(ox, oy)
        return Transform(gx - cox * sx * self.sx, gy - coy * sy * self.sy,
                         self.sx * sx, self.sy * sy)


def _q(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


# --------------------------------------------------------------------------- #
# SmartArt (diagram part)
# --------------------------------------------------------------------------- #
def _diagram_data_part(shape):
    """Return the `diagramData` part behind a SmartArt graphic frame, or None."""
    el = shape._element
    rel_ids = el.find(".//" + _q("dgm:relIds"))
    if rel_ids is None:
        return None
    rid = rel_ids.get(R_EMBED)
    if not rid:
        return None
    part = shape.part
    try:
        return part.related_part(rid)
    except Exception:
        pass
    try:
        return part.rels[rid].target_part
    except Exception:
        return None


def _dgm_point_text(pt) -> str:
    texts = []
    for para in pt.findall(".//" + _q("a:p")):
        runs = [t.text or "" for t in para.findall(".//" + _q("a:t"))]
        texts.append("".join(runs))
    return clean_text("\n".join(texts))


def _extract_smartart(shape, slide_index: int, prefix: str,
                      result: ExtractionResult, smart_text: bool) -> int:
    part = _diagram_data_part(shape)
    if part is None:
        return 0
    try:
        from lxml import etree
        root = etree.fromstring(part.blob)
    except Exception as exc:                                     # pragma: no cover
        result.warnings.append(f"تعذّرت قراءة بيانات SmartArt في الشريحة {slide_index}: {exc}")
        return 0

    model_ids: Dict[str, str] = {}
    added = 0
    for pt in root.findall(".//" + _q("dgm:pt")):
        ptype = pt.get("type") or "node"
        if ptype != "node":
            continue
        model_id = pt.get("modelId")
        text = _dgm_point_text(pt)
        node_id = f"{prefix}sa{added + 1}"
        model_ids[model_id] = node_id
        fields = parse_box_text(text, smart=smart_text)
        result.nodes.append(OrgNode(
            node_id=node_id, slide_index=slide_index, source="smartart",
            raw_text=text, **fields,
        ))
        added += 1

    # parent -> child links, honouring the authored sibling order
    links: List[Tuple[int, str, str]] = []
    for cxn in root.findall(".//" + _q("dgm:cxn")):
        if (cxn.get("type") or "parOf") != "parOf":
            continue
        src, dest = model_ids.get(cxn.get("srcId")), model_ids.get(cxn.get("destId"))
        if not dest or src == dest:
            continue
        order = int(cxn.get("destOrd") or 0)
        if src:
            links.append((order, src, dest))
    for _, src, dest in sorted(links):
        result.edges.append((src, dest))
    return added


# --------------------------------------------------------------------------- #
# plain shapes + connectors
# --------------------------------------------------------------------------- #
def _is_connector(shape) -> bool:
    return shape._element.tag == _q("p:cxnSp")


def _is_group(shape) -> bool:
    return shape._element.tag == _q("p:grpSp")


def _placeholder_type(shape) -> Optional[int]:
    try:
        if shape.is_placeholder:
            return int(shape.placeholder_format.type)
    except Exception:
        pass
    return None


def _connector_ends(shape, tr: Transform):
    """Absolute (x1, y1), (x2, y2) of a connector, honouring flips."""
    left, top, width, height = tr.box(shape.left, shape.top, shape.width, shape.height)
    if left is None:
        return None, None
    flip_h = flip_v = False
    xfrm = shape._element.find(".//" + _q("a:xfrm"))
    if xfrm is not None:
        flip_h = xfrm.get("flipH") in ("1", "true")
        flip_v = xfrm.get("flipV") in ("1", "true")
    x1, x2 = (left + width, left) if flip_h else (left, left + width)
    y1, y2 = (top + height, top) if flip_v else (top, top + height)
    return (x1, y1), (x2, y2)


def _collect_shapes(shapes, slide_index: int, prefix: str, tr: Transform,
                    result: ExtractionResult, opts) -> Tuple[Dict[int, str], List[dict]]:
    """Walk the shape tree; return {drawing id: node id} and the connector list."""
    id_map: Dict[int, str] = {}
    connectors: List[dict] = []
    counter = [len(result.nodes)]

    def walk(container, transform: Transform):
        for shape in container:
            if _is_group(shape):
                walk(shape.shapes, transform.child_of_group(shape))
                continue

            if _is_connector(shape):
                cnv = shape._element.find(".//" + _q("p:cNvCxnSpPr"))
                start = end = None
                if cnv is not None:
                    st, en = cnv.find(_q("a:stCxn")), cnv.find(_q("a:endCxn"))
                    start = int(st.get("id")) if st is not None else None
                    end = int(en.get("id")) if en is not None else None
                p1, p2 = _connector_ends(shape, transform)
                connectors.append({"start": start, "end": end, "p1": p1, "p2": p2})
                continue

            try:
                if not shape.has_text_frame:
                    continue
                text = clean_text(shape.text_frame.text)
            except Exception:
                continue
            if not text:
                continue

            ph = _placeholder_type(shape)
            if ph is not None:
                if ph in SKIP_PLACEHOLDERS:
                    continue
                if ph in (0, 13) and not opts.include_titles:   # TITLE / CENTER_TITLE
                    continue
                if ph in (1, 2, 3) and not opts.include_titles and shape.width and \
                        opts.slide_width and shape.width > opts.slide_width * 0.85:
                    continue

            left, top, width, height = transform.box(shape.left, shape.top,
                                                     shape.width, shape.height)
            # ignore full-slide background frames
            if opts.slide_width and width and height and \
                    width > opts.slide_width * 0.95 and height > opts.slide_height * 0.9:
                continue
            if len(text) > opts.max_text_len:
                continue

            counter[0] += 1
            node_id = f"{prefix}sh{counter[0]}"
            fields = parse_box_text(text, smart=opts.smart_text)
            result.nodes.append(OrgNode(
                node_id=node_id, slide_index=slide_index, source="shape",
                raw_text=text, left=left, top=top, width=width, height=height,
                **fields,
            ))
            try:
                id_map[int(shape.shape_id)] = node_id
            except Exception:
                pass

    walk(shapes, tr)
    return id_map, connectors


def _edges_from_connectors(connectors: List[dict], id_map: Dict[int, str],
                           nodes: List[OrgNode]) -> List[Tuple[str, str]]:
    """Turn connectors into parent/child edges (the higher box is the parent)."""
    edges: List[Tuple[str, str]] = []
    boxes = [n for n in nodes if n.top is not None]

    def nearest_node(point) -> Optional[str]:
        if point is None or not boxes:
            return None
        x, y = point
        best, best_d = None, None
        for n in boxes:
            cx = n.left + (n.width or 0) / 2.0
            cy = n.top + (n.height or 0) / 2.0
            dx = max(n.left - x, 0, x - (n.left + (n.width or 0)))
            dy = max(n.top - y, 0, y - (n.top + (n.height or 0)))
            d = (dx * dx + dy * dy) ** 0.5
            if best_d is None or d < best_d or (d == best_d and abs(cx - x) + abs(cy - y) < 1):
                best, best_d = n, d
        tolerance = max((n.height or 0) for n in boxes) * 0.75 or 1
        return best.node_id if best is not None and best_d is not None and best_d <= tolerance else None

    lookup = {n.node_id: n for n in nodes}
    for c in connectors:
        a = id_map.get(c["start"]) if c["start"] is not None else nearest_node(c["p1"])
        b = id_map.get(c["end"]) if c["end"] is not None else nearest_node(c["p2"])
        if not a or not b or a == b:
            continue
        na, nb = lookup[a], lookup[b]
        if na.top is not None and nb.top is not None and abs(na.top - nb.top) > 1000:
            parent, child = (a, b) if na.top < nb.top else (b, a)
        else:
            parent, child = a, b
        if (parent, child) not in edges:
            edges.append((parent, child))
    return edges


def _edges_from_geometry(nodes: List[OrgNode]) -> List[Tuple[str, str]]:
    """Last resort: rebuild the tree from the position of the boxes on the slide."""
    boxes = [n for n in nodes if n.top is not None]
    if len(boxes) < 2:
        return []

    heights = [n.height or 0 for n in boxes if n.height]
    band = (sum(heights) / len(heights)) if heights else 100000
    band = max(band * 0.8, 1)

    rows: List[List[OrgNode]] = []
    for node in sorted(boxes, key=lambda n: (n.top, n.left)):
        if rows and abs(node.top - rows[-1][0].top) <= band:
            rows[-1].append(node)
        else:
            rows.append([node])

    edges: List[Tuple[str, str]] = []
    for i in range(1, len(rows)):
        parents = rows[i - 1]
        for child in rows[i]:
            cx = child.left + (child.width or 0) / 2.0
            best, best_score = None, None
            for p in parents:
                px = p.left + (p.width or 0) / 2.0
                overlap = min(p.left + (p.width or 0), child.left + (child.width or 0)) \
                    - max(p.left, child.left)
                score = (-overlap, abs(px - cx))
                if best_score is None or score < best_score:
                    best, best_score = p, score
            if best is not None:
                edges.append((best.node_id, child.node_id))
    return edges


# --------------------------------------------------------------------------- #
# tables
# --------------------------------------------------------------------------- #
HEADER_HINTS = {
    "name": ["name", "employee", "الاسم", "اسم", "الموظف"],
    "title": ["title", "position", "job", "role", "المسمى", "الوظيفة", "المنصب", "الوظيفي"],
    "department": ["department", "dept", "division", "unit", "القسم", "الإدارة", "الادارة", "وحدة"],
    "manager": ["manager", "reports to", "supervisor", "parent", "المدير", "يتبع", "المشرف", "الرئيس المباشر"],
}


def _extract_table(shape, slide_index: int, prefix: str,
                   result: ExtractionResult, smart_text: bool) -> int:
    table = shape.table
    rows = [[clean_text(c.text) for c in row.cells] for row in table.rows]
    if len(rows) < 2:
        return 0

    header = [h.lower() for h in rows[0]]
    cols: Dict[str, int] = {}
    for key, hints in HEADER_HINTS.items():
        for idx, cell in enumerate(header):
            if any(h in cell for h in hints):
                cols.setdefault(key, idx)
                break
    body = rows[1:]
    if not cols:                       # no recognisable header -> positional
        cols = {"name": 0}
        if len(rows[0]) > 1:
            cols["title"] = 1
        if len(rows[0]) > 2:
            cols["manager"] = 2
        body = rows

    added = 0
    pending: List[Tuple[str, str]] = []          # (child node id, manager text)
    by_name: Dict[str, str] = {}
    for row in body:
        def cell(key: str) -> str:
            idx = cols.get(key)
            return row[idx] if idx is not None and idx < len(row) else ""

        name, title = cell("name"), cell("title")
        if not (name or title):
            continue
        added += 1
        node_id = f"{prefix}tb{added}"
        result.nodes.append(OrgNode(
            node_id=node_id, slide_index=slide_index, source="table",
            raw_text=" | ".join(v for v in row if v),
            name=name, title=title, department=cell("department"),
        ))
        if name:
            by_name.setdefault(normalise_key(name), node_id)
        manager = cell("manager")
        if manager:
            pending.append((node_id, manager))

    for child_id, manager in pending:
        parent_id = by_name.get(normalise_key(manager))
        if parent_id and parent_id != child_id:
            result.edges.append((parent_id, child_id))
    return added


# --------------------------------------------------------------------------- #
# public entry point
# --------------------------------------------------------------------------- #
class ExtractOptions:
    def __init__(self, include_titles: bool = False, use_tables: bool = True,
                 smart_text: bool = True, infer_geometry: bool = True,
                 max_text_len: int = 300, slides: Optional[Iterable[int]] = None):
        self.include_titles = include_titles
        self.use_tables = use_tables
        self.smart_text = smart_text
        self.infer_geometry = infer_geometry
        self.max_text_len = max_text_len
        self.slides = set(slides) if slides else None
        self.slide_width = 0
        self.slide_height = 0


def extract_from_presentation(path: str, options: Optional[ExtractOptions] = None) -> ExtractionResult:
    """Open `path` and return every chart box found in it plus their links."""
    opts = options or ExtractOptions()
    prs = Presentation(path)
    opts.slide_width = int(prs.slide_width or Emu(9144000))
    opts.slide_height = int(prs.slide_height or Emu(6858000))

    result = ExtractionResult(source_file=path)

    for index, slide in enumerate(prs.slides, start=1):
        if opts.slides and index not in opts.slides:
            continue
        prefix = f"s{index}_"
        before = len(result.nodes)

        # 1) SmartArt diagrams
        for shape in slide.shapes:
            if shape._element.tag == _q("p:graphicFrame") and \
                    shape._element.find(".//" + _q("dgm:relIds")) is not None:
                _extract_smartart(shape, index, prefix, result, opts.smart_text)

        # 2) tables
        if opts.use_tables:
            for shape in slide.shapes:
                try:
                    has_table = shape.has_table
                except Exception:
                    has_table = False
                if has_table:
                    _extract_table(shape, index, prefix, result, opts.smart_text)

        # 3) shapes + connectors
        id_map, connectors = _collect_shapes(slide.shapes, index, prefix,
                                             Transform(), result, opts)
        shape_nodes = [n for n in result.nodes[before:] if n.source == "shape"]
        edges = _edges_from_connectors(connectors, id_map, shape_nodes)
        if not edges and opts.infer_geometry and len(shape_nodes) > 1:
            edges = _edges_from_geometry(shape_nodes)
            if edges:
                result.warnings.append(
                    f"الشريحة {index}: لا توجد روابط (connectors) بين المربعات، "
                    "تم استنتاج التسلسل من مواقع المربعات على الشريحة."
                )
        result.edges.extend(edges)

        # slide title, useful when several charts live in one file
        title = ""
        try:
            if slide.shapes.title is not None:
                title = clean_text(slide.shapes.title.text).replace("\n", " ")
        except Exception:
            pass
        for node in result.nodes[before:]:
            node.slide_title = title

        if len(result.nodes) == before:
            result.warnings.append(f"الشريحة {index}: لم يتم العثور على أي عناصر تنظيمية.")

    return result


# ==========================================================================
#  4) بناء الشجرة  -  Hierarchy
# ==========================================================================

def _assign_parents(result: ExtractionResult) -> Dict[str, Optional[str]]:
    """First edge wins; edges that would create a cycle are dropped."""
    parents: Dict[str, Optional[str]] = {n.node_id: None for n in result.nodes}
    known = set(parents)

    def creates_cycle(parent: str, child: str) -> bool:
        seen, cur = set(), parent
        while cur is not None and cur not in seen:
            if cur == child:
                return True
            seen.add(cur)
            cur = parents.get(cur)
        return False

    for parent, child in result.edges:
        if parent not in known or child not in known:
            continue
        if parents.get(child) is not None or parent == child:
            continue
        if creates_cycle(parent, child):
            result.warnings.append(
                f"تم تجاهل رابط يسبب حلقة مغلقة: {parent} -> {child}")
            continue
        parents[child] = parent
    return parents


def _sort_key(node: OrgNode, order: Dict[str, int]) -> Tuple:
    return (
        node.slide_index,
        node.top if node.top is not None else 0,
        node.left if node.left is not None else 0,
        order.get(node.node_id, 0),
    )


def build_hierarchy(result: ExtractionResult) -> List[OrgNode]:
    """Fill in parent_id / level / path / report counts and return nodes in tree order."""
    nodes = result.nodes
    by_id = {n.node_id: n for n in nodes}
    order = {n.node_id: i for i, n in enumerate(nodes)}
    parents = _assign_parents(result)

    children: Dict[Optional[str], List[OrgNode]] = {}
    for node in nodes:
        node.parent_id = parents.get(node.node_id)
        children.setdefault(node.parent_id, []).append(node)
    for key in children:
        children[key].sort(key=lambda n: _sort_key(n, order))

    ordered: List[OrgNode] = []

    def visit(node: OrgNode, level: int, path: List[str]) -> int:
        node.level = level
        label = node.display_name or node.node_id
        node.path = " > ".join(path + [label])
        ordered.append(node)
        kids = children.get(node.node_id, [])
        node.direct_reports = len(kids)
        total = 0
        for kid in kids:
            total += 1 + visit(kid, level + 1, path + [label])
        node.total_reports = total
        return total

    for root in children.get(None, []):
        visit(root, 1, [])

    # safety net: anything left out (should not happen) is appended as a root
    if len(ordered) != len(nodes):
        for node in nodes:
            if node not in ordered:
                node.parent_id, node.level = None, 1
                node.path = node.display_name
                ordered.append(node)

    roots = [n for n in ordered if n.parent_id is None]
    if len(roots) > 1:
        result.warnings.append(
            f"يوجد {len(roots)} عنصر بدون رئيس مباشر (قد يكون الملف يحتوي أكثر من هيكل)."
        )
    return ordered


def manager_name(node: OrgNode, by_id: Dict[str, OrgNode]) -> str:
    parent = by_id.get(node.parent_id) if node.parent_id else None
    return parent.display_name if parent else ""


# ==========================================================================
#  5) كتابة ملف الإكسل  -  Excel export
# ==========================================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
LEVEL_FILLS = ["DDEBF7", "E2EFDA", "FFF2CC", "FCE4D6", "EDEDED", "F2F2F2"]
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLUMNS: Sequence[tuple] = (
    ("م", 6),
    ("الشريحة", 9),
    ("عنوان الشريحة", 22),
    ("المستوى", 9),
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
                      rtl: bool) -> None:
    _style_header(ws, COLUMNS, rtl)
    for row_idx, node in enumerate(nodes, start=2):
        parent = by_id.get(node.parent_id) if node.parent_id else None
        values = [
            row_idx - 1,
            node.slide_index,
            node.slide_title,
            node.level,
            node.name,
            node.title,
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
                vertical="center", wrap_text=col_idx in (3, 5, 6, 7, 8, 11, 14),
                horizontal="center" if col_idx in (1, 2, 4, 9, 10, 12) else "right" if rtl else "left",
            )
            if col_idx in (5, 6) and node.level == 1:
                cell.font = Font(bold=True)
    if len(nodes):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(nodes) + 1}"


def _write_tree_sheet(ws: Worksheet, nodes: List[OrgNode], rtl: bool) -> None:
    cols = (("المستوى", 9), ("الهيكل الشجري", 60), ("المسمى الوظيفي", 28),
            ("القسم / الإدارة", 22), ("عدد المرؤوسين", 14))
    _style_header(ws, cols, rtl)
    for row_idx, node in enumerate(nodes, start=2):
        indent = max(node.level - 1, 0)
        label = ("└─ " if indent else "") + (node.display_name or "-")
        ws.cell(row=row_idx, column=1, value=node.level).alignment = Alignment(horizontal="center")
        cell = ws.cell(row=row_idx, column=2, value=label)
        cell.alignment = Alignment(indent=indent * 2, horizontal="right" if rtl else "left")
        if node.level == 1:
            cell.font = Font(bold=True)
        ws.cell(row=row_idx, column=3, value=node.title)
        ws.cell(row=row_idx, column=4, value=node.department)
        ws.cell(row=row_idx, column=5, value=node.direct_reports).alignment = \
            Alignment(horizontal="center")
        for col in range(1, 6):
            ws.cell(row=row_idx, column=col).border = BORDER
        if node.level > 1:
            ws.row_dimensions[row_idx].outlineLevel = min(node.level - 1, 7)
    ws.sheet_properties.outlinePr.summaryBelow = False


def _write_summary_sheet(ws: Worksheet, nodes: List[OrgNode], result: ExtractionResult,
                         rtl: bool) -> None:
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
        ("إجمالي عدد الوظائف", len(nodes)),
        ("عدد المستويات الإدارية", max(levels) if levels else 0),
        ("عدد الوظائف في القمة (بدون مدير)", sum(1 for n in nodes if n.parent_id is None)),
        ("عدد الوظائف بدون مرؤوسين", sum(1 for n in nodes if n.direct_reports == 0)),
    ]
    for level in sorted(levels):
        rows.append((f"عدد الوظائف في المستوى {level}", levels[level]))
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
                    rtl: bool = True) -> str:
    """Create the workbook (3 sheets) and return the path it was saved to."""
    wb = Workbook()
    by_id = {n.node_id: n for n in nodes}

    main = wb.active
    main.title = "الهيكل الوظيفي"
    _write_main_sheet(main, nodes, by_id, rtl)

    _write_tree_sheet(wb.create_sheet("العرض الشجري"), nodes, rtl)
    _write_summary_sheet(wb.create_sheet("ملخص وتقرير"), nodes, result, rtl)

    wb.save(output_path)
    return output_path


# ==========================================================================
#  6) الدالة الجامعة  -  convert()
# ==========================================================================

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


# ==========================================================================
#  7) واجهة سطر الأوامر  -  CLI
# ==========================================================================

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
        epilog="مثال:\n  python orgchart_converter.py company.pptx -o company.xlsx --slides 1-3",
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
    parser.add_argument("--ltr", action="store_true",
                        help="إخراج ملف إكسل باتجاه من اليسار لليمين")
    parser.add_argument("--tree", action="store_true", help="طباعة الهيكل في الشاشة بعد التحويل")
    parser.add_argument("-q", "--quiet", action="store_true", help="إخفاء الرسائل")
    parser.add_argument("--gui", action="store_true", help="فتح الواجهة الرسومية")
    return parser


def print_tree(nodes: List[OrgNode]) -> None:
    for node in nodes:
        prefix = "    " * (node.level - 1) + ("└── " if node.level > 1 else "")
        title = f" ({node.title})" if node.title and node.name else ""
        print(f"{prefix}{node.display_name}{title}")


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if args.gui or not args.input:
        if not HAS_TKINTER:                                        # pragma: no cover
            print("الواجهة الرسومية غير متاحة (مكتبة tkinter غير مثبّتة).")
            print("استخدم بدلاً منها: python orgchart_converter.py file.pptx")
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
        )
    except Exception as exc:
        print(f"[خطأ] {exc}", file=sys.stderr)
        return 1

    if args.tree:
        print_tree(nodes)

    if not args.quiet:
        levels = max((n.level for n in nodes), default=0)
        print(f"عدد الوظائف المستخرجة : {len(nodes)}")
        print(f"عدد المستويات الإدارية: {levels}")
        for warning in result.warnings:
            print(f"  - تنبيه: {warning}")
        print(f"تم حفظ الملف في      : {os.path.abspath(output)}")
    return 0


# ==========================================================================
#  8) الواجهة الرسومية  -  Tkinter GUI
# ==========================================================================

class ConverterApp:
    def __init__(self, root, initial_file: Optional[str] = None):
        self.root = root
        root.title("تحويل الهيكل الوظيفي من PowerPoint إلى Excel")
        root.geometry("760x520")
        root.minsize(680, 480)

        self.pptx_var = tk.StringVar(value=initial_file or "")
        self.xlsx_var = tk.StringVar()
        self.slides_var = tk.StringVar()
        self.titles_var = tk.BooleanVar(value=False)
        self.tables_var = tk.BooleanVar(value=True)
        self.smart_var = tk.BooleanVar(value=True)
        self.geom_var = tk.BooleanVar(value=True)
        self.rtl_var = tk.BooleanVar(value=True)

        self._build()
        if initial_file:
            self.xlsx_var.set(default_output_path(initial_file))

    # ---------------------------------------------------------------- layout
    def _build(self) -> None:
        pad = {"padx": 8, "pady": 6}
        frame = ttk.LabelFrame(self.root, text="الملفات")
        frame.pack(fill="x", **pad)

        ttk.Label(frame, text="ملف PowerPoint:").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(frame, textvariable=self.pptx_var, width=60).grid(row=0, column=1, **pad)
        ttk.Button(frame, text="استعراض...", command=self.pick_input).grid(row=0, column=2, **pad)

        ttk.Label(frame, text="ملف Excel الناتج:").grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(frame, textvariable=self.xlsx_var, width=60).grid(row=1, column=1, **pad)
        ttk.Button(frame, text="حفظ باسم...", command=self.pick_output).grid(row=1, column=2, **pad)

        options = ttk.LabelFrame(self.root, text="الخيارات")
        options.pack(fill="x", **pad)
        ttk.Label(options, text="الشرائح (مثال 1,3,5-7):").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(options, textvariable=self.slides_var, width=20).grid(row=0, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="قراءة الجداول", variable=self.tables_var)\
            .grid(row=0, column=2, sticky="w", **pad)
        ttk.Checkbutton(options, text="إدراج عناوين الشرائح", variable=self.titles_var)\
            .grid(row=1, column=0, sticky="w", **pad)
        ttk.Checkbutton(options, text="فصل ذكي للاسم والمسمى", variable=self.smart_var)\
            .grid(row=1, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="استنتاج التسلسل من المواقع", variable=self.geom_var)\
            .grid(row=1, column=2, sticky="w", **pad)
        ttk.Checkbutton(options, text="اتجاه الإكسل من اليمين لليسار", variable=self.rtl_var)\
            .grid(row=2, column=0, sticky="w", **pad)

        actions = ttk.Frame(self.root)
        actions.pack(fill="x", **pad)
        self.convert_btn = ttk.Button(actions, text="تحويل الآن", command=self.start_convert)
        self.convert_btn.pack(side="right", padx=8)
        ttk.Button(actions, text="فتح مجلد الناتج", command=self.open_folder).pack(side="right")

        log_frame = ttk.LabelFrame(self.root, text="النتيجة")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(log_frame, height=12, wrap="word")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------------------------------------------------------------- actions
    def write(self, message: str) -> None:
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title="اختر ملف العرض التقديمي",
            filetypes=[("PowerPoint", "*.pptx *.pptm"), ("All files", "*.*")])
        if path:
            self.pptx_var.set(path)
            self.xlsx_var.set(default_output_path(path))

    def pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="احفظ ملف الإكسل", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.xlsx_var.set(path)

    def open_folder(self) -> None:
        path = self.xlsx_var.get() or self.pptx_var.get()
        folder = os.path.dirname(os.path.abspath(path)) if path else os.getcwd()
        try:
            if os.name == "nt":
                os.startfile(folder)                                # noqa: S606
            else:
                import subprocess
                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            messagebox.showwarning("تنبيه", f"تعذّر فتح المجلد: {exc}")

    def start_convert(self) -> None:
        self.convert_btn.state(["disabled"])
        threading.Thread(target=self._convert, daemon=True).start()

    def _convert(self) -> None:
        try:
            source = self.pptx_var.get().strip()
            if not source:
                messagebox.showerror("خطأ", "الرجاء اختيار ملف PowerPoint أولاً.")
                return
            self.write(f"جاري قراءة: {source}")
            output, nodes, result = convert(
                pptx_path=source,
                excel_path=self.xlsx_var.get().strip() or None,
                slides=_parse_slides(self.slides_var.get().strip()),
                include_titles=self.titles_var.get(),
                use_tables=self.tables_var.get(),
                smart_text=self.smart_var.get(),
                infer_geometry=self.geom_var.get(),
                rtl=self.rtl_var.get(),
            )
            self.xlsx_var.set(output)
            self.write(f"عدد الوظائف المستخرجة: {len(nodes)}")
            self.write(f"عدد المستويات: {max((n.level for n in nodes), default=0)}")
            for warning in result.warnings:
                self.write(f"تنبيه: {warning}")
            self.write(f"تم الحفظ في: {output}")
            messagebox.showinfo("تم بنجاح", f"تم إنشاء الملف:\n{output}")
        except Exception as exc:
            self.write(f"خطأ: {exc}")
            messagebox.showerror("خطأ", str(exc))
        finally:
            self.convert_btn.state(["!disabled"])


def run_gui(initial_file: Optional[str] = None) -> int:
    if not HAS_TKINTER:                                            # pragma: no cover
        raise RuntimeError("مكتبة tkinter غير متاحة على هذا الجهاز.")
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    ConverterApp(root, initial_file)
    root.mainloop()
    return 0


# ==========================================================================
#  9) نقطة التشغيل  -  entry point
# ==========================================================================

if __name__ == "__main__":
    sys.exit(main())
