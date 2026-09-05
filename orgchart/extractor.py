"""Reads a .pptx file and pulls out every org-chart box plus the links between them.

Three kinds of charts are understood:

1. **SmartArt** (Insert > SmartArt > Hierarchy) - the real hierarchy is stored in
   the diagram data part, so parent/child links are read exactly as authored.
2. **Plain shapes + connectors** - boxes are auto-shapes / text boxes and the lines
   between them are connectors; links come from the connector attachments, or from
   the connector end points when the line was never glued to a shape.
3. **Tables** - a table with columns such as Name / Title / Manager.
"""

from __future__ import annotations

from typing import Dict, Iterable, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Emu

from .models import ExtractionResult, OrgNode
from .textparse import clean_text, normalise_key, parse_box_text

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
