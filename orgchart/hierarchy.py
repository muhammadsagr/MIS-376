"""Turns the raw (nodes, edges) pair into a clean, ordered tree."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from .models import ExtractionResult, OrgNode


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
