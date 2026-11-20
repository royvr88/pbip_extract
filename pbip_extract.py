#!/usr/bin/env python3
"""
pbip_extract.py - documents a Power BI Project (.pbip) as one Markdown file.

Covers the report (pages, layout, visuals, titles, text boxes, buttons, slicers
and the filters on report / page / visual level) and the semantic model (tables,
columns, measures, relationships, RLS, Power Query sources, lineage).

Usage:
    python pbip_extract.py <project>.pbip [-o docs.md] [--rowcounts counts.txt]

Requirements: Python 3.8+, standard library only. Works on Windows, macOS, Linux.
"""

import argparse
import json
import re
import sys
import textwrap
from collections import Counter
from pathlib import Path
from urllib.parse import unquote as url_unquote
from datetime import datetime
from zoneinfo import ZoneInfo

class ExtractError(Exception):
    """A problem the user can fix (wrong path, no model found, ...)."""


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


# ---------------------------------------------------------------------------
# Sanitiser - masks organisation-specific infrastructure values in M / DAX.
# The query structure itself (functions, steps, column names) is left intact.
# ---------------------------------------------------------------------------

_ID_KEYS = {
    "sqlendpoint": "FABRIC_SQL_ENDPOINT", "sql endpoint": "FABRIC_SQL_ENDPOINT",
    "workspaceid": "WORKSPACE_ID", "groupid": "WORKSPACE_ID",
    "lakehouseid": "LAKEHOUSE_ID", "warehouseid": "WAREHOUSE_ID",
    "datasetid": "DATASET_ID", "driveid": "SHAREPOINT_DRIVE_ID",
    "itemid": "SHAREPOINT_ITEM_ID", "siteid": "SHAREPOINT_SITE_ID",
}
_SHAREPOINT_RE = re.compile(r"https?://[a-z0-9\-]+\.sharepoint\.com(/sites/[^/\"')\s,\]]+)?([^\"')\s,\]]*)", re.I)
_FABRIC_SQL_RE = re.compile(r"\b[a-z0-9]{6,}\.datawarehouse\.fabric\.microsoft\.com\b", re.I)
_PBI_WORKSPACE_RE = re.compile(r"(powerbi://[^/\"]+/v1\.0/[^/\"]+/)[^\"]+", re.I)
# The value is limited to identifier characters (no spaces/quotes) so that a business
# column that happens to be called e.g. WarehouseId inside a SQL string isn't mangled.
_ID_RE = re.compile(r"\b(" + "|".join(map(re.escape, _ID_KEYS)) + r')(\s*[:=]\s*)"?[\w.\-]+"?', re.I)
# Only a schema assignment on its own line; [Schema="dbo",Item="x"] navigation is harmless.
_SCHEMA_RE = re.compile(r'(?im)^(\s*schema\s*[:=]\s*)"?[\w.\-]+"?')


def sanitize(text):
    if not text:
        return text

    def sharepoint(m):
        site = "/sites/[SITE]" if m.group(1) else ""
        return "https://[TENANT].sharepoint.com" + site + url_unquote(m.group(2))

    text = _SHAREPOINT_RE.sub(sharepoint, text)
    text = _FABRIC_SQL_RE.sub("[FABRIC_SQL_ENDPOINT]", text)
    text = _PBI_WORKSPACE_RE.sub(lambda m: m.group(1) + "[WORKSPACE]", text)
    text = _ID_RE.sub(lambda m: m.group(1) + m.group(2) + "[" + _ID_KEYS[m.group(1).lower()] + "]", text)
    return _SCHEMA_RE.sub(lambda m: m.group(1) + "[SCHEMA]", text)


# ---------------------------------------------------------------------------
# TMDL - one generic indentation-based parser for every .tmdl file.
#
#   /// description                    -> description of the next object
#   type name [= expression]           -> object (table, column, measure, ...)
#   key: value                         -> property
#   key                                -> boolean property (isHidden)
#   key = expression                   -> property holding an expression (source =)
#
# A multi-line expression is either fenced with ``` or indented at least two
# levels deeper than the line that owns it (its properties are one level deeper).
# ---------------------------------------------------------------------------

def parse_tmdl(text):
    root = {"type": "root", "name": "", "value": "", "props": {}, "children": [], "description": ""}
    stack = [(-1, root)]
    docs = []
    capture = None  # [target dict, key, owner indent, lines, fenced]

    def close():
        target, key, _, lines, _ = capture
        body = textwrap.dedent("\n".join(lines[1:])).strip("\n")
        target[key] = (lines[0] + "\n" + body).strip() if body.strip() else lines[0]

    def node(kind, name, parent, indent):
        new = {"type": kind, "name": name, "value": "", "props": {}, "children": [],
               "description": "\n".join(docs).strip()}
        del docs[:]
        parent["children"].append(new)
        stack.append((indent, new))
        return new

    def start(target, key, indent, inline):
        fenced = inline == "```"
        return [target, key, indent, ["" if fenced else inline], fenced]

    for raw in text.splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        indent = len(line) - len(line.lstrip("\t"))

        if capture:
            lines = capture[3]
            if capture[4]:                                  # inside ``` ... ```
                if stripped == "```":
                    close()
                    capture = None
                else:
                    lines.append(line)
                continue
            if stripped == "```" and not any(lines):        # opening fence on its own line
                capture[4] = True
                continue
            if not stripped or indent >= capture[2] + 2:
                lines.append(line)
                continue
            close()
            capture = None

        if not stripped:
            continue
        if stripped.startswith("///"):
            docs.append(stripped[3:].strip())
            continue

        while stack[-1][0] >= indent:
            stack.pop()
        parent = stack[-1][1]

        m = re.match(r"(\w+)\s*:\s*(.*)$", stripped)
        if m:
            parent["props"][m.group(1)] = m.group(2).strip()
            continue
        m = re.match(r"(\w+)\s*=\s*(.*)$", stripped)
        if m:
            capture = start(parent["props"], m.group(1), indent, m.group(2).strip())
            continue
        if re.match(r"\w+$", stripped):
            parent["props"][stripped] = True
            node(stripped, "", parent, indent)
            continue

        kind, _, rest = stripped.partition(" ")
        rest = rest.strip()
        if kind == "ref":
            continue
        m = re.match(r"'((?:[^']|'')*)'\s*(.*)$", rest)
        if m:
            name, tail = m.group(1).replace("''", "'"), m.group(2)
        else:
            name, sep, tail = rest.partition("=")
            name, tail = name.strip(), sep + tail
        new = node(kind, name, parent, indent)
        if tail.startswith("="):
            capture = start(new, "value", indent, tail[1:].strip())

    if capture:
        close()
    return root


def _unquote(name):
    name = name.strip()
    if len(name) >= 2 and name[0] == name[-1] == "'":
        return name[1:-1].replace("''", "'")
    return name


def _flag(value):
    return value is True or str(value).lower() == "true"


def _kids(node, kind):
    return [c for c in node["children"] if c["type"] == kind]


def _is_parameter(expression):
    return bool(re.search(r"IsParameterQuery\s*=\s*true", expression, re.I))


def load_tmdl(definition_dir):
    def nodes(path, kind):
        if not path.exists():
            return []
        return _kids(parse_tmdl(path.read_text(encoding="utf-8-sig")), kind)

    def split_ref(ref):
        m = re.match(r"('(?:[^']|'')*'|[^.]+)\.(.+)$", ref or "")
        return (_unquote(m.group(1)), _unquote(m.group(2))) if m else (_unquote(ref or ""), "")

    def table(t):
        partitions = []
        for p in _kids(t, "partition"):
            source = p["props"].get("source")
            entity = next(iter(_kids(p, "source")), {"props": {}})["props"]
            partitions.append({
                "kind": p["value"] or "m", "mode": p["props"].get("mode", "import"),
                "expression": source if isinstance(source, str) else "",
                "entity": _unquote(str(entity.get("entityName", ""))),
                "expressionSource": _unquote(str(entity.get("expressionSource", ""))),
            })
        return {
            "name": t["name"], "description": t["description"], "hidden": _flag(t["props"].get("isHidden")),
            "columns": [{
                "name": c["name"], "dataType": c["props"].get("dataType", ""), "expression": c["value"],
                "hidden": _flag(c["props"].get("isHidden")), "description": c["description"],
            } for c in _kids(t, "column")],
            "measures": [{
                "name": m["name"], "expression": m["value"], "description": m["description"],
                "formatString": m["props"].get("formatString", ""),
                "displayFolder": m["props"].get("displayFolder", ""),
                "hidden": _flag(m["props"].get("isHidden")),
            } for m in _kids(t, "measure")],
            "hierarchies": [{"name": h["name"], "levels": [lv["name"] for lv in _kids(h, "level")]}
                            for h in _kids(t, "hierarchy")],
            "calculationItems": [{"name": ci["name"], "expression": ci["value"]}
                                 for cg in _kids(t, "calculationGroup") for ci in _kids(cg, "calculationItem")],
            "partitions": partitions,
        }

    tables_dir = definition_dir / "tables"
    files = sorted(tables_dir.glob("*.tmdl"), key=lambda f: f.name.lower()) if tables_dir.is_dir() else []
    roles_dir = definition_dir / "roles"
    role_files = sorted(roles_dir.glob("*.tmdl")) if roles_dir.is_dir() else []

    relationships = []
    for r in nodes(definition_dir / "relationships.tmdl", "relationship"):
        p = r["props"]
        from_table, from_column = split_ref(p.get("fromColumn"))
        to_table, to_column = split_ref(p.get("toColumn"))
        relationships.append({
            "fromTable": from_table, "fromColumn": from_column, "toTable": to_table, "toColumn": to_column,
            "cardinality": p.get("fromCardinality", "many") + "-to-" + p.get("toCardinality", "one"),
            "crossFilter": p.get("crossFilteringBehavior", "oneDirection"),
            "active": str(p.get("isActive", "true")).lower() != "false",
        })

    return {
        "format": "TMDL",
        "tables": [table(t) for f in files for t in nodes(f, "table")],
        "relationships": relationships,
        "roles": [{
            "name": r["name"], "permission": r["props"].get("modelPermission", ""),
            "members": [m["name"] for m in _kids(r, "member")],
            "filters": [(tp["name"], tp["value"]) for tp in _kids(r, "tablePermission")],
        } for f in role_files for r in nodes(f, "role")],
        "expressions": [{"name": e["name"], "expression": e["value"]}
                        for e in nodes(definition_dir / "expressions.tmdl", "expression")
                        if not _is_parameter(e["value"])],
    }


# ---------------------------------------------------------------------------
# TMSL (model.bim) - same model structure as load_tmdl()
# ---------------------------------------------------------------------------

def load_tmsl(bim_path):
    data = read_json(bim_path)
    model = data.get("model", data)

    def text(value):
        return ("\n".join(value) if isinstance(value, list) else str(value or "")).strip()

    def table(t):
        partitions = []
        for p in t.get("partitions", []):
            source = p.get("source", {})
            partitions.append({
                "kind": source.get("type", "m"), "mode": p.get("mode", "import"),
                "expression": text(source.get("expression") or source.get("query")),
                "entity": source.get("entityName", ""), "expressionSource": source.get("expressionSource", ""),
            })
        return {
            "name": t.get("name", ""), "description": text(t.get("description")), "hidden": bool(t.get("isHidden")),
            "columns": [{
                "name": c.get("name", ""), "dataType": c.get("dataType", ""), "expression": text(c.get("expression")),
                "hidden": bool(c.get("isHidden")), "description": text(c.get("description")),
            } for c in t.get("columns", []) if c.get("type") != "rowNumber"],
            "measures": [{
                "name": m.get("name", ""), "expression": text(m.get("expression")),
                "description": text(m.get("description")), "formatString": m.get("formatString", ""),
                "displayFolder": m.get("displayFolder", ""), "hidden": bool(m.get("isHidden")),
            } for m in t.get("measures", [])],
            "hierarchies": [{"name": h.get("name", ""), "levels": [lv.get("name", "") for lv in h.get("levels", [])]}
                            for h in t.get("hierarchies", [])],
            "calculationItems": [{"name": ci.get("name", ""), "expression": text(ci.get("expression"))}
                                 for ci in (t.get("calculationGroup") or {}).get("calculationItems", [])],
            "partitions": partitions,
        }

    return {
        "format": "TMSL",
        "tables": sorted((table(t) for t in model.get("tables", [])), key=lambda t: t["name"].lower()),
        "relationships": [{
            "fromTable": r.get("fromTable", ""), "fromColumn": r.get("fromColumn", ""),
            "toTable": r.get("toTable", ""), "toColumn": r.get("toColumn", ""),
            "cardinality": r.get("fromCardinality", "many") + "-to-" + r.get("toCardinality", "one"),
            "crossFilter": r.get("crossFilteringBehavior", "oneDirection"),
            "active": r.get("isActive", True) is not False,
        } for r in model.get("relationships", [])],
        "roles": [{
            "name": r.get("name", ""), "permission": r.get("modelPermission", ""),
            "members": [m.get("memberName", "") for m in r.get("members", [])],
            "filters": [(tp.get("name", ""), text(tp.get("filterExpression")))
                        for tp in r.get("tablePermissions", [])],
        } for r in model.get("roles", [])],
        "expressions": [{"name": e.get("name", ""), "expression": text(e.get("expression"))}
                        for e in model.get("expressions", []) if not _is_parameter(text(e.get("expression")))],
    }


# ---------------------------------------------------------------------------
# Report - field references, filter conditions and formatting literals all use
# the same small expression language, so one formatter (fmt) handles them all.
# ---------------------------------------------------------------------------

_AGGREGATIONS = {0: "Sum", 1: "Average", 2: "Count (distinct)", 3: "Min", 4: "Max",
                 5: "Count", 6: "Median", 7: "StdDev", 8: "Variance"}
_COMPARISONS = {0: "=", 1: ">", 2: ">=", 3: "<", 4: "<="}
_TIME_UNITS = {0: "day", 1: "week", 2: "month", 3: "year", 4: "decade", 5: "second", 6: "minute", 7: "hour"}
_FIELD_KINDS = ("Column", "Measure", "HierarchyLevel", "Aggregation")
_KINDS = _FIELD_KINDS + ("Literal", "Comparison", "In", "Not", "And", "Or", "Between", "Contains",
                         "StartsWith", "EndsWith", "DateSpan", "DateAdd", "Now", "NativeVisualCalculation")


def literal(raw):
    """'text' -> text, 15D / 3L -> 15 / 3, true/false -> bool, null -> None."""
    raw = str(raw)
    if raw in ("true", "false"):
        return raw == "true"
    if raw == "null":
        return None
    if raw.startswith("'") and raw.endswith("'") and len(raw) >= 2:
        return raw[1:-1].replace("''", "'")
    m = re.match(r"datetime'([^']*)'$", raw)
    if m:
        return m.group(1).replace("T00:00:00", "")
    m = re.match(r"(-?\d+(?:\.\d+)?)[LDM]$", raw)
    return m.group(1) if m else raw


def fmt(expr, refs=None):
    """Readable text for a field reference or filter condition. Every (table, field)
    it touches is appended to `refs`, which feeds the lineage / 'used in' overviews."""
    if not isinstance(expr, dict):
        return "?"
    kind = next((k for k in expr if k in _KINDS), None)
    if kind is None:
        return next(iter(expr), "?")
    v = expr[kind] if isinstance(expr[kind], dict) else {}

    def entity(x):
        return ((x.get("Expression") or {}).get("SourceRef") or {}).get("Entity", "")

    def note(table, name):
        if refs is not None and table:
            refs.append((table, name))

    def sub(key):
        return fmt(v.get(key) or {}, refs)

    if kind in ("Column", "Measure"):
        note(entity(v), v.get("Property", ""))
        return "%s[%s]" % (entity(v), v.get("Property", ""))
    if kind == "HierarchyLevel":
        hierarchy = (v.get("Expression") or {}).get("Hierarchy") or {}
        variation = (hierarchy.get("Expression") or {}).get("PropertyVariationSource")
        name = "%s.%s" % (hierarchy.get("Hierarchy", ""), v.get("Level", ""))
        if variation:                                      # auto date/time hierarchy on a column
            table, name = entity(variation), "%s.%s" % (variation.get("Property", ""), name)
        else:
            table = entity(hierarchy)
        note(table, name)
        return "%s[%s]" % (table, name)
    if kind == "Aggregation":
        return "%s(%s)" % (_AGGREGATIONS.get(v.get("Function"), "Aggregate"), sub("Expression"))
    if kind == "Literal":
        value = literal(v.get("Value"))
        if value is None:
            return "(blank)"
        if isinstance(value, bool):
            return str(value).lower()
        return '"%s"' % value if str(v.get("Value", "")).startswith("'") else value
    if kind == "Comparison":
        return "%s %s %s" % (sub("Left"), _COMPARISONS.get(v.get("ComparisonKind"), "?"), sub("Right"))
    if kind == "In":
        return _fmt_in(v, refs, negate=False)
    if kind == "Not":
        inner = v.get("Expression") or {}
        if "In" in inner:
            return _fmt_in(inner["In"], refs, negate=True)
        return "not (%s)" % fmt(inner, refs)
    if kind in ("And", "Or"):
        return "(%s %s %s)" % (sub("Left"), kind.lower(), sub("Right"))
    if kind == "Between":
        return "%s between %s and %s" % (sub("Expression"), sub("LowerBound"), sub("UpperBound"))
    if kind in ("Contains", "StartsWith", "EndsWith"):
        words = {"Contains": "contains", "StartsWith": "starts with", "EndsWith": "ends with"}
        return "%s %s %s" % (sub("Left"), words[kind], sub("Right"))
    if kind == "Now":
        return "now"
    if kind == "DateAdd":
        return "%s %+d %s" % (sub("Expression"), v.get("Amount", 0), _TIME_UNITS.get(v.get("TimeUnit"), "?"))
    if kind == "DateSpan":
        unit = _TIME_UNITS.get(v.get("TimeUnit"), "?")
        inner = sub("Expression")
        return "this " + unit if inner == "now" else "%s of (%s)" % (unit, inner)
    return 'visual calculation "%s" = %s' % (v.get("Name", ""), v.get("Expression", ""))


def _fmt_in(v, refs, negate):
    left = ", ".join(fmt(e, refs) for e in v.get("Expressions", []))
    values = [", ".join(fmt(c) for c in row) for row in v.get("Values", [])]
    if not values:
        return left + (" not in" if negate else " in") + " (subquery)"
    if len(values) == 1:
        return "%s %s %s" % (left, "<>" if negate else "=", values[0])
    shown = values[:25] + (["... +%d more" % (len(values) - 25)] if len(values) > 25 else [])
    return "%s %s (%s)" % (left, "not in" if negate else "in", ", ".join(shown))


def _resolve(obj, aliases):
    """Replaces query aliases ({"Source": "k"}) by the table they stand for."""
    if isinstance(obj, list):
        return [_resolve(x, aliases) for x in obj]
    if isinstance(obj, dict):
        if list(obj) == ["Source"] and obj["Source"] in aliases:
            return {"Entity": aliases[obj["Source"]]}
        return {k: _resolve(x, aliases) for k, x in obj.items()}
    return obj


def _conditions(query, refs):
    """The Where clauses of a filter / slicer selection, as one readable string."""
    query = query or {}
    aliases = {s.get("Name"): s.get("Entity") for s in query.get("From", []) if s.get("Entity")}
    return " and ".join(fmt(_resolve(w.get("Condition") or {}, aliases), refs) for w in query.get("Where", []))


def parse_filters(config):
    filters = []
    for f in (config or {}).get("filters") or []:
        refs = []
        field = fmt(f.get("field") or f.get("expression") or {}, refs)
        filters.append({
            "field": field, "type": f.get("type", ""), "label": f.get("displayName", ""),
            "condition": _conditions(f.get("filter"), refs),
            "hidden": bool(f.get("isHiddenInViewMode")), "locked": bool(f.get("isLockedInViewMode")),
            "refs": refs,
        })
    return filters


def _value(val):
    """Value of a formatting property: literal, resource name, or '= field' when dynamic."""
    if not isinstance(val, dict):
        return val
    expr = val.get("expr")
    if not isinstance(expr, dict):
        return None
    if "Literal" in expr:
        return literal(expr["Literal"].get("Value"))
    if "ResourcePackageItem" in expr:
        return expr["ResourcePackageItem"].get("ItemName")
    return "= " + fmt(expr)


def _properties(objects):
    for name, items in (objects or {}).items():
        for item in items or []:
            for key, val in (item.get("properties") or {}).items():
                yield name, key, val


def _prop(objects, name, key):
    for obj, k, val in _properties(objects):
        if obj == name and k == key:
            return _value(val)
    return None


def _find_fields(obj, refs, found):
    """Fields referenced from formatting (conditional colours, dynamic titles, labels)."""
    if isinstance(obj, list):
        for x in obj:
            _find_fields(x, refs, found)
    elif isinstance(obj, dict):
        kind = next((k for k in _FIELD_KINDS if isinstance(obj.get(k), dict)), None)
        if kind:
            found.append(fmt({kind: obj[kind]}, refs))
        else:
            for key, x in obj.items():
                if key != "filter":
                    _find_fields(x, refs, found)


def _region(pos, page_width, page_height):
    cx = (pos["x"] + pos["width"] / 2.0) / (page_width or 1)
    cy = (pos["y"] + pos["height"] / 2.0) / (page_height or 1)
    vertical = "top" if cy < 1 / 3.0 else "bottom" if cy > 2 / 3.0 else "middle"
    horizontal = "left" if cx < 1 / 3.0 else "right" if cx > 2 / 3.0 else "center"
    if pos["width"] >= 0.85 * (page_width or 1):
        horizontal = "full width"
    return vertical + " " + horizontal


def parse_visual(data, names, groups, page_width, page_height):
    visual = data.get("visual") or {}
    objects = visual.get("objects") or {}
    container = visual.get("visualContainerObjects") or {}
    query = visual.get("query") or {}
    refs = []

    fields = []                                    # (role, field, display name)
    for role, role_data in (query.get("queryState") or {}).items():
        for proj in (role_data or {}).get("projections") or []:
            label = fmt(proj["field"], refs) if proj.get("field") else proj.get("queryRef", "")
            fields.append((role, label, proj.get("displayName") or ""))

    sort = ["%s (%s)" % (fmt(s.get("field") or {}), str(s.get("direction", "")).lower())
            for s in (query.get("sortDefinition") or {}).get("sort") or []]

    text = []                                      # text box paragraphs
    other_text = []                                # button captions, shape text, card labels
    selection = ""
    for name, key, val in _properties(objects):
        if key == "paragraphs" and isinstance(val, list):
            for paragraph in val:
                runs = paragraph.get("textRuns") or []
                line = "".join(str(r.get("value", "")) for r in runs)
                urls = [r["url"] for r in runs if r.get("url")]
                text.append(line + (" (link: %s)" % ", ".join(urls) if urls else ""))
        elif key in ("text", "titleText"):
            value = _value(val)
            if isinstance(value, str) and value.strip():
                other_text.append(value.strip())
        elif key == "filter" and isinstance(val, dict):
            selection = _conditions(val.get("filter"), refs)

    dynamic = []
    _find_fields([objects, container], refs, dynamic)
    shown = set(f[1] for f in fields)
    dynamic = [d for d in dict.fromkeys(dynamic) if d not in shown]

    details = []
    for label, value in (
        ("Slicer style", _prop(objects, "data", "mode")),
        ("Single select", "yes" if _prop(objects, "selection", "singleSelect") is True else None),
        ("Selected by default", selection),
        ("Image", _prop(objects, "general", "imageUrl")),
        ("Shape", _prop(objects, "shape", "tileShape")),
        ("Icon", _prop(objects, "icon", "shapeType")),
    ):
        if value:
            details.append("%s: %s" % (label, value))

    link = [_prop(container, "visualLink", k) for k in ("type", "navigationSection", "bookmark", "webUrl")]
    if link[0] and _prop(container, "visualLink", "show") is not False:
        target = next((names.get(t, t) for t in link[1:] if isinstance(t, str) and t), "")
        details.append("Action: %s%s" % (link[0], ' -> "%s"' % target if target else ""))
    tooltip = _prop(container, "visualTooltip", "section")
    if tooltip and _prop(container, "visualTooltip", "show") is not False:
        details.append('Tooltip page: "%s"' % names.get(tooltip, tooltip))

    pos = {k: float((data.get("position") or {}).get(k) or 0) for k in ("x", "y", "z", "width", "height")}
    title = _prop(container, "title", "text")
    subtitle = _prop(container, "subTitle", "text")
    return {
        "type": visual.get("visualType") or "unknown",
        "title": str(title) if title else "",
        "titleShown": _prop(container, "title", "show") is not False,
        "subtitle": str(subtitle) if subtitle and _prop(container, "subTitle", "show") is not False else "",
        "text": text,
        "otherText": list(dict.fromkeys(other_text)),
        "fields": fields, "dynamicFields": dynamic, "sort": sort, "details": details,
        "filters": parse_filters(data.get("filterConfig")),
        "pos": pos, "region": _region(pos, page_width, page_height),
        "hidden": bool(data.get("isHidden")),
        "group": groups.get(data.get("parentGroupName"), ""),
        "refs": refs,
    }


def build_report(report_json, raw_pages, bookmarks):
    """raw_pages: [(page.json dict, [visual.json dict, ...]), ...] in display order."""
    names = {p.get("name"): p.get("displayName", p.get("name")) for p, _ in raw_pages}
    names.update(bookmarks)
    themes = report_json.get("themeCollection") or {}
    pages = []
    for page, visuals in raw_pages:
        width, height = page.get("width") or 1280, page.get("height") or 720
        groups = {v.get("name"): (v["visualGroup"] or {}).get("displayName", "group")
                  for v in visuals if isinstance(v.get("visualGroup"), dict)}
        background = (((_first(page.get("objects"), "background").get("image") or {}).get("image") or {}).get("name"))
        parsed = [parse_visual(v, names, groups, width, height) for v in visuals if "visualGroup" not in v]
        parsed.sort(key=lambda v: (round(v["pos"]["y"]), round(v["pos"]["x"])))
        pages.append({
            "name": page.get("displayName") or page.get("name") or "?",
            "width": width, "height": height,
            "kind": page.get("type") or (page.get("pageBinding") or {}).get("type") or "",
            "hidden": page.get("visibility") == "HiddenInViewMode",
            "background": _value(background) if background else "",
            "filters": parse_filters(page.get("filterConfig")),
            "visuals": parsed,
        })
    return {
        "theme": ", ".join(t["name"] for t in (themes.get("customTheme"), themes.get("baseTheme"))
                           if isinstance(t, dict) and t.get("name")),
        "filters": parse_filters(report_json.get("filterConfig")),
        "bookmarks": list(bookmarks.values()),
        "pages": pages,
    }


def _first(objects, name):
    items = (objects or {}).get(name) or [{}]
    return items[0].get("properties") or {}


def load_report(report_dir):
    """Returns the parsed report, or None when the folder holds no report definition."""
    definition = report_dir / "definition"
    if (definition / "pages").is_dir():                       # PBIR (one file per page/visual)
        pages_dir = definition / "pages"
        dirs = {d.name: d for d in pages_dir.iterdir() if (d / "page.json").exists()}
        order = []
        if (pages_dir / "pages.json").exists():
            order = [n for n in read_json(pages_dir / "pages.json").get("pageOrder", []) if n in dirs]
        order += sorted(set(dirs) - set(order))
        raw_pages = [(read_json(dirs[n] / "page.json"),
                      [read_json(f) for f in sorted((dirs[n] / "visuals").glob("*/visual.json"))])
                     for n in order]
        bookmarks = {}
        for f in sorted((definition / "bookmarks").glob("*.bookmark.json")):
            b = read_json(f)
            bookmarks[b.get("name", f.stem)] = b.get("displayName", f.stem)
        report_json = read_json(definition / "report.json") if (definition / "report.json").exists() else {}
        return build_report(report_json, raw_pages, bookmarks)

    if (report_dir / "report.json").exists():                 # legacy single-file report.json
        return _load_legacy_report(read_json(report_dir / "report.json"))
    return None


def _load_legacy_report(data):
    """Converts the legacy report.json (JSON-in-JSON) to the PBIR shape used above."""
    def loads(value, default):
        if isinstance(value, str):
            try:
                return json.loads(value)
            except ValueError:
                return default
        return value if value is not None else default

    def visual(vc):
        config = loads(vc.get("config"), {})
        out = {
            "name": config.get("name"), "parentGroupName": config.get("parentGroupName"),
            "position": {k: vc.get(k) for k in ("x", "y", "z", "width", "height")},
            "filterConfig": {"filters": loads(vc.get("filters"), [])},
        }
        if "singleVisualGroup" in config:
            out["visualGroup"] = config["singleVisualGroup"] or {}
            return out
        single = config.get("singleVisual") or {}
        prototype = single.get("prototypeQuery") or {}
        aliases = {s.get("Name"): s.get("Entity") for s in prototype.get("From", []) if s.get("Entity")}
        select = {s.get("Name"): _resolve(s, aliases) for s in prototype.get("Select", [])}
        renamed = single.get("columnProperties") or {}
        state = {}
        for role, projections in (single.get("projections") or {}).items():
            state[role] = {"projections": [{
                "field": select.get(p.get("queryRef")), "queryRef": p.get("queryRef"),
                "displayName": (renamed.get(p.get("queryRef")) or {}).get("displayName"),
            } for p in projections]}
        out["isHidden"] = (single.get("display") or {}).get("mode") == "hidden"
        out["visual"] = {
            "visualType": single.get("visualType"), "objects": single.get("objects"),
            "visualContainerObjects": single.get("vcObjects"),
            "query": {"queryState": state, "sortDefinition": {"sort": [
                {"field": _resolve(o.get("Expression") or {}, aliases),
                 "direction": "Descending" if o.get("Direction") == 2 else "Ascending"}
                for o in prototype.get("OrderBy", [])]}},
        }
        return out

    config = loads(data.get("config"), {})
    raw_pages = []
    for section in data.get("sections", []):
        section_config = loads(section.get("config"), {})
        page = {
            "name": section.get("name"), "displayName": section.get("displayName"),
            "width": section.get("width"), "height": section.get("height"),
            "filterConfig": {"filters": loads(section.get("filters"), [])},
            "objects": section_config.get("objects"),
            "visibility": "HiddenInViewMode" if section_config.get("visibility") == 1 else "",
        }
        raw_pages.append((page, [visual(vc) for vc in section.get("visualContainers", [])]))
    bookmarks = {}
    for b in config.get("bookmarks", []):
        for item in b.get("children") or [b]:
            bookmarks[item.get("name")] = item.get("displayName", "")
    report_json = {"themeCollection": config.get("themeCollection"),
                   "filterConfig": {"filters": loads(data.get("filters"), [])}}
    return build_report(report_json, raw_pages, bookmarks)


# ---------------------------------------------------------------------------
# Lineage helpers
# ---------------------------------------------------------------------------

_CONNECTORS = [  # most specific first: a Fabric warehouse still calls Sql.Database()
    (r"\.datawarehouse\.fabric\.microsoft\.com|FABRIC_SQL_ENDPOINT", "Fabric SQL endpoint (Warehouse / Lakehouse)"),
    (r"\bLakehouse\.Contents\s*\(", "Fabric Lakehouse (OneLake)"),
    (r"\bAnalysisServices\.Databases?\s*\(|powerbi://", "Power BI semantic model / Analysis Services"),
    (r"\bPowerPlatform\.Dataflows\s*\(|\bPowerBI\.Dataflows\s*\(", "Power BI / Fabric dataflow"),
    (r"sharepoint\.com|\bSharePoint\.(Files|Contents|Tables)\s*\(", "SharePoint"),
    (r"\bOneDrive\.", "OneDrive"),
    (r"\bSql\.Databases?\s*\(", "SQL Server / Azure SQL"),
    (r"\bPostgreSQL\.Database\s*\(", "PostgreSQL"),
    (r"\bMySQL\.Database\s*\(", "MySQL"),
    (r"\bOracle\.Database\s*\(", "Oracle"),
    (r"\bSnowflake\.Databases\s*\(", "Snowflake"),
    (r"\bDatabricks\.(Catalogs|Query)\s*\(", "Databricks"),
    (r"\bOdbc\.(Query|DataSource)\s*\(", "ODBC"),
    (r"\bOData\.Feed\s*\(", "OData feed"),
    (r"\bWeb\.(Contents|BrowserContents|Page)\s*\(", "Web"),
    (r"\bExcel\.Workbook\s*\(", "Excel workbook"),
    (r"\bCsv\.Document\s*\(", "CSV file"),
    (r"\bJson\.Document\s*\(", "JSON"),
    (r"\bFolder\.(Files|Contents)\s*\(", "Folder"),
    (r"\bTable\.FromRows\s*\(|#table\s*\(", "Data entered in the query"),
]
_NAVIGATION_RE = re.compile(
    r'\[\s*(?:Schema\s*=\s*"(?P<schema>[^"]*)"\s*,\s*)?Item\s*=\s*"(?P<item>[^"]*)"', re.I)


def connector(m_expression):
    return next((label for pattern, label in _CONNECTORS if re.search(pattern, m_expression, re.I)),
                "other / custom M source")


def describe_source(table, shared):
    """One line on where a table's data comes from."""
    for p in table["partitions"]:
        if p["entity"]:
            via = shared.get(p["expressionSource"], "")
            where = " via shared expression `%s`" % p["expressionSource"] if p["expressionSource"] else ""
            kind = " (%s)" % connector(via) if via else ""
            return "%s, entity `%s`%s%s" % (p["mode"], p["entity"], where, kind)
        if p["kind"] == "calculated":
            return "calculated table (DAX, shown below)"
        if p["expression"]:
            used = [name for name in shared if re.search(r"(?<![\w\"])#?\"?%s\"?(?![\w\"])" % re.escape(name),
                                                         p["expression"])]
            label = connector(p["expression"])
            if label.startswith("other") and used:
                label = connector(shared[used[0]])
            nav = _NAVIGATION_RE.search(p["expression"])
            parts = ["%s, Power Query: %s" % (p["mode"], label)]
            if nav and nav.group("schema"):
                parts.append("schema `%s`" % nav.group("schema"))
            if nav:
                parts.append("object `%s`" % nav.group("item"))
            if used:
                parts.append("uses shared expression " + ", ".join("`%s`" % u for u in used))
            return ", ".join(parts)
    return ""


def dax_refs(dax, table_names, measure_tables):
    """Measures, columns and tables a DAX expression refers to."""
    code = re.sub(r'"(?:[^"]|"")*"', '""', dax)
    code = re.sub(r"//[^\n]*|--[^\n]*|/\*.*?\*/", " ", code, flags=re.S)
    columns, measures, tables = set(), set(), set()
    for m in re.finditer(r"(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))\[([^\]]+)\]", code):
        table, name = (m.group(1) or m.group(2)).replace("''", "'"), m.group(3)
        tables.add(table)
        if measure_tables.get(name) == table:
            measures.add(name)
        else:
            columns.add("%s[%s]" % (table, name))
    measures.update(n for n in re.findall(r"(?<![\w'\]])\[([^\]]+)\]", code) if n in measure_tables)
    bare = re.sub(r"\[[^\]]*\]", "[]", code)
    for name in table_names:
        quoted = "'%s'" % name.replace("'", "''")
        if quoted.lower() in bare.lower():
            tables.add(name)
        elif re.match(r"\w+$", name) and re.search(r"(?<![\w'])%s(?![\w'])" % re.escape(name), bare, re.I):
            tables.add(name)
    return {"measures": sorted(measures), "columns": sorted(columns),
            "tables": set(t for t in tables if t in table_names)}


def report_usage(report):
    """{(table, field): [where used]} and {table: Counter(page name)} for the whole report."""
    fields, tables = {}, {}

    def add(refs, label, page):
        for ref in dict.fromkeys(refs):
            fields.setdefault(ref, []).append(label)
        for table in set(t for t, _ in refs):
            tables.setdefault(table, Counter())[page] += 1

    for f in report["filters"]:
        add(f["refs"], "report-level filter (all pages)", "report-level filters")
    for page in report["pages"]:
        for f in page["filters"]:
            add(f["refs"], 'page filter on "%s"' % page["name"], page["name"])
        for v in page["visuals"]:
            title = ' "%s"' % v["title"] if v["title"] else ""
            refs = v["refs"] + [r for f in v["filters"] for r in f["refs"]]
            add(refs, '%s%s on page "%s"' % (v["type"], title, page["name"]), page["name"])
    return fields, tables


# ---------------------------------------------------------------------------
# Markdown renderer
# ---------------------------------------------------------------------------

def _cell(text):
    return str(text).replace("|", "\\|").replace("\n", " ").strip()


def _code(text):
    return "`%s`" % str(text).replace("`", "'")


def _counted(labels):
    return ["%s%s" % (label, " (x%d)" % n if n > 1 else "") for label, n in Counter(labels).items()]


def _filter_lines(filters, indent=""):
    """Filters with a condition one per line; untouched filter cards on one combined line."""
    lines, unset = [], []
    for f in filters:
        if not f["condition"]:
            unset.append(_code(f["field"]))
            continue
        flags = [f["type"]] if f["type"] else []
        flags += ["hidden from viewers"] if f["hidden"] else []
        flags += ["locked"] if f["locked"] else []
        name = ' (filter card "%s")' % f["label"] if f["label"] else ""
        lines.append("%s- %s%s%s" % (indent, f["condition"], name, " _[%s]_" % ", ".join(flags) if flags else ""))
    if unset:
        lines.append("%s- Filter cards without a condition (all values): %s" % (indent, ", ".join(unset)))
    return lines


_STATIC_VISUALS = ("textbox", "image", "shape", "basicShape", "actionButton", "bookmarkNavigator", "pageNavigator")


def render_report(report, out):
    pages = report["pages"]
    visuals = sum(len(p["visuals"]) for p in pages)
    out("## Report")
    out("")
    out("- Pages: %d (%d hidden or tooltip/drillthrough pages), visuals: %d"
        % (len(pages), sum(1 for p in pages if p["hidden"] or p["kind"]), visuals))
    if report["theme"]:
        out("- Theme: %s" % report["theme"])
    if report["bookmarks"]:
        out("- Bookmarks: %s" % ", ".join('"%s"' % b for b in report["bookmarks"]))
    out("- Page order: %s" % " > ".join(p["name"] for p in pages))
    out("")
    out("### Filters on all pages (report level)")
    out("")
    for line in _filter_lines(report["filters"]) or ["_None._"]:
        out(line)
    out("")

    for number, page in enumerate(pages, 1):
        tags = [page["kind"].lower() + " page"] if page["kind"] else []
        tags += ["hidden"] if page["hidden"] else []
        out("### Page %d: %s%s" % (number, page["name"], " _(%s)_" % ", ".join(tags) if tags else ""))
        out("")
        out("- Canvas: %g x %g px%s" % (page["width"], page["height"],
                                        ", background image `%s`" % page["background"] if page["background"] else ""))
        types = Counter(v["type"] for v in page["visuals"])
        out("- Visuals: %d%s" % (len(page["visuals"]),
                                 " (%s)" % ", ".join("%d %s" % (n, t) for t, n in sorted(types.items()))
                                 if types else ""))
        out("- Filters on this page:%s" % ("" if page["filters"] else " none"))
        for line in _filter_lines(page["filters"], "  "):
            out(line)
        out("")
        if page["visuals"]:
            out("Visuals, top to bottom and left to right (x, y = top-left corner in px):")
            out("")
        for index, v in enumerate(page["visuals"], 1):
            pos = v["pos"]
            head = "%d. **%s**" % (index, v["type"])
            if v["title"]:
                head += ' "%s"%s' % (v["title"], "" if v["titleShown"] else " (title not shown)")
            head += " - %s (x %d, y %d, %d x %d)" % (v["region"], round(pos["x"]), round(pos["y"]),
                                                    round(pos["width"]), round(pos["height"]))
            extras = (["hidden"] if v["hidden"] else []) + (['in group "%s"' % v["group"]] if v["group"] else [])
            out(head + (" _[%s]_" % ", ".join(extras) if extras else ""))
            if v["subtitle"]:
                out('    - Subtitle: "%s"' % v["subtitle"])
            if v["text"]:
                out("    - Text:")
                for line in v["text"]:
                    out("      > %s" % line if line.strip() else "      >")
            if v["otherText"]:
                out("    - Labels: %s" % ", ".join('"%s"' % t for t in v["otherText"]))
            roles = {}
            for role, field, shown_as in v["fields"]:
                roles.setdefault(role, []).append(_code(field) + (' (shown as "%s")' % shown_as if shown_as else ""))
            for role, items in roles.items():
                out("    - %s: %s" % (role, ", ".join(items)))
            if not v["fields"] and v["type"] not in _STATIC_VISUALS:
                out("    - No data fields")
            if v["dynamicFields"]:
                out("    - Fields used in formatting / labels: %s" % ", ".join(_code(d) for d in v["dynamicFields"]))
            if v["sort"]:
                out("    - Sorted by: %s" % ", ".join(v["sort"]))
            for detail in v["details"]:
                out("    - %s" % detail)
            if v["filters"]:
                out("    - Filters on this visual:")
                for line in _filter_lines(v["filters"], "      "):
                    out(line)
        out("")


def render_model(model, report, row_counts, out):
    tables = model["tables"]
    table_names = [t["name"] for t in tables]
    measure_tables = {m["name"]: t["name"] for t in tables for m in t["measures"]}
    shared = {e["name"]: e["expression"] for e in model["expressions"]}
    field_usage, table_usage = report_usage(report) if report else ({}, {})

    # What refers to each table from DAX elsewhere in the model.
    deps = {}
    used_by = {name: {"measures": [], "columns": []} for name in table_names}
    for t in tables:
        for kind, items in (("measures", t["measures"]), ("columns", t["columns"])):
            for item in items:
                if not item["expression"]:
                    continue
                refs = dax_refs(item["expression"], table_names, measure_tables)
                deps[(t["name"], item["name"])] = refs
                label = item["name"] if kind == "measures" else "%s[%s]" % (t["name"], item["name"])
                for name in refs["tables"]:
                    used_by[name][kind].append(label)

    out("## Semantic model")
    out("")
    out("- Format: %s" % model["format"])
    out("- Tables: %d (%d hidden), columns: %d, measures: %d, relationships: %d, RLS roles: %d" % (
        len(tables), sum(1 for t in tables if t["hidden"]), sum(len(t["columns"]) for t in tables),
        len(measure_tables), len(model["relationships"]), len(model["roles"])))
    if row_counts:
        out("- Row counts: pasted from a manually run DAX query, so possibly stale. Total %s rows in %d tables."
            % (format(sum(row_counts.get(n, 0) for n in table_names), ","),
               sum(1 for n in table_names if n in row_counts)))
    out("")

    out("## Tables")
    out("")
    for t in tables:
        name = t["name"]
        out("### %s%s" % (name, " _(hidden)_" if t["hidden"] else ""))
        out("")
        if t["description"]:
            out(t["description"])
            out("")
        source = describe_source(t, shared)
        if source:
            out("- Source: %s" % source)
        if name in row_counts:
            out("- Rows: %s" % format(row_counts[name], ","))
        for h in t["hierarchies"]:
            out("- Hierarchy %s: %s" % (_code(h["name"]), " > ".join(h["levels"])))
        related = ["%s[%s] %s %s[%s]%s" % (
            r["fromTable"], r["fromColumn"], "<->" if r["crossFilter"] == "bothDirections" else "->",
            r["toTable"], r["toColumn"], "" if r["active"] else " (inactive)")
            for r in model["relationships"] if name in (r["fromTable"], r["toTable"])]
        if related:
            out("- Relationships: %s" % "; ".join(_code(r) for r in related))
        if t["measures"]:
            out("- Measures defined here: %d (see Measures)" % len(t["measures"]))
        if used_by[name]["measures"]:
            out("- Referenced by measures: %s" % ", ".join(_code(m) for m in used_by[name]["measures"]))
        if used_by[name]["columns"]:
            out("- Referenced by calculated columns: %s" % ", ".join(_code(c) for c in used_by[name]["columns"]))
        if report:
            usage = table_usage.get(name)
            out("- Used in report: %s" % (", ".join("%s (%d)" % item for item in usage.items())
                                          if usage else "not used in any visual or filter"))
        out("")
        if t["columns"]:
            out("| Column | Data type | Notes |")
            out("|---|---|---|")
            for c in t["columns"]:
                notes = (["calculated"] if c["expression"] else []) + (["hidden"] if c["hidden"] else [])
                notes += [c["description"]] if c["description"] else []
                out("| %s | %s | %s |" % (_cell(c["name"]), _cell(c["dataType"]), _cell("; ".join(notes))))
            out("")
        for kind, items in (("Calculated column", [c for c in t["columns"] if c["expression"]]),
                            ("Calculation item", t["calculationItems"])):
            for item in items:
                out("%s %s:" % (kind, _code(item["name"])))
                out("```dax")
                out(sanitize(item["expression"]))
                out("```")
                out("")
        for p in t["partitions"]:
            if p["expression"]:
                calculated = p["kind"] == "calculated"
                out("Calculated table expression:" if calculated else "Power Query (%s):" % p["mode"])
                out("```dax" if calculated else "```powerquery")
                out(sanitize(p["expression"]))
                out("```")
                out("")

    out("## Measures")
    out("")
    if not measure_tables:
        out("_No measures in this model._")
        out("")
    for t in tables:
        for m in t["measures"]:
            out("### %s" % m["name"])
            out("")
            info = ["Table: %s" % _code(t["name"])]
            info += ["Folder: %s" % m["displayFolder"]] if m["displayFolder"] else []
            info += ["Format: %s" % _code(m["formatString"])] if m["formatString"] else []
            info += ["hidden"] if m["hidden"] else []
            out("- " + " | ".join(info))
            if m["description"]:
                out("- Description: %s" % m["description"].replace("\n", " "))
            refs = deps.get((t["name"], m["name"]), {})
            others = [x for x in refs.get("measures", []) if x != m["name"]]
            if others:
                out("- Uses measures: %s" % ", ".join(_code(x) for x in others))
            if refs.get("columns"):
                out("- Uses columns: %s" % ", ".join(_code(x) for x in refs["columns"]))
            if report:
                usage = _counted(field_usage.get((t["name"], m["name"]), []))
                out("- Used in report: %s" % ("; ".join(usage) if usage else "not used"))
            out("")
            out("```dax")
            out(sanitize(m["expression"]) or "-- no expression")
            out("```")
            out("")

    out("## Relationships")
    out("")
    if model["relationships"]:
        out("| From | To | Cardinality | Cross filter | Active |")
        out("|---|---|---|---|---|")
        for r in model["relationships"]:
            out("| %s | %s | %s | %s | %s |" % (
                _cell("%s[%s]" % (r["fromTable"], r["fromColumn"])), _cell("%s[%s]" % (r["toTable"], r["toColumn"])),
                r["cardinality"], r["crossFilter"], "yes" if r["active"] else "no"))
    else:
        out("_No relationships._")
    out("")

    if model["roles"]:
        out("## Row-level security")
        out("")
        for role in model["roles"]:
            out("### %s" % role["name"])
            out("")
            if role["permission"]:
                out("- Model permission: %s" % role["permission"])
            if role["members"]:
                out("- Members: %s" % ", ".join(role["members"]))
            for table, expression in role["filters"]:
                out("- Filter on %s: %s" % (_code(table), _code(sanitize(expression).replace("\n", " "))))
            out("")

    if model["expressions"]:
        out("## Shared Power Query expressions")
        out("")
        out("Query parameters are left out on purpose (they tend to hold server names and the like).")
        out("")
        for e in model["expressions"]:
            out("### %s" % e["name"])
            out("")
            out("```powerquery")
            out(sanitize(e["expression"]))
            out("```")
            out("")


def render(project, report, model, row_counts=None, live_connection=False):
    generated_at = datetime.now(ZoneInfo("Europe/Amsterdam")).isoformat(timespec="seconds")
    lines = []
    out = lines.append
    out("- Generated at: %s" % generated_at)
    out("- Generated timezone: Europe/Amsterdam")
    out("")
    out("# Power BI documentation: %s" % project)
    out("")
    out("Generated by pbip_extract.py from the .pbip project files. Contents: "
        + ", ".join((["Report (pages, layout, visuals, filters)"] if report else [])
                    + (["Semantic model", "Tables", "Measures", "Relationships"] if model else [])) + ".")
    out("")
    if live_connection:
        out("_The report uses a live connection to a published semantic model, so there is no "
            "local model in this project to document._")
        out("")
    if report:
        render_report(report, out)
    if model:
        render_model(model, report, row_counts or {}, out)
    return "\n".join(lines).rstrip() + "\n"


# ---------------------------------------------------------------------------
# Row counts - the project files hold no data, so counts come from a DAX query
# the user runs by hand in Power BI Desktop and pastes back (--rowcounts).
# ---------------------------------------------------------------------------

def rowcount_query(tables):
    rows = ['\tROW("Table", "%s", "Rows", COUNTROWS(\'%s\'))'
            % (t["name"].replace('"', '""'), t["name"].replace("'", "''")) for t in tables]
    lines = [
        "// Row-count query generated by pbip_extract.py",
        "// 1. Open the project in Power BI Desktop > DAX query view (or DAX Studio) and run this query.",
        "// 2. Copy the result grid including the header into a text file.",
        "// 3. Run pbip_extract.py again with: --rowcounts <that file>",
        "EVALUATE",
    ]
    lines += ["UNION(", ",\n".join(rows), ")", "ORDER BY [Table]"] if len(rows) > 1 else rows
    return "\n".join(lines) + "\n"


def parse_rowcounts(path):
    """Pasted result grid (tab, semicolon or comma separated, with header) -> {table: rows}."""
    lines = [ln for ln in Path(path).read_text(encoding="utf-8-sig").splitlines() if ln.strip()]
    if len(lines) < 2:
        return {}
    delimiter = next((d for d in ("\t", ";", ",") if d in lines[0]), None)

    def split(line):
        return [c.strip().strip('"').strip("[]").strip() for c in (line.split(delimiter) if delimiter else [line])]

    header = [h.lower() for h in split(lines[0])]
    table_index = next((i for i, h in enumerate(header) if "table" in h), 0)
    rows_index = next((i for i, h in enumerate(header) if "row" in h), len(header) - 1)
    counts = {}
    for line in lines[1:]:
        cols = split(line)
        if len(cols) > max(table_index, rows_index):
            digits = re.sub(r"\D", "", cols[rows_index])
            if cols[table_index] and digits:
                counts[cols[table_index]] = int(digits)
    return counts


# ---------------------------------------------------------------------------
# Project discovery and entry points
# ---------------------------------------------------------------------------

def locate(pbip_file):
    """Follows the references in the .pbip and definition.pbir files.
    Returns (.pbip file, report folder, model path or None, model format, live connection?)."""
    pbip_file = Path(pbip_file).expanduser()
    if pbip_file.is_dir():
        candidates = sorted(pbip_file.glob("*.pbip"))
        if len(candidates) != 1:
            raise ExtractError("Expected a .pbip file, got a folder with %d .pbip files: %s"
                               % (len(candidates), pbip_file))
        pbip_file = candidates[0]
    if not pbip_file.is_file() or pbip_file.suffix.lower() != ".pbip":
        raise ExtractError("Not a .pbip file: %s" % pbip_file)

    report_dir = None
    for artifact in read_json(pbip_file).get("artifacts", []):
        path = (artifact.get("report") or {}).get("path")
        if path and (pbip_file.parent / path).is_dir():
            report_dir = (pbip_file.parent / path).resolve()
            break
    if report_dir is None:
        raise ExtractError("The report folder referenced in %s was not found." % pbip_file.name)

    reference = {}
    if (report_dir / "definition.pbir").exists():
        reference = read_json(report_dir / "definition.pbir").get("datasetReference") or {}
    by_path = (reference.get("byPath") or {}).get("path")
    if not by_path:
        return pbip_file, report_dir, None, "", "byConnection" in reference
    model_dir = (report_dir / by_path).resolve()
    if (model_dir / "definition").is_dir() and any((model_dir / "definition").rglob("*.tmdl")):
        return pbip_file, report_dir, model_dir / "definition", "TMDL", False
    bim = sorted(model_dir.glob("*.bim"))
    if bim:
        return pbip_file, report_dir, bim[0], "TMSL", False
    raise ExtractError("No semantic model (TMDL definition folder or .bim file) found in: %s" % model_dir)


def run(pbip_file, output=None, rowcounts=None, write_rowcount_query=True, copilot=False, log=print):
    """Generates the documentation and returns the path of the written file."""
    pbip_file, report_dir, model_path, model_format, live = locate(pbip_file)
    project = pbip_file.stem
    log("Report:  %s" % report_dir)

    report = load_report(report_dir)
    if report:
        log("         %d pages, %d visuals" % (len(report["pages"]), sum(len(p["visuals"]) for p in report["pages"])))
    else:
        log("         no report definition found - report section skipped")

    model = None
    if model_path:
        log("Model:   %s (%s)" % (model_path, model_format))
        model = load_tmdl(model_path) if model_format == "TMDL" else load_tmsl(model_path)
        log("         %d tables, %d measures, %d relationships" % (
            len(model["tables"]), sum(len(t["measures"]) for t in model["tables"]), len(model["relationships"])))
    elif live:
        log("Model:   live connection to a published semantic model - only the report is documented")
    if not report and not model:
        raise ExtractError("Nothing to document: no report definition and no local semantic model found.")

    suffix = "_copilot_kb.txt" if copilot else "_docs.md"
    output = Path(output) if output else Path.cwd() / (project.replace(" ", "_") + suffix)
    if output.is_dir():
        output = output / (project.replace(" ", "_") + suffix)

    row_counts = {}
    if rowcounts:
        row_counts = parse_rowcounts(rowcounts)
        log("Row counts loaded for %d table(s)" % len(row_counts))
    elif model and write_rowcount_query:
        query_path = output.parent / (project.replace(" ", "_") + "_rowcount_query.dax")
        query_path.write_text(rowcount_query(model["tables"]), encoding="utf-8")
        log("Row-count helper query: %s (optional, see README)" % query_path)

    output.write_text(render(project, report, model, row_counts, live), encoding="utf-8")
    log("Written: %s (%.0f KB)" % (output.resolve(), output.stat().st_size / 1024.0))
    return output


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):          # never crash on a legacy Windows code page
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="Document a Power BI project (.pbip) as Markdown.")
    ap.add_argument("pbip_file", help="path to the .pbip file (or the folder that contains exactly one)")
    ap.add_argument("-o", "--output", help="output file or folder (default: <project>_docs.md in the current folder)")
    ap.add_argument("--rowcounts", help="text file with the pasted result of the row-count helper query")
    ap.add_argument("--no-rowcount-query", action="store_true", help="don't write the *_rowcount_query.dax helper")
    ap.add_argument("--copilot", action="store_true",
                    help="same content, but named <project>_copilot_kb.txt (for use as an LLM knowledge file)")
    args = ap.parse_args(argv)
    try:
        run(args.pbip_file, args.output, args.rowcounts, not args.no_rowcount_query, args.copilot)
    except (ExtractError, OSError, ValueError) as e:
        print("ERROR: %s" % e, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
