#!/usr/bin/env python3
"""
pbip_document.py — Power BI Project (.pbip) Documentation Generator

Supports both TMSL (.bim) and TMDL (.tmdl) semantic model formats.

Output modes:
  Default:   Markdown documentation (human-readable reference)
  --copilot: Plain-text knowledge base optimised for Copilot / ChatGPT context injection.
             Describes every field, measure, and relationship in prose so an LLM can
             answer questions like "how is measure X built?" or "what does column Y mean?"

Usage:
    python pbip_document.py <path-to-pbip-folder> [--output docs.md]
    python pbip_document.py <path-to-pbip-folder> --copilot [--output knowledge.txt]

Requirements: Python 3.8+ (stdlib only, no pip installs needed)
"""

import argparse
import json
import os
import re
import sys
import urllib.parse
from collections import Counter
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def read_json(path: Path) -> dict:
    with open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def slugify(name: str) -> str:
    return re.sub(r"[^\w\-]", "-", name).lower()


# ---------------------------------------------------------------------------
# Sanitizer — maskeert organisatie-specifieke connectiestrings vóór output
# ---------------------------------------------------------------------------

_SHAREPOINT_URL_RE = re.compile(
    r'https?://[a-zA-Z0-9\-]+\.sharepoint\.com/sites/[^"\')\s,\]]+',
    re.IGNORECASE,
)

# Bestandsnamen die intern zijn maar niet gevoelig (bijv. swt_producten.xlsx) laten we staan.
# Alleen de volledige URL wordt gemaskeerd.

def sanitize(text: str) -> str:
    """
    Maskeert SharePoint-URLs in Power Query / DAX expressions vóór output.

    Wat er gebeurt:
      https://organisatie.sharepoint.com/sites/SITENAAM/Gedeelde%20documenten/map/bestand.xlsx
      → https://[TENANT].sharepoint.com/sites/[SITE]/Gedeelde documenten/map/bestand.xlsx

    De tenant-naam en sitenaam worden vervangen; het pad daarna blijft leesbaar
    zodat duidelijk is welke map/bestand de bron is.
    """
    if not text:
        return text

    def _replace(m: re.Match) -> str:
        url = m.group(0)
        try:
            parsed = urllib.parse.urlparse(url)
            # Decodeer %20 etc. zodat het pad leesbaar is
            path_decoded = urllib.parse.unquote(parsed.path)
            parts = path_decoded.strip("/").split("/")
            # parts[0] = "sites", parts[1] = sitenaam, parts[2:] = rest van het pad
            if len(parts) >= 2 and parts[0].lower() == "sites":
                rest = "/" + "/".join(parts[2:]) if len(parts) > 2 else ""
                masked = f"https://[TENANT].sharepoint.com/sites/[SITE]{rest}"
            else:
                masked = "https://[TENANT].sharepoint.com/[PAD]"
        except Exception:
            masked = "[SHAREPOINT-URL]"
        return masked

    return _SHAREPOINT_URL_RE.sub(_replace, text)


# ---------------------------------------------------------------------------
# TMSL parser (.bim format)
# ---------------------------------------------------------------------------

class TMSLParser:
    def __init__(self, bim_path: Path):
        self.data = read_json(bim_path)

    def _model(self) -> dict:
        return self.data.get("model", self.data)

    def tables(self) -> list[dict]:
        return [t for t in self._model().get("tables", []) if not t.get("isHidden", False)]

    def relationships(self) -> list[dict]:
        return self._model().get("relationships", [])

    def roles(self) -> list[dict]:
        return self._model().get("roles", [])

    def get_measures(self, table: dict) -> list[dict]:
        return table.get("measures", [])

    def get_columns(self, table: dict) -> list[dict]:
        return [c for c in table.get("columns", []) if c.get("type") != "rowNumber"]

    def get_partitions(self, table: dict) -> list[dict]:
        return table.get("partitions", [])

    def shared_expressions(self) -> list[dict]:
        return []


# ---------------------------------------------------------------------------
# TMDL parser (.tmdl format — Tabular Model Definition Language)
# ---------------------------------------------------------------------------

class TMLDParser:
    def __init__(self, definition_dir: Path):
        self.definition_dir = definition_dir
        self._tables: list[dict] = []
        self._relationships: list[dict] = []
        self._roles: list[dict] = []
        self._parse_all()

    def _parse_all(self):
        tables_dir = self.definition_dir / "tables"
        if tables_dir.exists():
            for tmdl_file in sorted(tables_dir.glob("*.tmdl")):
                table = self._parse_table_file(tmdl_file)
                if table and not table.get("isHidden", False):
                    self._tables.append(table)

        rel_file = self.definition_dir / "relationships.tmdl"
        if rel_file.exists():
            self._relationships = self._parse_relationships(rel_file)

        roles_dir = self.definition_dir / "roles"
        if roles_dir.exists():
            for role_file in sorted(roles_dir.glob("*.tmdl")):
                role = self._parse_role_file(role_file)
                if role:
                    self._roles.append(role)

        expr_file = self.definition_dir / "expressions.tmdl"
        if expr_file.exists():
            self._shared_expressions = self._parse_expressions_tmdl(expr_file)
        else:
            self._shared_expressions = []

    def _lines(self, path: Path) -> list[str]:
        return path.read_text(encoding="utf-8-sig").splitlines()

    def _parse_table_file(self, path: Path) -> Optional[dict]:
        lines = self._lines(path)
        if not lines:
            return None

        table: dict = {"name": "", "isHidden": False, "columns": [], "measures": [], "partitions": []}
        i = 0

        while i < len(lines) and not lines[i].strip():
            i += 1
        if i >= len(lines):
            return None

        header = lines[i].strip()
        if header.startswith("table "):
            table["name"] = header[6:].strip()
        else:
            table["name"] = path.stem
        i += 1

        current_block: list[str] = []
        current_type: str = ""
        current_name: str = ""

        def flush_block():
            if current_type == "measure":
                table["measures"].append(self._parse_measure_block(current_name, current_block))
            elif current_type == "column":
                col = self._parse_column_block(current_name, current_block)
                if col:
                    table["columns"].append(col)
            elif current_type == "partition":
                table["partitions"].append(self._parse_partition_block(current_name, current_block))

        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            if stripped.startswith("measure ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "measure"
                raw_name = stripped[8:].split("=")[0].strip().strip("'")
                current_name = raw_name
                current_block = [line]
            elif stripped.startswith("column ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "column"
                current_name = stripped[7:].strip().strip("'")
                current_block = [line]
            elif stripped.startswith("partition ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "partition"
                current_name = stripped[10:].strip().strip("'")
                current_block = [line]
            elif stripped.startswith("isHidden:") and line.startswith("\t") and not line.startswith("\t\t"):
                table["isHidden"] = "true" in stripped.lower()
            elif current_type:
                current_block.append(line)
            i += 1

        flush_block()
        return table

    def _parse_measure_block(self, name: str, block: list[str]) -> dict:
        dax_lines: list[str] = []
        description = ""
        format_string = ""

        SIBLING_PROPS = (
            "formatString:", "lineageTag:", "description:", "displayFolder:",
            "annotation ", "isHidden:", "kpiStatusType:", "dataCategory:",
        )

        in_dax = False

        for line in block:
            stripped = line.strip()

            if stripped.startswith("measure ") and "=" in stripped:
                parts = stripped.split("=", 1)
                inline = parts[1].strip() if len(parts) == 2 else ""
                if inline:
                    dax_lines.append(inline)
                in_dax = True
                continue

            if stripped and any(stripped.startswith(p) for p in SIBLING_PROPS):
                in_dax = False
                if stripped.startswith("formatString:"):
                    format_string = stripped.split(":", 1)[1].strip().strip("'\"")
                elif stripped.startswith("description:"):
                    description = stripped.split(":", 1)[1].strip().strip("'\"")
                continue

            if in_dax:
                clean = stripped.strip("`").strip()
                dax_lines.append(clean if clean else "")

        while dax_lines and not dax_lines[0]:
            dax_lines.pop(0)
        while dax_lines and not dax_lines[-1]:
            dax_lines.pop()

        return {
            "name": name,
            "expression": "\n".join(dax_lines).strip(),
            "description": description,
            "formatString": format_string,
        }

    def _parse_column_block(self, name: str, block: list[str]) -> Optional[dict]:
        col: dict = {"name": name, "dataType": "unknown", "isHidden": False, "description": ""}
        for line in block:
            if "dataType:" in line:
                col["dataType"] = line.split(":", 1)[1].strip()
            elif "isHidden:" in line:
                col["isHidden"] = "true" in line.lower()
            elif "description:" in line:
                col["description"] = line.split(":", 1)[1].strip().strip("'\"")
        return col

    def _parse_partition_block(self, name: str, block: list[str]) -> dict:
        partition: dict = {"name": name, "mode": "import", "source": {}}
        m_lines: list[str] = []
        in_source = False

        for line in block:
            stripped = line.strip()
            if stripped.startswith("mode:"):
                partition["mode"] = stripped.split(":", 1)[1].strip()
            elif stripped == "source =" or stripped.startswith("source ="):
                in_source = True
                rest = stripped[len("source ="):].strip().strip("`")
                if rest:
                    m_lines.append(rest)
            elif in_source:
                sibling_keywords = {"mode:", "annotation", "description:", "sourceColumn:", "isHidden:"}
                if any(stripped.startswith(k) for k in sibling_keywords):
                    in_source = False
                    continue
                clean = stripped.strip("`")
                if clean:
                    m_lines.append(clean)

        if m_lines:
            partition["source"] = {"expression": "\n".join(m_lines).strip()}

        return partition

    def _parse_expressions_tmdl(self, path: Path) -> list[dict]:
        lines = self._lines(path)
        expressions: list[dict] = []
        current_name = ""
        current_kind = ""
        m_lines: list[str] = []
        in_expr = False

        for line in lines:
            stripped = line.strip()

            if stripped.startswith("expression "):
                if current_name and m_lines:
                    expressions.append({
                        "name": current_name,
                        "kind": current_kind,
                        "expression": "\n".join(m_lines).strip(),
                    })
                name_part = stripped[len("expression "):].split("=")[0].strip().strip("'\"")
                current_name = name_part
                current_kind = "expression"
                m_lines = []
                in_expr = True
                if "=" in stripped:
                    rest = stripped.split("=", 1)[1].strip().strip("`")
                    if rest:
                        m_lines.append(rest)
            elif stripped.startswith("kind:"):
                current_kind = stripped.split(":", 1)[1].strip()
            elif in_expr and stripped and not stripped.startswith("kind:") and not stripped.startswith("lineageTag:") and not stripped.startswith("annotation"):
                m_lines.append(line.rstrip().strip("`"))
            elif stripped.startswith("lineageTag:") or stripped.startswith("annotation"):
                in_expr = False

        if current_name and m_lines:
            expressions.append({
                "name": current_name,
                "kind": current_kind,
                "expression": "\n".join(m_lines).strip(),
            })

        return expressions

    def shared_expressions(self) -> list[dict]:
        return getattr(self, "_shared_expressions", [])

    def _parse_relationships(self, path: Path) -> list[dict]:
        lines = self._lines(path)
        rels = []
        current: dict = {}

        def split_table_column(value: str):
            value = value.strip()
            m = re.match(r"^'([^']+)'\.'([^']+)'$", value)
            if m:
                return m.group(1), m.group(2)
            m = re.match(r"^'([^']+)'\.([^\s]+)$", value)
            if m:
                return m.group(1), m.group(2)
            m = re.match(r"^([^'.]+)\.'([^']+)'$", value)
            if m:
                return m.group(1), m.group(2)
            if "." in value:
                parts = value.split(".", 1)
                return parts[0], parts[1]
            return value, ""

        for line in lines:
            stripped = line.strip()
            if stripped.startswith("relationship "):
                if current:
                    rels.append(current)
                current = {"name": stripped[13:].strip()}
            elif stripped.startswith("fromTable:"):
                current["fromTable"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("fromColumn:"):
                raw = stripped.split(":", 1)[1].strip()
                if "." in raw:
                    current["fromTable"], current["fromColumn"] = split_table_column(raw)
                else:
                    current["fromColumn"] = raw
            elif stripped.startswith("toTable:"):
                current["toTable"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("toColumn:"):
                raw = stripped.split(":", 1)[1].strip()
                if "." in raw:
                    current["toTable"], current["toColumn"] = split_table_column(raw)
                else:
                    current["toColumn"] = raw
            elif stripped.startswith("crossFilteringBehavior:"):
                current["crossFilter"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("cardinality:") or stripped.startswith("fromCardinality:"):
                current["cardinality"] = stripped.split(":", 1)[1].strip()

        if current:
            rels.append(current)
        return rels

    def _parse_role_file(self, path: Path) -> Optional[dict]:
        lines = self._lines(path)
        role: dict = {"name": path.stem, "tablePermissions": []}

        for line in lines:
            stripped = line.strip()
            if stripped.startswith("role "):
                role["name"] = stripped[5:].strip()
            elif stripped.startswith("tablePermission "):
                table_name = stripped[16:].strip()
                role["tablePermissions"].append({"table": table_name, "filterExpression": ""})

        return role

    def tables(self) -> list[dict]:
        return self._tables

    def relationships(self) -> list[dict]:
        return self._relationships

    def roles(self) -> list[dict]:
        return self._roles

    def get_measures(self, table: dict) -> list[dict]:
        return table.get("measures", [])

    def get_columns(self, table: dict) -> list[dict]:
        return [c for c in table.get("columns", []) if not c.get("isHidden", False) or True]

    def get_partitions(self, table: dict) -> list[dict]:
        return table.get("partitions", [])


# ---------------------------------------------------------------------------
# Report parser (report.json)
# ---------------------------------------------------------------------------

def _extract_field_name(field: dict) -> Optional[str]:
    if not isinstance(field, dict):
        return None

    if "Measure" in field:
        m = field["Measure"]
        entity = m.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        prop = m.get("Property", "")
        return f"{entity}[{prop}]" if entity else f"[{prop}]"

    if "Column" in field:
        c = field["Column"]
        entity = c.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        prop = c.get("Property", "")
        return f"{entity}[{prop}]" if entity else f"[{prop}]"

    if "Aggregation" in field:
        agg = field["Aggregation"]
        func_map = {0: "SUM", 1: "AVG", 2: "COUNT", 3: "MIN", 4: "MAX",
                    5: "COUNTROWS", 6: "COUNTDISTINCT"}
        func = func_map.get(agg.get("Function", -1), "AGG")
        expr = agg.get("Expression", {})
        inner = _extract_field_name(expr)
        return f"{func}({inner})" if inner else func

    if "HierarchyLevel" in field:
        hl = field["HierarchyLevel"]
        expr = hl.get("Expression", {})
        hier_expr = expr.get("Hierarchy", {})
        entity = hier_expr.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        hier_name = hier_expr.get("Hierarchy", "")
        level = hl.get("Level", "")
        parts = [p for p in [entity, hier_name, level] if p]
        return ".".join(parts)

    if "SourceRef" in field:
        return field["SourceRef"].get("Entity", "")

    return None


def _parse_visual_config(config: dict) -> dict:
    sv = config.get("visual", config.get("singleVisual", {}))
    visual_type = sv.get("visualType", "unknown")

    title = ""
    for obj_key in ("objects", "vcObjects"):
        title_obj = sv.get(obj_key, {}).get("title", [])
        if isinstance(title_obj, list) and title_obj:
            props = title_obj[0].get("properties", {})
            title_text = props.get("text", {})
            if isinstance(title_text, dict):
                val = title_text.get("expr", {}).get("Literal", {}).get("Value", "")
                title = val.strip("'\"")
            elif isinstance(title_text, str):
                title = title_text
            if title:
                break

    fields: list[dict] = []

    query_state = sv.get("query", {}).get("queryState", {})
    if query_state:
        for role_name, role_data in query_state.items():
            for proj in role_data.get("projections", []):
                field_ref = proj.get("field", {})
                native_ref = proj.get("nativeQueryRef", proj.get("queryRef", ""))
                field_name = _extract_field_name(field_ref) or native_ref
                if field_name:
                    fields.append({
                        "field": field_name,
                        "role": role_name,
                        "displayName": native_ref,
                    })

    if not fields:
        proto_select = sv.get("prototypeQuery", {}).get("Select", [])
        dt_selects = sv.get("dataTransforms", {}).get("selects", [])
        dt_lookup: dict[str, dict] = {}
        for dt in dt_selects:
            qname = dt.get("queryName", "")
            roles_dict = dt.get("roles", {})
            role_name = next(iter(roles_dict.keys()), "") if roles_dict else ""
            dt_lookup[qname] = {"role": role_name, "displayName": dt.get("displayName", "")}

        for i, item in enumerate(proto_select):
            field_name = _extract_field_name(item)
            if not field_name:
                continue
            q_name = item.get("Name", "")
            meta = dt_lookup.get(q_name, dt_selects[i] if i < len(dt_selects) else {})
            roles_dict = meta.get("roles", {}) if isinstance(meta, dict) else {}
            fields.append({
                "field": field_name,
                "role": next(iter(roles_dict.keys()), meta.get("role", "")),
                "displayName": meta.get("displayName", ""),
            })

    return {
        "visualType": visual_type,
        "title": title,
        "fields": fields,
    }


def parse_report(report_json_path: Path) -> dict:
    data = read_json(report_json_path)
    result = {"pages": []}

    for section in data.get("sections", []):
        page = {
            "name": section.get("displayName", section.get("name", "Unknown")),
            "visuals": [],
        }

        for vc in section.get("visualContainers", []):
            config_str = vc.get("config", "{}")
            try:
                config = json.loads(config_str) if isinstance(config_str, str) else config_str
            except json.JSONDecodeError:
                config = {}

            if "singleVisual" not in config:
                continue

            visual_info = _parse_visual_config(config)

            if visual_info["visualType"] in ("", "unknown", "shape", "image", "textbox"):
                if not visual_info["fields"]:
                    continue

            page["visuals"].append(visual_info)

        result["pages"].append(page)

    return result


def parse_report_definition(definition_dir: Path) -> dict:
    result = {"pages": []}
    pages_dir = definition_dir / "pages"

    if not pages_dir.exists():
        return result

    page_order: list[str] = []
    pages_meta_file = pages_dir / "pages.json"
    if pages_meta_file.exists():
        try:
            pages_meta = read_json(pages_meta_file)
            page_order = pages_meta.get("pageOrder", [])
        except Exception:
            pass

    all_page_dirs = {d.name: d for d in pages_dir.iterdir() if d.is_dir()}
    ordered_hashes = page_order + [h for h in sorted(all_page_dirs) if h not in page_order]

    for page_hash in ordered_hashes:
        page_dir = all_page_dirs.get(page_hash)
        if not page_dir:
            continue

        page_json = page_dir / "page.json"
        display_name = page_hash
        if page_json.exists():
            try:
                page_meta = read_json(page_json)
                display_name = page_meta.get("displayName", page_hash)
            except Exception:
                pass

        page = {"name": display_name, "visuals": []}
        visuals_dir = page_dir / "visuals"

        if visuals_dir.exists():
            for visual_dir in sorted(visuals_dir.iterdir()):
                if not visual_dir.is_dir():
                    continue
                visual_json = visual_dir / "visual.json"
                if not visual_json.exists():
                    continue

                try:
                    visual_data = read_json(visual_json)
                except Exception:
                    continue

                visual_info = _parse_visual_config(visual_data)

                vtype = visual_info["visualType"]
                if vtype in ("", "unknown", "shape", "image", "textbox"):
                    if not visual_info["fields"]:
                        continue

                page["visuals"].append(visual_info)

        result["pages"].append(page)

    return result


# ---------------------------------------------------------------------------
# Markdown renderer
# ---------------------------------------------------------------------------

def render_markdown(
    project_name: str,
    parser,
    report_data: Optional[dict],
    model_format: str,
) -> str:
    lines: list[str] = []

    def h1(t): lines.extend([f"# {t}", ""])
    def h2(t): lines.extend([f"## {t}", ""])
    def h3(t): lines.extend([f"### {t}", ""])
    def p(t=""): lines.append(t)

    h1(f"Power BI Documentation: {project_name}")
    p(f"_Semantic model format: **{model_format}**_")
    p()

    h2("Table of Contents")
    p("1. [Data Model Overview](#data-model-overview)")
    p("2. [Tables & Columns](#tables--columns)")
    p("3. [Measure Catalogue](#measure-catalogue)")
    p("4. [Relationships](#relationships)")
    if parser.roles():
        p("5. [Row-Level Security](#row-level-security)")
    if report_data:
        p("6. [Report Structure](#report-structure)")
    p()

    h2("Data Model Overview")
    tables = parser.tables()
    total_measures = sum(len(parser.get_measures(t)) for t in tables)
    total_columns = sum(len(parser.get_columns(t)) for t in tables)
    total_rels = len(parser.relationships())

    p(f"| Metric | Count |")
    p(f"|--------|-------|")
    p(f"| Tables | {len(tables)} |")
    p(f"| Columns | {total_columns} |")
    p(f"| Measures | {total_measures} |")
    p(f"| Relationships | {total_rels} |")
    p(f"| RLS Roles | {len(parser.roles())} |")
    p()

    h2("Tables & Columns")
    for table in tables:
        table_name = table.get("name", "Unknown")
        h3(table_name)

        columns = parser.get_columns(table)
        measures = parser.get_measures(table)
        partitions = parser.get_partitions(table)

        if partitions:
            source_info = ""
            for pt in partitions:
                source = pt.get("source", {})
                if isinstance(source, dict):
                    query = source.get("query", source.get("expression", ""))
                    if query:
                        # ↓ sanitize vóór output
                        query_clean = sanitize(query)
                        source_info = f"`{query_clean[:120]}{'...' if len(query_clean) > 120 else ''}`"
            if source_info:
                p(f"**Source:** {source_info}")
                p()

        if columns:
            p("| Column | Data Type | Hidden | Description |")
            p("|--------|-----------|--------|-------------|")
            for col in columns:
                name = col.get("name", "")
                dtype = col.get("dataType", col.get("type", "unknown"))
                hidden = "✓" if col.get("isHidden") else ""
                desc = col.get("description", "")
                p(f"| {name} | {dtype} | {hidden} | {desc} |")
            p()
        else:
            p("_No columns defined (may be a calculated or virtual table)._")
            p()

        if measures:
            p(f"**{len(measures)} measure(s)** — see Measure Catalogue below.")
            p()

    h2("Measure Catalogue")
    has_measures = False
    for table in tables:
        measures = parser.get_measures(table)
        if not measures:
            continue
        has_measures = True
        table_name = table.get("name", "Unknown")
        h3(f"Table: {table_name}")

        for measure in measures:
            name = measure.get("name", "")
            dax = measure.get("expression", "").strip()
            desc = measure.get("description", "")
            fmt = measure.get("formatString", "")

            p(f"#### `{name}`")
            if desc:
                p(f"_{desc}_")
                p()
            if fmt:
                p(f"**Format:** `{fmt}`")
                p()
            if dax:
                p("```dax")
                # DAX bevat normaal geen URLs, maar voor de zekerheid
                p(sanitize(dax))
                p("```")
            else:
                p("_No DAX expression found._")
            p()

    if not has_measures:
        p("_No measures found in this model._")
        p()

    h2("Relationships")
    rels = parser.relationships()
    if rels:
        p("| From Table | From Column | To Table | To Column | Cardinality | Cross Filter |")
        p("|------------|-------------|----------|-----------|-------------|--------------|")
        for rel in rels:
            from_t = rel.get("fromTable", rel.get("fromTableId", ""))
            from_c = rel.get("fromColumn", rel.get("fromColumnId", ""))
            to_t = rel.get("toTable", rel.get("toTableId", ""))
            to_c = rel.get("toColumn", rel.get("toColumnId", ""))
            card = rel.get("cardinality", "many-to-one")
            cf = rel.get("crossFilter", rel.get("crossFilteringBehavior", ""))
            p(f"| {from_t} | {from_c} | {to_t} | {to_c} | {card} | {cf} |")
        p()
    else:
        p("_No relationships defined._")
        p()

    roles = parser.roles()
    if roles:
        h2("Row-Level Security")
        for role in roles:
            role_name = role.get("name", "")
            h3(role_name)
            members = role.get("members", [])
            if members:
                p(f"**Members:** {', '.join(str(m) for m in members)}")
                p()
            perms = role.get("tablePermissions", [])
            model_perms = role.get("modelPermission", "")
            if model_perms:
                p(f"**Model permission:** `{model_perms}`")
                p()
            if perms:
                p("| Table | Filter Expression |")
                p("|-------|------------------|")
                for perm in perms:
                    t = perm.get("table", perm.get("name", ""))
                    expr = perm.get("filterExpression", perm.get("expression", "_none_"))
                    p(f"| {t} | `{expr}` |")
                p()

    if report_data:
        h2("Report Structure")
        pages = report_data.get("pages", [])
        p(f"Total pages: **{len(pages)}**")
        p()
        for page in pages:
            h3(page["name"])
            visuals = page.get("visuals", [])
            if not visuals:
                p("_No visual information extracted._")
                p()
                continue

            type_counts = Counter(v["visualType"] for v in visuals)
            p(f"**{len(visuals)} visual(s):** " + ", ".join(f"{cnt}× {vt}" for vt, cnt in sorted(type_counts.items())))
            p()

            for i, visual in enumerate(visuals, 1):
                vtype = visual["visualType"]
                title = visual.get("title", "")
                fields = visual.get("fields", [])

                header = f"**Visual {i}: `{vtype}`**"
                if title:
                    header += f" — _{title}_"
                p(header)

                if fields:
                    p("| Role | Field | Display Name |")
                    p("|------|-------|--------------|")
                    for f in fields:
                        role = f.get("role", "")
                        field = f.get("field", "")
                        display = f.get("displayName", "")
                        p(f"| {role} | `{field}` | {display} |")
                else:
                    p("_No field bindings extracted._")
                p()

    h2("Power Query (M) Sources")
    has_pq = False
    for table in tables:
        partitions = parser.get_partitions(table)
        for pt in partitions:
            source = pt.get("source", {})
            expr = source.get("expression", "") if isinstance(source, dict) else source
            if not expr:
                continue
            has_pq = True
            table_name = table.get("name", "Unknown")
            h3(table_name)
            p(f"**Mode:** `{pt.get('mode', 'import')}`")
            p()
            p("```powerquery")
            # ↓ sanitize vóór output
            p(sanitize(expr.strip()))
            p("```")
            p()

    if not has_pq:
        p("_No Power Query expressions found (model may use DirectQuery or live connection)._")
        p()

    shared = parser.shared_expressions()
    if shared:
        h2("Shared Power Query Functions & Parameters")
        p("These are reusable queries/functions defined in `expressions.tmdl`, "
          "typically used as helper functions called from table queries above.")
        p()
        for expr in shared:
            name = expr.get("name", "")
            kind = expr.get("kind", "")
            body = expr.get("expression", "").strip()
            h3(name)
            if kind:
                p(f"**Kind:** `{kind}`")
                p()
            if body:
                p("```powerquery")
                # ↓ sanitize vóór output
                p(sanitize(body))
                p("```")
            p()

    p("---")
    p("_Generated by pbip_document.py_")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# DAX dependency analyser
# ---------------------------------------------------------------------------

def extract_dax_refs(dax: str, all_tables: list[str], all_measures: list[str]) -> dict:
    refs: dict = {"tables": set(), "columns": set(), "measures": set()}

    for match in re.finditer(r"'?([A-Za-z0-9_ ]+)'?\[([^\]]+)\]", dax):
        t, c = match.group(1).strip(), match.group(2).strip()
        if t:
            refs["tables"].add(t)
            refs["columns"].add(f"{t}[{c}]")

    for match in re.finditer(r"(?<!')\[([^\]]+)\]", dax):
        name = match.group(1).strip()
        if name in all_measures:
            refs["measures"].add(name)

    for t in all_tables:
        if not t:
            continue
        pattern = rf"\b{re.escape(t)}\b"
        if re.search(pattern, dax, re.IGNORECASE):
            refs["tables"].add(t)

    return {k: sorted(v) for k, v in refs.items()}


def describe_dax(dax: str) -> str:
    dax_upper = dax.strip().upper()
    patterns = [
        (r"^CALCULATE\s*\(", "A CALCULATE expression that evaluates a value in a modified filter context."),
        (r"^DIVIDE\s*\(", "A safe division (DIVIDE) that returns a result or an alternate value on divide-by-zero."),
        (r"^SUM\s*\(", "A simple SUM aggregation over a column."),
        (r"^SUMX\s*\(", "A row-by-row SUMX iterator that sums an expression over a table."),
        (r"^COUNT\s*\(", "A COUNT aggregation."),
        (r"^COUNTROWS\s*\(", "Counts the number of rows in a table or filtered table."),
        (r"^COUNTX\s*\(", "A row-by-row COUNTX iterator."),
        (r"^DISTINCTCOUNT\s*\(", "Counts the number of distinct values in a column."),
        (r"^AVERAGE\s*\(", "An AVERAGE aggregation over a column."),
        (r"^AVERAGEX\s*\(", "A row-by-row AVERAGEX iterator."),
        (r"^IF\s*\(", "A conditional expression (IF) that returns one of two values based on a condition."),
        (r"^SWITCH\s*\(", "A SWITCH expression that returns different values based on matching conditions."),
        (r"^VAR\s+", "Uses variables (VAR/RETURN) to break the calculation into named intermediate steps."),
        (r"^FILTER\s*\(", "A FILTER expression that returns a filtered table."),
        (r"^ALL\s*\(", "Removes filters using ALL() to return all rows or all values."),
        (r"^RELATED\s*\(", "Looks up a value from a related table via a relationship."),
        (r"^LOOKUPVALUE\s*\(", "Looks up a value based on matching criteria (LOOKUPVALUE)."),
        (r"^FORMAT\s*\(", "Formats a value as text using a format string."),
        (r"^DATEADD\s*\(", "A time intelligence calculation shifting dates by a specified interval."),
        (r"^SAMEPERIODLASTYEAR\s*\(", "A time intelligence calculation comparing to the same period last year."),
        (r"^TOTALYTD\s*\(", "Calculates a year-to-date total."),
        (r"^TOTALMTD\s*\(", "Calculates a month-to-date total."),
        (r"^TOTALQTD\s*\(", "Calculates a quarter-to-date total."),
        (r"^RANKX\s*\(", "Ranks values within a table using RANKX."),
        (r"^TOPN\s*\(", "Returns the top N rows of a table."),
        (r"^SELECTEDVALUE\s*\(", "Returns the selected value from a slicer or filter context (SELECTEDVALUE)."),
        (r"^HASONEVALUE\s*\(", "Checks whether a column is filtered to a single value."),
        (r"^ISBLANK\s*\(", "Checks whether a value is blank."),
        (r"^BLANK\s*\(", "Returns a BLANK value."),
        (r"^MIN\s*\(", "Returns the minimum value in a column."),
        (r"^MAX\s*\(", "Returns the maximum value in a column."),
        (r"^MAXX\s*\(", "A row-by-row MAXX iterator returning the maximum expression value."),
        (r"^MINX\s*\(", "A row-by-row MINX iterator returning the minimum expression value."),
        (r"^CONCATENATE\s*\(", "Concatenates two text strings."),
        (r"^CONCATENATEX\s*\(", "Concatenates text values across rows of a table."),
    ]
    for pattern, description in patterns:
        if re.match(pattern, dax_upper):
            return description
    return "A DAX expression."


# ---------------------------------------------------------------------------
# Copilot knowledge base renderer
# ---------------------------------------------------------------------------

def render_copilot_kb(
    project_name: str,
    parser,
    report_data: Optional[dict],
    model_format: str,
) -> str:
    tables = parser.tables()
    all_table_names = [t.get("name", "") for t in tables]
    all_measure_names = [
        m.get("name", "")
        for t in tables
        for m in parser.get_measures(t)
    ]

    lines: list[str] = []
    def w(s=""): lines.append(s)

    w("=" * 80)
    w(f"POWER BI KNOWLEDGE BASE: {project_name.upper()}")
    w(f"Semantic model format: {model_format}")
    w("=" * 80)
    w()
    w("INSTRUCTIONS FOR AI ASSISTANT")
    w("-" * 40)
    w("This file is the complete technical knowledge base for the Power BI project")
    w(f'"{project_name}". Use it to answer questions about:')
    w("- How specific measures or fields are calculated (including full DAX logic)")
    w("- What columns exist in which tables and what they represent")
    w("- How tables are related to each other")
    w("- Which measures depend on other measures or columns")
    w("- Which pages and visuals exist in the report, and which fields each visual uses")
    w()
    w("When answering questions about a measure, always explain:")
    w("1. What the measure calculates in plain language")
    w("2. The exact DAX formula")
    w("3. Which tables and columns it touches")
    w("4. Which other measures it calls (if any)")
    w("5. Which report visuals use this measure (if applicable)")
    w()

    w("=" * 80)
    w("INDEX")
    w("=" * 80)
    w(f"Tables ({len(tables)}): " + ", ".join(all_table_names))
    w(f"Total columns: {sum(len(parser.get_columns(t)) for t in tables)}")
    w(f"Total measures: {len(all_measure_names)}")
    w(f"Relationships: {len(parser.relationships())}")
    w(f"RLS Roles: {len(parser.roles())}")
    if report_data:
        pages = report_data.get("pages", [])
        total_visuals = sum(len(p.get("visuals", [])) for p in pages)
        w(f"Report pages: {len(pages)} ({total_visuals} visuals total)")
    w()

    w("=" * 80)
    w("SECTION 1: TABLES AND COLUMNS")
    w("=" * 80)
    w()

    for table in tables:
        table_name = table.get("name", "Unknown")
        columns = parser.get_columns(table)
        measures = parser.get_measures(table)
        partitions = parser.get_partitions(table)

        w(f"TABLE: {table_name}")
        w("-" * 60)

        for pt in partitions:
            source = pt.get("source", {})
            if isinstance(source, dict):
                query = source.get("query", source.get("expression", ""))
                if query:
                    w(f"This table is loaded from the following query:")
                    # ↓ sanitize vóór output
                    w(sanitize(query.strip()))
                    w()

        if columns:
            w(f"This table has {len(columns)} column(s):")
            w()
            for col in columns:
                col_name = col.get("name", "")
                dtype = col.get("dataType", col.get("type", "unknown"))
                desc = col.get("description", "")
                hidden = col.get("isHidden", False)

                line = f'  Column "{col_name}" (data type: {dtype})'
                if hidden:
                    line += " [hidden from report view]"
                if desc:
                    line += f": {desc}"
                w(line)
        else:
            w("This table has no regular columns (may be a calculated or virtual table).")

        if measures:
            w()
            w(f"This table contains {len(measures)} measure(s): "
              + ", ".join(f'"{m.get("name", "")}"' for m in measures))

        w()

    w("=" * 80)
    w("SECTION 2: MEASURES (FULL DAX AND DEPENDENCIES)")
    w("=" * 80)
    w()

    measure_visual_usage: dict[str, list[str]] = {}
    if report_data:
        for page in report_data.get("pages", []):
            page_name = page["name"]
            for visual in page.get("visuals", []):
                vtype = visual["visualType"]
                title = visual.get("title", "")
                label = f'"{title}" ({vtype}) on page "{page_name}"' if title else f'{vtype} on page "{page_name}"'
                for f in visual.get("fields", []):
                    field_str = f.get("field", "")
                    m = re.search(r"\[([^\]]+)\]", field_str)
                    if m:
                        fname = m.group(1)
                        if fname in all_measure_names:
                            measure_visual_usage.setdefault(fname, []).append(label)

    for table in tables:
        measures = parser.get_measures(table)
        if not measures:
            continue
        table_name = table.get("name", "Unknown")

        for measure in measures:
            name = measure.get("name", "")
            dax = measure.get("expression", "").strip()
            desc = measure.get("description", "")
            fmt = measure.get("formatString", "")

            w(f"MEASURE: {name}")
            w(f"Defined in table: {table_name}")
            w("-" * 60)

            if desc:
                w(f"Description: {desc}")
            if fmt:
                w(f"Display format: {fmt}")

            if dax:
                plain = describe_dax(dax)
                w(f"What it does: {plain}")
            w()

            if dax:
                w("Full DAX formula:")
                # ↓ sanitize vóór output
                w(sanitize(dax))
                w()

                deps = extract_dax_refs(dax, all_table_names, all_measure_names)

                if deps["measures"]:
                    w("This measure calls the following other measures:")
                    for m in deps["measures"]:
                        if m != name:
                            w(f'  - "{m}"')
                    w()

                if deps["columns"]:
                    w("This measure references the following columns:")
                    for c in deps["columns"]:
                        w(f"  - {c}")
                    w()

                if deps["tables"]:
                    w("This measure references the following tables:")
                    for t in deps["tables"]:
                        w(f"  - {t}")
                    w()
            else:
                w("No DAX expression was found for this measure.")
                w()

            usages = measure_visual_usage.get(name, [])
            if usages:
                w(f"This measure is used in {len(usages)} visual(s):")
                for usage in usages:
                    w(f"  - {usage}")
                w()

            w()

    w("=" * 80)
    w("SECTION 3: RELATIONSHIPS")
    w("=" * 80)
    w()

    rels = parser.relationships()
    if rels:
        w(f"The data model has {len(rels)} relationship(s):")
        w()
        for i, rel in enumerate(rels, 1):
            from_t = rel.get("fromTable", rel.get("fromTableId", ""))
            from_c = rel.get("fromColumn", rel.get("fromColumnId", ""))
            to_t = rel.get("toTable", rel.get("toTableId", ""))
            to_c = rel.get("toColumn", rel.get("toColumnId", ""))
            card = rel.get("cardinality", rel.get("fromCardinality", "many-to-one"))
            cf = rel.get("crossFilter", rel.get("crossFilteringBehavior", "single"))

            w(f"Relationship {i}:")
            w(f'  The column "{from_c}" in table "{from_t}" is linked to '
              f'the column "{to_c}" in table "{to_t}".')
            w(f"  Cardinality: {card}. Cross-filter direction: {cf}.")
            w()
    else:
        w("No relationships are defined in this model.")
        w()

    roles = parser.roles()
    if roles:
        w("=" * 80)
        w("SECTION 4: ROW-LEVEL SECURITY (RLS)")
        w("=" * 80)
        w()
        for role in roles:
            role_name = role.get("name", "")
            perms = role.get("tablePermissions", [])
            w(f"RLS Role: {role_name}")
            if perms:
                for perm in perms:
                    t = perm.get("table", perm.get("name", ""))
                    expr = perm.get("filterExpression", perm.get("expression", ""))
                    if expr:
                        w(f'  Table "{t}" is filtered by: {expr}')
                    else:
                        w(f'  Table "{t}" is included in this role (no filter expression).')
            w()

    if report_data:
        w("=" * 80)
        w("SECTION 5: REPORT STRUCTURE AND VISUAL FIELD BINDINGS")
        w("=" * 80)
        w()
        pages = report_data.get("pages", [])
        w(f"The report has {len(pages)} page(s).")
        w()
        for page in pages:
            page_name = page["name"]
            visuals = page.get("visuals", [])
            w(f"Page: {page_name} ({len(visuals)} visual(s))")
            w("-" * 60)

            if not visuals:
                w("  No visual information available.")
                w()
                continue

            for i, visual in enumerate(visuals, 1):
                vtype = visual["visualType"]
                title = visual.get("title", "")
                fields = visual.get("fields", [])

                label = f'Visual {i}: {vtype}'
                if title:
                    label += f' (title: "{title}")'
                w(f"  {label}")

                if fields:
                    for f in fields:
                        role = f.get("role", "")
                        field = f.get("field", "")
                        display = f.get("displayName", "")
                        role_str = f" [{role}]" if role else ""
                        display_str = f' (shown as "{display}")' if display and display != field else ""
                        w(f"    - {field}{role_str}{display_str}")
                else:
                    w("    No field bindings extracted.")
                w()

    w("=" * 80)
    w("SECTION 6: POWER QUERY (M) SOURCES")
    w("=" * 80)
    w()

    has_pq = False
    for table in tables:
        partitions = parser.get_partitions(table)
        for pt in partitions:
            source = pt.get("source", {})
            expr = source.get("expression", "") if isinstance(source, dict) else source
            if not expr:
                continue
            has_pq = True
            table_name = table.get("name", "Unknown")
            w(f"Table: {table_name} (mode: {pt.get('mode', 'import')})")
            w("-" * 60)
            w("Power Query expression:")
            # ↓ sanitize vóór output
            w(sanitize(expr.strip()))
            w()

    if not has_pq:
        w("No Power Query expressions found (model may use DirectQuery or live connection).")
        w()

    shared = parser.shared_expressions()
    if shared:
        w("=" * 80)
        w("SECTION 7: SHARED POWER QUERY FUNCTIONS AND PARAMETERS")
        w("=" * 80)
        w()
        for expr in shared:
            name = expr.get("name", "")
            kind = expr.get("kind", "")
            body = expr.get("expression", "").strip()
            w(f"Shared expression: {name} (kind: {kind})")
            w("-" * 60)
            if body:
                # ↓ sanitize vóór output
                w(sanitize(body))
            w()

    w("=" * 80)
    w("END OF KNOWLEDGE BASE")
    w("=" * 80)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Project discovery
# ---------------------------------------------------------------------------

def find_semantic_model(pbip_root: Path) -> tuple[Optional[Path], str]:
    for candidate in pbip_root.rglob("database.tmdl"):
        return candidate.parent, "TMDL"

    for candidate in pbip_root.rglob("definition"):
        if candidate.is_dir() and any(candidate.glob("*.tmdl")):
            return candidate, "TMDL"

    for candidate in pbip_root.rglob("*.bim"):
        return candidate, "TMSL"

    return None, ""


def find_report_json(pbip_root: Path) -> Optional[Path]:
    for candidate in pbip_root.rglob("report.json"):
        try:
            data = read_json(candidate)
            if data.get("sections"):
                return candidate
        except Exception:
            pass
    return None


def find_report_definition_dir(pbip_root: Path) -> Optional[Path]:
    for candidate in pbip_root.rglob("pages.json"):
        pages_dir = candidate.parent
        if pages_dir.name == "pages":
            return pages_dir.parent
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Document a Power BI .pbip project")
    ap.add_argument("pbip_path", help="Path to the .pbip project folder")
    ap.add_argument("--output", "-o", default="", help="Output file path")
    ap.add_argument(
        "--copilot",
        action="store_true",
        help=(
            "Generate a plain-text knowledge base optimised for Copilot / ChatGPT "
            "context injection instead of Markdown docs."
        ),
    )
    args = ap.parse_args()

    pbip_root = Path(args.pbip_path).resolve()
    if not pbip_root.exists():
        print(f"ERROR: Path does not exist: {pbip_root}", file=sys.stderr)
        sys.exit(1)

    project_name = pbip_root.name.replace(".pbip", "").replace("_", " ").title()

    model_path, model_format = find_semantic_model(pbip_root)
    if not model_path:
        print("ERROR: No semantic model found (.bim or TMDL definition folder).", file=sys.stderr)
        sys.exit(1)

    print(f"Found semantic model ({model_format}): {model_path}")

    if model_format == "TMSL":
        parser = TMSLParser(model_path)
    else:
        parser = TMLDParser(model_path)

    report_data = None

    report_def_dir = find_report_definition_dir(pbip_root)
    if report_def_dir:
        print(f"Found modern report definition: {report_def_dir}")
        try:
            report_data = parse_report_definition(report_def_dir)
        except Exception as e:
            print(f"WARNING: Could not parse modern report definition: {e}", file=sys.stderr)

    if not report_data or not report_data.get("pages"):
        report_json_path = find_report_json(pbip_root)
        if report_json_path:
            print(f"Found legacy report.json: {report_json_path}")
            try:
                report_data = parse_report(report_json_path)
            except Exception as e:
                print(f"WARNING: Could not parse report.json: {e}", file=sys.stderr)
        else:
            print("No report found — skipping report structure section.")

    if report_data and report_data.get("pages"):
        total_visuals = sum(len(p.get("visuals", [])) for p in report_data["pages"])
        total_fields = sum(
            len(v.get("fields", []))
            for p in report_data["pages"]
            for v in p.get("visuals", [])
        )
        print(f"  Pages: {len(report_data['pages'])}, Visuals: {total_visuals}, Field bindings: {total_fields}")

    tables = parser.tables()
    n_measures = sum(len(parser.get_measures(t)) for t in tables)

    if args.copilot:
        content = render_copilot_kb(project_name, parser, report_data, model_format)
        default_name = f"{project_name.replace(' ', '_')}_copilot_kb.txt"
        suffix = "Copilot knowledge base"
    else:
        content = render_markdown(project_name, parser, report_data, model_format)
        default_name = f"{project_name.replace(' ', '_')}_docs.md"
        suffix = "Markdown documentation"

    output_path = args.output or default_name
    out = Path(output_path)
    out.write_text(content, encoding="utf-8")

    print(f"\nDone. {suffix} written to: {out.resolve()}")
    print(f"  Tables:        {len(tables)}")
    print(f"  Measures:      {n_measures}")
    print(f"  Relationships: {len(parser.relationships())}")

    if args.copilot:
        size_kb = out.stat().st_size / 1024
        print(f"  File size:     {size_kb:.1f} KB")
        print()
        print("How to use:")
        print("  Copilot for M365: Upload this .txt file in a Teams/Word chat with Copilot.")
        print("  Claude Project:   Add this file as project knowledge.")
        print("  ChatGPT:          Upload via paperclip, or paste into a custom GPT system prompt.")
        print("  Copilot Studio:   Use as a knowledge source document.")


if __name__ == "__main__":
    main()
