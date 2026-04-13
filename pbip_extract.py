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
# TMSL parser (.bim format)
# ---------------------------------------------------------------------------

class TMSLParser:
    def __init__(self, bim_path: Path):
        self.data = read_json(bim_path)

    def _model(self) -> dict:
        return self.data.get("model", self.data)

    def tables(self) -> list[dict]:
        return self._model().get("tables", [])

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
    """
    Parses TMDL files. TMDL is an indentation-based DSL, not JSON.
    Each table lives in its own .tmdl file under SemanticModel/definition/tables/
    The database.tmdl contains top-level model metadata.
    relationships.tmdl contains relationships.
    roles/ contains RLS roles.
    """

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
                if table:
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

        # Shared Power Query expressions/functions (expressions.tmdl)
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

        table: dict = {"name": "", "columns": [], "measures": [], "partitions": []}
        i = 0

        # First non-empty line should be: table <name>
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

            # Detect new top-level blocks (indented 1 level under table)
            if stripped.startswith("measure ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "measure"
                current_name = stripped[8:].split("=")[0].strip()
                current_block = [stripped]
            elif stripped.startswith("column ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "column"
                current_name = stripped[7:].strip()
                current_block = [stripped]
            elif stripped.startswith("partition ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "partition"
                current_name = stripped[10:].strip()
                current_block = [stripped]
            elif current_type:
                current_block.append(stripped)
            i += 1

        flush_block()
        return table

    def _parse_measure_block(self, name: str, block: list[str]) -> dict:
        dax_lines: list[str] = []
        description = ""
        format_string = ""
        in_expression = False

        for line in block:
            if line.startswith("measure "):
                # Inline expression: measure Name = <expr>
                parts = line.split("=", 1)
                if len(parts) == 2 and parts[1].strip():
                    dax_lines.append(parts[1].strip())
            elif line.strip().startswith("expression:"):
                in_expression = True
                rest = line.split(":", 1)[1].strip().strip("`")
                if rest:
                    dax_lines.append(rest)
            elif in_expression and (line.startswith("\t\t") or line.startswith("  ")):
                dax_lines.append(line.strip().strip("`"))
            elif line.strip().startswith("description:"):
                in_expression = False
                description = line.split(":", 1)[1].strip().strip("'\"")
            elif line.strip().startswith("formatString:"):
                in_expression = False
                format_string = line.split(":", 1)[1].strip().strip("'\"")
            else:
                in_expression = False

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
        """
        Extract mode and M expression from a partition block.
        M expressions are multi-line, delimited by deeper indentation after 'source ='.
        They may be wrapped in triple-backticks or just indented.
        """
        partition: dict = {"name": name, "mode": "import", "source": {}}
        m_lines: list[str] = []
        in_source = False

        for line in block:
            stripped = line.strip()
            if stripped.startswith("mode:"):
                partition["mode"] = stripped.split(":", 1)[1].strip()
            elif stripped == "source =" or stripped.startswith("source ="):
                in_source = True
                # Inline expression on same line after '='
                rest = stripped[len("source ="):].strip().strip("`")
                if rest:
                    m_lines.append(rest)
            elif in_source:
                # Stop collecting when we hit a non-indented sibling keyword
                if stripped and not stripped.startswith("#") and ":" in stripped and not stripped.startswith("let") and not stripped.startswith("in ") and not stripped.startswith("//"):
                    # Heuristic: looks like a new TMDL property, not M code
                    sibling_keywords = {"mode:", "annotation", "description:", "sourceColumn:", "isHidden:"}
                    if any(stripped.startswith(k) for k in sibling_keywords):
                        in_source = False
                        continue
                clean = line.rstrip().strip("`")
                if clean:
                    m_lines.append(clean)

        if m_lines:
            partition["source"] = {"expression": "\n".join(m_lines).strip()}

        return partition

    def _parse_expressions_tmdl(self, path: Path) -> list[dict]:
        """
        Parse expressions.tmdl which contains shared Power Query functions and
        parameters (e.g. Transform File, Sample File, helper queries).
        Format:
            expression 'Name' =
                    let ... in ...
        """
        lines = self._lines(path)
        expressions: list[dict] = []
        current_name = ""
        current_kind = ""
        m_lines: list[str] = []
        in_expr = False

        for line in lines:
            stripped = line.strip()

            if stripped.startswith("expression "):
                # Flush previous
                if current_name and m_lines:
                    expressions.append({
                        "name": current_name,
                        "kind": current_kind,
                        "expression": "\n".join(m_lines).strip(),
                    })
                # Parse: expression 'Name' = or expression Name =
                name_part = stripped[len("expression "):].split("=")[0].strip().strip("'\"")
                current_name = name_part
                current_kind = "expression"
                m_lines = []
                in_expr = True
                # Inline expression
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

        # Flush last
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
            """Split TMDL 'Table.Column' notation into (table, column).
            Handles: dim_table.Column, 'My Table'.'My Column', dim_table.'My Column'
            """
            value = value.strip()
            # Match quoted or unquoted table name, followed by dot, then quoted or unquoted column
            m = re.match(r"^'([^']+)'\.'([^']+)'$", value)  # 'Table'.'Column'
            if m:
                return m.group(1), m.group(2)
            m = re.match(r"^'([^']+)'\.([^\s]+)$", value)   # 'Table'.Column
            if m:
                return m.group(1), m.group(2)
            m = re.match(r"^([^'.]+)\.'([^']+)'$", value)   # Table.'Column'
            if m:
                return m.group(1), m.group(2)
            if "." in value:                                  # Table.Column (no quotes)
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

def parse_report(report_json_path: Path) -> dict:
    """Extract pages and visual types from report.json"""
    data = read_json(report_json_path)
    result = {"pages": []}

    sections = data.get("sections", [])
    for section in sections:
        page = {
            "name": section.get("displayName", section.get("name", "Unknown")),
            "visuals": [],
        }
        for visual_container in section.get("visualContainers", []):
            config_str = visual_container.get("config", "{}")
            try:
                config = json.loads(config_str) if isinstance(config_str, str) else config_str
            except json.JSONDecodeError:
                config = {}

            visual_type = (
                config.get("singleVisual", {}).get("visualType")
                or config.get("vcObjects", {})
            )
            if isinstance(visual_type, str):
                page["visuals"].append(visual_type)

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

    # Table of contents
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

    # Data model overview
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

    # Tables & Columns
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
                        source_info = f"`{query[:120]}{'...' if len(query) > 120 else ''}`"
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
                desc = col.get("description", col.get("annotations", [{}])[0].get("value", "") if isinstance(col.get("annotations"), list) and col.get("annotations") else "")
                p(f"| {name} | {dtype} | {hidden} | {desc} |")
            p()
        else:
            p("_No columns defined (may be a calculated or virtual table)._")
            p()

        if measures:
            p(f"**{len(measures)} measure(s)** — see Measure Catalogue below.")
            p()

    # Measure catalogue
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
                p(dax)
                p("```")
            else:
                p("_No DAX expression found._")
            p()

    if not has_measures:
        p("_No measures found in this model._")
        p()

    # Relationships
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

    # RLS
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

    # Report structure
    if report_data:
        h2("Report Structure")
        pages = report_data.get("pages", [])
        p(f"Total pages: **{len(pages)}**")
        p()
        for page in pages:
            h3(page["name"])
            visuals = page.get("visuals", [])
            if visuals:
                from collections import Counter
                counts = Counter(visuals)
                p("| Visual Type | Count |")
                p("|-------------|-------|")
                for vtype, cnt in sorted(counts.items()):
                    p(f"| {vtype} | {cnt} |")
                p()
            else:
                p("_No visual type information extracted._")
                p()

    # Power Query
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
            p(expr.strip())
            p("```")
            p()

    if not has_pq:
        p("_No Power Query expressions found (model may use DirectQuery or live connection)._")
        p()

    # Shared expressions (expressions.tmdl)
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
                p(body)
                p("```")
            p()

    p("---")
    p("_Generated by pbip_document.py_")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# DAX dependency analyser
# ---------------------------------------------------------------------------

def extract_dax_refs(dax: str, all_tables: list[str], all_measures: list[str]) -> dict:
    """
    Pull referenced tables, columns, and measures out of a DAX expression.
    This is heuristic, not a full DAX parser — covers 95% of real-world patterns.
    Returns: {"tables": [...], "columns": [...], "measures": [...]}
    """
    refs: dict = {"tables": set(), "columns": set(), "measures": set()}

    # Table[Column] pattern
    for match in re.finditer(r"'?([A-Za-z0-9_ ]+)'?\[([^\]]+)\]", dax):
        t, c = match.group(1).strip(), match.group(2).strip()
        refs["tables"].add(t)
        refs["columns"].add(f"{t}[{c}]")

    # Standalone [Measure or Column] — match against known measure names
    for match in re.finditer(r"(?<!')\[([^\]]+)\]", dax):
        name = match.group(1).strip()
        if name in all_measures:
            refs["measures"].add(name)

    # CALCULATE, FILTER etc. with bare table names
    for t in all_tables:
        pattern = rf"\b{re.escape(t)}\b"
        if re.search(pattern, dax, re.IGNORECASE):
            refs["tables"].add(t)

    return {k: sorted(v) for k, v in refs.items()}


def describe_dax(dax: str) -> str:
    """
    Produce a one-line plain-English description of what a DAX expression does,
    based on the top-level function used. Not exhaustive but useful for LLM context.
    """
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
    """
    Generates a plain-text knowledge base optimised for LLM context injection.

    Design principles:
    - Every entity is described in full prose sentences, not tables.
    - Every measure gets: what it does, its full DAX, which columns/tables/measures it
      depends on, and its format string.
    - Every column gets: its table, data type, and any description.
    - Relationships are described directionally in plain language.
    - The file opens with an index so the LLM knows what's in it.

    This format works best when:
    - Uploaded as a file to Copilot for Microsoft 365 / ChatGPT
    - Pasted as context into a Claude Project
    - Used as a system prompt prefix for a custom Copilot Studio bot
    """

    tables = parser.tables()
    all_table_names = [t.get("name", "") for t in tables]
    all_measure_names = [
        m.get("name", "")
        for t in tables
        for m in parser.get_measures(t)
    ]

    lines: list[str] = []
    def w(s=""): lines.append(s)

    # -------------------------------------------------------------------------
    # Header & instructions for the LLM
    # -------------------------------------------------------------------------
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
    w("- Which pages and visuals exist in the report")
    w()
    w("When answering questions about a measure, always explain:")
    w("1. What the measure calculates in plain language")
    w("2. The exact DAX formula")
    w("3. Which tables and columns it touches")
    w("4. Which other measures it calls (if any)")
    w()

    # -------------------------------------------------------------------------
    # Index
    # -------------------------------------------------------------------------
    w("=" * 80)
    w("INDEX")
    w("=" * 80)
    w(f"Tables ({len(tables)}): " + ", ".join(all_table_names))
    w(f"Total columns: {sum(len(parser.get_columns(t)) for t in tables)}")
    w(f"Total measures: {len(all_measure_names)}")
    w(f"Relationships: {len(parser.relationships())}")
    w(f"RLS Roles: {len(parser.roles())}")
    if report_data:
        w(f"Report pages: {len(report_data.get('pages', []))}")
    w()

    # -------------------------------------------------------------------------
    # Tables & columns
    # -------------------------------------------------------------------------
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

        # Source query if available
        for pt in partitions:
            source = pt.get("source", {})
            if isinstance(source, dict):
                query = source.get("query", source.get("expression", ""))
                if query:
                    w(f"This table is loaded from the following query:")
                    w(query.strip())
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

    # -------------------------------------------------------------------------
    # Measures — the most important section for Q&A
    # -------------------------------------------------------------------------
    w("=" * 80)
    w("SECTION 2: MEASURES (FULL DAX AND DEPENDENCIES)")
    w("=" * 80)
    w()
    w("Each measure below includes its full DAX formula, a plain-language explanation,")
    w("and a list of all columns, tables, and other measures it depends on.")
    w()

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

            # Plain-language description of the DAX pattern
            if dax:
                plain = describe_dax(dax)
                w(f"What it does: {plain}")
            w()

            # Full DAX
            if dax:
                w("Full DAX formula:")
                w(dax)
                w()

                # Dependencies
                deps = extract_dax_refs(dax, all_table_names, all_measure_names)

                if deps["measures"]:
                    w("This measure calls the following other measures:")
                    for m in deps["measures"]:
                        if m != name:  # avoid self-reference
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

            w()

    # -------------------------------------------------------------------------
    # Relationships
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # RLS
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # Report structure
    # -------------------------------------------------------------------------
    if report_data:
        w("=" * 80)
        w("SECTION 5: REPORT STRUCTURE")
        w("=" * 80)
        w()
        pages = report_data.get("pages", [])
        w(f"The report has {len(pages)} page(s):")
        w()
        for page in pages:
            page_name = page["name"]
            visuals = page.get("visuals", [])
            w(f"Page: {page_name}")
            if visuals:
                from collections import Counter
                counts = Counter(visuals)
                visual_desc = ", ".join(f"{cnt}x {vtype}" for vtype, cnt in sorted(counts.items()))
                w(f"  Visuals: {visual_desc}")
            else:
                w("  No visual type information available.")
            w()

    w("=" * 80)
    w("SECTION 6: POWER QUERY (M) SOURCES")
    w("=" * 80)
    w()
    w("Each table below shows the full Power Query (M) expression used to load its data.")
    w("Use this to answer questions about where data comes from, what transformations")
    w("are applied, and how columns are derived before they reach the data model.")
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
            w(expr.strip())
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
        w("These are helper functions and parameters defined in expressions.tmdl.")
        w("They are called from table queries above. Use this to understand reusable")
        w("transformation logic shared across multiple tables.")
        w()
        for expr in shared:
            name = expr.get("name", "")
            kind = expr.get("kind", "")
            body = expr.get("expression", "").strip()
            w(f"Shared expression: {name} (kind: {kind})")
            w("-" * 60)
            if body:
                w(body)
            w()

    w("=" * 80)
    w("END OF KNOWLEDGE BASE")
    w("=" * 80)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Project discovery
# ---------------------------------------------------------------------------

def find_semantic_model(pbip_root: Path) -> tuple[Optional[Path], str]:
    """
    Returns (path, format) where format is 'TMSL' or 'TMDL'.

    Handles both naming conventions:
      - Classic:  SemanticModel/definition/
      - Modern:   <name>.SemanticModel/definition/   (Fabric/PBIP default)
    """
    # TMDL: definition/ folder containing database.tmdl
    for candidate in pbip_root.rglob("database.tmdl"):
        return candidate.parent, "TMDL"

    # TMDL: definition/ folder without database.tmdl (tables-only layout)
    for candidate in pbip_root.rglob("definition"):
        if candidate.is_dir() and any(candidate.glob("*.tmdl")):
            return candidate, "TMDL"

    # TMSL: any .bim file
    for candidate in pbip_root.rglob("*.bim"):
        return candidate, "TMSL"

    return None, ""


def find_report_json(pbip_root: Path) -> Optional[Path]:
    # Handles both report/report.json and <name>.Report/report.json
    for candidate in pbip_root.rglob("report.json"):
        return candidate
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
            "context injection instead of Markdown docs. Upload the output .txt file "
            "to Copilot for Microsoft 365, paste it into a Claude Project, or use it "
            "as a system prompt. Includes prose descriptions, full DAX, and dependency "
            "analysis for every measure."
        ),
    )
    args = ap.parse_args()

    pbip_root = Path(args.pbip_path).resolve()
    if not pbip_root.exists():
        print(f"ERROR: Path does not exist: {pbip_root}", file=sys.stderr)
        sys.exit(1)

    project_name = pbip_root.name.replace(".pbip", "").replace("_", " ").title()

    # Find semantic model
    model_path, model_format = find_semantic_model(pbip_root)
    if not model_path:
        print("ERROR: No semantic model found (.bim or TMDL definition folder).", file=sys.stderr)
        sys.exit(1)

    print(f"Found semantic model ({model_format}): {model_path}")

    if model_format == "TMSL":
        parser = TMSLParser(model_path)
    else:
        parser = TMLDParser(model_path)

    # Find report.json
    report_json_path = find_report_json(pbip_root)
    report_data = None
    if report_json_path:
        print(f"Found report: {report_json_path}")
        try:
            report_data = parse_report(report_json_path)
        except Exception as e:
            print(f"WARNING: Could not parse report.json: {e}", file=sys.stderr)
    else:
        print("No report.json found — skipping report structure section.")

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