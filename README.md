# pbip_extract

A command-line tool that generates documentation from a Power BI Project (`.pbip`) file.

Supports both **TMSL** (`.bim`) and **TMDL** (`.tmdl`) semantic model formats.

## Output modes

| Mode | Description |
|------|-------------|
| Default | Markdown documentation — tables, columns, measures, relationships, RLS roles, report pages |
| `--copilot` | Plain-text knowledge base for LLM context injection (Copilot for M365, ChatGPT, Claude) |

## Requirements

- Python 3.8+
- No external dependencies (stdlib only)

## Usage

Point the script at the **`.pbip` file itself** (not the project folder). The path can be
absolute or relative to the directory you run the script from:

```bash
# Absolute path
python pbip_extract.py "C:\Users\you\Documents\MyReport\MyReport.pbip"

# Relative path (resolved from current directory)
python pbip_extract.py MyReport/MyReport.pbip
python pbip_extract.py ./MyReport/MyReport.pbip
```

```bash
# Markdown docs
python pbip_extract.py <path-to-project>.pbip

# Write to a specific file
python pbip_extract.py <path-to-project>.pbip --output docs.md

# LLM knowledge base
python pbip_extract.py <path-to-project>.pbip --copilot

# LLM knowledge base to a specific file
python pbip_extract.py <path-to-project>.pbip --copilot --output knowledge.txt

# Include per-table row counts (see "Row counts" below)
python pbip_extract.py <path-to-project>.pbip --rowcounts rowcounts.txt
```

Passing the project *folder* instead of the `.pbip` file gives a clear error (with a
suggestion, if there's exactly one `.pbip` file in that folder) rather than silently
guessing. This is deliberate: the `.pbip` file's `artifacts` list, and the
`definition.pbir` file inside its report folder, already contain the exact, authoritative
reference to which report and which semantic model belong together — reading those is
more correct than scanning the folder tree for *some* report and *some* model, which can
pick up the wrong one in a folder containing more than one PBIP project. A report bound
to a live connection to a remote/published model (no local semantic model in the project)
is also reported clearly instead of failing with "not found".

## Row counts

The `.pbip` project files only describe the data **model** (tables, columns, M/DAX) —
the actual loaded data lives in the Analysis Services engine behind an open Power BI
Desktop file, not in anything checked into git. So getting a real row count per table
needs a small manual round-trip instead of a live connection (which would require an
extra ADOMD.NET/XMLA dependency and a running model):

1. Run `pbip_extract.py` as usual. Alongside the documentation it writes a
   `<project>_rowcount_query.dax` helper file (unless `--no-rowcount-query` is passed).
2. Open the `.pbip` project in Power BI Desktop (model must be loaded) and go to
   **View > DAX query view** — or use DAX Studio. Paste in the generated query and run it.
3. Select all results (Ctrl+A), copy (Ctrl+C), and paste them into a plain text file
   (keep the header row).
4. Re-run `pbip_extract.py --rowcounts <that file>` — row counts now show up per table
   and in the overview.

## Data lineage & impact analysis

Both output modes end with a lineage section, built entirely from what's already
parsed — no extra input needed. For every table it shows:

- **Upstream** — where its data physically comes from: the connector type (Fabric SQL
  Endpoint, SharePoint, SQL Server, Web, ...), detected from the M-query's function
  calls, plus the source schema/object name when the query uses the classic
  `Source{[Schema="...",Item="..."]}` navigation pattern (this schema/object name is
  not masked — see Sanitisation below — it's normal, non-sensitive source metadata).
- **Downstream** — what depends on it: calculated columns and measures (anywhere in
  the model, including on the table itself) whose DAX references this table,
  relationships it participates in, and which report visuals surface one of its fields.

Use it to answer "where does this table's data come from?" or "what would break if I
changed this table?" without cross-referencing every other section by hand.

## Sanitisation

Power Query and DAX text is scanned for organisation-specific infrastructure values
before it's written out, and those values are replaced with a placeholder — the query
structure itself (functions, applied steps, column references) is always kept intact:

| Found in the query | Replaced with |
|---|---|
| SharePoint site URL | `https://[TENANT].sharepoint.com/sites/[SITE]/...` (path/filename kept) |
| Fabric SQL analytics endpoint hostname | `[FABRIC_SQL_ENDPOINT]` |
| `workspaceId` / `groupId` value | `[WORKSPACE_ID]` |
| `lakehouseId` value | `[LAKEHOUSE_ID]` |
| `warehouseId` value | `[WAREHOUSE_ID]` |
| `datasetId` value | `[DATASET_ID]` |
| `driveId` / `itemId` / `siteId` value | `[SHAREPOINT_DRIVE_ID]` / `[SHAREPOINT_ITEM_ID]` / `[SHAREPOINT_SITE_ID]` |
| standalone `schema = "..."` assignment | `[SCHEMA]` (the common `[Schema="dbo",Item=...]` navigation field used by classic SQL connectors is left alone — it's not sensitive) |

Query **parameters** (`expressions.tmdl` entries with `IsParameterQuery=true`) are excluded
from the documentation entirely rather than masked, since knowing a parameter merely
exists is rarely useful on its own — see "Shared Power Query functions" below.

## What gets extracted

**Semantic model**
- All tables with columns (name, data type, description) — including full DAX formulas for calculated columns
- All DAX measures with full expressions
- Relationships (cardinality, cross-filter direction)
- Row-Level Security (RLS) roles and filter expressions
- Power Query (M) partition expressions per table
- Shared Power Query functions (`expressions.tmdl`) — query **parameters are deliberately excluded**, since they often carry organisation-specific default values (server names, environments, connection info)

**Report**
- Report pages and visual types

## Copilot / LLM mode

The `--copilot` flag generates a structured plain-text file optimised for AI assistants. Each measure includes:

- A plain-English description based on the top-level DAX function
- The full DAX formula
- Referenced columns, tables, and dependent measures

Upload the output to Copilot for Microsoft 365, paste it into a Claude Project, or use it as a system prompt prefix for a Copilot Studio bot.

## Supported project layouts

```
MyReport/
├── MyReport.SemanticModel/
│   └── definition/          # TMDL layout (Fabric / modern PBIP)
│       ├── database.tmdl
│       ├── tables/
│       ├── relationships.tmdl
│       ├── roles/
│       └── expressions.tmdl
└── MyReport.Report/
    └── report.json

MyReport/
└── SemanticModel/
    └── model.bim             # TMSL layout (classic PBIP)
```

## Roadmap
- GUI-ondersteuning voor gebruiksvriendelijke invoer
- Visuele rapportage in de kennisbank
