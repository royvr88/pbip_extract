# pbip_extract

A command-line tool that generates documentation from a Power BI Project (`.pbip`) folder.

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

The path can be absolute or relative to the directory you run the script from:

```bash
# Absolute path
python pbip_extract.py "C:\Users\you\Documents\MyReport"

# Relative path (resolved from current directory)
python pbip_extract.py MyReport
python pbip_extract.py ./MyReport
```

```bash
# Markdown docs
python pbip_extract.py <path-to-pbip-folder>

# Write to a specific file
python pbip_extract.py <path-to-pbip-folder> --output docs.md

# LLM knowledge base
python pbip_extract.py <path-to-pbip-folder> --copilot

# LLM knowledge base to a specific file
python pbip_extract.py <path-to-pbip-folder> --copilot --output knowledge.txt

# Include per-table row counts (see "Row counts" below)
python pbip_extract.py <path-to-pbip-folder> --rowcounts rowcounts.txt
```

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
