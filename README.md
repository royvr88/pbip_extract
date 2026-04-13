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

```bash
# Markdown docs
python pbip_extract.py <path-to-pbip-folder>

# Write to a specific file
python pbip_extract.py <path-to-pbip-folder> --output docs.md

# LLM knowledge base
python pbip_extract.py <path-to-pbip-folder> --copilot

# LLM knowledge base to a specific file
python pbip_extract.py <path-to-pbip-folder> --copilot --output knowledge.txt
```

## What gets extracted

**Semantic model**
- All tables with columns (name, data type, description)
- All DAX measures with full expressions
- Relationships (cardinality, cross-filter direction)
- Row-Level Security (RLS) roles and filter expressions
- Power Query (M) partition expressions per table
- Shared Power Query functions and parameters (`expressions.tmdl`)

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
