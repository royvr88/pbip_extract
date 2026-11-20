# pbip_extract

Generates one Markdown document from a Power BI Project (`.pbip`): what the report
looks like and how the semantic model behind it is built. Useful as documentation and
as a knowledge file for an AI assistant (Copilot, Claude, ChatGPT).

## Requirements

- Python 3.8 or newer, standard library only (nothing to `pip install`)
- Works on Windows, macOS and Linux

## Usage

```
python pbip_extract.py "C:\Reports\MyReport\MyReport.pbip"
python pbip_extract.py MyReport.pbip -o docs.md
python pbip_extract.py MyReport.pbip --rowcounts rowcounts.txt
python pbip_gui.py          (same thing with a small window; double-click on Windows)
```

| Option | Meaning |
|---|---|
| `pbip_file` | The `.pbip` file, or a folder that contains exactly one |
| `-o, --output` | Output file or folder. Default: `<project>_docs.md` in the current folder |
| `--rowcounts FILE` | Add row counts per table (see below) |
| `--no-rowcount-query` | Don't write the `<project>_rowcount_query.dax` helper |
| `--copilot` | Same content, but the default file name is `<project>_copilot_kb.txt` |

The report and model folders are found through the references in the `.pbip` file and
the report's `definition.pbir`, so a folder with several projects is never mixed up.
A report with a live connection to a published model has no local model; then only
the report is documented.

## What is extracted

**Report** (PBIR folder format and the older single `report.json`)
- Pages in display order, canvas size, hidden / tooltip / drillthrough pages, background image, theme, bookmarks
- Per visual: type, title and subtitle, position and size, fields per role (with renamed
  display names), sort order, fields used in conditional formatting or labels
- Text of text boxes (including links), button icons and actions, images, shapes, tooltip pages
- Slicers: style (dropdown, list, ...), single select, default selection
- Filters from the filter pane on all three levels: all pages, page, and visual, with the
  condition written out (`Productenmapping[code] = "100211"`, `Kalender[Date] = this year`)

**Semantic model** (TMDL and `.bim`)
- Tables (hidden ones are marked), columns, descriptions, hierarchies, calculated columns and tables
- Measures with DAX, format, display folder, the measures and columns they use, and where they are used in the report
- Relationships, row-level security roles and their filter expressions
- Per table: where the data comes from, the Power Query (M) code, what refers to it and on which pages it is used
- Shared Power Query expressions (query parameters are left out on purpose)

## Sanitisation

Organisation-specific values in Power Query and DAX are replaced by a placeholder; the
query structure itself stays intact.

| Found | Replaced with |
|---|---|
| SharePoint site URL | `https://[TENANT].sharepoint.com/sites/[SITE]/...` (path kept) |
| Fabric SQL endpoint host name | `[FABRIC_SQL_ENDPOINT]` |
| Power BI workspace in a `powerbi://` URL | `[WORKSPACE]` |
| `workspaceId`, `groupId`, `lakehouseId`, `warehouseId`, `datasetId`, `driveId`, `itemId`, `siteId` values | `[WORKSPACE_ID]`, `[LAKEHOUSE_ID]`, ... |
| `schema = "..."` on its own line | `[SCHEMA]` (`[Schema="dbo",Item=...]` navigation is kept) |

Text on the report canvas (text boxes, titles) is copied as is.

## Row counts

The project files contain no data, so row counts need one manual round trip:

1. Run the script. It also writes `<project>_rowcount_query.dax`.
2. Open the project in Power BI Desktop, go to **DAX query view** (or use DAX Studio) and run that query.
3. Copy the result grid including the header into a text file.
4. Run the script again with `--rowcounts <that file>`.

## Windows note

Report folders are deeply nested. If Windows reports "path not found" for a project in a
long folder path, move the project to a shorter path or enable long paths in Windows.
