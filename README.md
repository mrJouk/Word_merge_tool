# Word Merge Tool

This app pairs Word files from `kyrgyz_documents/` and `russian_documents/` by the same filename stem, then creates one merged `.docx` per pair in `merged_documents/`.

The merged document keeps normal non-table content in two-column Kyrgyz/Russian blocks:

- Column 1: Kyrgyz document body content.
- Column 2: Russian document body content.
- Header: `Kyrgyz header text / Russian header text`.
- Footer: `Kyrgyz footer text / Russian footer text`.
- Matching source tables are merged into one table. Matching cells are written as `Kyrgyz cell / Russian cell`, including column titles.

## Install

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

If `py` reports that no Python is installed, install Python 3.11+ first and rerun the commands.

## Create Dummy Data

```powershell
.\.venv\Scripts\python.exe app.py --create-dummy-data
```

This creates:

- `kyrgyz_documents/test.docx`
- `russian_documents/test.docx`
- `kyrgyz_documents/legacy.doc` when Microsoft Word is available
- `russian_documents/legacy.doc` when Microsoft Word is available

## Merge Documents

```powershell
.\.venv\Scripts\python.exe app.py --overwrite
```

The output will be:

- `merged_documents/test_merged.docx`
- `merged_documents/legacy_merged.docx` if the legacy `.doc` dummy files exist

When a source file is `.doc`, the Word backend opens and parses it directly through Microsoft Word COM. It does not export source `.doc` files to `.docx` before merging; only the final merged document is saved as `.docx`.

When both documents contain a table at the same table position, the app creates one merged table instead of two separate tables. Matching header cells are written as `Kyrgyz title / Russian title`.

## Inspect Parsing

Use this when you need to see exactly what Word exposes before merging:

```powershell
.\.venv\Scripts\python.exe inspect_parsing.py --pair legacy --output parsing_report.txt
```

The report shows headers, footers, body paragraphs, Word list labels such as `I)` or `1)`, and table cells for both documents. For `.doc` inputs it also creates a temporary `.docx` export and reports that parse too, so you can check whether `.doc -> .docx` conversion changes Roman numerals. Use `--no-doc-conversion-check` to inspect only the original files without any export.

## Backends

The default `--backend auto` tries Microsoft Word automation first, then falls back to a `.docx` XML backend.

Use Microsoft Word automation for real sensitive documents, especially `.doc` files:

```powershell
.\.venv\Scripts\python.exe app.py --backend word --overwrite
```

The fallback backend supports `.docx` only and preserves direct document XML where possible. Microsoft Word automation is the higher-fidelity path because it copies formatted Word ranges directly.

For documents with automatic numbering such as Roman numerals, use `--backend word`. The Word backend copies non-table content with Word COM formatted ranges and only manually rebuilds matching source tables.

When copied content contains Word automatic list labels, the Word backend freezes those labels in a temporary in-memory copy before inserting them into the merged file. This keeps visible labels such as `I)` and `II)` from being recalculated as decimal numbers in the output.

## Configuration

Operational defaults, supported extensions, Word COM numeric constants, output naming, and dummy sample text live in `config.py`.
