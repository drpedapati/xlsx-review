---
name: xlsx-review
description: "Use this skill when the user wants to create, edit, read, or diff Excel spreadsheets (.xlsx). Triggers: 'spreadsheet', 'Excel', '.xlsx', 'table in Excel', 'create a spreadsheet'. Do NOT use for CSV files, Google Sheets, or non-Excel tabular data."
license: MIT
---

# Excel (.xlsx) Creation, Editing, and Analysis

Create, edit, read, and diff Excel spreadsheets using `xlsx-review`, a CLI tool built on Microsoft's Open XML SDK. Ships as a self-contained native binary — no runtime required.

## Quick Reference

| Task | Command |
|------|---------|
| Create new workbook | `xlsx-review --create -o output.xlsx manifest.json` |
| Create empty workbook | `xlsx-review --create -o output.xlsx` |
| Create from template | `xlsx-review --create --template base.xlsx -o output.xlsx manifest.json` |
| Edit existing workbook | `xlsx-review input.xlsx edits.json -o output.xlsx` |
| Read workbook contents | `xlsx-review input.xlsx --read --json` |
| Read (human-readable) | `xlsx-review input.xlsx --read` |
| Diff two spreadsheets | `xlsx-review --diff old.xlsx new.xlsx` |
| Diff (JSON output) | `xlsx-review --diff old.xlsx new.xlsx --json` |
| Dry run (validate only) | `xlsx-review input.xlsx edits.json --dry-run` |
| Pipe JSON from stdin | `cat edits.json \| xlsx-review input.xlsx -o output.xlsx` |

## Warnings — Known Corrupt-File Bugs

**READ THESE BEFORE GENERATING ANY MANIFEST. Violating these rules produces files that open with a repair dialog or fail to open entirely.**

1. **NEVER combine `set_table` with `set_auto_filter` on the same range.** Tables include auto-filter automatically. Using both produces a corrupt file. If you need a filterable table, use `set_table` alone.

2. **NEVER use `set_page_orientation` or `set_print_area`.** These produce invalid XML element ordering. The resulting file opens with a repair dialog in Excel. Omit print layout changes entirely.

3. **NEVER use `comments` on cells that already have comments from a previous edit pass.** This produces invalid `legacyDrawing` element ordering. The file opens with a repair dialog.

4. **Always validate after creating.** Run `xlsx-review output.xlsx --read --json` and check that the structure matches expectations. For critical files, open in Excel or LibreOffice to verify.

## Workflow

### Step 1 — Plan the spreadsheet structure

Before writing JSON, decide: sheet names, column headers, data types (text vs. number), which ranges need formulas, tables, validation, or conditional formatting. State the plan, then write the manifest.

### Step 2 — Write the JSON manifest

The manifest is a JSON object with `author` (optional), `changes` (array of operations), and `comments` (optional array). Operations execute in array order.

```json
{
  "author": "Author Name",
  "changes": [
    { "type": "...", ... }
  ],
  "comments": [
    { "sheet": "Sheet1", "cell": "A1", "text": "Note text" }
  ]
}
```

### Step 3 — Run the tool

```bash
# Create new workbook and populate it
xlsx-review --create -o output.xlsx manifest.json --json

# Edit existing workbook
xlsx-review input.xlsx edits.json -o output.xlsx --json
```

### Step 4 — Validate the output

```bash
# Read back the workbook and verify structure
xlsx-review output.xlsx --read --json
```

Check that cell values, formulas, sheet names, and table definitions match expectations. If the output JSON shows unexpected structure, fix the manifest and re-run.

## Change Type Reference

Every operation goes in the `changes` array. All fields are required unless marked optional.

### Cell Operations

**`set_cell`** — Set a cell's value.

```json
{ "type": "set_cell", "sheet": "Sheet1", "cell": "A1", "value": "Hello" }
{ "type": "set_cell", "sheet": "Sheet1", "cell": "B2", "value": "42", "format": "number" }
```

- `sheet` — worksheet name (case-sensitive)
- `cell` — A1 notation
- `value` — string value
- `format` (optional) — `"number"` to store as numeric. Without this, all values are stored as strings.

**`set_formula`** — Set a cell formula.

```json
{ "type": "set_formula", "sheet": "Sheet1", "cell": "C1", "formula": "=SUM(A1:B1)" }
{ "type": "set_formula", "sheet": "Summary", "cell": "A1", "formula": "COUNTA(Sheet1!A:A)-1" }
```

- `formula` — Excel formula. The leading `=` is optional; the tool adds it if missing.

### Row and Column Operations

**`insert_row`** — Insert a blank row after a specified row number.

```json
{ "type": "insert_row", "sheet": "Sheet1", "after": 5 }
```

**`delete_row`** — Delete a row (shifts rows up).

```json
{ "type": "delete_row", "sheet": "Sheet1", "row": 10 }
```

**`insert_column`** — Insert a blank column after a specified column letter.

```json
{ "type": "insert_column", "sheet": "Sheet1", "after": "C" }
```

**`delete_column`** — Delete a column (shifts columns left).

```json
{ "type": "delete_column", "sheet": "Sheet1", "column": "D" }
```

### Sheet Operations

**`add_sheet`** — Add a new worksheet.

```json
{ "type": "add_sheet", "name": "Summary" }
```

**`rename_sheet`** — Rename a worksheet.

```json
{ "type": "rename_sheet", "from": "Sheet1", "to": "Data" }
```

**`delete_sheet`** — Delete a worksheet.

```json
{ "type": "delete_sheet", "name": "Old Sheet" }
```

### Tables

**`set_table`** — Create a formatted Excel table with auto-filter. **Do NOT also use `set_auto_filter` on the same range.**

```json
{
  "type": "set_table",
  "sheet": "Data",
  "range": "A1:D6",
  "name": "PatientScores",
  "display_name": "PatientScores",
  "style_name": "TableStyleMedium2"
}
```

- `range` — A1 range covering headers and data
- `name` — internal table name (no spaces, must be unique in workbook)
- `display_name` (optional) — display name (defaults to `name`)
- `style_name` (optional) — Excel table style, e.g. `TableStyleMedium2`, `TableStyleLight1`, `TableStyleDark1`
- `header_row_count` (optional) — currently only `1` is supported

**`delete_table`** — Delete an Excel table by name.

```json
{ "type": "delete_table", "name": "PatientScores" }
```

### Freeze Panes

**`set_freeze_panes`** — Freeze rows and/or columns. The cell reference is the top-left cell of the scrollable area.

```json
{ "type": "set_freeze_panes", "sheet": "Data", "cell": "A2" }
{ "type": "set_freeze_panes", "sheet": "Data", "cell": "B2" }
```

- `"A2"` freezes row 1 (header row)
- `"B2"` freezes row 1 and column A
- `"C3"` freezes rows 1-2 and columns A-B

**`clear_freeze_panes`** — Remove frozen panes.

```json
{ "type": "clear_freeze_panes", "sheet": "Data" }
```

### Merged Cells

**`merge_cells`** — Merge a cell range.

```json
{ "type": "merge_cells", "sheet": "Data", "range": "B2:C2" }
```

**`unmerge_cells`** — Unmerge a cell range.

```json
{ "type": "unmerge_cells", "sheet": "Data", "range": "B2:C2" }
```

### Data Validation

**`set_data_validation`** — Add dropdown lists, numeric constraints, or other validation rules.

Dropdown list:
```json
{
  "type": "set_data_validation",
  "sheet": "Data",
  "range": "D2:D6",
  "validation_type": "list",
  "formula1": "\"Control,Treatment\"",
  "allow_blank": true
}
```

Numeric range (whole number between 0 and 100):
```json
{
  "type": "set_data_validation",
  "sheet": "Data",
  "range": "C2:C6",
  "validation_type": "whole",
  "validation_operator": "between",
  "formula1": "0",
  "formula2": "100",
  "show_error_message": true
}
```

- `validation_type` — `"list"`, `"whole"`, `"decimal"`, `"textLength"`, `"date"`, `"time"`, `"custom"`
- `validation_operator` (optional) — `"between"`, `"notBetween"`, `"equal"`, `"notEqual"`, `"greaterThan"`, `"lessThan"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`
- `formula1` — for lists: comma-separated values in double quotes; for numeric: the constraint value
- `formula2` (optional) — second value for `between`/`notBetween`
- `allow_blank` (optional) — boolean
- `show_input_message` (optional) — boolean
- `show_error_message` (optional) — boolean

**`clear_data_validation`** — Remove a validation rule by exact range match.

```json
{ "type": "clear_data_validation", "sheet": "Data", "range": "D2:D6" }
```

### Conditional Formatting

**`set_conditional_format`** — Color cells based on rules.

Expression-based (highlight cells matching a formula):
```json
{
  "type": "set_conditional_format",
  "sheet": "Data",
  "range": "C2:C6",
  "conditional_type": "expression",
  "formula1": "C2=\"Treatment\"",
  "fill_color": "yellow",
  "stop_if_true": true
}
```

Cell value comparison:
```json
{
  "type": "set_conditional_format",
  "sheet": "Data",
  "range": "D2:D6",
  "conditional_type": "cellIs",
  "conditional_operator": "greaterThan",
  "formula1": "50",
  "fill_color": "FDE9D9"
}
```

- `conditional_type` — `"expression"` or `"cellIs"`
- `conditional_operator` (optional, for `cellIs`) — `"greaterThan"`, `"lessThan"`, `"equal"`, `"notEqual"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`, `"between"`, `"notBetween"`
- `formula1` — the expression or comparison value
- `formula2` (optional) — second value for `between`/`notBetween`
- `fill_color` (optional) — hex color without `#`, or color name like `"yellow"`. Defaults to yellow.
- `priority` (optional) — integer priority (lower = higher priority)
- `stop_if_true` (optional) — boolean

**`clear_conditional_format`** — Remove conditional formatting by exact range match.

```json
{ "type": "clear_conditional_format", "sheet": "Data", "range": "C2:C6" }
```

### Hyperlinks

**`set_hyperlink`** — Add a clickable link to a cell.

```json
{ "type": "set_hyperlink", "sheet": "Data", "cell": "A2", "url": "https://example.com/p001" }
```

**`clear_hyperlink`** — Remove a hyperlink from a cell.

```json
{ "type": "clear_hyperlink", "sheet": "Data", "cell": "A2" }
```

### Sheet Visibility

**`set_sheet_visibility`** — Control whether a sheet is visible, hidden, or very hidden.

```json
{ "type": "set_sheet_visibility", "name": "Summary", "visibility": "hidden" }
```

- `visibility` — `"visible"`, `"hidden"`, or `"veryHidden"` (only accessible via VBA)

### Named Ranges

**`set_defined_name`** — Create or update a named range.

```json
{
  "type": "set_defined_name",
  "name": "ScoreRange",
  "refers_to": "Data!$C$2:$C$6",
  "comment": "Primary score range"
}
```

Sheet-scoped:
```json
{
  "type": "set_defined_name",
  "name": "SummaryMetric",
  "scope_sheet": "Summary",
  "refers_to": "Summary!$B$3",
  "hidden": true,
  "comment": "Summary-scoped metric cell"
}
```

**`delete_defined_name`** — Delete a named range.

```json
{ "type": "delete_defined_name", "name": "ScoreRange" }
```

### Workbook and Sheet Protection

**`set_workbook_protection`** — Lock workbook structure.

```json
{ "type": "set_workbook_protection", "lock_structure": true }
```

**`set_sheet_protection`** — Lock a worksheet.

```json
{ "type": "set_sheet_protection", "sheet": "Data", "enabled": true }
```

### Auto-Filter

**`set_auto_filter`** — Apply auto-filter to a range. **Do NOT use on ranges that already have a `set_table` — tables include auto-filter automatically.**

```json
{ "type": "set_auto_filter", "sheet": "Data", "range": "A1:D6" }
```

**`clear_auto_filter`** — Remove auto-filter.

```json
{ "type": "clear_auto_filter", "sheet": "Data" }
```

## Complete Examples

### Example 1: Equipment Tracking Table

Creates a workbook with headers, data rows, an Excel table, and frozen header row.

```json
{
  "author": "Lab Manager",
  "changes": [
    { "type": "set_cell", "sheet": "Sheet1", "cell": "A1", "value": "Equipment" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "B1", "value": "Location" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "C1", "value": "Status" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "D1", "value": "Last Calibrated" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "A2", "value": "EEG Amplifier" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "B2", "value": "Room 201" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "C2", "value": "Active" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "D2", "value": "2026-03-15" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "A3", "value": "TMS Coil" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "B3", "value": "Room 105" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "C3", "value": "Maintenance" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "D3", "value": "2026-01-20" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "A4", "value": "fMRI Scanner" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "B4", "value": "Building C" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "C4", "value": "Active" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "D4", "value": "2026-04-01" },
    { "type": "rename_sheet", "from": "Sheet1", "to": "Equipment" },
    {
      "type": "set_table",
      "sheet": "Equipment",
      "range": "A1:D4",
      "name": "EquipmentTracker",
      "style_name": "TableStyleMedium2"
    },
    { "type": "set_freeze_panes", "sheet": "Equipment", "cell": "A2" }
  ]
}
```

```bash
xlsx-review --create -o equipment.xlsx manifest.json --json
```

### Example 2: Multi-Sheet Workbook with Cross-Sheet Formulas

Creates a data sheet and a summary sheet with formulas referencing the data.

```json
{
  "author": "Analyst",
  "changes": [
    { "type": "rename_sheet", "from": "Sheet1", "to": "Scores" },
    { "type": "set_cell", "sheet": "Scores", "cell": "A1", "value": "Subject" },
    { "type": "set_cell", "sheet": "Scores", "cell": "B1", "value": "Pre-Test" },
    { "type": "set_cell", "sheet": "Scores", "cell": "C1", "value": "Post-Test" },
    { "type": "set_cell", "sheet": "Scores", "cell": "A2", "value": "S001" },
    { "type": "set_cell", "sheet": "Scores", "cell": "B2", "value": "72", "format": "number" },
    { "type": "set_cell", "sheet": "Scores", "cell": "C2", "value": "85", "format": "number" },
    { "type": "set_cell", "sheet": "Scores", "cell": "A3", "value": "S002" },
    { "type": "set_cell", "sheet": "Scores", "cell": "B3", "value": "68", "format": "number" },
    { "type": "set_cell", "sheet": "Scores", "cell": "C3", "value": "91", "format": "number" },
    { "type": "set_cell", "sheet": "Scores", "cell": "A4", "value": "S003" },
    { "type": "set_cell", "sheet": "Scores", "cell": "B4", "value": "80", "format": "number" },
    { "type": "set_cell", "sheet": "Scores", "cell": "C4", "value": "88", "format": "number" },
    { "type": "add_sheet", "name": "Summary" },
    { "type": "merge_cells", "sheet": "Summary", "range": "A1:C1" },
    { "type": "set_cell", "sheet": "Summary", "cell": "A1", "value": "Score Summary" },
    { "type": "set_cell", "sheet": "Summary", "cell": "A3", "value": "Metric" },
    { "type": "set_cell", "sheet": "Summary", "cell": "B3", "value": "Pre-Test" },
    { "type": "set_cell", "sheet": "Summary", "cell": "C3", "value": "Post-Test" },
    { "type": "set_cell", "sheet": "Summary", "cell": "A4", "value": "Mean" },
    { "type": "set_formula", "sheet": "Summary", "cell": "B4", "formula": "=AVERAGE(Scores!B2:B4)" },
    { "type": "set_formula", "sheet": "Summary", "cell": "C4", "formula": "=AVERAGE(Scores!C2:C4)" },
    { "type": "set_cell", "sheet": "Summary", "cell": "A5", "value": "Count" },
    { "type": "set_formula", "sheet": "Summary", "cell": "B5", "formula": "=COUNTA(Scores!A2:A4)" },
    { "type": "set_cell", "sheet": "Summary", "cell": "A6", "value": "Max Improvement" },
    { "type": "set_formula", "sheet": "Summary", "cell": "B6", "formula": "=MAX(Scores!C2:C4-Scores!B2:B4)" }
  ]
}
```

### Example 3: Data Validation with Dropdowns and Numeric Constraints

Adds dropdown lists and numeric range validation to an existing data sheet.

```json
{
  "author": "Data Manager",
  "changes": [
    { "type": "set_cell", "sheet": "Sheet1", "cell": "A1", "value": "Subject" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "B1", "value": "Group" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "C1", "value": "Age" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "D1", "value": "Score" },
    {
      "type": "set_data_validation",
      "sheet": "Sheet1",
      "range": "B2:B100",
      "validation_type": "list",
      "formula1": "\"Control,Treatment A,Treatment B,Placebo\"",
      "allow_blank": true
    },
    {
      "type": "set_data_validation",
      "sheet": "Sheet1",
      "range": "C2:C100",
      "validation_type": "whole",
      "validation_operator": "between",
      "formula1": "0",
      "formula2": "120",
      "show_error_message": true
    },
    {
      "type": "set_data_validation",
      "sheet": "Sheet1",
      "range": "D2:D100",
      "validation_type": "decimal",
      "validation_operator": "greaterThanOrEqual",
      "formula1": "0",
      "show_error_message": true
    },
    { "type": "set_freeze_panes", "sheet": "Sheet1", "cell": "A2" }
  ]
}
```

### Example 4: Conditional Formatting for Color-Coded Status

Applies conditional formatting to highlight scores by threshold.

```json
{
  "author": "Reviewer",
  "changes": [
    {
      "type": "set_conditional_format",
      "sheet": "Results",
      "range": "D2:D50",
      "conditional_type": "cellIs",
      "conditional_operator": "greaterThanOrEqual",
      "formula1": "90",
      "fill_color": "C6EFCE"
    },
    {
      "type": "set_conditional_format",
      "sheet": "Results",
      "range": "D2:D50",
      "conditional_type": "cellIs",
      "conditional_operator": "between",
      "formula1": "70",
      "formula2": "89",
      "fill_color": "FFEB9C"
    },
    {
      "type": "set_conditional_format",
      "sheet": "Results",
      "range": "D2:D50",
      "conditional_type": "cellIs",
      "conditional_operator": "lessThan",
      "formula1": "70",
      "fill_color": "FFC7CE"
    },
    {
      "type": "set_conditional_format",
      "sheet": "Results",
      "range": "E2:E50",
      "conditional_type": "expression",
      "formula1": "E2=\"Flagged\"",
      "fill_color": "FFC7CE",
      "stop_if_true": true
    }
  ]
}
```

## Comments

Comments are added as legacy Notes (not threaded comments) for maximum compatibility. They go in the `comments` array, separate from `changes`.

```json
{
  "comments": [
    { "sheet": "Data", "cell": "A2", "text": "This value was updated during review." },
    { "sheet": "Summary", "cell": "B3", "text": "Formula references the raw data range." }
  ]
}
```

Each comment requires `sheet` (case-sensitive), `cell` (A1 notation), and `text`. The `--author` CLI flag or manifest `author` field sets the comment author name.

## Reading Workbook Contents

Use `--read` to extract workbook structure and cell data. The JSON output includes sheet metadata, cell values, formulas, formula types, tables, validations, conditional formats, protection state, and defined names.

```bash
xlsx-review workbook.xlsx --read --json
```

Key fields in the output:
- `workbook.sheet_count`, `workbook.protected`, `workbook.defined_names`
- `sheets[].name`, `sheets[].visibility`, `sheets[].row_count`, `sheets[].cell_count`
- `sheets[].tables`, `sheets[].data_validations`, `sheets[].conditional_formats`
- `sheets[].rows[].cells[].value`, `.formula`, `.type` (`"string"`, `"number"`)

## Diffing Spreadsheets

Compare two workbooks semantically. Detects cell value/formula changes, added/removed sheets, row/column count changes, and metadata changes (visibility, defined names, protection, tables, validations, conditional formats, comments).

```bash
xlsx-review --diff old.xlsx new.xlsx
xlsx-review --diff old.xlsx new.xlsx --json
```

## CLI Flags

| Flag | Description |
|------|-------------|
| `-o`, `--output <path>` | Output file path (default: `<input>_edited.xlsx`) |
| `--author <name>` | Author name for comments (overrides manifest `author`) |
| `--json` | Structured JSON output for scripting/pipelines |
| `--dry-run` | Validate manifest without writing output |
| `--create` | Create a new workbook (blank `Sheet1` by default) |
| `--template <path>` | Use custom workbook as starting template for `--create` |
| `--read` | Read workbook contents (no manifest needed) |
| `--diff` | Compare two spreadsheets semantically |
| `--textconv` | Git textconv driver output |
| `-v`, `--version` | Show version |

## Exit Codes

- `0` — all changes and comments applied successfully
- `1` — one or more edits failed, or invalid input

## Critical Rules

- **Modified cells are auto-highlighted yellow (#FFFF00).** This is built-in behavior; you do not need to add formatting for changed cells.
- **`format: "number"` is required for numeric values.** Without it, numbers are stored as strings and formulas like `SUM` will not work on them.
- **Sheet names are case-sensitive.** `"Sheet1"` and `"sheet1"` are different.
- **Operations execute in array order.** If you rename a sheet, subsequent operations must use the new name. If you add a sheet, you can immediately set cells on it.
- **Table names must be unique across the entire workbook** and cannot contain spaces.
- **Create mode starts with a single `Sheet1`.** Use `rename_sheet` to rename it and `add_sheet` to add more.
- **Formulas can omit the leading `=`.** The tool adds it automatically if missing.
