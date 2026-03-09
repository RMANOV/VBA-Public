# VBA-Public

A collection of **VBA macros** and **Power Query (M language) patterns** for Excel automation, data processing, and financial analysis.

---

## Power Query (M Language) Patterns

19 production-grade, copy-paste ready patterns in the [`power-query/`](power-query/) directory.

### Beginner

| # | Pattern | What You'll Learn |
|---|---------|-------------------|
| 02 | [Unpivot Wide to Long](power-query/02_UNPIVOT_WIDE_TO_LONG.pq) | `Table.UnpivotOtherColumns` — normalize wide tables |
| 03 | [Error Handling](power-query/03_ERROR_HANDLING_PATTERNS.pq) | `try...otherwise`, `[HasError]` inspection |
| 05 | [Null Handling](power-query/05_NULL_HANDLING.pq) | Three-valued logic, `Coalesce` function |
| 11 | [Header Cleanup Pipeline](power-query/11_HEADER_CLEANUP_PIPELINE.pq) | Promote → Clean → Type (universal ETL entry) |
| 17 | [Conditional Columns](power-query/17_CONDITIONAL_COLUMNS.pq) | Nested `if` vs table-driven lookup |
| 19 | [Text Split/Combine/Extract](power-query/19_TEXT_SPLIT_COMBINE_EXTRACT.pq) | String parsing without regex |

### Intermediate

| # | Pattern | What You'll Learn |
|---|---------|-------------------|
| 01 | [Dynamic Total Row](power-query/01_DYNAMIC_TOTAL_ROW.pq) | Auto-detect numeric columns, inject summary row |
| 04 | [Calendar/Date Table](power-query/04_CALENDAR_DATE_TABLE.pq) | Date dimension generation for Power BI |
| 06 | [Custom M Functions](power-query/06_CUSTOM_M_FUNCTIONS.pq) | Parameterized reusable functions |
| 09 | [Merge Strategies (Joins)](power-query/09_MERGE_STRATEGIES_JOINS.pq) | All 6 join types with examples |
| 10 | [Dynamic Column Operations](power-query/10_DYNAMIC_COLUMN_OPS.pq) | Schema-agnostic transforms |
| 12 | [Multi-Sheet Consolidation](power-query/12_MULTI_SHEET_CONSOLIDATION.pq) | Stack sheets dynamically with `Table.Combine` |
| 13 | [Fiscal Year Grouping](power-query/13_FISCAL_YEAR_GROUPING.pq) | Non-calendar period logic |
| 14 | [Missing Column Resilience](power-query/14_MISSING_COLUMN_RESILIENCE.pq) | Defensive schema handling |
| 15 | [Query Folding Practices](power-query/15_QUERY_FOLDING_PRACTICES.pq) | Performance — push ops to data source |
| 18 | [Custom Sorting](power-query/18_CUSTOM_SORTING.pq) | Non-alphabetical sort with sort key |

### Advanced

| # | Pattern | What You'll Learn |
|---|---------|-------------------|
| 07 | [List.Generate (Iteration)](power-query/07_LIST_GENERATE_ITERATION.pq) | M's general-purpose loop |
| 08 | [List.Accumulate (Reduce)](power-query/08_LIST_ACCUMULATE_REDUCE.pq) | Fold/reduce, running totals |
| 16 | [API Pagination](power-query/16_API_PAGINATION.pq) | Paginated REST API consumption |

### Existing (Root)

| Pattern | Description |
|---------|-------------|
| [VLOOKUP in Power Query](VLOOKUP%20IN%20POWER%20QUERY.pq) | Lookup function with default values |
| [Left/Right Anti Joins](L_R_ANTI_JOINS.pq) | Find unmatched records between tables |

---

## VBA Macros

50+ VBA macros for Excel automation, organized by function:

### Data Cleaning
- `Sub DELETEDUPLICATES()` — Remove duplicate rows
- `Sub DeleteBlankRows1()` / `Sub DeleteBlankCOLUMNs1()` — Remove empty rows/columns
- `Sub DUPLICATES()` — Highlight duplicates

### Navigation & Selection
- `Sub LASTROW()` — Find last row with data
- `Sub FIRSTCOLUMN()` / `Sub LASTCOLUMNLEFT()` / `Sub LASTCOLUMNRIGHT()` — Column navigation
- `Sub FIRSTGREEN()` — Jump to first green cell
- `Sub TOPSCREEN()` — Move to top of visible area
- `Sub SELECTBE()` / `Sub SELECTDG()` — Range selection utilities

### Formatting
- `Sub COLORGAIDE()` — Color formatting utilities
- `Sub SELECTINGCOLORIZE()` — Colorize selected range
- `Sub UNCOLORING()` — Remove all colors
- `Sub MARCGREEN()` — March-specific highlighting

### Paste Operations
- `Sub PASTEFORMULA()` — Smart paste formulas
- `Sub PASTEFORMULAANDDESTRUCT()` — Paste and clean up
- `Sub PASTEVALUESDIVIDE()` / `Sub PASTEVALUESMULTIPLY()` — Paste with arithmetic

### Sorting
- `Sub SORTBE()` / `Sub SORTDS()` / `Sub SORTKA()` — Multi-column sorts

### Utilities
- `WorkDays` — Business days calculation
- `GetOSAndOfficeVersion` — Detect OS and Office version
- `exctract_mob_phone_numbers` — Phone number extraction
- `Sub PlaySound()` — Audio playback
- `Sub MASSAVES()` — Bulk save operations
- `CopyPasteValues` — Copy/paste values utility

### Financial / Fraud Detection
- `GraphsAgainstFrauds` (10 variants) — Graph-based anomaly detection
- `Sub BEGININGLEVELSNEWYEAR()` — Year initialization
- `DecemberCorrector` (3 variants) — Month-end corrections

---

## How to Use

### Power Query Files (.pq)
1. Open Excel → Data → Get Data → Launch Power Query Editor
2. Home → Advanced Editor
3. Paste the contents of any `.pq` file
4. Click Done — the query runs with built-in sample data

### VBA Macros (.cls / .vb)
1. Open Excel → Alt+F11 (VBA Editor)
2. File → Import File → select the `.cls` or `.vb` file
3. Run via Alt+F8 or assign to a keyboard shortcut

---

## License

Free to use, modify, and distribute.
