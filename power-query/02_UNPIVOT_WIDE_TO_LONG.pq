// ═══════════════════════════════════════════════════════════════
// PATTERN: Unpivot Wide to Long
// CATEGORY: Data Reshaping | DIFFICULTY: Beginner-Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Normalize wide (crosstab) tables into long format using UnpivotOtherColumns
// WHY:  Wide tables break when new columns appear; long format is analysis-ready
// KEY LESSON: Always prefer UnpivotOtherColumns — specify what stays fixed
// ═══════════════════════════════════════════════════════════════

let
    // ── STEP 1: Sample wide table ────────────────────────────
    // Typical export: one column per month. Fragile for analysis.
    // If May appears later, hardcoded references break.
    WideTable = Table.FromRecords({
        [Product = "Widget A", Jan = 120, Feb = 135, Mar = 98,  Apr = 145],
        [Product = "Widget B", Jan = 200, Feb = 180, Mar = 210, Apr = 195],
        [Product = "Widget C", Jan = 85,  Feb = 90,  Mar = 110, Apr = 130]
    }),

    Typed = Table.TransformColumnTypes(WideTable, {
        {"Product", type text},
        {"Jan", type number},
        {"Feb", type number},
        {"Mar", type number},
        {"Apr", type number}
    }),

    // ── STEP 2 (RECOMMENDED): UnpivotOtherColumns ───────────
    // Tell M which columns to KEEP FIXED. Everything else gets unpivoted.
    // If the source adds May, Jun, Jul later — they are automatically included.
    // This is the resilient approach.
    Unpivoted = Table.UnpivotOtherColumns(
        Typed,
        {"Product"},           // Anchor columns — these stay as-is
        "Month",               // New column for the former header names
        "Sales"                // New column for the values
    ),

    // ── STEP 2-ALT (BRITTLE): Table.Unpivot ─────────────────
    // This requires listing every column to unpivot explicitly.
    // If May appears in the source, it will be SILENTLY DROPPED.
    // Shown here for contrast — avoid in production.
    //
    // Unpivoted_Brittle = Table.Unpivot(
    //     Typed,
    //     {"Jan", "Feb", "Mar", "Apr"},   // ← Must update manually!
    //     "Month",
    //     "Sales"
    // ),

    // ── STEP 3: Optional — sort for readability ──────────────
    // Define month order so "Jan" < "Feb" < "Mar" < "Apr"
    MonthOrder = {"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"},

    // Add a sort index based on position in MonthOrder list
    WithSortIndex = Table.AddColumn(
        Unpivoted,
        "MonthIndex",
        each List.PositionOf(MonthOrder, [Month]),
        type number
    ),
    Sorted = Table.Sort(WithSortIndex, {
        {"Product", Order.Ascending},
        {"MonthIndex", Order.Ascending}
    }),

    // Remove the helper column — it served its purpose
    Result = Table.RemoveColumns(Sorted, {"MonthIndex"})

    // ── RESULT ───────────────────────────────────────────────
    // Product   | Month | Sales
    // Widget A  | Jan   | 120
    // Widget A  | Feb   | 135
    // Widget A  | Mar   | 98
    // Widget A  | Apr   | 145
    // Widget B  | Jan   | 200
    // ...
    //
    // WHY THIS MATTERS:
    // - PivotTables, charts, and DAX measures all expect long format
    // - Filters on "Month" column work naturally
    // - Adding new months requires ZERO query changes
in
    Result
