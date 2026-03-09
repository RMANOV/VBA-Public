// ═══════════════════════════════════════════════════════════════
// PATTERN: Query Folding Best Practices
// CATEGORY: Performance | DIFFICULTY: Intermediate-Advanced
// ═══════════════════════════════════════════════════════════════
// WHAT: Push operations to the data source engine instead of
//       processing them locally in the M (mashup) engine.
// WHY:  A folded query runs on the database server (SQL, OData)
//       which is orders of magnitude faster than downloading
//       all rows and filtering in Power Query.
// KEY LESSON: Do all foldable operations FIRST (filter, select,
//             group), then add custom M logic. Verify with
//             "View Native Query" in the Power Query Editor.
// ═══════════════════════════════════════════════════════════════
//
// ── WHAT IS QUERY FOLDING? ──────────────────────────────────
//
// Power Query translates your M steps into the source's native
// language (SQL for databases, OData for web services).
// This is called "folding" — the M code folds INTO the source query.
//
// Folded:   Table.SelectRows → becomes WHERE in SQL
// Unfolded: each [Amount] * 1.2 → runs row-by-row in M engine
//
// ── HOW TO VERIFY ───────────────────────────────────────────
//
// In Power Query Editor:
//   1. Right-click any step in "Applied Steps"
//   2. If "View Native Query" is available → that step folds
//   3. If greyed out → folding broke at or before that step
//
// ── FOLDABLE OPERATIONS (these translate to SQL) ────────────
//
//   M Function              │ SQL Equivalent
//   ────────────────────────┼──────────────
//   Table.SelectRows        │ WHERE
//   Table.SelectColumns     │ SELECT
//   Table.Group             │ GROUP BY
//   Table.Sort              │ ORDER BY
//   Table.FirstN            │ TOP / LIMIT
//   Table.Skip              │ OFFSET
//   Table.Join              │ JOIN
//   Table.Distinct          │ DISTINCT
//   Table.TransformColumnTypes (basic) │ CAST
//   Table.RenameColumns     │ AS (alias)
//   Table.ExpandTableColumn │ JOIN (nested)
//
// ── FOLD-BREAKING OPERATIONS (force local processing) ───────
//
//   - Custom M functions (each [A] * [B] + [C])
//   - List.Generate, List.Accumulate
//   - Table.Buffer (forces full materialization)
//   - Table.AddColumn with M expressions
//   - Text.Combine, Text.Split (complex string ops)
//   - Table.Pivot / Table.Unpivot (some sources)
//   - Any step referencing another query
//   - Accessing external M variables in row context
//
// ═══════════════════════════════════════════════════════════════
// GOLDEN RULE: Filter → Select → Group → THEN custom columns
// ═══════════════════════════════════════════════════════════════

let
    // ── Simulated source ────────────────────────────────────────
    // In production, replace with Sql.Database, OData.Feed, etc.
    // Query folding only works against foldable sources.
    Source = Table.FromRecords({
        [Region = "East",  Product = "Alpha", Amount = 1200, Date = #date(2025, 1, 15)],
        [Region = "West",  Product = "Beta",  Amount =  800, Date = #date(2025, 2, 10)],
        [Region = "East",  Product = "Alpha", Amount = 1500, Date = #date(2025, 3, 20)],
        [Region = "West",  Product = "Alpha", Amount =  950, Date = #date(2025, 1, 25)],
        [Region = "East",  Product = "Beta",  Amount = 2200, Date = #date(2025, 4,  5)],
        [Region = "North", Product = "Gamma", Amount =  300, Date = #date(2025, 5, 12)]
    }),

    // ═══════════════════════════════════════════════════════════
    // BAD EXAMPLE — custom column FIRST, then filter
    // ═══════════════════════════════════════════════════════════
    // This breaks folding immediately at Step 1.
    // Everything after runs in the slow M engine.

    // BAD_Step1: Custom column → breaks fold
    // BAD_AddMargin = Table.AddColumn(Source, "Margin",
    //     each [Amount] * 0.15, type number),
    //
    // BAD_Step2: Filter → would have folded, but fold is already broken
    // BAD_Filtered = Table.SelectRows(BAD_AddMargin,
    //     each [Region] = "East"),
    //
    // BAD_Step3: Select columns → also runs locally
    // BAD_Result = Table.SelectColumns(BAD_Filtered,
    //     {"Product", "Amount", "Margin"})
    //
    // Result: ALL rows downloaded, margin calculated locally,
    //         THEN filtered. Slow on large datasets.

    // ═══════════════════════════════════════════════════════════
    // GOOD EXAMPLE — foldable operations FIRST
    // ═══════════════════════════════════════════════════════════
    // Steps 1-3 fold into SQL. Only Step 4 runs locally,
    // but by then the dataset is already small.

    // GOOD Step 1: Filter rows → folds to WHERE
    // Against a SQL source, this becomes:
    //   WHERE Region = 'East'
    Step1_Filter = Table.SelectRows(Source,
        each [Region] = "East"),

    // GOOD Step 2: Select columns → folds to SELECT
    // Reduces data transferred over the wire.
    //   SELECT Product, Amount, Date FROM ...
    Step2_Select = Table.SelectColumns(Step1_Filter,
        {"Product", "Amount", "Date"}),

    // GOOD Step 3: Group → folds to GROUP BY
    //   SELECT Product, SUM(Amount) AS TotalAmount
    //   FROM ... GROUP BY Product
    Step3_Group = Table.Group(Step2_Select,
        {"Product"},
        {{"TotalAmount", each List.Sum([Amount]), type number}}
    ),

    // GOOD Step 4: Custom column → NOW it's safe to break fold
    // The dataset is already filtered, selected, and grouped.
    // Only a few rows remain for local processing.
    Step4_Custom = Table.AddColumn(Step3_Group,
        "Margin",
        each [TotalAmount] * 0.15,
        type number
    ),

    // ═══════════════════════════════════════════════════════════
    // ADDITIONAL TIPS
    // ═══════════════════════════════════════════════════════════
    //
    // TIP 1: Table.Buffer — use ONLY when you need to snapshot
    //        a table to avoid re-evaluation. It downloads ALL
    //        rows and kills folding for everything after it.
    //
    // TIP 2: Value.NativeQuery — when M can't express what you
    //        need, write raw SQL and let the server handle it:
    //
    //   Value.NativeQuery(
    //       Sql.Database("server", "db"),
    //       "SELECT * FROM Sales WHERE Amount > @min",
    //       [min = 1000]
    //   )
    //
    // TIP 3: Multiple filters in sequence fold into one WHERE
    //        with AND — you don't need to combine them manually.
    //
    // TIP 4: Type changes (Int64.Type, Currency.Type) generally
    //        fold. Complex type casts (parsing text to dates)
    //        may not — test with View Native Query.
    //
    // TIP 5: Table.NestedJoin folds when both tables are from
    //        the same source. Cross-source joins never fold.
    //
    // ═══════════════════════════════════════════════════════════
    // STEP ORDER CHECKLIST (print and pin to your monitor)
    // ═══════════════════════════════════════════════════════════
    //
    //   1. Connect to source
    //   2. Filter rows       (WHERE)     ← folds
    //   3. Select columns    (SELECT)    ← folds
    //   4. Join tables       (JOIN)      ← folds (same source)
    //   5. Group / aggregate (GROUP BY)  ← folds
    //   6. Sort              (ORDER BY)  ← folds
    //   7. --- FOLD BOUNDARY ---
    //   8. Add custom columns            ← local M engine
    //   9. Pivot / Unpivot               ← local M engine
    //  10. Merge with other queries      ← local M engine
    //
    // ═══════════════════════════════════════════════════════════

    Result = Step4_Custom

in
    Result