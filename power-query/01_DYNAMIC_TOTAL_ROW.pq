// ═══════════════════════════════════════════════════════════════
// PATTERN: Dynamic Total Row
// CATEGORY: Table Manipulation | DIFFICULTY: Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Auto-detect numeric columns and append a TOTAL summary row
// WHY:  When you need a grand total row but column names change over time
// KEY LESSON: Never hardcode column names — discover them dynamically
// ═══════════════════════════════════════════════════════════════

let
    // ── STEP 1: Sample data ──────────────────────────────────
    // In real use, replace this with your actual source query
    Source = Table.FromRecords({
        [Region = "North", Sales = 1200, Cost = 800,  Units = 50],
        [Region = "South", Sales = 950,  Cost = 620,  Units = 38],
        [Region = "East",  Sales = 1800, Cost = 1100, Units = 72],
        [Region = "West",  Sales = 1450, Cost = 970,  Units = 61]
    }),

    // Ensure proper types so numeric detection works reliably
    Typed = Table.TransformColumnTypes(Source, {
        {"Region", type text},
        {"Sales", type number},
        {"Cost", type number},
        {"Units", type number}
    }),

    // ── STEP 2: Discover all column names ────────────────────
    AllColumns = Table.ColumnNames(Typed),

    // ── STEP 3: Identify numeric columns dynamically ─────────
    // Grab the first row and test each value with Value.Is(_, type number).
    // This approach works regardless of how many numeric columns exist.
    FirstRow = Typed{0},
    NumericColumns = List.Select(
        AllColumns,
        each Value.Is(Record.Field(FirstRow, _), type number)
    ),

    // Non-numeric columns get a label; numeric columns get summed
    NonNumericColumns = List.RemoveItems(AllColumns, NumericColumns),

    // ── STEP 4: Build the total record dynamically ───────────
    // For each non-numeric column: put "TOTAL" in the first one, null in rest.
    // For each numeric column: sum the entire column.
    LabelFields = List.Transform(
        NonNumericColumns,
        each {_, if _ = List.First(NonNumericColumns) then "TOTAL" else null}
    ),
    SumFields = List.Transform(
        NumericColumns,
        each {_, List.Sum(Table.Column(Typed, _))}
    ),

    // Record.FromList needs parallel lists of names and values
    TotalFieldNames = List.Transform(LabelFields, each _{0})
                    & List.Transform(SumFields, each _{0}),
    TotalFieldValues = List.Transform(LabelFields, each _{1})
                     & List.Transform(SumFields, each _{1}),

    TotalRecord = Record.FromList(TotalFieldValues, TotalFieldNames),

    // ── STEP 5: Append the total row ─────────────────────────
    // Table.FromRecords wraps our single record into a 1-row table,
    // then Table.Combine stacks it under the original data.
    TotalTable = Table.FromRecords({TotalRecord}),
    Result = Table.Combine({Typed, TotalTable})

    // ── RESULT ───────────────────────────────────────────────
    // Region  | Sales | Cost | Units
    // North   | 1200  | 800  | 50
    // South   | 950   | 620  | 38
    // East    | 1800  | 1100 | 72
    // West    | 1450  | 970  | 61
    // TOTAL   | 5400  | 3490 | 221
in
    Result
