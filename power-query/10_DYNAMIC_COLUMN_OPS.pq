// =====================================================================
// PATTERN: Dynamic Column Operations
// CATEGORY: Schema-Agnostic Transforms | DIFFICULTY: Advanced
// =====================================================================
// WHAT: Discover and operate on columns dynamically, without hard-coding
//       column names -- works regardless of how many columns exist.
// WHY:  Hard-coded column names break when source schema changes.
//       Dynamic patterns adapt automatically to new/removed columns.
// KEY LESSON: Hard-coding column names = brittle queries. Discover
//       dynamically with Table.ColumnNames + List.Select.
// =====================================================================

let
    // =================================================================
    // SAMPLE DATA
    // =================================================================
    Source = Table.FromRecords({
        [Name = "North",  Q1 = 100, Q2 = 250, Q3 = 175, Q4 = 300],
        [Name = "South",  Q1 = 200, Q2 = 150, Q3 = 225, Q4 = 275],
        [Name = "East",   Q1 = 175, Q2 = 300, Q3 = 200, Q4 = 350],
        [Name = "West",   Q1 = 125, Q2 = 175, Q3 = 150, Q4 = 200]
    }, type table [Name = text, Q1 = Int64.Type, Q2 = Int64.Type,
                   Q3 = Int64.Type, Q4 = Int64.Type]),

    AllColumns = Table.ColumnNames(Source),
    // {"Name", "Q1", "Q2", "Q3", "Q4"}

    // =================================================================
    // PATTERN 1: Detect All Numeric Columns Dynamically
    // =================================================================
    // Strategy: Check the TYPE of each column in the first row.
    // Value.Is tests whether a value conforms to a given type.

    FirstRow = Source{0},   // record: [Name="North", Q1=100, ...]

    NumericColumns = List.Select(
        AllColumns,
        (colName) =>
            let
                val = Record.Field(FirstRow, colName)
            in
                Value.Is(val, type number)
    ),
    // NumericColumns = {"Q1", "Q2", "Q3", "Q4"}

    // Alternative: filter by column type from table schema
    // (works even if first row has nulls)
    NumericColumns_Alt = List.Select(
        AllColumns,
        (colName) =>
            let
                colType = Type.TableColumn(Value.Type(Source), colName)
            in
                Type.Is(colType, type nullable number)
                or Type.Is(colType, Int64.Type)
    ),

    // =================================================================
    // PATTERN 2: Bulk Transform All Numeric Columns
    // =================================================================
    // Multiply every numeric column by 2, without listing them manually.
    // Table.TransformColumns expects a list of {name, function, type} items.

    TransformSpecs = List.Transform(
        NumericColumns,
        (colName) => {colName, each _ * 2, Int64.Type}
    ),

    Doubled = Table.TransformColumns(Source, TransformSpecs),
    // Q1: 100->200, Q2: 250->500, etc.

    // =================================================================
    // PATTERN 3: Dynamic Column Rename
    // =================================================================
    // Rename Q1 -> "Quarter 1", Q2 -> "Quarter 2", etc.
    // Build rename pairs dynamically from the column name pattern.

    RenameMap = List.Transform(
        NumericColumns,
        (colName) =>
            let
                // Extract the number from "Q1", "Q2", etc.
                quarterNum = Text.AfterDelimiter(colName, "Q")
            in
                {colName, "Quarter " & quarterNum}
    ),
    // RenameMap = {{"Q1","Quarter 1"}, {"Q2","Quarter 2"}, ...}

    Renamed = Table.RenameColumns(Source, RenameMap),

    // =================================================================
    // PATTERN 4: Dynamic Sum of All Numeric Columns per Row
    // =================================================================
    // Add a "Total" column that sums all numeric columns, regardless
    // of how many there are.

    WithTotal = Table.AddColumn(
        Source,
        "Total",
        (row) =>
            List.Sum(
                List.Transform(
                    NumericColumns,
                    (colName) => Record.Field(row, colName)
                )
            ),
        Int64.Type
    ),
    // North: 100+250+175+300 = 825, South: 200+150+225+275 = 850, ...

    // =================================================================
    // BONUS: Dynamic Percentage of Total per Column
    // =================================================================
    // For each numeric column, add a "Q1_Pct", "Q2_Pct" etc. showing
    // the percentage of the row total.

    PctSpecs = List.Transform(
        NumericColumns,
        (colName) => {
            colName & "_Pct",
            (row) =>
                let
                    total = List.Sum(
                        List.Transform(NumericColumns,
                            (c) => Record.Field(row, c))
                    ),
                    value = Record.Field(row, colName)
                in
                    if total = 0 then 0
                    else Number.Round(value / total * 100, 1)
        }
    ),

    WithPct = List.Accumulate(
        PctSpecs,
        WithTotal,
        (tbl, spec) => Table.AddColumn(tbl, spec{0}, spec{1}, type number)
    ),

    // =================================================================
    // FINAL OUTPUT: Show the table with dynamic total
    // =================================================================
    // Change to Doubled, Renamed, or WithPct to explore other patterns
    Result = WithTotal

in
    Result

// =====================================================================
// NOTES:
// - Table.ColumnNames returns names in column order (left to right)
// - Record.Field(row, name) is the dynamic equivalent of row[Name]
// - List.Select filters; List.Transform maps -- together they handle
//   most dynamic column discovery and transformation needs
// - For production: also handle null values in numeric columns:
//     each if _ = null then 0 else _ * 2
// - Type.TableColumn gets the declared type from the table schema,
//   which is more reliable than inspecting row values
// =====================================================================