// =====================================================================
// PATTERN: List.Accumulate (Fold / Reduce)
// CATEGORY: Functional Patterns | DIFFICULTY: Advanced
// =====================================================================
// WHAT: Transform a list into a single value or complex structure
//       using iterative accumulation (fold/reduce).
// WHY:  When you need running totals, chained transformations, or
//       must build a result by processing items one at a time.
// KEY LESSON: List.Accumulate is M's reduce/fold. The seed (initial
//       value) can be ANY type -- number, text, list, record, or table.
// =====================================================================
//
// SIGNATURE:
//   List.Accumulate(list, seed, accumulator(state, current) => newState)
//
//   list        = items to iterate over
//   seed        = starting value (any type)
//   accumulator = function called for each item; receives the running
//                 state and the current item, returns the new state
//
// =====================================================================

let
    // =================================================================
    // EXAMPLE 1: Running Totals (Cumulative Sums)
    // =================================================================
    // Given monthly sales, produce a list of cumulative sums.
    // Sales: {100, 250, 175, 300, 125}
    // Result: {100, 350, 525, 825, 950}
    //
    // Seed is a record with two fields:
    //   [sum]     = running total (number)
    //   [results] = accumulated list of cumulative values

    Sales = {100, 250, 175, 300, 125},

    // The accumulator adds each sale to [sum] and appends to [results]
    RunningState = List.Accumulate(
        Sales,
        [sum = 0, results = {}],               // seed: record
        (state, current) =>
            let
                newSum = state[sum] + current
            in
                [sum = newSum, results = state[results] & {newSum}]
    ),

    // Extract the cumulative list from the final state
    RunningTotals = RunningState[results],
    // RunningTotals = {100, 350, 525, 825, 950}

    // Build a nice table to show both original and cumulative
    RunningTotalsTable = Table.FromColumns(
        {Sales, RunningTotals},
        {"MonthlySale", "CumulativeTotal"}
    ),

    // =================================================================
    // EXAMPLE 2: Chain of Table Transformations
    // =================================================================
    // Apply a list of text-cleaning functions to a table column
    // sequentially. Each function takes the table from the previous
    // step and returns a transformed table.

    RawData = Table.FromRecords({
        [Name = "  ALICE SMITH  ", City = "  new york  "],
        [Name = "  BOB JONES  ",  City = "  chicago  "],
        [Name = "  carol KING ",  City = " MIAMI "]
    }),

    // Each transformation is a function: (table) => table
    // They will be applied in order: Trim -> Proper Case -> all lower for City
    Transformations = {
        // Step 1: Trim whitespace from Name
        (tbl) => Table.TransformColumns(tbl, {
            {"Name", Text.Trim, type text}
        }),
        // Step 2: Proper case for Name (e.g., "ALICE SMITH" -> "Alice Smith")
        (tbl) => Table.TransformColumns(tbl, {
            {"Name", Text.Proper, type text}
        }),
        // Step 3: Trim and lower-case the City column
        (tbl) => Table.TransformColumns(tbl, {
            {"City", each Text.Lower(Text.Trim(_)), type text}
        })
    },

    // List.Accumulate chains the transformations: seed = RawData table
    CleanedData = List.Accumulate(
        Transformations,
        RawData,                                // seed: the raw table
        (currentTable, transformFn) => transformFn(currentTable)
    ),
    // Result: Name = "Alice Smith", City = "new york" etc.

    // =================================================================
    // EXAMPLE 3: Build a Comma-Separated String from a List
    // =================================================================
    // Classic reduce: join list items into a delimited string.
    // Note: Text.Combine is simpler for this specific case, but
    // List.Accumulate shows the general pattern.

    Fruits = {"Apple", "Banana", "Cherry", "Date", "Elderberry"},

    // Seed is empty text. On each iteration, append delimiter + item.
    // Skip the comma before the first item by checking if state is empty.
    CommaSeparated = List.Accumulate(
        Fruits,
        "",                                     // seed: empty text
        (state, current) =>
            if state = "" then current
            else state & ", " & current
    ),
    // CommaSeparated = "Apple, Banana, Cherry, Date, Elderberry"

    // -----------------------------------------------------------------
    // COMPARISON: The idiomatic way (prefer this in production)
    // -----------------------------------------------------------------
    CommaSeparated_Idiomatic = Text.Combine(Fruits, ", "),

    // =================================================================
    // FINAL: Combine all examples into a summary table
    // =================================================================
    Summary = Table.FromRecords({
        [Example = "Running Totals",
         Input   = Text.Combine(List.Transform(Sales, Number.ToText), ", "),
         Output  = Text.Combine(List.Transform(RunningTotals, Number.ToText), ", ")],
        [Example = "Chained Transforms",
         Input   = "  ALICE SMITH  ",
         Output  = CleanedData{0}[Name]],
        [Example = "Comma Join",
         Input   = Text.Combine({"Apple", "...", "Elderberry"}, ", "),
         Output  = CommaSeparated]
    })

in
    Summary

// =====================================================================
// NOTES:
// - List.Accumulate processes items LEFT to RIGHT (like foldl)
// - The seed type determines what you can do in the accumulator
// - For simple sums, List.Sum is faster; use Accumulate for complex state
// - List.Accumulate is NOT lazy -- it processes the entire list
// - Performance: O(n) iterations, but string concatenation in a loop
//   can be O(n^2) -- prefer Text.Combine for large lists
// =====================================================================