// ═══════════════════════════════════════════════════════════════
// PATTERN: Conditional Columns — Three Approaches
// CATEGORY: Data Transformation | DIFFICULTY: Beginner-Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Assign grades/categories based on numeric thresholds
//       using nested if, table-driven lookup, and list ranges.
// WHY:  Business rules like grading, tiering, and bucketing are
//       everywhere — pick the right approach for the complexity.
// KEY LESSON: Nested if is fine for simple cases. Table-driven
//             lookup separates rules from logic — change
//             boundaries without touching code.
// ═══════════════════════════════════════════════════════════════

let
    // ── Sample data: student scores ─────────────────────────────
    Students = Table.FromRecords({
        [Name = "Alice",   Score = 95],
        [Name = "Bob",     Score = 82],
        [Name = "Charlie", Score = 74],
        [Name = "Diana",   Score = 67],
        [Name = "Eve",     Score = 55],
        [Name = "Frank",   Score = 43],
        [Name = "Grace",   Score = 90],
        [Name = "Hank",    Score = 78],
        [Name = "Iris",    Score = 100],
        [Name = "Jack",    Score = 59]
    }),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 1: Nested if...then...else
    // ═══════════════════════════════════════════════════════════
    // Best for: 3-4 tiers, simple logic, quick-and-dirty work.
    // Readable if you indent cleanly. Gets messy beyond 5 levels.
    Pattern1_NestedIf = Table.AddColumn(
        Students,
        "Grade_If",
        each
            if [Score] >= 90 then "A"
            else if [Score] >= 80 then "B"
            else if [Score] >= 70 then "C"
            else if [Score] >= 60 then "D"
            else "F",
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 2: Table-driven lookup
    // ═══════════════════════════════════════════════════════════
    // Best for: 5+ tiers, rules that change frequently, or when
    // business users maintain the thresholds in a spreadsheet.
    //
    // The rules are DATA, not code. To change boundaries,
    // edit the table — zero code changes needed.

    // Define the lookup table (could come from Excel named range)
    GradeTable = Table.FromRecords({
        [MinScore =  0, MaxScore = 59, Grade = "F"],
        [MinScore = 60, MaxScore = 69, Grade = "D"],
        [MinScore = 70, MaxScore = 79, Grade = "C"],
        [MinScore = 80, MaxScore = 89, Grade = "B"],
        [MinScore = 90, MaxScore = 100, Grade = "A"]
    }),

    // Lookup function: find the row where score falls in range
    LookupGrade = (score as number) as text =>
        let
            match = Table.SelectRows(
                GradeTable,
                each [MinScore] <= score and score <= [MaxScore]
            ),
            grade = if Table.RowCount(match) > 0
                    then match{0}[Grade]
                    else "N/A"  // Safety net for out-of-range values
        in
            grade,

    Pattern2_TableDriven = Table.AddColumn(
        Pattern1_NestedIf,
        "Grade_Table",
        each LookupGrade([Score]),
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 3: List-based bracket detection
    // ═══════════════════════════════════════════════════════════
    // Best for: fixed brackets where you just need the index.
    // Compact and functional, but less readable for newcomers.
    //
    // Idea: define breakpoints as a sorted list, count how many
    // the score exceeds, use that count as an index into labels.

    Breakpoints = {60, 70, 80, 90},     // Ascending thresholds
    Labels      = {"F", "D", "C", "B", "A"},  // One more label than breakpoints

    // Count how many breakpoints the score meets or exceeds
    GetBracket = (score as number) as text =>
        let
            // List.Select keeps breakpoints that score >= threshold
            exceeded = List.Select(Breakpoints, each score >= _),
            index    = List.Count(exceeded)  // 0 = below all, 4 = above all
        in
            Labels{index},

    Pattern3_ListBased = Table.AddColumn(
        Pattern2_TableDriven,
        "Grade_List",
        each GetBracket([Score]),
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // COMPARISON: When to use each pattern
    // ═══════════════════════════════════════════════════════════
    //
    // Pattern 1 (Nested if):
    //   + Simple, readable for 3-4 levels
    //   + No helper tables needed
    //   - Gets unwieldy beyond 5 levels
    //   - Rules buried in code — hard for business users to update
    //
    // Pattern 2 (Table-driven):
    //   + Rules are data — change without touching code
    //   + Business users can maintain thresholds in Excel
    //   + Scales to any number of tiers
    //   - Slightly more setup
    //   - Table.SelectRows per row can be slower on huge datasets
    //
    // Pattern 3 (List-based):
    //   + Most compact code
    //   + Fast — no table scanning per row
    //   - Less intuitive for non-programmers
    //   - Only works for sequential numeric brackets
    //
    // ═══════════════════════════════════════════════════════════

    // ── Clean up: keep only one grade column for final output ────
    Result = Table.SelectColumns(Pattern3_ListBased,
        {"Name", "Score", "Grade_If", "Grade_Table", "Grade_List"}
    )

in
    Result