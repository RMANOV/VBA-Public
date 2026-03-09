// =====================================================================
// PATTERN: Header Cleanup Pipeline (Promote -> Clean -> Type)
// CATEGORY: ETL Foundation | DIFFICULTY: Beginner-Intermediate
// =====================================================================
// WHAT: The universal 3-step entry sequence for raw data imports:
//       Promote headers, clean column names, assign proper types.
// WHY:  Raw CSV/Excel imports have messy headers (spaces, special chars,
//       mixed case) that cause errors downstream. Fix them FIRST.
// KEY LESSON: Every raw import follows Promote -> Clean -> Type.
//       Clean headers early -- downstream pain otherwise.
// =====================================================================

let
    // =================================================================
    // SAMPLE: Raw CSV-like data with messy headers in first row
    // =================================================================
    // Simulates what you get from Csv.Document or a poorly formatted
    // Excel file -- headers are in the first data row.

    RawData = Table.FromRecords({
        // Row 0: This is actually the header row (messy!)
        [Column1 = "  Employee Name ",
         Column2 = "Hire Date (MM/DD)",
         Column3 = "  Annual Salary ($) ",
         Column4 = "Department/Unit",
         Column5 = "Active? (Y/N)"],
        // Rows 1+: Actual data
        [Column1 = "Alice Smith",
         Column2 = "03/15/2020",
         Column3 = "75000",
         Column4 = "Engineering",
         Column5 = "Y"],
        [Column1 = "Bob Jones",
         Column2 = "11/01/2019",
         Column3 = "82000",
         Column4 = "Marketing",
         Column5 = "N"],
        [Column1 = "Carol King",
         Column2 = "06/20/2021",
         Column3 = "91000",
         Column4 = "Engineering",
         Column5 = "Y"],
        [Column1 = " Dave Wilson ",
         Column2 = "01/10/2022",
         Column3 = "67000",
         Column4 = "Sales",
         Column5 = "Y"]
    }),

    // =================================================================
    // STEP 1: PROMOTE HEADERS
    // =================================================================
    // Move the first data row into column headers.
    // Before: Column1, Column2, Column3, ...
    // After:  "  Employee Name ", "Hire Date (MM/DD)", ...

    Promoted = Table.PromoteHeaders(RawData, [PromoteAllScalars = true]),

    // =================================================================
    // STEP 2: CLEAN COLUMN NAMES
    // =================================================================
    // Table.TransformColumnNames applies a function to EVERY header.
    // We chain multiple text operations to normalize them.
    //
    // Cleaning rules:
    //   1. Trim whitespace from both ends
    //   2. Replace spaces with underscores
    //   3. Remove parentheses and their contents (optional)
    //   4. Remove special characters: $, ?, /
    //   5. Collapse multiple underscores into one
    //   6. Remove trailing underscores

    Cleaned = Table.TransformColumnNames(
        Promoted,
        (name) =>
            let
                step1_trim    = Text.Trim(name),
                // Remove parenthetical content: "(MM/DD)" -> ""
                step2_parens  = Text.Replace(
                    Text.Replace(
                        // Simple approach: remove ( and everything to )
                        // For nested parens, use a regex-like loop
                        let
                            openPos  = Text.PositionOf(step1_trim, "("),
                            closePos = Text.PositionOf(step1_trim, ")")
                        in
                            if openPos >= 0 and closePos > openPos then
                                Text.RemoveRange(step1_trim, openPos,
                                    closePos - openPos + 1)
                            else
                                step1_trim
                    , "  ", " "), "  ", " "  // collapse double spaces
                ),
                step3_special = Text.Replace(
                    Text.Replace(
                        Text.Replace(step2_parens, "$", ""),
                    "?", ""),
                "/", "_"),
                step4_spaces  = Text.Replace(Text.Trim(step3_special), " ", "_"),
                // Collapse double underscores
                step5_clean   = Text.Replace(step4_spaces, "__", "_"),
                // Remove trailing underscore
                step6_trail   = if Text.EndsWith(step5_clean, "_")
                    then Text.RemoveRange(step5_clean,
                        Text.Length(step5_clean) - 1, 1)
                    else step5_clean
            in
                step6_trail
    ),
    // Headers now: Employee_Name, Hire_Date, Annual_Salary,
    //              Department_Unit, Active

    // =================================================================
    // STEP 3: ASSIGN PROPER TYPES
    // =================================================================
    // Manual type assignment -- the most reliable approach when you
    // know your schema.

    Typed = Table.TransformColumnTypes(Cleaned, {
        {"Employee_Name",   type text},
        {"Hire_Date",       type date},
        {"Annual_Salary",   Int64.Type},
        {"Department_Unit", type text},
        {"Active",          type text}
    }),

    // =================================================================
    // BONUS: Auto-Detect Types by Inspecting Values
    // =================================================================
    // When you DON'T know the schema, inspect the first non-null value
    // in each column and guess the type. Useful for dynamic sources.

    AutoDetectTypes = (tbl as table) as table =>
        let
            cols = Table.ColumnNames(tbl),
            typeSpecs = List.Transform(cols, (colName) =>
                let
                    // Get first non-null value
                    colValues = Table.Column(tbl, colName),
                    firstVal = List.First(
                        List.RemoveNulls(colValues), null
                    ),
                    // Determine type from value
                    detectedType =
                        if firstVal = null then type text
                        else if Value.Is(firstVal, type number) then type number
                        else if Value.Is(firstVal, type date) then type date
                        else if Value.Is(firstVal, type datetime) then type datetime
                        else if Value.Is(firstVal, type logical) then type logical
                        else type text
                in
                    {colName, detectedType}
            )
        in
            Table.TransformColumnTypes(tbl, typeSpecs),

    // =================================================================
    // BONUS: Reusable Header Cleaner Function
    // =================================================================
    // Extract the cleaning logic as a reusable function for any table.

    CleanHeaders = (tbl as table) as table =>
        Table.TransformColumnNames(tbl, (name) =>
            let
                trimmed = Text.Trim(name),
                noSpecial = Text.Select(trimmed,
                    {"A".."Z"} & {"a".."z"} & {"0".."9"} & {" ", "_"}),
                underscored = Text.Replace(Text.Trim(noSpecial), " ", "_"),
                collapsed = Text.Replace(underscored, "__", "_")
            in
                collapsed
        ),

    // =================================================================
    // FINAL OUTPUT
    // =================================================================
    // The cleaned, typed table ready for downstream use
    Result = Typed

in
    Result

// =====================================================================
// THE PIPELINE IN ONE GLANCE:
//
//   Raw Import
//       |
//       v
//   Table.PromoteHeaders        <- Row 1 becomes column names
//       |
//       v
//   Table.TransformColumnNames  <- Trim, normalize, remove junk
//       |
//       v
//   Table.TransformColumnTypes  <- Assign correct data types
//       |
//       v
//   Clean Table (ready for transforms, joins, loads)
//
// NOTES:
// - PromoteAllScalars = true handles dates/numbers in the header row
// - Always trim BEFORE other text operations
// - Type assignment can fail silently (values become null) -- inspect
//   the "Errors" count at the bottom of each column in PQ Editor
// - For production: add error handling with try/otherwise on type step
// =====================================================================