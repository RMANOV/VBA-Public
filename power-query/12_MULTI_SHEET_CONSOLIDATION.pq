// =====================================================================
// PATTERN: Multi-Sheet Consolidation
// CATEGORY: ETL / Data Loading | DIFFICULTY: Intermediate
// =====================================================================
// WHAT: Dynamically stack all sheets from an Excel workbook into one
//       unified table, tagging each row with its source sheet name.
// WHY:  Monthly/regional data often lives in separate sheets with
//       identical layouts. Manual copy-paste doesn't scale.
// KEY LESSON: Excel.Workbook returns metadata. Filter -> Extract [Data]
//       -> Tag origin -> Combine. Scales from 2 to 200 sheets.
// =====================================================================

let
    // =================================================================
    // STEP 0: SIMULATED DATA (replace with real file in production)
    // =================================================================
    // In production, replace SimulatedSheets with:
    //   Source = Excel.Workbook(File.Contents("C:\path\to\file.xlsx"), null, true)
    //
    // Excel.Workbook returns a table with columns:
    //   [Name] = sheet/table name
    //   [Data] = the actual data as a table (headers NOT promoted yet)
    //   [Item] = same as Name for sheets
    //   [Kind] = "Sheet" or "Table" or "DefinedName"
    //   [Hidden] = true/false

    // Simulate 3 monthly sheets with identical structure
    JanData = Table.FromRecords({
        [Column1 = "Product", Column2 = "Sales", Column3 = "Region"],
        [Column1 = "Widget A", Column2 = "1200", Column3 = "North"],
        [Column1 = "Widget B", Column2 = "850",  Column3 = "South"],
        [Column1 = "Widget C", Column2 = "2100", Column3 = "East"]
    }),
    FebData = Table.FromRecords({
        [Column1 = "Product", Column2 = "Sales", Column3 = "Region"],
        [Column1 = "Widget A", Column2 = "1350", Column3 = "North"],
        [Column1 = "Widget B", Column2 = "920",  Column3 = "South"]
    }),
    MarData = Table.FromRecords({
        [Column1 = "Product", Column2 = "Sales", Column3 = "Region"],
        [Column1 = "Widget A", Column2 = "1500", Column3 = "North"],
        [Column1 = "Widget C", Column2 = "1800", Column3 = "East"],
        [Column1 = "Widget D", Column2 = "450",  Column3 = "West"]
    }),

    // Simulate the Excel.Workbook metadata table
    SimulatedSheets = Table.FromRecords({
        [Name = "Jan_Sales", Data = JanData, Item = "Jan_Sales",
         Kind = "Sheet", Hidden = false],
        [Name = "Feb_Sales", Data = FebData, Item = "Feb_Sales",
         Kind = "Sheet", Hidden = false],
        [Name = "Mar_Sales", Data = MarData, Item = "Mar_Sales",
         Kind = "Sheet", Hidden = false],
        [Name = "Config",    Data = JanData, Item = "Config",
         Kind = "Sheet", Hidden = true],
        [Name = "SalesTable", Data = JanData, Item = "SalesTable",
         Kind = "Table", Hidden = false]
    }),

    // =================================================================
    // PRODUCTION VERSION (uncomment and replace path):
    // =================================================================
    // Source = Excel.Workbook(
    //     File.Contents("C:\Reports\SalesData.xlsx"),
    //     null,   // null = don't promote headers yet (we do it manually)
    //     true    // true = use InferSheetDimensions for performance
    // ),

    Source = SimulatedSheets,

    // =================================================================
    // STEP 1: FILTER TO RELEVANT SHEETS
    // =================================================================
    // Keep only visible sheets (not hidden config sheets or named tables)
    SheetsOnly = Table.SelectRows(Source, each
        [Kind] = "Sheet" and [Hidden] = false
    ),

    // Optional: filter by name pattern (e.g., only sheets ending in "_Sales")
    FilteredByName = Table.SelectRows(SheetsOnly, each
        Text.EndsWith([Name], "_Sales")
    ),

    // =================================================================
    // STEP 2: PROCESS EACH SHEET
    // =================================================================
    // For each sheet: promote headers, add source column, clean types
    ProcessSheet = (sheetRow as record) as table =>
        let
            raw = sheetRow[Data],
            sheetName = sheetRow[Name],

            // Promote the first row to headers
            promoted = Table.PromoteHeaders(raw, [PromoteAllScalars = true]),

            // Add a column to track which sheet each row came from
            tagged = Table.AddColumn(
                promoted, "SourceSheet", each sheetName, type text
            ),

            // Assign types (adjust to your actual schema)
            typed = Table.TransformColumnTypes(tagged, {
                {"Product", type text},
                {"Sales",   Int64.Type},
                {"Region",  type text}
            })
        in
            typed,

    // Apply the processing function to each sheet
    ProcessedTables = Table.AddColumn(
        FilteredByName,
        "ProcessedData",
        each ProcessSheet(_)
    ),

    // =================================================================
    // STEP 3: COMBINE ALL SHEETS INTO ONE TABLE
    // =================================================================
    // Extract the processed tables and stack them vertically
    AllData = Table.Combine(ProcessedTables[ProcessedData]),

    // =================================================================
    // STEP 4: HANDLE TYPE INCONSISTENCIES (OPTIONAL)
    // =================================================================
    // When sheets have slightly different schemas, Table.Combine may
    // introduce nulls or type errors. Clean up here.
    CleanedResult = Table.ReplaceErrorValues(AllData, {
        {"Sales", 0}    // replace type conversion errors with 0
    }),

    // =================================================================
    // ALTERNATIVE: Quick one-liner for simple cases
    // =================================================================
    // When all sheets have identical schemas and you don't need
    // the source sheet name:
    //
    // QuickCombine = Table.Combine(
    //     List.Transform(
    //         Table.SelectRows(Source, each [Kind] = "Sheet")[Data],
    //         each Table.PromoteHeaders(_, [PromoteAllScalars = true])
    //     )
    // ),

    // =================================================================
    // MULTI-FILE VARIANT (Folder.Files)
    // =================================================================
    // To combine sheets from MULTIPLE Excel files in a folder:
    //
    // FolderSource = Folder.Files("C:\Reports\Monthly"),
    // ExcelFiles = Table.SelectRows(FolderSource, each
    //     Text.EndsWith([Name], ".xlsx")),
    // AllWorkbooks = Table.AddColumn(ExcelFiles, "Sheets", each
    //     Excel.Workbook([Content], null, true)),
    // Expanded = Table.ExpandTableColumn(AllWorkbooks, "Sheets",
    //     {"Name", "Data", "Kind"}),
    // FilteredSheets = Table.SelectRows(Expanded, each
    //     [Kind] = "Sheet"),
    // -- then proceed with PromoteHeaders, AddColumn, Combine as above

    Result = CleanedResult

in
    Result

// =====================================================================
// PRODUCTION CHECKLIST:
// 1. Replace SimulatedSheets with Excel.Workbook(File.Contents(...))
// 2. Adjust the sheet name filter (Text.EndsWith, Text.Contains, etc.)
// 3. Update type assignments in ProcessSheet to match your schema
// 4. Add error handling: Table.ReplaceErrorValues for type mismatches
// 5. Consider adding a row number column per sheet for audit trail:
//      Table.AddIndexColumn(tagged, "RowInSheet", 1, 1)
//
// COMMON PITFALLS:
// - Forgetting [PromoteAllScalars = true] -> numeric headers stay as numbers
// - Hidden sheets appearing in results -> always filter [Hidden] = false
// - Sheets with different column orders -> Table.Combine aligns by NAME,
//   not position, so this works fine as long as names match
// - Extra blank rows at bottom of sheets -> add Table.SelectRows to
//   remove rows where key column is null
// =====================================================================