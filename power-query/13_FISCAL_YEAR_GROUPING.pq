// =====================================================================
// PATTERN: Fiscal Year Grouping
// CATEGORY: Date Logic / Business Periods | DIFFICULTY: Intermediate
// =====================================================================
// WHAT: Map calendar dates to fiscal year, quarter, and month using
//       a configurable fiscal year start month.
// WHY:  Many organizations use non-calendar fiscal years (April-March,
//       July-June, October-September). Reports must group by fiscal
//       periods, not calendar periods.
// KEY LESSON: Fiscal year = if month >= startMonth then year else year-1.
//       Parameterize the start month for reuse across queries.
// =====================================================================

let
    // =================================================================
    // PARAMETER: Fiscal Year Start Month
    // =================================================================
    // Change this single value to shift all fiscal calculations.
    // Examples:
    //   4  = April start  (Apr 2024 - Mar 2025 = FY2024-25)
    //   7  = July start   (Jul 2024 - Jun 2025 = FY2024-25)
    //   10 = October start (Oct 2024 - Sep 2025 = FY2024-25)
    //   1  = Calendar year (Jan-Dec, standard)

    FiscalYearStartMonth = 4,

    // =================================================================
    // SAMPLE DATA: Transactions spanning multiple fiscal years
    // =================================================================
    Source = Table.FromRecords({
        [Date = #date(2024, 1, 15),  Amount = 1200, Category = "Sales"],
        [Date = #date(2024, 3, 28),  Amount = 800,  Category = "Sales"],
        [Date = #date(2024, 4, 1),   Amount = 1500, Category = "Sales"],
        [Date = #date(2024, 6, 15),  Amount = 2200, Category = "Services"],
        [Date = #date(2024, 9, 30),  Amount = 950,  Category = "Sales"],
        [Date = #date(2024, 10, 5),  Amount = 1800, Category = "Services"],
        [Date = #date(2024, 12, 20), Amount = 3100, Category = "Sales"],
        [Date = #date(2025, 1, 10),  Amount = 1400, Category = "Services"],
        [Date = #date(2025, 3, 31),  Amount = 2600, Category = "Sales"],
        [Date = #date(2025, 4, 1),   Amount = 1100, Category = "Sales"],
        [Date = #date(2025, 7, 15),  Amount = 1750, Category = "Services"]
    }, type table [Date = date, Amount = Int64.Type, Category = text]),

    // =================================================================
    // FISCAL YEAR
    // =================================================================
    // Core formula: if the calendar month is >= the fiscal start month,
    // the date belongs to the fiscal year starting in the current
    // calendar year. Otherwise, it belongs to the previous year's
    // fiscal year.
    //
    // Example with FiscalYearStartMonth = 4 (April):
    //   March 2024 (month 3 < 4) -> FY2023-24
    //   April 2024 (month 4 >= 4) -> FY2024-25

    WithFiscalYear = Table.AddColumn(
        Source,
        "FiscalYear",
        each
            let
                calYear  = Date.Year([Date]),
                calMonth = Date.Month([Date]),
                fyStart  = if calMonth >= FiscalYearStartMonth
                           then calYear
                           else calYear - 1,
                fyEnd    = fyStart + 1,
                // Format as "FY2024-25"
                label    = "FY" & Text.From(fyStart) & "-"
                         & Text.End(Text.From(fyEnd), 2)
            in
                label,
        type text
    ),

    // =================================================================
    // FISCAL QUARTER
    // =================================================================
    // Shift the month so that the fiscal start month becomes month 1,
    // then divide into quarters.
    //
    // Formula: FiscalMonth = ((calMonth - startMonth) mod 12) + 1
    //          FiscalQuarter = ceil(FiscalMonth / 3)
    //
    // Example with start = 4:
    //   Apr(4) -> FM1  -> FQ1    Jul(7) -> FM4  -> FQ2
    //   May(5) -> FM2  -> FQ1    Oct(10)-> FM7  -> FQ3
    //   Jun(6) -> FM3  -> FQ1    Jan(1) -> FM10 -> FQ4

    WithFiscalQuarter = Table.AddColumn(
        WithFiscalYear,
        "FiscalQuarter",
        each
            let
                calMonth    = Date.Month([Date]),
                // Shift so fiscal start = month 1 (0-indexed then +1)
                fiscalMonth = Number.Mod(calMonth - FiscalYearStartMonth, 12) + 1,
                quarter     = Number.RoundUp(fiscalMonth / 3)
            in
                "FQ" & Text.From(quarter),
        type text
    ),

    // =================================================================
    // FISCAL MONTH (Sequential Numbering)
    // =================================================================
    // Fiscal Month 1 = the start month, counting up to 12.
    // Useful for sorting and period comparisons.

    WithFiscalMonth = Table.AddColumn(
        WithFiscalQuarter,
        "FiscalMonth",
        each
            let
                calMonth = Date.Month([Date]),
                fm = Number.Mod(calMonth - FiscalYearStartMonth, 12) + 1
            in
                fm,
        Int64.Type
    ),

    // =================================================================
    // FISCAL YEAR NUMERIC (for sorting/filtering)
    // =================================================================
    // Numeric version of FY start year (e.g., 2024 for FY2024-25)
    WithFYNumeric = Table.AddColumn(
        WithFiscalMonth,
        "FY_Start",
        each
            let
                calYear  = Date.Year([Date]),
                calMonth = Date.Month([Date])
            in
                if calMonth >= FiscalYearStartMonth then calYear else calYear - 1,
        Int64.Type
    ),

    // =================================================================
    // GROUP BY FISCAL YEAR (Aggregation)
    // =================================================================
    // Summarize total amount by fiscal year
    FYSummary = Table.Group(
        WithFYNumeric,
        {"FiscalYear"},
        {
            {"TotalAmount",  each List.Sum([Amount]),      Int64.Type},
            {"RowCount",     each Table.RowCount(_),       Int64.Type},
            {"AvgAmount",    each List.Average([Amount]),   type number},
            {"Categories",   each Text.Combine(
                List.Distinct([Category]), ", "),           type text}
        }
    ),

    // =================================================================
    // GROUP BY FISCAL YEAR + QUARTER (Drill-down)
    // =================================================================
    FYQSummary = Table.Group(
        WithFYNumeric,
        {"FiscalYear", "FiscalQuarter"},
        {
            {"TotalAmount", each List.Sum([Amount]), Int64.Type},
            {"RowCount",    each Table.RowCount(_),  Int64.Type}
        }
    ),

    // Sort by FY then quarter
    FYQSorted = Table.Sort(FYQSummary, {
        {"FiscalYear", Order.Ascending},
        {"FiscalQuarter", Order.Ascending}
    }),

    // =================================================================
    // REUSABLE FUNCTION: Add All Fiscal Columns at Once
    // =================================================================
    // Wrap the logic into a function for use in other queries.

    AddFiscalColumns = (tbl as table, dateCol as text, startMonth as number) as table =>
        let
            step1 = Table.AddColumn(tbl, "FiscalYear", each
                let y = Date.Year(Record.Field(_, dateCol)),
                    m = Date.Month(Record.Field(_, dateCol)),
                    s = if m >= startMonth then y else y - 1
                in "FY" & Text.From(s) & "-" & Text.End(Text.From(s+1), 2),
                type text),
            step2 = Table.AddColumn(step1, "FiscalQuarter", each
                let m = Date.Month(Record.Field(_, dateCol)),
                    fm = Number.Mod(m - startMonth, 12) + 1
                in "FQ" & Text.From(Number.RoundUp(fm / 3)),
                type text),
            step3 = Table.AddColumn(step2, "FiscalMonth", each
                let m = Date.Month(Record.Field(_, dateCol))
                in Number.Mod(m - startMonth, 12) + 1,
                Int64.Type)
        in step3,

    // Usage: AddFiscalColumns(MyTable, "InvoiceDate", 4)

    // =================================================================
    // FINAL OUTPUT
    // =================================================================
    // Options to explore:
    //   WithFYNumeric  = detailed rows with all fiscal columns
    //   FYSummary      = grouped by fiscal year
    //   FYQSorted      = grouped by fiscal year + quarter

    Result = WithFYNumeric

in
    Result

// =====================================================================
// FISCAL YEAR REFERENCE (FiscalYearStartMonth = 4, April):
//
//   Calendar Date  | FY        | FQ  | FM
//   -------------- | --------- | --- | --
//   2024-01-15     | FY2023-24 | FQ4 | 10
//   2024-03-28     | FY2023-24 | FQ4 | 12
//   2024-04-01     | FY2024-25 | FQ1 | 1
//   2024-06-15     | FY2024-25 | FQ1 | 3
//   2024-09-30     | FY2024-25 | FQ2 | 6
//   2024-12-20     | FY2024-25 | FQ3 | 9
//   2025-03-31     | FY2024-25 | FQ4 | 12
//   2025-04-01     | FY2025-26 | FQ1 | 1
//
// NOTES:
// - Number.Mod handles the wrap-around (Jan after fiscal start in Apr)
// - For calendar year (startMonth=1), all formulas reduce to identity
// - The FY label format "FY2024-25" is common; adjust to your standard
// - Add FiscalWeek with: Number.RoundUp(fiscalDayOfYear / 7)
// =====================================================================