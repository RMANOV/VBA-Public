// ═══════════════════════════════════════════════════════════════
// PATTERN: Calendar / Date Dimension Table
// CATEGORY: Data Modeling | DIFFICULTY: Beginner-Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Generate a complete date dimension from scratch in pure M
// WHY:  Every Power BI / Excel data model needs a date table for time intelligence
// KEY LESSON: Every Power BI model needs a date table. Build in M, not DAX.
// ═══════════════════════════════════════════════════════════════

let
    // ── PARAMETERS ───────────────────────────────────────────
    // Change these to match your data range.
    // In production, derive from: List.Min(YourFactTable[DateColumn])
    StartDate = #date(2024, 1, 1),
    EndDate   = #date(2024, 12, 31),

    // ── STEP 1: Generate the date list ───────────────────────
    // List.Dates creates a list of consecutive dates.
    // Arguments: start, count, step (duration)
    DayCount = Duration.Days(EndDate - StartDate) + 1,
    DateList = List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),

    // ── STEP 2: Convert list to table ────────────────────────
    DateTable = Table.FromList(
        DateList,
        Splitter.SplitByNothing(),    // Each list item becomes one row
        {"Date"},                      // Column name
        null,
        ExtraValues.Error
    ),
    TypedDate = Table.TransformColumnTypes(DateTable, {{"Date", type date}}),

    // ── STEP 3: Enrich with date attributes ──────────────────
    // Each column is derived from the [Date] field using built-in M functions.

    WithYear = Table.AddColumn(TypedDate, "Year",
        each Date.Year([Date]), Int64.Type),

    WithMonth = Table.AddColumn(WithYear, "Month",
        each Date.Month([Date]), Int64.Type),

    WithMonthName = Table.AddColumn(WithMonth, "MonthName",
        each Date.MonthName([Date]), type text),

    // Short month name (first 3 characters) — useful for chart labels
    WithMonthShort = Table.AddColumn(WithMonthName, "MonthShort",
        each Text.Start(Date.MonthName([Date]), 3), type text),

    WithQuarter = Table.AddColumn(WithMonthShort, "Quarter",
        each Date.QuarterOfYear([Date]), Int64.Type),

    // "Q1", "Q2", etc. — human-friendly quarter label
    WithQuarterLabel = Table.AddColumn(WithQuarter, "QuarterLabel",
        each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),

    WithDayOfWeek = Table.AddColumn(WithQuarterLabel, "DayOfWeek",
        each Date.DayOfWeek([Date], Day.Monday),  // Monday = 0
        Int64.Type),

    WithDayName = Table.AddColumn(WithDayOfWeek, "DayName",
        each Date.DayOfWeekName([Date]), type text),

    WithDayOfMonth = Table.AddColumn(WithDayName, "DayOfMonth",
        each Date.Day([Date]), Int64.Type),

    WithWeekNumber = Table.AddColumn(WithDayOfMonth, "WeekNumber",
        each Date.WeekOfYear([Date], Day.Monday), Int64.Type),

    // ── STEP 4: Business logic columns ───────────────────────

    // IsWeekend: Saturday (5) or Sunday (6) when Monday = 0
    WithIsWeekend = Table.AddColumn(WithWeekNumber, "IsWeekend",
        each Date.DayOfWeek([Date], Day.Monday) >= 5, type logical),

    // IsBusinessDay: simple Mon-Fri check.
    // For production, join with a holiday table for accuracy.
    WithIsBusinessDay = Table.AddColumn(WithIsWeekend, "IsBusinessDay",
        each Date.DayOfWeek([Date], Day.Monday) < 5, type logical),

    // ── STEP 5: Sort / display columns ───────────────────────

    // YearMonth: "2024-01" format — sorts correctly as text
    WithYearMonth = Table.AddColumn(WithIsBusinessDay, "YearMonth",
        each Text.From(Date.Year([Date])) & "-" &
             Text.PadStart(Text.From(Date.Month([Date])), 2, "0"),
        type text),

    // DateKey: integer key 20240101 — useful for surrogate keys
    WithDateKey = Table.AddColumn(WithYearMonth, "DateKey",
        each Date.Year([Date]) * 10000
           + Date.Month([Date]) * 100
           + Date.Day([Date]),
        Int64.Type),

    // ── STEP 6: Sort by date for clean output ────────────────
    Result = Table.Sort(WithDateKey, {{"Date", Order.Ascending}})

    // ── RESULT (first 3 rows of 366) ─────────────────────────
    // Date       | Year | Month | MonthName | MonthShort | Quarter | QuarterLabel
    // 2024-01-01 | 2024 | 1     | January   | Jan        | 1       | Q1
    // 2024-01-02 | 2024 | 1     | January   | Jan        | 1       | Q1
    // 2024-01-03 | 2024 | 1     | January   | Jan        | 1       | Q1
    //
    // DayOfWeek | DayName   | DayOfMonth | WeekNumber | IsWeekend | IsBusinessDay
    // 0         | Monday    | 1          | 1          | false     | true
    // 1         | Tuesday   | 2          | 1          | false     | true
    // 2         | Wednesday | 3          | 1          | false     | true
    //
    // YearMonth | DateKey
    // 2024-01   | 20240101
    // 2024-01   | 20240102
    // 2024-01   | 20240103
    //
    // TIPS:
    // - Mark as Date Table in Power BI: Modeling → Mark as Date Table
    // - For holidays: merge with a holiday list and update IsBusinessDay
    // - MonthShort is locale-dependent; use culture parameter if needed
in
    Result
