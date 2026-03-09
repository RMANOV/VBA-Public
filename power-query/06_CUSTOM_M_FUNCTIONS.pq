// ═══════════════════════════════════════════════════════════════
// PATTERN: Custom M Functions
// CATEGORY: Reusability / Architecture | DIFFICULTY: Intermediate-Advanced
// ═══════════════════════════════════════════════════════════════
// WHAT: Build reusable, parameterized functions as first-class values
// WHY:  DRY principle — define transformation logic once, apply everywhere
// KEY LESSON: M functions are values. Define once, invoke anywhere.
//             Use optional + null-check for flexibility.
// ═══════════════════════════════════════════════════════════════

let
    // ═══════════════════════════════════════════════════════════
    // FUNCTION 1: fnCleanCurrency
    // ═══════════════════════════════════════════════════════════
    // Strips currency symbols, thousand separators, and whitespace,
    // then parses to number. Handles "$1,234.56", "€ 1.234,56", "1 500".
    //
    // M functions are defined with (parameters) => body.
    // "optional" means the caller can omit it (defaults to null).

    fnCleanCurrency = (input as nullable text, optional decimalSep as text) as nullable number =>
        let
            // Default decimal separator is "." (override with "," for EU formats)
            sep = if decimalSep is null then "." else decimalSep,

            // Step 1: Handle null/empty input
            result = if input is null or Text.Trim(input) = "" then
                null
            else
                let
                    trimmed = Text.Trim(input),

                    // Step 2: Remove everything except digits, minus, and the decimal separator
                    // Keep only: 0-9, "-", and the chosen decimal separator
                    allowedChars = {"0","1","2","3","4","5","6","7","8","9","-", sep},
                    chars = Text.ToList(trimmed),
                    cleaned = Text.Combine(
                        List.Select(chars, each List.Contains(allowedChars, _))
                    ),

                    // Step 3: Normalize decimal separator to "." for Number.From
                    normalized = if sep = "," then
                        Text.Replace(cleaned, ",", ".")
                    else
                        cleaned,

                    // Step 4: Parse — with try for safety
                    parsed = try Number.From(normalized) otherwise null
                in
                    parsed
        in
            result,

    // ═══════════════════════════════════════════════════════════
    // FUNCTION 2: fnSafeDivide
    // ═══════════════════════════════════════════════════════════
    // Division with zero and null protection.
    // Returns a configurable fallback instead of error.

    fnSafeDivide = (
        numerator as nullable number,
        denominator as nullable number,
        optional fallback as nullable number
    ) as nullable number =>
        let
            fb = if fallback is null then null else fallback
        in
            if numerator is null or denominator is null then fb
            else if denominator = 0 then fb
            else numerator / denominator,

    // ═══════════════════════════════════════════════════════════
    // FUNCTION 3: fnDatesBetween
    // ═══════════════════════════════════════════════════════════
    // Generates a list of dates between start and end (inclusive).
    // Optional step parameter for weekly, monthly, etc.

    fnDatesBetween = (
        startDate as date,
        endDate as date,
        optional stepDays as number
    ) as list =>
        let
            step = if stepDays is null then 1 else stepDays,
            dayCount = Duration.Days(endDate - startDate),
            count = Number.RoundDown(dayCount / step) + 1,
            dates = List.Dates(startDate, count, #duration(step, 0, 0, 0))
        in
            dates,

    // ═══════════════════════════════════════════════════════════
    // DEMONSTRATION: Using the functions
    // ═══════════════════════════════════════════════════════════

    // ── Sample data ──────────────────────────────────────────
    SalesData = Table.FromRecords({
        [Product = "Alpha",   Amount = "$1,234.56", Target = 1000],
        [Product = "Beta",    Amount = "€ 850,00",  Target = 0],
        [Product = "Gamma",   Amount = "  2500  ",  Target = 2000],
        [Product = "Delta",   Amount = null,         Target = 500],
        [Product = "Epsilon", Amount = "£3,100.75", Target = 3000]
    }),

    // ── Apply fnCleanCurrency via Table.TransformColumns ─────
    // This is the idiomatic way to apply a function to an entire column.
    // The function signature matches what TransformColumns expects:
    //   each cell value → transformed value
    WithCleanAmount = Table.TransformColumns(
        SalesData,
        {{"Amount", each fnCleanCurrency(_, "."), type number}}
    ),

    // ── Apply fnSafeDivide to compute achievement % ──────────
    WithAchievement = Table.AddColumn(
        WithCleanAmount,
        "Achievement%",
        each fnSafeDivide([Amount], [Target], 0) * 100,
        type number
    ),

    // ── Apply fnDatesBetween (standalone demo) ───────────────
    // Generate every Monday in January 2024
    JanMondays = fnDatesBetween(#date(2024, 1, 1), #date(2024, 1, 31), 7),
    // Result: {#date(2024,1,1), #date(2024,1,8), #date(2024,1,15), ...}

    // ── Daily dates for a week ───────────────────────────────
    OneWeek = fnDatesBetween(#date(2024, 6, 10), #date(2024, 6, 16)),
    // Result: 7 dates from June 10 to June 16

    // ── Final output ─────────────────────────────────────────
    Result = WithAchievement

    // ── RESULT ───────────────────────────────────────────────
    // Product  | Amount  | Target | Achievement%
    // Alpha    | 1234.56 | 1000   | 123.456
    // Beta     | 850     | 0      | 0          ← zero-division protected
    // Gamma    | 2500    | 2000   | 125
    // Delta    | null    | 500    | 0          ← null protected
    // Epsilon  | 3100.75 | 3000   | 103.358...
    //
    // FUNCTION PATTERNS CHEAT SHEET:
    //
    //   // Simple function
    //   fnDouble = (x as number) => x * 2
    //
    //   // With optional parameter
    //   fnGreet = (name as text, optional greeting as text) =>
    //       (if greeting is null then "Hello" else greeting) & ", " & name
    //
    //   // Apply to column (transform in place)
    //   Table.TransformColumns(T, {{"Col", fnDouble, type number}})
    //
    //   // Apply to column (add new column)
    //   Table.AddColumn(T, "New", each fnDouble([Col]), type number)
    //
    //   // Functions as values — store in records for namespacing
    //   Utils = [Clean = fnCleanCurrency, Divide = fnSafeDivide],
    //   Utils[Clean]("$100")   → 100
in
    Result
