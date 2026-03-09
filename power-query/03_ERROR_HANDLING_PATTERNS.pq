// ═══════════════════════════════════════════════════════════════
// PATTERN: Error Handling Patterns
// CATEGORY: Defensive Programming | DIFFICULTY: Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Comprehensive try...otherwise patterns for bulletproof transforms
// WHY:  One bad row shouldn't kill your entire query refresh
// KEY LESSON: try returns a record [HasError, Value, Error], not the value itself
// ═══════════════════════════════════════════════════════════════

let
    // ── STEP 1: Sample data with messy values ────────────────
    // Real-world data is never clean. Text in number fields,
    // empty strings, nulls — all must be handled gracefully.
    Source = Table.FromRecords({
        [ID = 1, Amount = "1500",    Category = "Sales"],
        [ID = 2, Amount = "abc",     Category = "Sales"],     // text, not a number
        [ID = 3, Amount = "",        Category = "Returns"],   // empty string
        [ID = 4, Amount = "2300.50", Category = "Sales"],
        [ID = 5, Amount = null,      Category = "Returns"],   // null value
        [ID = 6, Amount = "17,99",   Category = "Sales"],     // comma decimal (EU format)
        [ID = 7, Amount = "-300",    Category = "Returns"]
    }),

    // ══════════════════════════════════════════════════════════
    // PATTERN 1: Simple fallback with try...otherwise
    // ══════════════════════════════════════════════════════════
    // The simplest form: try an operation, fall back on failure.
    // Good for quick-and-dirty cleaning. You lose error details.

    Pattern1 = Table.AddColumn(
        Source,
        "AmountClean",
        each try Number.From([Amount]) otherwise null,
        type number
    ),

    // ══════════════════════════════════════════════════════════
    // PATTERN 2: Full error inspection with try (no otherwise)
    // ══════════════════════════════════════════════════════════
    // Without "otherwise", try returns a RECORD:
    //   [HasError = true/false, Value = ..., Error = [Reason, Message, Detail]]
    //
    // This lets you inspect what went wrong and WHY.

    Pattern2 = Table.AddColumn(
        Source,
        "Attempt",
        each try Number.From([Amount])
        // No "otherwise" — returns the full attempt record
    ),

    // Extract the fields from the attempt record
    WithHasError = Table.AddColumn(
        Pattern2,
        "HasError",
        each [Attempt][HasError],
        type logical
    ),
    WithValue = Table.AddColumn(
        WithHasError,
        "ParsedValue",
        each if [Attempt][HasError] then null else [Attempt][Value],
        type number
    ),
    WithErrorMsg = Table.AddColumn(
        WithValue,
        "ErrorMessage",
        each if [Attempt][HasError]
             then [Attempt][Error][Message]
             else null,
        type text
    ),

    // Clean up: remove the raw attempt record column
    Pattern2Clean = Table.RemoveColumns(WithErrorMsg, {"Attempt"}),

    // ══════════════════════════════════════════════════════════
    // PATTERN 3: Row-level error log with categorization
    // ══════════════════════════════════════════════════════════
    // Production pattern: parse the value, log errors with context,
    // and categorize error types for downstream reporting.

    Pattern3 = Table.AddColumn(
        Source,
        "ParseResult",
        each
            let
                raw = [Amount],
                attempt = try Number.From(raw)
            in
                if raw = null then
                    [Value = null, Status = "NULL_INPUT", Detail = "Field was null"]
                else if raw = "" then
                    [Value = null, Status = "EMPTY_STRING", Detail = "Field was empty"]
                else if not attempt[HasError] then
                    [Value = attempt[Value], Status = "OK", Detail = null]
                else
                    [Value = null, Status = "PARSE_ERROR",
                     Detail = attempt[Error][Message]]
    ),

    // Expand the nested record into separate columns
    Expanded = Table.ExpandRecordColumn(
        Pattern3,
        "ParseResult",
        {"Value", "Status", "Detail"},
        {"ParsedAmount", "Status", "ErrorDetail"}
    ),

    // ── BONUS: Filter to only error rows for review ──────────
    ErrorLog = Table.SelectRows(
        Expanded,
        each [Status] <> "OK"
    ),

    // ── Final result: the full table with error annotations ──
    Result = Expanded

    // ── RESULT (Expanded) ────────────────────────────────────
    // ID | Amount  | Category | ParsedAmount | Status       | ErrorDetail
    // 1  | 1500    | Sales    | 1500         | OK           | null
    // 2  | abc     | Sales    | null         | PARSE_ERROR  | "We couldn't..."
    // 3  | (empty) | Returns  | null         | EMPTY_STRING | "Field was empty"
    // 4  | 2300.50 | Sales    | 2300.5       | OK           | null
    // 5  | null    | Returns  | null         | NULL_INPUT   | "Field was null"
    // 6  | 17,99   | Sales    | 1799*        | OK           | null  ← locale-dependent!
    // 7  | -300    | Returns  | -300         | OK           | null
    //
    // *Note: Number.From("17,99") behavior depends on your locale settings.
    //  In en-US it reads as 1799; in bg-BG/de-DE it reads as 17.99.
    //  Use Number.FromText with culture parameter for explicit control.
    //
    // CHEAT SHEET:
    //   try X otherwise Y     → returns Y on error (simple, loses details)
    //   try X                 → returns [HasError, Value, Error] (full inspection)
    //   [Error][Reason]       → error category (e.g., "Expression.Error")
    //   [Error][Message]      → human-readable error description
    //   [Error][Detail]       → additional context (sometimes null)
in
    Result
