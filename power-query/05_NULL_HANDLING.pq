// ═══════════════════════════════════════════════════════════════
// PATTERN: Null Handling in M
// CATEGORY: Defensive Programming | DIFFICULTY: Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Complete null handling patterns — the traps and the solutions
// WHY:  Equality, ordering, and logical operators handle null differently
// KEY LESSON: null = null is true; relational comparisons can return null.
// ═══════════════════════════════════════════════════════════════

let
    // ═══════════════════════════════════════════════════════════
    // THE DISTINCTION: equality with null differs from ordering
    // ═══════════════════════════════════════════════════════════
    //
    //   null = null      → true
    //   null <> null     → false
    //   null = 5         → false
    //   null > 0         → null
    //   if null then ... → ERROR  (if requires logical, not null)
    //
    // M equality is not SQL null equality: null equals only itself.
    // Ordering and logical operations may propagate null instead.
    // Filtering on [Field] = null selects rows whose value is null.

    // ── STEP 1: Sample data with nulls ───────────────────────
    Source = Table.FromRecords({
        [Name = "Alice",   Revenue = 1000, Discount = 50,   Region = "North"],
        [Name = "Bob",     Revenue = null, Discount = null,  Region = null],
        [Name = "Charlie", Revenue = 800,  Discount = null,  Region = "South"],
        [Name = null,      Revenue = 500,  Discount = 25,    Region = "East"],
        [Name = "Eve",     Revenue = 0,    Discount = 0,     Region = "West"]
    }),

    Typed = Table.TransformColumnTypes(Source, {
        {"Name", type text},
        {"Revenue", type number},
        {"Discount", type number},
        {"Region", type text}
    }),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 1: Null checks — equality or type compatibility
    // ═══════════════════════════════════════════════════════════

    // Both [Revenue] = null and [Revenue] is null select Bob's row.
    // Use <> null or not (... is null) to select non-null values.

    NullRows = Table.SelectRows(Typed, each [Revenue] is null),
    NonNullRows = Table.SelectRows(Typed, each [Revenue] <> null),
    // Equivalent type-compatibility check:
    NonNullRowsSimple = Table.SelectRows(Typed, each not ([Revenue] is null)),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 2: Coalesce function (IFNULL / COALESCE equivalent)
    // ═══════════════════════════════════════════════════════════
    // Returns the first non-null value from a list.
    // Like SQL COALESCE(a, b, c) or Excel's first-non-blank.

    fnCoalesce = (values as list) =>
        let
            cleaned = List.RemoveNulls(values),
            result = if List.IsEmpty(cleaned) then null else List.First(cleaned)
        in
            result,

    // Usage: fnCoalesce({null, null, 42, 99})  → 42
    CoalesceDemo = fnCoalesce({null, null, 42, 99}),   // = 42

    // ═══════════════════════════════════════════════════════════
    // PATTERN 3: Replace nulls in a table column
    // ═══════════════════════════════════════════════════════════
    // Table.ReplaceValue with Replacer.ReplaceValue swaps exact matches.

    // Replace null Revenue with 0
    NullsToZero = Table.ReplaceValue(
        Typed,
        null,                        // old value
        0,                           // new value
        Replacer.ReplaceValue,       // exact match replacer
        {"Revenue", "Discount"}      // target columns (list)
    ),

    // Replace null text with a placeholder
    NullTextFilled = Table.ReplaceValue(
        NullsToZero,
        null,
        "(unknown)",
        Replacer.ReplaceValue,
        {"Name", "Region"}
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 4: Safe arithmetic with null coalescing
    // ═══════════════════════════════════════════════════════════
    // null + 5 = null.  null * 100 = null.  Any math with null propagates null.
    // Always coalesce BEFORE arithmetic.

    WithNetRevenue = Table.AddColumn(
        Typed,
        "NetRevenue",
        each
            let
                rev = if [Revenue] is null then 0 else [Revenue],
                disc = if [Discount] is null then 0 else [Discount]
            in
                rev - disc,
        type number
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 5: Null-safe comparison helper
    // ═══════════════════════════════════════════════════════════
    // When you need to compare two fields that might both be null.

    fnNullSafeEquals = (a, b) =>
        if a is null and b is null then true
        else if a is null or b is null then false
        else a = b,

    // Demo: fnNullSafeEquals(null, null) = true, as with null = null
    SafeCompareDemo = fnNullSafeEquals(null, null),    // = true

    // ═══════════════════════════════════════════════════════════
    // PATTERN 6: Count nulls per column (data quality audit)
    // ═══════════════════════════════════════════════════════════
    AllColumns = Table.ColumnNames(Typed),
    NullCounts = List.Transform(
        AllColumns,
        each {_, List.Count(List.Select(Table.Column(Typed, _), each _ is null))}
    ),
    NullAudit = Table.FromRows(
        NullCounts,
        {"Column", "NullCount"}
    ),

    // ── Final output: the cleaned table ──────────────────────
    Result = NullTextFilled

    // ── RESULT (NullTextFilled) ──────────────────────────────
    // Name      | Revenue | Discount | Region
    // Alice     | 1000    | 50       | North
    // Bob       | 0       | 0        | (unknown)
    // Charlie   | 800     | 0        | South
    // (unknown) | 500     | 25       | East
    // Eve       | 0       | 0        | West
    //
    // ── NullAudit table ──────────────────────────────────────
    // Column   | NullCount
    // Name     | 1
    // Revenue  | 1
    // Discount | 2
    // Region   | 1
    //
    // CHEAT SHEET:
    //   x is null           → true if x is null, otherwise false
    //   x = null            → true if x is null, otherwise false
    //   if x is null then 0 else x       → inline coalesce
    //   Table.ReplaceValue(..., null, 0, Replacer.ReplaceValue, {cols})
    //                       → bulk null replacement
    //   null + anything     → null       (arithmetic propagation)
    //   null and true       → null       (logical propagation)
    //   null or true        → true       (short-circuit: true wins)
in
    Result
