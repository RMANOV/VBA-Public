// ═══════════════════════════════════════════════════════════════
// PATTERN: List.Generate — M's General-Purpose Loop
// CATEGORY: Iteration / Algorithms | DIFFICULTY: Advanced
// ═══════════════════════════════════════════════════════════════
// WHAT: Use List.Generate to implement loops, sequences, and iterations
// WHY:  M has no for/while loops. List.Generate is the universal replacement.
// KEY LESSON: List.Generate is M's loop. State is a record.
//             Selector controls what goes into the output list.
// ═══════════════════════════════════════════════════════════════
//
// List.Generate SIGNATURE:
//
//   List.Generate(
//       initial    as function,   // () => initialState      — starting state
//       condition  as function,   // (state) => true/false   — keep going?
//       next       as function,   // (state) => newState     — advance state
//       optional selector as function  // (state) => output  — what to emit
//   )
//
// HOW IT WORKS:
//   1. Call initial() to get state₀
//   2. Call condition(state₀) — if false, stop; if true, continue
//   3. Call selector(state₀) → emit to output list (or state₀ if no selector)
//   4. Call next(state₀) → state₁
//   5. Call condition(state₁) — if false, stop; if true, continue
//   6. Call selector(state₁) → emit to output list
//   7. Repeat until condition returns false
//
// KEY INSIGHT: State is typically a RECORD with multiple fields.
//   The selector picks which field(s) go into the output list.
//   The next function updates ALL state fields for the next iteration.

let
    // ═══════════════════════════════════════════════════════════
    // EXAMPLE 1: Fibonacci Sequence (first 15 numbers)
    // ═══════════════════════════════════════════════════════════
    // State record tracks: iteration index, current value, next value.
    // The selector emits only the current value.

    Fibonacci = List.Generate(
        // initial: state record with starting values
        () => [i = 0, curr = 0, nxt = 1],

        // condition: keep going while i < 15
        each [i] < 15,

        // next: advance the state — classic Fibonacci shift
        each [
            i    = [i] + 1,
            curr = [nxt],            // current becomes what was next
            nxt  = [curr] + [nxt]    // next becomes sum of both
        ],

        // selector: emit only the current Fibonacci number
        each [curr]
    ),
    // Result: {0, 1, 1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144, 233, 377}

    FibTable = Table.FromList(Fibonacci, Splitter.SplitByNothing(), {"FibValue"}),

    // ═══════════════════════════════════════════════════════════
    // EXAMPLE 2: Powers of 2 until exceeding 1000
    // ═══════════════════════════════════════════════════════════
    // Condition-driven termination (not count-based).
    // Shows how List.Generate naturally handles "while" loops.

    PowersOf2 = List.Generate(
        // initial: start with 2^0 = 1
        () => [power = 0, value = 1],

        // condition: keep going while value <= 1000
        each [value] <= 1000,

        // next: increment power, double the value
        each [
            power = [power] + 1,
            value = [value] * 2
        ],

        // selector: emit the value
        each [value]
    ),
    // Result: {1, 2, 4, 8, 16, 32, 64, 128, 256, 512}

    PowersTable = Table.FromList(PowersOf2, Splitter.SplitByNothing(), {"PowerOf2"}),

    // ═══════════════════════════════════════════════════════════
    // EXAMPLE 3: Every Monday in 2024
    // ═══════════════════════════════════════════════════════════
    // Practical business use case: generate a list of weekly dates.
    // Find the first Monday, then step by 7 days.

    FirstDayOf2024 = #date(2024, 1, 1),
    // Day.Monday = 1 in Date.DayOfWeek (0-based, Sunday = 0)
    // Find offset to first Monday
    DaysUntilMonday = Number.Mod(8 - Date.DayOfWeek(FirstDayOf2024, Day.Sunday), 7),
    FirstMonday = Date.AddDays(FirstDayOf2024, DaysUntilMonday),

    Mondays2024 = List.Generate(
        // initial: start from the first Monday
        () => [dt = FirstMonday],

        // condition: stay within 2024
        each Date.Year([dt]) = 2024,

        // next: jump 7 days forward
        each [dt = Date.AddDays([dt], 7)],

        // selector: emit the date
        each [dt]
    ),
    // Result: {#date(2024,1,1), #date(2024,1,8), ..., #date(2024,12,30)}
    // 53 Mondays in 2024

    MondaysTable = Table.FromList(
        Mondays2024,
        Splitter.SplitByNothing(),
        {"Monday"}
    ),
    MondaysTyped = Table.TransformColumnTypes(MondaysTable, {{"Monday", type date}}),

    // Add week number for context
    MondaysWithWeek = Table.AddColumn(
        MondaysTyped,
        "WeekNumber",
        each Date.WeekOfYear([Monday], Day.Monday),
        Int64.Type
    ),

    // ═══════════════════════════════════════════════════════════
    // EXAMPLE 4 (BONUS): Cumulative running total via List.Generate
    // ═══════════════════════════════════════════════════════════
    // Demonstrates List.Generate as a stateful accumulator — something
    // List.Transform cannot do because it has no memory between items.

    SalesValues = {120, 85, 200, 150, 310, 95},

    RunningTotals = List.Generate(
        // initial: index 0, running sum starts at first element
        () => [idx = 0, total = SalesValues{0}],

        // condition: process all elements
        each [idx] < List.Count(SalesValues),

        // next: advance index, accumulate sum
        each [
            idx   = [idx] + 1,
            total = [total] + (
                if [idx] + 1 < List.Count(SalesValues)
                then SalesValues{[idx] + 1}
                else 0
            )
        ],

        // selector: emit running total
        each [total]
    ),
    // Result: {120, 205, 405, 555, 865, 960}

    // ── COMBINE ALL EXAMPLES ─────────────────────────────────
    // Output one of the tables. Change "Result" to explore others:
    //   FibTable, PowersTable, MondaysWithWeek

    Result = MondaysWithWeek

    // ── RESULT (MondaysWithWeek — first 5 rows) ──────────────
    // Monday     | WeekNumber
    // 2024-01-01 | 1
    // 2024-01-08 | 2
    // 2024-01-15 | 3
    // 2024-01-22 | 4
    // 2024-01-29 | 5
    //
    // LIST.GENERATE CHEAT SHEET:
    //
    //   // Count-based loop (like for i = 0 to N)
    //   List.Generate(() => 0, each _ < N, each _ + 1)
    //
    //   // Condition-based loop (like while x < limit)
    //   List.Generate(() => [x=1], each [x] < limit, each [x = f([x])], each [x])
    //
    //   // Without selector: emits the full state record
    //   // With selector: emits only what you choose (cleaner output)
    //
    //   // COMMON MISTAKE: forgetting that "each" in next creates a NEW record.
    //   // Fields in the new record reference the OLD state via [fieldName].
    //   // This is NOT mutation — it's functional state threading.
    //
    //   // PERFORMANCE: List.Generate is lazy (streamed).
    //   // But List.Count or Table.FromList forces full evaluation.
    //   // For very large sequences, consider List.Buffer if accessed multiple times.
in
    Result
