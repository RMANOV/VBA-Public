// ═══════════════════════════════════════════════════════════════
// PATTERN: Custom Sorting with Sort Key Injection
// CATEGORY: Data Transformation | DIFFICULTY: Beginner-Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Sort rows in a non-alphabetical custom order by adding
//       a numeric sort key, sorting on it, then removing it.
// WHY:  Alphabetical sort puts "Critical" before "High" and
//       "Low" before "Medium" — completely wrong for priorities.
// KEY LESSON: M's Table.Sort only does alphabetical/numerical.
//             For custom order: add sort key -> sort -> remove key.
// ═══════════════════════════════════════════════════════════════

let
    // ── Sample data: tasks with priority ────────────────────────
    Tasks = Table.FromRecords({
        [Task = "Deploy hotfix",       Priority = "Critical", DueDate = #date(2025, 3, 10)],
        [Task = "Write unit tests",    Priority = "Low",      DueDate = #date(2025, 4, 15)],
        [Task = "Review PR #421",      Priority = "Medium",   DueDate = #date(2025, 3, 18)],
        [Task = "Fix login bug",       Priority = "High",     DueDate = #date(2025, 3, 12)],
        [Task = "Update docs",         Priority = "Low",      DueDate = #date(2025, 5,  1)],
        [Task = "Patch CVE-2025-001",  Priority = "Critical", DueDate = #date(2025, 3,  8)],
        [Task = "Refactor auth module",Priority = "Medium",   DueDate = #date(2025, 3, 25)],
        [Task = "Add rate limiting",   Priority = "High",     DueDate = #date(2025, 3, 20)]
    }),

    // ═══════════════════════════════════════════════════════════
    // THE PROBLEM: Alphabetical sort is wrong
    // ═══════════════════════════════════════════════════════════
    // Table.Sort(Tasks, {"Priority", Order.Ascending}) gives:
    //   Critical → High → Low → Medium
    // We need:
    //   Critical → High → Medium → Low

    // ── Define the custom order ─────────────────────────────────
    // Position in this list = sort rank (0 = first, 3 = last)
    PriorityOrder = {"Critical", "High", "Medium", "Low"},

    // ═══════════════════════════════════════════════════════════
    // STEP 1: Add numeric sort key
    // ═══════════════════════════════════════════════════════════
    // List.PositionOf returns the 0-based index of the value.
    // "Critical" → 0, "High" → 1, "Medium" → 2, "Low" → 3
    // If value not found, returns -1 (sorts to top — fail-safe).
    WithSortKey = Table.AddColumn(
        Tasks,
        "PrioritySortKey",
        each List.PositionOf(PriorityOrder, [Priority]),
        Int64.Type
    ),

    // ═══════════════════════════════════════════════════════════
    // STEP 2: Sort on the numeric key
    // ═══════════════════════════════════════════════════════════
    Sorted = Table.Sort(WithSortKey, {
        {"PrioritySortKey", Order.Ascending}
    }),

    // ═══════════════════════════════════════════════════════════
    // STEP 3: Remove the helper column — clean output
    // ═══════════════════════════════════════════════════════════
    CleanResult = Table.RemoveColumns(Sorted, {"PrioritySortKey"}),

    // ═══════════════════════════════════════════════════════════
    // BONUS 1: Multi-level sort
    // ═══════════════════════════════════════════════════════════
    // Sort by Priority (custom), then by DueDate within each
    // priority group (earliest first).
    MultiLevelSorted = Table.Sort(WithSortKey, {
        {"PrioritySortKey", Order.Ascending},
        {"DueDate",         Order.Ascending}
    }),

    MultiLevelClean = Table.RemoveColumns(MultiLevelSorted, {"PrioritySortKey"}),

    // ═══════════════════════════════════════════════════════════
    // BONUS 2: Day-of-week sorting (Monday first)
    // ═══════════════════════════════════════════════════════════
    // Alphabetical gives: Friday, Monday, Saturday, Sunday...
    // We need: Monday, Tuesday, Wednesday, Thursday, Friday,
    //          Saturday, Sunday

    WeekdaySample = Table.FromRecords({
        [Day = "Wednesday", Hours = 8.0],
        [Day = "Monday",    Hours = 7.5],
        [Day = "Friday",    Hours = 6.0],
        [Day = "Sunday",    Hours = 0.0],
        [Day = "Tuesday",   Hours = 8.5],
        [Day = "Thursday",  Hours = 9.0],
        [Day = "Saturday",  Hours = 3.0]
    }),

    // Monday = 0, Sunday = 6  (ISO week order)
    DayOrder = {"Monday", "Tuesday", "Wednesday", "Thursday",
                "Friday", "Saturday", "Sunday"},

    WeekdayWithKey = Table.AddColumn(
        WeekdaySample,
        "DaySortKey",
        each List.PositionOf(DayOrder, [Day]),
        Int64.Type
    ),

    WeekdaySorted = Table.Sort(WeekdayWithKey,
        {{"DaySortKey", Order.Ascending}}
    ),

    WeekdayClean = Table.RemoveColumns(WeekdaySorted, {"DaySortKey"}),

    // ═══════════════════════════════════════════════════════════
    // REUSABLE FUNCTION: CustomSort
    // ═══════════════════════════════════════════════════════════
    // Encapsulates the 3-step pattern into a single call.
    // Usage: CustomSort(table, "Priority", PriorityOrder)
    CustomSort = (
        tbl       as table,
        colName   as text,
        sortOrder as list
    ) as table =>
        let
            withKey = Table.AddColumn(tbl, "__SortKey__",
                each List.PositionOf(sortOrder, Record.Field(_, colName)),
                Int64.Type),
            sorted  = Table.Sort(withKey, {"__SortKey__", Order.Ascending}),
            cleaned = Table.RemoveColumns(sorted, {"__SortKey__"})
        in
            cleaned,

    // Example: sort tasks using the reusable function
    FunctionDemo = CustomSort(Tasks, "Priority", PriorityOrder),

    // ── Choose which result to display ──────────────────────────
    // Change "Result" to see different outputs:
    //   CleanResult      → single-level priority sort
    //   MultiLevelClean  → priority + date sort
    //   WeekdayClean     → day-of-week sort
    //   FunctionDemo     → reusable function demo
    Result = MultiLevelClean

in
    Result