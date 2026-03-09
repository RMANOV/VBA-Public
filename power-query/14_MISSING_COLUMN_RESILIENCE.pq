// ═══════════════════════════════════════════════════════════════
// PATTERN: Missing Column Resilience
// CATEGORY: Defensive Programming | DIFFICULTY: Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Survive when source columns are renamed, removed, or reordered.
// WHY:  Production queries break silently when upstream schemas change.
// KEY LESSON: EnsureColumns pattern guarantees required columns exist
//             with safe defaults — no more "column not found" errors.
// ═══════════════════════════════════════════════════════════════

let
    // ── Sample data: Note "Discount" column is MISSING ──────────
    // In production, this table comes from Excel/SQL/API where
    // columns can disappear without warning.
    Source = Table.FromRecords({
        [OrderID = 1001, Product = "Widget A", Quantity = 10, Price = 25.00],
        [OrderID = 1002, Product = "Widget B", Quantity =  5, Price = 40.00],
        [OrderID = 1003, Product = "Widget C", Quantity = 20, Price = 12.50]
    }),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 1: Table.HasColumns — explicit check before access
    // ═══════════════════════════════════════════════════════════
    // Use when you need branching logic (different behavior
    // depending on whether the column exists).
    Pattern1_CheckFirst =
        if Table.HasColumns(Source, "Discount") then
            // Column exists — use it directly
            Table.TransformColumns(Source, {"Discount", each _ / 100})
        else
            // Column missing — add it with a safe default
            Table.AddColumn(Source, "Discount", each 0.00, type number),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 2: MissingField.UseDefault — the PREFERRED approach
    // ═══════════════════════════════════════════════════════════
    // Table.SelectColumns with MissingField.UseDefault silently
    // adds null columns for any missing names. No error, no branch.
    // This is the cleanest single-line defense.
    Pattern2_SelectWithDefault = Table.SelectColumns(
        Source,
        {"OrderID", "Product", "Quantity", "Price", "Discount"},
        MissingField.UseDefault  // Missing columns become null
    ),

    // Replace nulls from the default with actual values
    Pattern2_WithDefaults = Table.ReplaceValue(
        Pattern2_SelectWithDefault,
        null, 0.00,
        Replacer.ReplaceValue,
        {"Discount"}
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 3: Record.FieldOrDefault — row-level safety
    // ═══════════════════════════════════════════════════════════
    // When working with individual records (rows), this returns
    // a default value if the field doesn't exist on that record.
    Pattern3_RowLevel = Table.AddColumn(
        Source,
        "DiscountAmount",
        each
            let
                // Try to read "Discount" from this row; default to 0
                disc = Record.FieldOrDefault(_, "Discount", 0),
                qty  = Record.FieldOrDefault(_, "Quantity", 0),
                prc  = Record.FieldOrDefault(_, "Price", 0)
            in
                qty * prc * disc,
        type number
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 4: EnsureColumns — reusable production function
    // ═══════════════════════════════════════════════════════════
    // This is the DEFINITIVE pattern. Define once, use everywhere.
    // It takes a table and a list of {column, type, default} triples,
    // and guarantees every required column exists.

    // Function definition
    EnsureColumns = (tbl as table, requiredCols as list) as table =>
        // List.Accumulate walks each requirement and adds
        // any column that doesn't already exist on the table.
        List.Accumulate(
            requiredCols,
            tbl,
            (state, col) =>
                let
                    colName    = col{0},  // e.g. "Discount"
                    colType    = col{1},  // e.g. type number
                    colDefault = col{2}   // e.g. 0.00
                in
                    if Table.HasColumns(state, colName) then
                        state  // Column exists — leave it alone
                    else
                        Table.AddColumn(state, colName, each colDefault, colType)
        ),

    // Define the schema contract: {Name, Type, Default}
    RequiredSchema = {
        {"OrderID",   type number, 0},
        {"Product",   type text,   "Unknown"},
        {"Quantity",  type number, 0},
        {"Price",     type number, 0.00},
        {"Discount",  type number, 0.00},   // Missing in source!
        {"Category",  type text,   "N/A"}   // Also missing!
    },

    // Apply — all 6 columns guaranteed to exist
    Pattern4_Ensured = EnsureColumns(Source, RequiredSchema),

    // ── Final calculation using the safe table ──────────────────
    // Now we can confidently reference "Discount" and "Category"
    // without worrying about runtime errors.
    Result = Table.AddColumn(
        Pattern4_Ensured,
        "NetAmount",
        each [Quantity] * [Price] * (1 - [Discount]),
        type number
    )

in
    Result