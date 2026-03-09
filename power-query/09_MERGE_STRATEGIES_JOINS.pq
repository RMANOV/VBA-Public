// =====================================================================
// PATTERN: Merge Strategies (All 6 Join Types)
// CATEGORY: Table Joins | DIFFICULTY: Intermediate
// =====================================================================
// WHAT: Demonstrate all 6 JoinKind options with Table.NestedJoin,
//       including proper expansion of the nested result column.
// WHY:  Joining tables is the most common Power Query operation.
//       Choosing the wrong join type silently drops or duplicates rows.
// KEY LESSON: NestedJoin returns a nested table column -- you MUST use
//       Table.ExpandTableColumn to flatten it. LeftAnti finds orphans.
// =====================================================================
//
// EXTENDS: L_R_ANTI_JOINS.pq in repo root (which covers only anti-joins)
//
// JOIN TYPE CHEAT SHEET:
// ---------------------------------------------------------------
// JoinKind.Inner      = Only matching rows from BOTH tables
//   Use when: You need complete records only (e.g., orders WITH customers)
//
// JoinKind.LeftOuter  = ALL left rows + matching right (null if no match)
//   Use when: Keep all left rows, enrich with right where possible
//   This is the DEFAULT and most common join in ETL
//
// JoinKind.RightOuter = ALL right rows + matching left (null if no match)
//   Use when: Keep all right rows (rare -- usually swap tables and use LeftOuter)
//
// JoinKind.FullOuter  = ALL rows from BOTH tables (nulls where no match)
//   Use when: Reconciliation -- find ALL records regardless of matching
//
// JoinKind.LeftAnti   = Left rows with NO match in right
//   Use when: Find orphans, missing references, records to clean up
//
// JoinKind.RightAnti  = Right rows with NO match in left
//   Use when: Find new/unmatched records in the right table
// ---------------------------------------------------------------

let
    // =================================================================
    // SAMPLE DATA
    // =================================================================
    // Orders table: includes CustomerID "C999" with no matching customer
    Orders = Table.FromRecords({
        [OrderID = 1, CustomerID = "C001", Amount = 250],
        [OrderID = 2, CustomerID = "C002", Amount = 175],
        [OrderID = 3, CustomerID = "C001", Amount = 300],
        [OrderID = 4, CustomerID = "C003", Amount = 125],
        [OrderID = 5, CustomerID = "C999", Amount = 50]
    }, type table [OrderID = Int64.Type, CustomerID = text, Amount = Int64.Type]),

    // Customers table: includes "C004" with no orders (orphan customer)
    Customers = Table.FromRecords({
        [CustomerID = "C001", Name = "Alice"],
        [CustomerID = "C002", Name = "Bob"],
        [CustomerID = "C003", Name = "Carol"],
        [CustomerID = "C004", Name = "Dave"]
    }, type table [CustomerID = text, Name = text]),

    // =================================================================
    // 1. INNER JOIN -- Only matching rows
    // =================================================================
    // Result: 4 rows (C001 x2, C002, C003). Drops C999 and C004.
    InnerJoined = Table.NestedJoin(
        Orders, {"CustomerID"},
        Customers, {"CustomerID"},
        "CustomerData",
        JoinKind.Inner
    ),
    InnerExpanded = Table.ExpandTableColumn(
        InnerJoined, "CustomerData", {"Name"}, {"CustomerName"}
    ),

    // =================================================================
    // 2. LEFT OUTER JOIN -- All orders, customer name where available
    // =================================================================
    // Result: 5 rows. C999 order kept, CustomerName = null.
    LeftOuterJoined = Table.NestedJoin(
        Orders, {"CustomerID"},
        Customers, {"CustomerID"},
        "CustomerData",
        JoinKind.LeftOuter
    ),
    LeftOuterExpanded = Table.ExpandTableColumn(
        LeftOuterJoined, "CustomerData", {"Name"}, {"CustomerName"}
    ),

    // =================================================================
    // 3. RIGHT OUTER JOIN -- All customers, order data where available
    // =================================================================
    // Result: 5 rows. Dave (C004) kept with null OrderID and Amount.
    RightOuterJoined = Table.NestedJoin(
        Orders, {"CustomerID"},
        Customers, {"CustomerID"},
        "CustomerData",
        JoinKind.RightOuter
    ),
    RightOuterExpanded = Table.ExpandTableColumn(
        RightOuterJoined, "CustomerData", {"Name"}, {"CustomerName"}
    ),

    // =================================================================
    // 4. FULL OUTER JOIN -- Everything from both sides
    // =================================================================
    // Result: 6 rows. Both C999 (no customer) and C004 (no orders) kept.
    FullOuterJoined = Table.NestedJoin(
        Orders, {"CustomerID"},
        Customers, {"CustomerID"},
        "CustomerData",
        JoinKind.FullOuter
    ),
    FullOuterExpanded = Table.ExpandTableColumn(
        FullOuterJoined, "CustomerData", {"Name"}, {"CustomerName"}
    ),

    // =================================================================
    // 5. LEFT ANTI JOIN -- Orders with NO matching customer
    // =================================================================
    // Result: 1 row (OrderID=5, C999). Finds orphan orders.
    LeftAntiJoined = Table.NestedJoin(
        Orders, {"CustomerID"},
        Customers, {"CustomerID"},
        "CustomerData",
        JoinKind.LeftAnti
    ),
    // Anti-joins return no columns from the other table, but we still
    // need to remove the empty nested column for a clean result
    LeftAntiResult = Table.RemoveColumns(LeftAntiJoined, {"CustomerData"}),

    // =================================================================
    // 6. RIGHT ANTI JOIN -- Customers with NO orders
    // =================================================================
    // Result: 1 row (C004, Dave). Finds customers who never ordered.
    RightAntiJoined = Table.NestedJoin(
        Orders, {"CustomerID"},
        Customers, {"CustomerID"},
        "CustomerData",
        JoinKind.RightAnti
    ),
    RightAntiResult = Table.ExpandTableColumn(
        RightAntiJoined, "CustomerData", {"CustomerID", "Name"},
        {"OrphanCustomerID", "OrphanName"}
    ),

    // =================================================================
    // SUMMARY: Pick any result to inspect
    // =================================================================
    // Change the final step name to explore different joins:
    //   InnerExpanded, LeftOuterExpanded, RightOuterExpanded,
    //   FullOuterExpanded, LeftAntiResult, RightAntiResult

    // Build a summary showing row counts for each join type
    Summary = Table.FromRecords({
        [JoinType = "Inner",      Rows = Table.RowCount(InnerExpanded),
         Description = "Only matching: 4 orders with known customers"],
        [JoinType = "LeftOuter",  Rows = Table.RowCount(LeftOuterExpanded),
         Description = "All orders + customer name (null for C999)"],
        [JoinType = "RightOuter", Rows = Table.RowCount(RightOuterExpanded),
         Description = "All customers + order data (null for Dave)"],
        [JoinType = "FullOuter",  Rows = Table.RowCount(FullOuterExpanded),
         Description = "Everything from both sides (C999 + Dave)"],
        [JoinType = "LeftAnti",   Rows = Table.RowCount(LeftAntiResult),
         Description = "Orphan orders (no matching customer)"],
        [JoinType = "RightAnti",  Rows = Table.RowCount(RightAntiResult),
         Description = "Orphan customers (no orders)"]
    })

in
    Summary

// =====================================================================
// NOTES:
// - Table.NestedJoin ALWAYS returns a nested table column. For Inner,
//   LeftOuter, RightOuter, FullOuter you MUST expand it.
// - For Anti joins, the nested column is empty -- remove or ignore it.
// - Join keys can be multi-column: {"Col1", "Col2"} on both sides.
// - Performance tip: If the right table is small, Power Query hashes it.
//   Put the larger table on the left for better performance.
// - Table.Join (non-nested) also exists but gives less control over
//   column naming and is harder to debug.
// =====================================================================