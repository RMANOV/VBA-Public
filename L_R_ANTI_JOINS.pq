let
    CompareTables = (Table1 as table, Table2 as table, Columns as list, Message1 as text, Message2 as text) as table =>
    let
        // Default messages
        DefaultMessage1 = if Message1 = null then "Присъства в таблица 1, но не в таблица 2" else Message1,
        DefaultMessage2 = if Message2 = null then "Присъства в таблица 2, но не в таблица 1" else Message2,

        // Perform Left Anti Join
        LeftAntiJoin = Table.NestedJoin(Table1, Columns, Table2, Columns, "Table2", JoinKind.LeftAnti),
        LeftAntiExpanded = Table.ExpandTableColumn(LeftAntiJoin, "Table2", {}, {"Table2"}),
        LeftAntiResult = Table.AddColumn(LeftAntiExpanded, "Message", each DefaultMessage1),

        // Perform Right Anti Join
        RightAntiJoin = Table.NestedJoin(Table2, Columns, Table1, Columns, "Table1", JoinKind.RightAnti),
        RightAntiExpanded = Table.ExpandTableColumn(RightAntiJoin, "Table1", {}, {"Table1"}),
        RightAntiResult = Table.AddColumn(RightAntiExpanded, "Message", each DefaultMessage2),

        // Combine Results
        CombinedResults = Table.Combine({LeftAntiResult, RightAntiResult}),

        // Filter Results

        FilteredResults = if Table.FilterWithDataTable = "Table1" then
                            Table.SelectRows(CombinedResults, each [Message] = DefaultMessage1)
                          else if Table.FilterWithDataTable = "Table2" then
                            Table.SelectRows(CombinedResults, each [Message] = DefaultMessage2)
                          else
                            CombinedResults
    in
        FilteredResults
in
    CompareTables
