let
    VLOOKUP = (lookupValue as text, lookupTable as table, colIndex as number) =>
        let
            result = Table.SelectRows(
                lookupTable, each Record.Field(_, Table.ColumnNames(lookupTable){0}) = lookupValue
            ),
            output =
                if Table.IsEmpty(result) then
                    null
                else
                    Record.Field(Table.First(result), Table.ColumnNames(lookupTable){colIndex - 1})
        in
            output
in
    VLOOKUP




=================================================================	
= VLOOKUP([ColumnToSearch], TableToSearch, ColumnIndexToReturn)

================================================================================================

let
    OptimizedVLOOKUP = (lookupValue as text, lookupTable as table, lookupColumn as text, returnColumn as text, optional defaultValue as any) =>
    let
        result = Table.SelectRows(lookupTable, each Record.Field(_, lookupColumn) = lookupValue),
        output = if Table.IsEmpty(result) then defaultValue else Record.Field(Table.First(result), returnColumn)
    in
        output
in
    OptimizedVLOOKUP

================================================================================================

