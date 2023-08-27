/*  Function to combine column values in otherwise identical rows
    Outputs a list by default, but can output text of combined values with [Format = type text]
    [Delimeter = "; "] is optional, default is "; "
    Change output name of aggregated column with [AggregatedColumnNewName = "MyNewName"]
*/
(InputTable as table, AggregatedColumnName as text, optional Options as record) as table =>
let
    Options = Options ?? [],
    AggregatedColumnNewName = Options[AggregatedColumnNewName]? ?? AggregatedColumnName,
    Format = Options[Format]? ?? type list,
    Delimeter = Record.FieldOrDefault(Options, "Delimeter", 
        if Format = type text then "; " else null
        ),

    Aggregation = if Format = type text then
        {{AggregatedColumnNewName, 
            each Text.Combine(Record.Field(_, AggregatedColumnName), Delimeter),
            type nullable text
            }}
        else
        {{AggregatedColumnNewName, 
            each Record.Field(_, AggregatedColumnName),
            type nullable list
            }},

    Output = Table.Group(
        InputTable,
        List.RemoveItems(Table.ColumnNames(InputTable), {AggregatedColumnName}), 
        Aggregation
        )
in
    Output