/*
GroupingColumns list of text

Options record of any:

    HashFunctions Record of Functions
        - Value should be a function that returns a number from 0 to n - 1 where n is the number of unique values in GroupingColumn
        - Key should be GroupingColumn name to apply the function to
        - Default is grouping column values sorted a-z

        - Whether to add summary functions can be controlled with OutlineLevel (TODO implement)
    SummaryFunctions Record of Functions

    HeadingStyle Text
        "Outline" (Default): 
            - Places group headings in additional columns to the left. For example, with two grouping columns, Outline level 1 headings will be in column A, Outline level 2 headings will be in column B, and the source data will be in columns C+.
        "Inline": 
            - Places group headings in the first column of the source data.

    MaxLevelBlankRows Number
        - Maximum outline level for adding blank rows
        - Default is all outline levels

    BlankGroupName Text
        - Title to use for category rows that are blank (null)
        - Default is "Blank"

*/

(InputTable as table, GroupingColumns as list, optional Options as record, optional OutlineLevel as number) as table =>
let
    // Set up key variables; create and count a Sorted Set of Categories
    OutlineLevel = OutlineLevel ?? 1,
    InputTable = if (OutlineLevel = 1) then Table.AddIndexColumn(InputTable, "_ID", 1) else InputTable,
    NumCategoryCols = List.Count(GroupingColumns),
    GroupingColumn = GroupingColumns{OutlineLevel - 1},
    CategorySet = List.Sort(List.Distinct(Table.Column(InputTable, GroupingColumn))), 
    NumGroups = List.Count(CategorySet),

    // Options
    Options = Options ?? [],
    HashFunctions = Record.FieldOrDefault(Options, "HashFunctions", null),
    SummaryFunctions = Record.FieldOrDefault(Options, "SummaryFunctions", null),
    HeadingStyle = Record.FieldOrDefault(Options, "HeadingStyle", "Outline"),
    MaxLevelBlankRows = Record.FieldOrDefault(Options, "MaxLevelBlankRows", NumCategoryCols),
    BlankGroupName = Record.FieldOrDefault(Options, "BlankGroupName", "Blank"),

    // Hash Function (Default is sorted a-z ascending)
    MatchTable = Record.FromList(
        List.Numbers(0, NumGroups), 
        List.ReplaceValue(CategorySet, null, BlankGroupName, Replacer.ReplaceValue)
        ),
    DefaultHashFunction = (key) => Record.Field(MatchTable, key ?? BlankGroupName),
    HashFunction = Record.FieldOrDefault(HashFunctions, GroupingColumn, DefaultHashFunction),

    // Heading title function
    FormatHeading = (xText) => Text.Upper(xText ?? BlankGroupName),

    // Create empty record ingredients
    FirstRecord = InputTable{0},
    EmptyValues = List.Repeat({null}, Record.FieldCount(FirstRecord)),
    EmptyKeys = Record.FieldNames(FirstRecord),
    EmptyRecord = Record.FromList(EmptyValues, EmptyKeys),

    // Set up heading row function for groups
    HeadingCols = Record.FromList(List.Repeat({null}, NumCategoryCols), List.Transform(GroupingColumns, each "H_" & _)),
    HeadingName = "H_" & GroupingColumn,
    GroupHeaders = (xTable, FirstID) => Table.Combine({
        // add a blank row
        let
            AddHeadingCols = if (HeadingStyle = "Outline") then HeadingCols else [],
            AddBlank = (OutlineLevel <= MaxLevelBlankRows) and (Record.Field(xTable{0}, "_ID") <> FirstID)
        in
            if (AddBlank) then Table.FromRecords({Record.Combine({AddHeadingCols, EmptyRecord})}) else #table({}, {}),
        // add heading row
        let
            // Second record replaces value from <<HeadingName>> field of the first record
            HeadingRowOutline = Table.FromRecords({Record.Combine({
                HeadingCols, 
                Record.FromList({FormatHeading(Record.Field(xTable{0}, GroupingColumn))}, {HeadingName}),
                EmptyRecord
                })}),
            // Second record replaces value from first field of the first record
            HeadingRowInline = Table.FromRecords({Record.Combine({
                EmptyRecord, 
                Record.FromList({FormatHeading(Record.Field(xTable{0}, GroupingColumn))}, {EmptyKeys{0}})
                })})
        in
            if (HeadingStyle = "Outline") then HeadingRowOutline else HeadingRowInline
    }),

    // Set up summary row function for groups
    SummaryRow = (xTable) => #table({}, {}), // placeholder

    // Check if last group
    LastGroup = OutlineLevel = NumCategoryCols,

    // Partition and combine with headers
    #"Partitioned Table" = Table.Partition(InputTable, GroupingColumn, NumGroups, HashFunction),
    FirstID = #"Partitioned Table"{0}{0}[_ID],
    #"Added Headers" = List.Transform(#"Partitioned Table", each Table.Combine({
        GroupHeaders(_, FirstID),
        if (LastGroup) then _ else @GroupByCategory(_, GroupingColumns, Options, OutlineLevel + 1),
        if (SummaryFunctions <> null) then SummaryRow(_) else #table({}, {})
    })),
    #"Recombined Tables" = Table.Combine(#"Added Headers")
in
    #"Recombined Tables"