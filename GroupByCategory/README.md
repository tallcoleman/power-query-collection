# GroupByCategory Function

## About

This function takes a table and groups it using the specified columns, similar to [grouped records in Airtable](https://support.airtable.com/docs/grouping-records-in-airtable) or a [grouped report in Microsoft Access](https://support.microsoft.com/en-us/office/create-a-grouped-or-summary-report-f23301a1-3e0a-4243-9002-4a23ac0fdbf3).

## Basic Usage

### Adding the function:

New Query/Get Data > From Other Sources > Blank Query > Paste in function and rename query to `GroupByCategory`

### Syntax

`GroupByCategory(InputTable as table, GroupingColumns as list, optional Options as record) as table`

### Simple example:

```pq
let
    #"Source" = Table.FromRecords({
        [OrderID = 1001, Item = "Shampoo", Price = 12],
        [OrderID = 1001, Item = "Loofah", Price = 20],
        [OrderID = 1002, Item = "Shampoo", Price = 12]
    }),
    #"Grouped Records" = GroupByCategory(#"Source", {"OrderID"})
in
    #"Grouped Records"
```

**Output:**

| H_OrderID | OrderID | Item    | Price |
| --------- | ------- | ------- | ----- |
| 1001      |         |         |       |
|           | 1001    | Shampoo | 12    |
|           | 1001    | Loofah  | 20    |
|           |         |         |       |
| 1002      |         |         |       |
|           | 1002    | Shampoo | 12    |


## Options

### Custom Category Sort Order

By default, values in the grouping column are sorted in ascending order. The `GroupOrder` option field can be used to specify a different sort order, by supplying a list of unique values (i.e. a set) to be used for a specific grouping column. The supplied list should be in the desired sort order.

Note that null (blank) values generally get sorted to the top in a-z order. The label for null values is specified by the BlankGroupName option.

**Example:**

```pq
    ...
    OrderIDsDescending = List.Sort(List.Distinct(#"Source"[OrderID]), Order.Descending),
    #"Grouped Records" = GroupByCategory(
        #"Source", 
        {"OrderID"}, 
        [GroupOrder = [OrderID = OrderIDsDescending]]
        )
in
    #"Grouped Records"
```


### Summary Functions

The `SummaryFunctions` option field can be used to specify summary functions that should be applied to specific columns and displayed at the bottom of each group. Functions can be [built in List Functions](https://learn.microsoft.com/en-us/powerquery-m/list-functions) that provide a single output value or they can be a custom function. 

Built in functions have default labels (e.g. `List.Count` will be labeled "Count"), but custom labels can be specified by setting a "Name" metadata value for each function, with the syntax `meta [Name = "CustomName"]`. Labels can be turned off by adding an option field `[SummaryFunctionLabels = false]`.

**Example:**

Adding a total for "Price" and a count for "Item":

```pq
    #"Grouped Records" = GroupByCategory(
        #"Source", 
        {"OrderID"}, 
        [SummaryFunctions = [Price = List.Sum, Item = List.Count]]
        )
in
    #"Grouped Records"
```

**Output:**

| H_OrderID | OrderID | Item    | Price |
| --------- | ------- | ------- | ----- |
|           |         | Count   | Sum   |
| 1001      |         |         |       |
|           | 1001    | Shampoo | 12    |
|           | 1001    | Loofah  | 20    |
|           |         | 2       | 32    |
|           |         |         |       |
| 1002      |         |         |       |
|           | 1002    | Shampoo | 12    |
|           |         | 1       | 12    |


**Example:**

Custom function and custom labels:

```pq
    CountChars = (x as list) => List.Accumulate(x, 0, 
        (state, current) => state + Text.Length(Text.From(current))
        ),
    SummaryFunctions = [
        Price = List.Sum meta [Name = "Total"], 
        Item = CountChars meta [Name = "Characters"]
        ]
    #"Grouped Records" = GroupByCategory(
        #"Source", 
        {"OrderID"}, 
        [SummaryFunctions = SummaryFunctions]
        )
in
    #"Grouped Records"
```


### Output Table Style (Outline or Inline)

The function has two layout styles:

**Outline (Default): `[HeadingStyle = "Outline"]`**

Places group headings in additional columns to the left. For example, with two grouping columns, Outline level 1 headings will be in column B, Outline level 2 headings will be in column C, and the source data will be in columns D+. (Column A is generally reserved for the StyleIndex column.)

**Inline: `[HeadingStyle = "Inline"]`**

Places group headings in the first column of the source data. Example below:

| OrderID | Item    | Price |
| ------- | ------- | ----- |
| 1001    |         |       |
| 1001    | Shampoo | 12    |
| 1001    | Loofah  | 20    |
|         |         |       |
| 1002    |         |       |
| 1002    | Shampoo | 12    |


### Control Blank Row Generation

By default, the function will generate a blank row between all distinct groups of data, including when more than one grouping column is specified (e.g. `#"Grouped Records" = GroupByCategory(#"Source", {"OrderID", "Item"})`).

A maximum outline level for blank row generation can be set using the `MaxLevelBlankRows` option, e.g. `[MaxLevelBlankRows = 1]` to only add blank rows between sections created by the first grouping column.


### Custom Label for Blank Category Rows

By default, category values that are blank (null) in the source data are labeled "Blank". A custom label for this case can be set with `[BlankGroupName = "CustomBlankLabel"]`.


### Hide or Show the StyleIndex Column

By default, the function adds a StyleIndex column as the first column of the output table. The values in this column can be used with conditional formatting rules in Excel to style the output table. The StyleIndex column can be turned off by adding an option field `[ShowStyleIndex = false]`.


### Specifying Multiple Options

To specify multiple options at once, they just need to be combined in one record.

**Example:**

```pq
    ...
    OrderIDsDescending = List.Sort(List.Distinct(#"Source"[OrderID]), Order.Descending),
    #"Grouped Records" = GroupByCategory(
        #"Source", 
        {"OrderID"}, 
        [
            GroupOrder = [OrderID = OrderIDsDescending],
            SummaryFunctions = [Price = List.Sum, Item = List.Count],
            HeadingStyle = "Inline"
        ]
        )
in
    #"Grouped Records"
```


## Future Features to Consider

- [ ] additional options for where summary function explanations go (heading rows, append to value)
- [ ] option to suppress summary functions by outline level