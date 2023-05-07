# GroupByCategory Function

This function takes a table and groups it using the specified columns, similar to [grouped records in Airtable](https://support.airtable.com/docs/grouping-records-in-airtable) or a [grouped report in Microsoft Access](https://support.microsoft.com/en-us/office/create-a-grouped-or-summary-report-f23301a1-3e0a-4243-9002-4a23ac0fdbf3).

## Basic Usage

New Query/Get Data > From Other Sources > Blank Query > Paste in function and rename query to `GroupByCategory`

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

### Output

| H_OrderID | OrderID | Item    | Price |
|-----------|---------|---------|-------|
| 1001      |         |         |       |
|           | 1001    | Shampoo | 12    |
|           | 1001    | Loofah  | 20    |
|           |         |         |       |
| 1002      |         |         |       |
|           | 1002    | Shampoo | 12    |

## To Do List

- [ ] additional options for where summary function explanations go (heading rows, append to value)
- [ ] option to suppress summary functions by outline level
- [ ] option to remove grouping columns from data