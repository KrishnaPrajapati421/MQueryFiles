// Each _ in Table context refers to the record of the table
let 
Source = Excel.CurrentWorkbook(),
customColumn = Table.AddColumn(Source, "New Column", each Record.Field(_, "ColumnName"))
in 
customColumn

// Each _ in List context refers to the element of the list
let 
Source = {1,2,3,4,5,6},
customColumn = List.Transform(Source, each _ * 2)
in
customColumn

// Performing operation on a column in Mquery referencing other column without adding a custom column:

let 
Source=Table.AddColumn(),
Newtable=Table.FromRecords(Table.TransformRows(
	Source,
	each _&[ColumnName= Expression]
	))
in 
Newtable


// Performing append table operation with dynamic column names

let 
Source=Excel.CurrentWorkbook(),
ColumnNameList=Table.ColumnNames(Table.Combine(Source[Content])),
Custom1=Source,
ExpandTables=Table.ExpandTableColumns(Custom1,"Content",ColumnNameList)
in 
ExpandTables

// Dynamic Unpivoting 

let 
Source=Excel.CurrentWorkbook(),
ColumnNameList=Table.ColumnNames(Source),
ConvertTotable=Table.FromList(ColumnNameList,Splitter.SplitByNothing(),null,null,ExtraValues.Ignore)// For Converting List to table
	
let 
Source=Excel.CurrentWorkbook(),
// When 'each _' is used in the context of the table it means the Record of the table
CustomColumn= Table.AddColumn(Source,"New Column", each _)


// Performing total on the columns in the table skipping the first column
let 
Source = Excel.CurrentWorkbook(),
customColumn = Table.AddColumn(Source, "New Column", each List.Sum(List.Skip(Record.ToList(_))))
in 
customColumn

// Unpivot multiple columns

let
    Source = Excel.CurrentWorkbook(){[Name="Table12"]}[Content],
    Tables = List.Transform(List.Split(Table.ToColumns(Table.Skip(Source,1)),3), each Table.PromoteHeaders(Table.FromColumns(_))),
    FirstRow = Table.FromList(List.RemoveNulls(Record.ToList(Source{0})), null, {"Country"}),
    TWIndex = Table.AddIndexColumn(FirstRow, "Index", 0, 1, Int64.Type),
    Custom2 = Table.TransformColumns(
        TWIndex,{"Index", each Tables{_}}),
    #"Expanded Index" = Table.ExpandTableColumn(Custom2, "Index", Table.ColumnNames(Tables{0}))
in
    #"Expanded Index"