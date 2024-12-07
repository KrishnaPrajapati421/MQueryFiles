// Performing operation on a column in Mquery referencing other column without adding a custom column:

let 
Source=Table.AddColumn(),
Newtable=Table.ToRecords(Table.TransformRows(
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

