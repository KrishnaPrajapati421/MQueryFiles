// Performing operation on a column in Mquery referencing other column without adding a custom column:

let 
Source=Table.AddColumn(),
Newtable=Table.ToRecords(Table.TransformRows(
	Source,
	each _&[ColumnName= Expression]
	))
in 
@Newtable
