// Performing operation on a column in Mquery referencing other column without adding a custom column:

let 
Source=Table.AddColumn(),
Newtable=Table.Records(Table.TransformRows(
	Source,
	Each _&[ColumnName= Expression]
	))
in 
@Newtable
