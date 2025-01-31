Function IdentifyLastColumn(tempWorksheet As Worksheet, targetRow as Integer)

	Dim lastColumn as Integer
	lastColumn = tempWorksheet.Cells(targetRow, Columns.Count).End(xlToLeft).Column
    IdentifyLastColumn = lastColumn

End Function