Function GrapData(tempWorksheet As Worksheet)

	Dim lastRow As Long
	Dim titleRow As integer, lastColNum as integer, lastRow as integer
	Dim data as variant
	
	titleRow = IdentifyTitleRow(tempWorksheet)
	lastColNum = IdentifyLastColumn(tempWorksheet, titleRow)
	lastRow = IdentifyLastRow(tempWorksheet)
	
	data = tempWorksheet.Range(tempWorksheet.Cells(titleRow, 1), tempWorksheet.Cells(lastRow, lastColNum))
	GrapDat2 = data

End Function