Function GrapWorkbook(tempFileLink as String, tempWorksheetName as String, maxColumn As Integer)

	Dim wb as Workbook
	Dim ws as Worksheet
	Dim data as variant
	Set wb = workbooks.open(tempFileLink)
	set ws = wb.sheets(tempWorksheetName)
	data = GrapData2(ws)
	data = ReformatDatabase(data, maxColumn)
	GrapWorkbook = data
	wb.close SaveChange:=False

end Function