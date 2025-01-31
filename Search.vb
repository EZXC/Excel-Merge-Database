Function Search(tempWorksheet As Worksheet, searchTarget As String, optional targetCol1 as String = "", optional targetCol2 as string = "", optional targetCol3 as string ="")
    
		Dim titleRow as integer, lastColumn as integer, i as integer, j as integer
		Dim targetColumn as Variant
		titleRow = IdentifyTitleRow(tempWorksheet)
		lastColumn = identifyLastColumn(tempWorksheet)
		targetColumn = Array(targetCol1, targetCol2, targetCol3)
	
		for i = 1 to lastColumn
			for each col in targetColumn
				if tempWorksheet(1, i) = col Then
					targetColumn = i
				end if
			next col
		next i
	
		for j = LBound(tempWorksheet, 1) to UBound(tempWorksheet, 1)
			if Ucase(tempWorksheet(i, targetColumn)) = UCase(searchTarget)
				Search = j
				Exit For
			end if
		next j
    
End Function