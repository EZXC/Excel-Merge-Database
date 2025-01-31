Function IdentifyTitleRow(tempWorksheet As Worksheet)

	Dim titleRowRange as integer, titleRow as integer, lastColNum As Integer, i as integer
	
	titleRowRange = 5
	titleRow = 0

	For i = 1 To titleRowRange
        lastColNum = IdentifyLastColumn(tempWorksheet, i)
        If titleRow < lastColNum Then
            titleRow = lastColNum
        End If
    Next
    
    IdentifyTitleRow = titleRow    
    
End Function