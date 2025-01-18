Function IdentifyPrimaryKeyColumn(tempWorksheet As Worksheet)

    Dim checkKeyCol As Range
    Dim titleRow As Integer, lastColNum As Integer, i As Integer, keyColumn As Integer
	
    titleRow = IdentifyTitleRow(tempWorksheet)
    lastColNum = IdentifyLastColumn(tempWorksheet, titleRow)
    Set checkKeyCol = tempWorksheet.Range(tempWorksheet.Cells(titleRow, 1), tempWorksheet.Cells(titleRow, lastColNum))
    
    For i = 1 To lastColNum
        Select Case checkKeyCol(i)
            Case "SN", "Serial Number", "Serial No", "Monitor SN", "SAMAccountName" '<--------------------------May need change when add new excel
                keyColumn = i
                Exit For
        End Select
    Next
    
    IdentifyPrimaryKeyColumn = keyColumn

End Function