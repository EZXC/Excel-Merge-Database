Function IdentifyLastRow(tempWorksheet As Worksheet)

    Dim lastRow As Integer, keyColumn As Integer

    keyColumn = IdentifyPrimaryKeyColumn(tempWorksheet)
    
    lastRow = tempWorksheet.Cells(tempWorksheet.Rows.Count, keyColumn).End(xlUp).row
    IdentifyLastRow = lastRow

End Function