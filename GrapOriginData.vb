Function GrapOriginData(worksheetName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstRow As Integer, lastColNum As Integer, i As Integer, checkFirstRow As Integer, j As Integer, snColumn As Integer
    Dim sRange As Range, checkSNCol As Range

    Set ws = ThisWorkbook.Sheets(worksheetName)

    '-----------Ifentify first row & last column-------------------------------
    
    firstRow = IdentifyFirstRow(ws)
    lastColNum = IdentifyLastColumn(ws)
    snColumn = IdentifyPrimaryKeyColumn(ws)
    lastRow = IdentifyLastRow(ws)
    
    Set sRange = ws.Range(ws.Cells(firstRow + 1, 1), ws.Cells(lastRow, lastColNum))
    GrapOriginData = sRange

End Function