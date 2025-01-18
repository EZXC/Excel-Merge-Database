Function Search(database As Variant, target As String, identifyDB As String)
    
        Select Case identifyDB
            '---------Search PC Database-----------
            Case "PC"
                For i = LBound(database, 1) To UBound(database, 1)
                    If database(i, 5) = target Or database(i, 6) = UCase(target) Or database(i, 7) = target Then
                        Search = i 'skip title row
                        Exit For
                    End If
                Next i
            '---------Search Monitor Database-----------
            Case "Monitor"
                For i = LBound(database, 1) To UBound(database, 1)
                    If database(i, 2) = UCase(target) Or database(i, 4) = target Then
                        Search = i 'skip title row
                        Exit For
                    End If
                Next i
            '---------Search User Database-----------
            Case "User"
                For i = LBound(database, 1) To UBound(database, 1)
                    If database(i, 2) = target Then
                        Search = i 'skip title row
                        Exit For
                    End If
                Next i
            '---------Search SNtoPCName Database-----------
            Case "SNtoPCName"
                For i = LBound(database, 1) To UBound(database, 1)
                    If database(i, 1) = target Then
                        Search = i 'skip title row
                        Exit For
                    End If
                Next i
        End Select
    
End Function