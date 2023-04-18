Sub CountRows()
    Dim rowCount As Integer
    If Range("A1").Value <> "" Then
        rowCount = Range("A1").CurrentRegion.Rows.Count      
        'MsgBox "Count rows: " & rowCount
    End If
End Sub