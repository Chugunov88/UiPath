Sub ConvertAndDivide()
    'Преобразовать ячейки B2:B999 и E2:E999 и G2:G999 в финансовый формат
    Range("B2:B999,E2:E999,G2:G999").NumberFormat = "$#,##0.00"
    ActiveSheet.Cells.VerticalAlignment = xlCenter
    ActiveSheet.Cells.HorizontalAlignment = xlCenter
    Dim headerRange As Range
    Set headerRange = Range("A1").CurrentRegion.Rows(1)
    headerRange.EntireColumn.AutoFit


    'Деление значений в столбцах B и E и запись результатов в столбец G
    For i = 2 To 999
        If Range("A" & i).Value <> "" Then
            If IsNumeric(Range("E" & i).Value) And Range("E" & i).Value <> 0 Then
                If IsNumeric(Range("B" & i).Value) Then
                    Range("G" & i).Value = Range("B" & i).Value / Range("E" & i).Value
                Else
                    Range("G" & i).Value = "-"
                End If
            Else
                Range("G" & i).Value = "-"
            End If
        End If

        'Умножение значений в столбцах B и E на 1, если столбец G не пустой
        If Range("G" & i).Value <> "" And Range("G" & i).Value <> "-" Then
            If IsNumeric(Range("B" & i).Value) Then
                Range("B" & i).Value = Range("B" & i).Value * 1
            End If
            If IsNumeric(Range("E" & i).Value) Then
                Range("E" & i).Value = Range("E" & i).Value * 1
            End If
        End If
    Next i
End Sub
