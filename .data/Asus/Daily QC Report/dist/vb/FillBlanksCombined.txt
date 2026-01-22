Sub FillBlanksCombined(combined As Variant)

    Dim parts() As String
    parts = Split(CStr(combined), "|")

    If UBound(parts) < 1 Then
        MsgBox "Invalid argument. Format must be: Range|Sheet"
        Exit Sub
    End If

    Dim rngAddress As String
    Dim sheetName As String

    rngAddress = parts(0)
    sheetName = parts(1)

    On Error Resume Next   'jaga kalau tidak ada blank
    With ThisWorkbook.Sheets(sheetName).Range(rngAddress)
        .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Value = .Value
    End With
    On Error GoTo 0

End Sub
