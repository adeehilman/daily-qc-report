Sub SplitColumnDynamic(rng As String)


    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Application.Workbooks(1)
    Set ws = wb.Worksheets(1)    ' → atau pakai nama: wb.Worksheets("Sheet1")

    ws.Range(rng).TextToColumns _
        Destination:=ws.Range("S2"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Other:=True, _
        OtherChar:="/"

End Sub
