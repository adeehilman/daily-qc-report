Sub RemoveDuplicateDynamic(arg As String)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim parts() As String
    parts = Split(arg, "|")

    Dim rng As String
    Dim sheetName As String

    rng = parts(0)
    sheetName = ""

    If UBound(parts) >= 1 Then
        sheetName = parts(1)
    End If

    Set wb = Application.Workbooks(1)

    If sheetName <> "" Then
        Set ws = wb.Worksheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If

    ws.Range(rng).RemoveDuplicates Columns:=1, Header:=xlNo

End Sub
