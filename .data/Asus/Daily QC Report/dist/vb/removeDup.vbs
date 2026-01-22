Sub RemoveDuplicateDynamic(rng As String, sheetName As String)

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Application.Workbooks(1)

    ' Jika sheetName diberikan (tidak kosong), ambil sheet tersebut
    If sheetName <> "" Then
        Set ws = wb.Worksheets(sheetName)
    Else
        ' Jika sheetName kosong → gunakan sheet yang sedang aktif
        Set ws = wb.ActiveSheet
    End If

    ws.Range(rng).RemoveDuplicates Columns:=1, Header:=xlNo

End Sub
