Function GetLastColumnByRow(sheetRow As String) As String
    Dim parts() As String
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim lastCol As Long
    
    ' Split input: sheet|row
    parts = Split(sheetRow, "|")
    
    Set ws = ThisWorkbook.Worksheets(parts(0))
    targetRow = CLng(parts(1))
    
    ' Cari kolom terakhir yang ada isinya di row tertentu
    lastCol = ws.Cells(targetRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Convert ke huruf kolom
    GetLastColumnByRow = Split(ws.Cells(1, lastCol).Address(False, False), "1")(0)
End Function
