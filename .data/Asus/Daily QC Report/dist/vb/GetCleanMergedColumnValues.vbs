Function GetCleanMergedColumnArrayMultiExclude(arg As Variant) As Variant

    Dim parts() As String
    parts = Split(CStr(arg), "|")

    Dim sheetName As String
    Dim col As String
    Dim excludeRaw As String

    sheetName = parts(0)
    col = parts(1)
    If UBound(parts) >= 2 Then excludeRaw = parts(2) Else excludeRaw = ""

    Dim excludes() As String
    If excludeRaw <> "" Then
        excludes = Split(excludeRaw, ";")
    Else
        ReDim excludes(-1) ' empty array
    End If

    Dim ws As Worksheet
    Set ws = Sheets(sheetName)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Dim r As Long
    Dim c As Range
    Dim val As String
    Dim i As Long
    Dim skip As Boolean

    For r = 1 To lastRow
        Set c = ws.Cells(r, col)

        If c.MergeCells Then
            If c.Address <> c.MergeArea.Cells(1, 1).Address Then GoTo NextRow
            val = CStr(c.MergeArea.Cells(1, 1).Value)
        Else
            val = CStr(c.Value)
        End If

        If Trim(val) = "" Then GoTo NextRow

        skip = False
        For i = LBound(excludes) To UBound(excludes)
            If excludes(i) <> "" Then
                If InStr(1, val, excludes(i), vbTextCompare) > 0 Then
                    skip = True
                    Exit For
                End If
            End If
        Next i

        If skip Then GoTo NextRow

        If Not dict.Exists(val) Then dict.Add val, val

NextRow:
    Next r

    ' convert dictionary keys to String Array
    Dim result() As String
    Dim idx As Long

    ReDim result(0 To dict.Count - 1)
    For idx = 0 To dict.Count - 1
        result(idx) = CStr(dict.Keys()(idx))
    Next idx

    GetCleanMergedColumnArrayMultiExclude = result

End Function
