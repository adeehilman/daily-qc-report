Sub DoublePasteOperation(combined As Variant)

    '=== 1. PARSE ARGUMENTS ==='
    Dim parts() As String
    parts = Split(CStr(combined), "|")

    If UBound(parts) < 3 Then
        MsgBox "Invalid argument format. Expected: Range|Sheet|Op1|Op2"
        Exit Sub
    End If

    Dim rngAddress As String, sheetName As String
    Dim op1 As String, op2 As String

    rngAddress = parts(0)
    sheetName = parts(1)
    op1 = LCase(parts(2))
    op2 = LCase(parts(3))

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(rngAddress)

    '=== 2. BACKUP ORIGINAL VALUES ==='
    Dim backup As Variant
    backup = rng.Value   'snapshot 2D array

    Dim r As Long, c As Long

    '=== 3. OPERASI PERTAMA: pakai backup ==='
    Dim out1 As Variant
    ReDim out1(1 To UBound(backup, 1), 1 To UBound(backup, 2))

    For r = 1 To UBound(backup, 1)
        For c = 1 To UBound(backup, 2)
            If IsNumeric(backup(r, c)) And Not IsEmpty(backup(r, c)) Then
                Select Case op1
                    Case "subtract"
                        '10 - 10 = 0
                        out1(r, c) = CDbl(backup(r, c)) - CDbl(backup(r, c))
                    Case "add"
                        '10 + 10 = 20 (kalau nanti mau)
                        out1(r, c) = CDbl(backup(r, c)) + CDbl(backup(r, c))
                    Case Else
                        out1(r, c) = backup(r, c)
                End Select
            Else
                out1(r, c) = backup(r, c)
            End If
        Next c
    Next r

    'Paste hasil operasi 1
    rng.Value = out1

    '=== 4. OPERASI KEDUA: pakai hasil sekarang + backup ==='
    Dim out2 As Variant
    ReDim out2(1 To UBound(backup, 1), 1 To UBound(backup, 2))

    'baca lagi nilai sekarang (hasil operasi 1)
    Dim currentVal As Variant
    Dim baseVal As Double

    For r = 1 To UBound(backup, 1)
        For c = 1 To UBound(backup, 2)
            currentVal = out1(r, c)      'hasil sebelumnya (misal 0)
            If IsNumeric(backup(r, c)) And Not IsEmpty(backup(r, c)) Then
                baseVal = CDbl(backup(r, c))
                Select Case op2
                    Case "add"
                        '0 + 10 = 10
                        out2(r, c) = CDbl(currentVal) + baseVal
                    Case "subtract"
                        'kalau mau: 0 - 10 = -10
                        out2(r, c) = CDbl(currentVal) - baseVal
                    Case Else
                        out2(r, c) = currentVal
                End Select
            Else
                out2(r, c) = currentVal
            End If
        Next c
    Next r

    'Paste hasil operasi 2
    rng.Value = out2

End Sub
