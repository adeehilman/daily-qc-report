Function NextColumnLetter(colLetter As String) As String
    Dim colNum As Long
    
    ' Convert letter → number
    colNum = Range(colLetter & "1").Column
    
    ' Tambah 1
    colNum = colNum + 1
    
    ' Convert number → letter
    NextColumnLetter = Split(Cells(1, colNum).Address(False, False), "1")(0)
End Function
