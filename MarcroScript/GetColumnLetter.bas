Function GetColumnLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    GetColumnLetter = vArr(0)
End Function