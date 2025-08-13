Public Sub Delta05optional_DeltaRow5()
            Range("A1").Select
            
            'find As Is
            Range("A1").Select
            ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.Select
            Selection.Find(What:="As-Is", After:=ActiveCell, LookIn:=xlFormulas2, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(0, -1).Range("A1").Select
            Range(Selection, Selection.End(xlToLeft)).Select
            Selection.Copy
            ActiveCell.Offset(-4, 0).Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.Select
            Selection.Find(What:="As-Is", After:=ActiveCell, LookIn:=xlFormulas2, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(0, -1).Range("A1").Select
            Range(Selection, Selection.End(xlToLeft)).Select
            Selection.ClearContents
End Sub