Sub Delta04_AddColumnforDeltaProgram()
'
' Macro3 Macro
'

'
    Range("A1").Select
    ActiveCell.Offset(1, 0).Range("1:2").Select
    Selection.UnMerge
    
    Range("A1").Select
    ActiveCell.Offset(7, 0).Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .IndentLevel = 0
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Selection.UnMerge
    ActiveCell.Offset(0, 1).Columns("A:E").EntireColumn.Select
    ActiveCell.Offset(0, 1).Range("A1").Activate
    Selection.Insert Shift:=xlToRight
    
    Range("A1").Select
    ActiveCell.Offset(3, 0).Range("A1:A4").Select
    Selection.UnMerge

    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Status from"

    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "previous mock"
    
    ActiveCell.Offset(-1, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Reviewer"
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mock"
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Delta Indicator"
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Concatenate Asis Key"
        
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.UnMerge
 
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
 
     ActiveCell.Offset(4, -6).Range("A1").Select
     ActiveCell.FormulaR1C1 = ""
    
     ActiveCell.Offset(0, 3).Range("A1").Select
     ActiveCell.FormulaR1C1 = "Definiton and Design for PTT"

    ActiveCell.Offset(0, -3).Range("A1:G1").Select
     
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16777216
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16777216
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.Offset(-4, 0).Range("A1:G4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16777216
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16777216
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -16777216
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = -16777216
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
     
    Range("A1").Select
    ActiveCell.Offset(0, 1).Columns("A:E").EntireColumn.Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    Range("A1").Select
    
End Sub