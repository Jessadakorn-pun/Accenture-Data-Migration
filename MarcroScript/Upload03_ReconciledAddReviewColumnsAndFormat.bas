Sub Upload03_ReconciledAddReviewColumnsAndFormat()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim reviewCol As Range, checkRange As Range
    
    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    
    Set ws = ActiveSheet
    
    ' Find the last row and last column dynamically
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Add "Review", "A", and "B" columns
    Set reviewCol = ws.Cells(1, lastCol + 1)
    reviewCol.Value = "Review"
    
    reviewCol.Offset(0, 1).Value = "Validation and Reconciliation Result - Accuracy and Completeness"
    reviewCol.Offset(0, 2).Value = "Validation and Reconciliation Result - Consistency and Integrity"
    
    ' Apply background color to the last 3 column headers (Review, A, B)
    Dim i As Integer
    For i = 0 To 2
        reviewCol.Offset(0, i).Interior.Color = 11528959
    Next i
    
    ' Insert formulas to check "Passed" or "Failed" status
    reviewCol.Offset(1, 1).FormulaR1C1 = "=IF(ISBLANK(RC[-1]),""Passed"",""Failed"")"
    reviewCol.Offset(1, 2).FormulaR1C1 = "=IF(ISBLANK(RC[-2]),""Passed"",""Failed"")"
    
    ' Autofill the formula dynamically based on the last row
    Set checkRange = ws.Range(reviewCol.Offset(1, 1), reviewCol.Offset(lastRow - 1, 2))
    reviewCol.Offset(1, 1).Resize(1, 2).AutoFill Destination:=checkRange
    
    ' Convert formulas to values (Copy & Paste as Values)
    checkRange.Copy
    checkRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Apply borders to the entire table
    With ws.Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Apply AutoFilter to the entire dataset
    ws.Range("A1").CurrentRegion.AutoFilter
    
    ' Select cell A1 to finalize the macro
    ws.Range("A1").Select
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub