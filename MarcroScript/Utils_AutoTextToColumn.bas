Sub Utils_AutoTextToColumn()
    Dim ws As Worksheet
    Dim lastCol As Long, c As Long
    Dim fldInfo() As Variant
    
    ' Reference the active sheet
    Set ws = ActiveSheet
    
    ' Find last used column in row 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Build a FieldInfo array so that EVERY output column is Text.
    ' We'll assume no more than 50 splits per column—adjust 1 To 50 as needed.
    ReDim fldInfo(1 To 50, 1 To 2)
    For c = 1 To 50
        fldInfo(c, 1) = c            ' output column index
        fldInfo(c, 2) = xlTextFormat ' treat as Text
    Next c
    
    Application.ScreenUpdating = False
    
    ' Loop through each column and re-parse it
    For c = 1 To lastCol
        With ws.Columns(c)
            .TextToColumns _
                Destination:=.Cells(1, 1), _
                DataType:=xlDelimited, _
                TextQualifier:=xlTextQualifierNone, _
                ConsecutiveDelimiter:=False, _
                Tab:=True, _
                FieldInfo:=fldInfo, _
                TrailingMinusNumbers:=True
        End With
    Next c
    
    Application.ScreenUpdating = True
    MsgBox "All columns have been re-parsed as Tab-delimited Text.", vbInformation
End Sub
