Public Sub Delta03_CopySheetToDelta_ACO_CUTOVER()
    Dim ws As New Worksheet
    Dim ws1 As New Worksheet
    Dim ws2 As New Worksheet
    Dim ws3 As New Worksheet
    Dim nameListSheet As Worksheet
    Dim newName As String
    Dim sheetName, sheetName2, sheetName3 As String
    Dim i As Integer
    Dim LastRow, LastRow2, LastColumn As Long
    Dim LetterColumn As String
    Dim sheetExists As Boolean
    Dim cellValue As String
                
    ' Update screen
    Application.ScreenUpdating = False

    ' Set the Name list sheet
    Set nameListSheet = Worksheets("Name list")

    ' Loop through each name in column A of Name list sheet
    For i = 2 To nameListSheet.Cells(nameListSheet.Rows.Count, 1).End(xlUp).Row
        sheetName = nameListSheet.Cells(i, 1).Value
        sheetName2 = nameListSheet.Cells(i, 2).Value



        ' Check if the sheet exists
        On Error Resume Next
        Set ws = ActiveSheet
        Set ws1 = Worksheets(sheetName)
        Set ws2 = Worksheets(sheetName2)
        On Error GoTo 0

'********************************************** please change on Delta MOCK3

        If sheetName Like "DeltaM3*" Or sheetName Like "Deltam3*" Or sheetName Like "Delta*" Or sheetName Like "delta*" Then
            newName = "DeltaACO " & Mid(sheetName, InStr(sheetName, " ") + 1)
        Else
            newName = "DeltaACO " & ws1.Name
        End If

        sheetExists = False
    
   
         ' Check if the sheet name already exists
        For Each ws In ActiveWorkbook.Sheets
            If InStr(ws.Name, newName) Then
                sheetExists = True
                Exit For
            End If
        Next ws
        
        ' If the sheet name already exists
        If sheetExists Then
            MsgBox "A sheet " & newName & " already exists"
            Exit Sub
        End If


        If Not ws1 Is Nothing Then
            ' Copy the sheet and rename it
            ws1.Copy After:=Worksheets(Sheets.Count)
                    
           ActiveSheet.Name = newName
            
            nameListSheet.Range("C" & i).Value = ActiveSheet.Name
            sheetName3 = nameListSheet.Cells(i, 3).Value
                     
            
            ' Check last row
            LastRow = Cells(Rows.Count, 7).End(xlUp).Row
            
            
            'copy status form previous mock from column A to B and clear contents in column A
            Range("A9:A" & LastRow).Select
            Selection.Copy
            ActiveCell.Offset(0, 1).Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("A9:A" & LastRow).Select
            Selection.ClearContents
            
            'copy status form Mock number from column D to C and clear contents in column A
            Range("D9:D" & LastRow).Select
            Selection.Copy
            ActiveCell.Offset(0, -1).Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            ' set Value in MOCK cloumn(D)****************************************************
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.FormulaR1C1 = "3" 'change in Next Time
            Selection.Copy
            Range("D9:D" & LastRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            Range("A1").Select
            

                
            If Not ws2 Is Nothing Then
                ' Copy the sheet and rename it
                ws2.Select
                LastRow2 = Cells(Rows.Count, 7).End(xlUp).Row
                
                LastColumn = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column
                
                LetterColumn = GetColumnLetter(LastColumn)
                
                Range("G9:" & LetterColumn & LastRow2).Select
                Selection.Copy
                On Error Resume Next
                Set ws3 = Worksheets(sheetName3)
                On Error GoTo 0
                If Not ws3 Is Nothing Then
                    ws3.Select
                   LastRow = Cells(Rows.Count, 4).End(xlUp).Row
                   
                   Range("G" & LastRow + 1).Select
                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                       :=False, Transpose:=False
                       
                   LastRow3 = Cells(Rows.Count, 7).End(xlUp).Row
                       
                   Rows(LastRow & ":" & LastRow).Select
                   Selection.Copy
                   
                   Rows(LastRow & ":" & LastRow3).Select
                   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                       SkipBlanks:=False, Transpose:=False
                   Application.CutCopyMode = False
                   
                   Range("D" & LastRow + 1).Select
                   ActiveCell.FormulaR1C1 = "4"
                   Selection.Copy
                
                   Range("D" & LastRow + 1 & ":D" & LastRow3).Select
                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                       :=False, Transpose:=False
                       
                   Range("A9").Select
                   Selection.Copy
                
                   Range("A9:" & "A" & LastRow3).Select
                   ActiveSheet.Paste
                       
                    With ActiveWorkbook.Sheets(sheetName3).Tab
                        .Color = 10498160
                        .TintAndShade = 0
                    End With
                   
                   ' Select cell A1
                   Range("A1").Select
                End If
            End If
        End If


        Delta05optional_DeltaRow5
        
       
        ' Find the last column in row 4
        LastColumn = ws3.Cells(4, ws3.Columns.Count).End(xlToLeft).Column
        
        ' Read the value from the last column in row 4
        cellValue = ws3.Cells(4, LastColumn).Value
        
        ' Check the value, ignoring case sensitivity
        If LCase(cellValue) = "remark" Or LCase(cellValue) = "review" Then
            ' Set the value in row 5 of the same column to "To be"
            ws3.Cells(5, LastColumn).Value = "To be"
            ws3.Cells(5, LastColumn).Select
            With Selection.Font
                .Color = -10477568
                .TintAndShade = 0
            End With
        End If
        
        'add CountConcatKey
        'LastRow3 = Cells(Rows.Count, 7).End(xlUp).Row - 8
        'Range("A1").Select
        'ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.Select
        'Selection.Find(What:="NO.", After:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        'ActiveCell.Columns("A:A").EntireColumn.Select
        'ActiveCell.Offset(1, 0).Range("A1").Activate
        'Selection.Insert Shift:=xlToRight
        'ActiveCell.Offset(2, 0).Range("A1").Select
        'ActiveCell.FormulaR1C1 = "Count Concat Key"
        'ActiveCell.Offset(5, 0).Range("A1").Select
        'Selection.NumberFormat = "General"
        'ActiveCell.FormulaR1C1 = "=COUNTIF(R9C6:R" & (LastRow3 + 8) & "C6,RC[-1])"
        'Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & LastRow3)
        'ActiveCell.Range("A1:A" & LastRow3).Select
        'ActiveCell.Columns("A:A").EntireColumn.Select
        'ActiveCell.Activate
        'With Selection.Interior
            '.Color = 65535
        'End With
        
        Range("A1").Select
        ActiveCell.Offset(0, 0).Columns("A:B").EntireColumn.Select
        Selection.ColumnWidth = 7.75
        ActiveCell.Offset(0, 1).Columns("A:G").EntireColumn.Select
        ActiveCell.Columns("A:G").EntireColumn.EntireColumn.AutoFit
        ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select
        Selection.ColumnWidth = 4.88
        ActiveCell.Offset(7, 0).Range("A1").Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Range("A1").Select
        
    
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        
        LastColumn = ws3.Cells(4, ws3.Columns.Count).End(xlToLeft).Column
        LetterColumn = GetColumnLetter(LastColumn)
        Range("A8:" & LetterColumn & LastRow3 + 8).Select
        
        Selection.AutoFilter
        
        Range("A1").Select

    
    
    Next i
    
   
    Sheets("Name list").Select
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets("Name list").Range("E:L").Delete
    Range("A1").Select
        
        
    Range("E1") = "Original's Records"
    Range("F1") = "Compare' Records"
    Range("G1") = "SUM Redors"
    Range("H1") = "Delta's Records"
    Range("I1") = "Csmpared Results"
    
    Range("A1").Select
    ActiveCell.Offset(0, 0).Columns("A:J").EntireColumn.Select
    ActiveCell.Columns("A:J").EntireColumn.EntireColumn.AutoFit
    Range("A1").Select
    
    
    ActiveCell.Offset(1, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTA(INDIRECT(""'""&RC[-4]&""'!$H:$H""))-COUNTA(INDIRECT(""'""&RC[-4]&""'!$H$1:$H$8""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTA(INDIRECT(""'""&RC[-4]&""'!$H:$H""))-COUNTA(INDIRECT(""'""&RC[-4]&""'!$H$1:$H$8""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTA(INDIRECT(""'""&RC[-5]&""'!$H:$H""))-COUNTA(INDIRECT(""'""&RC[-5]&""'!$H$1:$H$8""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]=RC[-1],TRUE,FALSE)"
    ActiveCell.Offset(0, -4).Range("A1:D1").Select
    Selection.NumberFormat = "#,##0"
    ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=FALSE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveCell.Offset(0, -4).Range("A1:E1").Select
    ActiveCell.Activate
    Selection.Copy
    ActiveCell.Offset(0, 0).Range("A1:E" & LastRow - 1).Select
    ActiveCell.Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
    ' Update screen
    Application.ScreenUpdating = True

    ' Notify user of completion
    MsgBox "Copy to Delta ACO complete.", vbInformation

End Sub