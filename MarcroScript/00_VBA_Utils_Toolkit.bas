Attribute VB_Name = "VAB_Utilities_Toolkit"

' ========================= VAB DELTA TOOLKIT =========================
Sub Delta01_ListSheets()
 
    Dim ws As Worksheet
    Dim x As Integer
    Dim sheet As Worksheet
    Dim worksh As Integer
    Dim worksheetexists As Boolean
    worksh = Application.Sheets.Count
    worksheetexists = False
    For x = 1 To worksh
        If Worksheets(x).Name = "Name list" Then
            worksheetexists = True
            'Debug.Print worksheetexists
            Exit For
        End If
    Next x
    If worksheetexists = False Then
        Debug.Print "transformed exists"
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = "Name list"
    End If
    x = 2
 
    Sheets("Name list").Range("A:L").Delete
    For Each ws In Worksheets
         Sheets("Name list").Cells(x, 1) = ws.Name
         x = x + 1
    Next ws
    Sheets("Name list").Activate
    ' Select cell A1
    Range("A1").Select
    ActiveCell.Offset(0, 0).Rows("1:1").EntireRow.Select
    Selection.Font.Bold = True
    
    Range("A1") = "Original sheet's name"
    Range("B1") = "Compare sheet's name"
    Range("C1") = "Delta sheet's name"
    
    Range("A1").Select
    ActiveCell.Columns("A:C").EntireColumn.Select
    Selection.ColumnWidth = 20
   
    Range("A1").Select
    MsgBox "List name created successfully!", vbInformation
End Sub

Public Sub Delta02_CopyTemplateToBlankSheet_M3()
    Dim ws As Worksheet
    Dim nameListSheet As Worksheet
    Dim newName As String
    Dim sheetName As String
    Dim i As Integer

    ' Update screen
    Application.ScreenUpdating = False

    ' Set the Name list sheet
    Set nameListSheet = Worksheets("Name list")

    ' Loop through each name in column A of Name list sheet
    For i = 2 To nameListSheet.Cells(nameListSheet.Rows.Count, 1).End(xlUp).Row
        sheetName = nameListSheet.Cells(i, 1).Value

        ' Check if the sheet exists
        On Error Resume Next
        Set ws = Worksheets(sheetName)
        On Error GoTo 0


        If Not ws Is Nothing Then
            ' Copy the sheet and rename it
            ws.Copy After:=Worksheets(Sheets.Count)
            
            
            '********************************************** please change on  Cutover
    
            If sheetName Like "DeltaM2*" Or sheetName Like "Delta*" Or sheetName Like "M2*" Or sheetName Like "M 2*" Or sheetName Like "m2*" Or sheetName Like "m 2*" Or sheetName Like "Mock2*" Or sheetName Like "Mock 2*" Or sheetName Like "MOCK2*" Or sheetName Like "MOCK 2*" Or sheetName Like "mock2*" Or sheetName Like "mock 2*" Then
                newName = "M3 " & Mid(sheetName, InStr(sheetName, " ") + 1)
            Else
                newName = "M3 " & ws1.Name
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
         
                       
            '**************************************************

            ActiveSheet.Name = newName
            '********** color
            
            ActiveSheet.Tab.ColorIndex = 9
            
            If ActiveSheet.AutoFilterMode Then
                ActiveSheet.AutoFilterMode = False
            End If

            ' Clear contents and delete rows
            Rows("9:" & Rows.Count).ClearContents
            Rows("21:" & Rows.Count).EntireRow.Delete

            ' Select cell A1
            Range("A1").Select
            
            
        End If
        Sheets("Name list").Select
            
        Range("B" & i) = newName
    Next i
    
    Selection.AutoFilter
    
    ' Update screen
    Application.ScreenUpdating = True

    ' Notify user of completion
    MsgBox "Copy to MOCK3 complete.", vbInformation

End Sub

Public Sub Delta03_CopySheetToDelta_M3()
    Dim ws As New Worksheet
    Dim ws1 As New Worksheet
    Dim ws2 As New Worksheet
    Dim ws3 As New Worksheet
    Dim nameListSheet As Worksheet
    Dim newName As String
    Dim sheetName, sheetName2, sheetName3 As String
    Dim i As Integer
    Dim lastRow, LastRow2, LastColumn As Long
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

        If sheetName Like "DeltaM2*" Or sheetName Like "Deltam2*" Or sheetName Like "Delta*" Or sheetName Like "delta*" Then
            newName = "DeltaM3 " & Mid(sheetName, InStr(sheetName, " ") + 1)
        Else
            newName = "DeltaM3 " & ws1.Name
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
            lastRow = Cells(Rows.Count, 7).End(xlUp).Row
            
            
            'copy status form previous mock from column A to B and clear contents in column A
            Range("A9:A" & lastRow).Select
            Selection.Copy
            ActiveCell.Offset(0, 1).Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("A9:A" & lastRow).Select
            Selection.ClearContents
            
            'copy status form Mock number from column D to C and clear contents in column A
            Range("D9:D" & lastRow).Select
            Selection.Copy
            ActiveCell.Offset(0, -1).Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            ' set Value in MOCK cloumn(D)****************************************************
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.FormulaR1C1 = "2" 'change in Next Time
            Selection.Copy
            Range("D9:D" & lastRow).Select
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
                   lastRow = Cells(Rows.Count, 4).End(xlUp).Row
                   
                   Range("G" & lastRow + 1).Select
                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                       :=False, Transpose:=False
                       
                   LastRow3 = Cells(Rows.Count, 7).End(xlUp).Row
                       
                   Rows(lastRow & ":" & lastRow).Select
                   Selection.Copy
                   
                   Rows(lastRow & ":" & LastRow3).Select
                   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                       SkipBlanks:=False, Transpose:=False
                   Application.CutCopyMode = False
                   
                   Range("D" & lastRow + 1).Select
                   ActiveCell.FormulaR1C1 = "3"
                   Selection.Copy
                
                   Range("D" & lastRow + 1 & ":D" & LastRow3).Select
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
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets("Name list").Range("E:L").Delete
    Range("A1").Select
        
        
    Range("E1") = "Original's Records"
    Range("F1") = "Compare' Records"
    Range("G1") = "SUM Redors"
    Range("H1") = "Delta's Records"
    Range("I1") = "Compared Results"
    
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
    ActiveCell.Offset(0, 0).Range("A1:E" & lastRow - 1).Select
    ActiveCell.Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
    ' Update screen
    Application.ScreenUpdating = True

    ' Notify user of completion
    MsgBox "Copy to Delta M3 complete.", vbInformation

End Sub

Public Sub Delta03_CopySheetToDelta_M4_CutOver()
    Dim ws As New Worksheet
    Dim ws1 As New Worksheet
    Dim ws2 As New Worksheet
    Dim ws3 As New Worksheet
    Dim nameListSheet As Worksheet
    Dim newName As String
    Dim sheetName, sheetName2, sheetName3 As String
    Dim i As Integer
    Dim lastRow, LastRow2, LastRow3, LastColumn As Long
    Dim LetterColumn As String
    Dim sheetExists As Boolean
    Dim cellValue As String
                
    ' 1. Turn off screen updates for performance
    Application.ScreenUpdating = False

    ' 2. Point to the "Name list" sheet
    Set nameListSheet = Worksheets("Name list")

    ' 3. Loop through each row in "Name list" (from row 2 down)
    For i = 2 To nameListSheet.Cells(nameListSheet.Rows.Count, 1).End(xlUp).Row
        sheetName  = nameListSheet.Cells(i, 1).Value   ' Source sheet 1
        sheetName2 = nameListSheet.Cells(i, 2).Value   ' Source sheet 2

        ' 4. Attempt to set ws1 and ws2; if sheet doesn't exist, ws1/ws2 stay Nothing
        On Error Resume Next
          Set ws1 = Worksheets(sheetName)
          Set ws2 = Worksheets(sheetName2)
        On Error GoTo 0

        ' todo : Change from 2->3 and 3->4
        ' 5. Build the new sheet name ("DeltaM3 …")
        If sheetName Like "DeltaM3*" Or sheetName Like "Deltam3*" _
           Or sheetName Like "Delta*"  Or sheetName Like "delta*" Then
            newName = "DeltaM4 " & Mid(sheetName, InStr(sheetName, " ") + 1)
        Else
            newName = "DeltaM4 " & ws1.Name
        End If

        ' 6. Check for an existing sheet with that name
        sheetExists = False
        For Each ws In ActiveWorkbook.Sheets
            If InStr(ws.Name, newName) Then
                sheetExists = True
                Exit For
            End If
        Next ws
        If sheetExists Then
            MsgBox "A sheet " & newName & " already exists"
            Exit Sub
        End If

        ' 7. If the first source sheet exists, copy it and rename
        If Not ws1 Is Nothing Then
            ws1.Copy After:=Worksheets(Sheets.Count)
            ActiveSheet.Name = newName
            nameListSheet.Range("C" & i).Value = newName
            sheetName3 = newName

            ' 8. On the new sheet, find last row in column G
            lastRow = Cells(Rows.Count, 7).End(xlUp).Row

            ' 9. Move “old status” from A→B, then clear A9:A…
            With Range("A9:A" & lastRow)
                .Copy
                .Offset(0, 1).PasteSpecial xlPasteValues
                .ClearContents
            End With

            ' 10. Move “mock status” from D→C
            With Range("D9:D" & lastRow)
                .Copy
                .Offset(0, -1).PasteSpecial xlPasteValues
            End With

            ' todo : Change from 2->3
            ' 11. Fill column D with the new mock number “2”
            With Range("D9:D" & lastRow)
                .Cells(1, 1).FormulaR1C1 = "3"
                .Copy
                .PasteSpecial xlPasteValues
            End With
            Range("A1").Select
        End If

        ' 12. If the second source sheet exists, append its data to the new sheet
        If Not ws2 Is Nothing Then
            Set ws2 = Worksheets(sheetName2)
            LastRow2   = ws2.Cells(Rows.Count, 7).End(xlUp).Row
            LastColumn = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column
            LetterColumn = GetColumnLetter(LastColumn)

            ' Copy G9:[LastColumn][LastRow2] from ws2
            ws2.Range("G9:" & LetterColumn & LastRow2).Copy

            ' Paste values into ws3 (the newly created sheetName3)
            Set ws3 = Worksheets(sheetName3)
            With ws3
                Dim pasteRow As Long
                pasteRow = .Cells(.Rows.Count, 4).End(xlUp).Row + 1
                .Cells(pasteRow, "G").PasteSpecial xlPasteValues

                ' Copy formats down the same block
                LastRow3 = .Cells(.Rows.Count, 7).End(xlUp).Row
                .Rows(pasteRow & ":" & pasteRow).Copy
                .Rows(pasteRow & ":" & LastRow3).PasteSpecial xlPasteFormats

                ' todo : Change from 3->4
                ' Fill column D of the new rows with “3”
                .Range("D" & pasteRow).FormulaR1C1 = "4"
                .Range("D" & pasteRow & ":D" & LastRow3).PasteSpecial xlPasteValues

                ' Copy the key from A9 down through the appended rows
                .Range("A9").Copy
                .Range("A9:A" & LastRow3).PasteSpecial

                ' Color the sheet tab light blue
                With .Tab
                    .Color = 10498160
                End With
            End With

            Range("A1").Select
        End If

        ' 13. Call any optional “DeltaRow5” routine
        Delta05optional_DeltaRow5

        ' 14. Mark “To be” under a “Remark” or “Review” header if present
        With ws3
            LastColumn = .Cells(4, .Columns.Count).End(xlToLeft).Column
            cellValue  = .Cells(4, LastColumn).Value
            If LCase(cellValue) = "remark" Or LCase(cellValue) = "review" Then
                .Cells(5, LastColumn).Value = "To be"
                With .Cells(5, LastColumn).Font
                    .Color = -10477568
                End With
            End If
        End With

        ' 15. Adjust columns widths and apply AutoFilter
        With ws3
            .Columns("A:B").ColumnWidth = 7.75
            .Columns("A:G").AutoFit
            .Columns("C:C").ColumnWidth = 4.88
            .Range("A8:" & LetterColumn & (LastRow3 + 8)).AutoFilter
        End With

    Next i

    ' 16. Back on “Name list”: clear old summary, auto-fit, and insert summary formulas
    With Sheets("Name list")
        .Range("E:L").Delete
        .Columns("A:J").AutoFit
        ' Row 2 formulas for counts and comparisons
        .Range("E2").FormulaR1C1 = _
          "=COUNTA(INDIRECT(""'""&RC[-4]&""'!$H:$H""))-COUNTA(INDIRECT(""'""&RC[-4]&""'!$H$1:$H$8""))"
        .Range("F2").FormulaR1C1 = .Range("E2").FormulaR1C1
        .Range("G2").FormulaR1C1 = "=RC[-2]+RC[-1]"
        .Range("H2").FormulaR1C1 = _
          "=COUNTA(INDIRECT(""'""&RC[-5]&""'!$H:$H""))-COUNTA(INDIRECT(""'""&RC[-5]&""'!$H$1:$H$8""))"
        .Range("I2").FormulaR1C1 = "=IF(RC[-2]=RC[-1],TRUE,FALSE)"
        ' Format and conditional highlight of column I
        .Range("A2:D2").NumberFormat = "#,##0"
        With .Range("I2").FormatConditions _
                .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
            .Font.Bold = True
            .Interior.Color = 192
        End With
        ' Copy formulas down for all entries
        Dim lastSummaryRow As Long
        lastSummaryRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("E2:I2").Copy .Range("E2:I" & lastSummaryRow)
    End With

    ' 17. Restore screen updating and notify user
    Application.ScreenUpdating = True
    MsgBox "Copy to Delta M3 complete.", vbInformation
End Sub

Public Sub Delta05_Optional_DeltaRow5()
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

Function GetColumnLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    GetColumnLetter = vArr(0)
End Function

' ========================= VAB FOR LOADING TOOLKIT =========================

Sub Upload01_PrepareToLSMW()
    '
    ' PrepareToLSMW Macro
    ' This macro prepares a file for import into LSMW (Legacy System Migration Workbench)
    '

    Dim xPicRg       As Range
    Dim xPic         As Picture
    Dim xRg          As Range
    Dim ObjectsName  As String
    Dim findCell     As Range
    Dim searchTerms  As Variant
    Dim i            As Integer
    Dim firstColumn  As Integer    ' Column index where "As-Is" is found

    ' ========= Check Invalid Status Column ========
    Dim statusCol    As Long
    Dim lastRow      As Long
    Dim r            As Long
    Dim c            As Range
    Dim invalidFound As Boolean

    statusCol = 0
    For Each c In ActiveSheet.Rows(4).Cells
        If LCase(Trim(c.Value)) = "status" Then
            statusCol = c.Column: Exit For
        End If
    Next c
    If statusCol = 0 Then
        MsgBox "Status column not found in row 4. Please check the header.", vbCritical, "Validation Error"
        Exit Sub
    End If

    On Error Resume Next
    lastRow = ActiveSheet.Cells.Find(What:="*", After:=Cells(1, 1), _
                                     LookIn:=xlFormulas, LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0

    invalidFound = False
    For r = 9 To lastRow
        With ActiveSheet.Cells(r, statusCol)
            If IsEmpty(.Value) Or LCase(Trim(.Value)) = "delete" Then
                .Interior.Color = RGB(255, 204, 204)
                invalidFound = True
            End If
        End With
    Next r

    If invalidFound Then
        MsgBox "Invalid Status values found (blank or 'delete') and highlighted. Please recheck.", _
               vbExclamation, "Validation Error"
        Exit Sub
    End If
    ' ========= end Status check ========

    ' ========= Remove columns up through "NO." ========
    Dim noCell As Range
    Dim noCol  As Long
    noCol = 0
    For Each noCell In ActiveSheet.Rows(4).Cells
        If LCase(Trim(noCell.Value)) = "no." Then
            noCol = noCell.Column: Exit For
        End If
    Next noCell

    If noCol > 0 Then
        ActiveSheet.Range(Cells(1, 1), Cells(1, noCol)).EntireColumn.Delete Shift:=xlToLeft
    Else
        MsgBox "'NO.' column not found in row 4. Skipping NO. removal.", vbExclamation, "Notice"
    End If
    ' ========= end NO. removal ========

    ' === Delete any picture shapes that sit in rows 1–8 ===
    Dim shp As Shape
    Dim topRow As Long
    
    Application.ScreenUpdating = False
    
    For Each shp In ActiveSheet.Shapes
        ' msoPicture = 13
        If shp.Type = msoPicture Then
            ' determine which row its top‐left corner sits in
            topRow = shp.TopLeftCell.Row
            If topRow <= 8 Then
                On Error Resume Next    ' just in case it vanishes under us
                shp.Delete
                On Error GoTo 0
            End If
        End If
    Next shp
    
    Application.ScreenUpdating = True
    ' === end picture removal ========

    ' === Data Cleanup: Remove Unnecessary Rows ===
    ActiveSheet.Rows("1:3").Delete Shift:=xlUp

    ' === Check for specific text (As-Is variants) ===
    searchTerms = Array("As-Is", "as is", "As Is", "ASIS", "AS IS", "asis")
    For i = LBound(searchTerms) To UBound(searchTerms)
        Set findCell = ActiveSheet.Cells.Find( _
                          What:=searchTerms(i), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, MatchCase:=False)
        If Not findCell Is Nothing Then
            firstColumn = findCell.Column
            ActiveSheet.Range(Cells(1, firstColumn), Cells(1, firstColumn + 14)) _
                        .EntireColumn.Delete Shift:=xlToLeft
            Exit For
        End If
    Next i

    ' Delete the next 4 rows after A1
    ActiveSheet.Rows("2:5").Delete Shift:=xlUp

    Application.ScreenUpdating = True
    Range("A1").Select

    MsgBox "Data cleanup and file saving completed.", vbInformation, "Process Completed"
End Sub

Sub Upload02_SaveToTXT()
    Dim ws        As Worksheet
    Dim FilePath  As String
    Dim FileName  As String
    Dim FullName  As String
    Dim lastRow   As Long, lastCol As Long
    Dim r         As Long, c As Long
    Dim lineBuf   As String
    Dim emptyRow  As Boolean
    Dim fNum      As Integer
    Dim lastCell  As Range

    ' === CONFIGURE THIS ===
    FilePath = "C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\Accenture-Data-Migration\MarcroScript\Test\"
    Set ws    = ActiveSheet
    FileName = ws.Name & ".txt"
    FullName = FilePath & FileName
    ' ======================

    ' Ensure target folder exists
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox "Folder not found:" & vbCrLf & FilePath, vbCritical
        Exit Sub
    End If

    ' Find the true last row and column with any data
    With ws
        Set lastCell = .Cells.Find(What:="*", _
                                   LookIn:=xlValues, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlPrevious)
        If Not lastCell Is Nothing Then lastRow = lastCell.Row Else lastRow = 1

        Set lastCell = .Cells.Find(What:="*", _
                                   LookIn:=xlValues, _
                                   SearchOrder:=xlByColumns, _
                                   SearchDirection:=xlPrevious)
        If Not lastCell Is Nothing Then lastCol = lastCell.Column Else lastCol = 1
    End With

    ' Open text file for output (overwrites if exists)
    fNum = FreeFile
    Open FullName For Output As #fNum

    ' Loop through each row, build a tab-delimited line,
    ' skip entirely blank rows, and suppress the final newline.
    For r = 1 To lastRow
        emptyRow = True
        lineBuf = ""

        ' Check for non-blank cells and build the line
        For c = 1 To lastCol
            If Len(Trim(ws.Cells(r, c).Value2 & "")) > 0 Then emptyRow = False
            lineBuf = lineBuf & (ws.Cells(r, c).Value2 & "")
            If c < lastCol Then lineBuf = lineBuf & vbTab
        Next c

        If Not emptyRow Then
            If r = lastRow Then
                ' No trailing CRLF on the very last data row
                Print #fNum, lineBuf;
            Else
                Print #fNum, lineBuf
            End If
        End If
    Next r

    Close #fNum

    MsgBox "Export complete:" & vbCrLf & FullName, vbInformation
End Sub


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


' ========================= VAB FOR UTILITIES TOOLKIT =========================

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

Sub Utils_SplittingDataToBatch()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim headerRow As Range
    Dim newSheet As Worksheet
    Dim batchSize As Long
    Dim totalRows As Long
    Dim startRow As Long, endRow As Long
    Dim batchNum As Long
    Dim userInput As Variant
    Dim baseName As String

    ' Use the active workbook & sheet
    Set wb = ActiveWorkbook
    Set wsSource = wb.ActiveSheet
    baseName = wsSource.Name

    ' Get batch size
    userInput = Application.InputBox("Enter the batch size (rows per sheet):", "Batch Size", Type:=1)
    If userInput = False Or Not IsNumeric(userInput) Or userInput <= 0 Then
        MsgBox "Invalid batch size. Operation cancelled.", vbExclamation
        Exit Sub
    End If
    batchSize = CLng(userInput)

    ' Count data rows (excluding header)
    totalRows = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row - 1
    If totalRows <= 0 Then
        MsgBox "No data to split.", vbExclamation
        Exit Sub
    End If

    Set headerRow = wsSource.Rows(1)
    startRow = 2
    batchNum = 1

    Application.ScreenUpdating = False

    Do While startRow <= totalRows + 1
        endRow = WorksheetFunction.Min(startRow + batchSize - 1, totalRows + 1)

        ' Add new sheet at the end
        Set newSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))

        ' Name it <OriginalSheetName>_batchX
        On Error Resume Next
        newSheet.Name = baseName & "_batch" & batchNum
        If Err.Number <> 0 Then
            Err.Clear
            ' Fallback if name exists: append timestamp
            newSheet.Name = baseName & "_batch" & batchNum & "_" & Format(Now, "hhmmss")
        End If
        On Error GoTo 0

        ' Copy header + data
        headerRow.Copy Destination:=newSheet.Range("A1")
        wsSource.Rows(startRow & ":" & endRow).Copy Destination:=newSheet.Range("A2")

        batchNum = batchNum + 1
        startRow = endRow + 1
    Loop

    Application.ScreenUpdating = True

    MsgBox "Batching complete!" & vbCrLf & _
           "Total rows (excl. header): " & totalRows & vbCrLf & _
           "Total batches: " & (batchNum - 1), vbInformation
End Sub


Sub Utils_SaveAllBatchSheetsToTXT()
    Dim wb           As Workbook
    Dim ws           As Worksheet
    Dim FilePath     As String
    Dim FileName     As String
    Dim FullName     As String
    Dim lastRow      As Long, lastCol As Long
    Dim chunkSize    As Long
    Dim startRow     As Long, endRow As Long
    Dim dataArr      As Variant
    Dim i            As Long, j As Long
    Dim absoluteRow  As Long
    Dim chunkText    As String
    Dim stm          As Object   ' ADODB.Stream
    Dim lastCell     As Range
    Dim emptyRow     As Boolean

    ' === CONFIGURE THIS ===
    FilePath  = "C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\Accenture-Data-Migration\MarcroScript\Test\"
    chunkSize = 10000    ' rows per write-chunk
    ' ======================

    ' Ensure target folder exists
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox "Target folder not found:" & vbCrLf & FilePath, vbCritical
        Exit Sub
    End If

    ' Speed up Excel
    With Application
        .ScreenUpdating = False
        .EnableEvents   = False
        .Calculation    = xlCalculationManual
    End With

    Set wb = ActiveWorkbook

    For Each ws In wb.Worksheets
        ' Only sheets named like "<original>_batchX"
        If ws.Name Like "*_batch[0-9]*" Then

            FileName = ws.Name & ".txt"
            FullName = FilePath & FileName

            ' Find true last row & column with any data
            With ws
                Set lastCell = .Cells.Find(What:="*", _
                                          LookIn:=xlValues, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlPrevious)
                If Not lastCell Is Nothing Then
                    lastRow = lastCell.Row
                Else
                    lastRow = 1
                End If

                Set lastCell = .Cells.Find(What:="*", _
                                          LookIn:=xlValues, _
                                          SearchOrder:=xlByColumns, _
                                          SearchDirection:=xlPrevious)
                If Not lastCell Is Nothing Then
                    lastCol = lastCell.Column
                Else
                    lastCol = 1
                End If
            End With

            ' Initialize UTF-8 stream (with BOM)
            Set stm = CreateObject("ADODB.Stream")
            With stm
                .Type    = 2    ' adTypeText
                .Charset = "utf-8"
                .Open

                ' Write in chunks
                For startRow = 1 To lastRow Step chunkSize
                    endRow = Application.Min(startRow + chunkSize - 1, lastRow)
                    dataArr = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, lastCol)).Value2

                    chunkText = ""
                    For i = 1 To UBound(dataArr, 1)
                        absoluteRow = startRow + i - 1

                        ' Skip entirely blank rows
                        emptyRow = True
                        For j = 1 To UBound(dataArr, 2)
                            If Len(Trim(dataArr(i, j) & "")) > 0 Then
                                emptyRow = False
                                Exit For
                            End If
                        Next j

                        If Not emptyRow Then
                            ' Build tab-delimited line
                            For j = 1 To UBound(dataArr, 2)
                                chunkText = chunkText & (dataArr(i, j) & "")
                                If j < UBound(dataArr, 2) Then chunkText = chunkText & vbTab
                            Next j

                            ' Only add CRLF if this isn't the last data row
                            If absoluteRow <> lastRow Then
                                chunkText = chunkText & vbCrLf
                            End If
                        End If
                    Next i

                    ' Write raw text (no extra newline)
                    .WriteText chunkText, 0   ' adWriteText
                Next startRow

                ' Save & close
                .SaveToFile FullName, 2    ' adSaveCreateOverWrite
                .Close
            End With
        End If
    Next ws

    ' Restore Excel
    With Application
        .Calculation    = xlCalculationAutomatic
        .EnableEvents   = True
        .ScreenUpdating = True
    End With

    MsgBox "All *_batch# sheets exported without trailing blank line to:" & vbCrLf & FilePath, _
           vbInformation, "Export Complete"
End Sub