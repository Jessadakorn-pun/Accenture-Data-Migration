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
