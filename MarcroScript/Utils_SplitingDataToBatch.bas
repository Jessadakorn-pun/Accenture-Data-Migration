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
