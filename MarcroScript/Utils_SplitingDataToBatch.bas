Sub SplitDataToBatches()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim headerRow As Range
    Dim newSheet As Worksheet
    Dim batchSize As Long
    Dim totalRows As Long
    Dim startRow As Long, endRow As Long
    Dim batchNum As Long, createdBatches As Long
    Dim userInput As Variant

    ' === ensure we use the workbook that the user has active ===
    Set wb = ActiveWorkbook 
    Set wsSource = wb.ActiveSheet

    ' === get batch size from user ===
    userInput = Application.InputBox("Enter the batch size (rows per sheet):", "Batch Size", Type:=1)
    If userInput = False Or Not IsNumeric(userInput) Or userInput <= 0 Then
        MsgBox "Invalid batch size. Operation cancelled.", vbExclamation
        Exit Sub
    End If
    batchSize = CLng(userInput)

    ' === determine how many data-rows we have (excluding header) ===
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

        ' add new sheet at the end
        Set newSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        
        ' try to name it; if that fails (e.g. duplicate), append a time-stamp
        On Error Resume Next
        newSheet.Name = "batch_" & batchNum
        If Err.Number <> 0 Then
            Err.Clear
            newSheet.Name = "batch_" & batchNum & "_" & Format(Now, "hhmmss")
        End If
        On Error GoTo 0

        ' copy header + data block
        headerRow.Copy Destination:=newSheet.Range("A1")
        wsSource.Rows(startRow & ":" & endRow).Copy Destination:=newSheet.Range("A2")

        batchNum = batchNum + 1
        startRow = endRow + 1
    Loop

    Application.ScreenUpdating = True

    createdBatches = batchNum - 1
    MsgBox "Batching complete!" & vbCrLf & _
           "Total rows (excl. header): " & totalRows & vbCrLf & _
           "Total batches: " & createdBatches, vbInformation
End Sub
