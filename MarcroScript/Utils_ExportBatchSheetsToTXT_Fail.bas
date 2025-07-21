Sub ExportBatchSheetsToTXT()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim lineText As String, allText As String
    Dim folderPath As String, fName As String, outPath As String
    Dim fd As FileDialog
    
    ' Use the workbook the user has open
    Set wb = ActiveWorkbook
    
    ' Pick the output folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select folder to save TXT files"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "Operation cancelled.", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    
    ' Ensure trailing backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Application.ScreenUpdating = False
    
    For Each ws In wb.Worksheets
        If LCase(Left(ws.Name, 6)) = "batch_" Then
            ' Find used range
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            ' Build tab-delimited text
            allText = ""
            For r = 1 To lastRow
                lineText = ""
                For c = 1 To lastCol
                    lineText = lineText & ws.Cells(r, c).Text
                    If c < lastCol Then lineText = lineText & vbTab
                Next c
                allText = allText & lineText & vbCrLf
            Next r
            
            ' Prepare filename & path
            fName = ws.Name & ".txt"
            outPath = folderPath & fName
            
            ' Write out UTF-8 with BOM
            On Error GoTo ErrHandler
            With CreateObject("ADODB.Stream")
                .Type = 2             ' adTypeText
                .Charset = "utf-8"
                .Open
                .WriteText allText
                .Position = 0
                .SaveToFile outPath, 2  ' adSaveCreateOverWrite
                .Close
            End With
            On Error GoTo 0
        End If
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Export complete!" & vbCrLf & "Files saved in: " & folderPath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error saving '" & fName & "': " & Err.Description, vbCritical
    Resume Next
End Sub
