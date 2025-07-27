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
