Sub Upload02_SaveToTXT()
    Dim ws         As Worksheet
    Dim FilePath   As String
    Dim FileName   As String
    Dim FullName   As String
    Dim lastRow    As Long, lastCol As Long
    Dim chunkSize  As Long
    Dim startRow   As Long, endRow As Long
    Dim dataArr    As Variant
    Dim i As Long, j As Long
    Dim absoluteRow As Long
    Dim chunkText  As String
    Dim emptyRow   As Boolean
    Dim lastCell   As Range
    Dim stm        As Object

    ' === CONFIGURE THIS ===
    FilePath  = "C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\Accenture-Data-Migration\MarcroScript\Test\"
    chunkSize = 10000    ' rows per write-chunk
    Set ws    = ActiveSheet
    FileName  = ws.Name & ".txt"
    FullName  = FilePath & FileName
    ' ======================

    ' Ensure target folder exists
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox "Folder not found:" & vbCrLf & FilePath, vbCritical
        Exit Sub
    End If

    ' Find true last row & column with any data
    With ws
        Set lastCell = .Cells.Find(What:="*", LookIn:=xlValues, _
                                   SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not lastCell Is Nothing Then lastRow = lastCell.Row Else lastRow = 1

        Set lastCell = .Cells.Find(What:="*", LookIn:=xlValues, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If Not lastCell Is Nothing Then lastCol = lastCell.Column Else lastCol = 1
    End With

    ' Create ADODB.Stream for UTF-8 + BOM
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type    = adTypeText
        .Charset = "utf-8"
        .Open

        ' Emit BOM
        .WriteText ChrW(&HFEFF), 0

        ' Process rows in chunks
        For startRow = 1 To lastRow Step chunkSize
            endRow = Application.Min(startRow + chunkSize - 1, lastRow)
            dataArr = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, lastCol)).Value2

            chunkText = ""
            For i = 1 To UBound(dataArr, 1)
                absoluteRow = startRow + i - 1
                ' Check if entire row is blank
                emptyRow = True
                For j = 1 To UBound(dataArr, 2)
                    If Len(Trim(dataArr(i, j) & "")) > 0 Then
                        emptyRow = False
                        Exit For
                    End If
                Next j

                If Not emptyRow Then
                    ' Build the line
                    For j = 1 To UBound(dataArr, 2)
                        chunkText = chunkText & (dataArr(i, j) & "")
                        If j < UBound(dataArr, 2) Then chunkText = chunkText & vbTab
                    Next j
                    ' Append CRLF unless this is the very last data row
                    If absoluteRow <> lastRow Then
                        chunkText = chunkText & vbCrLf
                    End If
                End If
            Next i

            ' Write the entire chunk at once
            .WriteText chunkText, 0
        Next startRow

        ' Save and close
        .SaveToFile FullName, adSaveCreateOverWrite
        .Close
    End With
    Set stm = Nothing

    MsgBox "Export complete:" & vbCrLf & FullName, vbInformation
End Sub
