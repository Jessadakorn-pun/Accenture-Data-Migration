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
