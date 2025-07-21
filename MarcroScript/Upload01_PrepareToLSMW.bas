Sub PrepareToLSMW()
    '
    ' PrepareToLSMW Macro
    ' This macro prepares a file for import into LSMW (Legacy System Migration Workbench)
    '

    Dim xPicRg As Range
    Dim xPic As Picture
    Dim xRg As Range
    Dim ObjectsName As String
    Dim findCell As Range
    Dim searchTerms As Variant
    Dim i As Integer
    Dim firstColumn As Integer    ' Column index where "As-Is" is found

    ' ========= Check Invalid Status Column ========
    ' Step 1: Validate "Status" column content from row 9 downward,
    '         using the real last active row, and highlight invalid cells
    Dim statusCol    As Long
    Dim lastRow      As Long
    Dim r            As Long
    Dim c            As Range
    Dim invalidFound As Boolean

    ' 1a) Find the "Status" header in row 4
    statusCol = 0
    For Each c In ActiveSheet.Rows(4).Cells
        If LCase(Trim(c.Value)) = "status" Then
            statusCol = c.Column
            Exit For
        End If
    Next c

    If statusCol = 0 Then
        MsgBox "Status column not found in row 4. Please check the header.", _
               vbCritical, "Validation Error"
        Exit Sub
    End If

    ' 1b) Determine the real last used row on the sheet
    On Error Resume Next
    lastRow = ActiveSheet.Cells.Find(What:="*", _
                                     After:=Cells(1, 1), _
                                     LookIn:=xlFormulas, _
                                     LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious, _
                                     MatchCase:=False).Row
    On Error GoTo 0

    ' 1c) Scan from row 9 through lastRow for empty or "delete"
    invalidFound = False
    For r = 9 To lastRow
        With ActiveSheet.Cells(r, statusCol)
            If IsEmpty(.Value) Or LCase(Trim(.Value)) = "delete" Then
                ' Highlight only the invalid cell with a soft red fill
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
    ' ========= end new function add ========

    ' ========= function remove before and column "NO." ========
    Dim noCell As Range
    Dim noCol  As Long

    ' 1) Find "NO." in header row (row 4)
    noCol = 0
    For Each noCell In ActiveSheet.Rows(4).Cells
        If LCase(Trim(noCell.Value)) = "no." Then
            noCol = noCell.Column
            Exit For
        End If
    Next noCell

    If noCol > 0 Then
        ' 2) Delete columns 1 through the "NO." column (inclusive), then shift left
        ActiveSheet.Range(Cells(1, 1), Cells(1, noCol)).EntireColumn.Delete Shift:=xlToLeft
    Else
        MsgBox "'NO.' column not found in row 4. Skipping NO. removal.", _
               vbExclamation, "Notice"
    End If
    ' ========= end new function add ========

    ' === Delete Pictures in the first 8 rows ===
    Application.ScreenUpdating = False
    Set xRg = ActiveSheet.Range("1:8")
    For Each xPic In ActiveSheet.Pictures
        Set xPicRg = ActiveSheet.Range( _
                         xPic.TopLeftCell.Address & ":" & _
                         xPic.BottomRightCell.Address)
        If Not Intersect(xRg, xPicRg) Is Nothing Then xPic.Delete
    Next xPic
    Application.ScreenUpdating = True

    ' === Data Cleanup: Remove Unnecessary Rows ===
    ActiveSheet.Rows("1:3").Delete Shift:=xlUp

    ' === Check for specific text (As-Is variants) ===
    searchTerms = Array("As-Is", "as is", "As Is", "ASIS", "AS IS", "asis")
    For i = LBound(searchTerms) To UBound(searchTerms)
        Set findCell = ActiveSheet.Cells.Find( _
                          What:=searchTerms(i), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False)
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

    'Call SaveSheetToTXT   ' (optional export to TXT)

    MsgBox "Data cleanup and file saving completed.", _
           vbInformation, "Process Completed"
End Sub