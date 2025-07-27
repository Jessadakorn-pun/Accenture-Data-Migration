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