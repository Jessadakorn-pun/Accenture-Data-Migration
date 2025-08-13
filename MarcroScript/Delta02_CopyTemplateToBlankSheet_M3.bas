Public Sub Delta02_CopyTemplateToBlankSheet_CUTOVER()
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
    
            If sheetName Like "DeltaM3*" Or sheetName Like "Delta*" Or sheetName Like "M3*" Or sheetName Like "M 3*" Or sheetName Like "m3*" Or sheetName Like "m 3*" Or sheetName Like "Mock3*" Or sheetName Like "Mock 3*" Or sheetName Like "MOCK3*" Or sheetName Like "MOCK 3*" Or sheetName Like "mock3*" Or sheetName Like "mock 3*" Then
                newName = "M4 " & Mid(sheetName, InStr(sheetName, " ") + 1)
            Else
                newName = "M4 " & ws1.Name
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