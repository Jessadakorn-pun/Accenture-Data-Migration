Sub Reconcile_ListSheets()
 
    Dim ws As Worksheet
    Dim x As Integer
    Dim sheet As Worksheet
    Dim worksh As Integer
    Dim worksheetexists As Boolean
    worksh = Application.Sheets.Count
    worksheetexists = False
    For x = 1 To worksh
        If Worksheets(x).Name = "Name list" Then
            worksheetexists = True
            'Debug.Print worksheetexists
            Exit For
        End If
    Next x
    If worksheetexists = False Then
        Debug.Print "transformed exists"
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = "Name list"
    End If
    x = 2
 
    Sheets("Name list").Range("A:L").Delete
    For Each ws In Worksheets
         Sheets("Name list").Cells(x, 1) = ws.Name
         x = x + 1
    Next ws
    Sheets("Name list").Activate
    ' Select cell A1
    Range("A1").Select
    ActiveCell.Offset(0, 0).Rows("1:1").EntireRow.Select
    Selection.Font.Bold = True
    
    Range("A1") = "Error Log sheet's name"
    Range("B1") = "Preload sheet's name"
    
    Range("A1").Select
    ActiveCell.Columns("A:C").EntireColumn.Select
    Selection.ColumnWidth = 20
   
    Range("A1").Select
    MsgBox "List name created successfully!", vbInformation
End Sub