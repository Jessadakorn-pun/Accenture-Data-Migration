Sub Upload02_SaveSheetToTXT()
    Dim xFileName1 As String, xFileName2 As String
    Dim rng As Range
    Dim DelimChar As String
    Dim newFileName1 As String, newFileName2 As String
    Dim i As Long, j As Long
    Dim lineText As String
    Dim SavePath As String
    Dim wbName As String

    ' Set the delimiter character between columns
    DelimChar = vbTab ' Set delimiter to Tab

    ' Define the destination folder (Change this path if needed)
    SavePath = "C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\3_LSMW_Load\LSMWtoTXT"  ' Modify this to your desired path

    ' Remove any existing file extensions from workbook name
    wbName = ActiveWorkbook.Name
    If InStr(wbName, ".") > 0 Then wbName = Left(wbName, InStrRev(wbName, ".") - 1)

    ' Generate two filenames
    newFileName1 = wbName & ".txt"                          ' First file: WorkbookName.txt
    newFileName2 = wbName & "_" & ActiveSheet.Name & ".txt" ' Second file: Workbook_Sheet.txt

    xFileName1 = SavePath & newFileName1
    xFileName2 = SavePath & newFileName2

    ' Define the data range to be saved
    Set rng = ActiveSheet.Range("A1").CurrentRegion

    ' ==== First Save (WorkbookName.txt) ====
    Open xFileName1 For Output As #1
    For i = 1 To rng.Rows.Count
        lineText = "" ' Reset value before starting a new line
        For j = 1 To rng.Columns.Count
            lineText = lineText & IIf(j = 1, "", DelimChar) & rng.Cells(i, j).Value
        Next j
        Print #1, lineText
    Next i
    Close #1  ' Close first file

    ' ==== Second Save (Workbook_Sheet.txt) ====
    Open xFileName2 For Output As #2
    For i = 1 To rng.Rows.Count
        lineText = "" ' Reset value before starting a new line
        For j = 1 To rng.Columns.Count
            lineText = lineText & IIf(j = 1, "", DelimChar) & rng.Cells(i, j).Value
        Next j
        Print #2, lineText
    Next i
    Close #2  ' Close second file

    ' Notify the user
    MsgBox "Files saved successfully at:" & vbCrLf & xFileName1 & vbCrLf & xFileName2, vbInformation
End Sub