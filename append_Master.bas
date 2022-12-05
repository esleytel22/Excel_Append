Attribute VB_Name = "Modul1"
Public FileCounter As Long
Public FileNameArray
Public NewWorkbook As String
Public Sht As Worksheet

Sub main()
Dim i As Long
    NewSheet
    GetFileNames
    For i = 1 To FileCounter
        Workbooks.Open Filename:=FileNameArray(i)
        NewWorkbook = ActiveWorkbook.Name
        ProcessFile
        Workbooks(NewWorkbook).Close SaveChanges:=False
    Next
End Sub


Sub GetFileNames()
    FileNameArray = Application.GetOpenFilename(, , , , True)
    FileCounter = UBound(FileNameArray)
End Sub

Sub ProcessFile()
    Dim DestRow As Long, RowCount As Long
    
    RowCount = ActiveSheet.Range("A1").CurrentRegion.Rows.Count
    DestRow = Sht.Range("A" & Rows.Count).End(xlUp).Row + 1
    If DestRow + RowCount > 5000 Then
        MsgBox ("This sheet is full. New sheet will be added.")
        NewSheet
        DestRow = 1
    End If
    Workbooks(NewWorkbook).Sheets(1).Range("A1").CurrentRegion.Copy Destination:=Sht.Cells(DestRow, 1)
End Sub


Sub NewSheet()
    ThisWorkbook.Activate
    ThisWorkbook.Sheets.Add
    Set Sht = ActiveSheet
End Sub

