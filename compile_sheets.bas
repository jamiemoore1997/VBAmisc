Attribute VB_Name = "Module1"

Option Explicit
Sub datalocation()
    Dim startpoint As String
    Range("a1").Select
    Selection.End(xlDown).Select
    startpoint = ActiveCell.Address
    Range(startpoint).Select
End Sub
Sub CompileTables():
    Dim j As Integer
    For j = 1 To Sheets.Count 'starting from the second sheet
        Sheets(j).Activate
        Call datalocation
        Selection.CurrentRegion.Select
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
        Selection.Copy Destination:=Sheets("All").Range("A65536").End(xlUp)(2)
    Next
End Sub
