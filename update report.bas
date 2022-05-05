

Sub UpdateReport()

Dim FillRange As Range
Dim sheetname As Worksheet
Dim destinationsheet As Worksheet

'Alerts and Screen update false

Application.DisplayAlerts = False
Application.ScreenUpdating = False


'Confirmation

    Answer = MsgBox("Do you have the latest report version open?", vbQuestion + vbYesNo, "Confirm")
    If Answer = vbNo Then
        Exit Sub
    Else

'Select data
    Application.Goto Reference:=Worksheets("Transactions").Range("A1")
    
    Range("A2").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row - 3, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
'Copy data
    Selection.Copy
    Set sheetname = ActiveWorkbook.Worksheets.Add
    sheetname.Name = "Data Extract"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("CD1").Select
    Application.CutCopyMode = False
    
'Fill data
    ActiveCell.Formula = _
        "=MID(CELL(""filename"",A1),FIND(""["",CELL(""filename"",A1))+1,10)"
    Range("CD1").Select
    Last = Lastrow(sheetname)
    Selection.AutoFill Destination:=Range(Cells(1, "CD"), Cells(Last, "CD"))

'Copy to merge
    Range("A1").Select
    Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
    Selection.Copy
'Open file and create new sheet

    Workbooks.Open Filename:= _
        "C:\Users\jamie\Downloads\Match P&L"
    Set sheetname = ActiveWorkbook.Worksheets("Updated Reports")
    Last = Lastrow(sheetname) + 1
       
    Range(Cells(Last, "A"), Cells(Last, "A")).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Set destinationsheet = ActiveWorkbook.Worksheets("Updated Reports")
Last = Lastrow(destinationsheet)

'Sorted
    Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select

'Reset
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

