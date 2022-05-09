Attribute VB_Name = "Module2"
Option Explicit

Sub PDFTemplate()
    Dim folder As FileDialog
    Set folder = Application.FileDialog(msoFileDialogFilePicker)
    With folder
    .Title = "Select file"
    .Filters.Add "PDF Type Files", "*.pdf", 1
    If .Show <> -1 Then GoTo NoSelection
    ActiveSheet.Range("G18").Value = .SelectedItems(1)
    End With
NoSelection:
End Sub

Sub SaveFolder()
    Dim folder As FileDialog
    Set folder = Application.FileDialog(msoFileDialogFolderPicker)
    With folder
    .Title = "Select folder"
    If .Show <> -1 Then GoTo NoSelection
    ActiveSheet.Range("G19").Value = .SelectedItems(1)
    End With
NoSelection:
End Sub

Sub CreateForms()
    Dim PDFTemplate, NewFileName, SaveFolder, LastName As String
    Dim AppDate As Date
    Dim CustRow, LastRow As Long
    With ActiveSheet
    LastRow = Range("E9999").End(xlUp).Row
    PDFTemplate = .Range("G18").Value
    SaveFolder = .Range("G19").Value
    OpenURL "" & PDFFile & "", Show_Maximized
    Application.Wait Now + 0.00005
    
    For CustRow = 5 To LastRow
        LastName = .Range("C" & CustRow).Value
        AppDate = .Range("G" & CustRow).Value
        Application.SendKeys "{Tab}", True
        Application.SendKeys LastName, True
        Application.Wait Now + 0.00001
        
        Application.SendKeys "{Tab}", True
        Application.SendKeys .Range("F" & CustRow).Value, True
        Application.Wait Now + 0.00001
        Application.SendKeys "{Tab}", True
        Application.SendKeys "{Tab}", True
        
        'Contact
        Application.SendKeys "{Tab}", True
        Application.SendKeys Format(.Range("G" & CustRow).Value, "####-###-###"), True
        Application.Wait Now + 0.00001
        Application.SendKeys "{Tab}", True
     
        Application.SendKeys "{Tab}", True
        Application.SendKeys .Range("F" & CustRow).Value, True
        Application.Wait Now + 0.00001
        Application.SendKeys "{Tab}", True
        
        Application.SendKeys "{Tab}", True
        Application.SendKeys .Range("K" & CustRow).Value, True
        Application.Wait Now + 0.00001
        Application.SendKeys "{Tab}", True
        
        Application.SendKeys "^(p)", True
        Application.Wait Now + 0.00002
        Application.SendKeys "{Enter}", True
        Application.Wait Now + 0.00002
        
        Application.SendKeys "%(n)", True
        Application.Wait Now + 0.00002
        Application.SendKeys SaveFolder & "\" & LastName & "_" & Format(AppDate, "DD_MM_YY") & ".pdf"
        Application.Wait Now + 0.00002
        
        Next CustRow
        Application.SendKeys "^(q)", True
        Application.SendKeys "{numlock}%s", True
      End With
End Sub


