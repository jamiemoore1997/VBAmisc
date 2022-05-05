Attribute VB_Name = "Module1"
Sub rep_update_file()

Dim mail As Objec, app As Object, mailcontent As String, recipient As String, _
subtitle As String
recipient = ""
Do While recipient = ""
    recipient = InputBox("Name of recipient:", "Recipient")
    If recipient <> "" Then
        subtitle = InputBox("Subject here:", "Subject")
    Else
        MsgBox "Please enter recipient name!", vbCritical
    End If
Loop

Set app = CreatObject("Outlook.Application")
Set mail = app.CreateItem(0)

mailcontent = "<BODY style= font-size:12pt,font-family: Calibri>" & _
            "Hi,<br>" & _
            "Please check the attached file for the updates <br>" & _
            "Best Regards, <br> Kat<br>"
                     
On Error Resume Next
    With mail
        .to = recipient
        .cc = ""
        .bcc = ""
        .Subject = Subject
        .htmlbody = mailcontent
        .attachment.Add.ThisWorkbook.Path
    End With
    On Error GoTo 0
Set mail = Nothing

End Sub
    
