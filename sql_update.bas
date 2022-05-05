Attribute VB_Name = "Module1"
Option Explicit
Sub ADOtoSQL()

Dim sSQLQry As String
Dim ReturnArray
Dim conn As New ADODB.Connection
Dim mr As New ADODB.Recordset
Dim filepath As String, sconnect As String

filepath = ThisWorkbook.Path

sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & filepath & ";HDR=Yes';"

conn.Open sconnect
    
    sSQLSting = "SELECT * From" & Worksheet.Name
    mr.Open sSQLSting, conn
        ReturnArray = mrs.GetRows
    mr.Close

conn.Close

End Sub

