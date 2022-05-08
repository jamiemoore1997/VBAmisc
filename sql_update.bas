'Enable ADO 2.0 at Tools/References/ 
Option Explicit
Sub ADOtoSQL()

Dim sSQL As String
Dim ReturnArray
Dim Cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim filepath As String, sconnect As String

filepath = ThisWorkbook.Path

sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & filepath & ";HDR=Yes';"
Cn.Open sconnect
    
    sSQL= "SELECT * From" & Worksheet.Name
    rs.Open sSQL, Cn
        ReturnArray = rs.GetRows
    rs.Close

Cn.Close

End Sub

'Load journal_transaction to new sheet/for CloudStreet
Sub ADOSQL()

    Dim Cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim ServerName As String
    Dim DatabaseName As String
    Dim ID As String
    Dim Pass As String
    Dim SQLStr As String

    ServerName = "cloudstreetreportingsql.database.windows.net"
    DatabaseName = "accounting_cloudstreet" 
    ID = "cloudstreetaccounting"
    Pass= "SuperRead1!" 
    SQLStr = "SELECT * FROM Xero_JournalTransaction_Cash" 

  
    Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"

    rs.Open SQLStr, Cn, adOpenStatic
     
     'Load To Spreadsheet
    With Worksheets("sheet1").Range("a1:z500") 
        .ClearContents
        .CopyFromRecordset rs
    End With
      'Clean
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
End Sub

