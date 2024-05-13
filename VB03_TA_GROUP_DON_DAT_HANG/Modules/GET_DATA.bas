Attribute VB_Name = "GET_DATA"
Option Explicit

Sub SUB_GET_DATA()

Dim connection As ADODB.connection
Set connection = New ADODB.connection

Dim server_name As String, database_name As String
'Let server_name = "(LocalDb)\LocalDbTest"
'Let database_name = "AdventureWorks2016"
'Let server_name = "BANKSNB5\SQLEXPRESS"
'Let database_name = "TBD_2023"
Let server_name = "SRV1\TUAN_AN_GROUP"
Let database_name = "HA_NOI_2023"

With connection
'    .ConnectionString = "Provider=SQLNCLI11;Server=" & server_name & _
'        ";database=" & database_name & ";Integrated Security=SSPI;"
    .ConnectionString = "Provider=SQLNCLI11;Server=" & server_name & _
        ";database=" & database_name & ";User Id=sa; Password=Ta#9999;"
    'SQLOLEDB.1
    .ConnectionTimeout = 10
    .Open
End With

'With connection
''    .ConnectionString = "Provider=SQLNCLI11;Server=" & server_name & _
''        ";database=" & database_name & ";Integrated Security=SSPI;"
'    .ConnectionString = "Data Source=172.16.0.128,1433;Network Library=DBMSSOCN;Initial Catalog=" & database_name & ";User ID=sa;Password=Ta#9999;"
'    'SQLOLEDB.1
'    .ConnectionTimeout = 10
'    .Open
'End With

If connection.State = 1 Then
    Debug.Print "Connected!"
Else
    Debug.Print "Not Connected!"
End If

Dim sqlQuery As String
sqlQuery = "Select * from [HA_NOI_2023].[dbo].[HA_NOI_DS_DT_DDH]"

Dim rsSql As New ADODB.Recordset
rsSql.CursorLocation = adUseClient
rsSql.Open sqlQuery, connection, adOpenStatic

ThisWorkbook.Sheets(1).Range("A2").CopyFromRecordset rsSql

End Sub



