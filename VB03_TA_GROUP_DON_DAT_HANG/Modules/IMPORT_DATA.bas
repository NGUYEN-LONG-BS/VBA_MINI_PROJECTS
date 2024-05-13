Attribute VB_Name = "IMPORT_DATA"
'Option Explicit

Sub SUB_IMPORT_DATA()
        
    Dim connection As ADODB.connection
    Set connection = New ADODB.connection
    
    Dim server_name As String, database_name As String
'    Let server_name = "QUANNGUYEN\SQLEXPRESS"
'    Let database_name = "DONG_NAI_2023"
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
    
    If connection.State = 1 Then
        Debug.Print "Connected!"
    End If

    Dim rng As Range: Set rng = Application.Range("J1:K5")
    Dim row As Range
    For Each row In rng.Rows
        ID = row.Cells(1).Value
        STT = CDbl(row.Cells(2).Value)
'        quantity = CDbl(row.Cells(3).Value)
        Sql = "insert into HA_NOI_DS_DT_DDH values ('" & ID & "'," & STT & ")"
        connection.Execute Sql
    Next row
    connection.Close
End Sub


