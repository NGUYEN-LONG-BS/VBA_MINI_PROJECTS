Attribute VB_Name = "Module1"
Private Sub CommandButton_Import_Click()
    
    Dim con As ADODB.connection
    Set con = New ADODB.connection
    con.Open "Driver- (SQL Server); Server-.; Database-mydemo2; Uid-sa; Pwd=123456;"
    
    Dim rng As Range: Set rng = Application.Range("F6:H9")
    Dim row As Range
    For Each row In rng.Rows
        Name = row.Cells(1).Value
        Price = CDbl(row.Cells(2).Value)
        quantity = CDbl(row.Cells(3).Value)
        Sql = "insert into product values ('" & Name & "'," & Price & "," & quantity & ")"
        con.Execute Sql
    Next row
    
    con.Close
    MsgBox "Done"
End Sub

Sub Connection_String()

    Dim connection As ADODB.connection
    Set connection = New ADODB.connection

    With connection
        .ConnectionString = "Provider=SQLNCLI11;Server=" & server_name & _
            ";database=" & database_name & ";Integrated Security=SSPI;"
        .ConnectionTimeout = 10
        .Open
    End With

    If connection.State = 1 Then
        Debug.Print "Connected!"
    Else: Debug.Print "Not Connected!"
    End If

End Sub


