Attribute VB_Name = "Module1"
Option Explicit

Sub open_file_chooser()
    Dim fd As Office.FileDialog
Dim strFile As String
 
Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
    With fd
 
       .Filters.Clear
       .Filters.Add "Excel Files", "*.xlsx?", 1
       .Title = "Choose an Excel file"
       .AllowMultiSelect = False
    
       .InitialFileName = ThisWorkbook.path
    
       If .Show = True Then
    
           strFile = .SelectedItems(1)
    
       End If
 
    End With
    MenuSheet.Range("B1").Value = getFileNameFromPath(strFile)
End Sub

Function getFileNameFromPath(path)
    Dim fileName As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    getFileNameFromPath = fso.GetFilename(path)
End Function

Sub open_form()
    ufSQL.Show vbModeless
End Sub

Sub CreateSQLQuery(SQLQuery As String)

    Dim num As Long
    num = MenuSheet.Range("B2").Value
    GetQueryResults MenuSheet.Range("B1").Value, SQLQuery
    
    LogSheet.Range("A" & (num + 1)).Value = SQLQuery
    
End Sub

Sub GetQueryResults(fileName As String, SQLQuery As String)

    Dim MovieFilePath As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim i As Integer
    Dim RowCount As Long, ColCount As Long
    
    'Exit the procedure if no query was passed in
    If SQLQuery = "" Then
        'MsgBox _
            Prompt:="You didn't enter a query", _
            Buttons:=vbCritical, _
            Title:="Query string missing"
            CreateObject("WScript.Shell").Popup "B" & ChrW(7841) & "n ch" & ChrW(432) & "a nh" & ChrW(7853) & "p c" & ChrW(226) & "u l" & ChrW(7879) & "nh SQL n" & ChrW(224) & "o", , "C" & ChrW(226) & "u truy v" & ChrW(7845) & "n r" & ChrW(7895) & "ng", 0 + 32
        Exit Sub
    End If
    
    'Check that the Movies workbook exists in the same folder as this workbook
    MovieFilePath = ThisWorkbook.path & "\" & fileName
    
    If Dir(MovieFilePath) = "" Then
        'MsgBox _
            Prompt:="Could not find Movies.xlsx", _
            Buttons:=vbCritical, _
            Title:="File not found"
        CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y file c" & ChrW(417) & " s" & ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u " & fileName, , "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y file", 0 + 32
        Exit Sub
    End If
    
    'Create and open a connection to the Movies workbook
    Set cn = New ADODB.Connection
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & MovieFilePath & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    
    'Try to open the connection, exit the subroutine if this fails
    On Error GoTo EndPoint
    cn.Open
    
    'If anything fails after this point, close the connection before exiting
    On Error GoTo CloseConnection
    
    'Create and populate the recordset using the SQLQuery
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorType = adOpenStatic
    
    rs.Source = SQLQuery    'Use the query string that we passed into the procedure
    
    'Try to open the recordset to return the results of the query
    rs.Open
    
    'If anything fails after this point, close the recordset and connection before exiting
    On Error GoTo CloseRecordset
    
    'Get count of rows returned by the query
    RowCount = rs.RecordCount
    
    'Exit the procedure if no rows returned
    If RowCount = 0 Then
        'MsgBox _
            Prompt:="The query returned no results", _
            Buttons:=vbExclamation, _
            Title:="No Results"
        CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng c" & ChrW(243) & " k" & ChrW(7871) & "t qu" & ChrW(7843) & " n" & ChrW(224) & "o " & ChrW(273) & ChrW(432) & ChrW(7907) & "c tr" & ChrW(7843) & " v" & ChrW(7873) & " cho c" & ChrW(226) & "u truy v" & ChrW(7845) & "n c" & ChrW(7911) & "a b" & ChrW(7841) & "n", , "Kh" & ChrW(244) & "ng c" & ChrW(243) & " k" & ChrW(7871) & "t qu" & ChrW(7843) & " n" & ChrW(224) & "o " & ChrW(273) & ChrW(432) & ChrW(7907) & "c tr" & ChrW(7843) & " v" & ChrW(7873), 0 + 64
        Exit Sub
    End If
    
    'Get the count of columns returned by the query
    ColCount = rs.Fields.Count
    
    'Create a new worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    'Select the worksheet to avoid the formatting bug with CopyFromRecordset
    ThisWorkbook.Activate
    ws.Select
    
    'Format the header row of the worksheet
    With ws.Range("A1").Resize(1, ColCount)
        .Interior.Color = rgbCornflowerBlue
        .Font.Color = rgbWhite
        .Font.Bold = True
    End With
    
    'Copy values from the recordset into the worksheet
    ws.Range("A2").CopyFromRecordset rs
    
    'Write column names into row 1 of the worksheet
    For i = 0 To ColCount - 1
        With rs.Fields(i)
            ws.Range("A1").Offset(0, i).Value = .Name
            
            'Apply a custom date format to date columns
            If .Type = adDate Then
                ws.Range("A1").Offset(1, i).Resize(RowCount, 1).NumberFormat = "dd mmm yyyy"
            End If
        End With
    Next i
    
    'Change the column widths on the worksheet
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close the recordset and connection
    'This will happen anyway when the local variables go out of scope at the end of the subroutine
    rs.Close
    cn.Close
    
    'Free resources used by the recordset and connection
    'This will happen anyway when the local variables go out of scope at the end of the subroutine
    Set rs = Nothing
    Set cn = Nothing
    
    'Exit here to make sure that the error handling code does not run
    Exit Sub
    
'========================================================================
'ERROR HANDLERS
'========================================================================
CloseRecordset:
'If the recordset is opened successfully but a runtime error occurs later we end up here
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
    Debug.Print SQLQuery
    
    MsgBox _
        Prompt:="An error occurred after the recordset was opened." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Error After Recordset Open"
    
    Exit Sub

CloseConnection:
'If the connection is opened successfully but a runtime error occurs later we end up here
    cn.Close
    
    Set cn = Nothing
    
    Debug.Print SQLQuery
    
    MsgBox _
        Prompt:="An error occurred after the connection was established." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Error After Connection Open"
    
    Exit Sub
    
'If the connection failed to open we end up here
EndPoint:
    MsgBox _
        Prompt:="The connection failed to open." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Connection Error"
    
End Sub

Sub DeleteAllButMenuSheet()

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is MenuSheet Then
            If Not ws Is LogSheet Then
                ws.Delete
            End If
        End If
    Next ws
    
    Application.DisplayAlerts = True
    
End Sub
