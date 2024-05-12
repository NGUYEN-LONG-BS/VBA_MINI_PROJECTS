Attribute VB_Name = "M_SQL_ALL_FUNCTIONS"
Option Explicit
'====================================================================================================================================================
'====================================...Declare Variables
'====================================================================================================================================================
Public VAR_INFOR_MA_THANH_VIEN As String
Public VAR_INFOR_TEN_DANG_NHAP As String
Public VAR_INFOR_TEN_DATABASE As String
Public VAR_INFOR_MODULE_KINH_DOANH As String
Public VAR_INFOR_MODULE_VAT_TU As String
Public VAR_INFOR_MODULE_KY_THUAT As String
Public VAR_INFOR_MODULE_TAI_CHINH As String
Public VAR_INFOR_MODULE_ADMIN As String
Public VAR_NAM_TAI_CHINH As String


Public VAR_TABLE_RowsCount As Long
Public VAR_TABLE_ColumnsCount As Long
Private Type MyArraySettings
    RowsCount As Long
    ColumnCount As Long
    MyArray() As Variant
End Type
Public VAR_LOGIN_QUERY_STRING_01 As String
Public VAR_LOGIN_QUERY_STRING_02 As String
'====================================================================================================================================================
'====================================...SET Query
'====================================================================================================================================================
Sub S_SET_LOGIN_QUERY_01(a As String, b As String)
    VAR_LOGIN_QUERY_STRING_01 = "Select [Ten_Dang_Nhap] AS [Ten_Dang_Nhap], [Ten_Nhan_Vien] AS [Ten_Nhan_Vien] FROM [DATABASE_USER_ID].[dbo].[TABLE_USER_ID] WHERE ([Ten_Dang_Nhap] = '" & a & "') AND ([Pass_Dang_Nhap] = '" & b & "')"
End Sub
Sub S_SET_LOGIN_QUERY_02(a As String, b As String)
    VAR_LOGIN_QUERY_STRING_02 = "Select [Ten_Dang_Nhap] AS [Ten_Dang_Nhap], [Ten_Nhan_Vien] AS [Ten_Nhan_Vien] FROM [DATABASE_USER_ID].[dbo].[TABLE_USER_ID] WHERE ([Ten_Dang_Nhap] = '" & a & "') AND ([Pass_Dang_Nhap] = '" & b & "')"
End Sub
'====================================================================================================================================================
'====================================...SQL FUNCTION: Select an Array to Range
'====================================================================================================================================================
'====================================...https://stackoverflow.com/questions/5339807/return-multiple-values-from-a-function-sub-or-type
Sub S_TEST_F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_RANGE()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    Dim DatabaseName As String
    Dim LoginName As String
    Dim LoginPass As String
    
    Var_query = "SELECT [MA_DATABASE] as [MA DATABASE] " & _
                " ,[TEN_DATABASE] as [TEN DATABASE] " & _
                " ,[TEN_SERVER] as [TEN SERVER] " & _
                " FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]"
    
    ServerName = F_SQL_GET_SERVER_NAME_01
    DatabaseName = "DATABASE_USER_ID"
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_RANGE(Var_query, "RANGE_DATABASE_USER_ID_TB_DS_DATABASE", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Function F_SQL_GET_SERVER_NAME_01() As String
'    F_SQL_GET_SERVER_NAME_01 = "172.16.0.128, 1433"
'    F_SQL_GET_SERVER_NAME_01 = "QUANNGUYEN\SQLEXPRESS"
'    F_SQL_GET_SERVER_NAME_01 = "DESKTOP-GSMURCG\SQLEXPRESS"
    F_SQL_GET_SERVER_NAME_01 = "103.90.227.154, 1433"       'VPS Vietnix
End Function
Function F_SQL_GET_LOGIN_NAME_01() As String
    F_SQL_GET_LOGIN_NAME_01 = "sa"
End Function
Function F_SQL_GET_LOGIN_PASS_01() As String
    F_SQL_GET_LOGIN_PASS_01 = "Ta#9999"
End Function
Function F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_RANGE(Var_query As String, DestinyRange As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As range
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName, DatabaseName, LoginName, LoginPass)
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    VAR_CONN_ALL_ARGUMENT_RECORDSET.CursorLocation = adUseClient
    VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Var_query, VAR_CONN_ALL_ARGUMENT_CONNECTION, adOpenStatic
    ' VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Source, ActiveConnection, CursorType, LockType, Options
    ' CusorType có 5 hang so lan luot la: adOpenDynamic = 2, adOpenForwardOnly = 0, adOpenKeyset =1, adOpenStatic = 3, adOpenUnspecified= -1
    ' Locktype có 5 hang so lan luot là: adLockBatchOptimistic= 4, adLockOptimistic=3, adLockPessimistic=2, adLockReadOnly=1, adLockUnspecified=-1
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.RecordCount <> 0 Then            '///Do NOT Use "Do While Not rst.EOF" Can cause Problems///'
        colcnt = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields.count - 1
        rowcnt = VAR_CONN_ALL_ARGUMENT_RECORDSET.RecordCount
     Else
        ' Reset Error Trapping
        On Error GoTo 0
        Exit Function
    End If
'====================================...WRITE VAR_CONN_ALL_ARGUMENT_RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveFirst

    'Populating Array with Headers from VAR_CONN_ALL_ARGUMENT_RECORDSET'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields(j).Name
    Next

    'Populating Array with Record Data
    For I = 1 To rowcnt
        For j = 0 To colcnt
            MyArray(I, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET(j)
        Next j
        VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveNext
    Next I
    
'    S_SQL_SELECT_AN_ARRAY_01 = MyArray
'====================================...WORKSHEET OUTPUT
'    Debug.Print UBound(MyArray, 1) + 1
'    Debug.Print UBound(MyArray, 2) + 1
'    SH_DASHBOARD.Range("A1").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
'    SH_DASHBOARD.Range("AN2").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
    
    VAR_TABLE_RowsCount = UBound(MyArray, 1) + 1
    VAR_TABLE_ColumnsCount = UBound(MyArray, 2) + 1
    
'    Call F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Set Dest_range = F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function
'    Resume
End Function
Function F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_ANY_RANGE_IN_ANY_SHEET(Var_query As String, DestinyRange As String, DestinySheet As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As range
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName, DatabaseName, LoginName, LoginPass)
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    VAR_CONN_ALL_ARGUMENT_RECORDSET.CursorLocation = adUseClient
    VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Var_query, VAR_CONN_ALL_ARGUMENT_CONNECTION, adOpenStatic
    ' VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Source, ActiveConnection, CursorType, LockType, Options
    ' CusorType có 5 hang so lan luot la: adOpenDynamic = 2, adOpenForwardOnly = 0, adOpenKeyset =1, adOpenStatic = 3, adOpenUnspecified= -1
    ' Locktype có 5 hang so lan luot là: adLockBatchOptimistic= 4, adLockOptimistic=3, adLockPessimistic=2, adLockReadOnly=1, adLockUnspecified=-1
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.RecordCount <> 0 Then            '///Do NOT Use "Do While Not rst.EOF" Can cause Problems///'
        colcnt = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields.count - 1
        rowcnt = VAR_CONN_ALL_ARGUMENT_RECORDSET.RecordCount
     Else
        ' Reset Error Trapping
        On Error GoTo 0
        Exit Function
    End If
'====================================...WRITE VAR_CONN_ALL_ARGUMENT_RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveFirst

    'Populating Array with Headers from VAR_CONN_ALL_ARGUMENT_RECORDSET'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields(j).Name
    Next

    'Populating Array with Record Data
    For I = 1 To rowcnt
        For j = 0 To colcnt
            MyArray(I, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET(j)
        Next j
        VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveNext
    Next I
    
'    S_SQL_SELECT_AN_ARRAY_01 = MyArray
'====================================...WORKSHEET OUTPUT
'    Debug.Print UBound(MyArray, 1) + 1
'    Debug.Print UBound(MyArray, 2) + 1
'    SH_DASHBOARD.Range("A1").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
'    SH_DASHBOARD.Range("AN2").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
    
    VAR_TABLE_RowsCount = UBound(MyArray, 1) + 1
    VAR_TABLE_ColumnsCount = UBound(MyArray, 2) + 1
    
'    Call F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
'    Set Dest_range = F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Set Dest_range = F_FIND_ANY_RANGE_IN_ANY_SHEET_01(DestinyRange, DestinySheet)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION: Select an Array to Sheet
'====================================================================================================================================================
Function F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query As String, DestinySheet As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As range
    Dim I As Long, j As Long, colcnt As Long, rowcnt As Long
    
    Call S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName, DatabaseName, LoginName, LoginPass)
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    VAR_CONN_ALL_ARGUMENT_RECORDSET.CursorLocation = adUseClient
    VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Var_query, VAR_CONN_ALL_ARGUMENT_CONNECTION, adOpenStatic
    ' VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Source, ActiveConnection, CursorType, LockType, Options
    ' CusorType có 5 hang so lan luot la: adOpenDynamic = 2, adOpenForwardOnly = 0, adOpenKeyset =1, adOpenStatic = 3, adOpenUnspecified= -1
    ' Locktype có 5 hang so lan luot là: adLockBatchOptimistic= 4, adLockOptimistic=3, adLockPessimistic=2, adLockReadOnly=1, adLockUnspecified=-1
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.RecordCount <> 0 Then            '///Do NOT Use "Do While Not rst.EOF" Can cause Problems///'
        colcnt = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields.count - 1
        rowcnt = VAR_CONN_ALL_ARGUMENT_RECORDSET.RecordCount
     Else
        ' Reset Error Trapping
        On Error GoTo 0
        Exit Function
    End If
'====================================...WRITE VAR_CONN_ALL_ARGUMENT_RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveFirst

    'Populating Array with Headers from VAR_CONN_ALL_ARGUMENT_RECORDSET'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields(j).Name
    Next

    'Populating Array with Record Data
    For I = 1 To rowcnt
        For j = 0 To colcnt
            MyArray(I, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET(j)
        Next j
        VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveNext
    Next I
    
'    S_SQL_SELECT_AN_ARRAY_01 = MyArray
'====================================...WORKSHEET OUTPUT
'    Debug.Print UBound(MyArray, 1) + 1
'    Debug.Print UBound(MyArray, 2) + 1
'    SH_DASHBOARD.Range("A1").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
'    SH_DASHBOARD.Range("AN2").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
    
    VAR_TABLE_RowsCount = UBound(MyArray, 1) + 1
    VAR_TABLE_ColumnsCount = UBound(MyArray, 2) + 1
'    Debug.Print VAR_TABLE_RowsCount
'    Debug.Print VAR_TABLE_ColumnsCount

    Dim ws As Worksheet
    Dim sheet_index As Integer
    For Each ws In ThisWorkbook.Worksheets
         If ws.CodeName = DestinySheet Then
            sheet_index = ws.Index
            Exit For
         End If
    Next ws
    Set Dest_range = Worksheets(sheet_index).Cells(1, 1).CurrentRegion
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)

    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT
'====================================...https://stackoverflow.com/questions/10708077/fastest-way-to-transfer-excel-table-data-to-sql-2008r2
'====================================================================================================================================================
Sub S_TEST_F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim ServerName As String
    Dim DatabaseName As String
    Dim TableName As String
    Dim LoginName As String
    Dim LoginPass As String
    
    ServerName = F_SQL_GET_SERVER_NAME_01
    DatabaseName = "DATABASE_USER_ID"
    TableName = "EmployeeDetails"
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
'====================================...Get Data to range
    Call F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT(ServerName, DatabaseName, TableName, LoginName, LoginPass)
End Sub
Function F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT(ServerName As String, DatabaseName As String, TableName As String, LoginName As String, LoginPass As String)
    
    On Error GoTo ErrorHandler
    Dim sheet As Worksheet
    Set sheet = SH_DATA_IMPORT
    
    Dim Table As String
    Dim Con As Object
    Dim cmd As Object
    Dim level As Long
    Dim arr As Variant
    Dim row As Long
    Dim rowCount As Long

    Set Con = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    Call S_SET_LOGIN_NAME_01
    Call S_SET_LOGIN_PASS_01
    
    'Creating a connection
    Con.ConnectionString = "Provider=SQLOLEDB;" & _
                                    "Data Source=" & ServerName & ";" & _
                                    "Initial Catalog=" & DatabaseName & ";" & _
                                    "UID=" & Var_login_name_01 & "; PWD=" & Var_login_pass_01 & ";"

    'Setting provider Name
     Con.Provider = "Microsoft.JET.OLEDB.12.0"

    'Opening connection
     Con.Open
    If Con.State = 1 Then
'        Debug.Print "Connection Import Connected!"
    End If
    cmd.CommandType = 1             ' adCmdText
    
'    Call S_OpenStatusBar                 ' 0% Completed
    
    Dim Rst As Object
    Set Rst = CreateObject("ADODB.Recordset")
'    Table = "EmployeeDetails" 'This should be same as the database table name.
    Table = TableName 'This should be same as the database table name.
    
    With Rst
        Set .ActiveConnection = Con
        .Source = "SELECT * FROM " & Table
        .CursorLocation = 3         ' adUseClient
        .LockType = 4               ' adLockBatchOptimistic
        .CursorType = 0             ' adOpenForwardOnly
        .Open

        Dim tableFields(200) As Integer
        Dim rangeFields(200) As Integer

        Dim exportFieldsCount As Integer
        Dim ExportRangeToSQL As Integer         'Tu them vo
        Dim endRow As Long                      'Tu them vo
        Dim flag As Boolean                     'Tu them vo
        
        exportFieldsCount = 0

        Dim col As Integer
        Dim Index As Integer
        Index = 1
        
'        Debug.Print .Fields.count
        For col = 1 To .Fields.count
            exportFieldsCount = exportFieldsCount + 1
            tableFields(exportFieldsCount) = col
            rangeFields(exportFieldsCount) = Index
            Index = Index + 1
        Next
        
        If exportFieldsCount = 0 Then
            ExportRangeToSQL = 1
            GoTo ConnectionEnd
        End If

'        endRow = SH_DATA_IMPORT.Range("A65536").End(xlUp).row 'LastRow with the data.
'        Debug.Print endRow
''        arr = SH_DATA_IMPORT.Range("A1:CE" & endRow).Value 'This range selection column count should be same as database table column count.
'        endRow = SH_DATA_IMPORT.Rows.count
'        Debug.Print endRow
        
        ' Get the last row and column
        Dim lastRow As Long, lastColumn As Variant
        lastRow = SH_DATA_IMPORT.Cells.Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
'        Debug.Print SH_DATA_IMPORT.Cells.Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
        lastColumn = SH_DATA_IMPORT.Cells.Find("*", searchorder:=xlByColumns, SearchDirection:=xlPrevious).Column
'        Debug.Print lastRow
'        Debug.Print lastColumn
        
        
'        arr = SH_DATA_IMPORT.Range("A2:CE" & endRow).Value 'This range selection column count should be same as database table column count.
        arr = SH_DATA_IMPORT.range(SH_DATA_IMPORT.Cells(2, 1), SH_DATA_IMPORT.Cells(lastRow, lastColumn)).Value 'This range selection column count should be same as database table column count.


        rowCount = UBound(arr, 1)

        Dim val As Variant
        
'        Call S_RunStatusBar(5)     ' 5% Completed
        
'        Debug.Print rowCount
        For row = 1 To rowCount
            .AddNew
            For col = 1 To exportFieldsCount
                val = arr(row, rangeFields(col))
'                    Debug.Print val
                    .Fields(tableFields(col - 1)) = val
            Next
'            Call S_RunStatusBar(row / rowCount * 100)    ' 50% Completed
        Next

        .UpdateBatch
    End With

    flag = True
    
'    Unload UF_PROGRESS
    MsgBox "XONG"
    
ConnectionEnd:                               'Tu them vo
    ' Reset Error Trapping
    On Error GoTo 0
    'Closing RecordSet.
     If Rst.State = 1 Then
       Rst.Close
    End If
   'Closing Connection Object.
    If Con.State = 1 Then
      Con.Close
    End If
    'Setting empty for the RecordSet & Connection Objects
    Set Rst = Nothing
    Set Con = Nothing
    Exit Function
    
ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    'Closing RecordSet.
     If Rst.State = 1 Then
       Rst.Close
    End If
   'Closing Connection Object.
    If Con.State = 1 Then
      Con.Close
    End If
    'Setting empty for the RecordSet & Connection Objects
    Set Rst = Nothing
    Set Con = Nothing
'    Unload UF_PROGRESS
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...F_IMPORT_INTO_SQL_FROM_ARRAY
'====================================================================================================================================================
Sub S_TEST_F_IMPORT_INTO_SQL_FROM_ARRAY_01()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim ServerName As String
    Dim DatabaseName As String
    Dim TableName As String
    Dim LoginName As String
    Dim LoginPass As String
'    Dim MyArray(1 To 3) As Variant
    Dim MyArray(1 To 1, 1 To 6) As Variant
    
    MyArray(1, 1) = "A1001"
    MyArray(1, 2) = "A1002"
    MyArray(1, 3) = "A1003"
    MyArray(1, 4) = "A1004"
    MyArray(1, 5) = "A1005"
    MyArray(1, 6) = Now()       'NGAY_KHOI_TAO
    
    ServerName = F_SQL_GET_SERVER_NAME_01
    DatabaseName = "DATABASE_USER_ID"
    TableName = "TB_TEST"
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
'====================================...In toan bo thanh phan trong mang
'    Dim row_index As Integer
'    Dim col_index As Integer
'    For row_index = LBound(MyArray, 1) To UBound(MyArray, 1)
'        For col_index = LBound(MyArray, 2) To UBound(MyArray, 2)
'            Debug.Print MyArray(row_index, col_index)
'        Next col_index
'    Next row_index
    
    Call F_IMPORT_INTO_SQL_FROM_ARRAY_01(ServerName, DatabaseName, TableName, LoginName, LoginPass, MyArray)
End Sub
Function F_IMPORT_INTO_SQL_FROM_ARRAY_01(ServerName As String, DatabaseName As String, TableName As String, LoginName As String, LoginPass As String, ByRef MyArray() As Variant)
    
    On Error GoTo ErrorHandler
    
    Dim Con As Object
    Dim cmd As Object
    Dim Rst As Object
    
    Set Con = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    Set Rst = CreateObject("ADODB.Recordset")

    Dim Table As String
    Dim arr As Variant
    Dim rowCount As Long
    Dim row As Long
    
    'Creating a connection
    Con.ConnectionString = "Provider=SQLOLEDB;" & _
                                    "Data Source=" & ServerName & ";" & _
                                    "Initial Catalog=" & DatabaseName & ";" & _
                                    "UID=" & LoginName & ";" & _
                                    "PWD=" & LoginPass & ";"

    'Setting provider Name
     Con.Provider = "Microsoft.JET.OLEDB.12.0"

    'Opening connection
     Con.Open
    If Con.State = 1 Then
'        Debug.Print "Connection Connected: F_IMPORT_INTO_SQL_FROM_ARRAY_01!"
    End If
    cmd.CommandType = 1             ' adCmdText
    
'    Table = "EmployeeDetails" 'This should be same as the database table name.
    Table = TableName 'This should be same as the database table name.
    
    With Rst
        Set .ActiveConnection = Con
        .Source = "SELECT * FROM " & Table
        .CursorLocation = 3         ' adUseClient
        .LockType = 4               ' adLockBatchOptimistic
        .CursorType = 0             ' adOpenForwardOnly
        .Open

        Dim tableFields(200) As Integer         'Tao mot mang gom 200 thanh phan, moi thanh phan là integer
        Dim rangeFields(200) As Integer         'Tao mot mang gom 200 thanh phan, moi thanh phan là integer

        Dim exportFieldsCount As Integer
'        Dim ExportRangeToSQL As Integer         'Tu them vo
        Dim endRow As Long                      'Tu them vo
        Dim flag As Boolean                     'Tu them vo
        
        exportFieldsCount = 0

        Dim col As Integer
        Dim Index As Integer
        Index = 1
        
'        Debug.Print .Fields.count
        For col = 1 To .Fields.count                    'col chay tu 1 den tong so cot cua bang
            exportFieldsCount = exportFieldsCount + 1
'            Debug.Print exportFieldsCount
            tableFields(exportFieldsCount) = col
'            Debug.Print col
            rangeFields(exportFieldsCount) = Index
'            Debug.Print index
            Index = Index + 1
        Next
        
        If exportFieldsCount = 0 Then                   'Neu khong có cot nao
            MsgBox "Bang dich khong co cot de import"
            GoTo ConnectionEnd
        End If
        
        If exportFieldsCount <> UBound(MyArray, 2) Then
            MsgBox "So cot hai mang khong tuong ung"
            GoTo ConnectionEnd
        End If

'        Debug.Print LBound(MyArray, 1)
'        Debug.Print UBound(MyArray, 1)
'        Debug.Print LBound(MyArray, 2)
'        Debug.Print UBound(MyArray, 2)
'====================================...In toan bo thanh phan trong mang
        Dim row_index As Integer
        Dim col_index As Integer
        For row_index = LBound(MyArray, 1) To UBound(MyArray, 1)
            For col_index = LBound(MyArray, 2) To UBound(MyArray, 2)
'                Debug.Print MyArray(row_index, col_index)
            Next col_index
        Next row_index
        
        arr = MyArray()
        
'        Dim arr_index As Integer
'        For arr_index = LBound(arr, 1) To UBound(arr, 1)
''            Debug.Print arr(arr_index, 1)
'        Next arr_index

        rowCount = UBound(arr, 1)

        Dim val As Variant
'        Call S_OpenStatusBar                                ' Show Progress Bar
'        Debug.Print rowCount
'        Debug.Print rowCount
        For row = 1 To rowCount
            .AddNew
'            Debug.Print exportFieldsCount
            For col = 1 To exportFieldsCount
'                Debug.Print col
'                Debug.Print rangeFields(col)
'                Debug.Print arr(row, rangeFields(col))
                val = arr(row, rangeFields(col))
'                    Debug.Print val
                    .Fields(tableFields(col - 1)) = val
            Next
            
'            Call S_RunStatusBar(row / rowCount * 100)       ' Load Progress Bar
        Next
'        Unload UF_PROGRESS                                  ' Hide Progress Bar
        
        .UpdateBatch
    End With

    flag = True
    
    
    MsgBox "XONG"
    
ConnectionEnd:                               'Tu them vo
    ' Reset Error Trapping
    On Error GoTo 0
    'Closing RecordSet.
     If Rst.State = 1 Then
       Rst.Close
    End If
   'Closing Connection Object.
    If Con.State = 1 Then
      Con.Close
    End If
    'Setting empty for the RecordSet & Connection Objects
    Set Rst = Nothing
    Set Con = Nothing
    Exit Function
    
ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    'Closing RecordSet.
     If Rst.State = 1 Then
       Rst.Close
    End If
   'Closing Connection Object.
    If Con.State = 1 Then
      Con.Close
    End If
    'Setting empty for the RecordSet & Connection Objects
    Set Rst = Nothing
    Set Con = Nothing
    Unload UF_PROGRESS
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION: Select an Array num 01
'====================================================================================================================================================
Sub S_TEST_F_SQL_SELECT_AN_ARRAY_01()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Var_query = "SELECT [MA_DATABASE] as [MA DATABASE] " & _
                " ,[TEN_DATABASE] as [TEN DATABASE] " & _
                " ,[TEN_SERVER] as [TEN SERVER] " & _
                " FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]"
                
    Debug.Print Var_query
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_01(Var_query, "RANGE_DATABASE_USER_ID_TB_DS_DATABASE")
End Sub
Function F_SQL_SELECT_AN_ARRAY_01(Var_query As String, DestinyRange As String)
'Private Function F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY As String) As Variant
'Private Function F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY As String, ByRef RowsCount As Long, ByRef ColumnsCount As Long) As Variant
'Private Function F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY As String) As MyArraySettings
    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As range
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_01
'    Call S_SET_SQL_CONNECTION_02
    
'    Debug.Print VAR_Connection_01_recordset.State
    VAR_Connection_01_recordset.CursorLocation = adUseClient
'    VAR_Connection_01_recordset.CursorLocation = adUseServer
'    VAR_Connection_01_recordset.CursorLocation = adUseNone
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic
    ' recordset.Open Source, ActiveConnection, CursorType, LockType, Options
    ' CusorType có 5 hang so lan luot la: adOpenDynamic = 2, adOpenForwardOnly = 0, adOpenKeyset =1, adOpenStatic = 3, adOpenUnspecified= -1
    ' Locktype có 5 hang so lan luot là: adLockBatchOptimistic= 4, adLockOptimistic=3, adLockPessimistic=2, adLockReadOnly=1, adLockUnspecified=-1
    
'    Debug.Print VAR_Connection_01_recordset.State
    If VAR_Connection_01_recordset.RecordCount <> 0 Then            '///Do NOT Use "Do While Not rst.EOF" Can cause Problems///'
        colcnt = VAR_Connection_01_recordset.Fields.count - 1
        rowcnt = VAR_Connection_01_recordset.RecordCount
     Else
        ' Reset Error Trapping
        On Error GoTo 0
        Exit Function
    End If
'====================================...WRITE RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_Connection_01_recordset.MoveFirst

    'Populating Array with Headers from Recordset'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_Connection_01_recordset.Fields(j).Name
    Next

    'Populating Array with Record Data
    For I = 1 To rowcnt
        For j = 0 To colcnt
            MyArray(I, j) = VAR_Connection_01_recordset(j)
        Next j
        VAR_Connection_01_recordset.MoveNext
    Next I
    
'    S_SQL_SELECT_AN_ARRAY_01 = MyArray
'====================================...WORKSHEET OUTPUT
'    Debug.Print UBound(MyArray, 1) + 1
'    Debug.Print UBound(MyArray, 2) + 1
'    SH_DASHBOARD.Range("A1").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
'    SH_DASHBOARD.Range("AN2").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
    
    VAR_TABLE_RowsCount = UBound(MyArray, 1) + 1
    VAR_TABLE_ColumnsCount = UBound(MyArray, 2) + 1
    
'    Call F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Set Dest_range = F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_Connection_01_recordset.State = 1 Then
        VAR_Connection_01_recordset.Close
    End If
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_Connection_01_recordset.State = 1 Then
        VAR_Connection_01_recordset.Close
    End If
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL SUBROUTINE: S_GET_LOGIN_INFOMATION
'====================================================================================================================================================
Sub S_GET_LOGIN_INFORMATION(stringUser As String, stringPass As String)
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    Dim DatabaseName As String
    Dim LoginName As String
    Dim LoginPass As String
    
    Var_query = "SELECT " & _
                "[TEN_DANG_NHAP]," & _
                "[PASS_DANG_NHAP]," & _
                "[HO_TEN_NHAN_VIEN]," & _
                "[MA_THANH_VIEN]," & _
                "[TEN_THANH_VIEN]," & _
                "[MODULE_KINH_DOANH]," & _
                "[MODULE_VAT_TU]," & _
                "[MODULE_KY_THUAT]," & _
                "[MODULE_TAI_CHINH]," & _
                "[MODULE_ADMIN]" & _
                " FROM " & _
                "[DATABASE_USER_ID].[dbo].[TB_USER_ID]" & _
                " WHERE " & _
                "[TEN_DANG_NHAP] = '" & stringUser & "' AND " & _
                "[PASS_DANG_NHAP] = '" & stringPass & "'"
'    Debug.Print Var_query

    ServerName = F_SQL_GET_SERVER_NAME_01
    DatabaseName = "DATABASE_USER_ID"
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_RANGE(Var_query, "RANGE_LOGIN_INFOR", ServerName, DatabaseName, LoginName, LoginPass)
    
    Dim my_range As range
    Set my_range = F_FIND_RANGE_IN_SH_ALL_RANGES_01("RANGE_LOGIN_INFOR")
    VAR_INFOR_MA_THANH_VIEN = my_range(2, 4)
    VAR_INFOR_TEN_DANG_NHAP = my_range(2, 1)
    VAR_INFOR_MODULE_KINH_DOANH = my_range(2, 6)
    VAR_INFOR_MODULE_VAT_TU = my_range(2, 7)
    VAR_INFOR_MODULE_KY_THUAT = my_range(2, 8)
    VAR_INFOR_MODULE_TAI_CHINH = my_range(2, 9)
    VAR_INFOR_MODULE_ADMIN = my_range(2, 10)
    
'    Debug.Print VAR_INFOR_MA_THANH_VIEN
'    Debug.Print VAR_INFOR_MODULE_KINH_DOANH
'    Debug.Print VAR_INFOR_MODULE_VAT_TU
'    Debug.Print VAR_INFOR_MODULE_KY_THUAT
'    Debug.Print VAR_INFOR_MODULE_ADMIN
    Call S_GET_TEN_DATABASE
End Sub
Sub S_GET_TEN_DATABASE()
    Dim RowIndex As Integer
    Dim RNG As range
    Set RNG = F_FIND_RANGE_IN_SH_ALL_RANGES_01("RANGE_DATABASE_USER_ID_TB_DS_DATABASE")
    
    For RowIndex = RNG.Rows.count To 1 Step -1
'        Debug.Print RNG(RowIndex, 2)
'        Debug.Print VAR_INFOR_MA_THANH_VIEN & "23"
'        If RNG(RowIndex, 1) = VAR_INFOR_MA_THANH_VIEN & "23" Then
        If RNG(RowIndex, 1) = VAR_INFOR_MA_THANH_VIEN & VAR_NAM_TAI_CHINH Then
            VAR_INFOR_TEN_DATABASE = RNG(RowIndex, 2)
        End If
    Next RowIndex

End Sub
'====================================================================================================================================================
'====================================...SQL SUBROUTINE: S_GET_DATA_WHEN_LOGIN
'====================================================================================================================================================
Sub S_GET_DATA_WHEN_LOGIN()
    Call S_LOAD_ALL_DATA_FROM_DATABASE_TO_RANGE
End Sub
'====================================================================================================================================================
'====================================...SQL FUNCTION: Select an Array from Schema
'====================================================================================================================================================
Function F_SQL_SELECT_AN_ARRAY_02_Schema(Var_query As String, DestinyRange As String) '
'Private Function F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY As String) As Variant
'Private Function F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY As String, ByRef RowsCount As Long, ByRef ColumnsCount As Long) As Variant
'Private Function F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY As String) As MyArraySettings
    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As range
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_01
    
    Debug.Print VAR_Connection_01_recordset.State
    VAR_Connection_01_recordset.CursorLocation = adUseClient
'    VAR_Connection_01_recordset.CursorLocation = adUseServer
'    VAR_Connection_01_recordset.CursorLocation = adUseNone
'    VAR_Connection_01_recordset.Open VAR_QUERY, VAR_Connection_01, adOpenStatic
    Set VAR_Connection_01_recordset = VAR_Connection_01.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, Empty, "Table"))
    ' recordset.Open Source, ActiveConnection, CursorType, LockType, Options
    ' CusorType có 5 hang so lan luot la: adOpenDynamic = 2, adOpenForwardOnly = 0, adOpenKeyset =1, adOpenStatic = 3, adOpenUnspecified= -1
    ' Locktype có 5 hang so lan luot là: adLockBatchOptimistic= 4, adLockOptimistic=3, adLockPessimistic=2, adLockReadOnly=1, adLockUnspecified=-1
    
    Debug.Print VAR_Connection_01_recordset.State
    If VAR_Connection_01_recordset.RecordCount <> 0 Then            '///Do NOT Use "Do While Not rst.EOF" Can cause Problems///'
        colcnt = VAR_Connection_01_recordset.Fields.count - 1
        rowcnt = VAR_Connection_01_recordset.RecordCount
     Else
        ' Reset Error Trapping
        On Error GoTo 0
        Exit Function
    End If
    Debug.Print colcnt
    Debug.Print rowcnt
    rowcnt = 1
'====================================...WRITE RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_Connection_01_recordset.MoveFirst

    'Populating Array with Headers from Recordset'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_Connection_01_recordset.Fields(j).Name
    Next

    'Populating Array with Record Data
    For I = 1 To rowcnt
        For j = 0 To colcnt
            MyArray(I, j) = VAR_Connection_01_recordset(j)
        Next j
        VAR_Connection_01_recordset.MoveNext
    Next I
    
'    S_SQL_SELECT_AN_ARRAY_01 = MyArray
'====================================...WORKSHEET OUTPUT
'    Debug.Print UBound(MyArray, 1) + 1
'    Debug.Print UBound(MyArray, 2) + 1
'    SH_DASHBOARD.Range("A1").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
'    SH_DASHBOARD.Range("AN2").Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray
    
    VAR_TABLE_RowsCount = UBound(MyArray, 1) + 1
    VAR_TABLE_ColumnsCount = UBound(MyArray, 2) + 1
    
'    Call F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Set Dest_range = F_FIND_RANGE_IN_SH_ALL_RANGES_01(DestinyRange)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_Connection_01_recordset.State = 1 Then
        VAR_Connection_01_recordset.Close
    End If
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_Connection_01_recordset.State = 1 Then
        VAR_Connection_01_recordset.Close
    End If
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION
'====================================================================================================================================================
Sub S_SQL_SELECT_AN_ARRAY_01()
    
    On Error GoTo ErrorHandler
    Dim Var_query As String
    Var_query = "SELECT [MA_DATABASE] as [MA DATABASE] " & _
                                        " ,[TEN_DATABASE] as [TEN DATABASE] " & _
                                        " ,[TEN_SERVER] as [TEN SERVER] " & _
                                        " ,[XOA_DIEU_CHINH] as [XOA DIEU CHINH] " & _
                                        " ,[NGAY_KHOI_TAO] as [NGAY KHOI TAO] " & _
                                    " FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]"
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_01
    
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic
    
    If VAR_Connection_01_recordset.RecordCount <> 0 Then            '///Do NOT Use "Do While Not rst.EOF" Can cause Problems///'
        colcnt = VAR_Connection_01_recordset.Fields.count - 1
        rowcnt = VAR_Connection_01_recordset.RecordCount
     Else
        ' Reset Error Trapping
        On Error GoTo 0
        VAR_Connection_01_recordset.Close
        Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
        Exit Sub
    End If
    
'====================================...WRITE RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_Connection_01_recordset.MoveFirst

    'Populating Array with Headers from Recordset'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_Connection_01_recordset.Fields(j).Name
    Next

    'Populating Array with Record Data
    For I = 1 To rowcnt
        For j = 0 To colcnt
            MyArray(I, j) = VAR_Connection_01_recordset(j)
        Next j
        VAR_Connection_01_recordset.MoveNext
    Next I
    
'    S_SQL_SELECT_AN_ARRAY_01 = MyArray
'====================================...WORKSHEET OUTPUT
'    Set wb = ThisWorkbook
'    Set ws = wb.Worksheets("Insert Worksheet Name")
'    Set Dest = ws.Range("A1") 'Destination Cell
'    Dest.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = Application.Transpose(MyArray)  'Resize (secret sauce)
    
    SH_DASHBOARD.range("A1").Value = MyArray
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Sub
'    Resume
End Sub
'====================================================================================================================================================
'====================================...SQL FUNCTION: Update
'====================================================================================================================================================
Sub S_TEST_ALL_FUNCTION_UPDATE()
'====================================...Set SQL Query
    Dim VAR_QUERY_01 As String
    VAR_QUERY_01 = "UPDATE " & _
                        " [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE] " & _
                    " SET " & _
                        " [XOA_DIEU_CHINH] = '' " & _
                    " WHERE " & _
                        " [MA_DATABASE] = 'DN23'; "

'====================================...Update SQL
    Call F_SQL_UPDATE_SET_OLD_01(VAR_QUERY_01)
End Sub
Function F_SQL_UPDATE_SET_OLD_01(Var_query As String)
    On Error GoTo ErrorHandler
    
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_01
    
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic, adLockOptimistic
    
    On Error GoTo 0
'    VAR_Connection_01_recordset.Close      'Function Update closed recordset
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
'    VAR_Connection_01_recordset.Close      'Function Update closed recordset
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function
'    Resume
End Function
Function F_SQL_UPDATE_SET_CANCEL_01(Var_query As String)
    On Error GoTo ErrorHandler
    
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_SET_SQL_CONNECTION_01
    
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic, adLockOptimistic
    
    On Error GoTo 0
'    VAR_Connection_01_recordset.Close      'Function Update closed recordset
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
'    VAR_Connection_01_recordset.Close      'Function Update closed recordset
    Set VAR_Connection_01_recordset = Nothing: Set VAR_Connection_01 = Nothing
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION: Check Unique
'====================================================================================================================================================
Sub S_TEST_ALL_FUNCTION_CHECK_UNIQUE()
    Dim a As Boolean
    
    ' Test F_SQL_CHECK_UNIQUE_01
'    a = F_SQL_CHECK_UNIQUE_01("DN23", "MA_DATABASE", "TB_DS_DATABASE")
    a = F_SQL_CHECK_UNIQUE_01("MA_DATABASE", "[DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]", "MA_DATABASE", "DI23", "=", "XOA_DIEU_CHINH", "OLD", "NOT LIKE")
    Debug.Print a
End Sub
Function F_SQL_CHECK_UNIQUE_01(SelectColName01 As String, FromColName01 As String, WhereColName01 As String, WhereValue01 As String, WhereOperator01 As String _
                                                                                    , WhereColName02 As String, WhereValue02 As String, WhereOperator02 As String)
'    SELECT
'        COUNT(DISTINCT SelectColName01)
'    FROM
'        FromColName01
'    Where
'       WhereColName01 WhereOperator01 'WhereValue01'
'       AND
'       WhereColName02 WhereOperator02 'WhereValue02'

    ' Variables Declare
    Dim Var_query As String
    ' Set Error Trapping
    On Error GoTo ErrorHandler
    
    Var_query = "SELECT COUNT(DISTINCT " & SelectColName01 & ") From " & FromColName01 & " Where " & WhereColName01 & " " & WhereOperator01 & " '" & WhereValue01 & "' AND " _
                                                                                                   & WhereColName02 & " " & WhereOperator02 & " '" & WhereValue02 & "'"
'    Debug.Print VAR_QUERY
    
    ' Set Connection and open
    Call S_SET_SQL_CONNECTION_01
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic
    ' Set result
    If VAR_Connection_01_recordset(0) = 0 Then
        F_SQL_CHECK_UNIQUE_01 = False
    Else
        F_SQL_CHECK_UNIQUE_01 = True
    End If
'    VAR_Connection_01.Close
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION: Counting
'====================================================================================================================================================
Sub S_TEST_ALL_FUNCTION_SQL_COUTING()
    Dim a As Long
    
    ' Test F_SQL_CHECK_UNIQUE_01
'    a = F_SQL_COUNT_ALL_ROWS_OF_TABLE_TEST("SELECT COUNT(MA_DATABASE) as 'Tong_so_dong' FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE];")
    a = F_SQL_COUNT_ALL_ROWS_OF_TABLE_01("MA_DATABASE", "[DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]", "XOA_DIEU_CHINH", "OLD", "NOT LIKE")
    Debug.Print a
End Sub
Function F_SQL_COUNT_ALL_ROWS_OF_TABLE_01(SelectColName01 As String, FromColName01 As String, WhereColName01 As String, WhereValue01 As String, WhereOperator01 As String)
    ' Set Error Trapping
    On Error GoTo ErrorHandler
    Dim Var_query As String

    Var_query = "SELECT COUNT(" & SelectColName01 & ") FROM " & FromColName01 & " Where " & WhereColName01 & " " & WhereOperator01 & " '" & WhereValue01 & "'"
'    Debug.Print VAR_QUERY
    
    ' Set Connection and open
    Call S_SET_SQL_CONNECTION_01
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic
    ' Set result
    F_SQL_COUNT_ALL_ROWS_OF_TABLE_01 = VAR_Connection_01_recordset(0)
'    VAR_Connection_01.Close
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Function
'    Resume
End Function
Function F_SQL_COUNT_ALL_ROWS_OF_TABLE_TEST(Var_query As String)
    
    On Error GoTo ErrorHandler
    Call S_SET_SQL_CONNECTION_01
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic
    F_SQL_COUNT_ALL_ROWS_OF_TABLE_TEST = VAR_Connection_01_recordset(0)
'    VAR_Connection_01.Close
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    On Error GoTo 0
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION: F_SQL_KIEM_TRA_THONG_TIN_DANG_NHAP
'====================================================================================================================================================
Function F_SQL_KIEM_TRA_THONG_TIN_DANG_NHAP(stringUser As String, stringPass As String) As Boolean
    ' Set Error Trapping
    On Error GoTo ErrorHandler
    Dim Var_query As String

    Var_query = "SELECT COUNT([TEN_DANG_NHAP])" & _
                "FROM" & _
                "[DATABASE_USER_ID].[dbo].[TB_USER_ID]" & _
                "WHERE" & _
                "[TEN_DANG_NHAP] = '" & stringUser & "' AND" & _
                "[PASS_DANG_NHAP] = '" & stringPass & "'"
'    Debug.Print Var_query
    
    ' Set Connection and open
    Call S_SET_SQL_CONNECTION_01
    VAR_Connection_01_recordset.CursorLocation = adUseClient
    VAR_Connection_01_recordset.Open Var_query, VAR_Connection_01, adOpenStatic
    
    ' Set result
    If VAR_Connection_01_recordset(0) > 0 Then
        F_SQL_KIEM_TRA_THONG_TIN_DANG_NHAP = True
    ElseIf VAR_Connection_01_recordset(0) = 0 Then
        F_SQL_KIEM_TRA_THONG_TIN_DANG_NHAP = False
    End If
    
    'Turnoff connection
    If VAR_Connection_01_recordset.State = 1 Then
        VAR_Connection_01_recordset.Close
        Set VAR_Connection_01_recordset = Nothing
    End If
    If VAR_Connection_01.State = 1 Then
        VAR_Connection_01.Close
        Set VAR_Connection_01 = Nothing
    End If
    
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Function
'    Resume
End Function
'====================================================================================================================================================
'====================================...SQL FUNCTION: Delete
'====================================================================================================================================================
Function F_SQL_DELETE_ALL_RECORD_FROM_TABLE(Var_query As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
'    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
'    Dim Dest_range As Range
'    Dim i As Long, j As Long, colcnt As Long, rowcnt As Long
    
    Call S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName, DatabaseName, LoginName, LoginPass)
    
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
    VAR_CONN_ALL_ARGUMENT_RECORDSET.CursorLocation = adUseClient
    VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Var_query, VAR_CONN_ALL_ARGUMENT_CONNECTION, adOpenStatic
    ' VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Source, ActiveConnection, CursorType, LockType, Options
    ' CusorType có 5 hang so lan luot la: adOpenDynamic = 2, adOpenForwardOnly = 0, adOpenKeyset =1, adOpenStatic = 3, adOpenUnspecified= -1
    ' Locktype có 5 hang so lan luot là: adLockBatchOptimistic= 4, adLockOptimistic=3, adLockPessimistic=2, adLockReadOnly=1, adLockUnspecified=-1
    
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Function
'    Resume
End Function

'====================================================================================================================================================
'====================================...SQL FUNCTION: Insert Into
'====================================================================================================================================================
Sub S_SQL_INSERT_INTO_01(Var_query As String)
    ' Set Error Trapping
    On Error GoTo ErrorHandler
    ' Set Connection and open
    Call S_SET_SQL_CONNECTION_01
    VAR_Connection_01.Execute Var_query
'    VAR_Connection_01.Close
    
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    Exit Sub
'    Resume
End Sub
'====================================================================================================================================================
'====================================...SQL Error Handling
'====================================================================================================================================================
Sub S_THONG_BAO_LOI()
    Dim msg As String
'    Debug.Print Err.Number & "<< Ma Loi"
    msg = " - Error #: " & str(Err.Number) & " ==> (Ma loi)" & Chr(13) & _
            " - Was generated by: " & Err.Source & " ==> (Chuong trinh bao loi)" & Chr(13) & _
            " - Error Line: " & Erl & Chr(13) & _
            " - Error Desciption: " & Err.Description
    MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
End Sub
Sub S_XU_LY_LOI()
    If Err.Number = -2147217873 Then
        MsgBox "Huong dan: Da bi trung ID, vui long thay doi ID"
    End If
    If Err.Number = -2147217865 Then
        MsgBox "Huong dan: Khong tim thay ten bang trong SQL, vui long sua lai ten bang"
    End If
    If Err.Number = -2147217887 Then
        MsgBox "Huong dan: Kieu du lieu import khong dung quy dinh cua Table"
    End If
    If Err.Number = 3265 Then
        MsgBox "Huong dan: khong tim thay gia tri de thuc hien lenh"
    End If
    If Err.Number = -2147217900 Then
        MsgBox "Huong dan: Loi cau lenh SQL, can phai sua lai"
    End If
    If Err.Number = 3704 Then
        MsgBox "Huong dan: closet da bi dong, kiem tra lai code"
    End If
    If Err.Number = 6 Then
        MsgBox "Huong dan: Tran bo nho"
    End If
    If Err.Number = 91 Then
        MsgBox "Huong dan: Khong co du lieu IMPORT"
    End If
End Sub
'====================================================================================================================================================
'====================================...RANGE FUNCTION
'====================================================================================================================================================
Sub S_TEST_FUNCTION_F_FIND_RANGE_IN_SH_ALL_RANGES_01()
    Dim range1 As range
    Set range1 = F_FIND_RANGE_IN_SH_ALL_RANGES_01("RANGE_TEST")
End Sub
Function F_FIND_RANGE_IN_SH_ALL_RANGES_01(RangeNameString As String) As range
    Dim OutputRange As range
    Dim FirtsCellsRow As Long
    Dim FirtsCellsCol As Long
    Dim RangeNameRow As Long
    Dim RangeNameCol As Long
    
    FirtsCellsCol = Application.WorksheetFunction.Match(RangeNameString, SH_ALL_RANGES_01.Rows(1), 0)
    FirtsCellsRow = 3

    Set OutputRange = SH_ALL_RANGES_01.Cells(FirtsCellsRow, FirtsCellsCol).CurrentRegion
    Set F_FIND_RANGE_IN_SH_ALL_RANGES_01 = OutputRange
    
    RangeNameRow = SH_ALL_RANGES_01.Cells(FirtsCellsRow, FirtsCellsCol).CurrentRegion.Rows.count
    RangeNameCol = SH_ALL_RANGES_01.Cells(FirtsCellsRow, FirtsCellsCol).CurrentRegion.Columns.count
    
'    Debug.Print OutputRange.Rows.Count
'    Debug.Print OutputRange.Columns.Count
'    Debug.Print OutputRange.Address
'    OutputRange.Select
End Function
Sub S_TEST_FUNCTION_F_FIND_ANY_RANGE_IN_ANY_SHEET_01()
    Dim range1 As range
    Set range1 = F_FIND_ANY_RANGE_IN_ANY_SHEET_01("RANGE_COMBOBOX_TEN_KHACH_HANG", "SH_RANGE_MA_KH_01")
End Sub
Function F_FIND_ANY_RANGE_IN_ANY_SHEET_01(RangeNameString As String, SheetNameString As String) As range
    Dim OutputRange As range
    Dim FirtsCellsRow As Long
    Dim FirtsCellsCol As Long
    Dim RangeNameRow As Long
    Dim RangeNameCol As Long
    
    Dim ws As Worksheet
    Dim sheet_index As Integer
    For Each ws In ThisWorkbook.Worksheets
         If ws.CodeName = SheetNameString Then
            sheet_index = ws.Index
            Exit For
         End If
    Next ws
    
'    Debug.Print sheet_index
'    Debug.Print Worksheets(1).Name
'    Debug.Print Worksheets(sheet_index).Name
    
    FirtsCellsCol = Application.WorksheetFunction.Match(RangeNameString, Worksheets(sheet_index).Rows(1), 0)
    FirtsCellsRow = 3
    
    Set OutputRange = Worksheets(sheet_index).Cells(FirtsCellsRow, FirtsCellsCol).CurrentRegion
    Set F_FIND_ANY_RANGE_IN_ANY_SHEET_01 = OutputRange
    
    RangeNameRow = Worksheets(sheet_index).Cells(FirtsCellsRow, FirtsCellsCol).CurrentRegion.Rows.count
    RangeNameCol = Worksheets(sheet_index).Cells(FirtsCellsRow, FirtsCellsCol).CurrentRegion.Columns.count
    
'    Debug.Print OutputRange.Rows.count
'    Debug.Print OutputRange.Columns.count
'    Debug.Print OutputRange.Address
'    OutputRange.Select
End Function
'====================================================================================================================================================
'====================================...Excel Table
'====================================================================================================================================================
Sub S_EXCEL_TABLE_DELETE_ALL_ROWS()
    Dim RCount As Long
    Dim I As Long
    RCount = SH_ALL_COMBOBOX.ListObjects("TB_TEST").ListRows.count
    For I = RCount To 1 Step -1
        SH_ALL_COMBOBOX.ListObjects("TB_TEST").ListRows(I).Delete
    Next I
End Sub
Sub S_EXCEL_TABLE_DELETE_TABLE()
    SH_ALL_COMBOBOX.ListObjects("TB_TEST").Delete
End Sub
Sub S_EXCEL_TABLE_ADD_ONE_ROW_AT_BOTTOM()
'    SH_ALL_COMBOBOX.ListObjects("TB_TEST").ListRows.Add (5)
    SH_ALL_COMBOBOX.ListObjects("TB_TEST").ListRows.Add AlwaysInsert:=True
End Sub
Sub S_EXCEL_TABLE_ADD_MANY_ROWS_AFTER_ROW()
    Dim HowMuchRowToInsert As Long
    Dim AferWitchRowShouldIInsertRow As Long
    Dim I As Long
    
    HowMuchRowToInsert = 10
    AferWitchRowShouldIInsertRow = 3
    
    For I = 1 To HowMuchRowToInsert
        SH_ALL_COMBOBOX.ListObjects("TB_TEST").ListRows.Add (AferWitchRowShouldIInsertRow + I)
    Next
End Sub
Sub S_EXCEL_TABLE_ADD_MANY_ROWS_AT_BOTTOM()
    Dim HowMuchRowToInsert As Long
    Dim AferWitchRowShouldIInsertRow As Long
    Dim I As Long
    
    HowMuchRowToInsert = 10
    AferWitchRowShouldIInsertRow = 3
    
    For I = 1 To HowMuchRowToInsert
        SH_ALL_COMBOBOX.ListObjects("TB_TEST").ListRows.Add AlwaysInsert:=True
    Next
End Sub
Sub S_EXCEL_TABLE_ADD_ROWS_AND_VALUES()
    Dim TableName As ListObject
    Set TableName = SH_ALL_COMBOBOX.ListObjects("TB_TEST")
    Dim addedRow As ListRow
    Set addedRow = TableName.ListRows.Add()
    With addedRow
        .range(1) = "00006"
        .range(2) = "Nelson Biden"
        .range(3) = "Research"
        .range(4) = "30/11/2002"
        .range(5) = 150000
    End With
End Sub
'====================================================================================================================================================
'====================================...Excel Table
'====================================================================================================================================================
Sub S_LOAD_ALL_DATA_FROM_DATABASE_TO_RANGE()
    Call S_OpenStatusBar    ' 0% Completed
'====================================...
    Call S_LOAD_FROM_DATABASE_TO_RANGE_MA_SO_TAI_KHOAN
    Call S_RunStatusBar(5)     ' 5% Completed
'====================================...LOAD ALL RANGE KHACH HANG
'    Call S_LOAD_FROM_DATABASE_TO_RANGE_MA_KH
'    Call S_LOAD_RANGE_COMBOBOX_MA_KHACH_HANG
'    Call S_LOAD_RANGE_COMBOBOX_TEN_KHACH_HANG
    Call S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_COMBOBOX_MA_KHACH_HANG
    Call S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_COMBOBOX_TEN_KHACH_HANG
    Call S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_DANH_MUC_KHACH_HANG_FULL
    Call S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_DANH_MUC_KHACH_HANG_LIST_BOX
    Call S_RunStatusBar(10)     ' 10% Completed
'====================================...LOAD ALL RANGE XTHH
    Call S_LOAD_DATA_TO_SH_RANGE_XTHH_01_RANGE_DN_XTHH_FULL
'====================================...
    Call S_LOAD_FROM_DATABASE_TO_RANGE_MA_HANG_HOA
    Call S_RunStatusBar(15)     ' 15% Completed
'====================================...
    Call S_LOAD_FROM_DATABASE_TO_RANGE_MA_NCC
    Call S_RunStatusBar(20)     ' 20% Completed
'====================================...
    Call S_LOAD_FROM_DATABASE_TO_RANGE_MA_NCC_DICH_VU
'====================================...
    Call S_RunStatusBar(100)     ' 100% Completed
    Unload UF_PROGRESS
End Sub
Sub S_LOAD_FROM_DATABASE_TO_RANGE_MA_SO_TAI_KHOAN()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_PHAN_LOAI_TAI_KHOAN"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_TAI_KHOAN] as [MA_TAI_KHOAN] " & _
                " ,[TEN_TAI_KHOAN] as [TEN_TAI_KHOAN] " & _
                " ,[PL_TM_NH] as [PL_TM_NH] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_MA_SO_TAI_KHOAN", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_FROM_DATABASE_TO_RANGE_MA_KH()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_RANGE_MA_KH", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_RANGE_COMBOBOX_MA_KHACH_HANG()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_RANGE_MA_KH", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_COMBOBOX_MA_KHACH_HANG()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_ANY_RANGE_IN_ANY_SHEET(Var_query, "RANGE_COMBOBOX_MA_KHACH_HANG", "SH_RANGE_MA_KH_01", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_COMBOBOX_TEN_KHACH_HANG()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT " & _
                " [TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_ANY_RANGE_IN_ANY_SHEET(Var_query, "RANGE_COMBOBOX_TEN_KHACH_HANG", "SH_RANGE_MA_KH_01", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_DANH_MUC_KHACH_HANG_FULL()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_ANY_RANGE_IN_ANY_SHEET(Var_query, "RANGE_DANH_MUC_KHACH_HANG_FULL", "SH_RANGE_MA_KH_01", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_DATA_TO_SH_RANGE_MA_KH_01_RANGE_DANH_MUC_KHACH_HANG_LIST_BOX()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_ANY_RANGE_IN_ANY_SHEET(Var_query, "RANGE_DANH_MUC_KHACH_HANG_LIST_BOX", "SH_RANGE_MA_KH_01", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_RANGE_COMBOBOX_TEN_KHACH_HANG()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_KHACH_HANG"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [TEN_KHACH_HANG] as [TEN_KHACH_HANG] " & _
                " ,[MA_KHACH_HANG] as [MA_KHACH_HANG] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " ,[MA_KHU_VUC] as [MA_KHU_VUC] " & _
                " ,[PHAN_LOAI_DL_XL_NB] as [PHAN_LOAI_DL_XL_NB] " & _
                " ,[MA_DIEN_LUC] as [MA_DIEN_LUC] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_RANGE_MA_KH", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_FROM_DATABASE_TO_RANGE_MA_HANG_HOA()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_HANG_HOA"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_HANG] as [MA_HANG] " & _
                " ,[TEN_HANG] as [TEN_HANG] " & _
                " ,[DVT] as [DVT] " & _
                " ,[DON_GIA_DAU_KY] as [DON_GIA_DAU_KY] " & _
                " ,[SO_LUONG_TON_DAU_KY] as [SO_LUONG_TON_DAU_KY] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_RANGE_MA_HANG", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_FROM_DATABASE_TO_RANGE_MA_NCC()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_NCC"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_NCC] as [MA_NCC] " & _
                " ,[TEN_NCC] as [TEN_NCC] " & _
                " ,[MA_DAU_KY] as [MA_DAU_KY] " & _
                " ,[MA_TRONG_KY] as [MA_TRONG_KY] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_RANGE_MA_NCC", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_LOAD_FROM_DATABASE_TO_RANGE_MA_NCC_DICH_VU()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_MA_NCC_DICH_VU"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [MA_NCC_DICH_VU] as [MA_NCC_DICH_VU] " & _
                " ,[TEN_NCC_DICH_VU] as [TEN_NCC_DICH_VU] " & _
                " ,[MA_DAU_KY] as [MA_DAU_KY] " & _
                " ,[MA_TRONG_KY] as [MA_TRONG_KY] " & _
                " ,[MST_CHINH] as [MST_CHINH] " & _
                " ,[MST_CN] as [MST_CN] " & _
                " ,[DIA_CHI] as [DIA_CHI] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' "
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, "SH_RANGE_MA_NCC_DICH_VU", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
'====================================================================================================================================================
'====================================...SHEET XUAT TRUOC HANG HOA
'====================================================================================================================================================
Sub S_LOAD_DATA_TO_SH_RANGE_XTHH_01_RANGE_DN_XTHH_FULL()
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Dim ServerName As String
    ServerName = F_SQL_GET_SERVER_NAME_01
    Dim DatabaseName As String
    Call S_GET_TEN_DATABASE
    DatabaseName = VAR_INFOR_TEN_DATABASE
    Dim TableName As String
    TableName = "TB_KD0304_NHAT_KY_XTHH_PHIEU_DE_NGHI"
    Dim LoginName As String
    LoginName = F_SQL_GET_LOGIN_NAME_01
    Dim LoginPass As String
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Var_query = "SELECT [NGAY_LAP] as [NGAY_LAP] " & _
                " ,[SO_PHIEU_DN_XTHH] as [SO_PHIEU_DN_XTHH] " & _
                " ,[MA_KH] as [MA_KH] " & _
                " ,[TEN_KH] as [TEN_KH] " & _
                " ,[STT_DONG_NHAP] as [STT_DONG_NHAP] " & _
                " ,[MA_HANG] as [MA_HANG] " & _
                " ,[TEN_HANG] as [TEN_HANG] " & _
                " ,[DVT] as [DVT] " & _
                " ,[SO_LUONG] as [SO_LUONG] " & _
                " ,[DON_GIA_VON] as [DON_GIA_VON] " & _
                " ,[THANH_TIEN] as [THANH_TIEN] " & _
                " ,[KHO_A] as [KHO_A] " & _
                " ,[KHO_B] as [KHO_B] " & _
                " ,[SO_HOP_DONG] as [SO_HOP_DONG] " & _
                " ,[SO_DUYET_GIA] as [SO_DUYET_GIA] " & _
                " ,[CHUNG_TU_KHAC] as [CHUNG_TU_KHAC] " & _
                " ,[GHI_CHU_CUA_PHIEU] as [GHI_CHU_CUA_PHIEU] " & _
                " FROM [" & DatabaseName & "].[dbo].[" & TableName & "]" & _
                " WHERE [XOA_DIEU_CHINH] = '' " & _
                " ORDER BY [NGAY_LAP] ASC "
'    Debug.Print Var_query
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_ANY_RANGE_IN_ANY_SHEET(Var_query, "RANGE_DN_XTHH_FULL", "SH_RANGE_XTHH_01", ServerName, DatabaseName, LoginName, LoginPass)
End Sub
