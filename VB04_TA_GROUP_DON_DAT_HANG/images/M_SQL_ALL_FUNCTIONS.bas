Attribute VB_Name = "M_SQL_ALL_FUNCTIONS"
Option Explicit
'====================================================================================================================================================
'====================================...Declare Variables
'====================================================================================================================================================
Public ServerName As String
Public DatabaseName As String
Public TableName As String
Public LoginName As String
Public LoginPass As String
Public MaThanhVien As String
Public TenThanhVien As String
Public MaUser As String
Public MaNhanVien As String
Public TenNhanVien As String
'===========================================================
Public VAR_CONN_ALL_ARGUMENT_CONNECTION As ADODB.Connection
Public VAR_CONN_ALL_ARGUMENT_RECORDSET As ADODB.Recordset
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
'====================================...S_GET_INFORMATION_OF_USER_WHILE_LOGINNING
'====================================================================================================================================================
Sub S_GET_INFORMATION_OF_USER_WHILE_LOGINNING()
    ServerName = "103.90.227.154, 1433"
    DatabaseName = "MUTUAL_2024"
    LoginName = "sa"
    LoginPass = "Ta#9999"
End Sub
'====================================================================================================================================================
'====================================...S_GET_INFORMATION_OF_USER_AFTER_LOGIN_SUCCESSFULLY
'====================================================================================================================================================
Sub S_GET_INFORMATION_OF_USER_AFTER_LOGIN_SUCCESSFULLY()
    
'    ServerName = "QUANNGUYEN\SQLEXPRESS"
'    DatabaseName = "TBD_2024"
'    TableName = "TB_TEST_DON_DAT_HANG"
'    LoginName = "sa"
'    LoginPass = "Ta#9999"
    
'    ServerName = "KSNB3\SQL2014"
'    ServerName = "103.90.227.154, 1433"
'    DatabaseName = "TBD_2024"
'    TableName = "TB_TEST_DON_DAT_HANG"
'    TableName = "TB_USER_ID"
'    TableName = "TB_DON_DAT_HANG"
'    LoginName = "sa"
'    LoginPass = "_!d96KjXvw'\"
'    LoginPass = "Ta#9999"
    
    ServerName = SH_USER_INFO.Range("G4").Value
    DatabaseName = SH_USER_INFO.Range("H4").Value
    
    LoginName = SH_USER_INFO.Range("I4").Value
    LoginPass = SH_USER_INFO.Range("J4").Value
    
    MaThanhVien = SH_USER_INFO.Range("E4").Value
    TenThanhVien = SH_USER_INFO.Range("F4").Value
    
    MaUser = SH_USER_INFO.Range("A4").Value
    MaNhanVien = SH_USER_INFO.Range("C4").Value
    TenNhanVien = SH_USER_INFO.Range("D4").Value
End Sub
'====================================================================================================================================================
'====================================...F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT
'====================================...https://stackoverflow.com/questions/10708077/fastest-way-to-transfer-excel-table-data-to-sql-2008r2
'====================================================================================================================================================
Sub S_TEST_F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT()

    Debug.Print ServerName
    Debug.Print DatabaseName
    Debug.Print TableName
    Debug.Print LoginName
    Debug.Print LoginPass
    Dim arr As Variant
    
    Call S_GET_INFORMATION_OF_USER_AFTER_LOGIN_SUCCESSFULLY
    
    arr = F_FIND_ANY_RANGE_IN_ANY_SHEET_01("RANGE_DATA_IMPORT_05", "SH_DATA_IMPORT", False)
    'Get Data to range
    Call F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT(ServerName, DatabaseName, TableName, LoginName, LoginPass, arr)
End Sub
Function F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT(ServerName As String, DatabaseName As String, TableName As String, LoginName As String, LoginPass As String, DataArray As Variant)
    
'    Debug.Print ServerName
'    Debug.Print DatabaseName
'    Debug.Print TableName
'    Debug.Print LoginName
'    Debug.Print LoginPass
'    Debug.Print DataArray(1, 2)
    
    On Error GoTo ErrorHandler
    
    Dim Table As String
    Dim Con As Object
    Dim cmd As Object
    Dim level As Long
    Dim arr As Variant
    Dim row As Long
    Dim rowCount As Long

    Set Con = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    'Creating a connection
    Con.ConnectionString = "Provider=SQLOLEDB;" & _
                                    "Data Source=" & ServerName & ";" & _
                                    "Initial Catalog=" & DatabaseName & ";" & _
                                    "UID=" & LoginName & "; PWD=" & LoginPass & ";"

    'Setting provider Name
    Con.Provider = "Microsoft.JET.OLEDB.12.0"

    'Opening connection
    Con.Open
    If Con.State = 1 Then
        
    End If
    cmd.CommandType = 1                         ' adCmdText
    

    
    Dim Rst As Object
    Set Rst = CreateObject("ADODB.Recordset")
    Table = TableName 'This should be same as the database table name.
    
    With Rst
        Set .ActiveConnection = Con
        .Source = "SELECT * FROM " & Table
        .CursorLocation = 3                     ' adUseClient
        .LockType = 4                           ' adLockBatchOptimistic
        .CursorType = 0                         ' adOpenForwardOnly
        .Open

        Dim tableFields(200) As Integer
        Dim rangeFields(200) As Integer
        Dim exportFieldsCount As Integer
        Dim ExportRangeToSQL As Integer
        Dim endRow As Long
        Dim flag As Boolean

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
        
        'Check ExportFieldsCount
        If exportFieldsCount = 0 Then
            ExportRangeToSQL = 1
            GoTo ConnectionEnd
        End If

        'This range selection column count should be same as database table column count.
        arr = DataArray

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
    
    'Announcement
'    Debug.Print "Successsfull!"
    
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
'====================================...S_GET_AN_ARRAY_WITH_HEADER_FROM_SQL_SERVER_TO_ANY_RANGE_IN_ANY_SHEET_01
'====================================================================================================================================================
Sub S_GET_AN_ARRAY_WITH_HEADER_FROM_SQL_SERVER_TO_ANY_RANGE_IN_ANY_SHEET_01(Var_query As String, DestinyRange As String, DestinySheet As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As Range
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
        Exit Sub
    End If
'====================================...WRITE VAR_CONN_ALL_ARGUMENT_RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveFirst

    'Populating Array with Headers from VAR_CONN_ALL_ARGUMENT_RECORDSET'
    For j = 0 To colcnt
         MyArray(0, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields(j).Name
'         Debug.Print MyArray(0, j)
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
    Set Dest_range = F_FIND_ANY_RANGE_IN_ANY_SHEET_01(DestinyRange, DestinySheet, True)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Sub

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Sub
'    Resume
End Sub
'====================================================================================================================================================
'====================================...S_GET_AN_ARRAY_WITHOUT_HEADER_FROM_SQL_SERVER_TO_ANY_RANGE_IN_ANY_SHEET_01
'====================================================================================================================================================
Sub S_GET_AN_ARRAY_WITHOUT_HEADER_FROM_SQL_SERVER_TO_ANY_RANGE_IN_ANY_SHEET_01(Var_query As String, DestinyRange As String, DestinySheet As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As Range
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
        Exit Sub
    End If
'====================================...WRITE VAR_CONN_ALL_ARGUMENT_RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveFirst

'    'Populating Array with Headers from VAR_CONN_ALL_ARGUMENT_RECORDSET'
'    For j = 0 To colcnt
'         MyArray(0, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields(j).Name
''         Debug.Print MyArray(0, j)
'    Next

    'Populating Array with Record Data
    For I = 0 To rowcnt - 1
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
    Set Dest_range = F_FIND_ANY_RANGE_IN_ANY_SHEET_01(DestinyRange, DestinySheet, False)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Sub

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Sub
'    Resume
End Sub
'====================================================================================================================================================
'====================================...S_GET_AN_ARRAY_WITHOUT_HEADER_FROM_SQL_SERVER_TO_RANGE_USER_INFO
'====================================================================================================================================================
Sub S_GET_AN_ARRAY_WITHOUT_HEADER_FROM_SQL_SERVER_TO_RANGE_USER_INFO(Var_query As String, DestinyRange As String, DestinySheet As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)

    On Error GoTo ErrorHandler
    Dim MyArray() As Variant    'unbound Array with no definite dimensions'
    Dim Dest_range As Range
    Dim I As Integer, j As Integer, colcnt As Integer, rowcnt As Integer
    
    Call S_GET_INFORMATION_OF_USER_WHILE_LOGINNING
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
        Exit Sub
    End If
'====================================...WRITE VAR_CONN_ALL_ARGUMENT_RECORDSET TO MYARRAY
    ReDim MyArray(rowcnt, colcnt)  'Redimension MyArray parameters to fit the SQL returned'
    VAR_CONN_ALL_ARGUMENT_RECORDSET.MoveFirst

'    'Populating Array with Headers from VAR_CONN_ALL_ARGUMENT_RECORDSET'
'    For j = 0 To colcnt
'         MyArray(0, j) = VAR_CONN_ALL_ARGUMENT_RECORDSET.Fields(j).Name
''         Debug.Print MyArray(0, j)
'    Next

    'Populating Array with Record Data
    For I = 0 To rowcnt - 1
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
    Set Dest_range = F_FIND_ANY_RANGE_IN_ANY_SHEET_01(DestinyRange, DestinySheet, False)
    Dest_range.ClearContents
    Dest_range.Resize(UBound(MyArray, 1) + 1, UBound(MyArray, 2) + 1).Value = MyArray   'Resize (secret sauce)
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Sub

ErrorHandler:
    Call S_THONG_BAO_LOI
    Call S_XU_LY_LOI
    ' Reset Error Trapping
    On Error GoTo 0
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
    Exit Sub
'    Resume
End Sub
'====================================================================================================================================================
'====================================...F_CREATE_SQL_QUERY_SELECT_ALL_DATA_FROM_TABLE_NO_WHERE_CONDITION
'====================================================================================================================================================
Function F_CREATE_SQL_QUERY_SELECT_ALL_DATA_FROM_TABLE_NO_WHERE_CONDITION(F_VarDatabaseName As String, F_VarTableName As String) As String
    Dim SQL_QUERY As String
    SQL_QUERY = "SELECT * FROM [" & F_VarDatabaseName & "].[dbo].[" & F_VarTableName & "]"
    Debug.Print SQL_QUERY
    
    'Return results
    F_CREATE_SQL_QUERY_SELECT_ALL_DATA_FROM_TABLE_NO_WHERE_CONDITION = SQL_QUERY
End Function
'====================================================================================================================================================
'====================================...F_CREATE_SQL_QUERY_SELECT_ALL_DATA_FROM_TABLE_NO_WHERE_CONDITION
'====================================================================================================================================================
Function F_CREATE_SQL_QUERY_SELECT_ALL_DATA_FROM_TABLE_ONE_WHERE_CONDITION(F_VarDatabaseName As String, F_VarTableName As String, F_VarCondition_01 As String) As String
    Dim SQL_QUERY As String
    SQL_QUERY = "SELECT * FROM [" & F_VarDatabaseName & "].[dbo].[" & F_VarTableName & "] " & _
                "WHERE " & F_VarCondition_01 & ";"
    Debug.Print SQL_QUERY
    
    'Return results
    F_CREATE_SQL_QUERY_SELECT_ALL_DATA_FROM_TABLE_ONE_WHERE_CONDITION = SQL_QUERY
End Function
'====================================================================================================================================================
'====================================...F_CREATE_SQL_QUERY_00_SELECT_FROM
'====================================================================================================================================================
Function F_CREATE_SQL_QUERY_00_SELECT_FROM(SELECT_STRING As String, FROM_STRING As String) As String
    Dim SQL_QUERY As String
    'Create SQL Query
    SQL_QUERY = "SELECT " & SELECT_STRING & " " & _
                "FROM " & FROM_STRING
'    Debug.Print SQL_QUERY
    
    'Return results
    F_CREATE_SQL_QUERY_00_SELECT_FROM = SQL_QUERY
End Function
'====================================================================================================================================================
'====================================...F_CREATE_SQL_QUERY_01_SELECT_FROM_WHERE
'====================================================================================================================================================
Function F_CREATE_SQL_QUERY_01_SELECT_FROM_WHERE(SELECT_STRING As String, FROM_STRING As String, WHERE_STRING As String) As String
    Dim SQL_QUERY As String
    'Create SQL Query
    SQL_QUERY = "SELECT " & SELECT_STRING & " " & _
                "FROM " & FROM_STRING & " " & _
                "WHERE " & WHERE_STRING
'    Debug.Print SQL_QUERY
    
    'Return results
    F_CREATE_SQL_QUERY_01_SELECT_FROM_WHERE = SQL_QUERY
End Function
'====================================================================================================================================================
'====================================...F_CREATE_SQL_QUERY_02_SELECT_FROM_WHERE_ORDERBY
'====================================================================================================================================================
Function F_CREATE_SQL_QUERY_02_SELECT_FROM_WHERE_ORDERBY(SELECT_STRING As String, FROM_STRING As String, WHERE_STRING As String, OrderBy_String As String) As String
    Dim SQL_QUERY As String
    'Create SQL Query
    SQL_QUERY = "SELECT " & SELECT_STRING & " " & _
                "FROM " & FROM_STRING & " " & _
                "WHERE " & WHERE_STRING & " " & _
                "ORDER BY " & OrderBy_String
'    Debug.Print SQL_QUERY
    
    'Return results
    F_CREATE_SQL_QUERY_02_SELECT_FROM_WHERE_ORDERBY = SQL_QUERY
End Function
'====================================================================================================================================================
'====================================...Set Connection String
'====================================================================================================================================================
Sub S_test_S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11()
    
'    Dim ServerName As String
'    Dim DatabaseName As String
'    Dim LoginName As String
'    Dim LoginPass As String
'
'    ServerName = F_SQL_GET_SERVER_NAME_01
'    DatabaseName = "DATABASE_USER_ID"
'    LoginName = F_SQL_GET_LOGIN_NAME_01
'    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Call S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Sub S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String)
    
    Set VAR_CONN_ALL_ARGUMENT_CONNECTION = New ADODB.Connection
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = New ADODB.Recordset
    
    With VAR_CONN_ALL_ARGUMENT_CONNECTION
        .ConnectionString = "Provider=SQLNCLI11" & _
                            ";Server=" & ServerName & _
                            ";database=" & DatabaseName & _
                            ";User Id=" & LoginName & _
                            "; Password=" & LoginPass & _
                            ";"
        .ConnectionTimeout = 10
        .Open
    End With
    If VAR_CONN_ALL_ARGUMENT_CONNECTION.State = 1 Then
'        Debug.Print "CONNECTION_ALL_ARGUMENT CONNECTED!"
    End If
End Sub
'====================================================================================================================================================
'====================================...F_CHECK_DUPLICATE_ANY_VALUE_IN_ONE_COLUMN
'====================================================================================================================================================
Function F_CHECK_DUPLICATE_ANY_VALUE_IN_ONE_COLUMN(SoPhieu_String As String, ColumnName_String As String, TableName_String As String) As Boolean
    'Create SQL Query
    Dim SQL_QUERY As String
    Dim SelectString As String
    Dim FromString As String
    Dim WhereString As String
    
    SelectString = "COUNT(" & ColumnName_String & ")"
    FromString = "[" & DatabaseName & "].[dbo].[" & TableName_String & "]"
    WhereString = "" & ColumnName_String & " = '" & SoPhieu_String & "' "
    
    SQL_QUERY = F_CREATE_SQL_QUERY_01_SELECT_FROM_WHERE(SelectString, FromString, WhereString)
    
    'Kiem tra voi Query
    Dim a As Variant
    a = F_GET_ONE_VALUE_FROM_DATABASE_WITH_SQL_QUERY(SQL_QUERY, ServerName, DatabaseName, LoginName, LoginPass)
    If a > 0 Then
        F_CHECK_DUPLICATE_ANY_VALUE_IN_ONE_COLUMN = True
    Else
        F_CHECK_DUPLICATE_ANY_VALUE_IN_ONE_COLUMN = False
    End If
End Function
'====================================================================================================================================================
'====================================...F_GET_ONE_VALUE_FROM_DATABASE_WITH_SQL_QUERY
'====================================================================================================================================================
Function F_GET_ONE_VALUE_FROM_DATABASE_WITH_SQL_QUERY(Var_query As String, ServerName As String, DatabaseName As String, LoginName As String, LoginPass As String) As Variant
    
    'Set connection
    Call S_SET_SQL_CONNECTION_ALL_ARGUMENT_SQLNCLI11(ServerName, DatabaseName, LoginName, LoginPass)
    
    'Excute SQL query
'    Debug.Print VAR_CONN_ALL_ARGUMENT_RECORDSET.State
'    Debug.Print Var_query
    VAR_CONN_ALL_ARGUMENT_RECORDSET.CursorLocation = adUseClient
    VAR_CONN_ALL_ARGUMENT_RECORDSET.Open Var_query, VAR_CONN_ALL_ARGUMENT_CONNECTION, adOpenStatic
    
    'Get 1 value
    F_GET_ONE_VALUE_FROM_DATABASE_WITH_SQL_QUERY = VAR_CONN_ALL_ARGUMENT_RECORDSET(0).Value
    
    'Close Recordset
    If VAR_CONN_ALL_ARGUMENT_RECORDSET.State = 1 Then
        VAR_CONN_ALL_ARGUMENT_RECORDSET.Close
    End If
    Set VAR_CONN_ALL_ARGUMENT_RECORDSET = Nothing: Set VAR_CONN_ALL_ARGUMENT_CONNECTION = Nothing
End Function

