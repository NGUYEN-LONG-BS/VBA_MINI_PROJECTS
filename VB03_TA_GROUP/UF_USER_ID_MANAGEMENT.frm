VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_USER_ID_MANAGEMENT 
   Caption         =   "UF_USER_ID_MANAGEMENT"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15525
   OleObjectBlob   =   "UF_USER_ID_MANAGEMENT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_USER_ID_MANAGEMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================================================================================================
'====================================...DECLARE VARIABLES
'====================================================================================================================================================
Private VAR_CREATE_OK As Boolean
Private Sub BTN_CREATE_Click()
    UF_USER_ID.Show vbModeless
End Sub

'====================================================================================================================================================
'====================================...TB_DS_DATABASE  ==> Create
'====================================================================================================================================================
Private Sub CommandButton4_Click()  'BTN Create
'====================================...Kiem tra xem có du dien kien luu file khong
    Call S_CHECK_ALL_CONDITIONS_BEFORE_CREATING_NEW_DATABASE_NAME
    If VAR_CREATE_OK = False Then
        MsgBox "Khong thanh cong!"
        Exit Sub
    End If
'====================================...Create new record
    Call S_SQL_CREATE_NEW_RECORD_TB_DS_DATABASE
'====================================...Refresh listbox
    Call CommandButton8_Click
End Sub
Sub S_CHECK_ALL_CONDITIONS_BEFORE_CREATING_NEW_DATABASE_NAME()
    VAR_CREATE_OK = False
'====================================...Kiem tra: da nhap lieu day du chu
    If TextBox1.Value = "" Then
        Label9.Caption = "Phai dien day du du lieu (*)!"
        Exit Sub
    ElseIf TextBox2.Value = "" Then
        Label9.Caption = "Phai dien day du du lieu (*)!"
        Exit Sub
    ElseIf TextBox3.Value = "" Then
        Label9.Caption = "Phai dien day du du lieu (*)!"
        Exit Sub
    End If
'====================================...Kiem tra: co bi trung ma khong
    If F_SQL_CHECK_UNIQUE_01("MA_DATABASE", "[DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]", "MA_DATABASE", TextBox1.Value, "=", "XOA_DIEU_CHINH", "OLD", "NOT LIKE") = True Then
        Label9.Caption = "Ma Database da ton tai"
        Exit Sub
    End If
    
    VAR_CREATE_OK = True
End Sub
Private Sub S_SQL_CREATE_NEW_RECORD_TB_DS_DATABASE()
'    Dim ID As Integer
    Dim MA_DATABASE As String
    Dim TEN_DATABASE As String
    Dim TEN_SERVER As String
    Dim Table_Name As String
    Dim XOA_DIEU_CHINH As String
    Dim NGAY_KHOI_TAO As Date
    
'    ID = Label6.Caption
    MA_DATABASE = TextBox1.Value
    TEN_DATABASE = TextBox2.Value
    TEN_SERVER = TextBox3.Value
    Table_Name = "TB_DS_DATABASE"
    XOA_DIEU_CHINH = ""                     'Tinh trang cua phieu: ""/OLD/CANCEL
    NGAY_KHOI_TAO = Now()
    
    Dim sql As String
'    Sql = "insert into " & TABLE_NAME & " (ID, MA_DATABASE, TEN_DATABASE, TEN_SERVER) values ('" & ID & "','" & MA_DATABASE & "','" & TEN_DATABASE & "','" & TEN_SERVER & "')"
    sql = "insert into " & Table_Name & " (MA_DATABASE, TEN_DATABASE, TEN_SERVER, XOA_DIEU_CHINH, NGAY_KHOI_TAO) values " & _
            "('" & MA_DATABASE & "','" & TEN_DATABASE & "','" & TEN_SERVER & "','" & XOA_DIEU_CHINH & "','" & NGAY_KHOI_TAO & "')"
    Call S_SQL_INSERT_INTO_01(sql)
    
    ' Thong bao hoan thanh
'    Label9.Caption = "SAVED!"
    Call S_SQL_XAC_DINH_SO_DONG_LON_NHAT
End Sub
'====================================================================================================================================================
'====================================...TB_DS_DATABASE  ==> Update
'====================================================================================================================================================
Private Sub CommandButton6_Click()  'BTN Update
'====================================...Set XOA_DIEU_CHINH = 'OLD'
'====================================...Set SQL Query
    Dim VAR_QUERY_01 As String
    Dim VAR_RECORD_CODE As String
    VAR_RECORD_CODE = TextBox1.Value
    VAR_QUERY_01 = "UPDATE " & _
                        " [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE] " & _
                    " SET " & _
                        " [XOA_DIEU_CHINH] = 'OLD' " & _
                    " WHERE " & _
                        " [MA_DATABASE] = '" & VAR_RECORD_CODE & "'; "
'====================================...Update SQL
    Call F_SQL_UPDATE_SET_OLD_01(VAR_QUERY_01)
'====================================...Create new record
    Call CommandButton4_Click
'====================================...Refresh List Box
    Call CommandButton8_Click
    UF_USER_ID_MANAGEMENT.TextBox1.SetFocus
End Sub
'====================================================================================================================================================
'====================================...TB_DS_DATABASE  ==> Deleted
'====================================================================================================================================================
Private Sub CommandButton7_Click()  'BTN DELETE
'====================================...Set XOA_DIEU_CHINH = 'OLD'
'====================================...Set SQL Query
    Dim VAR_QUERY_01 As String
    Dim VAR_RECORD_CODE As String
    VAR_RECORD_CODE = TextBox1.Value
    VAR_QUERY_01 = "UPDATE " & _
                        " [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE] " & _
                    " SET " & _
                        " [XOA_DIEU_CHINH] = 'CANCELED' " & _
                    " WHERE " & _
                        " [MA_DATABASE] = '" & VAR_RECORD_CODE & "'; "
'====================================...Update SQL
    Call F_SQL_UPDATE_SET_OLD_01(VAR_QUERY_01)
'====================================...Refresh List Box
    Call CommandButton8_Click
    UF_USER_ID_MANAGEMENT.TextBox1.SetFocus
'====================================...Thong bao
    Label9.Caption = "Da xoa thanh cong!"
End Sub
'====================================================================================================================================================
'====================================...TB_DS_DATABASE  ==> Read
'====================================================================================================================================================
Private Sub CommandButton8_Click()  'BTN Refresh
'    Call S_SQL_READ_TB_DS_DATABASE
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim Var_query As String
    Call S_OpenStatusBar    ' 0% Completed
'    VAR_QUERY = "SELECT [MA_DATABASE] as [MA DATABASE] ," & _
'                    " [TEN_DATABASE] as [TEN DATABASE] ," & _
'                    " [TEN_SERVER] as [TEN SERVER] ," & _
'                    " [XOA_DIEU_CHINH] as [XOA DIEU CHINH] ," & _
'                    " [NGAY_KHOI_TAO] as [NGAY KHOI TAO] " & _
'                " FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]"
    Var_query = "SELECT [MA_DATABASE] as [MA DATABASE] ," & _
                    " [TEN_DATABASE] as [TEN DATABASE] ," & _
                    " [TEN_SERVER] as [TEN SERVER] " & _
                " FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE]" & _
                " WHERE [XOA_DIEU_CHINH] = ''"
'====================================...Get Data to range
    Call F_SQL_SELECT_AN_ARRAY_01(Var_query, "RANGE_DATABASE_USER_ID_TB_DS_DATABASE")
    Call S_RunStatusBar(50)     ' 50% Completed
    Call S_LOAD_LISTBOX_TB_DS_DATABASE
    UF_USER_ID_MANAGEMENT.TextBox1.SetFocus
    
'    Debug.Print ActiveControl.Name
'    ActiveWindow.Activate
'    Call S_GetForegroundWindow
    Call S_RunStatusBar(100)     ' 100% Completed
    Unload UF_PROGRESS
End Sub
'Private Sub CommandButton5_Click()
'    Call S_SQL_XAC_DINH_SO_DONG_LON_NHAT
'End Sub
Private Sub S_SQL_XAC_DINH_SO_DONG_LON_NHAT()
    Dim sql As String
    sql = "SELECT COUNT(MA_DATABASE) as 'Tong_so_dong' FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE];"
    UF_USER_ID_MANAGEMENT.Label6.Caption = F_SQL_COUNT_ALL_ROWS_OF_TABLE_TEST(sql) + 1
End Sub
'    Call S_SQL_COUNT_ALL_ROWS_OF_TABLE(Sql)
'    Debug.Print F_SQL_COUNT_ALL_ROWS_OF_TABLE_TEST(Sql)
'    SH_DASHBOARD.Cells(1, 1).Value = VAR_Connection_01_recordset.Fields.Count
'    SH_DASHBOARD.Cells(2, 1).Value = VAR_Connection_01.ConnectionString
'    SH_DASHBOARD.Cells(3, 1).Value = VAR_Connection_01.State
'    SH_DASHBOARD.Cells(4, 1).Value = VAR_Connection_01.Version
'    SH_DASHBOARD.Cells(5, 1).Value = VAR_Connection_01_recordset.Fields(0).Name
'    Debug.Print VAR_Connection_01_recordset(0)
'    SH_DASHBOARD.Cells(6, 1).Value = VAR_Connection_01_recordset(0)
Private Sub S_SQL_READ_TB_DS_DATABASE()
    Dim sql As String
    sql = "SELECT MA_DATABASE, TEN_DATABASE, TEN_SERVER, FROM [DATABASE_USER_ID].[dbo].[TB_DS_DATABASE];"
    UF_USER_ID_MANAGEMENT.Label6.Caption = F_SQL_COUNT_ALL_ROWS_OF_TABLE_TEST(sql) + 1
End Sub

Private Sub CommandButton9_Click()
    Call S_START_PingSystem
End Sub

Private Sub Image2_Click()
    MultiPage_MAIN_BODY.Value = 2
End Sub

Private Sub Label10_Click()
    UF_ALL_WB.Show vbModeless
End Sub
Private Sub Label13_Click()
    MultiPage_MAIN_BODY.Value = 3
End Sub

Private Sub Label7_Click()
    MultiPage_MAIN_BODY.Value = 1
End Sub
Private Sub Label8_Click()
    MultiPage_MAIN_BODY.Value = 0
End Sub
Private Sub Label23_Click()
    MultiPage_MAIN_BODY.Value = 4
End Sub
Private Sub ListBox_TB_DS_DATABASE_Click()
    Label9.Caption = ""
'    TextBox1.Value = ListBox_TB_DS_DATABASE.Value
    TextBox1.Value = ListBox_TB_DS_DATABASE.Column(0)
    TextBox2.Value = ListBox_TB_DS_DATABASE.Column(1)
    TextBox3.Value = ListBox_TB_DS_DATABASE.Column(2)
'    Dim a As Integer
'    For a = 0 To 3
'        Controls("UF_USER_ID_MANAGEMENT.textbox" & a + 1) = ListBox_TB_DS_DATABASE.Column(a)
'    Next
End Sub


'====================================================================================================================================================
'====================================...Initialize Events
'====================================================================================================================================================
Private Sub UserForm_Initialize()
    Call Image2_Click
    Call S_FORMAT_USERFORM
    Me.MultiPage_MAIN_BODY.Value = 2
    Call S_USER_ID_MANAGEMENT_EDIT_FORM
    Call S_SET_ALL_LABEL
    Call S_LOAD_ALL_LISTBOX
    Call removeTudo(Me)
    TextBox10.Value = F_SQL_GET_SERVER_NAME_01
End Sub
Private Sub S_SET_ALL_LABEL()
    Call S_SQL_XAC_DINH_SO_DONG_LON_NHAT
End Sub
'====================================================================================================================================================
'====================================...All SUB UX/UI
'====================================================================================================================================================
Private Sub S_USER_ID_MANAGEMENT_EDIT_FORM()
    MultiPage_MAIN_BODY.Style = fmTabStyleNone
'    MultiPage_MAIN_BODY.Value = 0
End Sub
'====================================================================================================================================================
'====================================...MouseDown Events
'====================================================================================================================================================
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(Me, Me, Button)
End Sub
Private Sub MultiPage_MAIN_BODY_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(Me, Me, Button)
End Sub
'====================================================================================================================================================
'====================================...Click Events
'====================================================================================================================================================
Private Sub Image1_Click()
    Unload Me
End Sub
Private Function GetRange() As range
    ' Get the data range from the Staff worksheet
    Set GetRange = SH_KDP05_LIST.range("A1").CurrentRegion
    ' remove the header from the range by moving the range down one row and
    ' then removing the last row.
    Set GetRange = GetRange.Offset(1).Resize(GetRange.Rows.count - 1)
End Function
Private Sub UserForm_Resize()
    With MultiPage_MAIN_BODY
        .Top = 50
        .Left = 0
        .Height = Me.Height - .Top - 5
        .Width = Me.Width - .Left - 5
    End With
End Sub
'====================================================================================================================================================
'====================================...LOAD ALL LISTBOX
'====================================================================================================================================================
Private Sub S_LOAD_ALL_LISTBOX()
    Call S_LOAD_DATA_LIST_USER_ID
    Call S_LOAD_LISTBOX_TB_DS_DATABASE
    Call S_LOAD_LISTBOX_01
    Call S_LOAD_LISTBOX_02
End Sub
Private Sub S_LOAD_DATA_LIST_USER_ID()
    ' Get the data range
    Dim rg As range
    Set rg = GetRange
    ' Link the data to the ListBox
    With ListBox_USER_ID
'        .RowSource = rg.Address(External:=True)
        .RowSource = "TB_BANG_USER_ID"
        .ColumnCount = rg.Columns.count
'        Debug.Print rg.Columns.Count
        .ColumnWidths = "100;100;100"
        .ColumnHeads = True
        .Font.Size = 14
        .ListIndex = 0
        .Left = 5
        .Top = 5
    End With
End Sub
Private Sub S_LOAD_LISTBOX_TB_DS_DATABASE()
    ' Get the data range
    Dim rg As range
'    Set rg = GetRange_LISTBOX_TB_DS_DATABASE
    Set rg = F_GetRange_To_LISTBOX("RANGE_DATABASE_USER_ID_TB_DS_DATABASE")
'    Debug.Print rg.Address & "<<rg"
    ' Link the data to the ListBox
    With ListBox_TB_DS_DATABASE
'        .RowSource = "TB_DS_DATABASE"
        .RowSource = rg.Address(External:=True)
'        .RowSource = "RANGE_DATABASE_USER_ID_TB_DS_DATABASE"
        .ColumnCount = rg.Columns.count
        .ColumnWidths = "100, 150, 150, 150, 150"
        .ColumnHeads = True
'        .ColumnHeads = False
        .Font.Size = 12
        .ListIndex = 0
        .Height = 350
        .Width = 400
    End With
End Sub
Private Sub S_LOAD_LISTBOX_01()
    ' Get the data range
    Dim rg As range
'    Set rg = GetRange_LISTBOX_TB_DS_DATABASE
    Set rg = F_GetRange_To_LISTBOX("RANGE_ALL_TABLE_IN_DATASE")
'    Debug.Print rg.Address & "<<rg"
    ' Link the data to the ListBox
    With UF_USER_ID_MANAGEMENT.ListBox1
'        .RowSource = "TB_DS_DATABASE"
        .RowSource = rg.Address(External:=True)
'        .RowSource = "RANGE_DATABASE_USER_ID_TB_DS_DATABASE"
        .ColumnCount = rg.Columns.count
        .ColumnWidths = "180"
        .ColumnHeads = True
'        .ColumnHeads = False
        .Font.Size = 12
        .ListIndex = 0
        .Height = 180
        .Width = 200
    End With
End Sub
Private Sub S_LOAD_LISTBOX_02()
    ' Get the data range
    Dim rg As range
'    Set rg = GetRange_LISTBOX_TB_DS_DATABASE
    Set rg = F_GetRange_To_LISTBOX("RANGE_DATABASE_USER_ID_TB_DS_DATABASE")
'    Debug.Print rg.Address & "<<rg"
    ' Link the data to the ListBox
    With UF_USER_ID_MANAGEMENT.ListBox2
'        .RowSource = "TB_DS_DATABASE"
        .RowSource = rg.Address(External:=True)
'        .RowSource = "RANGE_DATABASE_USER_ID_TB_DS_DATABASE"
        .ColumnCount = rg.Columns.count
        .ColumnWidths = "50,180,180"
        .ColumnHeads = True
'        .ColumnHeads = False
        .Font.Size = 12
        .ListIndex = 0
        .Height = 180
        .Width = 200
    End With
End Sub
'====================================================================================================================================================
'====================================...FORMAT
'====================================================================================================================================================
Private Sub S_FORMAT_USERFORM()
    Me.Height = 530
    Me.Width = 900
End Sub
'====================================================================================================================================================
'====================================...List Box
'====================================================================================================================================================
Private Sub ListBox1_Click()
    UF_USER_ID_MANAGEMENT.TextBox9 = UF_USER_ID_MANAGEMENT.ListBox1.Value
End Sub

Private Sub ListBox2_Click()
'    UF_USER_ID_MANAGEMENT.TextBox8 = UF_USER_ID_MANAGEMENT.ListBox2.Value
    UF_USER_ID_MANAGEMENT.TextBox8 = UF_USER_ID_MANAGEMENT.ListBox2.Column(1)
End Sub
'====================================================================================================================================================
'====================================...Get Range to ListBox By Function
'====================================================================================================================================================
Function F_GetRange_To_LISTBOX(RangeNameString As String) As range
    ' Get the data range by Function
    Set F_GetRange_To_LISTBOX = F_FIND_RANGE_IN_SH_ALL_RANGES_01(RangeNameString)
'    Debug.Print F_GetRange_To_LISTBOX.Address

    ' remove the header from the range by moving the range down one row and then removing the last row.
    'Luu y: Vung nay phai co nhieu hon 1 dong thi moi Offset(1) duoc
    Set F_GetRange_To_LISTBOX = F_GetRange_To_LISTBOX.Offset(1).Resize(F_GetRange_To_LISTBOX.Rows.count - 1)
'    Debug.Print F_GetRange_To_LISTBOX.Address
End Function
'====================================================================================================================================================
'====================================...Page 4 - PING IP
'====================================================================================================================================================
'====================================...https://www.youtube.com/watch?v=JxueKaNJ89s
Function Ping(strip)
    Dim objshell, boolcode
    Set objshell = CreateObject("Wscript.Shell")
    boolcode = objshell.Run("ping -n 1 -w 1000 " & strip, 0, True)
    If boolcode = 0 Then
        Ping = True
    Else
        Ping = False
    End If
End Function
Sub S_START_PingSystem()
    Dim strip As String
    Dim introw As Long
    ' Test First IP
    strip = TextBox5.Value
    If Ping(strip) = True Then
        Label16.ForeColor = RGB(0, 142, 0)
        Label16.Caption = "Status: online"
    Else
        Label16.ForeColor = RGB(192, 0, 0)
        Label16.Caption = "Status: offline"
    End If
    ' Test second IP
    strip = TextBox4.Value
    If Ping(strip) = True Then
        Label14.ForeColor = RGB(0, 142, 0)
        Label14.Caption = "Status: online"
    Else
        Label14.ForeColor = RGB(192, 0, 0)
        Label14.Caption = "Status: offline"
    End If
    ' Test Third IP
    strip = TextBox7.Value
    If Ping(strip) = True Then
        Label20.ForeColor = RGB(0, 142, 0)
        Label20.Caption = "Status: online"
    Else
        Label20.ForeColor = RGB(192, 0, 0)
        Label20.Caption = "Status: offline"
    End If
    ' Test Fourth IP
    strip = TextBox6.Value
    If Ping(strip) = True Then
        Label17.ForeColor = RGB(0, 142, 0)
        Label17.Caption = "Status: online"
    Else
        Label17.ForeColor = RGB(192, 0, 0)
        Label17.Caption = "Status: offline"
    End If
End Sub
'Sub S_STOP_PingSystem()
'    SH_PING.Range("F1").Value = "STOP"
'End Sub

'====================================================================================================================================================
'====================================...Page 4 - Export, Import, Delete all Record
'====================================================================================================================================================
Private Sub CommandButton13_Click() 'Liet Ke cac bang trong Database
    Dim Var_query As String
    Dim DestRange As String
    Dim ServerName As String
    Dim DatabaseName As String
    Dim LoginName As String
    Dim LoginPass As String
'====================================...Set query and destiny range
    Var_query = "SELECT name as ALL_TABLE_IN_DATASE" & _
                " FROM sys.Tables;"
    DestRange = "RANGE_ALL_TABLE_IN_DATASE"
'====================================...Set Login informations
    ServerName = UF_USER_ID_MANAGEMENT.TextBox10.Value
    DatabaseName = UF_USER_ID_MANAGEMENT.TextBox8.Value
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
'    Call F_SQL_SELECT_AN_ARRAY_01(VAR_QUERY, "RANGE_ALL_TABLE_IN_DATASE")
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_RANGE(Var_query, DestRange, ServerName, DatabaseName, LoginName, LoginPass)
End Sub
Private Sub CommandButton10_Click() 'Ket xuat tat ca Records ra Excel
'    SH_DATA_EXPORT.Cells.ClearContents
    SH_DATA_EXPORT.Cells.Delete Shift:=xlUp
    Dim Var_query As String
    Dim DestSheet As String
    Dim ServerName As String
    Dim DatabaseName As String
    Dim LoginName As String
    Dim LoginPass As String
    Dim Table_Name As String
    Table_Name = UF_USER_ID_MANAGEMENT.TextBox9.Value
'====================================...Set query and destiny range
    Var_query = "SELECT *" & _
                " FROM " & Table_Name & _
                " Order BY " & _
                "[NGAY_KHOI_TAO]"
                
    DestSheet = "SH_DATA_EXPORT"
'====================================...Set Login informations
    ServerName = UF_USER_ID_MANAGEMENT.TextBox10.Value
    DatabaseName = UF_USER_ID_MANAGEMENT.TextBox8.Value
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    Call F_SQL_SELECT_AN_ARRAY_ALL_ARGUMENTS_TO_SHEET(Var_query, DestSheet, ServerName, DatabaseName, LoginName, LoginPass)
    SH_DATA_EXPORT.Cells.EntireColumn.AutoFit
    SH_DATA_EXPORT.Copy
End Sub
Private Sub CommandButton14_Click() 'Xoa tat ca cac Records
    Dim retype_pass As String
    Dim Var_query As String
    Dim ServerName As String
    Dim DatabaseName As String
    Dim LoginName As String
    Dim LoginPass As String
    Dim Table_Name As String
    Dim ThongBao As String
    Table_Name = UF_USER_ID_MANAGEMENT.TextBox9.Value
'====================================...Set query and destiny range
    Var_query = "DELETE" & _
                " FROM " & Table_Name
'====================================...Set Login informations
    ServerName = UF_USER_ID_MANAGEMENT.TextBox10.Value
    DatabaseName = UF_USER_ID_MANAGEMENT.TextBox8.Value
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
    ThongBao = "Xac nhan xoa tat ca du lieu:" & Chr(13) & _
                "Ten Database: " & DatabaseName & Chr(13) & _
                "Ten bang: " & Table_Name & Chr(13) & _
                Chr(13) & _
                Chr(13) & _
                Chr(13) & _
                "Vui long nhap lai mat khau"
    retype_pass = InputBox(ThongBao, "Xac nhan lai mat khau!")
    If retype_pass <> LoginPass Then
        MsgBox "Mat khau khong dung!"
        Exit Sub
    End If
    
    Call F_SQL_DELETE_ALL_RECORD_FROM_TABLE(Var_query, ServerName, DatabaseName, LoginName, LoginPass)
    MsgBox "Hoan thanh xoa bang"
End Sub
Private Sub CommandButton15_Click() 'Import Array to SQL
'====================================...Get Data From SQL to Excel Range
'====================================...Set SQL Query
    Dim ServerName As String
    Dim DatabaseName As String
    Dim TableName As String
    Dim LoginName As String
    Dim LoginPass As String

    ServerName = UF_USER_ID_MANAGEMENT.TextBox10
    DatabaseName = UF_USER_ID_MANAGEMENT.TextBox8
    TableName = UF_USER_ID_MANAGEMENT.TextBox9
    LoginName = F_SQL_GET_LOGIN_NAME_01
    LoginPass = F_SQL_GET_LOGIN_PASS_01
    
'====================================...Get Data to range
    Call F_IMPORT_INTO_DATABASE_USER_ID_FROM_SH_DATA_IMPORT(ServerName, DatabaseName, TableName, LoginName, LoginPass)
'====================================...Clear contents
    SH_DATA_IMPORT.Cells.ClearContents
End Sub
Private Sub CommandButton16_Click() 'Get File
    Dim fd As Office.FileDialog
    Dim strFile As String
    Dim Full_Name As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
    With fd
       .Filters.Clear
       .Filters.Add "Excel Files", "*.xlsx?", 1
       .Title = "Choose an Excel file"
       .AllowMultiSelect = False
'       .InitialFileName = ThisWorkbook.path
       If .Show = True Then
           strFile = .SelectedItems(1)
       End If
    End With
    
    'Xu ly neu khong có file nào duoc chon
    If strFile = "" Then
        Exit Sub
    End If
    
    Full_Name = getFileNameFromPath(strFile)
    UF_USER_ID_MANAGEMENT.Label22.Caption = Full_Name
    
    Dim WB As Workbook
    Dim WB_index As Integer
    Dim WB_open As Boolean
    
    ' Kiem tra xem file có dang mo khong
    WB_open = False
    For WB_index = Workbooks.count To 1 Step -1
        If Workbooks(WB_index).FullName = strFile Then
            WB_open = True
            Exit For
        End If
    Next WB_index
        
    If WB_open = True Then
        GoTo XU_LY_FILE_DANG_MO
    Else
        GoTo XU_LY_FILE_DANG_DONG
    End If
    
XU_LY_FILE_DANG_MO:
'    Debug.Print "DANG MO"
    On Error Resume Next
    Set WB = Workbooks(WB_index)
    If WB Is Nothing Then
'    If WB.pthe Is Nothing Then
        Set WB = GetWorkbook(strFile)
        Debug.Print WB.path & "path"
'        Set WB = Workbooks.Open(sFullFilename)
    End If
    On Error GoTo 0
'    Set WB = GetWorkbook(strFile)
    WB.Sheets("IMPORT").range("A1").CurrentRegion.Copy Destination:=SH_DATA_IMPORT.range("A1")
'    WB.Close False
    
    UF_USER_ID_MANAGEMENT.Label26.Caption = "Da lay data xong"
    Exit Sub

XU_LY_FILE_DANG_DONG:
'    Debug.Print "DANG DONG"
    On Error Resume Next
    Application.ScreenUpdating = False
'    Set WB = Workbooks(WB_index)
    If WB Is Nothing Then
'    If WB.pthe Is Nothing Then
        Set WB = GetWorkbook(strFile)
'        Debug.Print WB.path & "path"
'        Set WB = Workbooks.Open(sFullFilename)
    End If
    On Error GoTo 0
'    Set WB = GetWorkbook(strFile)
    WB.Sheets("IMPORT").range("A1").CurrentRegion.Copy Destination:=SH_DATA_IMPORT.range("A1")
    WB.Close False
    Application.ScreenUpdating = True
    UF_USER_ID_MANAGEMENT.Label26.Caption = "Da lay data xong"
    Exit Sub
    
End Sub
Function getFileNameFromPath(path)
    Dim fileName As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    getFileNameFromPath = fso.GetFilename(path)
End Function
Function GetWorkbook(ByVal sFullFilename As String) As Workbook
    
    Dim sFilename As String
    sFilename = Dir(sFullFilename)
    
    On Error Resume Next
    Dim wk As Workbook
    Set wk = Workbooks(sFilename)
    
    If wk Is Nothing Then
        Set wk = Workbooks.Open(sFullFilename)
    End If
    
    On Error GoTo 0
    Set GetWorkbook = wk
    
End Function
Private Sub CommandButton17_Click() 'Get Template
    SH_DATA_IMPORT.Cells.ClearContents
    SH_DATA_IMPORT.Copy
End Sub

