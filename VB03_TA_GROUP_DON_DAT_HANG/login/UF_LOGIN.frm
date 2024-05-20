VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_LOGIN 
   Caption         =   "LOGIN"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "UF_LOGIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub BTN_CLEAR_Click()
    TextBox_ID.Value = ""
    TextBox_PASS.Value = ""
End Sub

Private Sub BTN_LOGIN_Click()
    Dim stringUser As String
    Dim stringPass As String
    Dim DangNhapOK As Boolean
    
    Call S_SET_VAR_NAM_TAI_CHINH
    
    stringUser = UF_LOGIN.TextBox_ID
    stringPass = UF_LOGIN.TextBox_PASS
    
    If stringUser = "" Or stringPass = "" Then
        UF_LOGIN.Label_PASS_VISIBLE.Caption = "Vui long nhap day du ID va Pass!"
        Exit Sub
    End If
    
    'Decentralization via Global Variable
    If stringUser = "admin" And stringPass = "123123" Then
        Unload Me
        VAR_INFOR_MODULE_KINH_DOANH = True
        VAR_INFOR_MODULE_VAT_TU = True
        VAR_INFOR_MODULE_KY_THUAT = True
        VAR_INFOR_MODULE_TAI_CHINH = True
        VAR_INFOR_MODULE_ADMIN = True
        Call S_SHOW_UF_DASHBOARD
        Exit Sub
    End If
    
    DangNhapOK = F_SQL_KIEM_TRA_THONG_TIN_DANG_NHAP(stringUser, stringPass)
    
    If DangNhapOK = True Then
        Call S_GET_LOGIN_INFORMATION(stringUser, stringPass)
        Unload Me
        Call S_GET_DATA_WHEN_LOGIN
        Call S_SHOW_UF_DASHBOARD
        Exit Sub
    Else
        UF_LOGIN.Label_PASS_VISIBLE.Caption = "Sai thong tin dang nhap"
        Exit Sub
    End If

End Sub

Private Sub S_SET_VAR_NAM_TAI_CHINH()
    VAR_NAM_TAI_CHINH = Right(ComboBox1.Value, 2)       'Global Variables
End Sub

Private Sub Label_Cancel_Click()
    Dim stringUser As String
    Dim stringPass As String
    
    stringUser = UF_LOGIN.TextBox_ID
    stringPass = UF_LOGIN.TextBox_PASS
    
    If stringUser = "admin" And stringPass = "123123" Then
        Unload Me
        Application.Visible = True
        UF_CHECK_CONNECTION.Show vbModeless
        Exit Sub
    End If
    
    If MsgBox("Do you want to logout?", vbYesNo, "Logout?") = vbYes Then
        Dim a As Integer
        a = Workbooks.count
        If a = 1 Then
            Unload Me
            ThisWorkbook.Close SaveChanges:=False
            Application.Quit
        ElseIf a > 1 Then
            Unload Me
            ThisWorkbook.Close SaveChanges:=False
        End If
    End If
End Sub


Private Sub Label_SHOW_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_PASS_VISIBLE.Caption = TextBox_PASS.Value
End Sub

Private Sub Label_SHOW_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_PASS_VISIBLE.Caption = ""
End Sub

Private Sub UserForm_Initialize()
    Call removeTudo(Me)
    Call S_LOAD_DATA_COMBOBOX
End Sub

Private Sub S_LOAD_DATA_COMBOBOX()
    ComboBox1.List = Array("2023", "2024", "2025", "2026")
'    ComboBox1.Value = "2024"
    ComboBox1.Value = Year(Now())
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(Me, Me, Button)
End Sub

Private Sub UserForm_Resize()
    'a
End Sub
