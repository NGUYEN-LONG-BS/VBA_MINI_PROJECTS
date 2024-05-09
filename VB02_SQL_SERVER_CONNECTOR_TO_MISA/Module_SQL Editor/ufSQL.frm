VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSQL 
   Caption         =   "SQL in Excel by DTNguyen | Hoc Excel Online"
   ClientHeight    =   7189
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   10620
   OleObjectBlob   =   "ufSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDownFontSize_Click()
    Dim current As Long
    current = Me.textboxSQL.Font.Size
    Me.textboxSQL.Font.Size = IIf(current < 8, 8, current - 1)
End Sub

Private Sub btnGo_Click()
    Dim sql_string As String
    sql_string = Me.textboxSQL.Text
    CreateSQLQuery sql_string
End Sub

Private Sub btnReset_Click()
    DeleteAllButMenuSheet
    Me.textboxSQL.Font.Size = 20
End Sub

Private Sub btnUpFontSize_Click()
    Dim current As Long
    current = Me.textboxSQL.Font.Size
    Me.textboxSQL.Font.Size = IIf(current > 100, 100, current + 1)
End Sub
