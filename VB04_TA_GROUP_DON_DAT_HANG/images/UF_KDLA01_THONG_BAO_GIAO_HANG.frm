VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_KDLA01_THONG_BAO_GIAO_HANG 
   Caption         =   "UF_KDLA01_THONG_BAO_GIAO_HANG"
   ClientHeight    =   10800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16185
   OleObjectBlob   =   "UF_KDLA01_THONG_BAO_GIAO_HANG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_KDLA01_THONG_BAO_GIAO_HANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'====================================================================================================================================================
'====================================...UserForm_Initialize
'====================================================================================================================================================
Private Sub UserForm_Initialize()
    Call SUB_MAKE_LABEL_TRANSPARENT
End Sub

Sub SUB_MAKE_LABEL_TRANSPARENT()
    ' make all labels transparent
    Dim ctrl As Object
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.Label Then
            ctrl.BackStyle = 0 - fmBackStyleTransparent
        End If
    Next ctrl
End Sub
