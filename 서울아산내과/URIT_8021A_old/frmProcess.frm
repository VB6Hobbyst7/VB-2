VERSION 5.00
Begin VB.Form frmProcess 
   Caption         =   "FrmProcess"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If ProcessCHK("IF_U411.exe") = True Then
        MsgBox "인터페이스 프로그램이 실행중입니다.", vbOKOnly, "경고"
    Else
        Shell App.Path & "\IF_U411.exe", vbNormalFocus
    End If
    
    Unload Me
End Sub
