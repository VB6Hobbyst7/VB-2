VERSION 5.00
Begin VB.Form frmBBS504 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ABO결과 조회"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14580
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmBBS504"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objFormViewResult As clsFormViewResult
Attribute objFormViewResult.VB_VarHelpID = -1



Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Set objFormViewResult = New clsFormViewResult
    Call objFormViewResult.FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objFormViewResult = Nothing
End Sub

Private Sub objFormViewResult_FormClose()
    Unload Me
End Sub
