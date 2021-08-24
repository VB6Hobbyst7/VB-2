VERSION 5.00
Begin VB.Form frmBBS503 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ABO결과 수정"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   9960
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmBBS503"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objFormModifyNo As clsFormModifyNo
Attribute objFormModifyNo.VB_VarHelpID = -1



Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Set objFormModifyNo = New clsFormModifyNo
    Call objFormModifyNo.FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objFormModifyNo = Nothing
End Sub

Private Sub objFormModifyNo_FormClose()
    Unload Me
End Sub
