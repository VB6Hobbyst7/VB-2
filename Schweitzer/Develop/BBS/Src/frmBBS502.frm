VERSION 5.00
Begin VB.Form frmBBS502 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ABO결과 개별등록"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmBBS502.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   9975
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmBBS502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objFormEntryNo As clsFormEntryNO
Attribute objFormEntryNo.VB_VarHelpID = -1



Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Set objFormEntryNo = New clsFormEntryNO
    Call objFormEntryNo.FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objFormEntryNo = Nothing
End Sub

Private Sub objFormEntryNo_FormClose()
    Unload Me
End Sub

