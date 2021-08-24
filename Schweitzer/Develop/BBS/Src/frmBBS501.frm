VERSION 5.00
Begin VB.Form frmBBS501 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ABO결과 일괄등록"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   Icon            =   "frmBBS501.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14580
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmBBS501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objFormEntryWS As clsFormEntryWS
Attribute objFormEntryWS.VB_VarHelpID = -1

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Set objFormEntryWS = New clsFormEntryWS
    Call objFormEntryWS.FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objFormEntryWS = Nothing
End Sub

Private Sub objFormEntryWS_FormClose()
    Unload Me
End Sub
