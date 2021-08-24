VERSION 5.00
Begin VB.Form frmBBS205 
   BackColor       =   &H00DBE6E6&
   Caption         =   "병동/외래 Barcode Label 재출력"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   Icon            =   "frmBBS205.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1845
   ScaleWidth      =   3600
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmBBS205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objLisCollectForm As clsLisCollectForm
Attribute objLisCollectForm.VB_VarHelpID = -1

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Me.Show
    Set objLisCollectForm = New clsLisCollectForm
    Call objLisCollectForm.CollectionButtonClick("LIS213", Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objLisCollectForm = Nothing
End Sub

Private Sub objLisCollectForm_LastFormUnload()
    Unload Me
End Sub
