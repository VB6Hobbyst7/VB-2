VERSION 5.00
Begin VB.Form frmBBS103 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Ward Collection"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   Icon            =   "frmBBS103N.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   2940
   WindowState     =   2  '√÷¥Î»≠
End
Attribute VB_Name = "frmBBS103"
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
    Me.WindowState = 2
    Set objLisCollectForm = New clsLisCollectForm
    objLisCollectForm.EmpId = ObjSysInfo.EmpId
    Call objLisCollectForm.CollectionButtonClick("LIS214", Me)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objLisCollectForm = Nothing
End Sub

Private Sub objLisCollectForm_LastFormUnload()
    Unload Me
End Sub
