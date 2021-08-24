VERSION 5.00
Begin VB.Form frmIIS101 
   Caption         =   "접 수"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmIIS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Me.WindowState = vbMaximized
    mdiIISMain.lblMenuNm = Me.Caption
End Sub

Private Sub Form_Deactivate()
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIIS101 = Nothing
End Sub


