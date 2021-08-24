VERSION 5.00
Begin VB.Form frmBBS304 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액반환"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   3885
   ScaleMode       =   0  '사용자
   ScaleWidth      =   16956.59
   WindowState     =   2  '최대화
End
Attribute VB_Name = "frmBBS304"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private f As New frmBBS303

Private Sub Form_Activate()
    Unload Me
    medMain.ZOrder
End Sub

Private Sub Form_Load()
    f.mode = 1
    f.Show
    f.ZOrder
    
    Me.WindowState = 1
End Sub
