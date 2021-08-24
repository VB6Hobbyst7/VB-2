VERSION 5.00
Begin VB.Form frmBBS305 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Ç÷¾×¹ÝÈ¯ ¹× Æó±â"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   5415
   WindowState     =   2  'ÃÖ´ëÈ­
End
Attribute VB_Name = "frmBBS305"
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
    f.mode = 2
    f.Show
    f.ZOrder
    
    Me.WindowState = 1
End Sub

