VERSION 5.00
Begin VB.Form frmControls 
   Caption         =   "Form for Popup Menu"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   3945
   Begin VB.Menu mnuPopup1 
      Caption         =   "Popup1"
      Begin VB.Menu mnuSub1 
         Caption         =   "Sub1"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event SubMenuClick(ByVal idx As Integer)


Private Sub mnuSub1_Click(Index As Integer)
    RaiseEvent SubMenuClick(Index)
End Sub
