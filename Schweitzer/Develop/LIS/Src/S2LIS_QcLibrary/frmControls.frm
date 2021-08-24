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
   Begin VB.ListBox lstUnsortedList 
      Height          =   2400
      Left            =   810
      TabIndex        =   1
      Top             =   660
      Width           =   2100
   End
   Begin VB.ListBox lstList 
      Height          =   2400
      Left            =   270
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2100
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuSub1 
         Caption         =   "Sub1"
      End
      Begin VB.Menu mnuSub2 
         Caption         =   "Sub2"
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
