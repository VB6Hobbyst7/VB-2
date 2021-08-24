VERSION 5.00
Begin VB.Form frmControl 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.ListBox lstSortedList 
      Height          =   2400
      Left            =   570
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   390
      Width           =   2145
   End
   Begin VB.ListBox lstUnsortedList 
      Height          =   2580
      Left            =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   2025
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuSub 
         Caption         =   "Sub"
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
