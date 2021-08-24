VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
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
   Begin MSCommLib.MSComm MyComm 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuSub 
         Caption         =   "Sub"
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "Sub1"
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
