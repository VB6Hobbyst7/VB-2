VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmControls 
   Caption         =   "Form for Popup Menu"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin RichTextLib.RichTextBox rtfTempText 
      Height          =   750
      Left            =   2175
      TabIndex        =   3
      Top             =   2910
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmControls.frx":0000
   End
   Begin RichTextLib.RichTextBox rtfTextBox 
      Height          =   750
      Left            =   300
      TabIndex        =   2
      Top             =   2880
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmControls.frx":0220
   End
   Begin VB.ListBox lstUnsortedList 
      Height          =   2400
      Left            =   810
      TabIndex        =   1
      Top             =   660
      Width           =   2100
   End
   Begin VB.ListBox lstList 
      Height          =   2400
      Left            =   195
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   2100
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuSub 
         Caption         =   "Sub"
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
