VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmControls 
   Caption         =   "Form for Popup Menu"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
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
   Begin RichTextLib.RichTextBox rtfTextBox 
      Height          =   1965
      Left            =   90
      TabIndex        =   2
      Top             =   3330
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   3466
      _Version        =   393217
      TextRTF         =   $"frmControls.frx":0000
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   3240
      Top             =   1605
      Width           =   360
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
