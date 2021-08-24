VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form4 
   Caption         =   "Muti Row Sync Call"
   ClientHeight    =   6570
   ClientLeft      =   3165
   ClientTop       =   2415
   ClientWidth     =   9645
   LinkTopic       =   "Form4"
   ScaleHeight     =   6570
   ScaleWidth      =   9645
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   9735
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4935
      Left            =   360
      OleObjectBlob   =   "Sync'.frx":0000
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
