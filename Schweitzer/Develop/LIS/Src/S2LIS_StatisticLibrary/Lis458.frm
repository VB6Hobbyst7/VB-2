VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm458Infection 
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14385
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7605
      TabIndex        =   22
      Top             =   8370
      Width           =   1650
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9390
      TabIndex        =   21
      Top             =   8370
      Width           =   1425
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12420
      TabIndex        =   20
      Top             =   8385
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10905
      TabIndex        =   19
      Top             =   8370
      Width           =   1410
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "To Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12435
      TabIndex        =   18
      Top             =   2040
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7845
      TabIndex        =   15
      Text            =   "Feb 28 1999"
      Top             =   210
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5700
      TabIndex        =   14
      Text            =   "Feb 01 1999"
      Top             =   210
      Width           =   1425
   End
   Begin VB.CommandButton cmdSelDT1 
      Caption         =   "..."
      Height          =   360
      Left            =   7155
      TabIndex        =   13
      Top             =   210
      Width           =   330
   End
   Begin VB.CommandButton cmdSelDT2 
      Caption         =   "..."
      Height          =   360
      Left            =   9300
      TabIndex        =   12
      Top             =   210
      Width           =   330
   End
   Begin VB.Frame fraImpression 
      Height          =   1725
      Left            =   540
      TabIndex        =   0
      Top             =   810
      Width           =   11475
      Begin VB.ListBox List2 
         Height          =   1140
         Left            =   360
         TabIndex        =   11
         Top             =   345
         Width           =   1665
      End
      Begin VB.ListBox List1 
         Height          =   1140
         Left            =   2340
         TabIndex        =   3
         Top             =   345
         Width           =   8865
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   495
      TabIndex        =   1
      Top             =   7785
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin FPSpread.vaSpread ssInfect 
      Height          =   4110
      Left            =   525
      TabIndex        =   2
      Top             =   2835
      Width           =   13215
      _Version        =   196608
      _ExtentX        =   23310
      _ExtentY        =   7250
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   5
      Protect         =   0   'False
      ShadowColor     =   12632256
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis458.frx":0000
      VisibleCols     =   7
      VisibleRows     =   2
   End
   Begin VB.Label Label11 
      Caption         =   "-"
      Height          =   240
      Left            =   7605
      TabIndex        =   17
      Top             =   300
      Width           =   270
   End
   Begin VB.Label lblDuration 
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4710
      TabIndex        =   16
      Top             =   255
      Width           =   870
   End
   Begin VB.Label Label7 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10650
      TabIndex        =   10
      Top             =   7230
      Width           =   630
   End
   Begin VB.Label lblPercent 
      Caption         =   "Percentage :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9150
      TabIndex        =   9
      Top             =   7230
      Width           =   1485
   End
   Begin VB.Label Label5 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7590
      TabIndex        =   8
      Top             =   7245
      Width           =   630
   End
   Begin VB.Label lblICount 
      Caption         =   "Impression Count :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5355
      TabIndex        =   7
      Top             =   7245
      Width           =   2130
   End
   Begin VB.Label Label3 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3615
      TabIndex        =   6
      Top             =   7260
      Width           =   630
   End
   Begin VB.Label lblTCount 
      Caption         =   "Total Count :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   5
      Top             =   7230
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Now Executing ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   630
      TabIndex        =   4
      Top             =   8400
      Width           =   5910
   End
End
Attribute VB_Name = "frm458Infection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub


