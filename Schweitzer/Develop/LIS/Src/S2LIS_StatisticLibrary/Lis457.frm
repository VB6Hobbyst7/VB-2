VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm457Epidemic 
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
      Left            =   12690
      TabIndex        =   18
      Top             =   2670
      Width           =   1245
   End
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
      Left            =   7845
      TabIndex        =   17
      Top             =   8295
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
      Left            =   9630
      TabIndex        =   16
      Top             =   8295
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
      Left            =   12660
      TabIndex        =   15
      Top             =   8310
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
      Left            =   11145
      TabIndex        =   14
      Top             =   8295
      Width           =   1410
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
      Left            =   7695
      TabIndex        =   11
      Text            =   "Feb 28 1999"
      Top             =   240
      Width           =   1425
   End
   Begin VB.TextBox Text6 
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
      Left            =   5550
      TabIndex        =   10
      Text            =   "Feb 01 1999"
      Top             =   240
      Width           =   1425
   End
   Begin VB.CommandButton cmdSelDT1 
      Caption         =   "..."
      Height          =   360
      Left            =   7005
      TabIndex        =   9
      Top             =   225
      Width           =   330
   End
   Begin VB.CommandButton cmdSelDT2 
      Caption         =   "..."
      Height          =   360
      Left            =   9150
      TabIndex        =   8
      Top             =   240
      Width           =   330
   End
   Begin VB.Frame fraSpecies 
      Caption         =   "Species Three"
      Height          =   2085
      Index           =   2
      Left            =   8640
      TabIndex        =   6
      Top             =   1005
      Width           =   3480
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "Lis457.frx":0000
         Left            =   285
         List            =   "Lis457.frx":0013
         TabIndex        =   7
         Top             =   345
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpecies 
      Caption         =   "Species Two"
      Height          =   2085
      Index           =   1
      Left            =   4665
      TabIndex        =   4
      Top             =   1020
      Width           =   3690
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "Lis457.frx":0071
         Left            =   315
         List            =   "Lis457.frx":0093
         TabIndex        =   5
         Top             =   330
         Width           =   3075
      End
   End
   Begin VB.Frame fraSpecies 
      Caption         =   "Species One"
      Height          =   2070
      Index           =   0
      Left            =   660
      TabIndex        =   2
      Top             =   1020
      Width           =   3780
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "Lis457.frx":0123
         Left            =   270
         List            =   "Lis457.frx":0139
         TabIndex        =   3
         Top             =   345
         Width           =   3135
      End
   End
   Begin FPSpread.vaSpread ssInfect 
      Height          =   3930
      Left            =   630
      TabIndex        =   1
      Top             =   3405
      Width           =   13395
      _Version        =   196608
      _ExtentX        =   23627
      _ExtentY        =   6932
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
      MaxCols         =   8
      MaxRows         =   20
      Protect         =   0   'False
      ShadowColor     =   12632256
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis457.frx":01A3
      VisibleCols     =   7
      VisibleRows     =   2
   End
   Begin VB.Label Label11 
      Caption         =   "-"
      Height          =   240
      Left            =   7455
      TabIndex        =   13
      Top             =   330
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
      Left            =   4560
      TabIndex        =   12
      Top             =   285
      Width           =   870
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
      Left            =   690
      TabIndex        =   0
      Top             =   8295
      Width           =   5910
   End
End
Attribute VB_Name = "frm457Epidemic"
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

