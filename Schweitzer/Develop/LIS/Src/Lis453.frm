VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSWL 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Work Load"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14385
   Tag             =   "45300"
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   7545
      Left            =   3210
      ScaleHeight     =   7485
      ScaleWidth      =   10785
      TabIndex        =   10
      Top             =   780
      Width           =   10845
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00FCEFE9&
         Caption         =   "To Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9315
         Style           =   1  '그래픽
         TabIndex        =   17
         Tag             =   "127"
         Top             =   300
         Width           =   1260
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   45
         Width           =   8790
         Begin VB.OptionButton optItem 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Item"
            Height          =   300
            Left            =   360
            TabIndex        =   14
            Tag             =   "45304"
            Top             =   285
            Width           =   1245
         End
         Begin VB.OptionButton optPerson 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Person"
            Height          =   300
            Left            =   1785
            TabIndex        =   13
            Tag             =   "45305"
            Top             =   285
            Width           =   1245
         End
      End
      Begin FPSpread.vaSpread ssWL 
         Height          =   6510
         Left            =   135
         TabIndex        =   11
         Top             =   855
         Width           =   10485
         _Version        =   196608
         _ExtentX        =   18494
         _ExtentY        =   11483
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
         MaxCols         =   9
         MaxRows         =   20
         Protect         =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis453.frx":0000
         VisibleCols     =   9
         VisibleRows     =   10
      End
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
      Left            =   7935
      TabIndex        =   7
      Text            =   "Feb 28 1999"
      Top             =   195
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
      Left            =   5790
      TabIndex        =   6
      Text            =   "Feb 01 1999"
      Top             =   195
      Width           =   1425
   End
   Begin VB.CommandButton cmdSelDT1 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   360
      Left            =   7245
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   195
      Width           =   330
   End
   Begin VB.CommandButton cmdSelDT2 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   360
      Left            =   9390
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   195
      Width           =   330
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11085
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "132"
      Top             =   8445
      Width           =   1410
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12600
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8475
      Width           =   1245
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Stop Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9570
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "162"
      Top             =   8445
      Width           =   1425
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Start Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7785
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "158"
      Top             =   8445
      Width           =   1650
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6945
      Left            =   195
      TabIndex        =   15
      Top             =   1365
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   12250
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.TabStrip tsSec 
      Height          =   375
      Left            =   150
      TabIndex        =   16
      Tag             =   "45306"
      Top             =   855
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   661
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Process"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Individual"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      Height          =   240
      Left            =   7695
      TabIndex        =   9
      Top             =   285
      Width           =   270
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H00DBE6E6&
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
      Left            =   4800
      TabIndex        =   8
      Tag             =   "45301"
      Top             =   240
      Width           =   870
   End
End
Attribute VB_Name = "frmSWL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub
