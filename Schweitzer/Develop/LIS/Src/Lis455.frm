VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSList 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Work Analysis"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14385
   Tag             =   "45500"
   WindowState     =   2  '최대화
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
      Left            =   7980
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "158"
      Top             =   8505
      Width           =   1650
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
      Left            =   9750
      Style           =   1  '그래픽
      TabIndex        =   14
      Tag             =   "162"
      Top             =   8505
      Width           =   1425
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
      Left            =   12780
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "128"
      Top             =   8505
      Width           =   1245
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
      Left            =   11265
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "132"
      Top             =   8505
      Width           =   1410
   End
   Begin VB.CommandButton cmdSelDT2 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   360
      Left            =   9390
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   210
      Width           =   330
   End
   Begin VB.CommandButton cmdSelDT1 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   360
      Left            =   7245
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   210
      Width           =   330
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
      TabIndex        =   7
      Text            =   "Feb 01 1999"
      Top             =   210
      Width           =   1425
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
      TabIndex        =   6
      Text            =   "Feb 28 1999"
      Top             =   210
      Width           =   1425
   End
   Begin TabDlg.SSTab sstWork 
      Height          =   7545
      Left            =   240
      TabIndex        =   0
      Tag             =   "45512"
      Top             =   765
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   13309
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   14411494
      TabCaption(0)   =   "Commemt Deletion List"
      TabPicture(0)   =   "Lis455.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ssDelCmt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExcel(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Modify Result List"
      TabPicture(1)   =   "Lis455.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ssModRst"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdExcel(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Cancel List"
      TabPicture(2)   =   "Lis455.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkOrd"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkAcc"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "ssCancel"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdExcel(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
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
         Height          =   435
         Index           =   2
         Left            =   -62790
         TabIndex        =   18
         Tag             =   "127"
         Top             =   495
         Width           =   1260
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
         Height          =   435
         Index           =   1
         Left            =   -62820
         TabIndex        =   17
         Tag             =   "127"
         Top             =   495
         Width           =   1260
      End
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
         Index           =   0
         Left            =   12210
         Style           =   1  '그래픽
         TabIndex        =   16
         Tag             =   "127"
         Top             =   465
         Width           =   1260
      End
      Begin FPSpread.vaSpread ssDelCmt 
         Height          =   6180
         Left            =   345
         TabIndex        =   1
         Tag             =   "45506"
         Top             =   1050
         Width           =   13155
         _Version        =   196608
         _ExtentX        =   23204
         _ExtentY        =   10901
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
         Protect         =   0   'False
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis455.frx":0054
         VisibleCols     =   7
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread ssModRst 
         Height          =   6225
         Left            =   -74670
         TabIndex        =   2
         Tag             =   "45508"
         Top             =   1050
         Width           =   13140
         _Version        =   196608
         _ExtentX        =   23178
         _ExtentY        =   10980
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
         Protect         =   0   'False
         SpreadDesigner  =   "Lis455.frx":1D24
         VisibleCols     =   7
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread ssCancel 
         Height          =   6210
         Left            =   -74655
         TabIndex        =   3
         Tag             =   "45510"
         Top             =   1080
         Width           =   13155
         _Version        =   196608
         _ExtentX        =   23204
         _ExtentY        =   10954
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
         Protect         =   0   'False
         SpreadDesigner  =   "Lis455.frx":1F2E
         VisibleCols     =   7
         VisibleRows     =   500
      End
      Begin VB.CheckBox chkAcc 
         Caption         =   "Accession"
         Height          =   255
         Left            =   -72420
         TabIndex        =   5
         Tag             =   "45505"
         Top             =   660
         Width           =   1575
      End
      Begin VB.CheckBox chkOrd 
         Caption         =   "Order"
         Height          =   255
         Left            =   -74325
         TabIndex        =   4
         Tag             =   "45504"
         Top             =   660
         Width           =   1575
      End
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
      Left            =   4785
      TabIndex        =   11
      Tag             =   "45501"
      Top             =   255
      Width           =   870
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      Height          =   240
      Left            =   7695
      TabIndex        =   10
      Top             =   300
      Width           =   270
   End
End
Attribute VB_Name = "frmSList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub
