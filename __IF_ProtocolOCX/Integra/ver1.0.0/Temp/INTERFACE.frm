VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmInterface 
   Caption         =   "�������̽� ȭ��"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11910
   ClipControls    =   0   'False
   Icon            =   "INTERFACE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   StartUpPosition =   2  'ȭ�� ���
   Begin Threed.SSPanel pnlTest 
      Height          =   2985
      Left            =   6015
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   3240
      _Version        =   65536
      _ExtentX        =   5715
      _ExtentY        =   5265
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin VB.ListBox lstResult 
         Height          =   2400
         ItemData        =   "INTERFACE.frx":08CA
         Left            =   1635
         List            =   "INTERFACE.frx":08CC
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   90
         Width           =   1500
      End
      Begin VB.ListBox lstOrder 
         Height          =   2400
         ItemData        =   "INTERFACE.frx":08CE
         Left            =   75
         List            =   "INTERFACE.frx":08D0
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   90
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Left            =   840
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2565
         Width           =   1605
      End
   End
   Begin Threed.SSCommand cmdConnect 
      Height          =   450
      Left            =   8160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "CONNECT"
      RoundedCorners  =   0   'False
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   10875
      Top             =   3375
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "TEST"
      Height          =   345
      Left            =   10455
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   1365
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Index           =   0
      Left            =   9840
      Top             =   3375
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10275
      Top             =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   9780
      TabIndex        =   26
      Top             =   4950
      Width           =   1995
      Begin Threed.SSCommand cmdInitial 
         Height          =   570
         Left            =   165
         TabIndex        =   27
         Top             =   330
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "Initialize"
         ForeColor       =   16576
         BevelWidth      =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdClear2 
         Height          =   570
         Left            =   165
         TabIndex        =   28
         Top             =   1035
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "ȭ�� Clear"
         ForeColor       =   32768
         BevelWidth      =   3
         RoundedCorners  =   0   'False
         Picture         =   "INTERFACE.frx":08D2
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   570
         Left            =   165
         TabIndex        =   29
         Top             =   1740
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "Interface ����"
         ForeColor       =   128
         BevelWidth      =   3
         RoundedCorners  =   0   'False
         Picture         =   "INTERFACE.frx":08EE
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4155
      Left            =   90
      TabIndex        =   14
      Top             =   3375
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   7329
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�˻���"
      TabPicture(0)   =   "INTERFACE.frx":090A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblResult"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pnlResult"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOrder"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pnlOrder"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSPanel1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "spdRst"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "spdRst2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCSelList"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "SSPanel10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "listNoOrd"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "�̺�Ʈ ���"
      TabPicture(1)   =   "INTERFACE.frx":0926
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstLog"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdClearLog"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdClearLog 
         Caption         =   "�̺�Ʈ ��� �ʱ�ȭ"
         Height          =   405
         Left            =   -67815
         TabIndex        =   30
         Top             =   3645
         Width           =   1845
      End
      Begin VB.ListBox listNoOrd 
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         Height          =   2010
         ItemData        =   "INTERFACE.frx":0942
         Left            =   210
         List            =   "INTERFACE.frx":0944
         TabIndex        =   16
         Top             =   1380
         Width           =   3405
      End
      Begin VB.ListBox lstLog 
         Appearance      =   0  '���
         Height          =   3450
         ItemData        =   "INTERFACE.frx":0946
         Left            =   -74820
         List            =   "INTERFACE.frx":0948
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         Width           =   8850
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   285
         Left            =   210
         TabIndex        =   17
         Top             =   120
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "   Interface Result ....."
         ForeColor       =   12648447
         BackColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   0
         Alignment       =   2
      End
      Begin Threed.SSPanel lblCSelList 
         Height          =   285
         Left            =   210
         TabIndex        =   18
         Top             =   390
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   8
      End
      Begin FPSpread.vaSpread spdRst2 
         Height          =   3885
         Left            =   6285
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   2745
         _Version        =   196608
         _ExtentX        =   4842
         _ExtentY        =   6853
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   8
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   4210752
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   85
         NoBeep          =   -1  'True
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "INTERFACE.frx":094A
         UserResize      =   0
         VisibleCols     =   4
         VisibleRows     =   85
         TextTip         =   1
      End
      Begin FPSpread.vaSpread spdRst 
         Height          =   3885
         Left            =   3765
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   2505
         _Version        =   196608
         _ExtentX        =   4419
         _ExtentY        =   6853
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   8
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   4210752
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   15
         NoBeep          =   -1  'True
         ProcessTab      =   -1  'True
         ScrollBars      =   0
         SpreadDesigner  =   "INTERFACE.frx":1640
         UserResize      =   0
         VisibleCols     =   4
         VisibleRows     =   15
         TextTip         =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   270
         Left            =   210
         TabIndex        =   21
         Top             =   1125
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   476
         _StockProps     =   15
         Caption         =   "Warning Event Log (���� = F2)"
         ForeColor       =   12648447
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   16576
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   285
         Left            =   210
         TabIndex        =   22
         Top             =   3405
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   " �ֱ� O(Order)"
         ForeColor       =   0
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin Threed.SSPanel lblOrder 
         Height          =   285
         Left            =   1605
         TabIndex        =   23
         Top             =   3405
         Width           =   2010
         _Version        =   65536
         _ExtentX        =   3545
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   8
      End
      Begin Threed.SSPanel pnlResult 
         Height          =   285
         Left            =   210
         TabIndex        =   24
         Top             =   3705
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   " �ֱ� R(Result)"
         ForeColor       =   0
         BackColor       =   8454016
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin Threed.SSPanel lblResult 
         Height          =   285
         Left            =   1605
         TabIndex        =   25
         Top             =   3705
         Width           =   2010
         _Version        =   65536
         _ExtentX        =   3545
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   8
      End
   End
   Begin VB.TextBox txtState 
      Height          =   1035
      Left            =   2745
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Threed.SSCommand cmdStart 
      Height          =   435
      Left            =   990
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8655
      Visible         =   0   'False
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "START"
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   390
      Left            =   45
      TabIndex        =   6
      Top             =   7605
      Width           =   1635
      _Version        =   65536
      _ExtentX        =   2884
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   "Interface Board"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   1710
      TabIndex        =   1
      Top             =   7605
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17886
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox listTest 
      Height          =   420
      Left            =   6750
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8265
      Visible         =   0   'False
      Width           =   2910
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3315
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
      _ExtentY        =   5847
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdIntList 
         Height          =   3120
         Left            =   75
         TabIndex        =   31
         Top             =   135
         Width           =   11715
         _Version        =   196608
         _ExtentX        =   20664
         _ExtentY        =   5503
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   4210752
         MaxCols         =   236
         MaxRows         =   25
         NoBeep          =   -1  'True
         SpreadDesigner  =   "INTERFACE.frx":1B91
         UserResize      =   0
         VisibleCols     =   236
         VisibleRows     =   25
         TextTip         =   1
      End
      Begin Threed.SSFrame fraOrdOpt 
         Height          =   915
         Left            =   6780
         TabIndex        =   2
         Top             =   2010
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   1614
         _StockProps     =   14
         Caption         =   "Option ����"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin VB.OptionButton optOrdOpt 
            Caption         =   "STAT"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   660
            Width           =   915
         End
         Begin VB.OptionButton optOrdOpt 
            Caption         =   "Active"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1005
         End
         Begin VB.OptionButton optOrdOpt 
            Caption         =   "Passive"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   450
            Width           =   975
         End
      End
   End
   Begin VB.CheckBox chkOExist 
      Caption         =   "chkOExist"
      Height          =   225
      Left            =   660
      TabIndex        =   7
      Top             =   210
      Value           =   1  'Ȯ��
      Width           =   1185
   End
   Begin VB.TextBox txtTest 
      Height          =   1050
      Left            =   10440
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   9855
      Picture         =   "INTERFACE.frx":5473
      Top             =   3780
      Width           =   1800
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup01 
         Caption         =   "���õ� ����Ʈ ȭ�� ����"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miHlpClick%
Dim miBSRow%
Dim miBERow%

Dim miPhase As Integer
Dim msWkbuf As String
Dim msRcvBuffer As String

Dim IdleFlag As Integer
Dim PendingFlag As Integer
Dim OrdState As Integer
Dim RstState As Integer

Dim miTimerFlag As Integer
Dim miConnectFlag As Integer
Dim miOrderFlag As Integer
Dim miResultFlag As Integer
Dim miTimerCnt As Integer
Dim miResultCnt As Integer
Dim miOrdRstCnt As Integer

Dim miSpaceCnt%

Dim miResultTimerCnt%
Dim miOrderTimerCnt%

'--- ???
Dim sRcvState As String
Dim sSndState As String
Dim sSndPacket As String
Dim OrderYes As Boolean

Dim miNoTestFlag    As Integer

Dim miPendOrderCnt  As Integer

'--- 2004/1/28 yk
Dim miTimerCnt1 As Integer  'Order
Dim miTimerCnt2 As Integer  'Result
Dim miTimerCnt3 As Integer  'Off Line Check

Dim miOrdFlag   As Integer
Dim miRstFlag   As Integer


Private Sub CommOut_ConnectionMsg()
    On Error GoTo ErrHandler
    
    Dim SendBuff    As String

 '########### CONNECTION ESTABLISH ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    
    'Integra 400
    'SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "00" & Chr(10)
    'Integra 700
'    SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "00" & Chr(10)
    'Integra 800
    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "00" & Chr(10)

    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>

    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>

    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    miTimerFlag = 1
    miConnectFlag = 1
    
'    Comm1.Output = SendBuff
    Call SendSckData(SendBuff)
    
    If giTestMode = 77 Then
        Print #2, SendBuff;
    End If
    
    Exit Sub
ErrHandler:
    miTimerFlag = 0
    miConnectFlag = 0
End Sub

Private Sub CommOut_RequestPendingMsg()
    On Error GoTo ErrHandler
    
    Dim SendBuff    As String

 '########### PENDING BARCODE SAMPLES REQUEST ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    
'    SendBuff = SendBuff & "06" & " " & "COBAS CORE II   " & " " & "60" & Chr(10)
'    SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "60" & Chr(10)
'    SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "60" & Chr(10)
    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "60" & Chr(10)
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
 
    SendBuff = SendBuff & "40" & " " & "1" & Chr(10)
    
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    miTimerFlag = 1
    miOrderFlag = 1
    
'    Comm1.Output = SendBuff
    Call SendSckData(SendBuff)
    
    If giTestMode = 77 Then
        Print #2, SendBuff;
    End If
    
    Exit Sub
ErrHandler:
    miTimerFlag = 0
    miOrderFlag = 0
End Sub
Private Sub CommOut_RequestResultMsg()
    On Error GoTo ErrHandler
    
    Dim SendBuff    As String

 '########### ALL TYPES OF FINAL RESULTS ARE TRANSFFERD TO THE HOST ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>

'    SendBuff = SendBuff & "06" & " " & "COBAS CORE II   " & " " & "09" & Chr(10)
'    SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "09" & Chr(10)
'    SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "09" & Chr(10)
    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "09" & Chr(10)
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & "10" & " " & "01" & Chr(10)
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    miTimerFlag = 1
    miResultFlag = 1
    
'    Comm1.Output = SendBuff
    Call SendSckData(SendBuff)
    
    If giTestMode = 77 Then
        Print #2, SendBuff;
    End If

    Exit Sub
ErrHandler:
    miTimerFlag = 0
    miResultFlag = 0
End Sub

Private Sub ConnectWinSock()
    On Error GoTo ErrRtn

    tcpClient(0).RemotePort = Val(gsPort)
    
    Load tcpClient(1)
    tcpClient(1).RemoteHost = gsIPAddress
    tcpClient(1).Connect tcpClient(1).RemoteHost, tcpClient(0).RemotePort
    
    Call Sleep(500)
    
    ViewMsg "State: " & tcpClient(1).State
    Call DispLogMsg("WinSock State: " & tcpClient(1).State)
        
ErrRtn:
    If Err <> 0 Then
        ViewMsg Err.Description
        Call DispLogMsg("ConnectWinSock(" & Err.Description & ")")
    End If
End Sub

Private Sub SendSckData(ByVal sData As String)
    On Error GoTo ErrSck
    
    With tcpClient(1)
'        txtState = txtState & .State & "/Len:" & Len(sData) & vbCrLf
        
'        If .State <> 7 Then
            .SendData sData
'            .SendData "12345"
'        End If
    End With
        
ErrSck:
    If Err <> 0 Then
        ViewMsg tcpClient(1).State & ":" & Err.Description
        Call DispLogMsg("SendSckData - " & Err.Description & "(State:" & tcpClient(1).State & ")")
        
        '2004/1/15 yk
        Select Case tcpClient(1).State
            Case 8, 9
                Timer2.Enabled = True
            Case Else
        End Select
        
'        Select Case tcpClient(1).State
'            Case 7
'                tcpClient(1).Close
'                Unload tcpClient(1)
'        End Select
    End If
End Sub

Private Sub SetIFProgramInfo()
    
    Dim vTmp
    Dim sTmp$
    Dim tmpData()   As String
    Dim iLen%
    
    gsMachineCd = "": gsMachineNm = ""
    
    vTmp = Command()
    
'    'TEST
'    vTmp = "����ƻ꺴��002 INTEG80-2"
'    '----
    
    iLen = Len(vTmp)
    
    If iLen = 0 Then
        Exit Sub
    End If
    
    tmpData() = Split(vTmp, " ")
    If UBound(tmpData()) > 0 Then
        gsMachineCd = Trim(tmpData(0))
        gsMachineNm = Trim(tmpData(1))
    End If
    
    
    '=== IP/PORT
    sTmp = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "IPAddress")
    If Trim(sTmp) <> "" Then
        gsIPAddress = Trim(sTmp)
    Else
        gsIPAddress = "0.0.0.0"
    End If
    
    sTmp = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Port")
    If Trim(sTmp) <> "" Then
        gsPort = Trim(sTmp)
    Else
        gsPort = "1001"
    End If
    
End Sub
Private Sub ProtectConflict(ByVal sFlag$)
    '0=�ܹ���
    '1=�����(Rack Or Tray ��� ��������, But Rack/Pos ǥ��)
    '2=�����(Rack Or Tray ��� ��������, But Tray/Pos ǥ��)
    '3=�����(Rack Or Tray ��� ��������, But Tray/Cup ǥ��)
    '4=�����(Rack/Pos ��� ����),
    '5=�����(Tray/Pos ��� ����),
    '6=�����(Tray/Cup ��� ����)
    
    If UCase(sFlag) = "Y" Then
        Select Case gsIFMode
            Case "0", "1", "2", "3"
                miTimerFlag = 0
            Case "4", "5", "6"
                miTimerFlag = 1
        End Select
    ElseIf UCase(sFlag) = "N" Then
        miTimerFlag = 0
    End If
End Sub

Private Sub SetProgHWnd(ByVal lHWnd As Long)

    Dim bRet    As Boolean
    
    bRet = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "IF.HWnd", Trim(lHWnd))
        
End Sub

Private Function SpecificProcessResult(ByVal sIFRstCd$, sSpRst1$, sSpRst2$) As Integer
    On Error GoTo ErrHandler
    
    SpecificProcessResult = 0
    
    sSpRst1 = JudgeResult1(sIFRstCd, sSpRst1, sSpRst2)
    
    If sSpRst2 = "" Then
    ElseIf sSpRst2 = "Negative" Then
        sSpRst1 = "Neg(" & sSpRst1 & ")"
    ElseIf sSpRst2 = "Positive" Then
        sSpRst1 = "Pos(" & sSpRst1 & ")"
    End If
    
    If listTest.ListCount > 10 Then
        listTest.RemoveItem (0)
    End If
    
    Exit Function
    
ErrHandler:
    ViewMsg "SpecificProcessResult(" & Err.Description & ")"
    Call DispLogMsg("SpecificProcessResult(" & Err.Description & ")")
End Function

Private Function RegServerOK(ByVal iCRow%, ByVal iRstCnt%, ByVal sIFRstCd$, ByVal sRst1$, ByVal sRst2$) As String
    Dim sBuf$, sCRst1$, sCRst2$, sCRstCd$, sCSvrCd$
    Dim sTSvrCd$, sTRst1$, sTRst2$, sTIFSeq$
    Dim sRetVal$, sTmp$, sIFSeq$
    Dim iTRstCnt%, i%, j%
    Dim vWSeq, vJDate, vJGbn, vJNo, vRack, vPos, vRegNo, vPtNm, vSex, vEmer, vRerun, vOther, vTmp, vIFItemCnt
    Dim objRst As Object
    
    '������ DLL�� Call�Ͽ� �����ʿ� ��������
    sBuf = gRstcfg.sComponent

    If sBuf = "" Then
        ViewMsg "������ �������� ���� DLL ������ �������� �ʽ��ϴ�!!"
        Call DispLogMsg("������ �������� ���� DLL ������ �������� �ʽ��ϴ�!!")
        Exit Function
    End If
    
    Set objRst = CreateObject(sBuf)
    
    With spdIntList
        Call .GetText(1, iCRow, vWSeq)
        Call .GetText(3, iCRow, vJDate)
        Call .GetText(4, iCRow, vJGbn)
        Call .GetText(5, iCRow, vJNo)
        Call .GetText(6, iCRow, vRack)
        Call .GetText(7, iCRow, vPos)
        Call .GetText(8, iCRow, vRegNo)
        Call .GetText(9, iCRow, vPtNm)
        Call .GetText(10, iCRow, vSex)
        Call .GetText(11, iCRow, vEmer)
        Call .GetText(12, iCRow, vRerun)
        Call .GetText(13, iCRow, vOther)
    End With
    
    '### Validation Check S ########################################
    If vWSeq = "" Then
        Exit Function
    End If
    
    If Trim(vJDate) = "" Then
        vJDate = Format(Now, "YYYYMMDD")
    End If
    
    'Len(vJNo)�� ���� �ɼ��� �� ����Ʈ�� ���� ���氡��
    If Len(vJNo) < 10 Then
        Exit Function
    End If
    '### Validation Check E ########################################
    
    sTSvrCd = ""
    sTIFSeq = ""
    sTRst1 = "": sTRst2 = ""
    iTRstCnt = 0
            
    'ServerCd�� ��ȯ - �������ڵ尡 �����ϴ� �͸� ���
    For i = 1 To iRstCnt
        sCRstCd = GetByOne(sIFRstCd, sIFRstCd)
        sCSvrCd = ""

        sCRst1 = GetByOne(sRst1, sRst1)
        sCRst2 = GetByOne(sRst2, sRst2)
        
        With spdIntList
            Call .GetText(16, iCRow, vIFItemCnt)
            
            For j = 1 To CInt(vIFItemCnt)
                Call .GetText(16 + j, iCRow, vTmp)
                
                sTmp = CStr(vTmp)
                
                sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
                
                'IFSeq�� IFRstCd�� Convert
                If Len(sIFSeq) = 2 And Left(sIFSeq, 1) = "C" Then
                '�����϶��� ������ IFSeq��
                    If sIFSeq = sCRstCd Then
                        'IFSeq�� �������ڵ�� Convert
                        sCSvrCd = ConvertIFItemInfo(2, sIFSeq)
                        Exit For
                    End If
                Else
                '�Ϲ��׸��� ���
                    If ConvertIFItemInfo(8, sIFSeq) = sCRstCd Then
                        'IFSeq�� �������ڵ�� Convert
                        sCSvrCd = ConvertIFItemInfo(2, sIFSeq)
                        Exit For
                    End If
                End If
            Next
        End With
        
        If sCSvrCd = "" Then
        Else
            iTRstCnt = iTRstCnt + 1
            sTSvrCd = sTSvrCd & sCSvrCd & Chr(124)
            sTRst1 = sTRst1 & sCRst1 & Chr(124)
            sTRst2 = sTRst2 & sCRst2 & Chr(124)
            sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
        End If
    Next
    
    '������� ����
    Call objRst.SetMachineInfo(gsMachineCd, gsMachineNm)

    sRetVal = objRst.RegServer_24hurine(1, Format(Now, "YYYYMMDD"), CStr(vWSeq) & Chr(124), _
                            CStr(vJDate) & Chr(124), CStr(vJGbn) & Chr(124), CStr(vJNo) & Chr(124), _
                            CStr(vRegNo) & Chr(124), CStr(vPtNm) & Chr(124), CStr(vSex) & Chr(124), _
                            CStr(vEmer) & Chr(124), CStr(vRerun) & Chr(124), CStr(vOther) & Chr(124), _
                            iTRstCnt & Chr(124), sTSvrCd & Chr(3), sTIFSeq & Chr(3), _
                            sTRst1 & Chr(3), sTRst2 & Chr(3))
                
    If sRetVal <> "NO" Then
        ViewMsg CStr(vJNo) & "�� ����� ������ �����Ͽ����ϴ�!!"
    Else
        Call ViewMsgLog("���� ERR : " & CStr(vJNo))
        Call DispLogMsg("���� ERR : " & CStr(vJNo))
    End If
    
    Set objRst = Nothing

End Function

Private Sub Timer1_Old()
'    On Error GoTo ErrHandler
'
'    If miNoTestFlag = 0 Then
'        miTimerCnt = miTimerCnt + 1
'    Else
'        miTimerCnt = 20
'        miNoTestFlag = 0
'    End If
'
'    If miConnectFlag = 1 Then
'        miTimerCnt = 0
'
'        '--- 2004/1/8 yk
'        miPendOrderCnt = miPendOrderCnt + 1
'        If miOrderFlag = 0 Then
'            If miPendOrderCnt > 28 Then
'                Call CommOut_RequestPendingMsg
'                Exit Sub
'            End If
'        End If
'        '---
'
'        Call CommOut_ConnectionMsg
'
'        Exit Sub
'    End If
'
'    If miTimerCnt > 30 Then
'        miTimerCnt = 0
'
'        Call CommOut_ConnectionMsg
'
'        Exit Sub
'    End If
'
'    If miTimerFlag = 1 Then Exit Sub
'
'    If miOrderFlag = 1 Then
'        Call CommOut_RequestPendingMsg
'
'        Exit Sub
'    End If
'
'    miResultCnt = miResultCnt + 1
'
'    If miResultFlag = 1 Then
'        If miResultCnt > 3 Then
'            miResultCnt = 0
'
'            Call CommOut_RequestPendingMsg
'
'            Exit Sub
'        End If
'
'        Call CommOut_RequestResultMsg
'
'        Exit Sub
'    End If
'
'    miOrdRstCnt = miOrdRstCnt + 1
'
'    If miOrdRstCnt > 5 And miOrdRstCnt < 100 Then
'        miOrdRstCnt = 100
'
'        miResultCnt = 0
'
'        Call CommOut_RequestPendingMsg
'
'        Exit Sub
'    ElseIf miOrdRstCnt > 105 Then
'        miOrdRstCnt = 0
'
'        Call CommOut_RequestResultMsg
'
'        Exit Sub
'    End If
'
'    Exit Sub
'ErrHandler:
'    miTimerFlag = 0
'    miConnectFlag = 0
'    miOrderFlag = 0
'    miResultFlag = 0
'    miTimerCnt = 0
'    miResultCnt = 0
'    miOrdRstCnt = 0
End Sub

Private Sub ViewMsgLog(ByVal sMsg$)
    Dim i%, iExist%
    
    iExist = 0
    
    For i = 1 To listNoOrd.ListCount
        If sMsg = listNoOrd.List(i - 1) Then
            iExist = 1
        End If
    Next
    
    If iExist = 0 Then
        With listNoOrd
            .AddItem sMsg
            If .ListCount > 100 Then
                .RemoveItem 0
            End If
        End With
    End If
End Sub
Private Sub LogFileOpen()
    On Error GoTo ErrHandler
    
    'INTERFACE LOG
    Open App.Path & "\" & gsMachineNm & ".log" For Output Shared As #1
    Open App.Path & "\" & gsMachineNm & "Buf.log" For Output Shared As #2

ErrHandler:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, "Log File Open Error"
    End If
End Sub

Private Sub LogFileClose()
    On Error GoTo ErrHandler
        
    Close #1
    Close #2
    
    Exit Sub
ErrHandler:
End Sub

Private Sub PhaseCfg_Protocol_Integra800()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(msWkbuf)
        wkDat = Mid$(msWkbuf, ix1, 1)
             
        Select Case Asc(wkDat)
            Case 1         ' SOH
                msRcvBuffer = ""
            Case 4         ' EOT
                Call Edit_Data
                msRcvBuffer = ""
            
            Case 17, 19    ' DC1, DC3 (XON, XOFF) ����
           
            Case Else      ' Data
                msRcvBuffer = msRcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub

Private Sub Edit_Data()
    On Error GoTo ErrHandler
    
'<---- COBAS ��񿡼� �ַ� ��� S --->
    Dim sBC          As String
    Dim sLC          As String
    Dim iBCpos       As Integer
    Dim iLCpos       As Integer
    
    Dim iErrCode     As Integer
    Dim sGeneralErrCode    As String
'<---- COBAS ��񿡼� �ַ� ��� E --->

    Dim sJDate     As String
    Dim sJGbn      As String
    Dim sJNo      As String
    Dim sIFSpcCd     As String   '�������̽��� ��ü�ڵ�
    Dim sIFRstCd    As String   '�������̽��� �˻��׸��ڵ�
    
    Dim sRack      As String
    Dim sPos       As String
    Dim sSendBuf     As String
    
    Dim sRst     As String
    Dim sRst2    As String
    Dim sExpFlag     As String
    Dim sSignFlag    As String
    
    Dim sTestCd       As String
    Dim sTestNm      As String
    
    Dim sBarCd      As String
    Dim i           As Integer
    Dim sTmpBuffer   As String
    Dim sRetVal     As String
    
    Dim bRetVal As Boolean
    
    Dim lngRetVal As Long
    Dim sBuf      As String
    Dim sSvrCd    As String
    Dim iRetVal   As Integer
    
    iErrCode = 0
    iBCpos = 22
    sBC = Mid(msRcvBuffer, iBCpos, 2)
    
    miTimerCnt3 = 0     '2004/1/28 yk
    
    Select Case sBC
        '### Idle Block, No more result Block ###
        Case "00"
        
        '### CAL Result Block ###
        Case "02"
        
        '### Control Result Block ###
        Case "03"
        
        '### Patient Result Block ###
        Case "04"
        
        '### Order Manipulation response Block ###
        Case "19"
            iErrCode = 99
        
        '### pending Sample Tubes Response Block ###
        Case "62"
            
        '### No More pending Sample Tubes Response Block ###
        Case "69"
        
        Case Else
        
    End Select
    
    iLCpos = iBCpos + 5
    
    Do
        If Asc(Mid(msRcvBuffer, iLCpos, 1)) = 3 Then  'ETX(END OF DATA BLOCK)
            Exit Do
        End If

        sLC = Mid(msRcvBuffer, iLCpos, 2)
        
        Select Case sLC
            Case "00"       'RESULT DATA
                sSignFlag = Trim(Mid(msRcvBuffer, iLCpos + 3, 1))
                sRst = Trim(Mid(msRcvBuffer, iLCpos + 4, 8))
                sExpFlag = Mid(msRcvBuffer, iLCpos + 12, 4)
                
                If sSignFlag = "-" Then
                    If sRst = "9.999999" And Mid(sExpFlag, 3, 2) = "99" Then
                        sRst = "LOWER LIMIT"
                    Else
                        sRst = "-" & ConvertResult1(Mid(sExpFlag, 2, 1), Mid(sExpFlag, 3, 2), sRst, sIFRstCd)
                    End If
                Else
                    If sRst = "9.999999" And Mid(sExpFlag, 3, 2) = "99" Then
                        sRst = "UPPER LIMIT"
                    Else
                        sRst = ConvertResult1(Mid(sExpFlag, 2, 1), Mid(sExpFlag, 3, 2), sRst, sIFRstCd)
                    End If
                End If
                
                If Left(sRst, 1) = "." Then
                    sRst = "0" & sRst
                End If
                
                Call SpecificProcessResult(sIFRstCd, sRst, sRst2)
                
                sRst = JudgeResult1(sIFRstCd, sRst, sRst2)
                
                RstState = 1
                
                Exit Do
            Case "01"       'Result Time --> CAL, QC �϶��� ���۵�
                Exit Do     '���� ��带 ���ø��� ���ý� ���۾ȵ�
                
            Case "02"       'Control ID --> CAL, QC �϶��� ���۵�
                Exit Do     '���� ��带 ���ø��� ���ý� ���۾ȵ�
                
            Case "03"       'Standard Rates --> CAL, QC �϶��� ���۵�
                Exit Do     '���� ��带 ���ø��� ���ý� ���۾ȵ�
                
            Case "04"       'Calibration Curve --> CAL, QC �϶��� ���۵�
                Exit Do     '���� ��带 ���ø��� ���ý� ���۾ȵ�
            
            Case "07"       'ABS Sample Check --> CAL, QC �϶��� ���۵�
                Exit Do     '���� ��带 ���ø��� ���ý� ���۾ȵ�
                
            Case "41"       'Slot State
                'Example "41 023 128 000 000 050<LF>"
                Exit Do
            Case "42"       'Tube Information
                'Integra400
                'Example "42 001 25 1 .....BARCD.....<LF>"
                'Integra700
                'Example "42 001 25 1 .....BARCD.....<LF>"
                'Integra800
                'Example "42 K0001 25 1 .....BARCD.....<LF>"
                
                'Integra400
                'sBarCd = Trim(Mid(msRcvBuffer, iLCpos + 12, 15))
                
                'Integra700
'                sBarCd = Trim(Mid(msRcvBuffer, iLCpos + 12, 15))
                
                'Integra800
                sBarCd = Trim(Mid(msRcvBuffer, iLCpos + 14, 15))
                
                If Len(sBarCd) = 0 Then
                Else
                    gOrderTable.sSampID = sBarCd
                    'Integra400
                    'gOrderTable.sRack = Trim(Mid(msRcvBuffer, iLCpos + 3, 3))
                    'gOrderTable.sPos = Trim(Mid(msRcvBuffer, iLCpos + 7, 2))
                    
                    'Integra700
'                    gOrderTable.sRack = Trim(Mid(msRcvBuffer, iLCpos + 3, 3))
'                    gOrderTable.sPos = Trim(Mid(msRcvBuffer, iLCpos + 7, 2))
                    
                    'Integra800
                    gOrderTable.sRack = Trim(Mid(msRcvBuffer, iLCpos + 3, 5))
                    gOrderTable.sPos = Trim(Mid(msRcvBuffer, iLCpos + 9, 2))
                    
                    'Order �������� �κ�
                    Call Order_Input("B")
                End If
                
                'Integra400
                'iLCpos = iLCpos + 28
                
                'Integra700
'                iLCpos = iLCpos + 28
                
                'Integra800
                iLCpos = iLCpos + 30
            Case "43"       'Test State
                'Example "43 032 1<LF>"
                
            Case "44"       'Cal/CS State
            
            Case "50"       'Patient ID
            
            Case "51"       'Patient Information
            
            Case "52"       'Special Order Selection
            
            Case "53"       'Order ID
                'Version 1.0
                'slipno = Trim(Mid(msRcvBuffer, iLCpos + 3, 9))
                
                'Version 2.0
                sJNo = Trim(Mid(msRcvBuffer, iLCpos + 3, 15))
                
                sIFSpcCd = ""
                
                'Version 1.0
                'iLCpos = iLCpos + 24  'Sample type �ɼ��� No
                'iLCpos = iLCpos + 28  'Sample type �ɼ��� Ok
                
                'Version 2.0
                iLCpos = iLCpos + 30  'Sample type �ɼ��� No
                'iLCpos = iLCpos + 34  'Sample type �ɼ��� Ok
                
            Case "55"       'Test ID
                sIFRstCd = Trim(Mid(msRcvBuffer, iLCpos + 3, 3))
                
                iLCpos = iLCpos + 7
                
            Case "96"       'Error Code
                If OrdState = 0 Then
                'Pending Sample Request�� Response�� ���� ��
                    If Mid(msRcvBuffer, iLCpos + 3, 2) = "61" Then
                        'TimerFlag = 0
                        Exit Do
                    End If
                    
                    Exit Do
                Else
                'Order�� ���� �� Response�� ���� ��
                    If Mid(msRcvBuffer, iLCpos + 3, 2) = "00" Then
                        iErrCode = 0     'Order Input Accepted
                        Exit Do
                    Else
                        If Mid(msRcvBuffer, iLCpos + 3, 2) = "22" Then
                            iErrCode = 1     'Order already available
                            Exit Do
                        ElseIf Mid(msRcvBuffer, iLCpos + 3, 2) = "24" Then
                            'Test not defined - all other tests will be performed
                            iErrCode = 0
                            ViewMsgLog "�Ϻ� �׸��� IF �����ڵ尡 �߸� �����Ǿ����ϴ�!!"
                            Call DispLogMsg("�Ϻ� �׸��� IF �����ڵ尡 �߸� �����Ǿ����ϴ�!!")
                            Exit Do
                        Else
                            iErrCode = 2     '��Ÿ ������ �˻���, ID ����, ORDER NO ����, SAMPLE TYPE ���� ���� ����
                            ViewMsgLog "Tx Warning : " & Mid(msRcvBuffer, iLCpos + 3, 2)
                            Call DispLogMsg("Tx Warning : " & Mid(msRcvBuffer, iLCpos + 3, 2))
                            Exit Do
                        End If
                    End If
                End If
            Case "98"       'Protocol Version
                ViewMsg "Protocol Version - " & Mid(msRcvBuffer, iLCpos + 3, 4)
                Exit Do
            
            Case "99"       'General Error Code
                sGeneralErrCode = Mid(msRcvBuffer, iLCpos + 3, 2)
                ViewMsgLog "Ge Warning : " & sGeneralErrCode
                Call DispLogMsg("Ge Warning : " & sGeneralErrCode)
                Exit Do
                
            Case Else
                Exit Do
        End Select
    Loop
            
'### ORDER INPUT RESPONSE ################################################################
    'OrdState = 1 --> From Host To Integra : Sample Order ���� ����
    'OrdState = 2 --> From Host To Integra : Order Delete�� ��û�� ����
    'OrdState = 0 --> Order ������ ����� ���� ����
    
    If sBC = "19" And iErrCode = 0 Then
        If OrdState = 1 Then
            ViewMsg gOrderTable.sSampID & "   Order OK!"
            OrdState = 0   'Order ������ ����� ���� ����
            
            '''Call DisplayOrderOK
        ElseIf OrdState = 2 Then
            ViewMsg gOrderTable.sSampID & "   Delete OK!"
            
            Call Order_Input
        End If
    ElseIf sBC = "19" And iErrCode = 1 Then
        'LineCode 22�� �����߻�
        ViewMsgLog "���� Order�� �̹� �����ϰų� Full(50��)�� �����Դϴ�.!!"
        Call DispLogMsg("���� Order�� �̹� �����ϰų� Full(50��)�� �����Դϴ�.!!")
            
        miTimerFlag = 0
        msRcvBuffer = ""
        Call cmdInitial.DoClick
        
        Exit Sub
    ElseIf sBC = "19" And iErrCode = 2 Then
        'LineCode 22�� ������ �����߻�
        ViewMsgLog "Order �ź�!! " & _
           "TestNo Err, Already Running, ID Err, OrderNo Err, SampleType Err ���� �����߻�..."
        Call DispLogMsg("Order �ź�!! " & _
                    "TestNo Err, Already Running, ID Err, OrderNo Err, SampleType Err ���� �����߻�...")
        
        miTimerFlag = 0
        msRcvBuffer = ""
        Call cmdInitial.DoClick
        
        Exit Sub
    End If
    
'### SAMPLE RESULT ���� & ��� #####################################################
    If Len(sJNo) > 0 And sIFRstCd <> "" Then
        If RstState = 1 And sBC = "04" Then
            RstState = 0
            
            '�߾Ӻ����� Ư���� ���ó�� - B(BIL)�� ��� ���ο� ���� CREA ó��
                '���� 1) CREA���� BIL�� ����� ���� ���´�..
            iRetVal = Amc_ProcessCreatinine(sJNo, sIFRstCd, sRst, sRst2)
                        
            'CREA�� ��� ����
            Select Case iRetVal
                Case 0
                    '������ �Ķ������ sCurRow�� �̿� - 0:�̵��, 1:���
                    Call DisplayResultOK(3, Format(Now, "YYYYMMDD"), "", _
                                        "", "", sJNo, "", "", "", "", "", "", "", "", _
                                        1, sIFRstCd & Chr$(124), sRst & Chr$(124), sRst2 & Chr$(124), _
                                        "", "1")
                'CREA ������ �̵��
                Case 1
                    Call DisplayResultOK(3, Format(Now, "YYYYMMDD"), "", _
                                        "", "", sJNo, "", "", "", "", "", "", "", "", _
                                        1, sIFRstCd & Chr$(124), sRst & Chr$(124), sRst2 & Chr$(124), _
                                        "", "0")
                'CREA ������ ���
                Case 2
                    Call DisplayResultOK(3, Format(Now, "YYYYMMDD"), "", _
                                        "", "", sJNo, "", "", "", "", "", "", "", "", _
                                        1, sIFRstCd & Chr$(124), sRst & Chr$(124), sRst2 & Chr$(124), _
                                        "", "1")
            End Select
            
            '�ٷ� ����䱸
            miTimerCnt2 = 15
'            miRstFlag = 0
        End If
    Else
        If RstState = 1 And sBC = "04" Then
            RstState = 0
        End If
    End If
    
    miTimerFlag = 0
    
    Exit Sub
ErrHandler:
    ViewMsg "Edit_Data ���� �߻�" & "(" & Err.Description & ")"
    Call DispLogMsg("Edit_Data ���� �߻�" & "(" & Err.Description & ")")
        
    miTimerFlag = 0
    
    msRcvBuffer = ""
    cmdInitial.DoClick
End Sub
Private Function Amc_ProcessCreatinine(ByVal sBarCd$, ByVal sIFRstCd$, sCREA_Rst1$, sCREA_Rst2$) As Integer
    On Error GoTo ErrHandler
    
    Dim i%, j%, iCRow%, iBIL_Exist%, iCREA_Exist%, iITEM_Exist%
    Dim vIFItemCnt, vTmp
    Dim sIFSeq$, sTmp$, sBILRst1$, sBILRst2$, sCREA_PrevRst1$, sCREA_PrevRst2$
    
    Amc_ProcessCreatinine = 0
    iBIL_Exist = 0
    iCREA_Exist = 0
    
    iCRow = FindIFListWithJNo(sBarCd)
    
    If iCRow = 0 Then
        Exit Function
    End If
    
    With spdIntList
        Call .GetText(16, iCRow, vIFItemCnt)
        
        '���� ���۵� ���� C(CREA)������ �׸��� �̹� �����ϴ��� ����
        For i = 1 To Val(vIFItemCnt)
            Call .GetText(16 + i, iCRow, vTmp)
                
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
            sCREA_PrevRst1 = GetByOne(sTmp, sTmp)
            sCREA_PrevRst2 = GetByOne(sTmp, sTmp)
                        
            iITEM_Exist = 0
            
            'IFSeq�� IFRstCd�� Convert
            If ConvertIFItemInfo(8, sIFSeq) = sIFRstCd Then
                For j = 1 To giOriginIFItemCnt
                    If gIFItem(j).s01 = sIFSeq Then
                        iITEM_Exist = 1
                        
                        If gIFItem(j).s05 = "C" Then
                            If sCREA_PrevRst1 = "" Then
                                iCREA_Exist = 1
                            Else
                                iCREA_Exist = 2
                            End If
                            
                            Exit For
                        End If
                        
                        Exit For
                    End If
                Next
                
                If iITEM_Exist = 1 Then
                    Exit For
                End If
            End If
        Next
        
        iITEM_Exist = 0
        
        'B(BIL)���� �����ϴ��� ����
        For i = 1 To Val(vIFItemCnt)
            Call .GetText(16 + i, iCRow, vTmp)
                
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
            sBILRst1 = GetByOne(sTmp, sTmp)
            sBILRst2 = GetByOne(sTmp, sTmp)
            
            'IFSeq�� IFSpcCd�� Convert
            For j = 1 To giOriginIFItemCnt
                If gIFItem(j).s01 = sIFSeq Then
                    If gIFItem(j).s05 = "B" Then
                        iITEM_Exist = 1
                        
                        If sBILRst1 <> "" And Val(sBILRst1) >= 5 Then
                            iBIL_Exist = 2
                        ElseIf sBILRst1 <> "" And Val(sBILRst1) < 5 Then
                            iBIL_Exist = 1
                        Else
                            iBIL_Exist = 0
                        End If
                        
                        Exit For
                    End If
                    
                    Exit For
                End If
            Next
            
            If iITEM_Exist = 1 Then
                Exit For
            End If
        Next
        
        If iBIL_Exist = 1 Then
            'CREA ���� ����
            If iCREA_Exist = 0 Then
                sCREA_Rst1 = sCREA_Rst1
                sCREA_Rst2 = sCREA_Rst2
                
                Amc_ProcessCreatinine = 0
                
            'CREA ������ ���
            Else
                sCREA_Rst1 = sCREA_Rst1
                sCREA_Rst2 = sCREA_Rst2
                
                Amc_ProcessCreatinine = 2
                
            End If
        Else
            'CREA ���� ����
            If iCREA_Exist = 0 Then
                sCREA_Rst1 = sCREA_Rst1
                sCREA_Rst2 = sCREA_Rst2
                
                Amc_ProcessCreatinine = 0
            
            'CREA���� �������� ������ �ϴ� ���� �̵��
            ElseIf iCREA_Exist = 1 Then
                sCREA_Rst1 = sCREA_Rst1
                sCREA_Rst2 = sCREA_Rst2
                
                Amc_ProcessCreatinine = 1
            
            'CREA���� �����ϸ� ū ���� ���� ���
            ElseIf iCREA_Exist = 2 Then
                If Val(sCREA_Rst1) >= Val(sCREA_PrevRst1) Then
                    sCREA_Rst1 = sCREA_Rst1
                Else
                    sCREA_Rst1 = sCREA_PrevRst1
                End If
                
                sCREA_Rst2 = sCREA_Rst2
                
                Amc_ProcessCreatinine = 2
            
            End If
        End If
    End With
    
    If listTest.ListCount > 10 Then
        listTest.RemoveItem (0)
    End If
    
    listTest.AddItem "CREA : " & CStr(iCREA_Exist) & ",  BIL : " & CStr(iBIL_Exist) & ", " & sIFSeq & ", " & sCREA_Rst1
    
    Exit Function
ErrHandler:
    ViewMsg "Amc_ProcessCreatinine(" & Err.Description & ")"
    Call DispLogMsg("Amc_ProcessCreatinine(" & Err.Description & ")")
End Function
Private Sub Order_Input_20040420(Optional ByVal sSendYN$)
''ȯ���� Order ����
'    Dim SendBuff As String
'    Dim i%, j%, k%, iOrdCnt%, iOldPos%
'    Dim vIFCnt, vTmp
'    Dim sTmp$, sTestCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$, sOldTmp$
'    Dim sTIFOrdCd$
'    Dim objOrd As Object
'
'    SendBuff = ""
'    sTmp = ""
'
'    'Order Dll�� Call�Ͽ� �����ʿ� Order�� ������
'    sBuf = gOrdCfg.sComponent
'
'    If sBuf = "" Then
'        ViewMsg "���� Dll ������ �������� �ʽ��ϴ�!!"
'        Call DispLogMsg("���� Dll ������ �������� �ʽ��ϴ�!!")
'        Exit Sub
'    End If
'
'    Set objOrd = CreateObject(sBuf)
'    Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
'    sOrdList = objOrd.FetchOrder(gsMachineCd, "", "", "", gOrderTable.sSampID)
'    Set objOrd = Nothing
'
'    If sOrdList = "" Then
'        Call ViewMsgLog("��ü ERR : " & gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID)
'        '2003/11/4 yk
'        miConnectFlag = 0
'        miTimerFlag = 0
'        miOrderFlag = 0
'        miResultFlag = 1: miResultCnt = 1
'        miNoTestFlag = 1
'
'        Exit Sub
'    Else
'        'sOrdList ����
'        'ȯ�ڹ�ȣ | �̸� | n | IFSeq 1 | IFSeq 2 | ... | IFSeq n |
'        gOrderTable.sJDate = ""
'        gOrderTable.sRegNo = GetByOne(sOrdList, sOrdList)   '�����
'        gOrderTable.sName = ""
'        gOrderTable.sOther = GetByOne(sOrdList, sOrdList)   '��������ڷ�
'        gOrderTable.iOrdCnt = Val(GetByOne(sOrdList, sOrdList))
'        gOrderTable.sOrdOpt = "S"
'        gOrderTable.sWDate = Format(Now, "YYYYMMDD")
'        gOrderTable.sJNo = gOrderTable.sSampID
'
'        If gOrderTable.iOrdCnt = 0 Then
'            Call ViewMsgLog("���� ERR : " & gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID)
'
'            Exit Sub
'        Else
'            For i = 1 To gOrderTable.iOrdCnt
'                sIFSeq = GetByOne(sOrdList, sOrdList)
'
'                sTmp = sIFSeq
'
'                'IFOrdCd�� ��ȯ
'                sTmp = ConvertIFItemInfo(6, sTmp)
'
'                If sTmp = "" Then
'                Else
'                    iOrdCnt = iOrdCnt + 1
'
'                    'IFSeq�� ��ģ��
'                    sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
'                End If
'            Next
'
'            'IFSeq ������ �籸��
'            sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
'
'            gOrderTable.iOrdCnt = iOrdCnt
'            ReDim gOrderTable.sIFSeq(iOrdCnt)
'
'            For i = 1 To iOrdCnt
'                gOrderTable.sIFSeq(i) = GetByOne(sTIFSeq, sTIFSeq)
'            Next
'        End If
'
'        'Order ���۾��� ���븸 Display - DisplayResultOK ���� ���
'        If sSendYN = "N" Then
'            Call DisplayOrderOK
'
'            Exit Sub
'        End If
'
'        'Order Packet ����
'        SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
'
'        'Integra 400
'        'SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "10" & Chr(10)     '<LF>
'        'Integra 700
''        SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "10" & Chr(10)     '<LF>
'        'Integra 800
'        SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "10" & Chr(10)     '<LF>
'
'        SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
'
'        SendBuff = SendBuff & "50" & " " & String(15, " ") & Chr(10)     '<LF>
'
'        'Sample Type No
'        SendBuff = SendBuff & "53" & " " & gOrderTable.sSampID & String(15 - Len(Trim(gOrderTable.sSampID)), " ") & _
'                              " " & Right(gOrderTable.sWDate, 2) & "/" & Mid(gOrderTable.sWDate, 5, 2) & "/" & Left(gOrderTable.sWDate, 4) & _
'                              Chr(10)      '<LF>
'
'        If sSendYN = "B" Then
'        'Barcode type
'            'Integra400
''            SendBuff = SendBuff & "54" & " " & "000 00" & _
''                                  " " & gOrderTable.sOrdOpt & " " & Space(21) & _
''                                  " " & Space(21) & Chr(10)    '<LF>
'
'            'Integra700
''                SendBuff = SendBuff & "54" & " " & "000 00" & _
''                                      " " & gOrderTable.sOrdOpt & " " & Space(21) & _
''                                      " " & Space(21) & Chr(10)    '<LF>
'
'            'Integra800
'                SendBuff = SendBuff & "54" & " " & "00000 00" & _
'                                      " " & gOrderTable.sOrdOpt & " " & Space(21) & _
'                                      " " & Space(21) & Chr(10)    '<LF>
'        ElseIf sSendYN = "R" Then
'        'Rack/Pos type
'            SendBuff = SendBuff & "54" & " " & gOrderTable.sRack & " " & gOrderTable.sPos & _
'                                  " " & gOrderTable.sOrdOpt & " " & Space(21) & _
'                                  " " & Space(21) & Chr(10)    '<LF>
'        End If
'
'        '--- ���� ���� ���� ������ �籸�� S --------------------------------------------------
'        sTIFOrdCd = ""
'
'        For i = 1 To iOrdCnt
'            sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
'
'            If sTmp = "" Then
'            Else
'                sTIFOrdCd = sTIFOrdCd & sTmp & ","
'            End If
'        Next
'
'        sTmp = sTIFOrdCd
'
'        sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
'        '--- ���� ���� ���� ������ �籸�� E --------------------------------------------------
'
'        sOldTmp = ""
'
'        For i = 1 To iOrdCnt
'            sTmp = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
'
'            If sTmp = "" Then
'            Else
'                SendBuff = SendBuff & "55" & " " & String(3 - Len(sTmp), " ") & sTmp & Chr(10)
'            End If
'        Next
'
'        SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
'        SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
'
''        Comm1.Output = SendBuff
'        Call SendSckData(SendBuff)
'
'        If giTestMode = 77 Then
'            Print #2, SendBuff;
'        End If
'    End If
'
'    Call DisplayOrderOK
'
'    OrdState = 1
End Sub

Private Sub Order_Input(Optional ByVal sSendYN$)
'ȯ���� Order ����
    Dim SendBuff As String
    Dim i%, j%, k%, iOrdCnt%, iOldPos%
    Dim vIFCnt, vTmp
    Dim sTmp$, sTestCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$, sOldTmp$
    Dim sTIFOrdCd$
    Dim objOrd As Object
    
    SendBuff = ""
    sTmp = ""
    
    'Order Dll�� Call�Ͽ� �����ʿ� Order�� ������
    sBuf = gOrdCfg.sComponent
    
    If sBuf = "" Then
        ViewMsg "���� Dll ������ �������� �ʽ��ϴ�!!"
        Call DispLogMsg("���� Dll ������ �������� �ʽ��ϴ�!!")
        Exit Sub
    End If
    
    Set objOrd = CreateObject(sBuf)
    Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
    sOrdList = objOrd.FetchOrder_24hurine(gsMachineCd, "", "", "", gOrderTable.sSampID)
    Set objOrd = Nothing
    
    If sOrdList = "" Then
        Call ViewMsgLog("��ü ERR : " & gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID)
        '2003/11/4 yk
        miConnectFlag = 0
        miTimerFlag = 0
        miOrderFlag = 0
        miResultFlag = 1: miResultCnt = 1
        miNoTestFlag = 1
        
        Exit Sub
    Else
        'sOrdList ����
        'ȯ�ڹ�ȣ | �̸� | n | IFSeq 1 | IFSeq 2 | ... | IFSeq n |
        gOrderTable.sJDate = ""
        gOrderTable.sRegNo = GetByOne(sOrdList, sOrdList)   '�����
        gOrderTable.sName = GetByOne(sOrdList, sOrdList)    'time/totvol/bcr
        gOrderTable.sOther = GetByOne(sOrdList, sOrdList)   '��������ڷ�
        gOrderTable.iOrdCnt = Val(GetByOne(sOrdList, sOrdList))
        gOrderTable.sOrdOpt = "S"
        gOrderTable.sWDate = Format(Now, "YYYYMMDD")
        gOrderTable.sJNo = gOrderTable.sSampID
            
        If gOrderTable.iOrdCnt = 0 Then
            Call ViewMsgLog("���� ERR : " & gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID)
            
            Exit Sub
        Else
            For i = 1 To gOrderTable.iOrdCnt
                sIFSeq = GetByOne(sOrdList, sOrdList)
                
                sTmp = sIFSeq
                
                'IFOrdCd�� ��ȯ
                sTmp = ConvertIFItemInfo(6, sTmp)
                
                If sTmp = "" Then
                Else
                    iOrdCnt = iOrdCnt + 1
                    
                    'IFSeq�� ��ģ��
                    sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
                End If
            Next
            
            'IFSeq ������ �籸��
            sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
            
            gOrderTable.iOrdCnt = iOrdCnt
            ReDim gOrderTable.sIFSeq(iOrdCnt)
            
            For i = 1 To iOrdCnt
                gOrderTable.sIFSeq(i) = GetByOne(sTIFSeq, sTIFSeq)
            Next
        End If
                                
        'Order ���۾��� ���븸 Display - DisplayResultOK ���� ���
        If sSendYN = "N" Then
            Call DisplayOrderOK
            
            Exit Sub
        End If
        
        'Order Packet ����
        SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
        
        'Integra 400
        'SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "10" & Chr(10)     '<LF>
        'Integra 700
'        SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "10" & Chr(10)     '<LF>
        'Integra 800
        SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "10" & Chr(10)     '<LF>
        
        SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
        
        SendBuff = SendBuff & "50" & " " & String(15, " ") & Chr(10)     '<LF>
        
        'Sample Type No
        SendBuff = SendBuff & "53" & " " & gOrderTable.sSampID & String(15 - Len(Trim(gOrderTable.sSampID)), " ") & _
                              " " & Right(gOrderTable.sWDate, 2) & "/" & Mid(gOrderTable.sWDate, 5, 2) & "/" & Left(gOrderTable.sWDate, 4) & _
                              Chr(10)      '<LF>
        
        If sSendYN = "B" Then
        'Barcode type
            'Integra400
'            SendBuff = SendBuff & "54" & " " & "000 00" & _
'                                  " " & gOrderTable.sOrdOpt & " " & Space(21) & _
'                                  " " & Space(21) & Chr(10)    '<LF>

            'Integra700
'                SendBuff = SendBuff & "54" & " " & "000 00" & _
'                                      " " & gOrderTable.sOrdOpt & " " & Space(21) & _
'                                      " " & Space(21) & Chr(10)    '<LF>
            
            'Integra800
                SendBuff = SendBuff & "54" & " " & "00000 00" & _
                                      " " & gOrderTable.sOrdOpt & " " & Space(21) & _
                                      " " & Space(21) & Chr(10)    '<LF>
        ElseIf sSendYN = "R" Then
        'Rack/Pos type
            SendBuff = SendBuff & "54" & " " & gOrderTable.sRack & " " & gOrderTable.sPos & _
                                  " " & gOrderTable.sOrdOpt & " " & Space(21) & _
                                  " " & Space(21) & Chr(10)    '<LF>
        End If
        
        '--- ���� ���� ���� ������ �籸�� S --------------------------------------------------
        sTIFOrdCd = ""
        
        For i = 1 To iOrdCnt
            sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
            
            If sTmp = "" Then
            Else
                sTIFOrdCd = sTIFOrdCd & sTmp & ","
            End If
        Next
        
        sTmp = sTIFOrdCd
        
        sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
        '--- ���� ���� ���� ������ �籸�� E --------------------------------------------------
        
        sOldTmp = ""
        
        For i = 1 To iOrdCnt
            sTmp = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
                
            If sTmp = "" Then
            Else
                SendBuff = SendBuff & "55" & " " & String(3 - Len(sTmp), " ") & sTmp & Chr(10)
            End If
        Next
        
        SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
        SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
        
'        Comm1.Output = SendBuff
        Call SendSckData(SendBuff)
        
        If giTestMode = 77 Then
            Print #2, SendBuff;
        End If
    End If
    
    Call DisplayOrderOK
    
    OrdState = 1
End Sub


Private Sub DisplayInit()
    Dim i%
    
    Set gfIFDisplayForm = frmInterface
    
    'Title ����
    Me.Caption = "   " & UCase$(gsMachineNm) & " �������̽� ȭ�� - BY ACK Co., Ltd."
    
    
    With spdIntList
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        '.EditModePermanent = True
        '.Protect = True
        .Lock = True
        .NoBeep = True
        .BlockMode = False
        
        .Col = 6
        .Col2 = 7
        .Row = -1
        .Row2 = -1
        .BlockMode = True
        .Lock = False
        .BlockMode = False
            
        Call SetSpdIntLIstColHidden
         
        'Rack, Pos ��뿩��
        If Val(gIFRack.sMaxRack) = 0 Then
            For i = 6 To 7
                .Col = i
                .ColHidden = True
            Next
        Else
            For i = 6 To 7
                .Col = i
                .ColHidden = False
            Next
        End If
        
        .MaxRows = 0
    End With
    
    
    With spdRst
        .BlockMode = True
        .Col = -1: .Row = -1
        .Action = ActionClearText
        .EditModePermanent = True
        .NoBeep = True
        .BlockMode = False
    End With
    
     With spdRst2
        .BlockMode = True
        .Col = -1: .Row = -1
        .Action = ActionClearText
        .EditModePermanent = True
        .NoBeep = True
        .BlockMode = False
    End With

'Interface Mode�� ���� Display
    If gsIFMode = "0" Then
    'Uni-Direction
'        With gfIFDisplayForm
'            .fraSendOrd.Visible = False
'            .fraBarCd.Top = 4920
'        End With
    Else
    'Bi-Direction
        '1=�����(Rack Or Tray ��� ��������, But Rack/Pos ǥ��)
        '2=�����(Rack Or Tray ��� ��������, But Tray/Pos ǥ��)
        '3=�����(Rack Or Tray ��� ��������, But Tray/Cup ǥ��)
        '4=�����(Rack/Pos ��� ����),
        '5=�����(Tray/Pos ��� ����),
        '6=�����(Tray/Cup ��� ����),

        With gfIFDisplayForm
'            .fraBarCd.Visible = False
            
            If gsIFMode = "1" Then
            'Rack Or Tray ��� ��������, But Rack/Pos ǥ��
'                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Rack"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "2" Then
            'Rack Or Tray ��� ��������, But Tray/Pos ǥ��
'                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "3" Then
            'Rack Or Tray ��� ��������, But Tray/Cup ǥ��
'                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Cup"))
            ElseIf gsIFMode = "4" Then
            'Rack/Pos ��� ����
'                .pnlRackTray = "Rack"
'                .pnlPosCup = "Pos"
                
                Call .spdIntList.SetText(6, 0, CVar("Rack"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "5" Then
            'Tray/Pos ��� ����
'                .pnlRackTray = "Tray"
'                .pnlPosCup = "Pos"
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "6" Then
            'Tray/Cup ��� ����
'                .pnlRackTray = "Tray"
'                .pnlPosCup = "Cup"
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Cup"))
            End If
        End With
    End If
    
'Transmit Mode�� ���� Display
    If gsTXMode = "0" Then
    'Batch
        '��� Option�� Client�� �ϸ� OK
    ElseIf gsTXMode = "1" Then
    'RealTime �� �׸�
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    ElseIf gsTXMode = "2" Then
    'RealTime  �� ȯ�ھ�
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    End If
    
'Initialize mode�� ���� Display
    If gsINITMode = "0" Then
    'Not Use
        gfIFDisplayForm.cmdInitial.Visible = False
    Else
    'Use
        gfIFDisplayForm.cmdInitial.Visible = True
    End If
    
''MaxLength Check
'    txtRack.MaxLength = CInt(Val(gIFRack.sRackDigit))
'    txtPos.MaxLength = CInt(Val(gIFRack.sPosDigit))
'    txtOrdNo.MaxLength = CInt(Val(gOrdCfg.sFSize(3)))
'    txtBarCd.MaxLength = CInt(Val(gOrdCfg.sFSize(3)))
    
''APMode�� ���� ��� Display
'    If gsAPMode = "1" Then
'        With gfIFDisplayForm.spdRst
'            .ColWidth(1) = 10#
'            .ColWidth(2) = 7.5
'            .ColWidth(3) = 0#
'            .ColWidth(4) = 4#
'        End With
'
'        With gfIFDisplayForm.spdRst2
'            .ColWidth(1) = 10#
'            .ColWidth(2) = 7.5
'            .ColWidth(3) = 0#
'            .ColWidth(4) = 4#
'        End With
'    End If
End Sub

Private Sub DisplayInitItem()
    Dim i%, j%
    Dim iCurItemCnt%
    
    For i = 1 To MAXIFITEM
    'Interface �׸� Seq�� ��ġ�ϴ� �˻�� �Ѹ���
        For j = 1 To giOriginIFItemCnt
            If Format$(i, "000") = gIFItem(j).s01 Then
                iCurItemCnt = iCurItemCnt + 1
                Call gfIFDisplayForm.spdIntList.SetText(16 + giTotIFItemCnt + iCurItemCnt, 0, gIFItem(j).s02 & "")
                
                Exit For
            End If
        Next
    Next
    
    For i = 1 To MAXCALITEM
    '����׸�� ��ġ�ϴ� �˻�� �Ѹ���
        For j = 1 To giOriginCalItemCnt
            If "C" & CStr(i - 1) = gCalItem(j).s01 Then
                iCurItemCnt = iCurItemCnt + 1
                Call gfIFDisplayForm.spdIntList.SetText(16 + giTotIFItemCnt + iCurItemCnt, 0, gCalItem(j).s02 & "")
            
                Exit For
            End If
        Next
    Next
End Sub

Private Sub DisplayExistIFList(ByVal iPersonCnt As Integer, ByVal iRowCnt As Integer, ByVal sTotList As String)
    On Error GoTo ErrHandler
    
    Dim i%, j%, iCnt%
    Dim sOneRow$, sWSeq$, sPWSeq$, sTWSeq$, sTmp$
    Dim iPCnt%, iEqual%, iERCnt%
    Dim sField() As String
    Dim sOrdField() As String
    Dim vEmer, vRerun
    
    ReDim sField(iPersonCnt, MAXORDERFIELD + 6)
    ReDim sOrdField(iPersonCnt, MAXIFITEM)
    
    sPWSeq = ""
    iPCnt = 0
    iCnt = 0
        
    For i = 1 To iRowCnt
        sOneRow = GetByOneUserSymbol(sTotList, sTotList, Chr(3))
        
        For j = 1 To MAXORDERFIELD + 6
            If j = 1 Then
                sWSeq = GetByOne(sOneRow, sOneRow)
                
                If sPWSeq = sWSeq Then
                    iEqual = 1
                Else
                    iEqual = 0
                    iPCnt = iPCnt + 1
                    sField(iPCnt, j) = sWSeq
                End If
            Else
                sTmp = GetByOne(sOneRow, sOneRow)
                
                If j = IFTESTFIELD + 5 Then
                    If iEqual = 1 Then
                        iCnt = iCnt + 1
                        sOrdField(iPCnt, iCnt) = sTmp
                        sField(iPCnt, j) = iCnt
                    Else
                        iCnt = 1
                        sOrdField(iPCnt, iCnt) = sTmp
                        sField(iPCnt, j) = iCnt
                    End If
                Else
                    If iEqual = 1 Then
                    Else
                        sField(iPCnt, j) = sTmp
                    End If
                End If
            End If
            
            sPWSeq = sWSeq
        Next
    Next
    
    For i = 1 To iPersonCnt
        If ExistIFList(sField(i, 3), sField(i, 4), sField(i, 5), sField(i, 1)) = "YES" Then
            ViewMsg "�ߺ��Ǵ� Interface List�� �����մϴ�!!"
        Else
            iERCnt = 0
            
            With gfIFDisplayForm.spdIntList
                .MaxRows = .MaxRows + 1
                
                For j = 1 To MAXORDERFIELD + 6
                    Call .SetText(j, .MaxRows, sField(i, j) & "")
                Next
                
                For j = 1 To sField(i, 16)
                    Call .SetText(16 + j, .MaxRows, sOrdField(i, j) & "|")
                Next
                
                Call .GetText(11, .MaxRows, vEmer)
                If vEmer = "Y" Then
                    Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 5, .MaxRows, .MaxRows, RGB(0, 0, 0), ������)
                    iERCnt = iERCnt + 1
                Else
                End If
                
                Call .GetText(12, .MaxRows, vRerun)
                If vRerun = "Y" Then
                    Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 5, .MaxRows, .MaxRows, RGB(0, 0, 0), �帰�Ķ�)
                    iERCnt = iERCnt + 1
                Else
                End If

                If iERCnt = 2 Then
                    Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 5, .MaxRows, .MaxRows, RGB(255, 255, 255), RGB(0, 0, 0))
                End If
            End With
        End If
    Next

    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayExistIFList �����߻� - (" & Err.Number & ")"
End Sub

Private Sub DisplayIFList()
    On Error GoTo ErrHandler
    
    Dim i%
    Dim vTmp, vJDate, vJGbn, vJNo, vRegNo, vName, vSex, vOther, vIntcnt, vEmer, vRerun
    Dim vRack, vPos
    Dim j%
    Dim iCurMaxRack%, iCurMaxPos%, iCurMaxRow%, iCurIntCnt%
    Dim iERCnt%
    Dim sWSeq$
    
    ViewMsg ""
    
    If gfIFDisplayForm.spdIntList.MaxRows = 0 Then
        
        '1999/09/29 ������ ���� - 10���� RACK�� �� ���� ���Ƿ� DEFAULT�� 10�� RACK���� �ߵ���.
        iCurMaxRack = 10
        
        'iCurMaxRack = 0
        iCurMaxPos = 0
    Else
        Call gfIFDisplayForm.spdIntList.GetText(6, gfIFDisplayForm.spdIntList.MaxRows, vRack)
        Call gfIFDisplayForm.spdIntList.GetText(7, gfIFDisplayForm.spdIntList.MaxRows, vPos)
        
        iCurMaxRow = gfIFDisplayForm.spdIntList.MaxRows
        iCurMaxRack = CInt(vRack)
        iCurMaxPos = CInt(vPos)
    End If
    
    iCurIntCnt = 0
    
    With gfIFDisplayForm.spdList
        For i = 1 To .MaxRows
            iERCnt = 0
            
            Call .GetText(1, i, vTmp)
            
            If vTmp = "1" Then
                Call .GetText(2, i, vJDate)
                Call .GetText(3, i, vJGbn)
                Call .GetText(4, i, vJNo)
                Call .GetText(5, i, vRegNo)
                Call .GetText(6, i, vName)
                Call .GetText(7, i, vSex)
                Call .GetText(8, i, vEmer)
                Call .GetText(9, i, vRerun)
                Call .GetText(10, i, vOther)
                Call .GetText(11, i, vIntcnt)
                
                If ExistIFList(CStr(vJDate), CStr(vJGbn), CStr(vJNo)) = "YES" Then
                    ViewMsg "�ߺ��Ǵ� Interface List�� �����մϴ�!!"
                Else
                    iCurIntCnt = iCurIntCnt + 1
                    
                    With gfIFDisplayForm.spdIntList
                        'WorkSeq �ʱ�ȭ
                        sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
                        
                        .MaxRows = .MaxRows + 1
                        
                        Call .SetText(1, .MaxRows, sWSeq & "")
                        Call .SetText(2, .MaxRows, "1")
                        Call .SetText(3, .MaxRows, vJDate & "")
                        Call .SetText(4, .MaxRows, vJGbn & "")
                        Call .SetText(5, .MaxRows, vJNo & "")
                        Call .SetText(8, .MaxRows, vRegNo & "")
                        Call .SetText(9, .MaxRows, vName & "")
                        Call .SetText(10, .MaxRows, vSex & "")
                                                
                        If vEmer = "Y" Then
                            Call SpdForeBack(gfIFDisplayForm.spdIntList, 2, 4, .MaxRows, .MaxRows, RGB(0, 0, 0), ������)
                            Call .SetText(11, .MaxRows, "Y")
                            iERCnt = iERCnt + 1
                        Else
                            Call .SetText(11, .MaxRows, "N")
                        End If
                        
                        If vRerun = "Y" Then
                            Call SpdForeBack(gfIFDisplayForm.spdIntList, 2, 4, .MaxRows, .MaxRows, RGB(0, 0, 0), �帰�Ķ�)
                            Call .SetText(12, .MaxRows, "Y")
                            iERCnt = iERCnt + 1
                        Else
                            Call .SetText(12, .MaxRows, "N")
                        End If
                        
                        If iERCnt = 2 Then
                            Call SpdForeBack(gfIFDisplayForm.spdIntList, 2, 4, .MaxRows, .MaxRows, RGB(255, 255, 255), RGB(0, 0, 0))
                        End If
                        
                        Call .SetText(13, .MaxRows, vOther & "")
                        
                        'Order, Result�� �ʱ�ȭ
                        Call .SetText(14, .MaxRows, "N")
                        Call .SetText(15, .MaxRows, "N")
                        
                        Call .SetText(16, .MaxRows, Val(vIntcnt) & "")
                        
                        Erase vIFItemCd
                        ReDim vIFItemCd(CInt(Val(vIntcnt)))
                        
                        For j = 1 To CInt(Val(vIntcnt))
                            Call gfIFDisplayForm.spdList.GetText(11 + j, i, vIFItemCd(j))
                            Call .SetText(16 + j, .MaxRows, vIFItemCd(j) & "")
                        Next
                    End With
                End If
            End If
        Next

    End With
    
    If iCurIntCnt = 0 Then
        Exit Sub
    End If
    
    If iCurMaxRack = 0 And iCurMaxPos = 0 Then
        Call DisplayRackPos(1, gfIFDisplayForm.spdIntList.MaxRows, 1, 1)
    Else
        If iCurMaxPos < gIFPosInfo(iCurMaxRack).sPosMaxNo Then
            Call DisplayRackPos(iCurMaxRow + 1, gfIFDisplayForm.spdIntList.MaxRows, iCurMaxRack, iCurMaxPos + 1)
        ElseIf iCurMaxPos = gIFPosInfo(iCurMaxRack).sPosMaxNo Then
            Call DisplayRackPos(iCurMaxRow + 1, gfIFDisplayForm.spdIntList.MaxRows, iCurMaxRack + 1, 1)
        End If
    End If
    
    '99/09/29 ������ �߰� - MAXROW�� 12���� Ŭ��, ���� �������� ROW�� ���� ���̵���.
    If spdIntList.MaxRows > 12 Then
        spdIntList.TopRow = spdIntList.MaxRows - 12 + 1
    End If
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayIFList �����߻� - (" & Err.Number & ")"
    Call DispLogMsg("DisplayIFList �����߻�(" & Err.Description & ")")
End Sub

Private Sub DisplayRackPos(ByVal iSRow As Integer, ByVal iERow As Integer, ByVal iSRack As Integer, ByVal iSPos As Integer)
    Dim j%
    Dim i%
    Dim iPosSum%
    Dim iCurMaxRack%
    Dim iCnt%
    
    iCurMaxRack = 0
    iPosSum = 0
    iCnt = 0
    
    For i = iSRack To gIFRack.sMaxRack
        iPosSum = iPosSum + CInt(gIFPosInfo(i).sPosMaxNo)
        If (iERow - iSRow + 1) <= (iPosSum - iSPos) Then
            iCurMaxRack = i
            Exit For
        End If
    Next
    
    If iCurMaxRack = 0 Then
        MsgBox "Interface Worklist�� ���� Maxium Rack�� �������� ���� �۾��� �� �� �����ϴ�!!"
        Exit Sub
    End If
    
    With gfIFDisplayForm.spdIntList
        For j = iSRack To iCurMaxRack
            If j = iSRack Then
                For i = iSPos To gIFPosInfo(j).sPosMaxNo
                    If iCnt = iERow - iSRow + 1 Then
                        Exit For
                    Else
                        'Call .SetText(6, iSRow + iCnt, gIFPosInfo(j).sRackNo & "")
                        'Call .SetText(7, iSRow + iCnt, Format$(i, RackFormat(gIFRack.sPosDigit)) & "")
                        .Col = 6
                        .Row = iSRow + iCnt
                        .Text = gIFPosInfo(j).sRackNo & ""
                        
                        .Col = 7
                        .Row = iSRow + iCnt
                        .Text = Format$(i, RackFormat(gIFRack.sPosDigit)) & ""
                        iCnt = iCnt + 1
                    End If
                Next
            Else
                For i = 1 To gIFPosInfo(j).sPosMaxNo
                    If iCnt = iERow - iSRow + 1 Then
                        Exit For
                    Else
                        'Call .SetText(6, iSRow + iCnt, gIFPosInfo(j).sRackNo & "")
                        'Call .SetText(7, iSRow + iCnt, Format$(i, RackFormat(gIFRack.sPosDigit)) & "")
                        .Col = 6
                        .Row = iSRow + iCnt
                        .Text = gIFPosInfo(j).sRackNo & ""
                        
                        .Col = 7
                        .Row = iSRow + iCnt
                        .Text = Format$(i, RackFormat(gIFRack.sPosDigit)) & ""
                        iCnt = iCnt + 1
                    End If
                Next
            End If
        Next
    End With
End Sub

Private Sub DisplayOrderOK()
    On Error GoTo ErrHandler
    
    Dim i%
    
    If listTest.ListCount > 10 Then
        listTest.RemoveItem (0)
    End If
    
    listTest.AddItem gOrderTable.sJNo
    
    With spdIntList
        '�۾����ڸ� ����
        gOrderTable.sWDate = Format$(Now, "YYYYMMDD")
        '�۾��Ϸù�ȣ�� ����
        gOrderTable.sWSeq = Format$(Val(GetCurLastWSeq) + 1, "0000")
        
        '�ش���ڵ��� ���������� �ѱ�
        .MaxRows = .MaxRows + 1
        If .MaxRows > 500 Then
            .Row = 1
            .Action = ActionDeleteRow
            .MaxRows = .MaxRows - 1
        End If
        
        gOrderTable.iCRow = spdIntList.MaxRows
        
        If spdIntList.MaxRows > 10 Then
            spdIntList.TopRow = spdIntList.MaxRows - 9
        End If
        
        Call .SetText(1, gOrderTable.iCRow, gOrderTable.sWSeq & "")
        Call .SetText(2, gOrderTable.iCRow, "0")
        Call .SetText(3, gOrderTable.iCRow, gOrderTable.sJDate & "")
        Call .SetText(4, gOrderTable.iCRow, gOrderTable.sJGbn & "")
        Call .SetText(5, gOrderTable.iCRow, gOrderTable.sJNo & "")
        Call .SetText(6, gOrderTable.iCRow, gOrderTable.sRack & "")
        Call .SetText(7, gOrderTable.iCRow, gOrderTable.sPos & "")
        Call .SetText(8, gOrderTable.iCRow, gOrderTable.sRegNo & "")
        Call .SetText(9, gOrderTable.iCRow, gOrderTable.sName & "")
        Call .SetText(10, gOrderTable.iCRow, gOrderTable.sSex & "")
        Call .SetText(11, gOrderTable.iCRow, gOrderTable.sEmer & "")
        Call .SetText(12, gOrderTable.iCRow, gOrderTable.sReRun & "")
        Call .SetText(13, gOrderTable.iCRow, gOrderTable.sOther & "")
        Call .SetText(14, gOrderTable.iCRow, CStr(gOrderTable.iOrdCnt) & "")
        Call .SetText(15, gOrderTable.iCRow, "N")
        Call .SetText(16, gOrderTable.iCRow, CStr(gOrderTable.iOrdCnt) & "")
        
        '�˻��׸� ���� �����
        For i = 1 To gOrderTable.iOrdCnt
            Call .SetText(16 + i, gOrderTable.iCRow, gOrderTable.sIFSeq(i) & "|||")
        Next
        
        Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, gOrderTable.iCRow, gOrderTable.iCRow, _
                    RGB(0, 0, 0), �����)
        
        lblOrder = gOrderTable.sJNo
    End With
    
'    'Order ���� Local MDB�� Insert
'    Call RegOrder(1)
    
    'gOrderTable �ʱ�ȭ
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
    
    Exit Sub
ErrHandler:
    listTest.AddItem "Error"
    
    'gOrderTable �ʱ�ȭ
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
End Sub

Private Sub DisplayResultOK(ByVal iMode As Integer, ByVal sWDate As String, ByVal sWSeq As String, _
                        ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                        ByVal sRack As String, ByVal sPos As String, ByVal sRegNo As String, ByVal sName As String, _
                        ByVal sSex As String, ByVal sEmer As String, ByVal sReRun As String, ByVal sOther As String, _
                        ByVal iRstCnt As Integer, ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                        ByVal sIFSpcCd As String, ByVal sCurRow As String)
    On Error GoTo ErrHandler
    
    Dim sRetVal$, sCWSeq$, sChkVal$
    Dim i%, iCRow%
    Dim vWSeq, vJDate, vJGbn, vJNo, vTmp
        
    giAddKey = 0
    
    ReDim gResultTable(1)
    
    With gfIFDisplayForm
        Select Case iMode
            Case 0  'JDate, JGbn, JNo�� �ѱ�� ���
                .lblResult = sJDate & "-" & sJGbn & "-" & sJNo
                
                iCRow = FindIFListWithJ(sJDate, sJGbn, sJNo)
                
                If iCRow > 0 Then
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther) = "NO" Then
                           
                        Exit Sub
                    End If
                    
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                    
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case 1  'WSeq�� �ѱ�� ���
                .lblResult = sWDate & "-" & sWSeq
                
                iCRow = FindIFListWithW(sWSeq)
                
                If iCRow > 0 Then
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther) = "NO" Then
                           
                        Exit Sub
                    End If
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                        
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                                                
            Case 2  'CurRow�� �ѱ�� ��� - ��) �Һ���ⰰ�� �ܹ��� ���
                If .spdIntList.MaxRows >= CInt(sCurRow) Then
                    With .spdIntList
                        Call .GetText(1, CInt(sCurRow), vWSeq)
                        Call .GetText(3, CInt(sCurRow), vJDate)
                        Call .GetText(4, CInt(sCurRow), vJGbn)
                        Call .GetText(5, CInt(sCurRow), vJNo)
                    End With
                    
                    If Len(vJNo) > 0 Then
                        .lblResult = CStr(vJDate) & "-" & CStr(vJGbn) & "-" & CStr(vJNo)
                    Else
                        .lblResult = Format(Now, "YYYYMMDD") & "-" & CStr(vWSeq)
                    End If
                    
                    iCRow = CInt(sCurRow)
                    
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther) = "NO" Then
                           
                        Exit Sub
                    End If
            
            '����Ʈ ���� �����ϴ� ���� �ܹ����� ���
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                        
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
            
            Case 3  'JNo�� �ѱ�� ���
                .lblResult = sJNo
                
                iCRow = FindIFListWithJNo(sJNo)
                
                If iCRow > 0 Then
                    Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther)
                    
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        '�ϴ� Order�� ������ �Ѹ��� ����� ��Ÿ��
                        gOrderTable.sSampID = sJNo
                        Call Order_Input("N")
                        
                        iCRow = FindIFListWithJNo(sJNo)
                        
                        If iCRow > 0 Then
                            Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther)
                        Else
                            giAddKey = 1
                    
                            sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                        sRack, sPos, sRegNo, sName, _
                                        sSex, sEmer, sReRun, sOther, _
                                        iRstCnt, sIFRstCd, sRst1, sRst2, _
                                        sIFSpcCd, sCurRow)
                        End If
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case 4  '�۾�����Ʈ�� ������Ī
                With .spdIntList
                    If .MaxRows = 0 Then
                        iCRow = 0
                    Else
                        For i = 1 To .MaxRows
                            Call .GetText(15, i, vTmp)
                            
                            If vTmp = "N" Then
                                iCRow = i
                                Exit For
                            Else
                                iCRow = 0
                            End If
                        Next
                    End If
                End With
                
                If iCRow > 0 Then
                    With .spdIntList
                        Call .GetText(5, iCRow, vJNo)
                    End With
                    
                    .lblResult = CStr(vJNo)
                    
                    Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther)
                
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                    
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case Else
            
        End Select
        
        '������ ���Ե� �׸��� �����Ͽ� ��Ÿ��
        sChkVal = ChkCalResult(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        
        'Low, High ���� �����Ͽ� ������ ��Ÿ��
        sRetVal = ViewIFResult2(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
                
        '1024 * 768
        'If gResultTable(1).iCRow > 20 Then
        '    spdIntList.TopRow = gResultTable(1).iCRow - 19
        'End If
        
        '800 * 600
        If gResultTable(1).iCRow > 10 Then
            spdIntList.TopRow = gResultTable(1).iCRow - 9
        End If
        
    'gsTxMode="0" => Batch, gsTxMode="1" => RealTime(�� �׸�), gsTxMode="2" => RealTime(�� ȯ�ھ�)
        If gsTXMode = "0" Then
        ElseIf gsTXMode = "1" Then
            If sRetVal = "NONE" Then
            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
                If sChkVal = "1" Then
                    Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
                Else
                    Call RegResult(0, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
                End If
                
                If giAddKey = 1 Then
                    If sCWSeq = "" Then
                    Else
                        gsLastWSeq = sCWSeq
                    End If
                End If
                
                If sRetVal = "DONE" Then
                    Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
                         gResultTable(1).iCRow, RGB(0, 0, 0), ���ʷ�)
                End If
            End If
        ElseIf gsTXMode = "2" Then
        '���ϴ� �����Ϲ�Ĵ�� ���� ������.
            If sRetVal = "NONE" Then
            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
                'ȯ�ڴ����� ��� ��� �� ���
                If giAddKey = 1 Then
                    If sCWSeq = "" Then
                    Else
                        gsLastWSeq = sCWSeq
                    End If
                End If
                
                Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
                
                Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
                         gResultTable(1).iCRow, RGB(0, 0, 0), ���ʷ�)
            End If
        End If
        
    '<----- �߾Ӻ����� CREAó������ Ư���� ���� ��� - sCurRow = 0(�̵��), sCurRow = 1(���)
        If sCurRow = "0" Then
            Exit Sub
        End If
    '<----- �߾Ӻ����� CREAó������ Ư���� ���� ��� - sCurRow = 0(�̵��), sCurRow = 1(���)
    
'        If optRegOpt(0).Value = True Then
        If giTestMode = 78 Then
        Else
            Call RegServerOK(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2)
        End If
'        End If
        
        Erase gResultTable
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayResultOK �����߻� - ( " & Err.Description & " )"
    Call DispLogMsg("DisplayResultOK �����߻�(" & Err.Description & ")")
End Sub
Private Function ExistIFList(ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, Optional ByVal sWSeq As String)
    Dim i%
    Dim vJDate, vJGbn, vJNo, vWSeq
    
    ExistIFList = "NO"
    
    With gfIFDisplayForm.spdIntList
        For i = 1 To gfIFDisplayForm.spdIntList.MaxRows
            If sWSeq = "" Then
                Call .GetText(3, i, vJDate)
                Call .GetText(4, i, vJGbn)
                Call .GetText(5, i, vJNo)
            
                If CStr(vJDate) = sJDate And _
                   CStr(vJGbn) = sJGbn And _
                   CStr(vJNo) = sJNo Then
                   
                   ExistIFList = "YES"
                   
                   Exit For
                End If
            Else
                Call .GetText(1, i, vWSeq)
                Call .GetText(3, i, vJDate)
                Call .GetText(4, i, vJGbn)
                Call .GetText(5, i, vJNo)
                
                If CStr(vJDate) = sJDate And _
                   CStr(vJGbn) = sJGbn And _
                   CStr(vJNo) = sJNo And _
                   CStr(vWSeq) = sWSeq Then
                   
                   ExistIFList = "YES"
                   
                   Exit For
                End If
            End If
        Next
    End With
End Function

Private Function ExistEmerIFList(ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String)
    Dim i%
    Dim vJDate, vJGbn, vJNo, vEmer
    
    ExistEmerIFList = "NO"
    
    With gfIFDisplayForm.spdIntList
        For i = 1 To .MaxRows
            Call .GetText(2, i, vJDate)
            Call .GetText(3, i, vJGbn)
            Call .GetText(4, i, vJNo)
            Call .GetText(12, i, vEmer)
            
            If CStr(vJDate) = sJDate And _
               CStr(vJGbn) = sJGbn And _
               CStr(vJNo) = sJNo And _
               CStr(vEmer) = "Y" Then
               
               ExistEmerIFList = "YES"
               
               Exit For
            End If
        Next
    End With
End Function

Private Function Find_ChkRow(ByVal iNum As Integer) As Integer
    On Error GoTo ErrHandler
    
    Dim i%
    Dim cmpNum%
    
    With gfIFDisplayForm.spdIntList
        .Col = 2
        
        For i = 1 To .MaxRows
            .Row = i
            
            If .Text = "1" Then
                cmpNum = cmpNum + 1
                
                If cmpNum = iNum Then
                    Find_ChkRow = i
                    Exit Function
                End If
            End If
        Next
    End With
    
    Exit Function
ErrHandler:
    ViewMsg "Find_ChkRow �����߻� - (" & CStr(Err.Number) & ")"
    Call DispLogMsg("Find_ChkRow �����߻�(" & Err.Description & ")")
End Function

Private Sub InvisibleBatch()
    On Error GoTo ErrHandler
    
    If gsTXMode = "0" Then
    '���� Batch������� ��������ϴ� ���
    Else
        gfIFDisplayForm.cmdReg.Visible = False
        
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    End If
    
    Exit Sub
    
ErrHandler:
    ViewMsg "InvisibleBatch �����߻� - (" & CStr(Err.Number) & ")"
End Sub


Private Sub SetSpdIntLIstColHidden()
    Dim i%
    
    With gfIFDisplayForm.spdIntList
        If gRstcfg.sUse = "1" Then
            For i = 1 To MAXORDERFIELD - 1
                If i < 4 Then
                    .Col = i + 2
                Else
                    .Col = i + 4
                End If
                
                If gRstcfg.sFUse(i) And Val(gRstcfg.sFSize(i)) > 0 Then
                    .ColHidden = False
                Else
                    .ColHidden = True
                End If
            Next
        Else
            If gOrdCfg.sUse = "1" Then
                For i = 1 To MAXRESULTFIELD - 3
                    If i < 4 Then
                        .Col = i + 2
                    Else
                        .Col = i + 4
                    End If
                    
                    If gOrdCfg.sFUse(i) And Val(gOrdCfg.sFSize(i)) > 0 Then
                        .ColHidden = False
                    Else
                        .ColHidden = True
                    End If
                Next
            Else
                MsgBox "ȯ�漳������ Result Setting�� �����Ͻʽÿ�!!", vbCritical
                Exit Sub
            End If
        End If
        
        If giTotIFItemCnt = 0 Then
            .MaxCols = 16
        Else
            .MaxCols = 16 + 2 * giTotIFItemCnt
            
            For i = 17 To 17 + giTotIFItemCnt - 1
                .Col = i
                .ColHidden = True
            Next
        End If
    End With
End Sub

Private Sub cmdClear2_Click()
'    If MsgBox("Interface List�� �����ϸ� �ش� List�� ����� ���� ���մϴ�." & vbCrLf & _
'        "����� ���� ���� �ʾҴٸ� '�ƴϿ�'�� �����Ͻʽÿ�." & vbCrLf & _
'        "Interface List�� ���� �����Ͻðڽ��ϱ�?", vbYesNo, "Interface List ��� ���� Ȯ��") = vbYes Then
    If spdIntList.MaxRows > 0 Then
        spdIntList.MaxRows = 0
        
        With spdRst
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = -1
            .Col2 = -1
            .Action = SS_ACTION_CLEAR_TEXT
            .BlockMode = False
        End With
        
        With spdRst2
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = -1
            .Col2 = -1
            .Action = SS_ACTION_CLEAR_TEXT
            .BlockMode = False
        End With
        
        lblResult.Caption = ""
        lblOrder.Caption = ""
        lblCSelList.Caption = ""
        
        RegIFStateFlag "SampleCnt", "0"
    End If
End Sub

Private Sub cmdClearLog_Click()

    lstLog.Clear

End Sub

Private Sub cmdConnect_Click()

    'WinSock ����
    Call ConnectWinSock
    
'    txtState = txtState & tcpClient(1).State
    
    If gsINITMode = "1" Then
        Dim iRetryCnt   As Integer
        iRetryCnt = 0
        Do While tcpClient(1).State = 7
            If iRetryCnt > 10 Then
                MsgBox "Error - Connect Timeout", vbExclamation
                Call DispLogMsg("Error - Connect Timeout")
                End
                Exit Sub
            End If
            Sleep 100
            iRetryCnt = iRetryCnt + 1
        Loop

        Sleep 100
        
        Call cmdInitial_Click
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub cmdInitial_Click()
    Call CommOut_ConnectionMsg
End Sub



Private Sub cmdStart_Click()
    
'    txtState = txtState & tcpClient(1).State
    
    Call cmdInitial_Click
    Timer1.Enabled = True

'    txtState = txtState & tcpClient(1).State
    
End Sub

Private Sub cmdTest_Click()

    msWkbuf = txtTest
    
    Call PhaseCfg_Protocol_Integra800
    
End Sub

Private Sub Comm1_OnComm()
'    Select Case Comm1.CommEvent
'       ' Events
'        Case MSCOMM_EV_SEND     ' There are SThreshold number of
'                               ' character in the transmit buffer.
'        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
'
'            msWkbuf = Comm1.Input
'
'            If giTestMode = 2 Then
'                Print #1, msWkbuf;    'Log File�� ����
'            End If
'
'            If miSpaceCnt = 30 Then
'                miSpaceCnt = 0
'            End If
'
'            miSpaceCnt = miSpaceCnt + 2
'
'            ViewMsg Space(miSpaceCnt) & "���� Interface �۾� ��..."
'
'            Call PhaseCfg_Protocol_Integra800
'
'        Case MSCOMM_EV_CTS      'j
'        Case MSCOMM_EV_DSR      ' Change in the DSR line.
'        Case MSCOMM_EV_CD       ' Change in the CD line.
'        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
'        ' Errors
'        Case MSCOMM_ER_BREAK    ' A Break was received.
'        ' Code to handle a BREAK goes here, and so on.
'        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
'        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
'        Case MSCOMM_ER_FRAME    ' Framing Error.
'        Case MSCOMM_ER_OVERRUN  ' Data Lost.
'        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
'        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
'        Case MSCOMM_ER_RXPARITY ' Parity Error.
'        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
'    End Select
End Sub



Private Sub Command1_Click()

    lstOrder.Clear
    lstResult.Clear
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrLoad
    
    Dim sUseYN$
    Dim bRetVal As Boolean
        
    '����ڵ�/����/IP/PORT ����
    Call SetIFProgramInfo
        
    
    miHlpClick = 0
    miSpaceCnt = 0
    
    Call RegViewMsgHwnd(Me.StatusBar1.hwnd)
    Call GetMachineInfo
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    Call GetTestMode
    Call GetCSMode
        
    Call DisplayInit
    Call DisplayInitItem
    
'    Call PortOpen
    Call LogFileOpen
       
    miPhase = 1
    
    miTimerFlag = 0
    miConnectFlag = 0
    miOrderFlag = 0
    miResultFlag = 0
    miTimerCnt = 0
    miResultCnt = 0
    miOrdRstCnt = 0
    
    miResultTimerCnt = 0
    miOrderTimerCnt = 0
    
    miPendOrderCnt = 0      '2004/1/8 yk
    
    OrdState = 0
    RstState = 0
    
    Tab1.Tab = 0
    
    miTimerCnt1 = 0         '2004/1/28 yk
    miTimerCnt2 = 0
    miTimerCnt3 = 0
    miOrdFlag = 0
    miRstFlag = 0
        
'<----- �����߾Ӻ����� Tuxedo ���� �ڵ� ----->
    Dim sError As String
    Dim sNull  As String

    sError = Space$(1024)
    sNull = Space$(1024)

    'Session �ʱ�ȭ
    If OpenSession(sError) <> 0 Then
        MsgBox "OpenSession - " & Trim$(sError), vbDefaultButton1, "Top End Error"
        End
    End If

    sError = Space$(1024)

    'Client Init
    If ClientInit(sError) <> 0 Then
        Call CloseSession(sNull)
        MsgBox "OpenSession - " & Trim$(sError), vbDefaultButton1, "Top End Error"
        End
    End If
'<----- �����߾Ӻ����� Tuxedo ���� �ڵ� ----->

    ViewMsg "Interface Program Ready..."
    Call DispLogMsg("Interface Program Ready...")
    
    Call GetLastWorkSeq(Format(Now, "YYYYMMDD"))
    
    If giTestMode = 77 Then
        txtTest.Visible = True
        cmdTest.Visible = True
    ElseIf giTestMode = 777 Then
        pnlTest.Visible = True
    End If
        
'    'WinSock ����
'    Call ConnectWinSock
'
'    If gsINITMode = "1" Then
'        Dim iRetryCnt   As Integer
'        iRetryCnt = 0
'        Do While tcpClient(1).State = 7
'            If iRetryCnt > 10 Then
'                MsgBox "Error - Connect Timeout", vbExclamation
'                Unload Me
'                Exit Sub
'            End If
'            Sleep 500
'            iRetryCnt = iRetryCnt + 1
'        Loop
'
'        Call cmdInitial_Click
'        Timer1.Enabled = True
'    End If
    Call cmdConnect.DoClick
    
    
    '������Ʈ���� HWnd ���
    Call SetProgHWnd(Me.hwnd)

    
    Exit Sub
ErrLoad:
    MsgBox Err.Description, vbExclamation
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call PortClose
    Call LogFileClose
    
    RegEditCurFrmTitle "IF", ""
  
    Close #301
    
    Call SetProgHWnd(0)
    
    '<----- �����߾Ӻ����� Tuxedo ���� �ڵ� ----->
    Dim sNull As String

    sNull = Space$(1024)

    Call CloseSession(sNull)
    '<----- �����߾Ӻ����� Tuxedo ���� �ڵ� ----->
    
End Sub

Private Sub cmdExit_Click()
    If MsgBox("[" & UCase$(gsMachineNm) & "]" & " Interface Program�� �����Ͻðڽ��ϱ�?" & vbCrLf & vbCrLf & _
            "Interface �۾� ���߿� ������ ��� ���۵����Ͱ� �ս��� �˴ϴ�.", vbYesNo, _
            "Interface ���� Ȯ��") = vbYes Then
            
        miHlpClick = 1
        Unload Me
    End If
End Sub

Private Sub listNoOrd_Click()
    listNoOrd.ToolTipText = listNoOrd.List(listNoOrd.ListIndex)
End Sub

Private Sub listNoOrd_DblClick()
    listNoOrd.RemoveItem (listNoOrd.ListIndex)
End Sub

Private Sub listNoOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        listNoOrd.Clear
    End If
End Sub

Private Sub mnuPopup01_Click()
    Dim i%
    
    With spdIntList
        For i = miBSRow To miBERow
            .Row = miBSRow
            .Action = SS_ACTION_DELETE_ROW
            .MaxRows = spdIntList.MaxRows - 1
        Next
        
        .Action = SS_ACTION_DESELECT_BLOCK
    End With
    
End Sub

Private Sub spdIntList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    If BlockRow = -1 And BlockRow2 = -1 Then
        miBSRow = 1
        miBERow = spdIntList.MaxRows
    Else
        miBSRow = CInt(BlockRow)
        miBERow = CInt(BlockRow2)
    End If
End Sub

Private Sub spdIntList_Change(ByVal Col As Long, ByVal Row As Long)
    Dim vRack
    Dim vPos
    
    'Rack
    If Col = 6 Then
        With spdIntList
            If giAddKey = 1 Then
                giAddKey = 0
            Else
                Call .GetText(Col, Row, vRack)
                
                If IsNumeric(vRack) = False Then
                    MsgBox "Rack�� ������ �߸� �Ǿ����ϴ�!!"
                    Exit Sub
                End If
                
                If LenH(vRack) <= gIFRack.sRackDigit Then
                    Call .SetText(Col, Row, Format(vRack, RackFormat(gIFRack.sRackDigit)))
                    Call .GetText(7, Row, vPos)
                    Call DisplayRackPos(Row, .MaxRows, CInt(vRack), CInt(vPos))
                ElseIf LenH(vRack) > gIFRack.sRackDigit Then
                    MsgBox "Rack�� ������ �߸� �Ǿ����ϴ�!!"
                    Exit Sub
                End If
            End If
        End With
    End If
    
    'Pos
    If Col = 7 Then
        With spdIntList
            If giAddKey = 1 Then
                giAddKey = 0
            Else
                Call .GetText(Col, Row, vPos)
                
                If IsNumeric(vPos) = False Then
                    MsgBox "Pos�� ������ �߸� �Ǿ����ϴ�!!"
                    Exit Sub
                End If
                
                If LenH(vPos) <= gIFRack.sPosDigit Then
                    Call .SetText(Col, Row, Format(vPos, RackFormat(gIFRack.sPosDigit)))
                    Call .GetText(6, Row, vRack)
                    Call DisplayRackPos(Row, .MaxRows, CInt(vRack), CInt(vPos))
                ElseIf LenH(vPos) > gIFRack.sPosDigit Then
                    MsgBox "Pos�� ������ �߸� �Ǿ����ϴ�!!"
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub spdIntList_Click(ByVal Col As Long, ByVal Row As Long)
    Call DisplayResult2(CInt(Row))
End Sub

Private Sub spdIntList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If MsgBox("�ش� Interface List�� �����Ͻðڽ��ϱ�?" & vbCrLf & _
        "������ Interface List�� ����� ���� �� �����ϴ�. ��� �Ͻðڽ��ϱ�?", _
        vbYesNo, "�ش� Interface List ���� Ȯ��") = vbYes Then
        
        With spdIntList
            .Row = Row
            .Action = SS_ACTION_DELETE_ROW
            .MaxRows = spdIntList.MaxRows - 1
            
        End With
        
    End If
End Sub

Private Sub spdIntList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If spdIntList.IsBlockSelected = True Then
        Call PopupMenu(mnuPopup)
    End If
End Sub

Private Sub tcpClient_Close(Index As Integer)
    '2004/1/15 yk
    If tcpClient(Index).State = 9 Then
        tcpClient(Index).Close
        Timer2.Enabled = True
    End If
End Sub

Private Sub tcpClient_Connect(Index As Integer)
    '2004/1/15 yk
    Timer2.Enabled = False
End Sub


Private Sub tcpClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo ErrRtn
    
    tcpClient(1).GetData msWkbuf

    If giTestMode = 77 Then
        Print #1, msWkbuf;    'Log File�� ����
    End If
    
    If miSpaceCnt = 30 Then
        miSpaceCnt = 0
    End If
    
    miSpaceCnt = miSpaceCnt + 2
    
    ViewMsg Space(miSpaceCnt) & "���� Interface �۾� ��..."
    
    Call PhaseCfg_Protocol_Integra800
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg Err.Description
        Call DispLogMsg(Err.Description)
    End If
End Sub

Private Sub tcpClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    ViewMsg Number & ":" & Description
    Call DispLogMsg(Number & ":" & Description)
    
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrHandler
    
    miTimerCnt1 = miTimerCnt1 + 1   'Order
    miTimerCnt2 = miTimerCnt2 + 1   'Result
    miTimerCnt3 = miTimerCnt3 + 1   'Line Check
    
    'Connection Check
    If miTimerCnt3 > 15 Then
        Call CommOut_ConnectionMsg
        miTimerCnt3 = 0
        
        Exit Sub
    End If
    
    If miTimerFlag = 1 Then
        Exit Sub
    End If
    
    'Order/Result Timer Count
    If miTimerCnt1 >= 15 And miOrdFlag = 1 Then
        miTimerCnt1 = 0
        miOrdFlag = 0
    End If
    If miTimerCnt2 >= 15 And miRstFlag = 1 Then
        miTimerCnt2 = 0
        miRstFlag = 0
    End If
    
    'Pending SampleID Request
    If miOrdFlag = 0 Then
        Call CommOut_RequestPendingMsg
        miOrdFlag = 1
        
        '--- test
        If giTestMode = 777 Then
            Call DispTimer(1)
        End If
        '========
        
        Exit Sub
    End If
    
    'Pending Result Request
    If miRstFlag = 0 Then
        Call CommOut_RequestResultMsg
        miRstFlag = 1
        
        '--- test
        If giTestMode = 777 Then
            Call DispTimer(2)
        End If
        '========
        
        Exit Sub
    End If
    
ErrHandler:
    If Err <> 0 Then
        miOrdFlag = 0
        miRstFlag = 0
    End If
End Sub
Private Sub DispTimer(ByVal iPara As Integer)
    On Error GoTo ErrTest
    
    Dim sTmp    As String
    Dim iTmp    As Integer
    
    iTmp = 0
    
    If iPara = 1 Then
        With lstOrder
            If .ListCount > 0 Then
                sTmp = .List(.ListCount - 1)
                iTmp = DateDiff("S", Left(sTmp, 8), Format(Now, "HH:MM:SS"))
            End If
            
            .AddItem Format(Now, "HH:MM:SS") & " / " & Trim(iTmp)
            
            If .ListCount > 200 Then
                .RemoveItem 0
            End If
        End With
    ElseIf iPara = 2 Then
        With lstResult
            If .ListCount > 0 Then
                sTmp = .List(.ListCount - 1)
                iTmp = DateDiff("S", Left(sTmp, 8), Format(Now, "HH:MM:SS"))
            End If
            
            .AddItem Format(Now, "HH:MM:SS") & " / " & Trim(iTmp)
            
            If .ListCount > 200 Then
                .RemoveItem 0
            End If
        End With
    End If
    
ErrTest:
    If Err <> 0 Then
        ViewMsg Err.Description
    End If
End Sub
Private Sub Timer2_Timer()

    '2004/1/15 yk
    tcpClient(1).Close
    Unload tcpClient(1)
    Call ConnectWinSock
    
End Sub


