VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWorklist 
   Caption         =   "CLINILOG Order"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9495
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   9495
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1035
      Left            =   1440
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   3405
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   270
         Top             =   255
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   1320
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         SThreshold      =   1
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   810
         Top             =   330
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   5000
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "test"
      Height          =   465
      Left            =   690
      TabIndex        =   43
      Top             =   3180
      Visible         =   0   'False
      Width           =   1185
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   2775
      Left            =   6210
      TabIndex        =   42
      Top             =   3990
      Visible         =   0   'False
      Width           =   3195
      _Version        =   393216
      _ExtentX        =   5636
      _ExtentY        =   4895
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":0442
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   555
      Left            =   2700
      TabIndex        =   41
      Top             =   3300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   6150
      TabIndex        =   38
      Top             =   150
      Width           =   3225
      _Version        =   65536
      _ExtentX        =   5689
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   39
         Top             =   90
         Width           =   1875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   165
         Width           =   900
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   1065
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   37
      Top             =   8040
      Width           =   9255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   435
      Left            =   2490
      TabIndex        =   36
      Top             =   4650
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   465
      Left            =   9390
      TabIndex        =   8
      Top             =   3465
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   9615
      TabIndex        =   35
      Top             =   4050
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "SP"
      Height          =   285
      Left            =   11040
      TabIndex        =   34
      Top             =   690
      Value           =   1  '확인
      Width           =   615
   End
   Begin VB.TextBox txtReOrd 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9510
      TabIndex        =   32
      Top             =   90
      Width           =   2085
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   150
      TabIndex        =   31
      Top             =   7650
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9150
      TabIndex        =   30
      Top             =   7590
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdPortOpen 
      Caption         =   "PortOpen"
      Height          =   525
      Left            =   3360
      TabIndex        =   6
      Top             =   6870
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "대기검체전송"
      Height          =   525
      Left            =   9540
      TabIndex        =   4
      Top             =   2190
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "   수동 오더 전송   "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   9420
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
      Begin VB.ComboBox cboWorkStation 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   27
         Top             =   2010
         Width           =   1845
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "Order"
         Height          =   555
         Left            =   720
         TabIndex        =   29
         Top             =   2340
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조  회"
         Height          =   525
         Left            =   90
         TabIndex        =   28
         Top             =   2550
         Width           =   2025
      End
      Begin VB.TextBox txtSeqNo1 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   26
         Text            =   "1"
         Top             =   2700
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtSeqNo2 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   24
         Text            =   "1000"
         Top             =   2700
         Visible         =   0   'False
         Width           =   705
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   300
         TabIndex        =   19
         Top             =   780
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   100663297
         CurrentDate     =   38583
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   300
         TabIndex        =   21
         Top             =   1170
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   100663297
         CurrentDate     =   38583
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈WorkStation"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1110
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈작업번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   22
         Top             =   2940
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1950
         TabIndex        =   20
         Top             =   840
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈작업일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫  기"
      Height          =   525
      Left            =   9540
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdReLoad1 
      Caption         =   "ReLoad"
      Height          =   255
      Left            =   5100
      TabIndex        =   13
      Top             =   3210
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdReLoad 
      Caption         =   "ReLoad"
      Height          =   255
      Left            =   5010
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   885
   End
   Begin FPSpread.vaSpread vasOrder 
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
      _ExtentY        =   6800
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":4915
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
      _ExtentY        =   4683
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":62A8
   End
   Begin VB.CheckBox chkReal 
      Caption         =   "자동오더전송"
      Height          =   255
      Left            =   9540
      TabIndex        =   5
      Top             =   720
      Value           =   1  '확인
      Width           =   1425
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   525
      Left            =   9540
      TabIndex        =   3
      Top             =   1620
      Width           =   2055
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   6345
      Left            =   6150
      TabIndex        =   2
      Top             =   960
      Width           =   3225
      _Version        =   393216
      _ExtentX        =   5689
      _ExtentY        =   11192
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":66FF
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   6180
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "처방 받기"
      Height          =   525
      Left            =   9540
      TabIndex        =   0
      Top             =   1050
      Width           =   2055
   End
   Begin FPSpread.vaSpread vasList1 
      Height          =   1935
      Left            =   60
      TabIndex        =   16
      Top             =   8010
      Visible         =   0   'False
      Width           =   11655
      _Version        =   393216
      _ExtentX        =   20558
      _ExtentY        =   3413
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":69F1
   End
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   45
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   100663297
      CurrentDate     =   40205
   End
   Begin MSComCtl2.DTPicker dtpEDate 
      Height          =   315
      Left            =   3270
      TabIndex        =   46
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   100663297
      CurrentDate     =   40205
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   2775
      Left            =   4890
      TabIndex        =   44
      Top             =   3330
      Visible         =   0   'False
      Width           =   3195
      _Version        =   393216
      _ExtentX        =   5636
      _ExtentY        =   4895
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":843D
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "MSG"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   7680
      Width           =   420
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "※ 조회일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   48
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3030
      TabIndex        =   47
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblWinsock 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "winsock 상태:"
      Height          =   180
      Left            =   9540
      TabIndex        =   33
      Top             =   3390
      Width           =   1185
   End
   Begin VB.Label lblBarCode 
      AutoSize        =   -1  'True
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   7230
      TabIndex        =   11
      Top             =   735
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6180
      TabIndex        =   10
      Top             =   735
      Width           =   900
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "menu"
      Begin VB.Menu subConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu subCode 
         Caption         =   "검사코드설정"
      End
      Begin VB.Menu subN 
         Caption         =   "-"
      End
      Begin VB.Menu subClose 
         Caption         =   "닫기"
      End
   End
End
Attribute VB_Name = "frmWorklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2010.03.16 이상은 - 타이머 시간 조정(5000 -> 10000)

Dim lsTotRsv As String
Dim sckState As Integer
Dim MyPort As String



Dim iState As Integer
Dim bTimer As Boolean
Dim flgETB As Boolean

Dim ReceiveData As String

Dim cntCheckSum      As Integer

Private Type typeOrder
    FormatTypeCode          As String * 4
    SampleID                As String * 20
    DateOfReception         As String * 8
    PatientID               As String * 10
    PatientNameABC          As String * 20
    PatientNameReserve      As String * 20
    Birthday                As String * 8
    Sex                     As String * 1
    '
    DeleteFlag              As String * 1
    WardCode                As String * 4
    WardName                As String * 16
    OrderDeptCode           As String * 4
    OrderDeptName           As String * 16
    OrderDrCode             As String * 4
    OrderDrName             As String * 16
    '
    TypeOfContainer         As String * 2
    TypeOfSample            As String * 2
    HeightOfSample          As String * 2
    DeCapping               As String * 2
    Centrifuge              As String * 2
    STATFlag                As String * 2
    '
    FreeComment             As String * 32
    '
    'NumberOfTest            As String * 4
    NumberOfTest            As String
    ItemNo(1 To 100)        As String * 10
    TypeOfOrder(1 To 100)   As String * 2
    NoOfAddInf(1 To 100)    As String * 2
    TypeOfAddInf(1 To 100)  As String * 2
    AdditionalInf(1 To 100) As String * 10
End Type

Dim typOrder As typeOrder

Private Type typeResult
    TestDate                As String
    TestTime                As String
    '
    FormatTypeCode          As String * 4
    TypeOfSample            As String * 2
    SampleID                As String * 20
    PatientID               As String * 10
    RackID                  As String * 10
    RackPosition            As String * 2
    NoOfAnalyzer            As String * 2
    '
    SampleInfCyle           As String * 4
    SampleInfHb             As String * 4
    SampleInfBil            As String * 4
    '
    NoOfItems               As String * 4
    ItemNo(1 To 100)        As String * 7
    Result(1 To 100)        As String * 10
    Comment(1 To 100)       As String * 4
    DilutionRatio(1 To 100) As String * 2
    ConfirmFlag(1 To 100)   As String * 2
    '
    LengthOfFreeComment     As String * 4
    FreeComment             As String * 32
End Type
Dim typResult As typeResult

Dim SendCursor As Integer


Private Sub chkReal_Click()
'    If chkReal.Value = 0 Then
'        If MsgBox("실시간으로 받은 처방을 장비로 전송할 수 없습니다" & vbCrLf & vbCrLf & "실시간 전송을 원하십니까? ", vbCritical + vbYesNo + vbDefaultButton1, "알림") = vbYes Then
'            chkReal.Value = 1
'        Else
'            chkReal.Value = 0
'        End If
'    End If
End Sub

Private Sub cmdClear_Click()
    ClearSpread vasExam
    ClearSpread vasList
    ClearSpread vasOrder
    ClearSpread vasList1
    lblBarCode = ""
    gRow = 1
End Sub

Private Sub cmdClose_Click()
    If MsgBox("프로그램을 종료하면 검사 처방 받기 및 장비로 처방 전송이 되지 않습니다" & vbCrLf & vbCrLf & "프로그램을 종료하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "종료알림") = vbNo Then
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdListen_Click()
    lsTotRsv = ""
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.LocalPort = MyPort
    Winsock1.Listen
    
    sckState = 0
    
End Sub

Private Sub cmdOrder_Click()
Dim lsID As String
Dim i, j As Integer
Dim mExam As Variant
Dim AdoRs_Exam As ADODB.Recordset
Dim Ord(7) As String
Dim lsOrder As String
Dim lRow1, lRow As Long
Dim lsWkNo As String
Dim lsPID As String

Dim lsSlideOrd As String
Dim lsExamDate As String


If MSComm1.PortOpen = False Then
    LASCPortOpen
End If
    
lRow1 = 1
Do While lRow1 <= vasList1.DataRowCnt
    vasList1.Row = lRow1
    vasList1.Col = 1
    
    
    lsSlideOrd = ""
    lsWkNo = ""
    
    If vasList1.Value = 0 Then
        For i = 1 To 7
            Ord(i) = "0"
        Next i
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then
            vasList.MaxRows = lRow
        End If
        
        lsID = Trim(GetText(vasList1, lRow1, 3))
        
        SetText vasList, lsID, lRow, 1
        SetText vasList, "A", lRow, 2
        SetText vasList, Trim(GetText(vasList1, lRow1, 4)), lRow, 3
        SetText vasList, Trim(GetText(vasList1, lRow1, 5)), lRow, 4
        SetText vasList, GetDateFull, lRow, 5
                
        mExam = Get_OrderBody(lsID)
        If Not IsNull(mExam) Then
            SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
            Set AdoRs_Exam = db_select_rs(gLocal, SQL)
            
            ClearSpread vasExam
            For j = LBound(mExam, 2) To UBound(mExam, 2)
                SetText vasExam, mExam(3, j), j + 1, 1
                SetText vasExam, mExam(4, j), j + 1, 2
                
                lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & SetSpace(Trim(mExam(6, LBound(mExam, 2))), 3)
                'lsPID = SetSpace(mExam(1, LBound(mExam, 2)), 8)
                lsPID = SetSpace(Left(mExam(1, LBound(mExam, 2)), 8), 8)
                
                CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(GetDateFull, 10)
                If Not IsNumeric(gPatGen.Age) Then
                    SetText vasList, gPatGen.Age, lRow, 7
                Else
                    SetText vasList, "", lRow, 7
                End If
                
                If Not AdoRs_Exam Is Nothing Then
                    AdoRs_Exam.MoveFirst
                    Do Until AdoRs_Exam.EOF
                        If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                            Select Case Trim(AdoRs_Exam("OrdGubun"))
                            Case "C": Ord(1) = "1"
                            Case "D": Ord(2) = "1"
                            Case "R": Ord(3) = "1"
                            Case "P"
                                Ord(4) = "1"
                                lsSlideOrd = "SP"
                            Case "S"
                                Ord(5) = "1"
                                lsSlideOrd = "SC"
'                            Case "P": Ord(4) = "1"
'                            Case "S": Ord(5) = "1"
                            Case "X": Ord(6) = "1"
                            Case "B"
                                If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                    Ord(7) = "1"
                                End If
                            End Select
                            
                            Exit Do
                        End If
                        
                        AdoRs_Exam.MoveNext
                    Loop
                End If
            Next j
            
            lsOrder = ""
            For i = 1 To 7
                lsOrder = lsOrder & Ord(i)
            Next i
            
            'MsgBox lsOrder
            
            If lsOrder <> "0000000" And lsOrder <> "" Then
                lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                'lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & lsWkNo & lsPID
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "000****************************************" & chrETX
                
                DoSleep 100
                
                MSComm1.Output = lsOrder
                
                SQL = "Select barcode from res_flag where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = lsID Then
                    SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "' " & vbCrLf & _
                          "where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                    res = SendQuery(gLocal, SQL)
                Else
                    SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                          "Values ('" & lsExamDate & "', '" & lsID & "', '', '', '', '', " & _
                          "'', '', '', '', '', '', " & _
                          "'', '', '', '', '', '" & lsSlideOrd & "' ) "
                    res = SendQuery(gLocal, SQL)
                End If
                
                
                SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
                      "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','','', 0) "
                res = SendQuery(gLocal, SQL)
                If res = 1 Then
                    SetText vasList, "B", lRow, 2
                    CopyRecord lRow
                    DeleteRow vasList1, lRow1, lRow1
                Else
                
                    lRow1 = lRow1 + 1
                    SaveQuery SQL
                    'Exit Function
                End If
            End If
        Else
            lRow1 = lRow1 + 1
        End If
    Else
        lRow1 = lRow1 + 1
    End If
Loop


End Sub

Private Sub cmdPortOpen_Click()
    If MSComm1.PortOpen = True Then
        If MsgBox("실시간으로 받은 처방을 장비로 전송할 수 없습니다" & vbCrLf & vbCrLf & "실시간 전송을 원하십니까? ", vbCritical + vbYesNo + vbDefaultButton1, "알림") = vbYes Then
            Exit Sub
        End If
        
        MSComm1.PortOpen = False
        cmdPortOpen.Caption = "PortOpen"
        
        lblMsg.Caption = "[Message] 포트가 닫혀있습니다"
        chkReal.Value = 0
        Timer1.Enabled = False
    Else
        cmdPortOpen.Caption = "PortClose"
        LASCPortOpen
        
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdReLoad_Click()

'    Call SendTheDataCLINILOG("12060102599", False, "20120601")

    ClearSpread vasList
    
    SQL = "Select barcode, OrdFlag, PID, PName, ReceDate,RemoteIP  from WorkList where OrdFlag = 'A' "
    res = db_select_Vas(gLocal, SQL, vasList)
    
End Sub

Private Sub cmdReLoad1_Click()
    Dim lsReceDate As String
    
    lsReceDate = Format(CDate(GetDateFull), "yyyymmdd")
    
    ClearSpread vasOrder
    
    'SQL = "Select barcode, OrdFlag, PID, PName, ReceDate from WorkList where recedate >= '" & lsReceDate & " 00:00:00' and recedate <= '" & lsReceDate & " 23:59:59' and OrdFlag <> 'A' """
    SQL = "Select barcode, OrdFlag, PID, PName, ReceDate, RemoteIP from WorkList  where recedate = '" & lsReceDate & "' "
    res = db_select_Vas(gLocal, SQL, vasOrder)
    If res = -1 Then
        SaveQuery SQL
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim lsDate As String
    Dim liOrdNo As Integer
    Dim i, j, k, n
    Dim lRow, lCol As Long
    
    Dim lsWorkStation As String
    
    Dim rsBarcode As ADODB.Recordset
    Dim cmdBarcode As New ADODB.Command
    
    Dim iOrd As Integer
    
    On Error GoTo errtrap
    
    ClearSpread vasList1
    
    k = DateDiff("d", DTPicker1.Value, DTPicker2)
    If k < 0 Then
        MsgBox "날짜 선택이 잘못되었습니다"
        Exit Sub
    End If
    
    If cboWorkStation.ListIndex < 0 Or cboWorkStation.ListIndex >= cboWorkStation.ListCount Then
        MsgBox "WorkStation 을 선택하십시오"
        cboWorkStation.SetFocus
        Exit Sub
    End If
    
    lsWorkStation = Left(cboWorkStation.Text, 2)
    
    Me.MousePointer = 11
    
    lsDate = DTPicker1.Value
    
    lRow = 1
    For i = 0 To k
        For j = 1 To 5
            If Not rs Is Nothing Then Set rs = Nothing
            
            With cmdSQL
                .ActiveConnection = cn_Ser
                .CommandType = adCmdStoredProc
                .CommandText = "Interface_WL_List_SELECT_sp"
                '.Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
                .Parameters.Append .CreateParameter("@i_instrumentcode", adChar, adParamInput, 11, lsWorkStation)
                .Parameters.Append .CreateParameter("@i_WorkList_Date", adChar, adParamInput, 11, lsDate)
                .Parameters.Append .CreateParameter("@i_Order_Number", adChar, adParamInput, 11, j)
                .Parameters.Append .CreateParameter("@i_from_seq_number", adChar, adParamInput, 11, Trim(txtSeqNo1))
                .Parameters.Append .CreateParameter("@i_to_seq_number", adChar, adParamInput, 11, Trim(txtSeqNo2))
        
                Set rs = New ADODB.Recordset
                rs.CursorType = adOpenStatic
                Set rs = .Execute
            End With
        
            For n = 0 To cmdSQL.Parameters.Count - 1
                cmdSQL.Parameters.Delete 0
            Next n
            
            While Not rs.EOF
                If vasList1.MaxRows < lRow Then
                    vasList1.MaxRows = lRow
                End If
                
                iOrd = -1
                With cmdSQL
                    .ActiveConnection = cn
                    .CommandType = adCmdText
                    .CommandText = "select barcode, OrdFlag from worklist where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
                    Set rsBarcode = New ADODB.Recordset
                    Set rsBarcode = .Execute
                End With
                If Not rsBarcode.EOF Then
                    If Trim(CStr(rsBarcode.Fields.Item(1).Value)) = "B" Then
                        iOrd = 1
                    End If
                    rsBarcode.Close
                End If
                
                With cmdSQL
                    .ActiveConnection = cn
                    .CommandType = adCmdText
                    .CommandText = "select barcode from pat_res where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
                    Set rsBarcode = New ADODB.Recordset
                    Set rsBarcode = .Execute
                End With
                If Not rsBarcode.EOF Then
                    If Trim(CStr(rsBarcode.Fields.Item(0).Value)) = Trim(rs.Fields.Item(1).Value) Then
                        iOrd = 1
                    End If
                    rsBarcode.Close
                End If
                
                If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
                If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
    
                If iOrd <> 1 Then
                    vasList1.Row = lRow
                    vasList1.Col = 2
                    If IsNull(rs.Fields.Item(0).Value) Then
                        vasList1.Text = ""
                    Else
                        vasList1.Text = CStr(j) & "-" & SetChar(Trim(CStr(rs.Fields.Item(0).Value)), 3, 1, " ")
                    End If
                    
                    For lCol = 1 To rs.Fields.Count - 1
                        vasList1.Row = lRow
                        vasList1.Col = lCol + 2
                        If IsNull(rs.Fields.Item(lCol).Value) Then
                            vasList1.Text = ""
                        Else
                            vasList1.Text = Trim(CStr(rs.Fields.Item(lCol).Value))
                        End If
                    Next lCol
                    lRow = lRow + 1
                End If
                
                rs.MoveNext
            Wend
            
            
        Next j
        lsDate = Format(DateAdd("d", 1, CDate(lsDate)), "yyyy-mm-dd")
    Next i
    
    vasList1.MaxRows = vasList1.DataRowCnt
    
    Me.MousePointer = 0
    
    If vasList1.DataRowCnt > 0 Then
        If MsgBox("장비로 오더를 전송하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbYes Then
            'cmdOrder_Click
            SP_Order
        End If
    End If
    
    Exit Sub
    
errtrap:
    Me.MousePointer = 0
    
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    
    If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
    If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
    
    'MsgBox Err.Number & " : " & Err.Description

    Exit Sub
End Sub

Sub Sck_Data(ByVal asData As String)
    Dim lRow As Long
    Dim lsRsv As String
    Dim lsID As String
    Dim i, j, k As Integer
    Dim mExam As Variant
    
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    
    Dim lsIP As String
    Dim lsHost As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    On Error GoTo ErrHandle
    
    lsRsv = asData
    
    If Trim(lsRsv) <> "" Then
        lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
        
        i = InStr(1, lsRsv, Chr(10))
        Do While i > 0
            lsWkNo = ""
            lsID = Left(lsRsv, i - 1)
            lsRsv = Mid(lsRsv, i + 1)
            
            'SaveQuery lsID & " : 체크시작"
            'SaveQuery "나머지 : " & lsRsv
            
            lRow = vasList.DataRowCnt + 1
            If lRow > vasList.MaxRows Then
                vasList.MaxRows = lRow
            End If
            
            SetText vasList, lsID, lRow, 1
            SetText vasList, "A", lRow, 2
            SetText vasList, GetDateFull, lRow, 5
            SetText vasList, lsIP, lRow, 6
            
            lsSlideOrd = ""
            
            mExam = Get_OrderBody(lsID)
            If Not IsNull(mExam) Then
                SetText vasList, mExam(1, LBound(mExam, 2)), lRow, 3
                SetText vasList, mExam(2, LBound(mExam, 2)), lRow, 4
                
                lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & SetSpace(Trim(mExam(6, LBound(mExam, 2))), 3)
                'lsPID = SetSpace(mExam(1, LBound(mExam, 2)), 8)
                lsPID = SetSpace(Left(mExam(1, LBound(mExam, 2)), 8), 8)
                
                CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(Trim(GetText(vasList, lRow, 5)), 10)
                If Not IsNumeric(gPatGen.Age) Then
                    SetText vasList, gPatGen.Age, lRow, 7
                Else
                    SetText vasList, "", lRow, 7
                End If
                    

'                ClearSpread vasExam
'                For j = 0 To UBound(mExam, 2)
'                    SetText vasExam, mExam(3, j), j, 1
'                    SetText vasExam, mExam(4, j), j, 2
'                Next j
                'SaveQuery "SP OK!"
            Else
                'SaveQuery "SP에서 정보 없음"
            End If
            
            SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP ) " & vbCrLf & _
                  "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'A','','', 0, '" & Trim(GetText(vasList, lRow, 6)) & "') "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL, 1
            End If
            'CopyRecord lRow
            
            
            'Winsock1.SendData "/" & lsID & "???"
            If chkReal.Value = 1 Then
                
                If Not IsNull(mExam) Then
'                    If MSComm1.PortOpen = False Then
'                        LASCPortOpen
'                    End If
                    
                    If MSComm1.PortOpen = True Then
                        MSComm1.PortOpen = False
                    End If
                    LASCPortOpen
                    
                    
                    If MSComm1.CTSHolding = False Then
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
                        lblMsg.ForeColor = RGB(255, 0, 0)
                    Else
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
                        lblMsg.ForeColor = RGB(0, 0, 0)
        
                        For i = 1 To 7
                            Ord(i) = "0"
                        Next i
                        SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
                        Set AdoRs_Exam = db_select_rs(gLocal, SQL)
                        
                        ClearSpread vasExam
                        lblBarCode.Caption = lsID
                        
                        lsSlideOrd = ""
                        k = 1
                        For j = LBound(mExam, 2) To UBound(mExam, 2)
                            SetText vasExam, mExam(3, j), k, 1
                            SetText vasExam, mExam(4, j), k, 2
                            k = k + 1
                            If Not AdoRs_Exam Is Nothing Then
                                AdoRs_Exam.MoveFirst
                                Do Until AdoRs_Exam.EOF
                                    If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                                        Case "C": Ord(1) = "1"
                                        Case "D": Ord(2) = "1"
                                        Case "R": Ord(3) = "1"
                                        Case "P"
                                            Ord(4) = "1"
                                            lsSlideOrd = "SP"
                                        Case "S"
                                            Ord(5) = "1"
                                            lsSlideOrd = "SC"
                                        Case "X": Ord(6) = "1"
                                        Case "B"
                                            If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                                Ord(7) = "1"
                                            End If
                                        End Select
                                        
                                        Exit Do
                                    End If
                                    
                                    AdoRs_Exam.MoveNext
                                Loop
                            End If
                        Next j
                        
                        lsOrder = ""
                        For i = 1 To 7
                            lsOrder = lsOrder & Ord(i)
                        Next i
                                
                        'SaveQuery "실시간 전송 : " & lsID & " : " & lsOrder
                        'MsgBox lsOrder
                        
                        If lsOrder <> "0000000" And lsOrder <> "" Then
                            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                            'lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & lsWkNo & lsPID
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "000****************************************" & chrETX
                            
                            DoSleep 1000
                            
                            MSComm1.Output = lsOrder
                            
                            SaveOrdLog lsOrder
                            
                            SQL = "Select barcode from res_flag where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                            res = db_select_Col(gLocal, SQL)
                            If Trim(gReadBuf(0)) = lsID Then
                                SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "' " & vbCrLf & _
                                      "where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                End If
                            Else
                                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                                      "Values ('" & lsExamDate & "', '" & lsID & "', '', '', '', '', " & _
                                      "'', '', '', '', '', '', " & _
                                      "'', '', '', '', '', '" & lsSlideOrd & "' ) "
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                End If
                            End If
                            
                            SQL = "Update WorkList set OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            'SQL = "Update WorkList set OrdDateTime = '" & Trim(GetText(vasList, lRow, 5)) & "', OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                            Else
                                SetText vasList, "B", lRow, 2
                                CopyRecord lRow
                            End If
                        
                        End If
                        
                    End If
                    
                    'MSComm1.PortOpen = False
                End If
            End If
            i = InStr(1, lsRsv, Chr(10))
        Loop
    End If
    
    'Winsock1.SendData "/???"
    
    sckState = 2
    
    Winsock1.Close
    sckState = -1
    
    Winsock1.LocalPort = MyPort
    Winsock1.Listen
    
    sckState = 0
    
    'SP_Search
    
    Exit Sub
ErrHandle:
    SaveQuery "[Winsock]" & Err.Number & ": " & Err.Description
    Resume Next

End Sub

Sub SP_Search()

    Dim lsDate As String
    Dim liOrdNo As Integer
    Dim i, j, k, n
    Dim lRow, lCol As Long
    
    Dim lsWorkStation As String
    
    Dim rsBarcode As ADODB.Recordset
    Dim cmdBarcode As New ADODB.Command
    
    Dim iOrd As Integer
    
    On Error GoTo errtrap
    
    ClearSpread vasList1
    
    k = DateDiff("d", DTPicker1.Value, DTPicker2)
    If k < 0 Then
        MsgBox "날짜 선택이 잘못되었습니다"
        Exit Sub
    End If
    
    If cboWorkStation.ListIndex < 0 Or cboWorkStation.ListIndex >= cboWorkStation.ListCount Then
        MsgBox "WorkStation 을 선택하십시오"
        cboWorkStation.SetFocus
        Exit Sub
    End If
    
    lsWorkStation = Left(cboWorkStation.Text, 2)
    
    Me.MousePointer = 11
    
    lsDate = DTPicker1.Value
    
    lRow = 1
    For i = 0 To k
        For j = 1 To 5
            With cmdSQL
                .ActiveConnection = cn_Ser
                .CommandType = adCmdStoredProc
                .CommandText = "Interface_WL_List_SELECT_sp"
                '.Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
                .Parameters.Append .CreateParameter("@i_instrumentcode", adChar, adParamInput, 11, lsWorkStation)
                .Parameters.Append .CreateParameter("@i_WorkList_Date", adChar, adParamInput, 11, lsDate)
                .Parameters.Append .CreateParameter("@i_Order_Number", adChar, adParamInput, 11, j)
                .Parameters.Append .CreateParameter("@i_from_seq_number", adChar, adParamInput, 11, Trim(txtSeqNo1))
                .Parameters.Append .CreateParameter("@i_to_seq_number", adChar, adParamInput, 11, Trim(txtSeqNo2))
        
                Set rs = New ADODB.Recordset
                rs.CursorType = adOpenStatic
                Set rs = .Execute
            End With
        
            For n = 0 To cmdSQL.Parameters.Count - 1
                cmdSQL.Parameters.Delete 0
            Next n
            
            'lRow = 1
            While Not rs.EOF
                If vasList1.MaxRows < lRow Then
                    vasList1.MaxRows = lRow
                End If
                
                iOrd = -1
                With cmdSQL
                    .ActiveConnection = cn
                    .CommandType = adCmdText
                    .CommandText = "select barcode, OrdFlag from worklist where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
                    Set rsBarcode = New ADODB.Recordset
                    Set rsBarcode = .Execute
                End With
                If Not rsBarcode.EOF Then
                    If Trim(CStr(rsBarcode.Fields.Item(1).Value)) = "B" Then
                        iOrd = 1
                    End If
                    rsBarcode.Close
                End If
                
                With cmdSQL
                    .ActiveConnection = cn
                    .CommandType = adCmdText
                    .CommandText = "select barcode from pat_res where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
                    Set rsBarcode = New ADODB.Recordset
                    Set rsBarcode = .Execute
                End With
                If Not rsBarcode.EOF Then
                    If Trim(CStr(rsBarcode.Fields.Item(0).Value)) = Trim(rs.Fields.Item(1).Value) Then
                        iOrd = 1
                    End If
                    rsBarcode.Close
                End If
                
                If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
                If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
    
                If iOrd <> 1 Then
                    vasList1.Row = lRow
                    vasList1.Col = 2
                    If IsNull(rs.Fields.Item(0).Value) Then
                        vasList1.Text = ""
                    Else
                        vasList1.Text = CStr(j) & "-" & SetChar(Trim(CStr(rs.Fields.Item(0).Value)), 3, 1, " ")
                    End If
                    
                    For lCol = 1 To rs.Fields.Count - 1
                        vasList1.Row = lRow
                        vasList1.Col = lCol + 2
                        If IsNull(rs.Fields.Item(lCol).Value) Then
                            vasList1.Text = ""
                        Else
                            vasList1.Text = Trim(CStr(rs.Fields.Item(lCol).Value))
                        End If
                    Next lCol
                    lRow = lRow + 1
                End If
                
                rs.MoveNext
            Wend
    
        Next j
        lsDate = Format(DateAdd("d", 1, CDate(lsDate)), "yyyy-mm-dd")
    Next i
    
    vasList1.MaxRows = vasList1.DataRowCnt
    
    Me.MousePointer = 0
    
    If vasList1.DataRowCnt > 0 Then
        'cmdOrder_Click
        SP_Order
    End If
    
    Exit Sub
    
errtrap:
    Me.MousePointer = 0
    
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    
    If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
    If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
        
    Exit Sub
    'MsgBox Err.Number & " : " & Err.Description

End Sub

Sub SP_Order()
Dim lsID As String
Dim i, j As Integer
Dim mExam As Variant
Dim AdoRs_Exam As ADODB.Recordset
Dim Ord(7) As String
Dim lsOrder As String
Dim lRow1, lRow As Long
Dim lsWkNo As String
Dim lsPID As String

Dim lsSlideOrd As String
Dim lsExamDate As String


If MSComm1.PortOpen = False Then
    LASCPortOpen
End If
    
lRow1 = 1
Do While lRow1 <= vasList1.DataRowCnt
    vasList1.Row = lRow1
    vasList1.Col = 1
    
    
    lsSlideOrd = ""
    lsWkNo = ""
    
    For i = 1 To 7
        Ord(i) = "0"
    Next i
    
    lRow = vasList.DataRowCnt + 1
    If lRow > vasList.MaxRows Then
        vasList.MaxRows = lRow
    End If
    
    lsID = Trim(GetText(vasList1, lRow1, 3))
    
    SetText vasList, lsID, lRow, 1
    SetText vasList, "A", lRow, 2
    SetText vasList, Trim(GetText(vasList1, lRow1, 4)), lRow, 3
    SetText vasList, Trim(GetText(vasList1, lRow1, 5)), lRow, 4
    SetText vasList, GetDateFull, lRow, 5
            
    mExam = Get_OrderBody(lsID)
    If Not IsNull(mExam) Then
        SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
        Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
        ClearSpread vasExam
        For j = LBound(mExam, 2) To UBound(mExam, 2)
            SetText vasExam, mExam(3, j), j + 1, 1
            SetText vasExam, mExam(4, j), j + 1, 2
            
            lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & SetSpace(Trim(mExam(6, LBound(mExam, 2))), 3)
            'lsPID = SetSpace(mExam(1, LBound(mExam, 2)), 8)
            lsPID = SetSpace(Left(mExam(1, LBound(mExam, 2)), 8), 8)
            
            CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(GetDateFull, 10)
            If Not IsNumeric(gPatGen.Age) Then
                SetText vasList, gPatGen.Age, lRow, 7
            Else
                SetText vasList, "", lRow, 7
            End If
            
            If Not AdoRs_Exam Is Nothing Then
                AdoRs_Exam.MoveFirst
                Do Until AdoRs_Exam.EOF
                    If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                        Case "C": Ord(1) = "1"
                        Case "D": Ord(2) = "1"
                        Case "R": Ord(3) = "1"
                        Case "P"
                            Ord(4) = "1"
                            lsSlideOrd = "SP"
                        Case "S"
                            Ord(5) = "1"
                            lsSlideOrd = "SC"
'                            Case "P": Ord(4) = "1"
'                            Case "S": Ord(5) = "1"
                        Case "X": Ord(6) = "1"
                        Case "B"
                            If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                Ord(7) = "1"
                            End If
                        End Select
                        
                        Exit Do
                    End If
                    
                    AdoRs_Exam.MoveNext
                Loop
            End If
        Next j
        
        lsOrder = ""
        For i = 1 To 7
            lsOrder = lsOrder & Ord(i)
        Next i
        
        'MsgBox lsOrder
        
        If lsOrder <> "0000000" And lsOrder <> "" Then
            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
            'lsOrder = lsOrder & "0000000000000"
            lsOrder = lsOrder & lsWkNo & lsPID
            lsOrder = lsOrder & "0000000000000"
            lsOrder = lsOrder & "0000000000000"
            lsOrder = lsOrder & "000****************************************" & chrETX
            
            DoSleep 100
            
            MSComm1.Output = lsOrder
            
            SaveOrdLog lsOrder
            
            SQL = "Select barcode from res_flag where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = lsID Then
                SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "' " & vbCrLf & _
                      "where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                res = SendQuery(gLocal, SQL)
            Else
                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                      "Values ('" & lsExamDate & "', '" & lsID & "', '', '', '', '', " & _
                      "'', '', '', '', '', '', " & _
                      "'', '', '', '', '', '" & lsSlideOrd & "' ) "
                res = SendQuery(gLocal, SQL)
            End If
            
            
            SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
                  "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','','', 0) "
            res = SendQuery(gLocal, SQL)
            If res = 1 Then
                SetText vasList, "B", lRow, 2
                CopyRecord lRow
                'DeleteRow vasList1, lRow1, lRow1
            Else
                SaveQuery SQL
            End If
        End If
    End If
    lRow1 = lRow1 + 1
Loop

ClearSpread vasList1

End Sub

Private Sub cmdWorkList_Click()
Dim lsID As String
Dim i, j, k As Integer
Dim mExam As Variant
Dim AdoRs_Exam As ADODB.Recordset
Dim Ord(7) As String
Dim lsOrder As String
Dim lRow As Long
Dim lsDate As String

    
If MSComm1.PortOpen = False Then
    LASCPortOpen
End If
    
If MSComm1.CTSHolding = False Then
    lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
    lblMsg.ForeColor = RGB(255, 0, 0)
    
    Exit Sub
Else
    lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
    lblMsg.ForeColor = RGB(0, 0, 0)
End If
    
For lRow = 1 To vasList.DataRowCnt
    lsDate = GetDateFull
    
    If Trim(GetText(vasList, lRow, 2)) = "A" Then
    
        For i = 1 To 7
            Ord(i) = "0"
        Next i
        
        lsID = Trim(GetText(vasList, lRow, 1))
        
        'If lsID = "" Then Exit Sub
        
        mExam = Get_OrderBody(lsID)
        If Not IsNull(mExam) Then
            SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
            Set AdoRs_Exam = db_select_rs(gLocal, SQL)
            
            ClearSpread vasExam
            lblBarCode.Caption = lsID
            
            k = 1
            For j = LBound(mExam, 1) To UBound(mExam, 2)
                SetText vasExam, mExam(3, j), k, 1
                SetText vasExam, mExam(4, j), k, 2
                k = k + 1
                If Not AdoRs_Exam Is Nothing Then
                    AdoRs_Exam.MoveFirst
                    Do Until AdoRs_Exam.EOF
                        If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                            Select Case Trim(AdoRs_Exam("OrdGubun"))
                            Case "C": Ord(1) = "1"
                            Case "D": Ord(2) = "1"
                            Case "R": Ord(3) = "1"
                            Case "P": Ord(4) = "1"
                            Case "S": Ord(5) = "1"
                            Case "X": Ord(6) = "1"
                            Case "B": Ord(7) = "1"
                            End Select
                            
                            Exit Do
                        End If
                        
                        AdoRs_Exam.MoveNext
                    Loop
                End If
            Next j
            
            lsOrder = ""
            For i = 1 To 7
                lsOrder = lsOrder & Ord(i)
            Next i
            
            'MsgBox lsOrder
            
            If lsOrder <> "0000000" And lsOrder <> "" Then
                lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "000****************************************" & chrETX
                MSComm1.Output = lsOrder
                SaveOrdLog lsOrder
                
                'SQL = "Update WorkList set OrdDateTime = '" & lsDate & "', OrdFlag = 'B' where Barcode ='" & lsID & "'"
                SQL = "Update WorkList set OrdFlag = 'B' where Barcode ='" & lsID & "'"
                res = SendQuery(gLocal, SQL)
                If res = 1 Then
                    SetText vasList, "B", lRow, 2
                    CopyRecord lRow
                Else
                    SaveQuery SQL
                    'Exit Function
                End If
            End If
            
        End If
    
    End If
Next lRow


End Sub

Sub CopyRecord(ByVal asRow As Long)
    Dim llRow As Long
    Dim llCol As Long
    
    If asRow < 1 Or asRow > vasList.DataRowCnt Then Exit Sub
    
    llRow = vasOrder.DataRowCnt + 1
    If llRow > vasOrder.MaxRows Then
        vasOrder.MaxRows = llRow
    End If
    
    For llCol = 1 To 6
        SetText vasOrder, Trim(GetText(vasList, asRow, llCol)), llRow, llCol
    Next llCol
    vasList.DeleteRows asRow, 1
End Sub

Private Sub Command1_Click()
    CheckStr Text1
End Sub


Private Sub Command2_Click()
    typOrder.SampleID = Trim(GetText(vasList, 1, 1))
    SendTheDataCLINILOG typOrder.SampleID, False, Left(Trim(GetText(vasList, 1, 5)), 8)
    DeleteRow vasList, 1, 1
End Sub

Private Sub Command3_Click()
    ClearSpread vasList
    
    SQL = " Select barcodenumber, '','','',worklist_Date + ' ' + worklist_time " & vbCrLf
    SQL = SQL & " From tlaorder  " & vbCrLf
    SQL = SQL & " order by 5 desc  "
    res = db_select_Vas(gServer, SQL, vasList)
    
    'vasSort vasList, 5, 0
End Sub

Private Sub Command4_Click()
    Dim lRow
    Dim mExam As Variant
    
                If iState = 1 Then
                    typOrder.SampleID = Trim(GetText(vasList, vasList.DataRowCnt, 1))
                    
                    If SendTheDataCLINILOG(typOrder.SampleID, False, Left(Trim(GetText(vasList, vasList.DataRowCnt, 5)), 8)) Then
                        iState = 2
                    Else
                        SQL = "delete from tlaorder where barcodenumber = '" & Trim(typOrder.SampleID) & "' "
                        'res = SendQuery(gServer, SQL)
                        
                        'DeleteRow vasList, vasList.DataRowCnt, vasList.DataRowCnt
                        
                        iState = 0
                        
                        MSComm1.Output = chrEOT
                        SaveOrdLog "TX : EOT"
                        typOrder.SampleID = ""
                        
                        lRow = vasList.DataRowCnt
                        Do While lRow <= vasList.DataRowCnt And lRow > 0
                            mExam = Get_OrderBody(Trim(Trim(GetText(vasList, lRow, 1))))
                            If Not IsNull(mExam) Then
                                Exit Do
                            Else
                                SQL = "delete from tlaorder where barcodenumber = '" & Trim(GetText(vasList, lRow, 1)) & "' "
                                'res = SendQuery(gServer, SQL)
                                
                                'DeleteRow vasList, lRow, lRow
                            End If
                            lRow = vasList.DataRowCnt
                        Loop
                        
                        
                        If vasList.DataRowCnt > 0 Then
                        
                            iState = 1
                            MSComm1.Output = chrENQ
                            SaveOrdLog "TX : ENQ"
                            
                            bTimer = False
                            
                            Exit Sub
                        Else
                            bTimer = True
                        End If
                    End If
                ElseIf iState = 2 Then
                    
                    SQL = "delete from tlaorder where barcodenumber = '" & Trim(typOrder.SampleID) & "' "
'                    res = SendQuery(gServer, SQL)
                    
'                    DeleteRow vasList, vasList.DataRowCnt, vasList.DataRowCnt
                    
                    iState = 0
                    
                    MSComm1.Output = chrEOT
                    SaveOrdLog "TX : EOT"
                    typOrder.SampleID = ""
                    
                    lRow = vasList.DataRowCnt
                    Do While lRow <= vasList.DataRowCnt And lRow > 0
                        mExam = Get_OrderBody(Trim(Trim(GetText(vasList, lRow, 1))))
                        If Not IsNull(mExam) Then
                            Exit Do
                        Else
                            SQL = "delete from tlaorder where barcodenumber = '" & Trim(GetText(vasList, lRow, 1)) & "' "
'                            res = SendQuery(gServer, SQL)
                            
'                            DeleteRow vasList, lRow, lRow
                        End If
                        lRow = vasList.DataRowCnt
                    Loop
                        
                    If vasList.DataRowCnt > 0 Then
                    
                        iState = 1
                        MSComm1.Output = chrENQ
                        SaveOrdLog "TX : ENQ"
                        
                        bTimer = False
                        
                        Exit Sub
                    Else
                        bTimer = True
                    End If
                                        
                End If

End Sub

Private Sub Command5_Click()
    Dim lsCode As String
    Dim i, lRow, j
    Dim mExam As Variant
    

mExam = Get_OrderBody_Cancel("A0811037971", "2008-11-03")
If Not IsNull(mExam) Then
    For i = LBound(mExam, 2) To UBound(mExam, 2)
        For j = LBound(mExam, 1) To UBound(mExam, 1)
            SetText vasExam, mExam(j, i), lRow + 1, j + 1
        Next j
        lRow = lRow + 1
    Next i
End If
SendTheDataCLINILOG_DEL "A0811037971", True, "20081103"
'    lsCode = "0599"
'If Order_OK(lsCode) = 1 Then
'    MsgBox "ok"
'
'End If

End Sub

Private Sub Command6_Click()
    Dim mExam As Variant
    Dim lsID As String
    Dim i, j
    
    lsID = "E0908068981"
    
    mExam = Get_OrderBody_Cancel(Trim(lsID), "2009-08-06")
    If Not IsNull(mExam) Then
        i = LBound(mExam, 2)
        
        For j = LBound(mExam, 1) To UBound(mExam, 1)
            MsgBox CStr(j) & vbTab & CStr(Trim(mExam(j, i)))
        Next j
    End If
    
'    SendTheDataCLINILOG "C0901165101", False, "20090116"
End Sub

Private Sub Form_Load()
    Dim db_tmp As String * 100
    Dim lsData As String
    Dim i As Integer
    
'    db_tmp = ""
'    Call GetPrivateProfileString("Option", "WorkStationCode", "", db_tmp, 100, App.Path & "\Interface.ini")
'    txtTemp = Trim(db_tmp)
'    gInsCode = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Option", "LocalPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    MyPort = Trim(txtTemp)
    
    sckState = -1
    
    cn_Local_Flag = False
    cn_Server_Flag = False
    
    GetSetup
    
    If Connect_Local Then
        cn_Local_Flag = True
    End If
    
'    If Connect_Server Then
'        cn_Server_Flag = True
'    End If
    
    
    gRow = 1
    
    If Not IsNumeric(gExpireDate) Then
        gExpireDate = 15
    End If
    
    'gExpireDate = Format(DateAdd("d", 0 - CInt(gExpireDate), CDate(GetDateFull)), "yyyy-mm-dd") & " 00:00:00"
    SQL = "Delete from WorkList "
    res = SendQuery(gLocal, SQL)
    
    SQL = "Select RemoteIP from WorkList "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table worklist add column RemoteIP varchar(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select SlideOrd from res_flag "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table res_flag add column SlideOrd varchar(2) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select OrdCode from EquipExam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table EquipExam add column OrdCode varchar(20) "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Update EquipExam set OrdCode = ExamCode "
        res = SendQuery(gLocal, SQL)
    End If
    
    'EquipExam
    SQL = "Select IndexFlag from EquipExam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table EquipExam add column IndexFlag varchar(5) "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Update EquipExam set IndexFlag = '' "
        res = SendQuery(gLocal, SQL)
    End If
    
    DTPicker1.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
    DTPicker2.Value = DTPicker1.Value
    
    dtpSDate.Value = CDate(Date)
    dtpEDate.Value = CDate(Date)
    
'    SQL = SQL & " delete from tlaorder  " & vbCrLf
'    SQL = SQL & " where worklist_date < '" & Format(CDate(DTPicker1.Value), "yyyymmdd") & "' "
'    res = SendQuery(gServer, SQL)
    
    cmdPortOpen_Click
    
    cmdReLoad_Click
    
    bTimer = True
    
'2010.02.17 이상은
'    SQL = " delete from equipexam where equipcode not like 'L%' "
'    res = SendQuery(gLocal, SQL)
End Sub

Sub LASCPortOpen()
    GetSetup_LASC
    
    MSComm1.CommPort = gSetup.Port
'    MSComm1.CommPort = 8
    MSComm1.Settings = gSetup.Speed & "," & gSetup.Parity & "," & gSetup.DataBit & "," & gSetup.StopBit
    If gSetup.DTREnable = "1" Then
        MSComm1.DTREnable = True
    Else
        MSComm1.DTREnable = False
    End If
    If gSetup.RTSEnable = "1" Then
        MSComm1.RTSEnable = True
    Else
        MSComm1.RTSEnable = False
    End If
    
    MSComm1.PortOpen = True
    
'    If MSComm1.CTSHolding = False Then
'        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
'        lblMsg.ForeColor = RGB(255, 0, 0)
'    Else
'        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
'        lblMsg.ForeColor = RGB(0, 0, 0)
'    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Timer1.Enabled = False
    
    Winsock1.Close
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    DisConnect_Server
    
    DisConnect_Local
    
    End
End Sub

Private Sub MSComm1_OnComm()
    Dim lsIn        As String
    Dim ii, lRow    As Integer
    Dim mExam       As Variant
    Dim lbDelFlag   As String
    Dim lsExamDate  As String
    
On Error Resume Next

    lsIn = MSComm1.Input
    
    Select Case lsIn
        Case chrSTX:
                'ii = ii + 1 'Frame Number
                cntCheckSum = 0
        Case chrETX:
                cntCheckSum = cntCheckSum + 1
                MSComm1.Output = chrACK
                SaveOrdLog "TX : ACK"
                'flgETX = True
                ReceiveTheDataCLINILOG
        Case chrETB:
                If Mid(ReceiveData, ii - 1, 2) = chrCR & chrLF Then
                    ReceiveData = Left(ReceiveData, Len(ReceiveData) - 2) 'Remove CR & LF
                End If
                cntCheckSum = cntCheckSum + 1
                MSComm1.Output = chrACK
                SaveOrdLog "TX : ACK"
                flgETB = True
        Case chrCR:
                If flgETB = True Then
                   flgETB = False
                Else
                    'ReceiveTheDataCLINILOG
                    'GoSub ClearReceiveData
                End If
        Case chrLF:
                '
        Case chrENQ:
                MSComm1.Output = chrACK
                SaveOrdLog "TX : ACK"
                'TXTLOG = txtLog & "TX : ACK" & vbCrLf
                'SendFlg = False
        Case chrACK:
                SaveOrdLog "RX : ACK"
                'TXTLOG = txtLog & "RX : ACK" & vbCrLf
                If iState = 1 Then
                    typOrder.SampleID = Trim(GetText(vasList, 1, 1))
                    
                    If Trim(GetText(vasList, 1, 6)) <> "" Then
                        If UCase(Trim(GetText(vasList, 1, 6))) = "RECHECK" Then
                            lbDelFlag = "RE"
                        Else
                            lbDelFlag = "True"
                        End If
                    Else
                        lbDelFlag = "False"
                    End If
                    
                    'If Trim(typOrder.SampleID) = "12060102900" Then Stop
                    If SendTheDataCLINILOG(typOrder.SampleID, lbDelFlag, Left(Trim(GetText(vasList, 1, 5)), 8)) Then
                        iState = 2
                    Else
'                        SQL = "delete from tlaorder where barcodenumber = '" & Trim(typOrder.SampleID) & "' "
'                        res = SendQuery(gServer, SQL)
                        
                        DeleteRow vasList, 1, 1
                        
                        iState = 0
                        
                        MSComm1.Output = chrEOT
                        SaveOrdLog "TX : EOT"
                        'TXTLOG = txtLog & "TX : EOT" & vbCrLf
                        
                        typOrder.SampleID = ""
                        
'                        lRow = 1
'                        Do While lRow <= vasList.DataRowCnt And lRow > 0
'
''                            If Trim(Trim(GetText(vasList, lRow, 6))) = "" Then
''                                mExam = Get_OrderBody(Trim(Trim(GetText(vasList, lRow, 1))))
''                            Else
''                                lsExamDate = Trim(GetText(vasList, lRow, 5))
''                                lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2)
''                                mExam = Get_OrderBody_Cancel(Trim(Trim(GetText(vasList, lRow, 1))), lsExamDate)
''                            End If
''
''                            If Not IsNull(mExam) Then
''                                Exit Do
''                            Else
''                                SQL = "delete from tlaorder where barcodenumber = '" & Trim(GetText(vasList, lRow, 1)) & "' "
''                                res = SendQuery(gServer, SQL)
''
''                                DeleteRow vasList, lRow, lRow
''                            End If
'
'                            If Trim(Trim(GetText(vasList, lRow, 6))) = "" Then
'                                'mExam = Get_OrderBody(Trim(Trim(GetText(vasList, lRow, 1))))
'                                res = Online_XML(gXml_S07, Trim(GetText(vasList, lRow, 1)))
'                            Else
'                                lsExamDate = Trim(GetText(vasList, lRow, 5))
'                                lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2)
'                                'mExam = Get_OrderBody_Cancel(Trim(Trim(GetText(vasList, lRow, 1))), lsExamDate)
'                            End If
'
'                            If res = 1 Then
'                                DeleteRow vasList, lRow, lRow
'                            Else
'                                Exit Do
'                            End If
'
'                            lRow = 1
'                        Loop
                        
                        
                        If vasList.DataRowCnt > 0 Then
                        
                            iState = 1
                            MSComm1.Output = chrENQ
                            SaveOrdLog "TX : ENQ"
                            'TXTLOG = txtLog & "TX : ENQ" & vbCrLf
                            
                            bTimer = False
                            
                            Exit Sub
                        Else
                            bTimer = True
                            
                            Timer1.Enabled = True
                        End If
                    End If
                ElseIf iState = 2 Then
                    
'                    SQL = "delete from tlaorder where barcodenumber = '" & Trim(typOrder.SampleID) & "' "
'                    res = SendQuery(gServer, SQL)
                    
                    DeleteRow vasList, 1, 1
                    
                    iState = 0
                    
                    MSComm1.Output = chrEOT
                    SaveOrdLog "TX : EOT"
                    'TXTLOG = txtLog & "TX : EOT" & vbCrLf
                    
                    typOrder.SampleID = ""
                    
'                    lRow = 1
'                    Do While lRow <= vasList.DataRowCnt And lRow > 0
''                        If Trim(Trim(GetText(vasList, lRow, 6))) = "" Then
''                            mExam = Get_OrderBody(Trim(Trim(GetText(vasList, lRow, 1))))
''                        Else
''                            lsExamDate = Trim(GetText(vasList, lRow, 5))
''                            lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2)
''                            mExam = Get_OrderBody_Cancel(Trim(Trim(GetText(vasList, lRow, 1))), lsExamDate)
''                        End If
''
''                        If Not IsNull(mExam) Then
''                            Exit Do
''                        Else
''                            SQL = "delete from tlaorder where barcodenumber = '" & Trim(GetText(vasList, lRow, 1)) & "' "
''                            res = SendQuery(gServer, SQL)
''
''                            DeleteRow vasList, lRow, lRow
''                        End If
'
'                        If Trim(Trim(GetText(vasList, lRow, 6))) = "" Then
'                            res = Online_XML(gXml_S07, Trim(GetText(vasList, lRow, 1)))
'                        Else
'                            lsExamDate = Trim(GetText(vasList, lRow, 5))
'                            lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2)
'                            'mExam = Get_OrderBody_Cancel(Trim(Trim(GetText(vasList, lRow, 1))), lsExamDate)
'                        End If
'
'                        If res = 0 Then
'                            Exit Do
'                        Else
'                            DeleteRow vasList, lRow, lRow
'                        End If
'
'                        lRow = vasList.DataRowCnt
'                    Loop
                        
                    If vasList.DataRowCnt > 0 Then
                    
                        iState = 1
                        MSComm1.Output = chrENQ
                        SaveOrdLog "TX : ENQ"
                        'TXTLOG = txtLog & "TX : ENQ" & vbCrLf
                        
                        bTimer = False
                        
                        Exit Sub
                    Else
                        bTimer = True
                        Timer1.Enabled = True
                    End If
                                        
                End If
'        Case chrNACK:
'                If typOrder.SampleID <> "" Then
'                    SendCount = SendCount - 1
'                    'Temp_CLINILOG "A"
'                Else
'                    MSComm1.Output = chrEOT
'                    SaveOrdLog "TX : EOT"
'                End If
        Case chrEOT:
                'mscomm1.Output = ENQ
                cntCheckSum = 0
                'GoSub ClearReceiveData
        Case Else:
            'Select Case chrcntCheckSum
            '    Case chr1:
            '        cntCheckSum = cntCheckSum + 1
            '    Case chr2:
            '        cntCheckSum = 0
            '    Case chrElse:
                    ReceiveData = ReceiveData & lsIn
            'End Select
    End Select

End Sub

Private Sub subClose_Click()
    If MsgBox("프로그램을 종료하면 검사 처방 받기 및 장비로 처방 전송이 되지 않습니다" & vbCrLf & vbCrLf & "프로그램을 종료하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "종료알림") = vbNo Then
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub subCode_Click()
    frmCode.Show
End Sub

Private Sub subConfig_Click()
    frmConfig.Show
End Sub

Private Sub Timer1_Timer()

    Dim lsID        As String
    Dim i, j, k     As Integer
    Dim mExam       As Variant
    Dim AdoRs_Exam  As ADODB.Recordset
    Dim Ord(7)      As String
    Dim lsOrder     As String
    Dim lRow        As Long
    Dim lRow1       As Long
    
    Dim lEndRow     As Long
    Dim lsDate      As String
    
    Dim lsID1       As String
    Dim lsReceNo    As String
    Dim lsWkNo      As String
    Dim lsPID       As String
    
    Dim lsSlideOrd  As String
    Dim lsExamDate  As String
    Dim Res_1       As String
    
    On Error Resume Next
    
    If Not bTimer Then Exit Sub
    
    If MSComm1.PortOpen = False Then
        LASCPortOpen
    End If
        
'    If MSComm1.CTSHolding = False Then
'        lblMsg.Caption = "[Message] 포트가 준비되지 않았습니다"
'        lblMsg.ForeColor = RGB(255, 0, 0)
'
'        Exit Sub
'    Else
'        lblMsg.Caption = "[Message] 포트가 준비되었습니다"
'        lblMsg.ForeColor = RGB(0, 0, 0)
'    End If
        
    ClearSpread vasList
    
'    SQL = " Select barcodenumber, '','','',worklist_Date + ' ' + worklist_time, delete_Date " & vbCrLf
'    SQL = SQL & " From tlaorder  " & vbCrLf
'    SQL = SQL & " order by 5 asc  "
'    res = db_select_Vas(gServer, SQL, vasList)
    
    'vasSort vasList, 5, 0
        
'    lRow = 1
'    Do While lRow <= vasList.DataRowCnt And lRow > 0
'        If Trim(Trim(GetText(vasList, lRow, 6))) = "" Or Trim(Trim(GetText(vasList, lRow, 6))) = "RECHECK" Then
'            mExam = Get_OrderBody(Trim(Trim(GetText(vasList, lRow, 1))))
'        Else
'            lsExamDate = Trim(GetText(vasList, lRow, 5))
'            lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2)
'
'            mExam = Get_OrderBody_Cancel(Trim(Trim(GetText(vasList, lRow, 1))), lsExamDate)
'        End If
'        If Not IsNull(mExam) Then
'            Exit Do
'        Else
'            SQL = "delete from tlaorder where barcodenumber = '" & Trim(GetText(vasList, lRow, 1)) & "' "
'            res = SendQuery(gServer, SQL)
'
'            DeleteRow vasList, lRow, lRow
'        End If
'        lRow = 1
'        'lRow = lRow + 1
'    Loop

    'TLA WorkList 불러오기
    Res_1 = Online_TLA(gXml_S04, dtpSDate.Value, dtpEDate.Value)
    
    lsID1 = ""
    
    With vasList
        For lRow1 = 0 To UBound(gTLA_Info_Select)
            lsID = Trim(gTLA_Info_Select(lRow1).SPCNO)
            
            '접수번호 가져오기
            res = Online_XML(gXml_S03, lsID)
            
            lsWkNo = Trim(gPat_Info_Select.ACPTNO_1)
            
            If lsID <> lsID1 Then
                SQL = "select barcode, OrdFlag from worklist where barcode = '" & lsID & "' "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) <> "" Then
                    'txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                    'DeleteRow vasList, lRow, lRow
                Else
                    SQL = "select barcode from pat_res where barcode = '" & lsID & "' "
                    res = db_select_Col(gLocal, SQL)
                    If Trim(gReadBuf(0)) = lsID Then
                        'txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                        'DeleteRow vasList, lRow, lRow
                    Else
                        lRow = .DataRowCnt + 1
                        
                        .SetText 1, lRow, lsID
'                        .SetText 8, lRow, Trim(gPat_Info_Select.ACPTNO_1)
'
'                        .SetText 3, lRow, gPat_Info_Select.PT_NO
'                        .SetText 4, lRow, gPat_Info_Select.PT_NM

                        .SetText 5, lRow, Format(gPat_Info_Select.ACPT_DTETM, "YYYYMMDD HH:MM:SS")       '날짜
                        
'                        If IsDate(gPat_Info_Select.ACPT_DTETM) Then
'                            .SetText 5, lRow, Format(CDate(gPat_Info_Select.ACPT_DTETM), "yyyy-mm-dd")    '날짜
'                        Else
'                            .SetText 5, lRow, gPat_Info_Select.ACPT_DTETM    '날짜
'                        End If
'                        .SetText 6, lRow, ""    '시간
'                        .SetText 7, lRow, ""    '접수코드
'                        .SetText 8, lRow, gPat_Info_Select.ACPTNO_1
'                        .SetText 9, lRow, gPat_Info_Select.Sex
'                        .SetText 10, lRow, gPat_Info_Select.Age
'                        .SetText 11, lRow, ""   'slip
                        
                    End If
                End If
            End If
            lsID1 = lsID
        Next lRow1
    End With
    
    If vasList.DataRowCnt > 0 Then
        iState = 1
        MSComm1.Output = chrENQ
        SaveOrdLog "TX : ENQ"
        
        bTimer = False
        
        Timer1.Enabled = False
        
        Exit Sub
    End If
End Sub

Private Sub txtBarcode_GotFocus()
    SelectFocus txtBarcode
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtBarcode) <> "" Then
            InsertRow vasList, 1
            vasList.SetText 1, 1, txtBarcode
            vasList.SetText 5, 1, Format(Date, "yyyymmdd") & " " & Format(Time, "hhnn")
            txtBarcode = ""
            
            iState = 1
            MSComm1.Output = chrENQ
            SaveOrdLog "TX : ENQ"
            
            bTimer = False
            
            Timer1.Enabled = False
            
        End If
    End If
End Sub

Private Sub txtReOrd_GotFocus()
    SelectFocus txtReOrd
End Sub

Private Sub txtReOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    Dim lsRsv As String
    Dim lsID As String
    Dim i, j, k As Integer
    Dim mExam As Variant
    
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    
    Dim lsIP As String
    Dim lsHost As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    On Error GoTo ErrHandle
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtReOrd) = "" Then Exit Sub
            
        lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
    
        lsWkNo = ""
        lsID = Trim(txtReOrd)
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then
            vasList.MaxRows = lRow
        End If
        
        SetText vasList, lsID, lRow, 1
        SetText vasList, "A", lRow, 2
        SetText vasList, GetDateFull, lRow, 5
        SetText vasList, "", lRow, 6
        
        lsSlideOrd = ""
        
        mExam = Get_OrderBody(lsID)
        If IsNull(mExam) Then
            MsgBox "샘플정보가 없습니다"
            txtReOrd = ""
            Exit Sub
        End If
        
        SetText vasList, mExam(1, LBound(mExam, 2)), lRow, 3
        SetText vasList, mExam(2, LBound(mExam, 2)), lRow, 4
        
        lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & SetSpace(Trim(mExam(6, LBound(mExam, 2))), 3)
        'lsPID = SetSpace(mExam(1, LBound(mExam, 2)), 8)
        lsPID = SetSpace(Left(mExam(1, LBound(mExam, 2)), 8), 8)
        
        CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(Trim(GetText(vasList, lRow, 5)), 10)
        If Not IsNumeric(gPatGen.Age) Then
            SetText vasList, gPatGen.Age, lRow, 7
        Else
            SetText vasList, "", lRow, 7
        End If
        
        'SaveQuery "SP OK!"
        
        SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'A','','', 0, '" & Trim(GetText(vasList, lRow, 6)) & "') "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL, 1
        End If
                    
        If MSComm1.PortOpen = False Then
            LASCPortOpen
        End If
        
        If MSComm1.CTSHolding = False Then
            lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
            lblMsg.ForeColor = RGB(255, 0, 0)
            Exit Sub
        End If
            
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
        lblMsg.ForeColor = RGB(0, 0, 0)

        For i = 1 To 7
            Ord(i) = "0"
        Next i
        SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
        Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
        ClearSpread vasExam
        lblBarCode.Caption = lsID
        
        lsSlideOrd = ""
        k = 1
        For j = LBound(mExam, 2) To UBound(mExam, 2)
            SetText vasExam, mExam(3, j), k, 1
            SetText vasExam, mExam(4, j), k, 2
            k = k + 1
            If Not AdoRs_Exam Is Nothing Then
                AdoRs_Exam.MoveFirst
                Do Until AdoRs_Exam.EOF
                    If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                        Case "C": Ord(1) = "1"
                        Case "D": Ord(2) = "1"
                        Case "R": Ord(3) = "1"
                        Case "P"
                            Ord(4) = "1"
                            lsSlideOrd = "SP"
                        Case "S"
                            Ord(5) = "1"
                            lsSlideOrd = "SC"
                        Case "X": Ord(6) = "1"
                        Case "B"
                            If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                Ord(7) = "1"
                            End If
                        End Select
                        
                        Exit Do
                    End If
                    
                    AdoRs_Exam.MoveNext
                Loop
            End If
        Next j
        
        lsOrder = ""
        For i = 1 To 7
            lsOrder = lsOrder & Ord(i)
        Next i
                
        'SaveQuery "실시간 전송 : " & lsID & " : " & lsOrder
        'MsgBox lsOrder
        
        If lsOrder <> "0000000" And lsOrder <> "" Then
            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
            'lsOrder = lsOrder & "0000000000000"
            lsOrder = lsOrder & lsWkNo & lsPID
            lsOrder = lsOrder & "0000000000000"
            lsOrder = lsOrder & "0000000000000"
            lsOrder = lsOrder & "000****************************************" & chrETX
            MSComm1.Output = lsOrder
            
            SaveOrdLog lsOrder
            
            SQL = "Select barcode from res_flag where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = lsID Then
                SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "' " & vbCrLf & _
                      "where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                res = SendQuery(gLocal, SQL)
            Else
                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                      "Values ('" & lsExamDate & "', '" & lsID & "', '', '', '', '', " & _
                      "'', '', '', '', '', '', " & _
                      "'', '', '', '', '', '" & lsSlideOrd & "' ) "
                res = SendQuery(gLocal, SQL)
            End If
            
            SQL = "Update WorkList set OrdFlag = 'B' where Barcode = '" & lsID & "'"
            'SQL = "Update WorkList set OrdDateTime = '" & Trim(GetText(vasList, lRow, 5)) & "', OrdFlag = 'B' where Barcode = '" & lsID & "'"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
            Else
                SetText vasList, "B", lRow, 2
                CopyRecord lRow
            End If
        
        End If
        
        txtReOrd = ""
    ElseIf KeyCode = vbKeyF3 Then
        If Trim(txtReOrd) = "" Then Exit Sub
            
        lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
    
        lsWkNo = ""
        lsID = Trim(txtReOrd)
                
        mExam = Get_OrderBody(lsID)
        If IsNull(mExam) Then
            MsgBox "샘플정보가 없습니다"
            txtReOrd = ""
            Exit Sub
        End If
        
        k = 1
        For j = LBound(mExam, 2) To UBound(mExam, 2)
            SetText vasExam, mExam(3, j), k, 1
            SetText vasExam, mExam(4, j), k, 2
            k = k + 1
        Next j
        
    
    End If
    
    
    
    Exit Sub
    
ErrHandle:
    SaveQuery "[개별전송]" & Err.Number & ": " & Err.Description
    Resume Next

End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col > 0 Then
            vasSort vasList, Col
        End If
    End If
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsID    As String
    Dim lRow    As Integer
    Dim i       As Integer
    Dim mExam   As Variant
    
    If Row < 1 Or Row > vasList.DataRowCnt Then Exit Sub
    
    lsID = Trim(GetText(vasList, Row, 1))
    
'    mExam = Get_OrderBody(lsID)
'    If Not IsNull(mExam) Then
'        lblBarCode.Caption = lsID
'        ClearSpread vasExam
'        If Trim(GetText(vasList, Row, 3)) = "" Then
'            SetText vasList, mExam(1, LBound(mExam, 2)), Row, 3
'            SetText vasList, mExam(2, LBound(mExam, 2)), Row, 4
'            SQL = "Update worklist set " & vbCrLf & _
'                  " PID = '" & Trim(GetText(vasList, Row, 3)) & "', " & vbCrLf & _
'                  " PName = '" & Trim(GetText(vasList, Row, 4)) & "' " & vbCrLf & _
'                  "where barcode = '" & lsID & "' "
'            res = SendQuery(gLocal, SQL)
'        End If
'        lRow = 1
'        For i = LBound(mExam, 2) To UBound(mExam, 2)
'            SetText vasExam, mExam(3, i), lRow, 1
'            SetText vasExam, mExam(4, i), lRow, 2
'
'            lRow = lRow + 1
'        Next i
'    End If

    res = Online_XML(gXml_S03, lsID)
    If res = 1 Then
        lblBarCode.Caption = lsID
        
        ClearSpread vasExam
        
        If Trim(GetText(vasList, Row, 3)) = "" Then
            SetText vasList, Trim(gPat_Info_Select.PT_NO), Row, 3
            SetText vasList, Trim(gPat_Info_Select.PT_NM), Row, 4
        End If
        
        res = Online_XML(gXml_S07, lsID)
        
        lRow = 1
        
        For i = 0 To UBound(gExam_Select)
            vasExam.SetText 1, lRow, gExam_Select(i).TST_CD
            
            SQL = " Select examname From equipexam Where examcode = '" & gExam_Select(i).TST_CD & "' "
            res = db_select_Col(gLocal, SQL)
            vasExam.SetText 2, lRow, Trim(gReadBuf(0))
            
            lRow = lRow + 1
        Next i
    End If
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyDelete Then
        lRow = vasList.ActiveRow
        
        If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Sub
            
        If MsgBox("검체코드 " & Trim(GetText(vasList, lRow, 1)) & " " & _
                  Trim(GetText(vasList, lRow, 4)) & " 검체를 삭제하시겠습니까? ", vbCritical + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = "Delete from worklist where barcode = '" & Trim(GetText(vasList, lRow, 1)) & "' "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            DeleteRow vasList, lRow, lRow
        End If
    End If
End Sub

Private Sub vasOrder_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col > 0 Then
            vasSort vasOrder, Col
        End If
    End If
End Sub

Private Sub vasOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim lRow As Integer
    Dim i As Integer
    Dim mExam As Variant
    
    If Row < 1 Or Row > vasOrder.DataRowCnt Then Exit Sub
    
    lsID = Trim(GetText(vasOrder, Row, 1))
    mExam = Get_OrderBody(lsID)
    If Not IsNull(mExam) Then
        lblBarCode.Caption = lsID
        ClearSpread vasExam
        
        If Trim(GetText(vasList, Row, 3)) = "" Then
            SetText vasOrder, mExam(1, LBound(mExam, 2)), Row, 3
            SetText vasOrder, mExam(2, LBound(mExam, 2)), Row, 4
            SQL = "Update worklist set " & vbCrLf & _
                  " PID = '" & Trim(GetText(vasOrder, Row, 3)) & "', " & vbCrLf & _
                  " PName = '" & Trim(GetText(vasOrder, Row, 4)) & "' " & vbCrLf & _
                  "where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        lRow = 1
        For i = LBound(mExam, 2) To UBound(mExam, 2)
            SetText vasExam, mExam(3, i), lRow, 1
            SetText vasExam, mExam(4, i), lRow, 2
            
            lRow = lRow + 1
        Next i
    End If
End Sub

Private Sub vasOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyDelete Then
        lRow = vasOrder.ActiveRow
        
        If lRow < 1 Or lRow > vasOrder.DataRowCnt Then Exit Sub
            
        If MsgBox("검체코드 " & Trim(GetText(vasOrder, lRow, 1)) & " " & _
                  Trim(GetText(vasOrder, lRow, 4)) & " 검체를 삭제하시겠습니까? ", vbCritical + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = "Delete from worklist where barcode = '" & Trim(GetText(vasOrder, lRow, 1)) & "' "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            DeleteRow vasOrder, lRow, lRow
        End If
    End If
End Sub

Private Sub Winsock1_Close()
    sckState = -1
End Sub

Private Sub Winsock1_Connect()
    sckState = 0
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim lRow As Long
    Dim lsRsv As String
    Dim lsID As String
    Dim i, j, k As Integer
    Dim mExam As Variant
    
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    
    Dim lsIP As String
    Dim lsHost As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    On Error GoTo ErrHandle
    
    sckState = 1
    
    Winsock1.GetData lsRsv
    
    lsIP = Winsock1.RemoteHostIP
    lsHost = Winsock1.RemoteHost
    
    SaveWinsockLog lsRsv
    
    'SaveQuery "받은 데이타 : " & lsRsv
    
    'lsTotRsv = lsTotRsv & lsRsv
    lsTotRsv = lsRsv
    
    lsRsv = ""
    
    'Check1.Value = 0
    
    If Check1.Value = 1 Then
        sckState = 2
        
        Winsock1.Close
        sckState = -1
        
        Winsock1.LocalPort = MyPort
        Winsock1.Listen
        
        sckState = 0
        
        SP_Search
        
        Exit Sub
    
    End If
    
    i = InStr(1, lsTotRsv, Chr(2))
    j = InStr(1, lsTotRsv, Chr(3))
    
    If i > 0 And j > 0 Then
        lsRsv = Mid(lsTotRsv, i + 1, j - i - 1)
        'lsTotRsv = Mid(lsTotRsv, j + 1)
        
        'SaveQuery "시작과 끝 확인 OK : " & lsRsv
    End If
    
    If Check1.Value = 0 Then
        Sck_Data lsRsv
    
        Exit Sub
    End If
    
    If Trim(lsRsv) <> "" Then
        lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
        
        i = InStr(1, lsRsv, Chr(10))
        Do While i > 0
            lsWkNo = ""
            lsID = Left(lsRsv, i - 1)
            lsRsv = Mid(lsRsv, i + 1)
            
            'SaveQuery lsID & " : 체크시작"
            'SaveQuery "나머지 : " & lsRsv
            
            lRow = vasList.DataRowCnt + 1
            If lRow > vasList.MaxRows Then
                vasList.MaxRows = lRow
            End If
            
            SetText vasList, lsID, lRow, 1
            SetText vasList, "A", lRow, 2
            SetText vasList, GetDateFull, lRow, 5
            SetText vasList, lsIP, lRow, 6
            
            lsSlideOrd = ""
            
            mExam = Get_OrderBody(lsID)
            If Not IsNull(mExam) Then
                SetText vasList, mExam(1, LBound(mExam, 2)), lRow, 3
                SetText vasList, mExam(2, LBound(mExam, 2)), lRow, 4
                
                lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & SetSpace(Trim(mExam(6, LBound(mExam, 2))), 3)
                'lsPID = SetSpace(mExam(1, LBound(mExam, 2)), 8)
                lsPID = SetSpace(Left(mExam(1, LBound(mExam, 2)), 8), 8)
                
                CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(Trim(GetText(vasList, lRow, 5)), 10)
                If Not IsNumeric(gPatGen.Age) Then
                    SetText vasList, gPatGen.Age, lRow, 7
                Else
                    SetText vasList, "", lRow, 7
                End If
                    

'                ClearSpread vasExam
'                For j = 0 To UBound(mExam, 2)
'                    SetText vasExam, mExam(3, j), j, 1
'                    SetText vasExam, mExam(4, j), j, 2
'                Next j
                'SaveQuery "SP OK!"
            Else
                'SaveQuery "SP에서 정보 없음"
            End If
            
            SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP ) " & vbCrLf & _
                  "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'A','','', 0, '" & Trim(GetText(vasList, lRow, 6)) & "') "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL, 1
            End If
            'CopyRecord lRow
            
            
            'Winsock1.SendData "/" & lsID & "???"
            If chkReal.Value = 1 Then
                
                If Not IsNull(mExam) Then
'                    If MSComm1.PortOpen = False Then
'                        LASCPortOpen
'                    End If
                    
                    If MSComm1.PortOpen = True Then
                        MSComm1.PortOpen = False
                    End If
                    LASCPortOpen
                    
                    
                    If MSComm1.CTSHolding = False Then
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
                        lblMsg.ForeColor = RGB(255, 0, 0)
                    Else
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
                        lblMsg.ForeColor = RGB(0, 0, 0)
        
                        For i = 1 To 7
                            Ord(i) = "0"
                        Next i
                        SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
                        Set AdoRs_Exam = db_select_rs(gLocal, SQL)
                        
                        ClearSpread vasExam
                        lblBarCode.Caption = lsID
                        
                        lsSlideOrd = ""
                        k = 1
                        For j = LBound(mExam, 2) To UBound(mExam, 2)
                            SetText vasExam, mExam(3, j), k, 1
                            SetText vasExam, mExam(4, j), k, 2
                            k = k + 1
                            If Not AdoRs_Exam Is Nothing Then
                                AdoRs_Exam.MoveFirst
                                Do Until AdoRs_Exam.EOF
                                    If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                                        Case "C": Ord(1) = "1"
                                        Case "D": Ord(2) = "1"
                                        Case "R": Ord(3) = "1"
                                        Case "P"
                                            Ord(4) = "1"
                                            lsSlideOrd = "SP"
                                        Case "S"
                                            Ord(5) = "1"
                                            lsSlideOrd = "SC"
                                        Case "X": Ord(6) = "1"
                                        Case "B"
                                            If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                                Ord(7) = "1"
                                            End If
                                        End Select
                                        
                                        Exit Do
                                    End If
                                    
                                    AdoRs_Exam.MoveNext
                                Loop
                            End If
                        Next j
                        
                        lsOrder = ""
                        For i = 1 To 7
                            lsOrder = lsOrder & Ord(i)
                        Next i
                                
                        'SaveQuery "실시간 전송 : " & lsID & " : " & lsOrder
                        'MsgBox lsOrder
                        
                        If lsOrder <> "0000000" And lsOrder <> "" Then
                            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                            'lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & lsWkNo & lsPID
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "000****************************************" & chrETX
                            
                            DoSleep 1000
                            
                            MSComm1.Output = lsOrder
                            
                            SaveOrdLog lsOrder
                            
                            SQL = "Select barcode from res_flag where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                            res = db_select_Col(gLocal, SQL)
                            If Trim(gReadBuf(0)) = lsID Then
                                SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "' " & vbCrLf & _
                                      "where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                End If
                            Else
                                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                                      "Values ('" & lsExamDate & "', '" & lsID & "', '', '', '', '', " & _
                                      "'', '', '', '', '', '', " & _
                                      "'', '', '', '', '', '" & lsSlideOrd & "' ) "
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                End If
                            End If
                            
                            SQL = "Update WorkList set OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            'SQL = "Update WorkList set OrdDateTime = '" & Trim(GetText(vasList, lRow, 5)) & "', OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                            Else
                                SetText vasList, "B", lRow, 2
                                CopyRecord lRow
                            End If
                        
                        End If
                        
                    End If
                    
                    'MSComm1.PortOpen = False
                End If
            End If
            i = InStr(1, lsRsv, Chr(10))
        Loop
    End If
    
    'Winsock1.SendData "/???"
    
    sckState = 2
    
    Winsock1.Close
    sckState = -1
    
    Winsock1.LocalPort = MyPort
    Winsock1.Listen
    
    sckState = 0
    
    SP_Search
    
    Exit Sub
ErrHandle:
    SaveQuery "[Winsock]" & Err.Number & ": " & Err.Description
    Resume Next
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    'Debug.Print Number & " : " & Description
    SaveQuery Number & " : " & Description
    
    Resume Next
'    If Winsock1.State <> sckClosed Then
'        Winsock1.Close
'        sckState = -1
'    End If
'
'    Winsock1.LocalPort = MyPort
'    Winsock1.Listen
'    sckState = 0
End Sub

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""
    
    GetSetup = False
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.HostName = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Data", "WorkListExpire", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gExpireDate = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gServerPath = Trim(txtTemp)
    
    GetSetup = True

End Function

Public Sub GetSetup_LASC()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
       
            
    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "Port", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.Port = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.Speed = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.Parity = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.DataBit = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.StopBit = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.RTSEnable = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("ORD_COM", "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gSetup.DTREnable = Trim(txtTemp)

End Sub

Public Function STS(ByVal strStmt As String) As String
    Dim strTmp As String
    
    strTmp = Replace(strStmt, "'", "''")
    
    STS = "'" & strTmp & "'"
End Function


Public Sub SaveWinsockLog(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    Open App.Path & "\Log\Socket.log" For Append As FilNum
    Print #FilNum, Date & " " & Time & " " & argSQL
    Close FilNum
End Sub

Public Sub SaveOrdLog(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    Open App.Path & "\Log\Ord" & Format(Date, "yyyymmdd") & ".log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

Private Sub Winsock1_SendComplete()
'    Winsock1.Close
'    sckState = -1
'    Winsock1.Listen
'    sckState = 0
End Sub

Sub CheckStr(asData As String)
    Dim lRow As Long
    Dim lsRsv As String
    Dim lsID As String
    Dim i, j, k As Integer
    Dim mExam As Variant
    
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    
    Dim lsIP As String
    Dim lsHost As String
    
    Dim lsWkNo As String
    
    On Error GoTo ErrHandle
    
    sckState = 1
    
    'Winsock1.GetData lsRsv
    lsRsv = asData
    lsIP = Winsock1.RemoteHostIP
    lsHost = Winsock1.RemoteHost
    
    SaveWinsockLog lsRsv
    
    Debug.Print "받은 데이타 : " & lsRsv
    
    'lsTotRsv = lsTotRsv & lsRsv
    lsTotRsv = lsRsv
    
    Debug.Print "받은 데이타 전체: " & lsTotRsv
    
    lsRsv = ""
    
    i = InStr(1, lsTotRsv, Chr(2))
    j = InStr(1, lsTotRsv, Chr(3))
    
    If i > 0 And j > 0 Then
        lsRsv = Mid(lsTotRsv, i + 1, j - i - 1)
        'lsTotRsv = Mid(lsTotRsv, j + 1)
        'Debug.Print lsRsv
        'Debug.Print lsTotRsv
        'Debug.Print "시작과 끝 확인 OK"
    End If
    
    If Trim(lsRsv) <> "" Then
        i = InStr(1, lsRsv, Chr(10))
        Do While i > 0
            lsWkNo = ""
            
            lsID = Left(lsRsv, i - 1)
            lsRsv = Mid(lsRsv, i + 1)
            lRow = vasList.DataRowCnt + 1
            If lRow > vasList.MaxRows Then
                vasList.MaxRows = lRow
            End If
            
            SetText vasList, lsID, lRow, 1
            SetText vasList, "A", lRow, 2
            SetText vasList, GetDateFull, lRow, 5
            SetText vasList, lsIP, lRow, 6
            
            mExam = Get_OrderBody(lsID)
            If Not IsNull(mExam) Then
                SetText vasList, mExam(1, LBound(mExam, 2)), lRow, 3
                SetText vasList, mExam(2, LBound(mExam, 2)), lRow, 4
                
                lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2)))
                
                CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(GetDateFull, 10)
                If Not IsNumeric(gPatGen.Age) Then
                    SetText vasList, gPatGen.Age, lRow, 7
                Else
                    SetText vasList, "", lRow, 7
                End If
                    

'                ClearSpread vasExam
'                For j = 0 To UBound(mExam, 2)
'                    SetText vasExam, mExam(3, j), j, 1
'                    SetText vasExam, mExam(4, j), j, 2
'                Next j
            End If
            
            SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP ) " & vbCrLf & _
                  "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'A','','', 0, '" & Trim(GetText(vasList, lRow, 6)) & "') "
            res = SendQuery(gLocal, SQL)
            'CopyRecord lRow
            
            
            'Winsock1.SendData "/" & lsID & "???"
            'chkReal.Value = 0
            If chkReal.Value = 1 Then
                Debug.Print "실시간 전송 : " & lsID
                If Not IsNull(mExam) Then
                    If MSComm1.PortOpen = False Then
                        LASCPortOpen
                    End If
                    
                    If MSComm1.CTSHolding = False Then
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
                        lblMsg.ForeColor = RGB(255, 0, 0)
                    Else
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
                        lblMsg.ForeColor = RGB(0, 0, 0)
        
                        For i = 1 To 7
                            Ord(i) = "0"
                        Next i
                        SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
                        Set AdoRs_Exam = db_select_rs(gLocal, SQL)
                        
                        ClearSpread vasExam
                        lblBarCode.Caption = lsID
                        
                        k = 1
                        For j = LBound(mExam, 2) To UBound(mExam, 2)
                            SetText vasExam, mExam(3, j), k, 1
                            SetText vasExam, mExam(4, j), k, 2
                            k = k + 1
                            If Not AdoRs_Exam Is Nothing Then
                                AdoRs_Exam.MoveFirst
                                Do Until AdoRs_Exam.EOF
                                    If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                                        Case "C": Ord(1) = "1"
                                        Case "D": Ord(2) = "1"
                                        Case "R": Ord(3) = "1"
                                        Case "P": Ord(4) = "1"
                                        Case "S": Ord(5) = "1"
                                        Case "X": Ord(6) = "1"
                                        Case "B"
                                            If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                                Ord(7) = "1"
                                            End If
                                        End Select
                                        
                                        Exit Do
                                    End If
                                    
                                    AdoRs_Exam.MoveNext
                                Loop
                            End If
                        Next j
                        
                        lsOrder = ""
                        For i = 1 To 7
                            lsOrder = lsOrder & Ord(i)
                        Next i
                        
                        'MsgBox lsOrder
                        
                        If lsOrder <> "0000000" And lsOrder <> "" Then
                            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                            'lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & SetSpace(lsWkNo, 13)
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "000****************************************" & chrETX
                            MSComm1.Output = lsOrder
                            
                            SaveOrdLog lsOrder
                            
                            SQL = "Update WorkList set OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            'SQL = "Update WorkList set OrdDateTime = '" & Trim(GetText(vasList, lRow, 5)) & "', OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                Debug.Print SQL
                            Else
                                SetText vasList, "B", lRow, 2
                                CopyRecord lRow
                            End If
                        
                        End If
                        
                    End If
                    
                    'MSComm1.PortOpen = False
                End If
            End If
            i = InStr(1, lsRsv, Chr(10))
        Loop
    End If
    
    'Winsock1.SendData "/???"
    
    sckState = 2
    
    Winsock1.Close
    sckState = -1
    Winsock1.Listen
    sckState = 0
    
    Exit Sub
ErrHandle:
    Debug.Print Err.Number & ": " & Err.Description
    Resume Next
End Sub


Sub Data_proc(asData As String)
    Dim lRow As Long
    Dim lsRsv As String
    Dim lsID As String
    Dim i, j, k As Integer
    Dim mExam As Variant
    
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    
    Dim lsIP As String
    Dim lsHost As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    
    
    On Error GoTo ErrHandle
    
    sckState = 1
    
    
    lsRsv = asData
    
    lsIP = Winsock1.RemoteHostIP
    lsHost = Winsock1.RemoteHost
    
    SaveWinsockLog lsRsv
    
    'SaveQuery "받은 데이타 : " & lsRsv
    
    'lsTotRsv = lsTotRsv & lsRsv
    lsTotRsv = lsRsv
    
    lsRsv = ""
    
    i = InStr(1, lsTotRsv, Chr(2))
    j = InStr(1, lsTotRsv, Chr(3))
    
    If i > 0 And j > 0 Then
        lsRsv = Mid(lsTotRsv, i + 1, j - i - 1)
        'lsTotRsv = Mid(lsTotRsv, j + 1)
        
        'SaveQuery "시작과 끝 확인 OK : " & lsRsv
    End If
    
    If Trim(lsRsv) <> "" Then
        lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
        
        i = InStr(1, lsRsv, Chr(10))
        Do While i > 0
            lsWkNo = ""
            lsID = Left(lsRsv, i - 1)
            lsRsv = Mid(lsRsv, i + 1)
            
            'SaveQuery lsID & " : 체크시작"
            'SaveQuery "나머지 : " & lsRsv
            
            lRow = vasList.DataRowCnt + 1
            If lRow > vasList.MaxRows Then
                vasList.MaxRows = lRow
            End If
            
            SetText vasList, lsID, lRow, 1
            SetText vasList, "A", lRow, 2
            SetText vasList, GetDateFull, lRow, 5
            SetText vasList, lsIP, lRow, 6
            
            lsSlideOrd = ""
            
            mExam = Get_OrderBody(lsID)
            If Not IsNull(mExam) Then
                SetText vasList, mExam(1, LBound(mExam, 2)), lRow, 3
                SetText vasList, mExam(2, LBound(mExam, 2)), lRow, 4
                
                lsWkNo = Trim(mExam(5, LBound(mExam, 2))) & "-" & SetSpace(Trim(mExam(6, LBound(mExam, 2))), 3)
                'lsPID = SetSpace(mExam(1, LBound(mExam, 2)), 8)
                lsPID = SetSpace(Left(mExam(1, LBound(mExam, 2)), 8), 8)
                
                CalSexAge Trim(mExam(7, LBound(mExam, 2))), Left(Trim(GetText(vasList, lRow, 5)), 10)
                If Not IsNumeric(gPatGen.Age) Then
                    SetText vasList, gPatGen.Age, lRow, 7
                Else
                    SetText vasList, "", lRow, 7
                End If
                    

'                ClearSpread vasExam
'                For j = 0 To UBound(mExam, 2)
'                    SetText vasExam, mExam(3, j), j, 1
'                    SetText vasExam, mExam(4, j), j, 2
'                Next j
                'SaveQuery "SP OK!"
            Else
                'SaveQuery "SP에서 정보 없음"
            End If
            
            SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP ) " & vbCrLf & _
                  "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'A','','', 0, '" & Trim(GetText(vasList, lRow, 6)) & "') "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL, 1
            End If
            'CopyRecord lRow
            
            
            'Winsock1.SendData "/" & lsID & "???"
            If chkReal.Value = 1 Then
                
                If Not IsNull(mExam) Then
                    If MSComm1.PortOpen = False Then
                        LASCPortOpen
                    End If
                    
                    If MSComm1.CTSHolding = False Then
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
                        lblMsg.ForeColor = RGB(255, 0, 0)
                    Else
                        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
                        lblMsg.ForeColor = RGB(0, 0, 0)
        
                        For i = 1 To 7
                            Ord(i) = "0"
                        Next i
                        SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
                        Set AdoRs_Exam = db_select_rs(gLocal, SQL)
                        
                        ClearSpread vasExam
                        lblBarCode.Caption = lsID
                        
                        lsSlideOrd = ""
                        k = 1
                        For j = LBound(mExam, 2) To UBound(mExam, 2)
                            SetText vasExam, mExam(3, j), k, 1
                            SetText vasExam, mExam(4, j), k, 2
                            k = k + 1
                            If Not AdoRs_Exam Is Nothing Then
                                AdoRs_Exam.MoveFirst
                                Do Until AdoRs_Exam.EOF
                                    If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                                        Case "C": Ord(1) = "1"
                                        Case "D": Ord(2) = "1"
                                        Case "R": Ord(3) = "1"
                                        Case "P"
                                            Ord(4) = "1"
                                            lsSlideOrd = "SP"
                                        Case "S"
                                            Ord(5) = "1"
                                            lsSlideOrd = "SC"
                                        Case "X": Ord(6) = "1"
                                        Case "B"
                                            If Trim(GetText(vasList, lRow, 7)) <> "" Then
                                                Ord(7) = "1"
                                            End If
                                        End Select
                                        
                                        Exit Do
                                    End If
                                    
                                    AdoRs_Exam.MoveNext
                                Loop
                            End If
                        Next j
                        
                        lsOrder = ""
                        For i = 1 To 7
                            lsOrder = lsOrder & Ord(i)
                        Next i
                                
                        'SaveQuery "실시간 전송 : " & lsID & " : " & lsOrder
                        'MsgBox lsOrder
                        
                        If lsOrder <> "0000000" And lsOrder <> "" Then
                            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                            'lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & lsWkNo & lsPID
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "0000000000000"
                            lsOrder = lsOrder & "000****************************************" & chrETX
                            MSComm1.Output = lsOrder
                            
                            SaveOrdLog lsOrder
                            
                            SQL = "Select barcode from res_flag where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                            res = db_select_Col(gLocal, SQL)
                            If Trim(gReadBuf(0)) = lsID Then
                                SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "' " & vbCrLf & _
                                      "where examdate = '" & lsExamDate & "' and Barcode = '" & lsID & "'"
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                End If
                            Else
                                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                                      "Values ('" & lsExamDate & "', '" & lsID & "', '', '', '', '', " & _
                                      "'', '', '', '', '', '', " & _
                                      "'', '', '', '', '', '" & lsSlideOrd & "' ) "
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                End If
                            End If
                            
                            SQL = "Update WorkList set OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            'SQL = "Update WorkList set OrdDateTime = '" & Trim(GetText(vasList, lRow, 5)) & "', OrdFlag = 'B' where Barcode = '" & lsID & "'"
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                            Else
                                SetText vasList, "B", lRow, 2
                                CopyRecord lRow
                            End If
                        
                        End If
                        
                    End If
                    
                    'MSComm1.PortOpen = False
                End If
            End If
            i = InStr(1, lsRsv, Chr(10))
        Loop
    End If
    
    'Winsock1.SendData "/???"
    
    sckState = 2
    
    Winsock1.Close
    sckState = -1
    Winsock1.Listen
    sckState = 0
    
    Exit Sub
ErrHandle:
    SaveQuery Err.Number & ": " & Err.Description
    Resume Next
End Sub


Private Function SendTheDataCLINILOG(SampleID As String, DeleteFlag As String, ReceDate As String) As Boolean
    Screen.MousePointer = vbHourglass
    SendTheDataCLINILOG = False
    '
    Dim ii, j, k    As Long
    Dim Data        As String
    '
    Dim lsID        As String
    Dim lsJumin     As String
    
    Dim mExam       As Variant
    Dim lsCode      As String
    
    Dim iTIBC, iFe, iWard As Integer
    Dim sGubun      As String
    
    Dim lRow        As Long
    Dim iRerun      As Integer
    
    Dim iCrea_S, iCrea_U As Integer
    Dim iIndexFlag  As Integer
    
    iCrea_S = 0
    iCrea_U = 0
    iIndexFlag = 0
    
    If DeleteFlag = "True" Then
        SendTheDataCLINILOG_DEL SampleID, DeleteFlag, ReceDate
        Exit Function
    ElseIf DeleteFlag = "RE" Then
        SendTheDataCLINILOG_RE SampleID, DeleteFlag, ReceDate
        Exit Function
    End If
        
    Clear_Send
    '
    SendCursor = 0
    'Seq = -1
    
    iTIBC = 0
    iFe = 0
    iWard = 0
    
    iRerun = 0
    
    lsID = Trim(SampleID)
    
    SQL = "select barcode from worklist where barcode = '" & lsID & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = lsID Then
        iRerun = 1
    End If
    
    
    ClearSpread vasTemp
    
    res = Online_XML(gXml_S03, lsID)
    
    If res = 1 Then
        SendTheDataCLINILOG = True
        
        typOrder.FormatTypeCode = "O01," 'O Zero One

        RSet typOrder.SampleID = SetSpace(Trim(lsID), 20, 1)
        typOrder.DateOfReception = ReceDate

        RSet typOrder.PatientID = SetSpace(Trim(gPat_Info_Select.PT_NO), 10, 1)
        
        '-- 2012.05.24 수정
        'typOrder.PatientNameABC = ""
        RSet typOrder.PatientNameABC = SetSpace(Trim(gPat_Info_Select.ACPTNO_1), 20, 2)
        
'        typOrder.PatientNameABC = gPat_Info_Select.ACPTNO_1
        

        typOrder.PatientNameReserve = ""
        RSet typOrder.PatientNameReserve = SetSpace(Trim(gPat_Info_Select.ACPTNO_1), 20, 2)
        
        'typOrder.Birthday = typOrder.DateOfReception
        
        typOrder.Sex = Trim(gPat_Info_Select.Sex)
        '
        If DeleteFlag = True Then
            typOrder.DeleteFlag = "D" 'Delete
        Else
            typOrder.DeleteFlag = "" 'Normal
        End If
        
'        sGubun = Trim(mExam(10, LBound(mExam, 2)))
        sGubun = ""
        
'        If Trim(mExam(10, LBound(mExam, 2))) = "2" Then
'            iWard = 1
'        Else
            iWard = 0
'        End If
'
'        If Trim(mExam(8, LBound(mExam, 2))) = "ED" Or Trim(mExam(8, LBound(mExam, 2))) = "ER" Or Trim(mExam(8, LBound(mExam, 2))) = "EW" Then
'            iWard = 1
'        End If
        
        'typOrder.WardCode = Left(St1("ward"), 3)
        'typOrder.WardName = St1("ward")
        'If IsNull(St1("kwacd1")) Then
        '    typOrder.OrderDeptCode = Left(St1("ward"), 3)
        'Else
        '    typOrder.OrderDeptCode = St1("kwacd1")
        'End If
        'typOrder.OrderDeptName = typOrder.OrderDeptCode
        'typOrder.OrderDrCode = typOrder.OrderDeptCode
        'typOrder.OrderDrName = IIf(IsNull(St1("idoc")), "", St1("idoc"))
        'If typOrder.OrderDrCode <> "" Then
        '    SQL = "SELECT "
        '    SQL = SQL & "       doctname "
        '    SQL = SQL & "  FROM ocsuser.ocsdoctor "
        '    SQL = SQL & " WHERE ROWNUM = 1 "
        '    SQL = SQL & "   AND doctno = '" & Trim(typOrder.OrderDrName) & "' "
        '    St2.Open SQL, conDataBase, adOpenForwardOnly, adLockReadOnly
        '    If St2.EOF = False Then
        '        typOrder.OrderDrName = Left(GetKorToEng(St2("doctname")), 16)
        '    End If
        '    St2.Close
        '    Set St2 = Nothing
        'End If
        'Trim(mExam(8, LBound(mExam, 2))) 병동
        typOrder.WardCode = ""
        typOrder.WardName = ""
        typOrder.OrderDeptCode = ""
        typOrder.OrderDeptName = ""
        typOrder.OrderDrCode = ""
        typOrder.OrderDrName = ""
        'Tube
        'Select Case St2("tubecd")
        '    Case "FB" 'FDP Tube
                typOrder.TypeOfContainer = "01" 'Tube
        '    Case "WB" 'Plain Tube
        '        typOrder.TypeOfContainer = "01" 'Tube
        '    'Case ...
        '        'typOrder.TypeOfContainer = "02" 'Cup
        '    Case Else
        '        typOrder.TypeOfContainer = "03" 'Others
        'End Select
        
        'Sample
        'typOrder.TypeOfSample = "01"
        
        Select Case Mid(gPat_Info_Select.SPC_CD_1, 2)
        Case "SRM"      'Serum
            typOrder.TypeOfSample = "01"
        Case "RUR"      'Urine
            typOrder.TypeOfSample = "02"
        Case Else       'Other
            typOrder.TypeOfSample = "03"
        End Select
        
        typOrder.HeightOfSample = ""
        If typOrder.TypeOfContainer = "01" And typOrder.TypeOfSample = "01" Then
            typOrder.DeCapping = "02" '100mm Tube
        Else
            typOrder.DeCapping = "00"
        End If
        typOrder.Centrifuge = "01"
        '응급여부
        'Select Case St1("emflg")
        '    Case "Y"
        '        typOrder.STATFlag = "01"
        '    Case Else
                typOrder.STATFlag = "00"
        'End Select
        '
'        If IsNumeric(Trim(mExam(6, LBound(mExam, 2)))) Then
'            typOrder.FreeComment = Trim(mExam(5, LBound(mExam, 2))) & "-" & Format(CCur(Trim(mExam(6, LBound(mExam, 2)))), "000")
'        Else
'            typOrder.FreeComment = Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2)))
'        End If

'        If gPat_Info_Select.ACPTNO_1 = "" Then
            typOrder.FreeComment = ""
'        Else
'            typOrder.FreeComment = gPat_Info_Select.ACPTNO_1
'        End If
        '
        res = Online_XML(gXml_S07, lsID)
        typOrder.NumberOfTest = gExamCnt
        '
        Select Case DeleteFlag
            Case False
                
                For j = 0 To UBound(gExam_Select())
'                    If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
'                        SQL = "SELECT OrdCode,  ExamName, Seqno, ExamCode, IndexFlag  "
'                    Else
                        SQL = "SELECT ExamCode, ExamName, Seqno, ExamCode, IndexFlag  "
'                    End If
                    SQL = SQL & CR & _
                          "  From EquipExam " & CR & _
                          " WHERE Equip = '" & gEquip & "'   AND ExamCode = '" & Trim(gExam_Select(j).TST_CD) & "' "
                    res = db_select_Vas(gLocal, SQL, vasTemp, vasTemp.DataRowCnt + 1, 1)
                Next j
                vasSort vasTemp, 3, 1
                
                ii = 0
                For j = 1 To vasTemp.DataRowCnt
                    If Trim(GetText(vasTemp, j, 1)) = "B2056" Then
                        iCrea_S = 1
                    End If
                    If Trim(GetText(vasTemp, j, 1)) = "1162" Then
                        iCrea_U = 1
                    End If
                    
                    If Trim(GetText(vasTemp, j, 5)) = "1" Then
                        iIndexFlag = 1
                    End If
                    
                    
                    k = 0
                    If Left(lsID, 1) = "M" And _
                        (Trim(GetText(vasTemp, j, 4)) = "0108" Or Trim(GetText(vasTemp, j, 4)) = "0288" Or _
                         Trim(GetText(vasTemp, j, 4)) = "0627" Or Trim(GetText(vasTemp, j, 4)) = "0640" Or _
                         Trim(GetText(vasTemp, j, 4)) = "0641") Then
                    Else
                        lsCode = Trim(GetText(vasTemp, j, 1))
                        
                        Select Case Trim(GetText(vasTemp, j, 1))
                        Case "0106" 'PP2 => Glucose 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad("0148", 10, " ")
                        Case "0348", "B0348" 'TIBC 0348 UIBC 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            iTIBC = 1
                        Case "0342", "B0342" 'Fe 오더 확인
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            iFe = 1
'                        Case "0288" 'HIV : 병동은 TLA, 외래는 A0001(Offline)
'                            'gubun => 7:가상,8:임상
'                            If iWard = 1 Or sGubun = "7" Or sGubun = "8" Then
'                                ii = ii + 1
'                                k = 1
'                                RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
'                            Else
'                                If Left(lsID, 1) = "B" Then
'                                    ii = ii + 1
'                                    k = 1
'                                    RSet typOrder.ItemNo(ii) = LPad("A001", 10, " ")
'                                End If
'                            End If
'                        Case "0108" 'HCV : 병동은 TLA, 외래는 A0002(Offline)"
'                            If iWard = 1 Or sGubun = "7" Or sGubun = "8" Then
'                                ii = ii + 1
'                                k = 1
'                                RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
'                            Else
'                                If Left(lsID, 1) = "B" Then
'                                    ii = ii + 1
'                                    k = 1
'                                    RSet typOrder.ItemNo(ii) = LPad("A002", 10, " ")
'                                End If
'                            End If
                        Case Else
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                        End Select
                        
                        If k = 1 Then
                            'If iRerun = 1 Then
                            '    typOrder.TypeOfOrder(ii) = "03" 'Rerun
                            'Else
                                typOrder.TypeOfOrder(ii) = "01" 'New
                            'End If
                            
                        '이전결과 가저오기
                        'If St2.EOF = False Then
                        '    If IsNull(St2("result")) Then
                                typOrder.NoOfAddInf(ii) = "01"
                                typOrder.TypeOfAddInf(ii) = "01"
                                typOrder.AdditionalInf(ii) = ""
                        '    Else
                        '        typOrder.NoOfAddInf(ii) = "01"
                        '        typOrder.TypeOfAddInf(ii) = "01" 'Previous Result
                        '        RSet typOrder.AdditionalInf(ii) = IIf(IsNull(St2("result")), "", St2("result"))
                                
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이상", ">")
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이하", "<")
                                
                        '    End If
                        'Else
                        '    typOrder.NoOfAddInf(ii) = "00"
                        '    typOrder.TypeOfAddInf(ii) = ""
                        '    typOrder.AdditionalInf(ii) = ""
                        'End If
                        End If
                    End If
                Next j
            Case "DELETE"
                '
        End Select
        
        
        
        'TIBC 오더만 난 경우에는 UIBC는 오더 주고 난 뒤 Fe 오더 추가
        If iTIBC = 1 Then
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("B003", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("A003", 10, " ")
            End If
            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = ii
            
'            If iFe = 0 Then
'                ii = ii + 1
'                typOrder.NumberOfTest = Format(ii, "0000")
'                RSet typOrder.ItemNo(ii) = LPad(trim(gettext(vastemp,j,1)), 10, " ")
'                typOrder.TypeOfOrder(ii) = "01" 'New
'                '이전결과 가저오기
'                'If St2.EOF = False Then
'                '    If IsNull(St2("result")) Then
'                typOrder.NoOfAddInf(ii) = "00"
'                typOrder.TypeOfAddInf(ii) = ""
'                typOrder.AdditionalInf(ii) = ""
'            End If
        End If
        
        '생화학 검사가 있는 경우
        If iIndexFlag = 1 Then
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("BLHI01", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("LHI01", 10, " ")
            End If
            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = ii
            
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("BLHI02", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("LHI02", 10, " ")
            End If
            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = ii
            
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("BLHI03", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("LHI03", 10, " ")
            End If

            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = ii
            
        End If
        
        
        j = ii
        typOrder.NumberOfTest = ii
        
        If iCrea_S = 1 And iCrea_U = 1 Then
            j = 0
            For ii = 1 To typOrder.NumberOfTest
                If Trim(typOrder.ItemNo(ii)) = "1161" Then
                    typOrder.ItemNo(ii) = ""
                End If
                If Trim(typOrder.ItemNo(ii)) = "1162" Then
                    typOrder.ItemNo(ii) = ""
                End If
                If Trim(typOrder.ItemNo(ii)) = "0612" Then
                    typOrder.ItemNo(ii) = ""
                End If
                If Trim(typOrder.ItemNo(ii)) <> "" Then
                    j = j + 1
                End If
            Next ii
        End If
                
        '
        Data = ""
        Data = Data & typOrder.FormatTypeCode
        Data = Data & typOrder.SampleID
        Data = Data & typOrder.DateOfReception
        Data = Data & Space(2)
        Data = Data & typOrder.PatientID
        Data = Data & typOrder.PatientNameABC
        Data = Data & typOrder.PatientNameReserve
        Data = Data & typOrder.Birthday
        Data = Data & Space(56)
        Data = Data & typOrder.Sex
        '
        Data = Data & typOrder.DeleteFlag
        Data = Data & typOrder.WardCode
        Data = Data & typOrder.WardName
        Data = Data & typOrder.OrderDeptCode
        Data = Data & typOrder.OrderDeptName
        Data = Data & typOrder.OrderDrCode
        Data = Data & typOrder.OrderDrName
        '
        Data = Data & typOrder.TypeOfContainer
        Data = Data & typOrder.TypeOfSample
        Data = Data & typOrder.HeightOfSample
        Data = Data & typOrder.DeCapping
        Data = Data & typOrder.Centrifuge
        Data = Data & typOrder.STATFlag
        Data = Data & typOrder.FreeComment
        '
        Select Case Len(typOrder.NumberOfTest)
        Case "1"
            Data = Data & Space(3) & typOrder.NumberOfTest
        Case "2"
            Data = Data & Space(2) & typOrder.NumberOfTest
        Case "3"
            Data = Data & Space(1) & typOrder.NumberOfTest
        Case "4"
            Data = Data & typOrder.NumberOfTest
        End Select
        
        For ii = 1 To typOrder.NumberOfTest
            If Trim(typOrder.ItemNo(ii)) <> "" Then
                Data = Data & typOrder.ItemNo(ii)
                Data = Data & typOrder.TypeOfOrder(ii)
                Data = Data & typOrder.NoOfAddInf(ii)

                If typOrder.NoOfAddInf(ii) = "01" Then '01
                    Data = Data & typOrder.TypeOfAddInf(ii)
                    Data = Data & typOrder.AdditionalInf(ii)

                End If
            End If
        Next ii
        
        
        lRow = vasOrder.DataRowCnt + 1
        If lRow > vasOrder.MaxRows Then
            vasOrder.MaxRows = lRow
        End If
        'vasOrder.ActiveRow = lRow
        vasOrder.SetText 1, lRow, Trim(lsID)
        vasOrder.SetText 2, lRow, "B"
        vasOrder.SetText 3, lRow, Trim(gPat_Info_Select.PT_NO)
        vasOrder.SetText 4, lRow, Trim(gPat_Info_Select.PT_NM)
        vasOrder.SetText 5, lRow, Trim(GetText(vasList, 1, 5))
        
        vasActiveCell vasOrder, lRow, 1
        
        SQL = "Delete from worklist where barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
              "Values ('" & Left(Trim(GetText(vasList, 1, 5)), 8) & "', '" & lsID & "', '" & Trim(GetText(vasOrder, lRow, 3)) & "','" & Trim(GetText(vasOrder, lRow, 4)) & "', 'B','','', 0) "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            SetText vasList, "B", lRow, 2
            'CopyRecord lRow
            'DeleteRow vasList1, lRow1, lRow1
        Else
            SaveQuery SQL
        End If
        
    Else
        '
    End If
    '
    If Data <> "" Then
        'If SendFlg = True Then
            'Text1.Text = Text1.Text & "Send:" & vbCrLf & Replace(MakeString(Data), " ", ".") & vbCrLf
            
            Data = chrSTX & Data & chrETX & CStr(MakeCS(Data & chrETX))
            'Data = chrSTX & Data & chrETX
            MSComm1.Output = Data
            'SaveOrdLog "TX : " & Data
            'TXTLOG = txtLog & "TX : " & Data & vbCrLf
            
            txtLog.SelStart = Len(txtLog)
            txtLog.SelLength = 0

        'End If
        '
        SendCursor = 1
        '
        Clear_Send
        '
        typOrder.SampleID = Trim(lsID)
        
        SendTheDataCLINILOG = True
    Else
        SendTheDataCLINILOG = False
    End If
    '
    Screen.MousePointer = vbDefault
End Function

Private Function SendTheDataCLINILOG_DEL(SampleID As String, DeleteFlag As String, ReceDate As String) As Boolean
    Screen.MousePointer = vbHourglass
    SendTheDataCLINILOG_DEL = False
    '
    Dim ii, j, k As Long
    Dim Data As String
    '
    Dim lsID As String
    Dim lsJumin As String
    
    Dim mExam As Variant
    Dim lsCode As String
    
    Dim iTIBC, iFe, iWard As Integer
    Dim sGubun  As String
    
    Dim lRow As Long
    Dim lsReceDate As String
    
    Dim lsReceDate1 As String
    
    Clear_Send
    '
    SendCursor = 0
    'Seq = -1
    
    iTIBC = 0
    iFe = 0
    iWard = 0
    
    lsID = Trim(SampleID)
    
    SQL = "Select ReceDate from WorkList where Barcode = '" & lsID & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) <> "" Then
        lsReceDate = Trim(gReadBuf(0))
    End If
    
    ClearSpread vasTemp
    
    lsReceDate = Left(ReceDate, 4) & "-" & Mid(ReceDate, 5, 2) & "-" & Mid(ReceDate, 7, 2)
    
    'mExam = Get_OrderBody(Trim(lsID))
    mExam = Get_OrderBody_Cancel(Trim(lsID), lsReceDate)
    
    If Not IsNull(mExam) Then
        
        SendTheDataCLINILOG_DEL = True
        
        typOrder.FormatTypeCode = "O01," 'O Zero One
        'lsID = SetSpace(lsID, 20, 1)
        RSet typOrder.SampleID = SetSpace(Trim(lsID), 20, 1)
        typOrder.DateOfReception = ReceDate
        'RSet typOrder.PatientID = LPad(lsID, 10, "0")
        RSet typOrder.PatientID = LPad(Trim(mExam(1, LBound(mExam, 2))), 10, "0")
        'typOrder.PatientNameABC = Left(GetKorToEng(St1("name")), 20)
        typOrder.PatientNameABC = ""    'Conv_Kor_Eng(Trim(mExam(2, LBound(mExam, 2))))
        typOrder.PatientNameReserve = ""
        
        lsJumin = "" '주민번호
            
        typOrder.Birthday = typOrder.DateOfReception
        typOrder.Sex = gPatGen.Sex
        '
        If DeleteFlag = "True" Then
            typOrder.DeleteFlag = "D" 'Delete
            'typOrder.DeleteFlag = "" 'Normal
        Else
            typOrder.DeleteFlag = "" 'Normal
        End If
        
        sGubun = ""
        
'        If Trim(mExam(9, LBound(mExam, 2))) <> "" Then
'            iWard = 1
'        Else
'            iWard = 0
'        End If
        
        If Trim(mExam(8, LBound(mExam, 2))) = "ED" Or _
            Trim(mExam(8, LBound(mExam, 2))) = "ER" Or _
            Trim(mExam(8, LBound(mExam, 2))) = "응급의학과" Or _
            Trim(mExam(8, LBound(mExam, 2))) = "EW" Then
            iWard = 1
        End If
        
        typOrder.WardCode = ""
        typOrder.WardName = ""
        typOrder.OrderDeptCode = ""
        typOrder.OrderDeptName = ""
        typOrder.OrderDrCode = ""
        typOrder.OrderDrName = ""
        'Tube
        'Select Case St2("tubecd")
        '    Case "FB" 'FDP Tube
                typOrder.TypeOfContainer = "01" 'Tube
        '    Case "WB" 'Plain Tube
        '        typOrder.TypeOfContainer = "01" 'Tube
        '    'Case ...
        '        'typOrder.TypeOfContainer = "02" 'Cup
        '    Case Else
        '        typOrder.TypeOfContainer = "03" 'Others
        'End Select
        'Sample
        'Select Case St1("sampcd")
        '    Case "0100" 'Serum
                typOrder.TypeOfSample = "01"
        '    Case "0301" 'Urine
        '        typOrder.TypeOfSample = "02"
        '    Case Else 'Others
        '        typOrder.TypeOfSample = "03"
        'End Select
        typOrder.HeightOfSample = ""
        If typOrder.TypeOfContainer = "01" And typOrder.TypeOfSample = "01" Then
            typOrder.DeCapping = "02" '100mm Tube
        Else
            typOrder.DeCapping = "00"
        End If
        
        typOrder.Centrifuge = "01"
        
        '응급여부
        'Select Case St1("emflg")
        '    Case "Y"
        '        typOrder.STATFlag = "01"
        '    Case Else
                typOrder.STATFlag = "00"
        'End Select
        '
        
        typOrder.FreeComment = ""
        '
        typOrder.NumberOfTest = Format(UBound(mExam, 2) + 1, "0000")
        '
        Select Case DeleteFlag
            Case "True"
                
                For j = LBound(mExam, 2) To UBound(mExam, 2)
                    If Left(Trim(lsID), 1) = "M" Or Left(Trim(lsID), 1) = "B" Then
                        SQL = "SELECT OrdCode, ExamName, Seqno, ExamCode "
                    Else
                        SQL = "SELECT ExamCode, ExamName, Seqno, ExamCode "
                    End If
                    SQL = SQL & CR & _
                          "  From EquipExam " & CR & _
                          " WHERE Equip = '" & gEquip & "'   "
                    SQL = SQL & " AND ExamCode = '" & Trim(mExam(3, j)) & "' "
                    res = db_select_Vas(gLocal, SQL, vasTemp, vasTemp.DataRowCnt + 1, 1)
                Next j
                vasSort vasTemp, 3, 1
                
                ii = 0
                For j = 1 To vasTemp.DataRowCnt
                
                    k = 0
                    If Left(lsID, 1) = "M" And _
                        (Trim(GetText(vasTemp, j, 4)) = "0108" Or Trim(GetText(vasTemp, j, 4)) = "0288") Then
                    Else
                        lsCode = Trim(GetText(vasTemp, j, 1))
                        
                        Select Case Trim(GetText(vasTemp, j, 4))
                        Case "0106" 'PP2 => Glucose 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad("0148", 10, " ")
                        Case "0348" 'TIBC 0348 UIBC 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            iTIBC = 1
                        Case "0342" 'Fe 오더 확인
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            iFe = 1
                        Case "0288" 'HIV : 병동은 TLA, 외래는 A0001(Offline)
                            If iWard = 1 Or sGubun = "7" Or sGubun = "8" Then
                                ii = ii + 1
                                k = 1
                                RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            Else
                                If Left(lsID, 1) = "B" Then
                                    ii = ii + 1
                                    k = 1
                                    RSet typOrder.ItemNo(ii) = LPad("A001", 10, " ")
                                End If
                            End If
                        Case "0108" 'HCV : 병동은 TLA, 외래는 A0002(Offline)"
                            If iWard = 1 Or sGubun = "7" Or sGubun = "8" Then
                                ii = ii + 1
                                k = 1
                                RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            Else
                                If Left(lsID, 1) = "B" Then
                                    ii = ii + 1
                                    k = 1
                                    RSet typOrder.ItemNo(ii) = LPad("A002", 10, " ")
                                End If
                            End If
                        Case Else
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                        End Select
                        
                        If k = 1 Then
                            typOrder.TypeOfOrder(ii) = "02" '01:New,02:Delete
                            'typOrder.TypeOfOrder(ii) = "01" '01:New,02:Delete
                            'typOrder.TypeOfOrder(ii) = "03" '01:New,02:Delete,03:Rerun
                        '이전결과 가저오기
                        'If St2.EOF = False Then
                        '    If IsNull(St2("result")) Then
                                typOrder.NoOfAddInf(ii) = "00"
                                typOrder.TypeOfAddInf(ii) = ""
                                typOrder.AdditionalInf(ii) = ""
                        '    Else
                        '        typOrder.NoOfAddInf(ii) = "01"
                        '        typOrder.TypeOfAddInf(ii) = "01" 'Previous Result
                        '        RSet typOrder.AdditionalInf(ii) = IIf(IsNull(St2("result")), "", St2("result"))
                                
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이상", ">")
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이하", "<")
                                
                        '    End If
                        'Else
                        '    typOrder.NoOfAddInf(ii) = "00"
                        '    typOrder.TypeOfAddInf(ii) = ""
                        '    typOrder.AdditionalInf(ii) = ""
                        'End If
                        End If
                    End If
                Next j
            Case "DELETE"
                '
        End Select
        
        
        
        'TIBC 오더만 난 경우에는 UIBC는 오더 주고 난 뒤 Fe 오더 추가
        If iTIBC = 1 Then
            ii = ii + 1
            
            If Left(Trim(lsID), 1) = "M" Or Left(Trim(lsID), 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("B003", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("A003", 10, " ")
            End If
            typOrder.TypeOfOrder(ii) = "02" 'delete
            'typOrder.TypeOfOrder(ii) = "01" 'delete
            'typOrder.TypeOfOrder(ii) = "03" '01:New,02:Delete,03:Rerun
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = Format(ii, "0000")
            
'            If iFe = 0 Then
'                ii = ii + 1
'                typOrder.NumberOfTest = Format(ii, "0000")
'                RSet typOrder.ItemNo(ii) = LPad(trim(gettext(vastemp,j,1)), 10, " ")
'                typOrder.TypeOfOrder(ii) = "02" 'New
'                '이전결과 가저오기
'                'If St2.EOF = False Then
'                '    If IsNull(St2("result")) Then
'                typOrder.NoOfAddInf(ii) = "00"
'                typOrder.TypeOfAddInf(ii) = ""
'                typOrder.AdditionalInf(ii) = ""
'            End If
        End If
        
        typOrder.NumberOfTest = Format(ii, "0000")
        
        '
        Data = ""
        Data = Data & typOrder.FormatTypeCode
        Data = Data & typOrder.SampleID
        Data = Data & typOrder.DateOfReception
        Data = Data & typOrder.PatientID
        Data = Data & typOrder.PatientNameABC
        Data = Data & typOrder.PatientNameReserve
        Data = Data & typOrder.Birthday
        Data = Data & typOrder.Sex
        '
        Data = Data & typOrder.DeleteFlag
        Data = Data & typOrder.WardCode
        Data = Data & typOrder.WardName
        Data = Data & typOrder.OrderDeptCode
        Data = Data & typOrder.OrderDeptName
        Data = Data & typOrder.OrderDrCode
        Data = Data & typOrder.OrderDrName
        '
        Data = Data & typOrder.TypeOfContainer
        Data = Data & typOrder.TypeOfSample
        Data = Data & typOrder.HeightOfSample
        Data = Data & typOrder.DeCapping
        Data = Data & typOrder.Centrifuge
        Data = Data & typOrder.STATFlag
        Data = Data & typOrder.FreeComment
        '
        Data = Data & typOrder.NumberOfTest
        For ii = 1 To typOrder.NumberOfTest
            Data = Data & typOrder.ItemNo(ii)
            Data = Data & typOrder.TypeOfOrder(ii)
            Data = Data & typOrder.NoOfAddInf(ii)
            If typOrder.NoOfAddInf(ii) = "01" Then '01
                Data = Data & typOrder.TypeOfAddInf(ii)
                Data = Data & typOrder.AdditionalInf(ii)
            End If
        Next ii
        
        
        lRow = vasOrder.DataRowCnt + 1
        If lRow > vasOrder.MaxRows Then
            vasOrder.MaxRows = lRow
        End If
        'vasOrder.ActiveRow = lRow
        vasOrder.SetText 1, lRow, Trim(lsID)
        vasOrder.SetText 2, lRow, "B"
        vasOrder.SetText 3, lRow, Trim(mExam(1, LBound(mExam, 2)))
        vasOrder.SetText 4, lRow, Trim(mExam(2, LBound(mExam, 2)))
        vasOrder.SetText 5, lRow, Trim(GetText(vasList, 1, 5))
        
        vasActiveCell vasOrder, lRow, 1
        
        SQL = "Delete from worklist where barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
              "Values ('" & Left(Trim(GetText(vasList, 1, 5)), 8) & "', '" & lsID & "', '" & Trim(GetText(vasOrder, lRow, 3)) & "','" & Trim(GetText(vasOrder, lRow, 4)) & "', 'B','','', 0) "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            SetText vasList, "B", lRow, 2
            'CopyRecord lRow
            'DeleteRow vasList1, lRow1, lRow1
        Else
            SaveQuery SQL
        End If
        
    Else
        '
    End If
    '
    If Data <> "" Then
        'If SendFlg = True Then
            'Text1.Text = Text1.Text & "Send:" & vbCrLf & Replace(MakeString(Data), " ", ".") & vbCrLf
            
            'Data = chrSTX & Data & chrETX & CStr(MakeCS(Data & chrETX))
            Data = chrSTX & Data & chrETX
            MSComm1.Output = Data
            SaveOrdLog "TX : " & Data
            'TXTLOG = txtLog & "TX : " & Data & vbCrLf
            
            txtLog.SelStart = Len(txtLog)
            txtLog.SelLength = 0

        'End If
        '
        SendCursor = 1
        '
'        conDataBase.BeginTrans
'        '
'        SQL = "UPDATE cpusr.temp_clinilog "
'        If Val(typOrder.NumberOfTest) = 0 Then
'            SQL = SQL & "   SET status = 'Y', " '전송시도(삭제)
'        Else
'            SQL = SQL & "   SET status = DECODE(status,'A','B','X','Y',status), " '전송시도
'        End If
'        SQL = SQL & "       senddate = DECODE(status,'A',TO_CHAR(SYSDATE,'YYYYMMDD'),'B',senddate), "
'        SQL = SQL & "       sendtime = DECODE(status,'A',TO_CHAR(SYSDATE,'HH24MISS'),'B',sendtime), "
'        SQL = SQL & "       lastdate = DECODE(status,'A',NULL,'B',TO_CHAR(SYSDATE,'YYYYMMDD')), "
'        SQL = SQL & "       lasttime = DECODE(status,'A',NULL,'B',TO_CHAR(SYSDATE,'HH24MISS')) "
'        SQL = SQL & " WHERE sid = '" & Trim(typOrder.SampleID) & "' "
'        SQL = SQL & "   AND seqn = 0 "
'        SQL = SQL & "   AND status IN ('A','B','X','Y') " '접수, 전송시도
'        conDataBase.Execute SQL
'        '
'        conDataBase.CommitTrans
        '
        Clear_Send
        '
        typOrder.SampleID = Trim(lsID)
        
        SendTheDataCLINILOG_DEL = True
    Else
        SendTheDataCLINILOG_DEL = False
    End If
    '
    
    Screen.MousePointer = vbDefault
End Function

Private Function SendTheDataCLINILOG_RE(SampleID As String, DeleteFlag As String, ReceDate As String) As Boolean
    Screen.MousePointer = vbHourglass
    SendTheDataCLINILOG_RE = False
    '
    Dim ii, j, k As Long
    Dim Data As String
    '
    Dim lsID As String
    Dim lsJumin As String
    
    Dim mExam As Variant
    Dim lsCode As String
    
    Dim iTIBC, iFe, iWard As Integer
    Dim sGubun  As String
    
    Dim lRow As Long
    Dim iRerun As Integer
    
    Dim iCrea_S, iCrea_U As Integer
    Dim iIndexFlag As Integer
    
    iCrea_S = 0
    iCrea_U = 0
    iIndexFlag = 0
        
    Clear_Send
    '
    SendCursor = 0
    'Seq = -1
    
    iTIBC = 0
    iFe = 0
    iWard = 0
    
    iRerun = 1
    
    lsID = Trim(SampleID)
    
    SQL = "select barcode from worklist where barcode = '" & lsID & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = lsID Then
        iRerun = 1
    End If
    
    
    ClearSpread vasTemp
    
    mExam = Get_OrderBody(Trim(lsID))
    If Not IsNull(mExam) Then
        
        SendTheDataCLINILOG_RE = True
        
        typOrder.FormatTypeCode = "O01," 'O Zero One
        'lsID = SetSpace(lsID, 20, 1)
        RSet typOrder.SampleID = SetSpace(Trim(lsID), 20, 1)
        typOrder.DateOfReception = ReceDate
        'RSet typOrder.PatientID = LPad(lsID, 10, "0")
        RSet typOrder.PatientID = LPad(Trim(mExam(1, LBound(mExam, 2))), 10, "0")
        'typOrder.PatientNameABC = Left(GetKorToEng(St1("name")), 20)
        typOrder.PatientNameABC = ""    'Conv_Kor_Eng(Trim(mExam(2, LBound(mExam, 2))))
        typOrder.PatientNameReserve = ""
        
        lsJumin = Trim(mExam(7, LBound(mExam, 2))) '주민번호
        CalSexAge lsJumin, ReceDate
        
        If gPatGen.Birth <> "" Then
            typOrder.Birthday = Format(CDate(gPatGen.Birth), "yyyymmdd")
        Else
            typOrder.Birthday = typOrder.DateOfReception
        End If
        typOrder.Sex = gPatGen.Sex
        '
        If DeleteFlag = "True" Then
            typOrder.DeleteFlag = "D" 'Delete
        Else
            typOrder.DeleteFlag = "" 'Normal
        End If
        
        sGubun = Trim(mExam(10, LBound(mExam, 2)))
        
        If Trim(mExam(10, LBound(mExam, 2))) = "2" Then
            iWard = 1
        Else
            iWard = 0
        End If
        
        If Trim(mExam(8, LBound(mExam, 2))) = "ED" Or Trim(mExam(8, LBound(mExam, 2))) = "ER" Or Trim(mExam(8, LBound(mExam, 2))) = "EW" Then
            iWard = 1
        End If
        
        'typOrder.WardCode = Left(St1("ward"), 3)
        'typOrder.WardName = St1("ward")
        'If IsNull(St1("kwacd1")) Then
        '    typOrder.OrderDeptCode = Left(St1("ward"), 3)
        'Else
        '    typOrder.OrderDeptCode = St1("kwacd1")
        'End If
        'typOrder.OrderDeptName = typOrder.OrderDeptCode
        'typOrder.OrderDrCode = typOrder.OrderDeptCode
        'typOrder.OrderDrName = IIf(IsNull(St1("idoc")), "", St1("idoc"))
        'If typOrder.OrderDrCode <> "" Then
        '    SQL = "SELECT "
        '    SQL = SQL & "       doctname "
        '    SQL = SQL & "  FROM ocsuser.ocsdoctor "
        '    SQL = SQL & " WHERE ROWNUM = 1 "
        '    SQL = SQL & "   AND doctno = '" & Trim(typOrder.OrderDrName) & "' "
        '    St2.Open SQL, conDataBase, adOpenForwardOnly, adLockReadOnly
        '    If St2.EOF = False Then
        '        typOrder.OrderDrName = Left(GetKorToEng(St2("doctname")), 16)
        '    End If
        '    St2.Close
        '    Set St2 = Nothing
        'End If
        'Trim(mExam(8, LBound(mExam, 2))) 병동
        typOrder.WardCode = ""
        typOrder.WardName = ""
        typOrder.OrderDeptCode = ""
        typOrder.OrderDeptName = ""
        typOrder.OrderDrCode = ""
        typOrder.OrderDrName = ""
        'Tube
        'Select Case St2("tubecd")
        '    Case "FB" 'FDP Tube
                typOrder.TypeOfContainer = "01" 'Tube
        '    Case "WB" 'Plain Tube
        '        typOrder.TypeOfContainer = "01" 'Tube
        '    'Case ...
        '        'typOrder.TypeOfContainer = "02" 'Cup
        '    Case Else
        '        typOrder.TypeOfContainer = "03" 'Others
        'End Select
        'Sample
        'Select Case St1("sampcd")
        '    Case "0100" 'Serum
                typOrder.TypeOfSample = "01"
        '    Case "0301" 'Urine
        '        typOrder.TypeOfSample = "02"
        '    Case Else 'Others
        '        typOrder.TypeOfSample = "03"
        'End Select
        typOrder.HeightOfSample = ""
        If typOrder.TypeOfContainer = "01" And typOrder.TypeOfSample = "01" Then
            typOrder.DeCapping = "02" '100mm Tube
        Else
            typOrder.DeCapping = "00"
        End If
        typOrder.Centrifuge = "01"
        '응급여부
        'Select Case St1("emflg")
        '    Case "Y"
        '        typOrder.STATFlag = "01"
        '    Case Else
                typOrder.STATFlag = "00"
        'End Select
        '
        If IsNumeric(Trim(mExam(6, LBound(mExam, 2)))) Then
            typOrder.FreeComment = Trim(mExam(5, LBound(mExam, 2))) & "-" & Format(CCur(Trim(mExam(6, LBound(mExam, 2)))), "000")
        Else
            typOrder.FreeComment = Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2)))
        End If
        '
        typOrder.NumberOfTest = Format(UBound(mExam, 2) + 1, "0000")
        '
          
        ClearSpread vasTemp1
        SQL = "SELECT itemcode FROM LABORATORYPRESCRIBE"
        SQL = SQL & vbCrLf & " WHERE BARCODENUMBER = '" & Trim(lsID) & "' AND HOLD='1' "
        res = db_select_Vas(gServer, SQL, vasTemp1)

                
                For j = 1 To vasTemp1.DataRowCnt
                    If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                        SQL = "SELECT OrdCode,  ExamName, Seqno, ExamCode, IndexFlag  "
                    Else
                        SQL = "SELECT ExamCode, ExamName, Seqno, ExamCode, IndexFlag  "
                    End If
                    SQL = SQL & CR & _
                          "  From EquipExam " & CR & _
                          " WHERE Equip = '" & gEquip & "'   AND ExamCode = '" & Trim(GetText(vasTemp1, j, 1)) & "' "
                    res = db_select_Vas(gLocal, SQL, vasTemp, vasTemp.DataRowCnt + 1, 1)
                Next j
                vasSort vasTemp, 3, 1
                
                ii = 0
                For j = 1 To vasTemp.DataRowCnt
                    If Trim(GetText(vasTemp, j, 1)) = "B2056" Then
                        iCrea_S = 1
                    End If
                    If Trim(GetText(vasTemp, j, 1)) = "1162" Then
                        iCrea_U = 1
                    End If
                    
                    If Trim(GetText(vasTemp, j, 5)) = "1" Then
                        iIndexFlag = 1
                    End If
                    
                    
                    k = 0
                    If Left(lsID, 1) = "M" And _
                        (Trim(GetText(vasTemp, j, 4)) = "0108" Or Trim(GetText(vasTemp, j, 4)) = "0288" Or _
                         Trim(GetText(vasTemp, j, 4)) = "0627" Or Trim(GetText(vasTemp, j, 4)) = "0640" Or _
                         Trim(GetText(vasTemp, j, 4)) = "0641") Then
                    Else
                        lsCode = Trim(GetText(vasTemp, j, 1))
                        
                        Select Case Trim(GetText(vasTemp, j, 1))
                        Case "0106" 'PP2 => Glucose 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad("0148", 10, " ")
                        Case "0348", "B0348" 'TIBC 0348 UIBC 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            iTIBC = 1
                        Case "0342", "B0342" 'Fe 오더 확인
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            iFe = 1
                        Case "0288" 'HIV : 병동은 TLA, 외래는 A0001(Offline)
                            'gubun => 7:가상,8:임상
                            If iWard = 1 Or sGubun = "7" Or sGubun = "8" Then
                                ii = ii + 1
                                k = 1
                                RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            Else
                                If Left(lsID, 1) = "B" Then
                                    ii = ii + 1
                                    k = 1
                                    RSet typOrder.ItemNo(ii) = LPad("A001", 10, " ")
                                End If
                            End If
                        Case "0108" 'HCV : 병동은 TLA, 외래는 A0002(Offline)"
                            If iWard = 1 Or sGubun = "7" Or sGubun = "8" Then
                                ii = ii + 1
                                k = 1
                                RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                            Else
                                If Left(lsID, 1) = "B" Then
                                    ii = ii + 1
                                    k = 1
                                    RSet typOrder.ItemNo(ii) = LPad("A002", 10, " ")
                                End If
                            End If
                        Case Else
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(GetText(vasTemp, j, 1)), 10, " ")
                        End Select
                        
                        If k = 1 Then
                            If iRerun = 1 Then
                                typOrder.TypeOfOrder(ii) = "03" 'Rerun
                            Else
                                typOrder.TypeOfOrder(ii) = "01" 'New
                            End If
                            
                        '이전결과 가저오기
                        'If St2.EOF = False Then
                        '    If IsNull(St2("result")) Then
                                typOrder.NoOfAddInf(ii) = "00"
                                typOrder.TypeOfAddInf(ii) = ""
                                typOrder.AdditionalInf(ii) = ""
                        '    Else
                        '        typOrder.NoOfAddInf(ii) = "01"
                        '        typOrder.TypeOfAddInf(ii) = "01" 'Previous Result
                        '        RSet typOrder.AdditionalInf(ii) = IIf(IsNull(St2("result")), "", St2("result"))
                                
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이상", ">")
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이하", "<")
                                
                        '    End If
                        'Else
                        '    typOrder.NoOfAddInf(ii) = "00"
                        '    typOrder.TypeOfAddInf(ii) = ""
                        '    typOrder.AdditionalInf(ii) = ""
                        'End If
                        End If
                    End If
                Next j

        
        
        
        'TIBC 오더만 난 경우에는 UIBC는 오더 주고 난 뒤 Fe 오더 추가
        If iTIBC = 1 Then
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("B003", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("A003", 10, " ")
            End If
            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = Format(ii, "0000")
            
'            If iFe = 0 Then
'                ii = ii + 1
'                typOrder.NumberOfTest = Format(ii, "0000")
'                RSet typOrder.ItemNo(ii) = LPad(trim(gettext(vastemp,j,1)), 10, " ")
'                typOrder.TypeOfOrder(ii) = "01" 'New
'                '이전결과 가저오기
'                'If St2.EOF = False Then
'                '    If IsNull(St2("result")) Then
'                typOrder.NoOfAddInf(ii) = "00"
'                typOrder.TypeOfAddInf(ii) = ""
'                typOrder.AdditionalInf(ii) = ""
'            End If
        End If
        
        '생화학 검사가 있는 경우
        If iIndexFlag = 1 Then
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("BLHI01", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("LHI01", 10, " ")
            End If
            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = Format(ii, "0000")
            
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("BLHI02", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("LHI02", 10, " ")
            End If
            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = Format(ii, "0000")
            
            ii = ii + 1
            
            If Left(lsID, 1) = "M" Or Left(lsID, 1) = "B" Then
                RSet typOrder.ItemNo(ii) = LPad("BLHI03", 10, " ")
            Else
                RSet typOrder.ItemNo(ii) = LPad("LHI03", 10, " ")
            End If

            If iRerun = 1 Then
                typOrder.TypeOfOrder(ii) = "03" 'Rerun
            Else
                typOrder.TypeOfOrder(ii) = "01" 'New
            End If
            
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = Format(ii, "0000")
            
        End If
        
        
        j = ii
        typOrder.NumberOfTest = Format(ii, "0000")
        
        If iCrea_S = 1 And iCrea_U = 1 Then
            j = 0
            For ii = 1 To typOrder.NumberOfTest
                If Trim(typOrder.ItemNo(ii)) = "1161" Then
                    typOrder.ItemNo(ii) = ""
                End If
                If Trim(typOrder.ItemNo(ii)) = "1162" Then
                    typOrder.ItemNo(ii) = ""
                End If
                If Trim(typOrder.ItemNo(ii)) = "0612" Then
                    typOrder.ItemNo(ii) = ""
                End If
                If Trim(typOrder.ItemNo(ii)) <> "" Then
                    j = j + 1
                End If
            Next ii
        End If
                
        '
        Data = ""
        Data = Data & typOrder.FormatTypeCode
        Data = Data & typOrder.SampleID
        Data = Data & typOrder.DateOfReception
        Data = Data & typOrder.PatientID
        Data = Data & typOrder.PatientNameABC
        Data = Data & typOrder.PatientNameReserve
        Data = Data & typOrder.Birthday
        Data = Data & typOrder.Sex
        '
        Data = Data & typOrder.DeleteFlag
        Data = Data & typOrder.WardCode
        Data = Data & typOrder.WardName
        Data = Data & typOrder.OrderDeptCode
        Data = Data & typOrder.OrderDeptName
        Data = Data & typOrder.OrderDrCode
        Data = Data & typOrder.OrderDrName
        '
        Data = Data & typOrder.TypeOfContainer
        Data = Data & typOrder.TypeOfSample
        Data = Data & typOrder.HeightOfSample
        Data = Data & typOrder.DeCapping
        Data = Data & typOrder.Centrifuge
        Data = Data & typOrder.STATFlag
        Data = Data & typOrder.FreeComment
        '
        'Data = Data & typOrder.NumberOfTest
        Data = Data & Format(j, "0000")
        
        For ii = 1 To typOrder.NumberOfTest
            If Trim(typOrder.ItemNo(ii)) <> "" Then
                Data = Data & typOrder.ItemNo(ii)
                Data = Data & typOrder.TypeOfOrder(ii)
                Data = Data & typOrder.NoOfAddInf(ii)
                If typOrder.NoOfAddInf(ii) = "01" Then '01
                    Data = Data & typOrder.TypeOfAddInf(ii)
                    Data = Data & typOrder.AdditionalInf(ii)
                End If
            End If
        Next ii
        
        
        lRow = vasOrder.DataRowCnt + 1
        If lRow > vasOrder.MaxRows Then
            vasOrder.MaxRows = lRow
        End If
        'vasOrder.ActiveRow = lRow
        vasOrder.SetText 1, lRow, Trim(lsID)
        vasOrder.SetText 2, lRow, "B"
        vasOrder.SetText 3, lRow, Trim(mExam(1, LBound(mExam, 2)))
        vasOrder.SetText 4, lRow, Trim(mExam(2, LBound(mExam, 2)))
        vasOrder.SetText 5, lRow, Trim(GetText(vasList, 1, 5))
        
        vasActiveCell vasOrder, lRow, 1
        
        SQL = "Delete from worklist where barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
              "Values ('" & Left(Trim(GetText(vasList, 1, 5)), 8) & "', '" & lsID & "', '" & Trim(GetText(vasOrder, lRow, 3)) & "','" & Trim(GetText(vasOrder, lRow, 4)) & "', 'B','','', 0) "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            SetText vasList, "B", lRow, 2
            'CopyRecord lRow
            'DeleteRow vasList1, lRow1, lRow1
        Else
            SaveQuery SQL
        End If
        
    Else
        '
    End If
    '
    If Data <> "" Then
        'If SendFlg = True Then
            'Text1.Text = Text1.Text & "Send:" & vbCrLf & Replace(MakeString(Data), " ", ".") & vbCrLf
            
            'Data = chrSTX & Data & chrETX & CStr(MakeCS(Data & chrETX))
            Data = chrSTX & Data & chrETX
            MSComm1.Output = Data
            SaveOrdLog "TX : " & Data
            'TXTLOG = txtLog & "TX : " & Data & vbCrLf
            
            txtLog.SelStart = Len(txtLog)
            txtLog.SelLength = 0

        'End If
        '
        SendCursor = 1
        '
'        conDataBase.BeginTrans
'        '
'        SQL = "UPDATE cpusr.temp_clinilog "
'        If Val(typOrder.NumberOfTest) = 0 Then
'            SQL = SQL & "   SET status = 'Y', " '전송시도(삭제)
'        Else
'            SQL = SQL & "   SET status = DECODE(status,'A','B','X','Y',status), " '전송시도
'        End If
'        SQL = SQL & "       senddate = DECODE(status,'A',TO_CHAR(SYSDATE,'YYYYMMDD'),'B',senddate), "
'        SQL = SQL & "       sendtime = DECODE(status,'A',TO_CHAR(SYSDATE,'HH24MISS'),'B',sendtime), "
'        SQL = SQL & "       lastdate = DECODE(status,'A',NULL,'B',TO_CHAR(SYSDATE,'YYYYMMDD')), "
'        SQL = SQL & "       lasttime = DECODE(status,'A',NULL,'B',TO_CHAR(SYSDATE,'HH24MISS')) "
'        SQL = SQL & " WHERE sid = '" & Trim(typOrder.SampleID) & "' "
'        SQL = SQL & "   AND seqn = 0 "
'        SQL = SQL & "   AND status IN ('A','B','X','Y') " '접수, 전송시도
'        conDataBase.Execute SQL
'        '
'        conDataBase.CommitTrans
        '
        Clear_Send
        '
        typOrder.SampleID = Trim(lsID)
        
        SendTheDataCLINILOG_RE = True
    Else
        SendTheDataCLINILOG_RE = False
    End If
    '
    
    Screen.MousePointer = vbDefault
End Function

Private Function SendTheDataCLINILOG_1(SampleID As String, DeleteFlag As Boolean, ReceDate As String) As Boolean
    Screen.MousePointer = vbHourglass
    SendTheDataCLINILOG_1 = False
    '
    Dim ii, j, k As Long
    Dim Data As String
    '
    Dim lsID As String
    Dim lsJumin As String
    
    Dim mExam As Variant
    Dim lsCode As String
    
    Dim iTIBC, iFe, iWard As Integer
    
    Dim lRow As Long
    
    Clear_Send
    '
    SendCursor = 0
    'Seq = -1
    
    iTIBC = 0
    iFe = 0
    iWard = 0
    
    lsID = Trim(SampleID)
    
    ClearSpread vasTemp
    
    mExam = Get_OrderBody(Trim(lsID))
    If Not IsNull(mExam) Then
        
        SendTheDataCLINILOG_1 = True
        
        typOrder.FormatTypeCode = "O01," 'O Zero One
        'lsID = SetSpace(lsID, 20, 1)
        RSet typOrder.SampleID = SetSpace(Trim(lsID), 20, 1)
        typOrder.DateOfReception = ReceDate
        'RSet typOrder.PatientID = LPad(lsID, 10, "0")
        RSet typOrder.PatientID = LPad(Trim(mExam(1, LBound(mExam, 2))), 10, "0")
        'typOrder.PatientNameABC = Left(GetKorToEng(St1("name")), 20)
        typOrder.PatientNameABC = ""    'Conv_Kor_Eng(Trim(mExam(2, LBound(mExam, 2))))
        typOrder.PatientNameReserve = ""
        
        lsJumin = Trim(mExam(7, LBound(mExam, 2))) '주민번호
        CalSexAge lsJumin, ReceDate
        
        If gPatGen.Birth <> "" Then
            typOrder.Birthday = Format(CDate(gPatGen.Birth), "yyyymmdd")
        Else
            typOrder.Birthday = typOrder.DateOfReception
        End If
        typOrder.Sex = gPatGen.Sex
        '
        If DeleteFlag = True Then
            typOrder.DeleteFlag = "D" 'Delete
        Else
            typOrder.DeleteFlag = "" 'Normal
        End If
        
        If Trim(mExam(10, LBound(mExam, 2))) = "2" Then
            iWard = 1
        Else
            iWard = 0
        End If
        
        'typOrder.WardCode = Left(St1("ward"), 3)
        'typOrder.WardName = St1("ward")
        'If IsNull(St1("kwacd1")) Then
        '    typOrder.OrderDeptCode = Left(St1("ward"), 3)
        'Else
        '    typOrder.OrderDeptCode = St1("kwacd1")
        'End If
        'typOrder.OrderDeptName = typOrder.OrderDeptCode
        'typOrder.OrderDrCode = typOrder.OrderDeptCode
        'typOrder.OrderDrName = IIf(IsNull(St1("idoc")), "", St1("idoc"))
        'If typOrder.OrderDrCode <> "" Then
        '    SQL = "SELECT "
        '    SQL = SQL & "       doctname "
        '    SQL = SQL & "  FROM ocsuser.ocsdoctor "
        '    SQL = SQL & " WHERE ROWNUM = 1 "
        '    SQL = SQL & "   AND doctno = '" & Trim(typOrder.OrderDrName) & "' "
        '    St2.Open SQL, conDataBase, adOpenForwardOnly, adLockReadOnly
        '    If St2.EOF = False Then
        '        typOrder.OrderDrName = Left(GetKorToEng(St2("doctname")), 16)
        '    End If
        '    St2.Close
        '    Set St2 = Nothing
        'End If
        'Trim(mExam(8, LBound(mExam, 2))) 병동
        typOrder.WardCode = ""
        typOrder.WardName = ""
        typOrder.OrderDeptCode = ""
        typOrder.OrderDeptName = ""
        typOrder.OrderDrCode = ""
        typOrder.OrderDrName = ""
        'Tube
        'Select Case St2("tubecd")
        '    Case "FB" 'FDP Tube
                typOrder.TypeOfContainer = "01" 'Tube
        '    Case "WB" 'Plain Tube
        '        typOrder.TypeOfContainer = "01" 'Tube
        '    'Case ...
        '        'typOrder.TypeOfContainer = "02" 'Cup
        '    Case Else
        '        typOrder.TypeOfContainer = "03" 'Others
        'End Select
        'Sample
        'Select Case St1("sampcd")
        '    Case "0100" 'Serum
                typOrder.TypeOfSample = "01"
        '    Case "0301" 'Urine
        '        typOrder.TypeOfSample = "02"
        '    Case Else 'Others
        '        typOrder.TypeOfSample = "03"
        'End Select
        typOrder.HeightOfSample = ""
        If typOrder.TypeOfContainer = "01" And typOrder.TypeOfSample = "01" Then
            typOrder.DeCapping = "02" '100mm Tube
        Else
            typOrder.DeCapping = "00"
        End If
        typOrder.Centrifuge = "01"
        '응급여부
        'Select Case St1("emflg")
        '    Case "Y"
        '        typOrder.STATFlag = "01"
        '    Case Else
                typOrder.STATFlag = "00"
        'End Select
        '
        If IsNumeric(Trim(mExam(6, LBound(mExam, 2)))) Then
            typOrder.FreeComment = Trim(mExam(5, LBound(mExam, 2))) & "-" & Format(CCur(Trim(mExam(6, LBound(mExam, 2)))), "000")
        Else
            typOrder.FreeComment = Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2)))
        End If
        '
        typOrder.NumberOfTest = Format(UBound(mExam, 2) + 1, "0000")
        '
        Select Case DeleteFlag
            Case False
                ii = 0
                For j = LBound(mExam, 2) To UBound(mExam, 2)
                
                k = 0
                If Left(lsID, 1) = "M" And _
                    (Trim(mExam(3, j)) = "0108" Or Trim(mExam(3, j)) = "0288") Then
                Else
                    lsCode = Trim(mExam(3, j))
                    'SaveOrdLog lsCode
                    If Order_OK(lsCode) = 1 Then
                        
                        
                        Select Case Trim(mExam(3, j))
                        Case "0106" 'PP2 => Glucose 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad("0148", 10, " ")
                        Case "0348" 'TIBC 0348 UIBC 오더 주기
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(mExam(3, j)), 10, " ")
                            iTIBC = 1
                        Case "0342" 'Fe 오더 확인
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(mExam(3, j)), 10, " ")
                            iFe = 1
                        Case "0288" 'HIV : 병동은 TLA, 외래는 A0001(Offline)
                            If iWard = 1 Then
                                ii = ii + 1
                                k = 1
                                RSet typOrder.ItemNo(ii) = LPad(Trim(mExam(3, j)), 10, " ")
                            Else
                                If Left(lsID, 1) = "B" Then
                                    ii = ii + 1
                                    k = 1
                                    RSet typOrder.ItemNo(ii) = LPad("A001", 10, " ")
                                End If
                            End If
                        Case "0108" 'HCV : 병동은 TLA, 외래는 A0002(Offline)"
                            If iWard = 1 Then
                                ii = ii + 1
                                k = 1
                                RSet typOrder.ItemNo(ii) = LPad(Trim(mExam(3, j)), 10, " ")
                            Else
                                If Left(lsID, 1) = "B" Then
                                    ii = ii + 1
                                    k = 1
                                    RSet typOrder.ItemNo(ii) = LPad("A002", 10, " ")
                                End If
                            End If
                        Case Else
                            ii = ii + 1
                            k = 1
                            RSet typOrder.ItemNo(ii) = LPad(Trim(mExam(3, j)), 10, " ")
                        End Select
                        
                        If k = 1 Then
                            typOrder.TypeOfOrder(ii) = "01" 'New
                        '이전결과 가저오기
                        'If St2.EOF = False Then
                        '    If IsNull(St2("result")) Then
                                typOrder.NoOfAddInf(ii) = "00"
                                typOrder.TypeOfAddInf(ii) = ""
                                typOrder.AdditionalInf(ii) = ""
                        '    Else
                        '        typOrder.NoOfAddInf(ii) = "01"
                        '        typOrder.TypeOfAddInf(ii) = "01" 'Previous Result
                        '        RSet typOrder.AdditionalInf(ii) = IIf(IsNull(St2("result")), "", St2("result"))
                                
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이상", ">")
                        '        RSet typOrder.AdditionalInf(ii) = Replace(typOrder.AdditionalInf(ii), "이하", "<")
                                
                        '    End If
                        'Else
                        '    typOrder.NoOfAddInf(ii) = "00"
                        '    typOrder.TypeOfAddInf(ii) = ""
                        '    typOrder.AdditionalInf(ii) = ""
                        'End If
                        End If
                    End If
                End If
                Next j
            Case "DELETE"
                '
        End Select
        
        
        
        'TIBC 오더만 난 경우에는 UIBC는 오더 주고 난 뒤 Fe 오더 추가
        If iTIBC = 1 Then
            ii = ii + 1
            
            RSet typOrder.ItemNo(ii) = LPad("A003", 10, " ")
            typOrder.TypeOfOrder(ii) = "01" 'New
            '이전결과 가저오기
            'If St2.EOF = False Then
            '    If IsNull(St2("result")) Then
            typOrder.NoOfAddInf(ii) = "00"
            typOrder.TypeOfAddInf(ii) = ""
            typOrder.AdditionalInf(ii) = ""
            
            typOrder.NumberOfTest = Format(ii, "0000")
            
'            If iFe = 0 Then
'                ii = ii + 1
'                typOrder.NumberOfTest = Format(ii, "0000")
'                RSet typOrder.ItemNo(ii) = LPad(Trim(mExam(3, j)), 10, " ")
'                typOrder.TypeOfOrder(ii) = "01" 'New
'                '이전결과 가저오기
'                'If St2.EOF = False Then
'                '    If IsNull(St2("result")) Then
'                typOrder.NoOfAddInf(ii) = "00"
'                typOrder.TypeOfAddInf(ii) = ""
'                typOrder.AdditionalInf(ii) = ""
'            End If
        End If
        
        typOrder.NumberOfTest = Format(ii, "0000")
        
        '
        Data = ""
        Data = Data & typOrder.FormatTypeCode
        Data = Data & typOrder.SampleID
        Data = Data & typOrder.DateOfReception
        Data = Data & typOrder.PatientID
        Data = Data & typOrder.PatientNameABC
        Data = Data & typOrder.PatientNameReserve
        Data = Data & typOrder.Birthday
        Data = Data & typOrder.Sex
        '
        Data = Data & typOrder.DeleteFlag
        Data = Data & typOrder.WardCode
        Data = Data & typOrder.WardName
        Data = Data & typOrder.OrderDeptCode
        Data = Data & typOrder.OrderDeptName
        Data = Data & typOrder.OrderDrCode
        Data = Data & typOrder.OrderDrName
        '
        Data = Data & typOrder.TypeOfContainer
        Data = Data & typOrder.TypeOfSample
        Data = Data & typOrder.HeightOfSample
        Data = Data & typOrder.DeCapping
        Data = Data & typOrder.Centrifuge
        Data = Data & typOrder.STATFlag
        Data = Data & typOrder.FreeComment
        '
        Data = Data & typOrder.NumberOfTest
        For ii = 1 To typOrder.NumberOfTest
            Data = Data & typOrder.ItemNo(ii)
            Data = Data & typOrder.TypeOfOrder(ii)
            Data = Data & typOrder.NoOfAddInf(ii)
            If typOrder.NoOfAddInf(ii) = "01" Then '01
                Data = Data & typOrder.TypeOfAddInf(ii)
                Data = Data & typOrder.AdditionalInf(ii)
            End If
        Next ii
        
        
        lRow = vasOrder.DataRowCnt + 1
        If lRow > vasOrder.MaxRows Then
            vasOrder.MaxRows = lRow
        End If
        'vasOrder.ActiveRow = lRow
        vasOrder.SetText 1, lRow, Trim(lsID)
        vasOrder.SetText 2, lRow, "B"
        vasOrder.SetText 3, lRow, Trim(mExam(1, LBound(mExam, 2)))
        vasOrder.SetText 4, lRow, Trim(mExam(2, LBound(mExam, 2)))
        vasOrder.SetText 5, lRow, Trim(GetText(vasList, 1, 5))
        
        vasActiveCell vasOrder, lRow, 1
        
        SQL = "Delete from worklist where barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
              "Values ('" & Left(Trim(GetText(vasList, 1, 5)), 8) & "', '" & lsID & "', '" & Trim(GetText(vasOrder, lRow, 3)) & "','" & Trim(GetText(vasOrder, lRow, 4)) & "', 'B','','', 0) "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            SetText vasList, "B", lRow, 2
            'CopyRecord lRow
            'DeleteRow vasList1, lRow1, lRow1
        Else
            SaveQuery SQL
        End If
        
    Else
        '
    End If
    '
    If Data <> "" Then
        'If SendFlg = True Then
            'Text1.Text = Text1.Text & "Send:" & vbCrLf & Replace(MakeString(Data), " ", ".") & vbCrLf
            
            'Data = chrSTX & Data & chrETX & CStr(MakeCS(Data & chrETX))
            Data = chrSTX & Data & chrETX
            MSComm1.Output = Data
            SaveOrdLog "TX : " & Data
            'TXTLOG = txtLog & "TX : " & Data & vbCrLf
            
            txtLog.SelStart = Len(txtLog)
            txtLog.SelLength = 0

        'End If
        '
        SendCursor = 1
        '
'        conDataBase.BeginTrans
'        '
'        SQL = "UPDATE cpusr.temp_clinilog "
'        If Val(typOrder.NumberOfTest) = 0 Then
'            SQL = SQL & "   SET status = 'Y', " '전송시도(삭제)
'        Else
'            SQL = SQL & "   SET status = DECODE(status,'A','B','X','Y',status), " '전송시도
'        End If
'        SQL = SQL & "       senddate = DECODE(status,'A',TO_CHAR(SYSDATE,'YYYYMMDD'),'B',senddate), "
'        SQL = SQL & "       sendtime = DECODE(status,'A',TO_CHAR(SYSDATE,'HH24MISS'),'B',sendtime), "
'        SQL = SQL & "       lastdate = DECODE(status,'A',NULL,'B',TO_CHAR(SYSDATE,'YYYYMMDD')), "
'        SQL = SQL & "       lasttime = DECODE(status,'A',NULL,'B',TO_CHAR(SYSDATE,'HH24MISS')) "
'        SQL = SQL & " WHERE sid = '" & Trim(typOrder.SampleID) & "' "
'        SQL = SQL & "   AND seqn = 0 "
'        SQL = SQL & "   AND status IN ('A','B','X','Y') " '접수, 전송시도
'        conDataBase.Execute SQL
'        '
'        conDataBase.CommitTrans
        '
        Clear_Send
        '
        typOrder.SampleID = Trim(lsID)
        
        SendTheDataCLINILOG_1 = True
    Else
        SendTheDataCLINILOG_1 = False
    End If
    '
    
    Screen.MousePointer = vbDefault
End Function

Private Sub Clear_Send()
    typOrder.FormatTypeCode = ""
    typOrder.SampleID = ""
    typOrder.DateOfReception = ""
    typOrder.PatientID = ""
    typOrder.PatientNameABC = ""
    typOrder.PatientNameReserve = ""
    typOrder.Birthday = ""
    typOrder.Sex = ""
    '
    typOrder.DeleteFlag = ""
    typOrder.WardCode = ""
    typOrder.WardName = ""
    typOrder.OrderDeptCode = ""
    typOrder.OrderDeptName = ""
    typOrder.OrderDrCode = ""
    typOrder.OrderDrName = ""
    '
    typOrder.TypeOfContainer = ""
    typOrder.TypeOfSample = ""
    typOrder.HeightOfSample = ""
    typOrder.DeCapping = ""
    typOrder.Centrifuge = ""
    '
    typOrder.FreeComment = ""
    '
    typOrder.NumberOfTest = ""
    Erase typOrder.ItemNo
    Erase typOrder.TypeOfOrder
    Erase typOrder.NoOfAddInf
    Erase typOrder.TypeOfAddInf
    Erase typOrder.AdditionalInf
    
    ClearSpread vasTemp
    '
    'SendFlg = False
End Sub

Function Order_OK(asCode As String) As Integer
    Order_OK = 0
    
    If Trim(asCode) = "" Then Exit Function
    
    SQL = "SELECT EquipCode, ExamCode, ExamName, Seqno " & CR & _
          "  From EquipExam " & CR & _
          " WHERE Equip = '" & gEquip & "'   AND ExamCode = '" & Trim(asCode) & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(1)) = Trim(asCode) Then
        Order_OK = 1
    End If

End Function

Private Function MakeCS(Source As String) As String
    On Error GoTo MakeCS_
    Dim ii As Long
    Dim SumCS As Long
    Dim AA As Long
    Dim BB As Long
    SumCS = 0
    For ii = 1 To Len(Source)
        SumCS = SumCS Xor Asc(Mid(Source, ii, 1))
    Next ii
    '
    Select Case SumCS
        Case Asc(chrACK)
            SumCS = Asc("A")
        Case Asc(chrNACK)
            SumCS = Asc("A")
    End Select
    '
    MakeCS = Chr(SumCS)
    '
    Exit Function
MakeCS_:
    MsgBox Err.Description
    Exit Function
End Function

Private Sub ReceiveTheDataCLINILOG()
'    'On Error GoTo Exception_
'    Dim ii As Integer
'    Dim Sum As Long
'    Dim nSid As String
'    Dim nResult As String
'    '
'    Dim SQL As String
'    Dim St1 As New ADODB.Recordset
'
'    SaveOrdLog "RX : " & ReceiveData
'    '
'    Text1.Text = Text1.Text & "Receive:" & vbCrLf & Replace(ReceiveData, " ", ".") & vbCrLf
'    '
'    Sum = 1
'    typResult.FormatTypeCode = Mid(ReceiveData, Sum, Len(typResult.FormatTypeCode))
'    Sum = Sum + Len(typResult.FormatTypeCode)
'    typResult.TypeOfSample = Mid(ReceiveData, Sum, Len(typResult.TypeOfSample))
'    Sum = Sum + Len(typResult.TypeOfSample)
'    typResult.SampleID = Mid(ReceiveData, Sum, Len(typResult.SampleID))
'    Sum = Sum + Len(typResult.SampleID)
'    typResult.PatientID = Mid(ReceiveData, Sum, Len(typResult.PatientID))
'    Sum = Sum + Len(typResult.PatientID)
'    typResult.RackID = Mid(ReceiveData, Sum, Len(typResult.RackID))
'    Sum = Sum + Len(typResult.RackID)
'    typResult.RackPosition = Mid(ReceiveData, Sum, Len(typResult.RackPosition))
'    Sum = Sum + Len(typResult.RackPosition)
'    typResult.NoOfAnalyzer = Mid(ReceiveData, Sum, Len(typResult.NoOfAnalyzer))
'    Sum = Sum + Len(typResult.NoOfAnalyzer)
'    '
'    typResult.SampleInfCyle = Mid(ReceiveData, Sum, Len(typResult.SampleInfCyle))
'    Sum = Sum + Len(typResult.SampleInfCyle)
'    typResult.SampleInfHb = Mid(ReceiveData, Sum, Len(typResult.SampleInfHb))
'    Sum = Sum + Len(typResult.SampleInfHb)
'    typResult.SampleInfBil = Mid(ReceiveData, Sum, Len(typResult.SampleInfBil))
'    Sum = Sum + Len(typResult.SampleInfBil)
'    '
'    typResult.NoOfItems = Mid(ReceiveData, Sum, Len(typResult.NoOfItems))
'    Sum = Sum + Len(typResult.NoOfItems)
'        For ii = 1 To Val(typResult.NoOfItems)
'            typResult.ItemNo(ii) = Mid(ReceiveData, Sum, Len(typResult.ItemNo(ii)))
'            Sum = Sum + Len(typResult.ItemNo(ii))
'            typResult.Result(ii) = Mid(ReceiveData, Sum, Len(typResult.Result(ii)))
'            Sum = Sum + Len(typResult.Result(ii))
'            typResult.Comment(ii) = Mid(ReceiveData, Sum, Len(typResult.Comment(ii)))
'            Sum = Sum + Len(typResult.Comment(ii))
'            typResult.DilutionRatio(ii) = Mid(ReceiveData, Sum, Len(typResult.DilutionRatio(ii)))
'            Sum = Sum + Len(typResult.DilutionRatio(ii))
'            typResult.ConfirmFlag(ii) = Mid(ReceiveData, Sum, Len(typResult.ConfirmFlag(ii)))
'            Sum = Sum + Len(typResult.ConfirmFlag(ii))
'        Next ii
'    '
'    typResult.LengthOfFreeComment = Mid(ReceiveData, Sum, Len(typResult.LengthOfFreeComment))
'    Sum = Sum + Len(typResult.LengthOfFreeComment)
'    typResult.FreeComment = Mid(ReceiveData, Sum, Len(typResult.FreeComment))
'    Sum = Sum + Len(typResult.FreeComment)
'    '
'    conDataBase.BeginTrans
'    '
'    For ii = 1 To Val(typResult.NoOfItems)
'        If typResult.ItemNo(ii) <> "" Then
'            If typResult.TypeOfSample = "01" Then '00:normal, 01:control
'                If nSid = "" Then
'                    nSid = GetLastSid(Format(Now, "YYYYMMDD"))
'                End If
'            End If
'            If Trim(typResult.SampleID) = "" Then
'                typResult.SampleID = "0000000000"
'            End If
'            nResult = Trim(typResult.Result(ii))
'            If Left(nResult, 1) = "." Then
'                nResult = "0" & nResult
'            End If
'            If Right(nResult, 1) = "." Then
'                nResult = Left(nResult, Len(nResult) - 1)
'            End If
'            '
'            SQL = "SELECT "
'            SQL = SQL & "       ROWID "
'            SQL = SQL & "  FROM cpusr.int_clinilog "
'            SQL = SQL & " WHERE testdate = TO_CHAR(SYSDATE, 'YYYYMMDD') "
'            SQL = SQL & "   AND testtime = TO_CHAR(SYSDATE, 'HH24MISS') "
'            SQL = SQL & "   AND sampleid = '" & Trim(typResult.SampleID) & "' "
'            SQL = SQL & "   AND testid = '" & Trim(typResult.ItemNo(ii)) & "' "
'            St1.Open SQL, conDataBase, adOpenForwardOnly, adLockReadOnly
'            If St1.EOF Then
'                SQL = "INSERT INTO cpusr.int_clinilog ( "
'                SQL = SQL & " testdate, testtime, "
'                SQL = SQL & " sampleid, patientid, testid, "
'                SQL = SQL & " rackid, rackposition, analyzerno, "
'                SQL = SQL & " sampleinfcyle, sampleinfhb, sampleinfbil, "
'                SQL = SQL & " itemsno, "
'                SQL = SQL & " sid, seqn, testcd, testno, "
'                SQL = SQL & " result, comments, dilutionratio, confirmflag, "
'                SQL = SQL & " freecomment "
'                SQL = SQL & " ) VALUES ( "
'                SQL = SQL & " TO_CHAR(SYSDATE, 'YYYYMMDD'), "
'                SQL = SQL & " TO_CHAR(SYSDATE, 'HH24MISS'), "
'                If typResult.TypeOfSample = "01" Then '00:normal, 01:control
'                    typResult.PatientID = typResult.SampleID
'                    typResult.SampleID = nSid
'                    '
'                    SQL = SQL & " '" & Trim(typResult.SampleID) & "', "
'                    SQL = SQL & " '" & Trim(typResult.PatientID) & "', "
'                Else
'                    SQL = SQL & " '" & Trim(typResult.SampleID) & "', "
'                    SQL = SQL & " '" & Trim(typResult.PatientID) & "', "
'                End If
'                SQL = SQL & " '" & Trim(typResult.ItemNo(ii)) & "', "
'                SQL = SQL & " '" & Trim(typResult.RackID) & "', "
'                SQL = SQL & " '" & Trim(typResult.RackPosition) & "', "
'                SQL = SQL & " '" & Trim(typResult.NoOfAnalyzer) & "', "
'                SQL = SQL & " '" & Trim(typResult.SampleInfCyle) & "', "
'                SQL = SQL & " '" & Trim(typResult.SampleInfHb) & "', "
'                SQL = SQL & " '" & Trim(typResult.SampleInfBil) & "', "
'                SQL = SQL & " '" & Trim(typResult.NoOfItems) & "', "
'                If typResult.TypeOfSample = "01" Then '00:normal, 01:control
'                    SQL = SQL & " '" & nSid & "', "
'                    SQL = SQL & " 0, "
'                Else
'                    If Trim(typResult.SampleID) <> "" Then
'                        If Val(typResult.SampleID) = 0 Then
'                            SQL = SQL & " NULL, "
'                        Else
'                            SQL = SQL & " '" & Trim(typResult.SampleID) & "', "
'                        End If
'                        SQL = SQL & " 0, "
'                    Else
'                        SQL = SQL & " NULL, "
'                        SQL = SQL & " NULL, "
'                    End If
'                End If
'                SQL = SQL & " NULL, "
'                SQL = SQL & " NULL, "
'                SQL = SQL & " '" & Trim(typResult.Result(ii)) & "', "
'                SQL = SQL & " '" & Trim(typResult.Comment(ii)) & "', "
'                SQL = SQL & " '" & Trim(typResult.DilutionRatio(ii)) & "', "
'                SQL = SQL & " '" & Trim(typResult.ConfirmFlag(ii)) & "', "
'                SQL = SQL & " '" & Trim(typResult.FreeComment) & "' ) "
'                conDataBase.Execute SQL
'            Else
'                SQL = "UPDATE cpusr.int_clinilog "
'                SQL = SQL & "   SET result = '" & Trim(typResult.Result(ii)) & "', "
'                SQL = SQL & "       comments = '" & Trim(typResult.Comment(ii)) & "', "
'                SQL = SQL & "       dilutionratio = '" & Trim(typResult.DilutionRatio(ii)) & "', "
'                SQL = SQL & "       confirmflag = '" & Trim(typResult.ConfirmFlag(ii)) & "' "
'                SQL = SQL & " WHERE ROWID = '" & St1("ROWID") & "' "
'                conDataBase.Execute SQL
'            End If
'            St1.Close
'            Set St1 = Nothing
'        End If
'    Next ii
'    '
'    If typResult.TypeOfSample = "00" Then '00:normal, 01:control
'        SQL = "UPDATE cpusr.temp_clinilog "
'        SQL = SQL & "   SET status = 'D', " '결과
'        SQL = SQL & "       resultdate = TO_CHAR(SYSDATE,'YYYYMMDD'), "
'        SQL = SQL & "       resulttime = TO_CHAR(SYSDATE,'HH24MISS') "
'        SQL = SQL & " WHERE sid = '" & Trim(typResult.SampleID) & "' "
'        SQL = SQL & "   AND seqn = 0 "
'        SQL = SQL & "   AND status = 'C' " '전송성공
'        conDataBase.Execute SQL
'    End If
'    '
'    conDataBase.CommitTrans
'    '
'    'SendTheDataCLINILOG (Index Mod 2) + 1, VITROS_R.SampleID, True
'    '
'    Clear_Receive True
'    '
'    Exit Sub
'Exception_:
'    MsgBox Err.Description
'    Exit Sub
End Sub
