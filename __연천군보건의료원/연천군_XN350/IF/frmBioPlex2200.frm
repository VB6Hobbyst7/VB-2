VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBioPlex2200 
   Caption         =   "BioPlex2200"
   ClientHeight    =   11955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   21840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11955
   ScaleWidth      =   21840
   WindowState     =   2  '최대화
   Begin VB.Frame Frame6 
      Height          =   1320
      Left            =   60
      TabIndex        =   26
      Top             =   10440
      Width           =   8715
      Begin VB.CommandButton cmdLog 
         Caption         =   "로그보기"
         Height          =   855
         Left            =   7020
         TabIndex        =   29
         Tag             =   "1"
         Top             =   300
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  '평면
         Caption         =   "변경사항 저장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   4410
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   28
         Top             =   300
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  '평면
         Caption         =   "결과등록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   510
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   27
         Top             =   300
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 워크리스트 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1470
      Left            =   5610
      TabIndex        =   24
      Top             =   60
      Width           =   3855
      Begin VB.TextBox txtStartNum 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   990
         TabIndex        =   43
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtStopNum 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         TabIndex        =   42
         Top             =   660
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  '평면
         Caption         =   "WORKLIST 조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2460
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   25
         Top             =   630
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   315
         Left            =   2460
         TabIndex        =   44
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   315
         Left            =   990
         TabIndex        =   45
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   40248
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "작업일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   49
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2310
         TabIndex        =   48
         Top             =   330
         Width           =   105
      End
      Begin VB.Label Label13 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "W/N"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   47
         Top             =   750
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   46
         Top             =   750
         Width           =   165
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1470
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   5535
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  '평면
         Caption         =   "검사LIST 조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3660
         TabIndex        =   23
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   22
         Top             =   1020
         Width           =   3675
      End
      Begin VB.ComboBox cboWhere 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   630
         Width           =   1890
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1500
         TabIndex        =   50
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364736
         CurrentDate     =   40248
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "▶ 작업일자 :"
         Height          =   180
         Left            =   270
         TabIndex        =   20
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "▶ 조회조건 :"
         Height          =   180
         Left            =   270
         TabIndex        =   19
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "▷ 의뢰번호 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   18
         Top             =   1095
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdMode 
      Caption         =   "통합모드"
      Height          =   585
      Left            =   20190
      TabIndex        =   16
      Tag             =   "1"
      Top             =   840
      Width           =   1485
   End
   Begin VB.Frame Frame4 
      Height          =   1470
      Left            =   9330
      TabIndex        =   5
      Top             =   60
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   6
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "00000001"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStatFg 
         Height          =   315
         Left            =   6795
         TabIndex        =   8
         Top             =   165
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "QC"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   315
         Left            =   1245
         TabIndex        =   9
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "오세원"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "남자 / 18"
         Appearance      =   0
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "검사구분 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   5760
         TabIndex        =   15
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "의뢰일자 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   3105
         TabIndex        =   14
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "성별/나이 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3030
         TabIndex        =   13
         Top             =   615
         Width           =   990
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         Caption         =   "환자이름 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   12
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         Caption         =   "의뢰번호 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "hidden frame"
      Height          =   2955
      Left            =   11820
      TabIndex        =   0
      Top             =   5610
      Visible         =   0   'False
      Width           =   6525
      Begin VB.CommandButton cmdCommTest 
         Caption         =   "Comm Test"
         Height          =   525
         Left            =   1500
         TabIndex        =   40
         Top             =   300
         Visible         =   0   'False
         Width           =   1245
      End
      Begin MDIControls.MDIActiveX MDIActiveX 
         Left            =   780
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   794
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   180
         Top             =   330
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '없음
      Height          =   615
      Left            =   18000
      TabIndex        =   1
      Top             =   60
      Width           =   3675
      Begin VB.Shape Shape4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   10
         FillColor       =   &H000000FF&
         Height          =   105
         Left            =   3210
         Shape           =   3  '원형
         Top             =   270
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   10
         FillColor       =   &H000000FF&
         Height          =   105
         Left            =   1950
         Shape           =   3  '원형
         Top             =   270
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   10
         FillColor       =   &H0000FF00&
         Height          =   105
         Left            =   750
         Shape           =   3  '원형
         Top             =   270
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00C0FFC0&
         Height          =   465
         Left            =   30
         Shape           =   4  '둥근 사각형
         Top             =   90
         Width           =   3585
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   690
         Top             =   750
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   1860
         Top             =   750
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   3150
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   1245
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive : "
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8805
      Left            =   30
      TabIndex        =   30
      Top             =   1590
      Width           =   21645
      _ExtentX        =   38179
      _ExtentY        =   15531
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "검사내역"
      TabPicture(0)   =   "frmBioPlex2200.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "spdIntegration(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "spdSeparationOrder(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "spdSeparationResult(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "정상결과"
      TabPicture(1)   =   "frmBioPlex2200.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spdIntegration(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "spdSeparationOrder(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "spdSeparationResult(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "비정상결과"
      TabPicture(2)   =   "frmBioPlex2200.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "spdIntegration(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "spdSeparationOrder(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "spdSeparationResult(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   0
         Left            =   8670
         TabIndex        =   31
         Top             =   480
         Width           =   12570
         _Version        =   393216
         _ExtentX        =   22172
         _ExtentY        =   14235
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmBioPlex2200.frx":0054
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Top             =   480
         Width           =   8505
         _Version        =   393216
         _ExtentX        =   15002
         _ExtentY        =   14261
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmBioPlex2200.frx":0B0A
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   1
         Left            =   -66330
         TabIndex        =   33
         Top             =   480
         Width           =   12570
         _Version        =   393216
         _ExtentX        =   22172
         _ExtentY        =   14235
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmBioPlex2200.frx":4C20
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   2
         Left            =   -66330
         TabIndex        =   34
         Top             =   480
         Width           =   12570
         _Version        =   393216
         _ExtentX        =   22172
         _ExtentY        =   14235
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmBioPlex2200.frx":56D6
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   1
         Left            =   -74820
         TabIndex        =   35
         Top             =   480
         Width           =   8505
         _Version        =   393216
         _ExtentX        =   15002
         _ExtentY        =   14261
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmBioPlex2200.frx":618C
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   2
         Left            =   -74820
         TabIndex        =   36
         Top             =   480
         Width           =   8505
         _Version        =   393216
         _ExtentX        =   15002
         _ExtentY        =   14261
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmBioPlex2200.frx":A2A2
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   1
         Left            =   -74820
         TabIndex        =   37
         Top             =   480
         Width           =   21045
         _Version        =   393216
         _ExtentX        =   37121
         _ExtentY        =   14261
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   40
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmBioPlex2200.frx":E3B8
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   2
         Left            =   -74820
         TabIndex        =   38
         Top             =   480
         Width           =   21045
         _Version        =   393216
         _ExtentX        =   37121
         _ExtentY        =   14261
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   40
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmBioPlex2200.frx":12888
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   0
         Left            =   180
         TabIndex        =   39
         Top             =   480
         Width           =   21045
         _Version        =   393216
         _ExtentX        =   37121
         _ExtentY        =   14261
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   40
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmBioPlex2200.frx":16D58
         UserResize      =   2
      End
   End
   Begin FPSpread.vaSpread tblErrors 
      Height          =   1215
      Left            =   9180
      TabIndex        =   41
      Top             =   10500
      Width           =   12495
      _Version        =   393216
      _ExtentX        =   22040
      _ExtentY        =   2143
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmBioPlex2200.frx":1B228
   End
End
Attribute VB_Name = "frmBioPlex2200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmBioPlex2200.frm
'   작성자  : 오세원
'   내  용  : BioPlex2200 장비폼
'   작성일  : 2014-01-07
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqpCd As String
Private mEqpKey As String

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Private Sub cmdCommTest_Click()
    
    Call comEqp_OnComm
    
End Sub

Private Sub cmdLog_Click()
    
    frmLog.Show vbModal

End Sub

Private Sub cmdMode_Click()
    
    Dim intCnt As Integer
    
'    If Index = 0 Then
'        vasID(0).Visible = True
'        vasID(1).Visible = False
'
'        vasID(0).ZOrder 0
'    Else
'        vasID(0).Visible = False
'        vasID(1).Visible = True
'
'        vasID(1).ZOrder 0
'    End If

'    '-- 분리모드 클릭
'    If cmdMode.Tag = 0 Then
'        vasID(0).Visible = True
'        vasID(1).Visible = False
'
'        vasID(0).ZOrder 0
'        cmdMode.Caption = "통합모드"
'        cmdMode.Tag = 1
'
'    '-- 통합모드 클릭
'    Else
'        vasID(0).Visible = False
'        vasID(1).Visible = True
'
'        vasID(1).ZOrder 0
'        cmdMode.Caption = "분리모드"
'        cmdMode.Tag = 0
'    End If

    '-- 분리모드 클릭
    If cmdMode.Tag = 0 Then
        For intCnt = 0 To 7
            spdIntegration(intCnt).Visible = False
            spdSeparationOrder(intCnt).Visible = True
            spdSeparationResult(intCnt).Visible = True
        Next
        cmdMode.Caption = "통합모드"
        cmdMode.Tag = 1
    
    '-- 통합모드 클릭
    Else
        For intCnt = 0 To 7
            spdIntegration(intCnt).Visible = True
            spdSeparationOrder(intCnt).Visible = False
            spdSeparationResult(intCnt).Visible = False
        Next
        cmdMode.Caption = "분리모드"
        cmdMode.Tag = 0
    End If
    
    
End Sub

Private Sub comEqp_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long

    Select Case comEqp.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long

'            '-- 버퍼에 Write
'            Buffer = comEqp.Input
'
'            '-- 로그저장(원시데이터)
'            Call WriteLog(Buffer, ccEqp)
'
'            lngBufLen = Len(Buffer)
'            For i = 1 To lngBufLen
'                BufChar = Mid$(Buffer, i, 1)
'
'                Select Case mIntLib.Phase
'                    Case 1      '## STX 대기
'                        Select Case BufChar
'                            Case STX
'                                Call mIntLib.ClearBuffer
'                                mIntLib.Phase = 2
'                        End Select
'                    Case 2      '## ETX 대기
'                        Select Case BufChar
'                            Case ETX
'                                Call EditRcvData
'                                mIntLib.Phase = 1
'                                MSComm.Output = ACK
'                            Case Else
'                                Call mIntLib.AddBuffer(BufChar)
'                        End Select
'                End Select
'            Next i

        Case comEvSend
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

    If Len(EVMsg$) Then
'        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
'        StatusBar.Panels(2).Text = ERMsg$
    End If
    
End Sub



Public Sub Form_Load()

    
    
    
End Sub

Private Sub CtlInitializing()
    Dim intCnt As Integer
'    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
'    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
'    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    cboWhere.ListIndex = 0

'    vasID(0).Visible = True
'    vasID(1).Visible = False
    
'    vasID(0).ZOrder 0
    
    
    For intCnt = 0 To 7
        spdIntegration(intCnt).Visible = False
        spdSeparationOrder(intCnt).Visible = True
        spdSeparationResult(intCnt).Visible = True
    Next
    
End Sub

Private Sub getTestNms(ByVal strEqpCd As String)
    Dim intCnt As Integer

    
    For intCnt = 0 To 7
        spdIntegration(intCnt).Visible = False
        spdSeparationOrder(intCnt).Visible = True
        spdSeparationResult(intCnt).Visible = True
        With spdIntegration(intCnt)
            Call .SetText(14, 0, "WBC")
            Call .SetText(15, 0, "RBC")
            Call .SetText(16, 0, "HGB")
            Call .SetText(17, 0, "HCV")
            Call .SetText(18, 0, "MCV")
            Call .SetText(19, 0, "MCH")
            Call .SetText(20, 0, "MCHC")
            Call .SetText(21, 0, "PDW")
            Call .SetText(22, 0, "RDW")
            Call .SetText(23, 0, "MPV")
            Call .SetText(24, 0, "NEUT")
            Call .SetText(25, 0, "BASO")
            Call .SetText(26, 0, "EO")
        End With
    Next
End Sub


