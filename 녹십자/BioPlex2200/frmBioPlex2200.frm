VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBioPlex2200 
   Caption         =   "BioPlex2200"
   ClientHeight    =   11955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   21840
   Icon            =   "frmBioPlex2200.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11955
   ScaleWidth      =   21840
   WindowState     =   2  '최대화
   Begin FPSpread.vaSpread tblErrors 
      Height          =   1215
      Left            =   8940
      TabIndex        =   43
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
      SpreadDesigner  =   "frmBioPlex2200.frx":1272
   End
   Begin VB.Frame Frame6 
      Height          =   1320
      Left            =   60
      TabIndex        =   39
      Top             =   10440
      Width           =   8715
      Begin VB.CommandButton cmdLog 
         Caption         =   "로그보기"
         Height          =   855
         Left            =   7020
         TabIndex        =   65
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
         TabIndex        =   42
         Top             =   300
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  '평면
         Caption         =   "결과등록(1차)"
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
         Left            =   90
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   41
         Top             =   300
         Width           =   1965
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  '평면
         Caption         =   "결과등록(2차)"
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
         Left            =   2130
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   40
         Top             =   300
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1470
      Left            =   5640
      TabIndex        =   36
      Top             =   60
      Width           =   3645
      Begin VB.CommandButton Command6 
         Caption         =   "Comm Test"
         Height          =   555
         Left            =   1200
         TabIndex        =   66
         Top             =   180
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  '평면
         Caption         =   "결과가져오기"
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
         Left            =   1800
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   38
         Top             =   300
         Width           =   1665
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
         Height          =   885
         Left            =   90
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   37
         Top             =   300
         Width           =   1665
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1470
      Left            =   60
      TabIndex        =   27
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1020
         Width           =   3675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "결과수신자료"
         Height          =   285
         Left            =   3690
         TabIndex        =   33
         Top             =   690
         Width           =   1395
      End
      Begin VB.ComboBox cboWhere 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmBioPlex2200.frx":16FE
         Left            =   1500
         List            =   "frmBioPlex2200.frx":170B
         Style           =   2  '드롭다운 목록
         TabIndex        =   32
         Top             =   630
         Width           =   1740
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   8  '영문
         Left            =   1470
         TabIndex        =   31
         Text            =   "2014년 01월 07일"
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "▶ 작업일자 :"
         Height          =   180
         Left            =   270
         TabIndex        =   30
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "▶ 조회조건 :"
         Height          =   180
         Left            =   270
         TabIndex        =   29
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "▷ 의뢰번호 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   28
         Top             =   1095
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdMode 
      Caption         =   "통합모드"
      Height          =   585
      Left            =   20190
      TabIndex        =   26
      Tag             =   "1"
      Top             =   840
      Width           =   1485
   End
   Begin VB.Frame Frame4 
      Height          =   1470
      Left            =   9330
      TabIndex        =   9
      Top             =   60
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   10
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
         Left            =   3930
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   14
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
         Caption         =   "수술실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   6795
         TabIndex        =   15
         Top             =   525
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
         Caption         =   "Blood"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   1245
         TabIndex        =   16
         Top             =   885
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
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   17
         Top             =   885
         Width           =   4530
         _ExtentX        =   7990
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
         Caption         =   " 원장님 친척"
         Appearance      =   0
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "검 체 명 :"
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
         Index           =   4
         Left            =   5760
         TabIndex        =   25
         Top             =   600
         Width           =   900
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
         TabIndex        =   24
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "기타 : "
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
         Index           =   2
         Left            =   3105
         TabIndex        =   23
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "거래처 :"
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
         Index           =   1
         Left            =   3105
         TabIndex        =   22
         Top             =   600
         Width           =   720
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
         TabIndex        =   21
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
         Left            =   150
         TabIndex        =   20
         Top             =   975
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "hidden frame"
      Height          =   1185
      Left            =   11820
      TabIndex        =   0
      Top             =   5610
      Visible         =   0   'False
      Width           =   2355
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
      TabIndex        =   5
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
         Picture         =   "frmBioPlex2200.frx":172B
         Top             =   750
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   1860
         Picture         =   "frmBioPlex2200.frx":1CB5
         Top             =   750
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   3150
         Picture         =   "frmBioPlex2200.frx":223F
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
         TabIndex        =   8
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   1245
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8805
      Left            =   60
      TabIndex        =   1
      Top             =   1590
      Width           =   21645
      _ExtentX        =   38179
      _ExtentY        =   15531
      _Version        =   393216
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "검사내역"
      TabPicture(0)   =   "frmBioPlex2200.frx":27C9
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "spdIntegration(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "spdSeparationOrder(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "spdSeparationResult(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "정상결과"
      TabPicture(1)   =   "frmBioPlex2200.frx":27E5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spdSeparationResult(1)"
      Tab(1).Control(1)=   "spdSeparationOrder(1)"
      Tab(1).Control(2)=   "spdIntegration(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "비정상결과"
      TabPicture(2)   =   "frmBioPlex2200.frx":2801
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "spdSeparationResult(2)"
      Tab(2).Control(1)=   "spdSeparationOrder(2)"
      Tab(2).Control(2)=   "spdIntegration(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "L/H"
      TabPicture(3)   =   "frmBioPlex2200.frx":281D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "spdSeparationResult(3)"
      Tab(3).Control(1)=   "spdSeparationOrder(3)"
      Tab(3).Control(2)=   "spdIntegration(3)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "D/P/C"
      TabPicture(4)   =   "frmBioPlex2200.frx":2839
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "spdSeparationResult(4)"
      Tab(4).Control(1)=   "spdSeparationOrder(4)"
      Tab(4).Control(2)=   "spdIntegration(4)"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Hold"
      TabPicture(5)   =   "frmBioPlex2200.frx":2855
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "spdSeparationResult(5)"
      Tab(5).Control(1)=   "spdSeparationOrder(5)"
      Tab(5).Control(2)=   "spdIntegration(5)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Flag"
      TabPicture(6)   =   "frmBioPlex2200.frx":2871
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "spdIntegration(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "spdSeparationOrder(6)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "spdSeparationResult(6)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Others"
      TabPicture(7)   =   "frmBioPlex2200.frx":288D
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "spdSeparationResult(7)"
      Tab(7).Control(1)=   "spdSeparationOrder(7)"
      Tab(7).Control(2)=   "spdIntegration(7)"
      Tab(7).ControlCount=   3
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   0
         Left            =   -66330
         TabIndex        =   3
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":28A9
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   0
         Left            =   -74820
         TabIndex        =   2
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":333E
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   1
         Left            =   -66330
         TabIndex        =   44
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":7433
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   2
         Left            =   -66330
         TabIndex        =   45
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":7EC8
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   3
         Left            =   -66330
         TabIndex        =   46
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":895D
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   4
         Left            =   -66330
         TabIndex        =   47
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":93F2
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   5
         Left            =   -66330
         TabIndex        =   48
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":9E87
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   6
         Left            =   8670
         TabIndex        =   49
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":A91C
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationResult 
         Height          =   8070
         Index           =   7
         Left            =   -66330
         TabIndex        =   50
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":B3B1
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   1
         Left            =   -74820
         TabIndex        =   51
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":BE46
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   2
         Left            =   -74820
         TabIndex        =   52
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":FF3B
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   3
         Left            =   -74820
         TabIndex        =   53
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":14030
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   4
         Left            =   -74820
         TabIndex        =   54
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":18125
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   5
         Left            =   -74820
         TabIndex        =   55
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":1C21A
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   6
         Left            =   180
         TabIndex        =   56
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":2030F
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdSeparationOrder 
         Height          =   8085
         Index           =   7
         Left            =   -74820
         TabIndex        =   57
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":24404
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   1
         Left            =   -74820
         TabIndex        =   58
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":284F9
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   2
         Left            =   -74820
         TabIndex        =   59
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":2C9A8
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   3
         Left            =   -74820
         TabIndex        =   60
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":30E57
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   4
         Left            =   -74820
         TabIndex        =   61
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":35306
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   5
         Left            =   -74820
         TabIndex        =   62
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":397B5
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   6
         Left            =   180
         TabIndex        =   63
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":3DC64
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   7
         Left            =   -74820
         TabIndex        =   64
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":42113
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdIntegration 
         Height          =   8085
         Index           =   0
         Left            =   -74820
         TabIndex        =   4
         Top             =   540
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
         SpreadDesigner  =   "frmBioPlex2200.frx":465C2
         UserResize      =   2
      End
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


Private Sub Command6_Click()
    
    Call comEqp_OnComm

End Sub

Public Sub Form_Load()
    
    '-- DB 접속
    
    '-- 로그인 사용자
    
    '-- 컨트롤초기화
    Call CtlInitializing
    
    '-- 장비통신정보 읽어오기
    
    '-- 장비검사정보 읽어오기(검사항목 리스트업)
    Call getTestNms(mEqpCd)
    
    '-- 포트 열기
    
    
    
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


