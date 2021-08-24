VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmBBS201 
   BackColor       =   &H00DBE6E6&
   Caption         =   "결과등록"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14730
   Icon            =   "frmBBS201_cmc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   14730
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   480
      Left            =   12000
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8520
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrintTAG 
      BackColor       =   &H00F4F0F2&
      Caption         =   "TAG출력"
      Height          =   480
      Left            =   9360
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8520
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   10680
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8520
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   13320
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8520
      Width           =   1245
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1680
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   556
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "검사 결과"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   6435
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   14475
      Begin VB.TextBox txtOrderRmk 
         BackColor       =   &H00CFDCDE&
         Height          =   975
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5220
         Width           =   6975
      End
      Begin VB.TextBox txtResultRmk 
         Height          =   975
         Left            =   7320
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   5220
         Width           =   6975
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1515
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   4230
         Left            =   180
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "10114"
         Top             =   660
         Width           =   14130
         _Version        =   196608
         _ExtentX        =   24924
         _ExtentY        =   7461
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15003117
         GridColor       =   14737632
         MaxCols         =   24
         OperationMode   =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS201_cmc.frx":076A
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   500
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   7740
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   10320
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   13500
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과 Remark"
         Height          =   180
         Left            =   7320
         TabIndex        =   37
         Top             =   4980
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방 Remark"
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Top             =   4980
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체채취후 경과시간"
         Height          =   180
         Left            =   11760
         TabIndex        =   31
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체보관장소"
         Height          =   180
         Left            =   9180
         TabIndex        =   29
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         Height          =   180
         Left            =   6960
         TabIndex        =   27
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액번호"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   300
         Width           =   720
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   556
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자 정보"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   14475
      Begin VB.TextBox txtDay 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   330
         Left            =   5160
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboReqDt 
         Height          =   300
         Left            =   3240
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPtID 
         Height          =   330
         Left            =   780
         TabIndex        =   0
         Top             =   240
         Width           =   1515
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   2820
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   660
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   4560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   660
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   6180
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   660
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   315
         Left            =   8220
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   660
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   315
         Left            =   780
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   660
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "일 전까지 조회"
         Height          =   180
         Left            =   5760
         TabIndex        =   32
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성명"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액형"
         Height          =   180
         Left            =   7620
         TabIndex        =   18
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "병동"
         Height          =   180
         Left            =   5760
         TabIndex        =   16
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "진료과"
         Height          =   180
         Left            =   3960
         TabIndex        =   14
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성/나이"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblDtCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4680
         TabIndex        =   11
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "예정일자"
         Height          =   180
         Left            =   2460
         TabIndex        =   10
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자ID"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmBBS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

