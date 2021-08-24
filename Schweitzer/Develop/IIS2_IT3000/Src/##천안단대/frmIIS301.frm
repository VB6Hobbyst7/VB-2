VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmIIS301 
   BackColor       =   &H00DBE6E6&
   Caption         =   "결과등록"
   ClientHeight    =   9180
   ClientLeft      =   -4350
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   990
      Left            =   6360
      TabIndex        =   26
      Top             =   -60
      Width           =   8835
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방비고"
         Height          =   495
         Left            =   5220
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   345
         Width           =   1755
      End
      Begin VB.CommandButton cmdRelTests 
         BackColor       =   &H00DBE6E6&
         Caption         =   "관련검사 결과"
         Height          =   495
         Left            =   6975
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   345
         Width           =   1755
      End
   End
   Begin VB.CommandButton txtCumul 
      BackColor       =   &H00DBE6E6&
      Caption         =   "누적결과(&S)"
      Height          =   495
      Left            =   10320
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton txtReport 
      BackColor       =   &H00DBE6E6&
      Caption         =   "결과보고(&R)"
      Height          =   495
      Left            =   11535
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   12750
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   8567
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   990
      Left            =   75
      TabIndex        =   11
      Top             =   -60
      Width           =   6255
      Begin VB.OptionButton optCondition 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수시간별 조회"
         Height          =   240
         Index           =   1
         Left            =   2010
         TabIndex        =   3
         Top             =   660
         Width           =   1620
      End
      Begin VB.OptionButton optCondition 
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급검체 우선조회"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   1830
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조 회(&Q)"
         Height          =   495
         Left            =   4875
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체조회"
         Height          =   255
         Left            =   3750
         TabIndex        =   4
         Top             =   645
         Width           =   1080
      End
      Begin VB.TextBox txtEqpCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1185
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   195
         Width           =   825
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   345
         Index           =   0
         Left            =   2010
         Picture         =   "frmIIS301.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   180
         Width           =   405
      End
      Begin MedControls1.LisLabel lblEqpNm 
         Height          =   345
         Left            =   2445
         TabIndex        =   14
         Top             =   180
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 장비코드"
         Height          =   180
         Left            =   165
         TabIndex        =   12
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   13965
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   8567
      Width           =   1215
   End
   Begin FPSpread.vaSpread tblSpcs 
      Height          =   7545
      Left            =   75
      TabIndex        =   5
      Top             =   945
      Width           =   3675
      _Version        =   393216
      _ExtentX        =   6482
      _ExtentY        =   13309
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
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      MaxCols         =   3
      MaxRows         =   1
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIIS301.frx":0E42
   End
   Begin FPSpread.vaSpread tblPtInfo 
      Height          =   720
      Left            =   3780
      TabIndex        =   15
      Top             =   945
      Width           =   11400
      _Version        =   393216
      _ExtentX        =   20108
      _ExtentY        =   1270
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
      GrayAreaBackColor=   16777215
      MaxCols         =   9
      MaxRows         =   1
      OperationMode   =   2
      ScrollBars      =   0
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIIS301.frx":11F7
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   5265
      Left            =   3780
      TabIndex        =   6
      Top             =   1695
      Width           =   11400
      _Version        =   393216
      _ExtentX        =   20108
      _ExtentY        =   9287
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
      GrayAreaBackColor=   16777215
      MaxCols         =   9
      MaxRows         =   17
      ScrollBars      =   2
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIIS301.frx":17DC
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "FootNote && Remark"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   3780
      TabIndex        =   16
      Top             =   7035
      Width           =   5685
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   345
         Index           =   2
         Left            =   5205
         Picture         =   "frmIIS301.frx":1F37
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   1035
         Width           =   405
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   345
         Index           =   1
         Left            =   5205
         Picture         =   "frmIIS301.frx":2D79
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   570
         Width           =   405
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00F7FFF7&
         Height          =   375
         Left            =   75
         TabIndex        =   21
         Top             =   1020
         Width           =   5115
      End
      Begin VB.TextBox txtFootNote 
         BackColor       =   &H00F7FFF7&
         Height          =   675
         Left            =   75
         TabIndex        =   20
         Top             =   255
         Width           =   5130
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Text Result"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   9480
      TabIndex        =   17
      Top             =   7035
      Width           =   5715
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   345
         Index           =   3
         Left            =   5205
         Picture         =   "frmIIS301.frx":3BBB
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   1035
         Width           =   405
      End
      Begin VB.TextBox txtTextResult 
         BackColor       =   &H00F7FFF7&
         Height          =   1140
         Left            =   75
         TabIndex        =   24
         Top             =   255
         Width           =   5130
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   405
      Left            =   75
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8550
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   714
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "■ 전송된 개수"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblCount 
      Height          =   405
      Left            =   2580
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8550
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   714
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "100"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmIIS301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS301.frm (우리LIS랑 조인할때 사용)
'   작성자  :
'   내  용  : 결과등록폼
'   작성일  : 2005-07-19
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Load()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIIS301 = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
