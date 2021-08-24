VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS108 
   BackColor       =   &H00DBE6E6&
   Caption         =   "접수취소"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
   Icon            =   "frmBBS108.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   14625
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "접수취소(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   3495
      Left            =   75
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3555
      Width           =   10695
      _Version        =   196608
      _ExtentX        =   18865
      _ExtentY        =   6165
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   11
      MaxRows         =   12
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS108.frx":076A
      TextTip         =   4
   End
   Begin FPSpread.vaSpread tblDStatus 
      Height          =   1020
      Left            =   75
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7410
      Width           =   10695
      _Version        =   196608
      _ExtentX        =   18865
      _ExtentY        =   1799
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   6
      MaxRows         =   1
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS108.frx":0D0B
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   75
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자 기본 정보"
      Appearance      =   0
   End
   Begin VB.Frame fraPtInfo 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   75
      TabIndex        =   7
      Tag             =   "104"
      Top             =   270
      Width           =   14385
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   3660
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   6345
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "성명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   9015
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   3660
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   6345
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   11685
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   9015
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "병실"
         Appearance      =   0
      End
      Begin VB.TextBox txtOrdNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   1965
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "채혈상태로"
         Height          =   375
         Index           =   1
         Left            =   1785
         Style           =   1  '그래픽
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1515
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "처방상태로"
         Height          =   375
         Index           =   0
         Left            =   270
         Style           =   1  '그래픽
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1020
         Value           =   -1  'True
         Width           =   1515
      End
      Begin FPSpread.vaSpread tblSpc 
         Height          =   675
         Left            =   360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   13380
         _Version        =   196608
         _ExtentX        =   23601
         _ExtentY        =   1191
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   6
         MaxRows         =   1
         OperationMode   =   1
         ScrollBars      =   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS108.frx":10DF
         TextTip         =   4
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   360
         Left            =   4740
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   7425
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   10095
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblDOB 
         Height          =   375
         Left            =   12765
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblDoct 
         Height          =   360
         Left            =   4740
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   360
         Left            =   7425
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblWardId 
         Height          =   360
         Left            =   10095
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblDone 
         Height          =   375
         Left            =   360
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2460
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   661
         BackColor       =   14411494
         ForeColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   255
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   75
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3225
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "처방내역"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   2
      Left            =   75
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액 준비 사항"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   3
      Left            =   10815
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3225
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "취소사유"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   4965
      Left            =   10830
      TabIndex        =   8
      Top             =   3480
      Width           =   3645
      Begin VB.CheckBox chkRefund 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Refund to Patient"
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Tag             =   "10404"
         Top             =   4605
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.TextBox txtReason 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   555
         Width           =   3435
      End
      Begin VB.ComboBox cboReason 
         BackColor       =   &H00FCEFE9&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   210
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmBBS108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TblColumn      '처방정보
    tcORDDT = 1
    tcORDNM
    tcREQDT
    tcUNIT
    tcREASON
    tcDOCT
    tcSTAT
    tcREMARK
    tcORDNO
    tcORDSEQ
    tcDCFG
End Enum
Private Enum TblColumn2     '혈액정보
    tcASSIGN = 1
    tcCANCLE
    tcDELIVERY
    tcEXP
    tcRET
    tcBAG
End Enum
Private Enum TBLCOLUMN3     '검체정보
    tcSPCNO = 1
    tcSAVEPOS
    tcCOLDT
    tcCOLNM
    tcACCDT
    tcACCNM
End Enum
Private RcvDtFormat As Long
Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Clear()
    txtOrdNo = ""
    txtReason = ""
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblDOB.Caption = ""
    lblDeptNm.Caption = ""
    lblWardId.Caption = ""
    lblDoct.Caption = ""
    lblDone.Caption = ""
    
    With tblSpc
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With tblResult
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With

    With tblDStatus
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    '감염관리
    Call ICSPatientMark
    
End Sub


Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objNum As New clsBBSNumbers
    Dim objMaster As New clsCom003
    Dim strNowDate As String
    
    Clear
    strNowDate = Format(GetSystemDate, PRESENTDATE_FORMAT)
    '취소사유를 넣자.
    objMaster.AddComboBox BC2_CANCELRSN, cboReason
    
    With objNum
        RcvDtFormat = Len(.Get_AccdtFormat)
    End With
    Set objNum = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub txtOrdNo_Change()
    Dim lngLen As Long

    If RcvDtFormat = 4 Then
        With txtOrdNo
            lngLen = Len(Trim(.Text))
            If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    Else
        With txtOrdNo
            lngLen = Len(Trim(.Text))
            If lngLen = RcvDtFormat Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    End If
'
'
'    Dim lngLen As Long
'
'    With txtOrdNo
'        lngLen = Len(Trim(.Text))
'        If lngLen = RcvDtFormat Then
'            .Text = .Text & "-"
'            .SelStart = Len(.Text)
'        End If
'    End With
End Sub

Private Sub txtOrdNo_GotFocus()
    txtOrdNo.tag = txtOrdNo
    txtOrdNo.SelStart = 0
    txtOrdNo.SelLength = Len(txtOrdNo)
    
End Sub

Private Sub txtOrdNo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
    Dim strTmp As String
    
    strTmp = Mid(Format(GetSystemDate, "YYYY"), 1, 2)
    If KeyCode = vbKeyReturn Then
        If txtOrdNo <> "" Then
            txtOrdNo = strTmp & medGetP(txtOrdNo, 1, "-") & "-" & medGetP(txtOrdNo, 2, "-")
            SendKeys "{TAB}"
        End If
    End If
    
End Sub

Private Sub txtOrdNo_KeyPress(KeyAscii As Integer)
    If RcvDtFormat = 4 Then
        If Len(txtOrdNo) <> RcvDtFormat - 2 Then
            If KeyAscii = vbKeyInsert Then KeyAscii = 0
        End If
        
        If KeyAscii = vbKeyBack Then
            With txtOrdNo
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 2)
                    .SelStart = Len(.Text)
                    KeyAscii = 0
                End If
            End With
        End If
    Else
        If Len(txtOrdNo) <> RcvDtFormat Then
            If KeyAscii = vbKeyInsert Then KeyAscii = 0
        End If
        
        If KeyAscii = vbKeyBack Then
            With txtOrdNo
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 2)
                    .SelStart = Len(.Text)
                    KeyAscii = 0
                End If
            End With
        End If
    End If
'
'    If Len(txtOrdNo) <> RcvDtFormat Then
'        If KeyAscii = vbKeyInsert Then KeyAscii = 0
'    End If
'
'    If KeyAscii = vbKeyBack Then
'        With txtOrdNo
'            If .Text = "" Or .Text = "-" Then Exit Sub
'            If Mid(.Text, Len(.Text)) = "-" Then
'                .Text = Mid(.Text, 1, Len(.Text) - 2)
'                .SelStart = Len(.Text)
'                KeyAscii = 0
'            End If
'        End With
'    End If

End Sub

Private Sub txtOrdNo_LostFocus()
    If Trim(txtOrdNo) = "" Then
        Call Clear
        Exit Sub
    End If
    If txtOrdNo.tag = txtOrdNo Then Exit Sub
    
    Call Display
    txtOrdNo.tag = txtOrdNo
End Sub

Private Function Search_chk() As Boolean
    If Len(txtOrdNo) <= RcvDtFormat Then
        MsgBox "접수번호가 완전하지 않습니다.", vbCritical + vbOKOnly, "입력오류"
        Exit Function
    End If
    Search_chk = True
End Function
Private Sub Get_SpcInfo(ByVal PtId As String, ByVal orddt As String)
'검체정보를 가지고 온다.
    Dim objSQL         As New clsGetSqlStatement
    Dim Rs             As Recordset
    
    Set Rs = objSQL.Get_SpcInFormation(PtId, orddt)
    
    With tblSpc
        .Row = 1
        .Col = TBLCOLUMN3.tcSPCNO:   .value = Rs.Fields("spcyy").value & "" & "-" & _
                                              Rs.Fields("spcno").value & ""
        .Col = TBLCOLUMN3.tcSAVEPOS: .value = Rs.Fields("storeleg").value & "" & "(" & _
                                              Rs.Fields("storerno").value & "" & "," & _
                                              Rs.Fields("storecno").value & "" & ")"
        .Col = TBLCOLUMN3.tcCOLDT:   .value = Format(Rs.Fields("coldt").value & "", "####-##-##") & " " & _
                                              Format(Mid(Rs.Fields("coltm").value & "", 1, 4), "##:##")
        .Col = TBLCOLUMN3.tcCOLNM:   .value = GetEmpNm(Rs.Fields("colid").value & "")
        .Col = TBLCOLUMN3.tcACCDT:   .value = Format(Rs.Fields("rcvdt").value & "", "####-##-##") & " " & _
                                              Format(Mid(Rs.Fields("rcvtm").value & "", 1, 4), "##:##")
        .Col = TBLCOLUMN3.tcACCNM:   .value = GetEmpNm(Rs.Fields("rcvid").value & "")
    End With
    Set objSQL = Nothing
End Sub
Private Sub Get_PtInfo()
'환자정보를 가지고 온다.처방정보도 같이 가지고 온다.

    Dim objSQL         As clsGetSqlStatement
    Dim objTransReason As clsQueryOrder
    Dim Rs             As Recordset
    Dim strReason      As String
    Dim strTmp         As String
    
    '-----------------------------------------------------------------------------------
    '접수번호를 가지고, 환자정보와 처방정보를  구한다.........
    '-----------------------------------------------------------------------------------
    Set Rs = New Recordset
    Set objSQL = New clsGetSqlStatement
    Set objTransReason = New clsQueryOrder

    Set Rs = objSQL.Get_Ptinformation(txtOrdNo, RcvDtFormat)
    If Not Rs.EOF Then
        lblPtId.Caption = Rs.Fields("ptid").value & ""
        lblPtNm.Caption = Rs.Fields("ptnm").value & ""
        
        strTmp = SDA_String(Rs.Fields("ssn").value & "")
        lblDOB.Caption = medGetP(strTmp, 2, COL_DIV)
        lblSexAge.Caption = medGetP(strTmp, 1, COL_DIV) & "/" & medGetP(strTmp, 3, COL_DIV)
        
        If TRANS_REQUIRE_USED Then
            lblDone.Caption = IIf(Val(Rs.Fields("donefg").value & "") = BBSOrdStatus.stsACCESS, "", "검사진행중인 접수번호입니다")
        Else
            lblDone.Caption = IIf(Val(Rs.Fields("donefg").value & "") = BBSOrderStatus.stsACCESS, "", "검사진행중인 접수번호입니다")
        End If
        
        lblWardId.Caption = GetWardNm(Rs.Fields("deptcd").value & "")
        lblDeptNm.Caption = GetDeptNm(Rs.Fields("deptcd").value & "")
        lblDoct.Caption = GetDoctNm(Rs.Fields("orddoct").value & "")
        
        With tblResult
            .Row = 1
            .Col = TblColumn.tcORDDT: .value = Format(Rs.Fields("orddt").value & "", "####-##-##")
            .Col = TblColumn.tcORDNM: .value = Rs.Fields("testnm").value & ""
            .Col = TblColumn.tcREQDT: .value = Format(Rs.Fields("reqdt").value & "", "####-##-##")
            .Col = TblColumn.tcUNIT:  .value = CLng(Rs.Fields("unitqty").value & "")
            
            strReason = objTransReason.GetTransReason(lblPtId.Caption, Trim(Rs.Fields("orddt").value & ""), Trim(Rs.Fields("ordno")))
            .Col = TblColumn.tcREASON: .value = strReason
            .Col = TblColumn.tcDOCT:   .value = lblDoct.Caption
            
            .Col = TblColumn.tcSTAT
            If TRANS_REQUIRE_USED Then
                    Select Case Rs.Fields("stscd").value & ""
                        Case BBSOrdStatus.stsORDER:     .ForeColor = RGB(0, 0, 0):   .value = "처방"
                        Case BBSOrdStatus.stsCOLLECT:   .ForeColor = RGB(0, 255, 0): .value = "채혈"
                        Case BBSOrdStatus.stsACCESS:    .ForeColor = RGB(0, 0, 255): .value = "접수"
                        Case BBSOrdStatus.stsINPROCESS: .ForeColor = RGB(255, 0, 0): .value = "검사중"
                        Case Else:                      .ForeColor = RGB(0, 0, 0):   .value = ""
                    End Select
            Else
                    Select Case Rs.Fields("stscd").value & ""
                        Case BBSOrderStatus.stsORDER:     .ForeColor = RGB(0, 0, 0):   .value = "처방"
                        Case BBSOrderStatus.stsCOLLECT:   .ForeColor = RGB(0, 255, 0): .value = "채혈"
                        Case BBSOrderStatus.stsACCESS:    .ForeColor = RGB(0, 0, 255): .value = "접수"
                        Case BBSOrderStatus.stsINPROCESS: .ForeColor = RGB(255, 0, 0): .value = "검사중"
                        Case Else:                        .ForeColor = RGB(0, 0, 0):   .value = ""
                    End Select
            End If
            
            .Col = TblColumn.tcREMARK: .value = IIf(IsNull(Rs.Fields("mesg").value) = True, "", Rs.Fields("mesg").value & "")
            .Col = TblColumn.tcORDNO:  .value = CLng(Rs.Fields("ordno").value & "")
            .Col = TblColumn.tcORDSEQ: .value = CLng(Rs.Fields("ordseq").value & "")
            .ForeColor = vbRed
            .Col = TblColumn.tcDCFG: .value = IIf(Rs.Fields("dcfg").value & "" = "1", "Y", "")
            .ForeColor = vbBlack
        End With
        
        '-----------------------------------------------------------------------------------
        '접수번호를 가지고, Pheresis 처방이 아닐경우 검체정보를 구한다.........
        '-----------------------------------------------------------------------------------
        If objSQL.Get_PheresisChk(txtOrdNo, RcvDtFormat) = False Then
            Get_SpcInfo Rs.Fields("ptid").value & "", Rs.Fields("orddt").value & ""
        End If
        
    End If
    '감염관리
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    
    Set Rs = Nothing
    Set objSQL = Nothing
    Set objTransReason = Nothing
End Sub



Private Function Display()
    Dim objSQL         As New clsGetSqlStatement
    Dim objTransReason As New clsQueryOrder
    Dim DrRS           As New Recordset
    Dim strReason      As String
    Dim i              As Integer
    
    If Search_chk() = False Then Exit Function
    
'    objSql.setDbConn DBConn
    If objSQL.Get_CollectChk(txtOrdNo, RcvDtFormat) = False Then
        Clear
        Set objSQL = Nothing
        Set objTransReason = Nothing
        Exit Function
    End If
    
    medClearTable tblSpc
    medClearTable tblResult
    medClearTable tblDStatus
    '-----------------------------------------------------------------------------------
    '접수번호를 가지고, 환자정보와 처방정보를  구한후
    '검체정보를 조회한다.
    '-----------------------------------------------------------------------------------
    Call Get_PtInfo
    Call SetDStatus
    Exit Function
End Function

Private Sub SetDStatus()
    Dim objDStatus As clsDetailStatus
    
    Set objDStatus = New clsDetailStatus
    
    objDStatus.WorkArea = C_WORKAREA
    objDStatus.accdt = medGetP(txtOrdNo, 1, "-")
    objDStatus.accseq = medGetP(txtOrdNo, 2, "-")
    If objDStatus.GetCount = True Then
        With tblDStatus
            .Row = 1
            .Col = TblColumn2.tcASSIGN:   .value = objDStatus.AssignCnt
            .Col = TblColumn2.tcCANCLE:   .value = objDStatus.AssignCancelCnt
            .Col = TblColumn2.tcDELIVERY: .value = objDStatus.DeliveryCnt
            .Col = TblColumn2.tcEXP:      .value = objDStatus.ExpCnt
            .Col = TblColumn2.tcRET:      .value = objDStatus.RetCnt
            .Col = TblColumn2.tcBAG:      .value = objDStatus.BagCnt
        End With
    Else
        With tblDStatus
            .Row = 1
            .Col = TblColumn2.tcASSIGN:   .value = ""
            .Col = TblColumn2.tcCANCLE:   .value = ""
            .Col = TblColumn2.tcDELIVERY: .value = ""
            .Col = TblColumn2.tcEXP:      .value = ""
            .Col = TblColumn2.tcRET:      .value = ""
            .Col = TblColumn2.tcBAG:      .value = ""
        End With
    End If
    Set objDStatus = Nothing
End Sub
Private Function CancelChk() As Boolean
    Dim accdt As String
    Dim accseq As String
    Dim ii As Integer
    
    CancelChk = True
    
    If Trim(txtOrdNo) = "" Then CancelChk = False: Exit Function
    
    accdt = medGetP(txtOrdNo, 1, "-")
    accseq = medGetP(txtOrdNo, 2, "-")
    If accdt = "" Or accseq = "" Then
        MsgBox "접수번호가 완전하지 않습니다.", vbCritical, Me.Caption
        CancelChk = False
        Exit Function
    End If
    
    If cboReason.ListIndex < 0 Then
        MsgBox "취소사유를 선택하세요.", vbInformation + vbOKOnly, Me.Caption
        CancelChk = False
        Exit Function
    End If
    
    With tblDStatus
        .Row = 1
        For ii = 3 To .MaxCols
            .Col = ii
            If CLng(.value) > 0 Then
                CancelChk = False
                MsgBox "해당 접수번호에 대해서 이미 검사가 " & vbNewLine & _
                       "진행되었기에 해당접수번호를 취소할수 없습니다.", vbCritical + vbOKOnly, "접수취소"
                Exit Function
            End If
        Next
    End With
    
End Function
Private Sub cmdSave_Click()
    If CancelChk = False Then Exit Sub
    
    Dim objBg        As New clsBeginTrans
    Dim strCancelrsn As String
    Dim strStatus    As String
    Dim CancelCnt    As Long
    Dim strSQL       As String
    
    strCancelrsn = medGetP(cboReason.List(cboReason.ListIndex), 1, " ")
    
    If optStatus(0).value = True = True Then
        If TRANS_REQUIRE_USED Then
            strStatus = BBSOrdStatus.stsORDER
        Else
            strStatus = BBSOrdStatus.stsORDER
        End If
    Else
        If TRANS_REQUIRE_USED Then
            strStatus = BBSOrderStatus.stsCOLLECT
        Else
            strStatus = BBSOrderStatus.stsCOLLECT
        End If
    
    End If
    
    With tblDStatus
        .Row = 1
        .Col = 1: CancelCnt = CLng(.value)
        .Col = 2: CancelCnt = CLng(.value) + CancelCnt
    End With

    If objBg.Set_AccessCancel(lblPtId.Caption, txtOrdNo, strStatus, strCancelrsn, _
                                    CStr(ObjMyUser.EmpId), txtReason, CancelCnt) = True Then
        MsgBox "접수 취소되었습니다.", vbInformation + vbOKOnly, Me.Caption
        Clear
        cboReason.ListIndex = -1
    Else
        MsgBox "접수취소 오류입니다. ", vbCritical + vbOKOnly, "접수취소오류"
    End If
    Set objBg = Nothing
    
End Sub


