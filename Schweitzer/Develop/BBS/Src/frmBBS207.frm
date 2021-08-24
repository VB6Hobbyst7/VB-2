VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS207 
   BackColor       =   &H00DBE6E6&
   Caption         =   "수혈부작용등록"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmBBS207.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   45
      TabIndex        =   24
      Top             =   405
      Width           =   3420
      Begin VB.OptionButton optQDiv 
         BackColor       =   &H80000005&
         Caption         =   "출고일기준조회"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   26
         Top             =   30
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton optQDiv 
         BackColor       =   &H80000005&
         Caption         =   "폐기일기준 조회"
         Height          =   255
         Index           =   1
         Left            =   1725
         TabIndex        =   25
         Top             =   30
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtpQFrom 
         Height          =   285
         Left            =   750
         TabIndex        =   27
         Top             =   315
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   37915
      End
      Begin MSComCtl2.DTPicker dtpQTo 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   315
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   37915
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "출고일 :"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   30
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "~"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   29
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H80000005&
      Caption         =   "  출고일      조회(&Q)"
      Height          =   675
      Left            =   3465
      Style           =   1  '그래픽
      TabIndex        =   23
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "부작용등록(S)"
      Height          =   525
      Left            =   11910
      Style           =   1  '그래픽
      TabIndex        =   20
      Top             =   8595
      Width           =   1290
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(X)"
      Height          =   525
      Left            =   13215
      Style           =   1  '그래픽
      TabIndex        =   19
      Top             =   8595
      Width           =   1290
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   1710
      Left            =   4485
      TabIndex        =   13
      Top             =   4245
      Width           =   4950
      Begin VB.CheckBox chkBeforeTest 
         Caption         =   "관련검사결과보기"
         Height          =   255
         Left            =   3165
         TabIndex        =   14
         Top             =   45
         Width           =   1740
      End
      Begin FPSpread.vaSpread tblBeforeResult 
         Height          =   1290
         Left            =   90
         TabIndex        =   15
         Top             =   360
         Width           =   4830
         _Version        =   196608
         _ExtentX        =   8520
         _ExtentY        =   2275
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
         MaxCols         =   4
         MaxRows         =   7
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmBBS207.frx":076A
      End
      Begin FPSpread.vaSpread tblBeforeRealTest 
         Height          =   1770
         Left            =   105
         TabIndex        =   16
         Top             =   1980
         Width           =   4770
         _Version        =   196608
         _ExtentX        =   8414
         _ExtentY        =   3122
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
         MaxCols         =   3
         MaxRows         =   7
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmBBS207.frx":0B62
      End
      Begin VB.Label Label6 
         Caption         =   "♣ Cross-Matching 검사"
         Height          =   255
         Left            =   30
         TabIndex        =   18
         Top             =   90
         Width           =   2085
      End
      Begin VB.Label Label7 
         Caption         =   "♣ 관련검사 결과"
         Height          =   255
         Left            =   15
         TabIndex        =   17
         Top             =   1725
         Width           =   2085
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   1710
      Left            =   9510
      TabIndex        =   7
      Top             =   4245
      Width           =   4950
      Begin VB.CheckBox chkAfterTest 
         Caption         =   "관련검사결과보기"
         Height          =   255
         Left            =   3270
         TabIndex        =   8
         Top             =   75
         Width           =   1740
      End
      Begin FPSpread.vaSpread tblAfterResult 
         Height          =   1290
         Left            =   90
         TabIndex        =   9
         Top             =   360
         Width           =   4830
         _Version        =   196608
         _ExtentX        =   8520
         _ExtentY        =   2275
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   4
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmBBS207.frx":0F47
      End
      Begin FPSpread.vaSpread tblAfterRealTest 
         Height          =   1770
         Left            =   105
         TabIndex        =   10
         Top             =   1980
         Width           =   4770
         _Version        =   196608
         _ExtentX        =   8414
         _ExtentY        =   3122
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
         MaxCols         =   3
         MaxRows         =   7
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmBBS207.frx":12E5
      End
      Begin VB.Label Label8 
         Caption         =   "♣ 관련검사 결과"
         Height          =   255
         Left            =   15
         TabIndex        =   12
         Top             =   1725
         Width           =   2085
      End
      Begin VB.Label Label9 
         Caption         =   "♣Cross-Matching검사(1:적격,2:부적격)"
         Height          =   255
         Left            =   -30
         TabIndex        =   11
         Top             =   90
         Width           =   3435
      End
   End
   Begin VB.TextBox txtVolume 
      Height          =   270
      Left            =   8820
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6510
      Width           =   1545
   End
   Begin VB.TextBox txtMesg 
      BackColor       =   &H00F7FDF8&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   6030
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      ToolTipText     =   "검사 리마크를 입력하세요."
      Top             =   7320
      Width           =   8205
   End
   Begin VB.TextBox txtReactioncd 
      Height          =   315
      Left            =   12030
      TabIndex        =   1
      Top             =   6525
      Width           =   1155
   End
   Begin VB.CommandButton cmdPop 
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   13185
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   6525
      Width           =   300
   End
   Begin MSComCtl2.DTPicker dtpTtm 
      Height          =   270
      Left            =   8850
      TabIndex        =   3
      Top             =   6825
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   476
      _Version        =   393216
      Format          =   62980098
      CurrentDate     =   37915
   End
   Begin MSComCtl2.DTPicker dtpFtm 
      Height          =   300
      Left            =   6000
      TabIndex        =   4
      Top             =   6810
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62980098
      CurrentDate     =   37915
   End
   Begin MSComCtl2.DTPicker dtpTransdt 
      Height          =   300
      Left            =   6000
      TabIndex        =   6
      Top             =   6480
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62980097
      CurrentDate     =   37915
   End
   Begin MedControls1.LisLabel lblBABO 
      Height          =   810
      Left            =   7665
      TabIndex        =   21
      Top             =   1980
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1429
      BackColor       =   14411494
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "AB(AB)+"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblPABO 
      Height          =   810
      Left            =   7950
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   510
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1429
      BackColor       =   14411494
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "AB(AB)+"
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   7440
      Left            =   45
      TabIndex        =   31
      Top             =   1080
      Width           =   4395
      _Version        =   196608
      _ExtentX        =   7752
      _ExtentY        =   13123
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
      MaxCols         =   4
      MaxRows         =   50
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS207.frx":16CA
   End
   Begin FPSpread.vaSpread tblBlood 
      Height          =   2010
      Left            =   9495
      TabIndex        =   32
      Top             =   1845
      Width           =   5025
      _Version        =   196608
      _ExtentX        =   8864
      _ExtentY        =   3545
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
      MaxCols         =   30
      MaxRows         =   7
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS207.frx":1C88
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   285
      Left            =   4740
      TabIndex        =   33
      Top             =   7290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      BackColor       =   16249848
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "◈ Remark"
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆  출고 혈액 리스트"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   45
      TabIndex        =   82
      Top             =   105
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆  환 자 정 보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   4455
      TabIndex        =   81
      Top             =   135
      Width           =   4995
   End
   Begin VB.Shape Shape2 
      Height          =   1005
      Index           =   0
      Left            =   4455
      Top             =   405
      Width           =   4995
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "환자 정보  :"
      Height          =   195
      Index           =   2
      Left            =   4530
      TabIndex        =   80
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "진료 정보  :"
      Height          =   195
      Index           =   3
      Left            =   4530
      TabIndex        =   79
      Top             =   1110
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "성별/나이 :"
      Height          =   195
      Index           =   4
      Left            =   4530
      TabIndex        =   78
      Top             =   810
      Width           =   1065
   End
   Begin VB.Label lblPtnm 
      BackColor       =   &H00F7F0F0&
      Caption         =   "김정규"
      Height          =   195
      Left            =   6855
      TabIndex        =   77
      Top             =   525
      Width           =   1020
   End
   Begin VB.Label lblptid 
      BackColor       =   &H00F7F0F0&
      Caption         =   "00000002"
      Height          =   210
      Left            =   5640
      TabIndex        =   76
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label lblSexAge 
      BackColor       =   &H00F7F0F0&
      Caption         =   "남/30"
      Height          =   210
      Left            =   5640
      TabIndex        =   75
      Top             =   810
      Width           =   1185
   End
   Begin VB.Label lblLocation 
      BackColor       =   &H00F7F0F0&
      Caption         =   "21W-123-A15"
      Height          =   210
      Left            =   5655
      TabIndex        =   74
      Top             =   1110
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆ 혈 액 정 보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   4470
      TabIndex        =   73
      Top             =   1575
      Width           =   4995
   End
   Begin VB.Shape Shape2 
      Height          =   2010
      Index           =   1
      Left            =   4470
      Top             =   1845
      Width           =   4995
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆ 처 방 정 보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   9480
      TabIndex        =   72
      Top             =   135
      Width           =   5055
   End
   Begin VB.Shape Shape2 
      Height          =   1005
      Index           =   2
      Left            =   9480
      Top             =   390
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "처 방 일  :"
      Height          =   195
      Index           =   5
      Left            =   9540
      TabIndex        =   71
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "접수번호 :"
      Height          =   195
      Index           =   6
      Left            =   11850
      TabIndex        =   70
      Top             =   510
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "처방코드 :"
      Height          =   195
      Index           =   7
      Left            =   9540
      TabIndex        =   69
      Top             =   795
      Width           =   900
   End
   Begin VB.Label lblTestNm 
      BackColor       =   &H00F7F0F0&
      Caption         =   "X2021 Whole Blood 400cc"
      Height          =   225
      Left            =   10500
      TabIndex        =   68
      Top             =   795
      Width           =   3975
   End
   Begin VB.Label lblAccdt 
      BackColor       =   &H00F7F0F0&
      Caption         =   "Label3"
      Height          =   225
      Left            =   13020
      TabIndex        =   67
      Top             =   510
      Width           =   1440
   End
   Begin VB.Label lblordDt 
      BackColor       =   &H00F7F0F0&
      Caption         =   "2003-10-10"
      Height          =   195
      Left            =   10500
      TabIndex        =   66
      Top             =   510
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "처방Msg :"
      Height          =   195
      Index           =   8
      Left            =   9540
      TabIndex        =   65
      Top             =   1095
      Width           =   900
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00F7F0F0&
      Caption         =   "Label3"
      Height          =   225
      Left            =   10500
      TabIndex        =   64
      Top             =   1110
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "혈액번호 :"
      Height          =   195
      Index           =   9
      Left            =   4515
      TabIndex        =   63
      Top             =   2025
      Width           =   930
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "입 고 일  :"
      Height          =   195
      Index           =   10
      Left            =   4515
      TabIndex        =   62
      Top             =   2925
      Width           =   930
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "혈액제제 :"
      Height          =   195
      Index           =   11
      Left            =   4515
      TabIndex        =   61
      Top             =   2325
      Width           =   930
   End
   Begin VB.Label lblBloodNo 
      BackColor       =   &H00F7F0F0&
      Caption         =   "01-01-123456"
      Height          =   195
      Left            =   5535
      TabIndex        =   60
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label lblComponm 
      BackColor       =   &H00F7F0F0&
      Caption         =   "01 Whole Blood"
      Height          =   195
      Left            =   5520
      TabIndex        =   59
      Top             =   2325
      Width           =   1605
   End
   Begin VB.Label lblEntdt 
      BackColor       =   &H00F7F0F0&
      Caption         =   "2003-10-22"
      Height          =   195
      Left            =   5535
      TabIndex        =   58
      Top             =   2925
      Width           =   1050
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검 사 일  :"
      Height          =   195
      Index           =   12
      Left            =   4515
      TabIndex        =   57
      Top             =   3240
      Width           =   930
   End
   Begin VB.Label lblVfyDt 
      BackColor       =   &H00F7F0F0&
      Caption         =   "2003-10-22"
      Height          =   195
      Left            =   5535
      TabIndex        =   56
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출 고 자  :"
      Height          =   195
      Index           =   13
      Left            =   6675
      TabIndex        =   55
      Top             =   3525
      Width           =   930
   End
   Begin VB.Label lblDeliveryNm 
      BackColor       =   &H00F7F0F0&
      Caption         =   "21004  김 정 구"
      Height          =   195
      Left            =   7755
      TabIndex        =   54
      Top             =   3525
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출 고 일  :"
      Height          =   195
      Index           =   14
      Left            =   4515
      TabIndex        =   53
      Top             =   3525
      Width           =   930
   End
   Begin VB.Label lblDeliverydt 
      BackColor       =   &H00F7F0F0&
      Caption         =   "2003-10-22"
      Height          =   195
      Left            =   5520
      TabIndex        =   52
      Top             =   3540
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆ 출고혈액조회(2003-10-21 ~2003-10-22)(label1(4)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   9495
      TabIndex        =   51
      Top             =   1575
      Width           =   5010
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "입 고 자  :"
      Height          =   195
      Index           =   15
      Left            =   6675
      TabIndex        =   50
      Top             =   2925
      Width           =   930
   End
   Begin VB.Label lblEntNm 
      BackColor       =   &H00F7F0F0&
      Caption         =   "21004  김 정 구"
      Height          =   195
      Left            =   7755
      TabIndex        =   49
      Top             =   2925
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검 사 자  :"
      Height          =   195
      Index           =   16
      Left            =   6675
      TabIndex        =   48
      Top             =   3240
      Width           =   930
   End
   Begin VB.Label lblVfyNm 
      BackColor       =   &H00F7F0F0&
      Caption         =   "21004  김 정 구"
      Height          =   195
      Left            =   7755
      TabIndex        =   47
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "혈액용량 :"
      Height          =   195
      Index           =   17
      Left            =   4515
      TabIndex        =   46
      Top             =   2625
      Width           =   930
   End
   Begin VB.Label lblVolumn 
      BackColor       =   &H00F7F0F0&
      Caption         =   "400cc"
      Height          =   195
      Left            =   5520
      TabIndex        =   45
      Top             =   2625
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆ 수혈전 검사결과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   4470
      TabIndex        =   44
      Top             =   3960
      Width           =   4995
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆ 수혈후 검사결과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   6
      Left            =   9495
      TabIndex        =   43
      Top             =   3960
      Width           =   5010
   End
   Begin VB.Shape Shape2 
      Height          =   1740
      Index           =   3
      Left            =   4470
      Top             =   4230
      Width           =   4995
   End
   Begin VB.Shape Shape2 
      Height          =   1740
      Index           =   4
      Left            =   9495
      Top             =   4230
      Width           =   4995
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆ 수혈 부작용 등록 정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   7
      Left            =   4485
      TabIndex        =   42
      Top             =   6090
      Width           =   2370
   End
   Begin VB.Shape Shape2 
      Height          =   2145
      Index           =   5
      Left            =   4500
      Top             =   6390
      Width           =   10005
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "수 혈 일        :"
      Height          =   195
      Index           =   18
      Left            =   4695
      TabIndex        =   41
      Top             =   6525
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "수 혈 량        :"
      Height          =   195
      Index           =   19
      Left            =   7560
      TabIndex        =   40
      Top             =   6540
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "수혈시작시간 :"
      Height          =   195
      Index           =   20
      Left            =   4695
      TabIndex        =   39
      Top             =   6855
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "수혈종료시간 :"
      Height          =   195
      Index           =   21
      Left            =   7560
      TabIndex        =   38
      Top             =   6885
      Width           =   1290
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F7F3F8&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   1245
      Index           =   8
      Left            =   4680
      Shape           =   4  '둥근 사각형
      Top             =   7215
      Width           =   9660
   End
   Begin VB.Label lblReactionChk 
      BackColor       =   &H009E9383&
      Caption         =   "▶ 이미 부작용등록된 혈액입니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   9495
      TabIndex        =   37
      Top             =   6060
      Width           =   4380
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "부작용 사유코드 :"
      Height          =   195
      Index           =   22
      Left            =   10455
      TabIndex        =   36
      Top             =   6570
      Width           =   1530
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "부작용 사유내역 :"
      Height          =   195
      Index           =   23
      Left            =   10455
      TabIndex        =   35
      Top             =   6885
      Width           =   1530
   End
   Begin VB.Label lblReactionNm 
      BackColor       =   &H00F7F0F0&
      Height          =   225
      Left            =   12045
      TabIndex        =   34
      Top             =   6855
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   3
      Left            =   9480
      Top             =   45
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   1
      Left            =   4455
      Top             =   45
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   0
      Left            =   45
      Top             =   45
      Width           =   4380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   7
      Left            =   4485
      Top             =   5985
      Width           =   10005
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   6
      Left            =   9495
      Top             =   3870
      Width           =   5010
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   5
      Left            =   4470
      Top             =   3870
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   4
      Left            =   9495
      Top             =   1485
      Width           =   5010
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   2
      Left            =   4470
      Top             =   1485
      Width           =   4995
   End
End
Attribute VB_Name = "frmBBS207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Private Enum TblCol
    tcChk = 1
    tcDELIVERYDT
    tcBLDNO
    tcVOL
    tcCOMPONM
    
    tcSTEP1
    tcSTEP2
    tcSTEP3
    tcSTEP4
    tcRSTV
    
    tcPTID
    tcORDDT
    tcORDNO
    tcORDSEQ
    tcWARDID
    
    tcHOSILID
    tcDEPTCD
    tcWORKAREA
    tcACCDT
    tcACCSEQ
    
    tcRSTSEQ
    tcCompocd
    TcABO
    tcRh
    tcSTSCD
    
    tcDeliveryid
    tcENTDT
    tcEntid
    tcVFYDT
    tcVfyid
End Enum

Private Sub Form_Clear(Optional ByVal blnClear As Boolean = True)

    If blnClear = True Then
        optQDiv(0).value = True
        dtpQFrom.value = DateAdd("d", -3, GetSystemDate)
        dtpQTo.value = GetSystemDate
        Call medClearTable(tblPtList)
    End If
    lblptid.Caption = ""
    lblPtnm.Caption = ""
    lblSexAge.Caption = ""
    lblLocation.Caption = ""
    lblPABO.Caption = ""
    
    lblBloodNo.Caption = ""
    lblComponm.Caption = ""
    lblVolumn.Caption = ""
    lblEntdt.Caption = ""
    lblEntNm.Caption = ""
    lblVfyDt.Caption = ""
    lblVfyNm.Caption = ""
    lblDeliverydt.Caption = ""
    lblDeliveryNm.Caption = ""
    lblBABO.Caption = ""
    lblordDt.Caption = ""
    lblAccdt.Caption = ""
    lblTestNm.Caption = ""
    lblMsg.Caption = ""
    
    Call medClearTable(tblBlood)
    chkBeforeTest.value = 0
    chkAfterTest.value = 0
    Call medClearTable(tblBeforeResult)
    Call medClearTable(tblAfterResult)
    Call medClearTable(tblBeforeRealTest)
    Call medClearTable(tblAfterRealTest)
    
    dtpTransdt.value = GetSystemDate
    txtVolume.Text = "0"
    dtpFtm.value = GetSystemDate
    dtpTtm.value = GetSystemDate
    txtMesg.Text = ""
    lblReactionChk.Visible = False
    lblReactionNm.Caption = ""
    txtReactioncd.Text = ""
    Call Form_Setting
End Sub

Private Sub chkAfterTest_Click()
    If chkAfterTest.value = 1 Then
        Frame3.Height = 3825
        Shape2(4).Height = 3855
        Frame3.ZOrder 0
'
    Else
        Frame3.Height = 1710
        Shape2(4).Height = 1740
        Frame2.ZOrder 0
    End If
End Sub

Private Sub chkBeforeTest_Click()
    DoEvents
    If chkBeforeTest.value = 1 Then
        Frame2.Height = 3825
        Shape2(3).Height = 3855
        Frame2.ZOrder 0
    Else
        Frame2.Height = 1710
        Shape2(3).Height = 1740
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPop_Click(Index As Integer)
    Dim objSQL  As New clsTransfusion
    Dim lngTop  As Long
    Dim lngLeft As Long
    Dim SSQL    As String
    
    SSQL = objSQL.GetReactionSQL
    
    
    Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = DBConn
        lngTop = txtReactioncd.Top + 3850
        lngLeft = Me.Left + txtReactioncd.Left + 1050
        .FormCaption = "수혈부작용 리스트"
        .ColumnHeaderText = "사유코드;부작용명"
        Call .LoadPopUp(SSQL) ', lngTop, lngLeft)
        txtReactioncd.Text = Trim(medGetP(.SelectedString, 1, ";"))
        lblReactionNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
    End With
    Set objSQL = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    Dim sStscd  As String
    Dim sFDate  As String
    Dim sTDate  As String
    Dim SSQL    As String
    Dim sTmp    As String
    Dim strSSN  As String
    
    Me.MousePointer = 11
    Call medClearTable(tblPtList)
    Call Form_Clear(False)
    
    If optQDiv(0).value Then
        Label1(4).Caption = "◆ 출고혈액조회(" & Format(dtpQFrom.value, "YYYY-MM-DD") & " ~ " & Format(dtpQTo.value, "YYYY-MM-DD") & ")"
    Else
        Label1(4).Caption = "◆ 폐기혈액조회(" & Format(dtpQFrom.value, "YYYY-MM-DD") & " ~ " & Format(dtpQTo.value, "YYYY-MM-DD") & ")"
    End If
    
    
    sFDate = Format(dtpQFrom.value, "YYYYMMDD")
    sTDate = Format(dtpQTo.value, "YYYYMMDD")
    
    sStscd = BBSBloodStatus.stsDELIVERY
    If optQDiv(1).value Then sStscd = BBSBloodStatus.stsEXPIRE
    
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    
    SSQL = objSQL.PtListQuery(sFDate, sTDate, sStscd)
    
    RS.Open SSQL, DBConn
    
    With tblPtList
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                strSSN = ""
                .Col = 1: .value = Format(RS.Fields("qdate").value & "", "####-##-##")
                If sTmp = RS.Fields("qdate").value & "" Then .ForeColor = .BackColor
                .Col = 2: .value = RS.Fields("ptid").value & ""
                .Col = 3: .value = GetPtNm(RS.Fields("ptid").value & "")
'                          Call getbbs_ptinfo(Rs.Fields("ptid").value & "", strSSN)
                .Col = 4: .value = GetSSN(RS.Fields("ptid").value & "") ' strSSN
                sTmp = RS.Fields("qdate").value & ""
                RS.MoveNext
            Loop
        End If
    End With
    Me.MousePointer = 0
    Set RS = Nothing
    Set objSQL = Nothing
    
End Sub

Private Function GetSSN(ByVal vPtID As String) As String
    Dim objSQL As New clsPatient
    Dim RS As New Recordset
    
    RS.Open objSQL.GetSQLPt(vPtID), DBConn
    
    If RS.EOF = False Then GetSSN = RS.Fields("ssn").value & ""
    
    Set RS = Nothing
    Set objSQL = Nothing
End Function

Private Sub Form_Load()
    Call Form_Clear
    Call Form_Setting
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCodeList = Nothing
End Sub

Private Sub optQDiv_Click(Index As Integer)
    If Index = 0 Then
        cmdQuery.Caption = "  출고일      조회(&Q)"
        tblPtList.Row = 0: tblPtList.Col = 1: tblPtList.value = "출고일"
        Label1(0).Caption = " ◆ 출고혈액리스트"
    Else
        cmdQuery.Caption = "  폐기일      조회(&Q)"
        tblPtList.Row = 0: tblPtList.Col = 1: tblPtList.value = "폐기일"
        Label1(0).Caption = " ◆ 폐기혈액리스트"
    End If
    
End Sub

Private Sub tblAfterResult_Advance(ByVal AdvanceNext As Boolean)
    Dim sValue  As String
    
    With tblAfterResult
        .Row = .ActiveRow
        .Col = .ActiveCol: sValue = .value
        .Col = 4:   .value = sValue
        
        .Col = .ActiveCol
        
        If sValue = "1" Or sValue = "적격" Then
            .value = "적격": .ForeColor = vbBlue
            .Col = 4: .value = "1"
        ElseIf .value = "0" Or sValue = "부적격" Then
            .value = "부적격": .ForeColor = vbRed
            .Col = 4: .value = "0"
        Else
            .value = ""
            .Col = 4: .value = ""
        End If
    End With
End Sub

Private Sub tblAfterResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim sValue  As String
    
    With tblAfterResult
        If Row < 1 Then Exit Sub
        If Col <> 2 Then Exit Sub
        .Row = Row
        .Col = 1
        If .value = "" Then
            .Col = Col: .value = ""
            Exit Sub
        End If
        
        .Col = Col: sValue = .value
        .Col = 4:   .value = sValue
        
        .Col = Col
        
        If sValue = "1" Or sValue = "적격" Then
            .value = "적격": .ForeColor = vbBlue
            .Col = 4: .value = "1"
        ElseIf .value = "0" Or sValue = "부적격" Then
            .value = "부적격": .ForeColor = vbRed
            .Col = 4: .value = "0"
        Else
            .value = ""
            .Col = 4: .value = ""
        End If
    End With
End Sub

Private Sub tblBlood_Click(ByVal Col As Long, ByVal Row As Long)
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    Dim sPtid   As String
    Dim sORDDT  As String
    Dim sOrdno  As String
    Dim sOrdseq As String
    Dim SSQL    As String
    Dim sStep1  As String
    Dim sStep2  As String
    Dim sStep3  As String
    Dim sStep4  As String
    Dim sRstv   As String
    Dim ii      As Long
    
    sPtid = lblptid.Caption
    With tblBlood
        If Row < 1 Then Exit Sub
        .Row = Row: .Col = TblCol.tcBLDNO
        If .value = "" Then Exit Sub
        Me.MousePointer = 11
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1: .value = ""
        Next
        .Row = Row
        .Col = TblCol.tcChk: .value = "▶": .ForeColor = vbRed
        
        .Col = TblCol.tcWARDID: lblLocation.Caption = .value
        .Col = TblCol.tcHOSILID:
        If .value <> "" Then
            If lblLocation.Caption <> "" Then
                lblLocation.Caption = lblLocation.Caption & "-" & .value
            Else
                lblLocation.Caption = .value
            End If
        End If
        .Col = TblCol.tcDEPTCD
        If .value <> "" Then
            If lblLocation.Caption <> "" Then
                lblLocation.Caption = lblLocation.Caption & "-" & .value
            Else
                lblLocation.Caption = .value
            End If
        End If
        
        .Col = TblCol.tcBLDNO:      lblBloodNo.Caption = .value
        .Col = TblCol.tcCompocd:    lblComponm.Caption = .value
        .Col = TblCol.tcCOMPONM:    lblComponm.Caption = lblComponm.Caption & " " & .value
        .Col = TblCol.tcVOL:        lblVolumn.Caption = .value
        .Col = TblCol.tcENTDT:      lblEntdt.Caption = .value
        .Col = TblCol.tcVFYDT:      lblVfyDt.Caption = .value
        .Col = TblCol.tcDELIVERYDT: lblDeliverydt.Caption = .value
        .Col = TblCol.tcEntid:      lblEntNm.Caption = .value & " " & GetEmpNm(.value)
        .Col = TblCol.tcVfyid:      lblVfyNm.Caption = .value & " " & GetEmpNm(.value)
        .Col = TblCol.tcDeliveryid: lblDeliveryNm.Caption = .value & " " & GetEmpNm(.value)
        .Col = TblCol.TcABO:        lblBABO.Caption = .value
        .Col = TblCol.tcWORKAREA:   lblAccdt.Caption = .value
        .Col = TblCol.tcACCDT:      lblAccdt.Caption = lblAccdt.Caption & "-" & .value
        .Col = TblCol.tcACCSEQ:     lblAccdt.Caption = lblAccdt.Caption & "-" & .value
        .Col = TblCol.tcRSTSEQ:     lblAccdt.tag = .value
        .Col = TblCol.tcORDDT:      lblordDt.Caption = .value: sORDDT = Replace(.value, "-", "")
        .Col = TblCol.tcORDNO:      sOrdno = .value
        .Col = TblCol.tcORDSEQ:     sOrdseq = .value
        .Col = TblCol.tcSTEP1:      sStep1 = .value
        .Col = TblCol.tcSTEP2:      sStep2 = .value
        .Col = TblCol.tcSTEP3:      sStep3 = .value
        .Col = TblCol.tcSTEP4:      sStep4 = .value
        .Col = TblCol.tcRSTV:       sRstv = .value
        Set RS = New Recordset
        Set objSQL = New clsTransfusion
        SSQL = objSQL.GetOrderBodySQL(sPtid, sORDDT, sOrdno, sOrdseq)
        RS.Open SSQL, DBConn
        If Not RS.EOF Then
            lblTestNm.Caption = RS.Fields("testcd").value & "" & " " & RS.Fields("testnm").value & ""
            lblMsg.Caption = RS.Fields("mesg").value & ""
        End If
    End With
    With tblBeforeResult
        .Row = 1:
                    If sRstv = "1" Then
                        .Col = 2: .value = IIf(sStep1 = "1", "적격", "")
                                  .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                        .Col = 3: .value = IIf(sStep1 = "1", lblVfyDt.Caption, "")
                        .Col = 4: .value = "1"
                    Else
                        .Col = 2: .value = "부적격"
                        .Col = 4: .value = sStep1
                    End If
                    
'                    .value = IIf(sRstv = "1", "적격", IIf(sStep1 = "1", "적격", "부적격"))
'                    .Col = 4: .value = IIf(sRstv = "1", "1", sStep1)
        .Row = 2:
                    If sRstv = "1" Then
                        .Col = 2: .value = IIf(sStep2 = "1", "적격", "")
                                  .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                        .Col = 3: .value = IIf(sStep2 = "1", lblVfyDt.Caption, "")
                        .Col = 4: .value = "1"
                    Else
                        .Col = 2: .value = "부적격"
                        .Col = 4: .value = sStep2
                    End If
        
'                    .Col = 2: .value = IIf(sRstv = "1", "적격", IIf(sStep2 = "1", "적격", "부적격"))
'                    .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
'                    .Col = 4: .value = IIf(sRstv = "1", "1", sStep2)
'                    .Col = 3: .value = lblVfyDt.Caption
        .Row = 3:
                    If sRstv = "1" Then
                        .Col = 2: .value = IIf(sStep3 = "1", "적격", "")
                                  .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                        .Col = 3: .value = IIf(sStep3 = "1", lblVfyDt.Caption, "")
                        .Col = 4: .value = "1"
                    Else
                        .Col = 2: .value = "부적격"
                        .Col = 4: .value = sStep3
                    End If
'                    .Col = 2: .value = IIf(sRstv = "1", "적격", IIf(sStep3 = "1", "적격", "부적격"))
'                    .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
'                    .Col = 4: .value = IIf(sRstv = "1", "1", sStep3)
'                    .Col = 3: .value = lblVfyDt.Caption
        .Row = 4:
                    If sRstv = "1" Then
                        .Col = 2: .value = IIf(sStep4 = "1", "적격", "")
                                  .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                        .Col = 3: .value = IIf(sStep4 = "1", lblVfyDt.Caption, "")
                        .Col = 4: .value = "1"
                    Else
                        .Col = 2: .value = "부적격"
                        .Col = 4: .value = sStep4
                    End If
'                    .Col = 2: .value = IIf(sRstv = "1", "적격", IIf(sStep4 = "1", "적격", "부적격"))
'                    .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
'                    .Col = 4: .value = IIf(sRstv = "1", "1", sStep4)
'                    .Col = 3: .value = lblVfyDt.Caption
    End With
    
    
    '부작용등록여부 체크
    Call ReacionVisible
    If tblBeforeRealTest.tag = sPtid Then GoTo Skip
    
    Call medClearTable(tblBeforeRealTest)
    Call medClearTable(tblAfterRealTest)
    
    SSQL = objSQL.GetRealResultSQL(sPtid)
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    With tblBeforeRealTest
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                .Col = 1: .value = RS.Fields("abbrnm10").value & ""
                .Col = 2: .value = GetRstCdMatching(sPtid, RS.Fields("rstcd").value & "", RS.Fields("testcd").value & "")
                .Col = 3: .value = Format(RS.Fields("vfydt").value & "", "####-##-##") & " " & _
                                   Format(RS.Fields("vfytm").value & "", "0#:##:##")
                 RS.MoveNext
            Loop
        End If
    End With
    SSQL = objSQL.GetRealResultSQL(sPtid, Replace(lblVfyDt.Caption, "-", ""))
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    With tblAfterRealTest
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                .Col = 1: .value = RS.Fields("abbrnm10").value & ""
                .Col = 2: .value = GetRstCdMatching(sPtid, RS.Fields("rstcd").value & "", RS.Fields("testcd").value & "")
                .Col = 3: .value = Format(RS.Fields("vfydt").value & "", "####-##-##") & " " & _
                                   Format(RS.Fields("vfytm").value & "", "0#:##:##")
                 RS.MoveNext
            Loop
        End If
    End With
    
    tblBeforeRealTest.tag = sPtid
Skip:
    Me.MousePointer = 0
    Set RS = Nothing
    Set objSQL = Nothing
End Sub
Private Function GetRstCdMatching(ByVal sPtid As String, ByVal sRstcd As String, ByVal sTestcd As String) As String
    Dim objSQL  As clsTransfusion
    Dim sRS     As Recordset
    Dim SSQL    As String
    
    GetRstCdMatching = sRstcd
    Set sRS = New Recordset
    Set objSQL = New clsTransfusion
    
    SSQL = objSQL.Getlab031CdMST(LC2_ItemResult, sTestcd, sRstcd)
    sRS.Open SSQL, DBConn
    
    If Not sRS.EOF Then
        GetRstCdMatching = sRS.Fields("field1").value & ""
    End If

    Set sRS = Nothing
    Set objSQL = Nothing
End Function

Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim objSQL  As clsTransfusion
    Dim ObjABO  As clsABO
    Dim RS      As Recordset
    Dim sStscd  As String
    Dim sFDate  As String
    Dim sTDate  As String
    Dim sPtid   As String
    Dim strTmp  As String
    
    Dim SSQL    As String
    Dim sTmp    As String
    
    If Row < 1 Then Exit Sub
    
    Call Form_Clear(False)
    
    
    With tblPtList
        .Row = Row
        .Col = 2: lblptid.Caption = .value
        .Col = 3: lblPtnm.Caption = .value
        .Col = 4: strTmp = SDA_String(.value)
                
        lblSexAge.Caption = medGetP(strTmp, 1, COL_DIV) & "/" & medGetP(strTmp, 3, COL_DIV)
    End With

    sPtid = lblptid.Caption
    If sPtid = "" Then
        
        Exit Sub
    End If
    
    
    sFDate = Format(dtpQFrom.value, "YYYYMMDD")
    sTDate = Format(dtpQTo.value, "YYYYMMDD")
    
    sStscd = BBSBloodStatus.stsDELIVERY
    If optQDiv(1).value Then sStscd = BBSBloodStatus.stsEXPIRE
    
    
    Set RS = New Recordset
    Set ObjABO = New clsABO
    Set objSQL = New clsTransfusion
    
    '혈액형구하기
    ObjABO.PtId = sPtid
    If ObjABO.GetABO = True Then lblPABO.Caption = ObjABO.ABO & ObjABO.Rh
    
    SSQL = objSQL.BloodDetailQuery(sFDate, sTDate, sPtid, sStscd)
    
    RS.Open SSQL, DBConn
    
    With tblBlood
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                .Col = TblCol.TcABO:        .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
                .Col = TblCol.tcACCDT:      .value = RS.Fields("accdt").value & ""
                .Col = TblCol.tcACCSEQ:     .value = RS.Fields("accseq").value & ""
                .Col = TblCol.tcBLDNO:      .value = RS.Fields("bldsrc").value & "" & "-" & _
                                                     RS.Fields("bldyy").value & "" & "-" & _
                                                     Format(RS.Fields("bldno").value & "", "000000")
                .Col = TblCol.tcChk
                .Col = TblCol.tcCompocd:    .value = RS.Fields("compocd").value & ""
                .Col = TblCol.tcCOMPONM:    .value = RS.Fields("componm").value & ""
                .Col = TblCol.tcDELIVERYDT: .value = Format(RS.Fields("qdate").value & "", "####-##-##")
                .Col = TblCol.tcDeliveryid: .value = RS.Fields("deliveryid").value & ""
                .Col = TblCol.tcDEPTCD:     .value = RS.Fields("deptcd").value & ""
                .Col = TblCol.tcENTDT:      .value = Format(RS.Fields("entdt").value & "", "####-##-##")
                .Col = TblCol.tcEntid:      .value = RS.Fields("entid").value & ""
                .Col = TblCol.tcHOSILID:    .value = RS.Fields("hosilid").value & ""
                .Col = TblCol.tcORDDT:      .value = Format(RS.Fields("orddt").value & "", "####-##-##")
                .Col = TblCol.tcORDNO:      .value = RS.Fields("ordno").value & ""
                .Col = TblCol.tcORDSEQ:     .value = RS.Fields("ordseq").value & ""
                .Col = TblCol.tcPTID:       .value = RS.Fields("ptid").value & ""
                .Col = TblCol.tcRSTSEQ:     .value = RS.Fields("rstseq").value & ""
                .Col = TblCol.tcRSTV:       .value = RS.Fields("rstv").value & ""
                .Col = TblCol.tcSTEP1:      .value = RS.Fields("step1").value & ""
                .Col = TblCol.tcSTEP2:      .value = RS.Fields("step2").value & ""
                .Col = TblCol.tcSTEP3:      .value = RS.Fields("step3").value & ""
                .Col = TblCol.tcSTEP4:      .value = RS.Fields("step4").value & ""
                .Col = TblCol.tcVFYDT:      .value = Format(RS.Fields("vfydt").value & "", "####-##-##")
                .Col = TblCol.tcVfyid:      .value = RS.Fields("vfyid").value & ""
                .Col = TblCol.tcVOL:        .value = RS.Fields("volumn").value & "" & "cc"
                .Col = TblCol.tcWARDID:     .value = RS.Fields("wardid").value & ""
                .Col = TblCol.tcWORKAREA:   .value = RS.Fields("workarea").value & ""
                RS.MoveNext
            Loop
        End If
    End With
    
    Call tblBlood_Click(1, 1)
    
    Set RS = Nothing
    Set objSQL = Nothing
    Set ObjABO = Nothing
End Sub

Private Sub Form_Setting()
    '검사Step을 가지고 온다.
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    
    Dim strTmp  As String
    
    Dim jj      As Long
    Dim ii      As Long
    Dim lngStep As Long
    
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    
    RS.Open objSQL.GetCrossmatchingStep, DBConn
    If Not RS.EOF Then
        lngStep = Val(RS.Fields("field1").value & "")
        For ii = 1 To lngStep
            strTmp = strTmp & medGetP(RS.Fields("text1").value & "", ii, ";") & vbTab
        Next
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        With tblBeforeResult
            For jj = 1 To lngStep
                .Row = jj: .Col = 1: .value = medGetP(strTmp, jj, vbTab)
            Next
        End With
        With tblAfterResult
            For jj = 1 To lngStep
                .Row = jj: .Col = 1: .value = medGetP(strTmp, jj, vbTab)
            Next
        End With
    End If
    Set RS = Nothing
    Set objSQL = Nothing
 
End Sub

Private Sub txtReactioncd_KeyPress(KeyAscii As Integer)
    txtReactioncd.Text = UCase(txtReactioncd.Text)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtReactioncd_LostFocus()
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    Dim SSQL    As String
    
    lblReactionNm.Caption = ""
    If txtReactioncd.Text = "" Then Exit Sub
    
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    SSQL = objSQL.GetReactionSQL(txtReactioncd.Text)
    
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        lblReactionNm.Caption = RS.Fields("field1").value & ""
    Else
        txtReactioncd.Text = "": lblReactionNm.Caption = ""
    End If
    Set RS = Nothing
    
End Sub
Private Sub ReacionVisible()
    Dim objSQL      As clsTransfusion
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim sBldSrc     As String
    Dim sBldYY      As String
    Dim sBldNo      As String
    Dim sCompo      As String
    Dim sPtid       As String
    Dim ii          As Integer
    Dim jj          As Integer
    
    
    sBldSrc = medGetP(lblBloodNo.Caption, 1, "-")
    sBldYY = medGetP(lblBloodNo.Caption, 2, "-")
    sBldNo = medGetP(lblBloodNo.Caption, 3, "-")
    sCompo = medGetP(lblComponm.Caption, 1, " ")
    sPtid = lblptid.Caption
    
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    
    SSQL = objSQL.GetReactionChkSQL(sPtid, sBldSrc, sBldYY, sBldNo, sCompo)
    RS.Open SSQL, DBConn
    
    lblReactionChk.Visible = False
    With tblAfterResult
        For ii = 1 To 4
            .Row = ii
            For jj = 2 To 4
                .Col = jj: .value = ""
            Next
        Next
    End With
    dtpTransdt.value = GetSystemDate
    dtpFtm.value = GetSystemDate
    dtpTtm.value = GetSystemDate
    
    txtVolume.Text = "0"
    txtMesg = ""
    txtReactioncd.Text = ""
    lblReactionNm.Caption = ""
    
    If Not RS.EOF Then
        lblReactionChk.Visible = True
        With tblAfterResult
            .Row = 1: .Col = 2: .value = IIf(RS.Fields("step1").value & "" = "1", "적격", IIf(RS.Fields("step1").value & "" = "0", "부적격", ""))
                                .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                      .Col = 3: .value = Format(RS.Fields("reactiondt").value & "", "####-##-##") & " " & Format(RS.Fields("reactiontm").value & "", "0#:##:##")
                      .Col = 4: .value = RS.Fields("step1").value & ""
            .Row = 2: .Col = 2: .value = IIf(RS.Fields("step2").value & "" = "1", "적격", IIf(RS.Fields("step2").value & "" = "0", "부적격", ""))
                                .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                      .Col = 3: .value = Format(RS.Fields("reactiondt").value & "", "####-##-##") & " " & Format(RS.Fields("reactiontm").value & "", "0#:##:##")
                      .Col = 4: .value = RS.Fields("step2").value & ""
            .Row = 3: .Col = 2: .value = IIf(RS.Fields("step3").value & "" = "1", "적격", IIf(RS.Fields("step3").value & "" = "0", "부적격", ""))
                                .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                      .Col = 3: .value = Format(RS.Fields("reactiondt").value & "", "####-##-##") & " " & Format(RS.Fields("reactiontm").value & "", "0#:##:##")
                      .Col = 4: .value = RS.Fields("step3").value & ""
            .Row = 4: .Col = 2: .value = IIf(RS.Fields("step4").value & "" = "1", "적격", IIf(RS.Fields("step4").value & "" = "0", "부적격", ""))
                                .ForeColor = IIf(.value = "적격", vbBlue, vbRed)
                      .Col = 3: .value = Format(RS.Fields("reactiondt").value & "", "####-##-##") & " " & Format(RS.Fields("reactiontm").value & "", "0#:##:##")
                      .Col = 4: .value = RS.Fields("step4").value & ""
        End With
        txtVolume.Text = RS.Fields("volumn").value & ""
        txtReactioncd.Text = RS.Fields("reactioncd").value & ""
        lblReactionNm.Caption = RS.Fields("reactionnm").value & ""
        txtMesg.Text = RS.Fields("mesg").value & ""
        dtpTransdt.value = Format(RS.Fields("transdt").value & "", "####-##-##")
        dtpFtm.value = Format(medGetP(RS.Fields("transtm").value & "", 1, ";"), "0#:##:##")
        dtpTtm.value = Format(medGetP(RS.Fields("transtm").value & "", 2, ";"), "0#:##:##")
    End If
    Set RS = Nothing
    Set objSQL = Nothing
End Sub


Private Function SaveItemCheck() As Boolean
    If lblBloodNo.Caption = "" Then Exit Function
    If lblptid.Caption = "" Then Exit Function
    If lblComponm.Caption = "" Then Exit Function
    If txtReactioncd.Text = "" Then Exit Function
    SaveItemCheck = True
End Function

Private Sub cmdSave_Click()
    Dim objSQL  As clsTransfusion
    Dim sTmp    As String
    Dim SSQL    As String
    
    Dim sBldSrc     As String
    Dim sBldYY      As String
    Dim sBldNo      As String
    Dim sCompo      As String
    Dim sPtid       As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sRstseq     As String
    Dim sDeliverydt As String
    Dim sStep1      As String
    Dim sStep2      As String
    Dim sStep3      As String
    Dim sStep4      As String
    Dim sRstv       As String
    Dim sTransDt    As String
    Dim sTransTm    As String
    Dim sVolumn     As String
    Dim sReactionDt As String
    Dim sReactionTm As String
    Dim sReactionid As String
    Dim sReactioncd As String
    Dim sReactionNm As String
    Dim sMesg       As String
    
    
    If SaveItemCheck = False Then
        MsgBox "입력정보가 누락되었습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    If lblReactionChk.Visible = True Then
        sTmp = MsgBox("부작용등록된 혈액입니다.재등록하시겠습니까?", vbYesNo + vbInformation, "Info")
        If sTmp = vbNo Then Exit Sub
    End If
    
    sBldSrc = medGetP(lblBloodNo.Caption, 1, "-")
    sBldYY = medGetP(lblBloodNo.Caption, 2, "-")
    sBldNo = medGetP(lblBloodNo.Caption, 3, "-")
    sCompo = medGetP(lblComponm.Caption, 1, " ")
    sPtid = lblptid.Caption
    sWorkArea = medGetP(lblAccdt.Caption, 1, "-")
    sAccDt = medGetP(lblAccdt.Caption, 2, "-")
    sAccSeq = medGetP(lblAccdt.Caption, 3, "-")
    sRstseq = lblAccdt.tag
    sDeliverydt = Replace(lblDeliverydt.Caption, "-", "")
    With tblAfterResult
        .Row = 1: .Col = 4: sStep1 = .value
        .Row = 2: .Col = 4: sStep2 = .value
        .Row = 3: .Col = 4: sStep3 = .value
        .Row = 4: .Col = 4: sStep4 = .value
    End With
    
    If sStep1 = "" And sStep2 = "" And sStep3 = "" And sStep4 = "" Then
        sTmp = MsgBox("CorssMatching 결과를 입력하지 않았습니다. 등록하시겠습니까?", vbYesNo + vbInformation, "Info")
        If sTmp = vbNo Then Exit Sub
        sRstv = ""
    Else
        If sStep1 <> "1" Or sStep2 <> "1" Or sStep3 <> "1" Or sStep4 <> "1" Then
            sRstv = "0"
        Else
            sRstv = "1"
        End If
    End If
    
    sTransDt = Format(dtpTransdt.value, PRESENTDATE_FORMAT)
    sTransTm = Format(dtpFtm.value, PRESENTTIME_FORMAT) & ";" & Format(dtpTtm.value, PRESENTTIME_FORMAT)
    sVolumn = txtVolume.Text
    sReactionDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    sReactionTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    sReactionid = ObjSysInfo.EmpId
    sReactioncd = txtReactioncd.Text
    sReactionNm = lblReactionNm.Caption
    sMesg = txtMesg.Text
    
    Set objSQL = New clsTransfusion
    
    
    
On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    SSQL = objSQL.SetReactionSaveSQL("DELETE", sBldSrc, sBldYY, sBldNo, sCompo, sPtid, sWorkArea, sAccDt, sAccSeq, sRstseq, _
                                               sStep1, sStep2, sStep3, sStep4, sRstv, sTransDt, sTransTm, sVolumn, sReactionDt, _
                                               sReactionTm, sReactionid, sReactioncd, sReactionNm, sMesg, sDeliverydt)
    DBConn.Execute SSQL
    
    SSQL = objSQL.SetReactionSaveSQL("SAVE", sBldSrc, sBldYY, sBldNo, sCompo, sPtid, sWorkArea, sAccDt, sAccSeq, sRstseq, _
                                             sStep1, sStep2, sStep3, sStep4, sRstv, sTransDt, sTransTm, sVolumn, sReactionDt, _
                                             sReactionTm, sReactionid, sReactioncd, sReactionNm, sMesg, sDeliverydt)
    DBConn.Execute SSQL
    
    DBConn.CommitTrans
    
    MsgBox "정상적으로 저장되었습니다.", vbInformation + vbOKOnly, "Info"
    Set objSQL = Nothing
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    Set objSQL = Nothing
End Sub



