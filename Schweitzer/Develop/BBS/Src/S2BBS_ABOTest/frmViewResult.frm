VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViewResult 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "ABO결과 조회"
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Left            =   3390
      TabIndex        =   67
      Top             =   1500
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "※ 전체 환자를 대상으로 조회할 시 많은 시간이 소요될 수 있습니다."
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblPtinfo 
      Height          =   315
      Left            =   10470
      TabIndex        =   46
      Top             =   45
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
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
      AutoSize        =   -1  'True
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   44
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   43
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   " 환자 선택"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1155
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   8895
      Begin VB.CheckBox chkFg 
         Height          =   285
         Left            =   2670
         TabIndex        =   68
         Top             =   510
         Value           =   1  '확인
         Width           =   255
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   4230
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   330
         Left            =   6975
         TabIndex        =   3
         Top             =   465
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   5970
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   465
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "성/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   315
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "검체번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3225
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "환자명"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   1500
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "  처방 선택"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   6735
      Left            =   75
      TabIndex        =   6
      Top             =   1725
      Width           =   8895
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   420
         Left            =   7260
         Style           =   1  '그래픽
         TabIndex        =   13
         Tag             =   "15101"
         Top             =   195
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93126659
         CurrentDate     =   36948
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   315
         Left            =   3360
         TabIndex        =   9
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93126659
         CurrentDate     =   36948
      End
      Begin FPSpread.vaSpread tblOrder 
         CausesValidation=   0   'False
         Height          =   5925
         Left            =   90
         TabIndex        =   7
         Tag             =   "20001"
         Top             =   705
         Width           =   8700
         _Version        =   196608
         _ExtentX        =   15346
         _ExtentY        =   10451
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowVert    =   0   'False
         MaxCols         =   16
         MaxRows         =   24
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmViewResult.frx":0000
         UserResize      =   0
         VisibleCols     =   5
         TextTip         =   2
      End
      Begin VB.OptionButton optDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수일"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   11
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방일"
         Height          =   255
         Index           =   1
         Left            =   1275
         TabIndex        =   12
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "조회기간"
         Height          =   180
         Left            =   435
         TabIndex        =   47
         Top             =   315
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   3060
         TabIndex        =   10
         Top             =   330
         Width           =   135
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   9000
      TabIndex        =   14
      Top             =   45
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "  채혈 접수"
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1530
      Left            =   9000
      TabIndex        =   15
      Top             =   300
      Width           =   5430
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   1110
         TabIndex        =   16
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColDtTm 
         Height          =   330
         Left            =   3555
         TabIndex        =   17
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRcvNm 
         Height          =   330
         Left            =   1110
         TabIndex        =   18
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRcvDtTm 
         Height          =   330
         Left            =   3555
         TabIndex        =   19
         Top             =   660
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   2550
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   105
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "접수자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   2550
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   6
         Left            =   105
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVfyNm 
         Height          =   330
         Left            =   1110
         TabIndex        =   63
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVfyDtTm 
         Height          =   330
         Left            =   3555
         TabIndex        =   64
         Top             =   1080
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   15
         Left            =   2550
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "보고일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   16
         Left            =   105
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "보고자"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   315
      Left            =   9000
      TabIndex        =   20
      Top             =   1860
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "  검사 결과"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   1530
      Left            =   9000
      TabIndex        =   21
      Top             =   2115
      Width           =   5430
      Begin FPSpread.vaSpread tblResult 
         CausesValidation=   0   'False
         Height          =   1035
         Left            =   75
         TabIndex        =   26
         Tag             =   "20001"
         Top             =   405
         Width           =   5250
         _Version        =   196608
         _ExtentX        =   9260
         _ExtentY        =   1826
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   4
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmViewResult.frx":0999
         UserResize      =   0
         VisibleCols     =   3
      End
      Begin VB.Label lblSpcNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체:"
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
         Left            =   1050
         TabIndex        =   40
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체:"
         Height          =   180
         Left            =   75
         TabIndex        =   39
         Top             =   195
         Width           =   420
      End
   End
   Begin MedControls1.LisLabel LisLabel9 
      Height          =   315
      Left            =   9000
      TabIndex        =   22
      Top             =   3705
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "  세부 결과"
      Appearance      =   0
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   2610
      Left            =   9000
      TabIndex        =   23
      Top             =   3960
      Width           =   5430
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   11
         Left            =   480
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2190
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
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
         Caption         =   "검사일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   12
         Left            =   480
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1845
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
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
         Caption         =   "검사자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   9
         Left            =   480
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
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
         Caption         =   "RH Subgroup"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   10
         Left            =   480
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
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
         Caption         =   "ABO Subtype"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   345
         Index           =   7
         Left            =   480
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   795
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
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
         Caption         =   "RH Result"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   8
         Left            =   480
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   450
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
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
         Caption         =   "ABO Result"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel11 
         Height          =   315
         Left            =   1935
         TabIndex        =   27
         Top             =   120
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         BackColor       =   16512
         ForeColor       =   -2147483634
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
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel12 
         Height          =   315
         Left            =   3435
         TabIndex        =   28
         Top             =   120
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         BackColor       =   16512
         ForeColor       =   -2147483634
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
         Caption         =   "Back Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABOFront 
         Height          =   330
         Left            =   1935
         TabIndex        =   29
         Top             =   450
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   12632256
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
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABOBack 
         Height          =   330
         Left            =   3435
         TabIndex        =   30
         Top             =   450
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   12632256
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
         Caption         =   "Back Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRHFront 
         Height          =   345
         Left            =   1935
         TabIndex        =   31
         Top             =   795
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         BackColor       =   12632256
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
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRHBack 
         Height          =   345
         Left            =   3435
         TabIndex        =   32
         Top             =   795
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         BackColor       =   12632256
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
         Caption         =   "Back Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABOSub 
         Height          =   330
         Left            =   1935
         TabIndex        =   33
         Top             =   1155
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BackColor       =   12632256
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
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRHSub 
         Height          =   330
         Left            =   1935
         TabIndex        =   34
         Top             =   1500
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BackColor       =   12632256
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
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVfyNmFront 
         Height          =   330
         Left            =   1935
         TabIndex        =   35
         Top             =   1845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   12632256
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
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVfyNmBack 
         Height          =   330
         Left            =   3435
         TabIndex        =   36
         Top             =   1845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   12632256
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
         Caption         =   "Back Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVfyDtTmFront 
         Height          =   330
         Left            =   1935
         TabIndex        =   37
         Top             =   2190
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "Front Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVfyDtTmBack 
         Height          =   330
         Left            =   3435
         TabIndex        =   38
         Top             =   2190
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "Back Typing"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblModifyFg 
         Height          =   315
         Left            =   1230
         TabIndex        =   45
         Top             =   120
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         BackColor       =   4194304
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
         Alignment       =   1
         Caption         =   "수정"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel10 
      Height          =   315
      Left            =   9000
      TabIndex        =   24
      Top             =   6570
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "  Remark / Comment"
      Appearance      =   0
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00DBE6E6&
      Height          =   1650
      Left            =   9000
      TabIndex        =   25
      Top             =   6810
      Width           =   5430
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   13
         Left            =   75
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "Comment"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   75
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin VB.TextBox txtComment 
         BackColor       =   &H00DBE6E6&
         Height          =   1020
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   4260
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00DBE6E6&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   150
         Width           =   4260
      End
   End
End
Attribute VB_Name = "frmViewResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'처방일자, 번호, 검사명,검체, 응급, 환자id, 환자명, 상태, 소요시간, wa, vAccdt, vAccseq, ordtm, 접수일자,처방의
Private Enum TblColumn
    tcORDDT = 1
    tcORDNO
    tcTESTNM
    tcSPCNM
    tcSTAT
    
    tcPTID
    tcPTNM
    tcSTATUS
    tcDELAYTIME
    tcWORKAREA
    
    tcACCDT
    tcACCSEQ
    tcORDTM
    tcRCVDTTM
    tcDOCT
    
    tcORDCD
End Enum

Private IsFirst As Boolean
Private onPgm As Boolean

Public Event FormClose()

Private Sub chkFg_Click()
    If chkFg.Value = 0 Then
        lbldt(0).Caption = "환자ID"
    Else
        lbldt(0).Caption = "검체번호"
    End If
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub cmdQuery_Click()
    Call QueryOrder
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    dtpToDt = GetSystemDate
    dtpFrDt = GetSystemDate
    optDay(0).Visible = False
    optDay(1).Visible = False
    
    optDay(1).Value = True
    ClearAll
End Sub

Private Sub Form_Load()
    IsFirst = True
End Sub

Private Sub tblOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vWorkarea As String
    Dim vAccdt As String
    Dim vAccseq As String
    Dim spcnm As String
    Dim ptid As String
    Dim ordcd As String
    
    If onPgm Then Exit Sub
    
    If Row < 1 Then Exit Sub
    
    Call Clear3
    Call Clear4
    Call Clear5
    Call Clear6
    
    With tblOrder
        If Row > .DataRowCnt Then Exit Sub
        
        .Row = Row
        .Col = TblColumn.tcWORKAREA:    vWorkarea = .Value
        .Col = TblColumn.tcACCDT:       vAccdt = .Value
        .Col = TblColumn.tcACCSEQ:      vAccseq = .Value
        .Col = TblColumn.tcSPCNM:       spcnm = .Value
        .Col = TblColumn.tcPTID:        ptid = .Value
        .Col = TblColumn.tcORDCD: ordcd = .Value
        lblSpcNm.Caption = spcnm
        
        Call QueryAccInfo(vWorkarea, vAccdt, vAccseq)
        Call QueryOrdInfo(ptid, vWorkarea, vAccdt, vAccseq, ordcd)
        Call QueryResult(vWorkarea, vAccdt, vAccseq)
        Call QueryDetailResult(vWorkarea, vAccdt, vAccseq)
    End With
End Sub

Private Sub tblOrder_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim vWorkarea As String
'    Dim vAccdt As String
'    Dim vAccseq As String
'    Dim spcnm As String
'    Dim ptid As String
'
'    If onPgm Then Exit Sub
'
'    If Row = NewRow Then Exit Sub
'    If NewRow < 1 Then Exit Sub
'
'    Call Clear3
'    Call Clear4
'    Call Clear5
'    Call Clear6
'
'    With tblOrder
'        If NewRow > .DataRowCnt Then Exit Sub
'
'        .Row = Row
'        .Col = TblColumn.tcWORKAREA:    vWorkarea = .Value
'        .Col = TblColumn.tcACCDT:       vAccdt = .Value
'        .Col = TblColumn.tcACCSEQ:      vAccseq = .Value
'        .Col = TblColumn.tcSPCNM:       spcnm = .Value
'        .Col = TblColumn.tcPTID:        ptid = .Value
'        lblSpcNm.Caption = spcnm
''        Dim objPtnt As New clsPatient
''        Dim QueryPtnt As Boolean
'
'        Call QueryAccInfo(vWorkarea, vAccdt, vAccseq)
'        Call QueryResult(vWorkarea, vAccdt, vAccseq)
'        Call QueryDetailResult(vWorkarea, vAccdt, vAccseq)
''        QueryPtnt = objPtnt.PtntQuery(ptid)
''        lblPtinfo.Caption = "(" & objPtnt.ptnm & "[" & ptid & "]," & objPtnt.SexAge & ")"
''        Set objPtnt = Nothing
'    End With
End Sub

Private Sub tblOrder_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim orddttm As String
    Dim ordno As String
    Dim testnm As String
    Dim spcnm As String
    Dim doct As String
    Dim accno As String
    
'    MsgBox TipWidth
    
    MultiLine = 1
    
    With tblOrder
        If Row < 1 Then Exit Sub
        If Row > .DataRowCnt Then Exit Sub
        
        .Row = Row
        
        .Col = TblColumn.tcORDDT:       orddttm = Mid(.Value, 3)
        .Col = TblColumn.tcORDTM:       orddttm = orddttm & " " & .Value
        .Col = TblColumn.tcORDNO:       ordno = .Value
        .Col = TblColumn.tcTESTNM:      testnm = .Value
        .Col = TblColumn.tcSPCNM:       spcnm = .Value
        .Col = TblColumn.tcDOCT:        doct = .Value
        .Col = TblColumn.tcWORKAREA:    accno = .Value
        .Col = TblColumn.tcACCDT:       accno = accno & "-" & Mid(.Value, 3)
        .Col = TblColumn.tcACCSEQ:      accno = accno & "-" & .Value
        
        TipText = ""
        TipText = TipText & vbNewLine & "   처방일시 : " & orddttm
        TipText = TipText & vbNewLine & "   처방번호 : " & ordno
        TipText = TipText & vbNewLine & "   검 사 명 : " & testnm
        TipText = TipText & vbNewLine & "   검    체 : " & spcnm
        TipText = TipText & vbNewLine & "   처 방 의 : " & doct
        TipText = TipText & vbNewLine & "   접수번호 : " & accno
        TipText = TipText & vbNewLine
        
        TipWidth = 4000
        
        Call .SetTextTipAppearance("굴림체", 9, False, False, &HC0C0C0, RGB(0, 0, 0))
        
    End With
    
    ShowTip = True
End Sub

Private Sub txtPtId_Change()
'    If txtPtId = "" Then ClearAll
End Sub

Private Sub txtPtid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim ii As Integer
        
        If chkFg.Value = 0 Then
            txtPtId.Text = Format(txtPtId.Text, String(BBS_PTID_LENGTH, "0"))
            If QueryPtnt = True Then Call QueryOrder
        Else
            Call QueryOrder
        End If
    End If
End Sub

Private Function QueryPtnt() As Boolean
    Dim objPtnt As clsPatient
    Dim ii      As Integer
    
    Call Clear2
    Call Clear3
    Call Clear4
    Call Clear5
    Call Clear6
    
    Set objPtnt = New clsPatient
    
    txtPtId.Text = Format(txtPtId.Text, String(BBS_PTID_LENGTH, "0"))
    
    QueryPtnt = objPtnt.GETPatient(txtPtId.Text)
    
    lblPtNm.Caption = objPtnt.ptnm
    lblSexAge.Caption = objPtnt.SEXAGE
    
    Set objPtnt = Nothing
End Function

Private Function QueryOrder() As Boolean
    Dim FrDt As String
    Dim ToDt As String
    Dim ptid As String
    Dim dayfg As String
    
    Dim i As Long
    
    Dim Rs As Recordset
    Dim objABOSql As clsABOSql
    
    Dim orddt As String
    Dim ordno As String
    Dim testnm As String
    Dim spcnm As String
    Dim stat As String
    Dim doct As String
    Dim rcvdttm As String
    Dim status As String
    Dim delaytime As String
    Dim vWorkarea As String
    Dim vAccdt As String
    Dim vAccseq As String
    Dim ordtm As String
    Dim ii  As Integer
    Dim strLng As String
    
    QueryOrder = False
    
    Call Clear2
    Call Clear3
    Call Clear4
    Call Clear5
    Call Clear6
    
    txtPtId.Text = Format(txtPtId.Text, String(BBS_PTID_LENGTH, "0"))
    
    
    ptid = txtPtId.Text
    FrDt = Format(dtpFrDt, "YYYYMMDD")
    ToDt = Format(dtpToDt, "YYYYMMDD")
    
    If optDay(0).Value Then
        dayfg = 0
    ElseIf optDay(1).Value Then
        dayfg = 1
    Else
        dayfg = 0
    End If
    
    Me.MousePointer = 11

    Dim objPrgBar As New clsProgress
    
    objPrgBar.Container = MainFrm.stsbar
    
    Set objABOSql = New clsABOSql
    
    If chkFg.Value = 1 Then
        Set Rs = objABOSql.LoadABOOrdList_New("", FrDt, ToDt, dayfg, ptid)
    Else
        Set Rs = objABOSql.LoadABOOrdList_New(ptid, FrDt, ToDt, dayfg, "")
    End If
    Set objABOSql = Nothing
    If Rs Is Nothing Then Exit Function
    
    With Rs
        If .RecordCount < 1 Then
            MsgBox "처방을 찾을 수 없습니다.", vbCritical, Me.Caption
        Else
            objPrgBar.max = .RecordCount
            If .RecordCount < 19 Then
                tblOrder.MaxRows = 19
            Else
                tblOrder.MaxRows = .RecordCount
            End If
            tblOrder.RowHeight(-1) = 12
            For i = 1 To .RecordCount
                objPrgBar.Value = i
                orddt = Format(.Fields("orddt").Value & "", "0###-0#-##")
                ordno = .Fields("ordno").Value & ""
                testnm = .Fields("testnm").Value & ""
                spcnm = .Fields("spcnm").Value & ""
                stat = IIf(.Fields("statfg").Value & "" = "1", "Y", "")
                status = statusnm(.Fields("stscd").Value & "")
                delaytime = ""
                vWorkarea = .Fields("workarea").Value & ""
                vAccdt = .Fields("accdt").Value & ""
                vAccseq = .Fields("accseq").Value & ""
                ordtm = Format(Mid(.Fields("ordtm").Value & "", 1, 4), "0#:0#")
                rcvdttm = Format(.Fields("rcvdt").Value & "", "0###-0#-0#") & " " & Format(Mid(.Fields("rcvtm").Value & "", 1, 4), "0#:0#")
                doct = GetDoctNm(.Fields("orddoct").Value & "")
                
                With tblOrder
                    .Row = i
'                    If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
                    
                    .Col = TblColumn.tcORDDT:       .Value = orddt
                    .Col = TblColumn.tcORDNO:       .Value = ordno
                    .Col = TblColumn.tcTESTNM:      .Value = testnm
                    .Col = TblColumn.tcSPCNM:       .Value = spcnm
                    .Col = TblColumn.tcSTAT:        .Value = stat
                    .Col = TblColumn.tcPTID: .Value = Rs.Fields("ptid").Value & ""
                    Dim objPtnt As New clsPatient
                    Call objPtnt.GETPatient(Rs.Fields("ptid").Value & "")
                    .Col = TblColumn.tcPTNM: .Value = objPtnt.ptnm: .ForeColor = DCM_Blue
                    Set objPtnt = Nothing
                    .Col = TblColumn.tcSTATUS:      .Value = status
                                                    Select Case Rs.Fields("stscd").Value & ""
                                                        Case "0": .ForeColor = vbBlack
                                                        Case "1": .ForeColor = vbBlack
                                                        Case "2": .ForeColor = vbGreen
                                                        Case "3": .ForeColor = DCM_LightRed
                                                        Case "4": .ForeColor = DCM_LightBlue
                                                        Case "5": .ForeColor = vbRed
                                                        Case "6":
                                                    End Select
                    .Col = TblColumn.tcDELAYTIME:   .Value = delaytime
                    .Col = TblColumn.tcWORKAREA:    .Value = vWorkarea
                    .Col = TblColumn.tcACCDT:       .Value = vAccdt
                    .Col = TblColumn.tcACCSEQ:      .Value = vAccseq
                    .Col = TblColumn.tcORDTM:       .Value = ordtm
                    .Col = TblColumn.tcRCVDTTM:     .Value = rcvdttm
                    .Col = TblColumn.tcDOCT:        .Value = doct
                    .Col = TblColumn.tcORDCD: .Value = Rs.Fields("ordcd").Value & ""
                End With
                .MoveNext
            Next i
        End If
    End With
    Me.MousePointer = 0
    Set objPrgBar = Nothing
    Set Rs = Nothing
    txtPtId.Text = ""
End Function

Private Function statusnm(ByVal stscd As String) As String
    Select Case stscd
        Case "0": statusnm = "처방"
        Case "1": statusnm = "채혈"
        Case "2": statusnm = "접수"
        Case "3": statusnm = "검사중"
        Case "4": statusnm = "중간"
        Case "5": statusnm = "결과"
        Case "6": statusnm = "수정"
        Case Else:
                  statusnm = ""
    End Select
End Function

Private Function QueryAccInfo(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    
    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.GetAccessInfo(vWorkarea, vAccdt, vAccseq)
    
    QueryAccInfo = False
    If Rs Is Nothing Then Exit Function
    
    With Rs
        If .RecordCount > 0 Then
            
            lblColDtTm.Caption = Format(.Fields("coldt").Value & "", "0###-##-##") & " " & Format(Mid(.Fields("coltm").Value & "", 1, 4), "0#:##")
            lblColNm.Caption = GetEmpNm(.Fields("colid").Value & "")
            lblRcvDtTm.Caption = Format(.Fields("rcvdt").Value & "", "0###-##-##") & " " & Format(Mid(.Fields("rcvtm").Value & "", 1, 4), "0#:##")
            lblRcvNm.Caption = GetEmpNm(.Fields("rcvid").Value & "")
            
            txtRemark.Text = objABOSql.GetRemarkNm(.Fields("rmkcd").Value & "")
        End If
    End With
    Set Rs = Nothing
    
    Set Rs = objABOSql.GetAccComment(vWorkarea, vAccdt, vAccseq)
    If Rs Is Nothing Then Exit Function
    
    With Rs
        If .RecordCount > 0 Then
            txtComment.Text = .Fields("rsttxt").Value & ""
        End If
    End With
    Set Rs = Nothing
    
    Set objABOSql = Nothing
    
    QueryAccInfo = True
End Function

Private Function QueryOrdInfo(ByVal ptid As String, ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, ByVal ordcd As String) As Boolean
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    
    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.GetOrderInfo(ptid, vWorkarea, vAccdt, vAccseq, ordcd)
    
    QueryOrdInfo = False
    If Rs Is Nothing Then Exit Function
    
    With Rs
        If .RecordCount > 0 Then
            If .Fields("examdoct").Value & "" <> "" Then
                lblVfyDtTm.Caption = Format(.Fields("examdt").Value & "", "0###-##-##") & " " & Format(Mid(.Fields("examtm").Value & "", 1, 4), "0#:##")
                lblVfyNm.Caption = GetEmpNm(.Fields("examdoct").Value & "")
            End If
        End If
    End With
    Set Rs = Nothing
    Set objABOSql = Nothing
End Function

Private Function QueryResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    Dim i As Long
    
    Dim testnm As String
    Dim result As String
    Dim unit As String
    
    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.GetABOResult(vWorkarea, vAccdt, vAccseq)
    Set objABOSql = Nothing
    
    QueryResult = False
    If Rs Is Nothing Then Exit Function
    
    With Rs
        If .RecordCount > 0 Then
            If .RecordCount < 4 Then
                tblResult.MaxRows = 4
            Else
                tblResult.MaxRows = .RecordCount
            End If
            tblResult.RowHeight(-1) = 11
            For i = 1 To .RecordCount
                testnm = .Fields("testnm").Value & ""
                If .Fields("rstcdnm").Value & "" = "" Then
                    result = .Fields("rstcd").Value & ""
                Else
                    result = .Fields("rstcdnm").Value & ""
                End If
                
                unit = .Fields("rstunit").Value & ""
                
                With tblResult
                    .Row = i
'                    If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
                    
                    .Col = 1: .Value = testnm
                    .Col = 2: .Value = result
                    .Col = 3: .Value = unit
                End With
                .MoveNext
            Next i
        End If
    End With
    Set Rs = Nothing
End Function

Private Function QueryDetailResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    
    
    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.GetABOResultInfo(vWorkarea, vAccdt, vAccseq)
    Set objABOSql = Nothing
    
    QueryDetailResult = False
    If Rs Is Nothing Then Exit Function
        
    With Rs
        If .RecordCount > 0 Then
            If .Fields("seq").Value & "" > 0 Then lblModifyFg.Visible = True
            lblABOFront.Caption = .Fields("abo1").Value & ""
            lblABOBack.Caption = .Fields("abo2").Value & ""
            lblRHFront.Caption = .Fields("rh1").Value & ""
            lblRHBack.Caption = .Fields("rh2").Value & ""
            lblABOSub.Caption = .Fields("abosub").Value & ""
            lblRhSub.Caption = .Fields("rhsub").Value & ""
            
            lblVfyDtTmFront.Caption = Format(.Fields("vfydt1").Value & "", "####-##-##") & " " & Format(Mid(.Fields("vfytm1").Value & "", 1, 4), "##:##")
            lblVfyNmFront.Caption = GetEmpNm(.Fields("vfyid1").Value & "")
            
            lblVfyDtTmBack.Caption = Format(.Fields("vfydt2").Value & "", "####-##-##") & " " & Format(Mid(.Fields("vfytm2").Value & "", 1, 4), "##:##")
            lblVfyNmBack.Caption = GetEmpNm(.Fields("vfyid2").Value & "")
        End If
    End With
    Set Rs = Nothing
End Function

Private Sub ClearAll()
    Call Clear1
    Call Clear2
    Call Clear3
    Call Clear4
    Call Clear5
    Call Clear6
End Sub


Private Sub Clear1()
    txtPtId = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
End Sub

Private Sub Clear2()
    medClearTable tblOrder
    tblOrder.MaxRows = 19
End Sub

Private Sub Clear3()
    lblColNm.Caption = ""
    lblColDtTm.Caption = ""
    lblRcvNm.Caption = ""
    lblRcvDtTm.Caption = ""
    lblVfyNm.Caption = ""
    lblVfyDtTm.Caption = ""
End Sub

Private Sub Clear4()
    lblSpcNm.Caption = ""
    medClearTable tblResult
    tblResult.MaxRows = 4
End Sub

Private Sub Clear5()
    lblModifyFg.Visible = False
    
    lblABOFront.Caption = ""
    lblABOBack.Caption = ""
    lblRHFront.Caption = ""
    lblRHBack.Caption = ""
    lblABOSub.Caption = ""
    lblRhSub.Caption = ""
    lblVfyNmFront.Caption = ""
    lblVfyNmBack.Caption = ""
    lblVfyDtTmFront.Caption = ""
    lblVfyDtTmBack.Caption = ""
End Sub

Private Sub Clear6()
    txtRemark = ""
    txtComment = ""
End Sub

