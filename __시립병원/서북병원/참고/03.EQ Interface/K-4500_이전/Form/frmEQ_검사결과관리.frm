VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEQ_검사결과관리 
   Caption         =   "검사결과관리"
   ClientHeight    =   9375
   ClientLeft      =   8160
   ClientTop       =   3390
   ClientWidth     =   13155
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ_검사결과관리.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   13155
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6660
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   2760
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "전송(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11220
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   60
      Width           =   915
   End
   Begin VB.ComboBox cboSENDFLAG 
      Height          =   300
      Left            =   4020
      Style           =   2  '드롭다운 목록
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox cboSTATEFLAG 
      Height          =   300
      Left            =   4020
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtPATNM 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Text            =   "1234567890"
      Top             =   2760
      Width           =   1035
   End
   Begin VB.TextBox txtPATNO 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Text            =   "1234567890"
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "[기준일자]"
      Height          =   915
      Left            =   60
      TabIndex        =   23
      Top             =   1020
      Width           =   7515
      Begin VB.OptionButton optDateSection 
         Caption         =   "처방일자"
         Height          =   180
         Index           =   3
         Left            =   6360
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optDateSection 
         Caption         =   "검사결과전송일자"
         Height          =   180
         Index           =   2
         Left            =   4500
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optDateSection 
         Caption         =   "검사결과수신일자"
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optDateSection 
         Caption         =   "검사처방전송일자"
         Height          =   180
         Index           =   0
         Left            =   780
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66912257
         CurrentDate     =   40820
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   2340
         TabIndex        =   4
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66912257
         CurrentDate     =   40820
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "기간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Index           =   8
         Left            =   2160
         TabIndex        =   25
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "구분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtBARCD 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Text            =   "201101011234567"
      Top             =   2040
      Width           =   1575
   End
   Begin FPSpread.vaSpread sprDResult 
      Height          =   6195
      Left            =   7620
      TabIndex        =   12
      Top             =   3120
      Width           =   5475
      _Version        =   393216
      _ExtentX        =   9657
      _ExtentY        =   10927
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      SpreadDesigner  =   "frmEQ_검사결과관리.frx":263A
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12180
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "조회(&V)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5700
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   2760
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   15
      Top             =   600
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FPSpread.vaSpread sprLResult 
      Height          =   6195
      Left            =   60
      TabIndex        =   33
      Top             =   3120
      Width           =   7515
      _Version        =   393216
      _ExtentX        =   13256
      _ExtentY        =   10927
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   20
      SpreadDesigner  =   "frmEQ_검사결과관리.frx":42DF
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "Sample No"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   7680
      TabIndex        =   55
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label lblSAMPLENO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   54
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "(Like)"
      Height          =   180
      Index           =   22
      Left            =   2040
      TabIndex        =   53
      Top             =   2820
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "병록 번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   11040
      TabIndex        =   52
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "수검자 명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   11040
      TabIndex        =   51
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "성별/연령"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   11040
      TabIndex        =   50
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "처방 전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   7680
      TabIndex        =   49
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "처방 일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   11040
      TabIndex        =   48
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "처방 종류"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   11040
      TabIndex        =   47
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblEXDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   46
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label lblPATNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   45
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label lblPATNM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   44
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label lblSEXAGE 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   43
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label lblORDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   42
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label lblORDGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   41
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "결과 수신"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   7680
      TabIndex        =   40
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label lblRCDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   39
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label lblSDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   38
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "결과 전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   7680
      TabIndex        =   37
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "Rack/Pos"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   7680
      TabIndex        =   36
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblDISKNOPOSNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   35
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label lblEXSEQ 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   12000
      TabIndex        =   32
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "검사 회차"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   11040
      TabIndex        =   31
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "검체 번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   7680
      TabIndex        =   30
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label lblBARCD 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8640
      TabIndex        =   29
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체번호별 세부정보"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   7740
      TabIndex        =   28
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체번호별 검사결과"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   7740
      TabIndex        =   27
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "전송상태"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   3180
      TabIndex        =   22
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "결과상태"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   3180
      TabIndex        =   21
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "수검자명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "병록번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "검체번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   2100
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사리스트"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl장비명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사결과관리"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   16
      Top             =   60
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   2595
   End
   Begin VB.Shape shpDResult 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Left            =   7620
      Shape           =   4  '둥근 사각형
      Top             =   2820
      Width           =   5475
   End
   Begin VB.Shape shpDInfo 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Left            =   7620
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   5475
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Index           =   0
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   7515
   End
End
Attribute VB_Name = "frmEQ_검사결과관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height의 초기값
Dim lngMeWidth      As Long '/Me.Width의 초기값

Private Type ConWhere   ' 사용자 정의 형식을 만듭니다.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Public Sub SUB_MM_CANCEL()
    barStatus.Max = 100
    barStatus.Value = 100
    
    txtBARCD = ""
    txtPATNO = ""
    txtPATNM = ""
    cboSTATEFLAG.ListIndex = -1
    cboSENDFLAG.ListIndex = -1
    
    If sprLResult.MaxRows > 0 Then sprLResult.MaxRows = 0
    
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Public Function FUNC_MM_DELETE() As Boolean
    FUNC_MM_DELETE = False
    
    Dim intActCol    As Integer
    Dim intActRow    As Integer
    
    '/1.삭제 조건 Check
    If sprVIEW.ActiveRow = 0 Then MsgBox "삭제할 내용을 선택하십시오", vbInformation, "확인": Exit Function
    
    '/2.삭제 질의
    If MsgBox("장비검사코드 : " & GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow) & vbCrLf & _
              "장비검사명   : " & GET_CELL(sprVIEW, 2, sprVIEW.ActiveRow) & vbCrLf & vbCrLf & _
              "위 자료를 삭제하겠습니까?", vbQuestion + vbOKCancel, "삭제질의") = vbCancel Then Exit Function
    
    '/3.Process
    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
    If sprVIEW.IsBlockSelected Then
        intActCol = sprVIEW.SelBlockCol
        intActRow = sprVIEW.SelBlockRow
    Else
        intActCol = sprVIEW.ActiveCol
        intActRow = sprVIEW.ActiveRow
    End If
    If sprVIEW.IsBlockSelected Then
        For intX = sprVIEW.SelBlockRow To sprVIEW.SelBlockRow2
            gstrQuy = "DELETE FROM EQ_MST "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & GET_CELL(sprVIEW, 1, intX) & "' "
            If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
        Next intX
    Else
        gstrQuy = "DELETE FROM EQ_MST "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow) & "' "
        If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    End If
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_MM_DELETE = True
    
    MsgBox "삭제되었습니다!", vbInformation, "확인"
    
    '/4.화면처리
    Call FUNC_MM_VIEW_LIST
    sprVIEW.Col = intActCol
    sprVIEW.Row = intActRow
    sprVIEW.Action = ActionActiveCell
End Function

Private Sub SUB_MM_INITIAL()
    '/Form Resize를 위한 컨트롤 초기값 읽기
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Resize를 위한 초기값 설정
    lngMeHeight = 9855
    lngMeWidth = 13275
    
    '/화면 가운데 위치
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    '''Me.Show
    
    GoSub ADD_ITEM
    
    optDateSection(0).Value = True  '/기준일자/구분:처방일자
    dtpDateFrom.Value = Date        '/기간From
    dtpDateTo.Value = Date          '/기간To
    
    Call SUB_MM_CANCEL
Exit Sub

'/------------------------------------------------------------------------------------------/

ADD_ITEM:
    '/결과진행상태 (0.처방, 1.결과)
    cboSTATEFLAG.AddItem ""
    cboSTATEFLAG.AddItem "0.처방"
    cboSTATEFLAG.AddItem "1.결과"
    
    '/HIS 전송 FLAG (0.대기, 1.완료)
    cboSENDFLAG.AddItem ""
    cboSENDFLAG.AddItem "0.대기"
    cboSENDFLAG.AddItem "1.완료"
Return
End Sub

Public Sub SUB_MM_INPUT()
    gstrInputUpdate = "1" '/1.Input, 2.Update
    gstrInputUpdateYN = False

    frmEQ공용_장비검사코드관리_입력.Show vbModal

    If gstrInputUpdateYN = True Then
        Call FUNC_MM_VIEW_LIST
    End If
End Sub

Private Sub SUB_MM_KEY_CLEAR(ArgSection As String) '/ArgSection: 1.검사리스트, 2.검체번호별
    If ArgSection = "1" Then
        If sprLResult.MaxRows > 0 Then sprLResult.MaxRows = 0 '/검사리스트
    End If
    
    lblBARCD = ""       '/검체번호
    lblEXSEQ = ""       '/검사회차
    lblDISKNOPOSNO = "" '/Rack/Pos
    lblEXDT = ""        '/검사처방전송일자
    lblRCDT = ""        '/검사결과수신일자
    lblSDDT = ""        '/검사결과전송일자
    lblPATNO = ""       '/병록번호
    lblPATNM = ""       '/수검자명
    lblORDDT = ""       '/처방일자
    lblSEXAGE = ""      '/성별/연령
    lblORDGB = ""       '/입/외구분
    
    If sprDResult.MaxRows > 0 Then sprDResult.MaxRows = 0 '/검체번호별 검사결과
        
End Sub

Public Sub SUB_MM_UPDATE()
    Dim intActCol    As Integer
    Dim intActRow    As Integer
    
    If sprVIEW.ActiveRow = 0 Then MsgBox "수정할 대상을 선택하십시오!", vbInformation, "확인": Exit Sub
    
    gstrInputUpdate = "2" '/1.Input, 2.Update
    gstrInputUpdateYN = False
    gstrArgTemp1 = GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow)
    
    frmEQ공용_장비검사코드관리_입력.Show vbModal
    
    If gstrInputUpdateYN = True Then
        intActCol = sprVIEW.ActiveCol
        intActRow = sprVIEW.ActiveRow

        Call FUNC_MM_VIEW_LIST
    
        sprVIEW.Col = intActCol
        sprVIEW.Row = intActRow
        sprVIEW.Action = ActionActiveCell
    End If
End Sub

Public Function FUNC_MM_VIEW_LIST() As Boolean
    FUNC_MM_VIEW_LIST = False
    
On Error GoTo RTN_ERR
    
    Call SUB_MM_KEY_CLEAR("1")
    
    If ConnDB_LOC = False Then Exit Function
    
    With sprLResult
        gstrQuy = "SELECT BARCD, EXSEQ, SAMPLENO, DISKNO, POSNO, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(STATEFLAG) AS STATEFLAG, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(SENDFLAG)  AS SENDFLAG, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(EXDT+' '+EXTM) AS EXDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(RCDT+' '+RCTM) AS RCDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(SDDT+' '+SDTM) AS SDDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(ORDDT)     AS ORDDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(ORDGB)     AS ORDGB, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATNO)     AS PATNO, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATNM)     AS PATNM, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATSEX)    AS PATSEX, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATAGE)    AS PATAGE "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
        Select Case True
            Case optDateSection(0).Value '/검사처방전송일자
                gstrQuy = gstrQuy & vbCrLf & " WHERE EXDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND EXDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
            
            Case optDateSection(1).Value '/검사결과수신일자
                gstrQuy = gstrQuy & vbCrLf & " WHERE RCDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND RCDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
            
            Case optDateSection(2).Value '/검사결과전송일자
                gstrQuy = gstrQuy & vbCrLf & " WHERE SDDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND SDDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
            
            Case optDateSection(3).Value '/처방일자
                gstrQuy = gstrQuy & vbCrLf & " WHERE ORDDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND ORDDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
        
        End Select
        
        '/검체번호
        If Trim(txtBARCD) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND BARCD = '" & Trim(txtBARCD) & "' "
        End If
        
        '/병록번호
        If Trim(txtPATNO) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND PATNO = '" & Trim(txtPATNO) & "' "
        End If
        
        '/수검자명
        If Trim(txtPATNM) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND PATNM LIKE '%" & Trim(txtPATNM) & "%' "
        End If
        
        '/결과진행상태 (0:처방, 1:결과)
        If Trim(cboSTATEFLAG) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND STATEFLAG = '" & Trim(Left(cboSTATEFLAG, 1)) & "' "
        End If
        
        '/HIS 전송 FLAG (0:대기, 1:완료)
        If Trim(cboSENDFLAG) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND SENDFLAG = '" & Trim(Left(cboSENDFLAG, 1)) & "' "
        End If
        
        gstrQuy = gstrQuy & vbCrLf & " GROUP BY BARCD, EXSEQ, SAMPLENO, DISKNO, POSNO "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY BARCD, EXSEQ, SAMPLENO, DISKNO, POSNO "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
        
        If Not ADR_LOC Is Nothing Then
            .MaxRows = ARC_LOC
            barStatus.Max = ARC_LOC
            intX = 0
            
            Do Until ADR_LOC.EOF
                intX = intX + 1: .Row = intX: barStatus.Value = intX
                
                .Col = 2: .Text = Trim(ADR_LOC!BARCD & "")     '/검체번호(Barcode)
                .Col = 3: .Text = Trim(ADR_LOC!EXSEQ & "")     '/검체번호(Barcode)별 검사회차
                .Col = 4: .Text = Trim(ADR_LOC!SAMPLENO & "")  '/Sample No
                .Col = 5: .Text = Trim(ADR_LOC!DISKNO & "")    '/디스크번호 or 렉번호
                .Col = 6: .Text = Trim(ADR_LOC!POSNO & "")     '/위치번호
                
                .Col = 7                                        '/결과진행상태 (0:처방, 1:결과)
                Select Case Trim(ADR_LOC!STATEFLAG & "")
                    Case "0": .Text = "처방"
                    Case "1": .Text = "결과"
                End Select
                
                .Col = 8                                        '/HIS 전송 FLAG (0:대기, 1:완료)
                Select Case Trim(ADR_LOC!SENDFLAG & "")
                    Case "0": .Text = "대기"
                    Case "1": .Text = "완료"
                End Select
                
                .Col = 9: '/검사처방전송일자
                If Trim(ADR_LOC!EXDT & "") <> "" Then
                    .Text = Format(Left(Trim(ADR_LOC!EXDT & ""), 8), "@@@@-@@-@@") & " " & Format(Mid(Trim(ADR_LOC!EXDT & ""), 10), "@@:@@:@@")
                End If
                .Col = 10 '/검사결과수신일자
                If Trim(ADR_LOC!RCDT & "") <> "" Then
                    .Text = Format(Left(Trim(ADR_LOC!RCDT & ""), 8), "@@@@-@@-@@") & " " & Format(Mid(Trim(ADR_LOC!RCDT & ""), 10), "@@:@@:@@")
                End If
                .Col = 11 '/검사결과전송일자
                If Trim(ADR_LOC!SDDT & "") <> "" Then
                    .Text = Format(Left(Trim(ADR_LOC!SDDT & ""), 8), "@@@@-@@-@@") & " " & Format(Mid(Trim(ADR_LOC!SDDT & ""), 10), "@@:@@:@@")
                End If
                .Col = 12 '/처방일자
                If Trim(ADR_LOC!ORDDT & "") <> "" Then
                    .Text = Format(Trim(ADR_LOC!ORDDT & ""), "@@@@-@@-@@")
                End If
                
                .Col = 13
                Select Case Trim(ADR_LOC!ORDGB & "") '/처방종류(O.외래, I.입원, G.건강검진)
                    Case "O": .Text = "외래"
                    Case "I": .Text = "입원"
                    Case "G": .Text = "검진"
                End Select
                
                .Col = 14: .Text = Trim(ADR_LOC!PATNO & "") '/병록번호
                .Col = 15: .Text = Trim(ADR_LOC!PATNM & "") '/수검자명
                If Trim(ADR_LOC!PATSEX & "") <> "" Or Trim(ADR_LOC!PATAGE & "") <> "" Then
                    .Col = 16: .Text = Trim(ADR_LOC!PATSEX & "") & "/" & Trim(ADR_LOC!PATAGE & "") '/Sex/Age
                End If
                
                If .MaxTextRowHeight(intX) > 13.3 Then .RowHeight(intX) = .MaxTextRowHeight(intX)
                
                ADR_LOC.MoveNext
            Loop
            ADR_LOC.Close: Set ADR_LOC = Nothing
        Else
            MsgBox "자료가 없습니다.", vbInformation, "확인"
        End If
    End With

    Call CloseDB_LOC

    FUNC_MM_VIEW_LIST = True
    
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_MM_VIEW_RSLT(argBARCD As String, argEXSEQ As String) As Boolean
    FUNC_MM_VIEW_RSLT = False
    
On Error GoTo RTN_ERR
    
    If ConnDB_LOC = False Then Exit Function
    
    With sprDResult
        gstrQuy = "SELECT A.* "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES A, EQ_MST B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQCD  = B.EQCD "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.BARCD = '" & Trim(argBARCD) & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.EXSEQ =  " & Val(argEXSEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY B.EQSEQ "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
        
        If Not ADR_LOC Is Nothing Then
            .MaxRows = ARC_LOC
            barStatus.Max = ARC_LOC
            intX = 0
            
            Do Until ADR_LOC.EOF
                intX = intX + 1: .Row = intX: barStatus.Value = intX
                
                .Col = 1:  .Text = Trim(ADR_LOC!EQCD & "")     '/장비검사코드
                .Col = 2:  .Text = Trim(ADR_LOC!EXAMCD & "")     '/검사코드
                .Col = 3:  .Text = Trim(ADR_LOC!Result & "")     '/검사결과
                .Col = 4:  .Text = Trim(ADR_LOC!EQRESULT & "")     '/장비결과
                .Col = 5:  .Text = Trim(ADR_LOC!AFLAG & "")     '/R
                .Col = 6:  .Text = Trim(ADR_LOC!PFLAG & "")     '/P
                .Col = 7:  .Text = Trim(ADR_LOC!DFLAG & "")     '/D
                
                .Col = 8                                        '/결과진행상태 (0:처방, 1:결과)
                Select Case Trim(ADR_LOC!STATEFLAG & "")
                    Case "0": .Text = "처방"
                    Case "1": .Text = "결과"
                End Select
                
                .Col = 9                                        '/HIS 전송 FLAG (0:대기, 1:완료)
                Select Case Trim(ADR_LOC!SENDFLAG & "")
                    Case "0": .Text = "대기"
                    Case "1": .Text = "완료"
                End Select
                .Col = 10: .Text = Trim(ADR_LOC!ORDDT & "")     '/처방일자
                .Col = 11: .Text = Trim(ADR_LOC!EXDT & "") & " " & Trim(ADR_LOC!EXTM & "") '/검사처방전송일시
                .Col = 12: .Text = Trim(ADR_LOC!RCDT & "") & " " & Trim(ADR_LOC!RCTM & "") '/검사결과수신일시
                .Col = 13: .Text = Trim(ADR_LOC!SDDT & "") & " " & Trim(ADR_LOC!SDTM & "") '/검사결과전송일시
                
                If .MaxTextRowHeight(intX) > 13.3 Then .RowHeight(intX) = .MaxTextRowHeight(intX)
                
                ADR_LOC.MoveNext
            Loop
            ADR_LOC.Close: Set ADR_LOC = Nothing
        Else
            MsgBox "자료가 없습니다.", vbInformation, "확인"
        End If
    End With

    Call CloseDB_LOC

    FUNC_MM_VIEW_RSLT = True
    
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Private Sub cboSENDFLAG_Click()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub cboSENDFLAG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSTATEFLAG_Click()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub cboSTATEFLAG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Call FUNC_MM_VIEW_LIST
    If sprLResult.MaxRows > 0 Then sprLResult.SetFocus
End Sub

Private Sub dtpDateFrom_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub dtpDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpDateTo_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub dtpDateTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
'''    Call FUNC_MM_VIEW_LIST
End Sub

Private Sub Form_Resize()
    Dim intCnt  As Integer
    
On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : 높이가 늘어나는 개체 3개, 디자인상 해당 개체 위에 늘어난 개체가 2개
    For intCnt = 0 To UBound(CW)
        Select Case CW(intCnt).Nm
            Case cmdSave.Name:      cmdSave.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case cmdQuit.Name:      cmdQuit.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case barStatus.Name: barStatus.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprLResult.Name:   sprLResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height + (Me.Height - lngMeHeight)
            Case shpDInfo.Name: shpDInfo.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case shpDResult.Name: shpDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprDResult.Name:   sprDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height + (Me.Height - lngMeHeight)
        End Select
    Next intCnt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Set frmEQ공용_장비검사코드관리_조회 = Nothing
End Sub

Private Sub optDateSection_Click(Index As Integer)
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub sprLResult_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col < 2 Then Exit Sub
    If Row < 1 Then Exit Sub
    
    Call SUB_MM_KEY_CLEAR("2")

    lblBARCD = GET_CELL(sprLResult, 2, Row)     '/검체번호
    lblEXSEQ = GET_CELL(sprLResult, 3, Row)     '/검사회차
    lblSAMPLENO = GET_CELL(sprLResult, 4, Row)  '/Sample No
    lblDISKNOPOSNO = GET_CELL(sprLResult, 5, Row) & "/" & GET_CELL(sprLResult, 6, Row) '/Rack/Pos
    lblEXDT = GET_CELL(sprLResult, 9, Row)      '/검사처방전송일자
    lblRCDT = GET_CELL(sprLResult, 10, Row)     '/검사결과수신일자
    lblSDDT = GET_CELL(sprLResult, 11, Row)     '/검사결과전송일자
    lblORDDT = GET_CELL(sprLResult, 12, Row)    '/처방일자
    lblORDGB = GET_CELL(sprLResult, 13, Row)    '/입/외구분
    lblPATNO = GET_CELL(sprLResult, 14, Row)    '/병록번호
    lblPATNM = GET_CELL(sprLResult, 15, Row)    '/수검자명
    lblSEXAGE = GET_CELL(sprLResult, 16, Row)   '/성별/연령
    
    Call FUNC_MM_VIEW_RSLT(GET_CELL(sprLResult, 2, Row), GET_CELL(sprLResult, 3, Row))
End Sub

Private Sub sprLResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call sprLResult_DblClick(sprLResult.ActiveCol, sprLResult.ActiveRow)
End Sub

Private Sub optDateSection_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtBARCD_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub txtBARCD_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtBARCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPATNM_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub txtPATNM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtPATNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPATNO_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub txtPATNO_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtPATNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
