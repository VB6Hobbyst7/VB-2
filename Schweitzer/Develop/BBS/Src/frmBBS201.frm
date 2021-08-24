VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS201 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Cross-Match 결과등록"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   8.25
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmBBS201.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.Frame fraABO 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   7710
      TabIndex        =   52
      Top             =   5250
      Width           =   2415
      Begin VB.TextBox txtRH 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   63
         Top             =   1575
         Width           =   1110
      End
      Begin VB.TextBox txtSABO 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   62
         Top             =   1215
         Width           =   1110
      End
      Begin VB.TextBox txtCABO 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   53
         Top             =   855
         Width           =   1110
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   75
         Width           =   2310
         _ExtentX        =   4075
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
         Alignment       =   1
         Caption         =   "혈액형 등록"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblaboptnm 
         Height          =   300
         Left            =   1230
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   495
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "홍길동의자"
         Appearance      =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   165
         TabIndex        =   61
         Tag             =   "103"
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "SerumABO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   165
         TabIndex        =   60
         Tag             =   "103"
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label lblabocancel 
         AutoSize        =   -1  'True
         Caption         =   "Cancel"
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
         Left            =   1470
         TabIndex        =   58
         Top             =   2085
         Width           =   705
      End
      Begin VB.Label lblaboapply 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
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
         Left            =   510
         TabIndex        =   57
         Top             =   2085
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자명"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   56
         Tag             =   "103"
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "CellABO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   55
         Tag             =   "103"
         Top             =   900
         Width           =   690
      End
   End
   Begin VB.Frame fraList 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   10155
      TabIndex        =   40
      Top             =   5265
      Width           =   2715
      Begin VB.ListBox lstResult 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         ItemData        =   "frmBBS201.frx":076A
         Left            =   180
         List            =   "frmBBS201.frx":076C
         Style           =   1  '확인란
         TabIndex        =   43
         Top             =   1200
         Width           =   2475
      End
      Begin VB.TextBox txtBloodNo 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   42
         Top             =   480
         Width           =   1605
      End
      Begin VB.TextBox txtCompcdnm 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   41
         Top             =   840
         Width           =   1605
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   315
         Left            =   60
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   60
         Width           =   2595
         _ExtentX        =   4577
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
         Alignment       =   1
         Caption         =   "상세결과등록"
         Appearance      =   0
      End
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         Caption         =   "Cancel"
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
         Left            =   1620
         TabIndex        =   48
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
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
         Left            =   660
         TabIndex        =   47
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액번호"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   46
         Tag             =   "103"
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액제제"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   45
         Tag             =   "103"
         Top             =   900
         Width           =   780
      End
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7620
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   39
      Text            =   "frmBBS201.frx":076E
      Top             =   5685
      Visible         =   0   'False
      Width           =   6615
   End
   Begin FPSpread.vaSpread tblOrder 
      Height          =   1695
      Left            =   75
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1980
      Width           =   9525
      _Version        =   196608
      _ExtentX        =   16801
      _ExtentY        =   2990
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
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
      MaxCols         =   15
      MaxRows         =   1
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS201.frx":07A2
      UserResize      =   0
      TextTip         =   3
   End
   Begin MedControls1.LisLabel LisLabel7 
      Height          =   315
      Left            =   75
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3720
      Width           =   14370
      _ExtentX        =   25347
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
      Caption         =   "  결과 등록"
      Appearance      =   0
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   75
      TabIndex        =   26
      Top             =   3975
      Width           =   14385
      Begin VB.ComboBox cboComment 
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
         ItemData        =   "frmBBS201.frx":0D96
         Left            =   10425
         List            =   "frmBBS201.frx":0D98
         Style           =   2  '드롭다운 목록
         TabIndex        =   91
         Top             =   240
         Width           =   1770
      End
      Begin VB.TextBox txtLabelCnt 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13050
         MaxLength       =   1
         TabIndex        =   87
         Text            =   "2"
         Top             =   195
         Width           =   570
      End
      Begin VB.CheckBox chkABO 
         BackColor       =   &H00DBE6E6&
         Caption         =   "혈액형등록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8730
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   600
         Value           =   1  '확인
         Width           =   1455
      End
      Begin VB.CommandButton cmdSizing 
         BackColor       =   &H00F4F0F2&
         Caption         =   "▽"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11820
         Style           =   1  '그래픽
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "최대로"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdTagPrint 
         BackColor       =   &H00F4F0F2&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13875
         Picture         =   "frmBBS201.frx":0D9A
         Style           =   1  '그래픽
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "혈액Tag 재출력"
         Top             =   570
         Width           =   345
      End
      Begin VB.ComboBox cboMethod 
         Appearance      =   0  '평면
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   360
         ItemData        =   "frmBBS201.frx":12CC
         Left            =   11805
         List            =   "frmBBS201.frx":12DC
         Locked          =   -1  'True
         Style           =   1  '단순 콤보
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   570
         Width           =   2070
      End
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드로 입력(&B)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Top             =   210
         Value           =   1  '확인
         Width           =   1755
      End
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   1
         Top             =   555
         Width           =   2205
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   330
         Left            =   10425
         TabIndex        =   28
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   " 검사방법"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   195
         Index           =   0
         Left            =   3705
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   195
         Index           =   1
         Left            =   5835
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   255
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   600
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   16711680
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   555
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "혈액번호"
         Appearance      =   0
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   13620
         TabIndex        =   88
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtSABO"
         BuddyDispid     =   196611
         OrigLeft        =   3840
         OrigTop         =   330
         OrigRight       =   4080
         OrigBottom      =   645
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   11
         Left            =   12210
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   195
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "출력장수"
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장"
         Height          =   180
         Left            =   13935
         TabIndex        =   90
         Tag             =   "151"
         Top             =   285
         Width           =   195
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3960
         TabIndex        =   37
         Tag             =   "103"
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과재등록(응급)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   6075
         TabIndex        =   36
         Tag             =   "103"
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과등록"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   4875
         TabIndex        =   35
         Tag             =   "103"
         Top             =   600
         Width           =   780
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   45
      Width           =   2955
      _ExtentX        =   5212
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
      Caption         =   " 접수번호"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   3060
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   11415
      _ExtentX        =   20135
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
      Caption         =   " 환자정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   75
      TabIndex        =   7
      Top             =   285
      Width           =   2970
      Begin VB.TextBox txtSpcNO 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   345
         Width           =   1665
      End
      Begin MedControls1.LisLabel lblReaction 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   12640511
         ForeColor       =   255
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
         Caption         =   "Reaction"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblInfection 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BackColor       =   12640511
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "@"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   10
         Left            =   135
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   345
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   9630
      TabIndex        =   13
      Top             =   1620
      Width           =   4830
      _ExtentX        =   8520
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
      Caption         =   " 검체정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   3060
      TabIndex        =   15
      Top             =   285
      Width           =   11430
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   870
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "상병명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   525
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "상병코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblrmk 
         Height          =   300
         Left            =   10440
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         BackColor       =   14411494
         ForeColor       =   255
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
      Begin VB.CommandButton cmdRmk 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7500
         Style           =   1  '그래픽
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "최대로"
         Top             =   495
         Width           =   885
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   3330
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "홍길동의자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   7515
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         BackColor       =   14411494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   5490
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   5490
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         BackColor       =   14411494
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
         Caption         =   "M/09"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSickCd 
         Height          =   315
         Left            =   1185
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   525
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "12345"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSick 
         Height          =   315
         Left            =   1185
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   870
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "아픈곳이 너무 많아요"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1185
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "홍길동의자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   2265
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "환자명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   4425
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   6450
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   4425
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   510
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "병동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   6450
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   510
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin VB.Label lblABO 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "AB(AB)+"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   9165
         TabIndex        =   23
         Top             =   465
         Width           =   2085
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   8925
         TabIndex        =   24
         Top             =   120
         Width           =   2445
      End
   End
   Begin FPSpread.vaSpread tblBlood 
      Height          =   3360
      Left            =   75
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5055
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   5927
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
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
      MaxCols         =   31
      MaxRows         =   13
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS201.frx":1306
      UserResize      =   1
      VirtualRows     =   7
      TextTip         =   4
   End
   Begin MSComctlLib.TabStrip tabData 
      Height          =   315
      Left            =   9615
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   556
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "검사정보"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "검체정보"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "최근수혈정보"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fradata 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   1
      Left            =   9660
      TabIndex        =   68
      Top             =   1965
      Width           =   4785
      Begin FPSpread.vaSpread tblTest 
         Height          =   1305
         Left            =   30
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   45
         Width           =   4530
         _Version        =   196608
         _ExtentX        =   7990
         _ExtentY        =   2302
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
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
         MaxCols         =   4
         MaxRows         =   0
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS201.frx":37DF
      End
   End
   Begin VB.Frame fradata 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   0
      Left            =   9660
      TabIndex        =   65
      Top             =   1965
      Width           =   4785
      Begin FPSpread.vaSpread tblSpc 
         Height          =   690
         Left            =   30
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   60
         Width           =   4530
         _Version        =   196608
         _ExtentX        =   7990
         _ExtentY        =   1217
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
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
         MaxCols         =   4
         MaxRows         =   1
         OperationMode   =   1
         ScrollBars      =   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS201.frx":3B05
      End
      Begin MedControls1.LisLabel lblAddChk 
         Height          =   540
         Left            =   30
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   780
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   953
         BackColor       =   12648447
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   ""
      End
   End
   Begin VB.Frame fradata 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   2
      Left            =   9660
      TabIndex        =   70
      Top             =   1965
      Width           =   4785
      Begin VB.Label lblTransDt 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
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
         Height          =   285
         Left            =   1065
         TabIndex        =   76
         Top             =   945
         Width           =   3375
      End
      Begin VB.Label lblLastBldNo 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
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
         Height          =   285
         Left            =   1065
         TabIndex        =   75
         Top             =   540
         Width           =   3375
      End
      Begin VB.Label lblLastComp 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
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
         Height          =   285
         Left            =   1065
         TabIndex        =   74
         Top             =   135
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수혈일시 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   73
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "혈액번호 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   72
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "혈액제제 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   71
         Top             =   180
         Width           =   1080
      End
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   75
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1620
      Width           =   9525
      _ExtentX        =   16801
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
      Caption         =   " 처방정보"
      Appearance      =   0
   End
   Begin VB.Label lblLog 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00808080&
      Caption         =   "이 처방에 대한 혈액이 모두 준비되었습니다."
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   75
      TabIndex        =   38
      Top             =   8580
      Width           =   8700
   End
End
Attribute VB_Name = "frmBBS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn1
    tcNo = 1
    tcBldNo
    tcCOMPONM
    tcABO
    tcVol
    
    tcSTAT      '6
    tcOK
    tcNot
    tcIRR
    tcFilter
    
    tcSPCNO
    tcVFYNM
    tcVFYDT
    tcSTATUS
    tcDELIVERYDT
    
    tcDTAILRST
    tcCMTBTN
    tcSTEP1
    tcSTEP2
    tcSTEP3
    
    tcSTEP4
    tcCOMPOCD
    tcRSTSEQ
    tcABBRNM
    tcDUP
    
    tcRESULTFG
    tcFLAG
    tcASSIGN
    tcRMK
    tcNORSV
    tcVfyTm
End Enum

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private lngAccDt As Long            '디비에 저장된 접수일자(AccDt의 형식)
Private SpcNum As String            '검체번호
Private strBNum As String           '입력된 혈액번호(사용시,2,2,임의, 형식으로 끊어서 사용한다.)
Private Test_Step As Long           '테스트단계

'print시에 사용되는 변수.....
Private strPtid   As String                 '환자ID
Private strOrdDt  As String                 '처방일
Private strWardID As String                 '병동
Private strDeptCd As String                 '진료과
Private lngOrdNo       As Long              '처방번호
Private lngOrdseq      As Long              '처방seq
Private strComponent   As String            '혈액제제코드
Private strComponentNm As String            '혈액제제명
Private strVolume As String
Private lngUnitQty     As Integer
Private strSSN         As String            '주민번호

Private InPutNo As Integer

'플래그
Private blnStat As Boolean
Private onPgm   As Boolean
Private UpInchk As Boolean          '응급검체 결과재등록
Private RePrint As Boolean

Private Const CurrentSelected$ = "▶" '현재 선택된 오더 표시

Public Sub CallByExtForm()
    Call txtSpcNo_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cboComment_Click()
    Dim Step(3) As String
    Dim ii      As Integer
    Dim ResultFg As Boolean
    Dim strComment As Integer
    Dim varTmp

    strComment = cboComment.ListIndex

    If strComment = 0 Then
        Exit Sub
    End If

    With tblBlood
        For ii = 1 To .MaxRows
            .GetText 1, ii, varTmp
            If varTmp <> "" Then
                .Row = ii
                .Col = TblColumn1.tcSTEP1: .value = IIf(strComment = 1, "1", "0")
                .Col = TblColumn1.tcSTEP2: .value = IIf(strComment = 2, "1", "0")
                .Col = TblColumn1.tcSTEP3: .value = IIf(strComment = 3, "1", "0")
                .Col = TblColumn1.tcSTEP4: .value = IIf(strComment = 4, "1", "0")
                .Col = TblColumn1.tcDTAILRST: .value = ""

                If strComment = 0 Then
                    onPgm = True
                    .Col = TblColumn1.tcOK: .value = False
                    .Col = TblColumn1.tcNot: .value = True
                    onPgm = False
                Else
                    onPgm = True
                    .Col = TblColumn1.tcOK: .value = True
                    .Col = TblColumn1.tcNot: .value = False
                    onPgm = False
                End If
            End If
        Next
    End With
        
End Sub

Private Sub chkABO_Click()
    Dim strTmp As String
    Dim ii     As Integer
    
    If chkABO.value = 1 Then
        fraABO.Visible = True
        lblaboptnm.Caption = lblPtNm.Caption
        If lblABO.Caption = "" Then
            txtCABO.Text = ""
            txtSABO.Text = ""
            txtRH.Text = ""
        Else
            If Len(lblABO.Caption) > 3 Then
                txtCABO.Text = medGetP(lblABO.Caption, 1, "(")
                txtSABO.Text = medGetP(medGetP(lblABO.Caption, 2, "("), 1, ")")
                txtRH.Text = medGetP(lblABO.Caption, 2, ")")
            Else
                For ii = 1 To Len(lblABO.Caption)
                    If Mid(lblABO.Caption, ii, 1) = "+" Or Mid(lblABO.Caption, ii, 1) = "-" Then
                        txtRH.Text = Mid(lblABO.Caption, ii, 1)
                    Else
                        strTmp = strTmp & Mid(lblABO.Caption, ii, 1)
                    End If
                Next ii
                txtCABO.Text = strTmp
            End If
        End If
    Else
        If fraABO.Visible Then fraABO.Visible = False
    End If
End Sub

Private Sub chkBar_Click()
    txtBldNo = ""
    txtBldNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Clear
    tblOrder.MaxRows = 0
    txtSpcNO.SetFocus
End Sub

Private Sub Clear()
    InPutNo = 0
    txtSpcNO.Text = ""
    txtComment.Text = ""
    txtBldNo.Text = ""
    chkABO.value = 0
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblWardNm.Caption = ""
    lblDeptNm.Caption = ""
    lblSickCd.Caption = ""
    lblSick.Caption = ""
    lblABO.Caption = ""
    lblAddChk.Caption = ""
    lblrmk.Caption = ""
   
    lblTransDt.Caption = ""
    lblLastComp.Caption = ""
    lblLastBldNo.Caption = ""
    
    tblTest.MaxRows = 0
    tblBlood.MaxRows = 0
    tblSpc.MaxRows = 0
    lblLog.Visible = False
    cmdTagPrint.Enabled = False
    lblInfection.Visible = False
    lblReaction.Visible = False
    fraABO.Visible = False
    lblaboptnm.Caption = ""
    txtCABO.Text = ""
    txtSABO.Text = ""
    txtRH.Text = ""
    
    cmdRmk.Caption = ""
'    cmdRmk.Visible = False
    fraResult.Visible = True
    Call ICSPatientMark
    
    txtLabelCnt.Text = "2"
    
End Sub

Private Sub cmdRmk_Click()
    If lblPtId.Caption = "" Then Exit Sub
    frmXMRemark.sPtid = lblPtId.Caption
    frmXMRemark.rmk = lblrmk.Caption
    frmXMRemark.Show 1
End Sub




'Private Sub cmdSizing_Click()
'    If cmdSizing.Caption = "▽" Then
'        tblXM.Height = 5820
'        cmdSizing.Caption = "△"
'        cmdSizing.ToolTipText = "이전크기로"
'    ElseIf cmdSizing.Caption = "△" Then
'        tblXM.Height = 1260
'        cmdSizing.Caption = "▽"
'        cmdSizing.ToolTipText = "최대로"
'    End If
'End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If txtComment.Visible Then txtComment.Visible = False
        If fraList.Visible Then fraList.Visible = False
        If fraABO.Visible Then fraABO.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Call Form_Setting
    Call Clear
    
    cboComment.AddItem ""
    cboComment.AddItem "Saline"
    cboComment.AddItem "Bovine"
    cboComment.AddItem "37'c"
    cboComment.AddItem "Coombs'"
    cboComment.ListIndex = 0
End Sub

Private Sub Form_Setting()
    '검사Step을 가지고 온다.
    '접수일자의 형식을 가지고 온다.
    Dim objXM     As New clsCrossMatching
    
    Dim DrRS      As New Recordset
     
    Dim strStep As String
    Dim strTmp  As String
    Dim Cnt     As Integer
    Dim jj      As Integer
    Dim ii      As Integer
    Dim kk      As Long
    
    Set DrRS = objXM.Get_XM_Step
    
    If Not DrRS.EOF Then
        Test_Step = Val(DrRS.Fields("field1").value & "")
        lstResult.Clear
        For ii = 1 To Test_Step
            lstResult.AddItem medGetP(Trim(DrRS.Fields("text1").value & ""), ii, ";")
        Next
    End If
    

    fradata(1).ZOrder 0
    fraList.Visible = False
    
    Dim objNumbers As New clsBBSNumbers
    
    With objNumbers    '접수 일자의 형식을 가져온다.
        lngAccDt = Len(.Get_AccdtFormat)
    End With
    
    Set objXM = Nothing
    Set objNumbers = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim objSql As clsCrossMatching
    Dim SSQL      As String
    Dim ii        As Integer
    Dim strRT     As String
    Dim strBldNo  As String
    Dim strBldSrc As String
    Dim strBldYY  As String
    Dim strCompCd As String
    Dim strError  As String
    
    Dim SaveTF    As Boolean
    
    If tblBlood.DataRowCnt < 1 Then Exit Sub
    
'fraABO 혈액형등록 창이 떠 있는 경우 Apply를 누르지 않은 경우 저장 버튼 누르지 못하도록
    If fraABO.Visible Then
        MsgBox "혈액형을 등록(Apply)한 후 Assign하십시오.", vbExclamation
        Exit Sub
    End If
    
'fraList 상세결과 창이 떠 있는 경우 Apply를 누르지 않은 경우 저장 버튼 누르지 못하도록
    If fraList.Visible Then
        MsgBox "상세결과를 저장(Apply)한 후 Assing하십시오.", vbExclamation
        Exit Sub
    End If
    
'PRC,WB는 상세결과를 반드시 입력하도록
    If CheckXMDetail = False Then
        MsgBox "XM결과 필수입력 제제입니다. ""?""로 표시된 항목은 상세결과를 입력하십시오.", vbExclamation
        Exit Sub
    End If
    
'수혈처방의 제제와 Assign대기 중인 제제가 다른 경우 한번 더 확인 하도록...
    If CheckDiffCompo Then
        If MsgBox("수혈처방의 제제와 Assign대기 중인 제제가 다른 혈액이 있습니다. 계속 진행하시겠습니까?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    If CheckDiffABO Then
        If MsgBox("환자의 혈액형과 Assign대기 중인 혈액의 혈액형이 다른 혈액이 있습니다. 계속 진행하시겠습니까?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    '** 처방명과 제제명 앞3자리 비교 후 틀리면 결과등록 못함 By M.G.Choi 2008.04.17 -- 보류
'    If CheckCompo = False Then
'        MsgBox "처방명과 제제가 틀립니다.", vbCritical
'        Exit Sub
'    End If
    
    Me.MousePointer = 11
    
    If Assign_Cnt = True Then
        If Insert_Sql = True Then
            Call Clear
            tblOrder.MaxRows = 0
            txtSpcNO.SetFocus
        Else
            txtBldNo.SetFocus
        End If
    Else
        On Error GoTo SAVE_ERROR
        DBConn.BeginTrans
        
        Set objSql = New clsCrossMatching
        With tblBlood
            For ii = 1 To .MaxRows
                .Row = ii: .Col = TblColumn1.tcIRR: strRT = .value
                If strRT = "1" Then
                    .Col = TblColumn1.tcCOMPOCD: strCompCd = .value
                    .Col = TblColumn1.tcBldNo:
                    strBldSrc = medGetP(.value, 1, "-")
                    strBldYY = medGetP(.value, 2, "-")
                    strBldNo = medGetP(.value, 3, "-")
                    '처방 출고가 된후에  irrdat
                    SSQL = objSql.SetBBS401_IRRADD(strBldSrc, strBldYY, strBldNo, strCompCd, BBSBloodStatus.stsASSIGN, strRT, Format(GetSystemDate, PRESENTDATE_FORMAT), ObjSysInfo.EmpId)
                    DBConn.Execute SSQL
                    SaveTF = True
                End If
            Next
        End With
        If SaveTF = False Then
            MsgBox "처방수량과 ASSIGN 검사 대기수량이 일치하지 않습니다." & vbCrLf & "확인후 작업을 진행하세요", vbInformation + vbOKOnly, "결과등록"
        Else
            Clear
            tblOrder.MaxRows = 0
            MsgBox "정상적으로 저장되었습니다.", vbInformation + vbOKOnly, "결과등록"
        End If
        Set objSql = Nothing
        txtSpcNO.SetFocus
        DBConn.CommitTrans
        
    End If
    Me.MousePointer = 0
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    
    Me.MousePointer = 0
    Set objSql = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Function CheckCompo() As Boolean
    Dim strOrdNm        As String
    Dim strCompo        As String
    Dim iRow            As Long
    
    CheckCompo = True
    
    With tblOrder
        For iRow = 1 To .DataRowCnt
            .Row = iRow: .Col = 1
            If .value = CurrentSelected Then
                .Col = 2: strOrdNm = UCase(Mid(.value, 1, 3))
                Exit For
            End If
        Next
    End With
    
    With tblBlood
        For iRow = 1 To .DataRowCnt
            .Row = iRow
            .Col = 1
            If .value <> "**" Then
                .Col = 3: strCompo = UCase(Mid(.value, 1, 3))
                If strOrdNm <> strCompo Then
                    CheckCompo = False
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Private Function CheckXMDetail() As Boolean
'PRC,WB인 경우 상세결과를 반드시 입력하도록
'tcCOMPOCD 입력하는 제제코드가 PRC,WB에 해당될 경우
    Dim RS As Recordset
    Dim strSQL As String
    Dim strCompocd As String
    Dim i As Long
    Dim vStep1 As Variant
    Dim vStep2 As Variant
    Dim vStep3 As Variant
    Dim vStep4 As Variant
    
    CheckXMDetail = True
    'Assign대기 중인 혈액의 제제를 기준으로 XM 상세결과 필수입력 제제 여부 판단.
    
    With tblOrder
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            If .value = CurrentSelected Then
                .Col = TblColumn1.tcCOMPOCD: strCompocd = .value
                Exit For
            End If
        Next
    End With
    
    strSQL = " select text1 from " & T_COM003 & _
             " where " & DBW("cdindex=", BC2_XM_COMPO) & _
             " and " & DBW("cdval1=", BC2_XM_COMPO)
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF = False Then
        With tblBlood
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = TblColumn1.tcRESULTFG
                If .value = "1" Then    '결과 입력 대기인 것들
                    .Col = TblColumn1.tcCOMPOCD: strCompocd = .value
                    
                    If InStr(RS.Fields("text1").value & "", strCompocd) > 0 Then 'XM결과 필수입력 제제
                        Call .GetText(TblColumn1.tcSTEP1, i, vStep1)
                        Call .GetText(TblColumn1.tcSTEP2, i, vStep2)
                        Call .GetText(TblColumn1.tcSTEP3, i, vStep3)
                        Call .GetText(TblColumn1.tcSTEP4, i, vStep4)
                        
                        'Not으로 결과 입력하는 경우에는 어떻게 처리를 해야 하나...
                        'Stat으로 선택된 경우에는 어떻게 처리를 해야 하나...
                        If (vStep1 = "" And vStep2 = "" And vStep3 = "" And vStep4 = "") Or _
                           (vStep1 = "0" And vStep2 = "0" And vStep3 = "0" And vStep4 = "0") Then
                            Call .SetText(TblColumn1.tcDTAILRST, i, "?")
                            CheckXMDetail = False
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    Set RS = Nothing
End Function

Private Function CheckDiffCompo() As Boolean
'수혈처방 제제와 Assign할 제제가 다른 경우 한번 더 워닝을 띄워준다.
    Dim i As Long
    
    CheckDiffCompo = False
    'tblcolumn1.tcRESULTFG ="1" 이고 TblColumn1.tcCOMPONM 컴퍼넌트 명의 ForeColor이 DCM_Magenta인 경우 있는 지만 체크
    
    With tblBlood
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn1.tcRESULTFG
            If .value = "1" Then    '결과 입력 대기인 것들
                .Col = TblColumn1.tcCOMPONM
                If .ForeColor = DCM_Magenta Then '수혈처방 제제와 Assign할 제제가 다른 경우 ForeColor이 DCM_Magenta로 표시되어 있다.
                    CheckDiffCompo = True
                    Exit For
                End If
            End If
        Next
    End With
End Function

Private Function CheckDiffABO() As Boolean
'환자의 혈액형과 Assign할 혈액의 혈액형이 다른 경우 한번 더 워닝을 띄워준다.
    Dim i As Long
    Dim strABO As String
    
    If Len(lblABO.Caption) > 3 Then
        strABO = medGetP(lblABO.Caption, 1, "(") & medGetP(lblABO.Caption, 2, ")")
    Else
        strABO = lblABO.Caption
    End If
    
    CheckDiffABO = False
    'tblcolumn1.tcRESULTFG ="1" 이고 TblColumn1.tcABO 값 비교
    
    With tblBlood
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn1.tcRESULTFG
            If .value = "1" Then    '결과 입력 대기인 것들
                .Col = TblColumn1.tcABO
                If .value <> strABO Then  '환자의 혈액형과 Assign할 혈액의 혈액형이 다른 경우 표시
                    CheckDiffABO = True
                    Exit For
                End If
            End If
        Next
    End With
End Function

Private Function Assign_Cnt() As Boolean
'결과등록하고자 하는 처방에 대해서 처방 수량과, 결과 등록수량을 비교해서넘겨준다.
    
    Dim objXM As New clsCrossMatching
    Dim strJudge As String
    Dim strEr As String
    Dim AA_Cnt As Long
    Dim A_Cnt As Long   'Assign수량
    Dim C_Cnt As Long   'Assign Cancel 수량
    Dim O_Cnt As Long   '출고수량
    Dim R_Cnt As Long   '반환수량
    Dim X_Cnt As Long   '폐기수량
    Dim T_Cnt As Long   '총Assign 수량
    Dim unitqty As Long
    Dim ACnt As Long
    Dim ii As Integer

    '--------------------------------------------------------------------
    ' 수정되어야 합니다.
    ' 현재 Assign된 수량이 정확하지 않습니다.
    ' 현재 Assign된 수량 = Assign수량 - Assign취소수량 - 반환수량 - 폐기수량
    '--------------------------------------------------------------------

    Assign_Cnt = True
    With objXM
'        .setDbConn DBConn
        .Assign_Cnt medGetP(txtSpcNO, 1, "-"), Val(medGetP(txtSpcNO, 2, "-"))
        A_Cnt = .AssignCnt
        C_Cnt = .CancelCnt
        O_Cnt = .OutCnt
        R_Cnt = .RetCnt
        X_Cnt = .ExpCnt
    End With
    Set objXM = Nothing
    
    T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt

    With tblOrder
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 1
            If .value = CurrentSelected Then
                .Col = 3: unitqty = Val(.value)
                Exit For
            End If
        Next
    End With
    
    'tblOrder.Row = 1: tblOrder.Col = 3: unitqty = Val(tblOrder.value)
    
    
    Dim SA_Cnt As Integer
    
    With tblBlood
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn1.tcASSIGN
            
            If .value = "1" Then
            'Assign 되어있지만 응급인 경우.......
                .Col = TblColumn1.tcRESULTFG
                If .value = "1" Then SA_Cnt = SA_Cnt + 1
            Else
                .Col = TblColumn1.tcRESULTFG
                If .value = "1" Then
                    .Col = TblColumn1.tcSTAT: strEr = .value
                    .Col = TblColumn1.tcOK: strJudge = .value
                    If strJudge = True Or strEr = True Then AA_Cnt = AA_Cnt + 1
                End If
            End If
        Next
    End With
    
    If T_Cnt - SA_Cnt = unitqty Then
'        MsgBox "처방에 대한 모든혈액이 ASSIGN 되어있습니다.", vbInformation + vbOKOnly, Me.Caption
        Assign_Cnt = False
        Exit Function
    End If
    
    If (AA_Cnt + T_Cnt - SA_Cnt) > unitqty Then
'        MsgBox "처방에 대한 수량보다 Assign 대상 혈액이 초과합니다." & vbNewLine & _
               "이미 " & A_Cnt & "개의 혈액이 Assign 되어있으며," & O_Cnt & _
               "개의 혈액이 출고되었습니다.", vbInformation + vbOKOnly, Me.Caption
        Assign_Cnt = False
    End If
    
End Function

Private Function Insert_Sql() As Boolean
'Cross-Matching 결과내역 작성
    Dim objXM As New clsCrossMatching
    Dim strAccDt   As String              '접수번호(년)
    Dim lngAccSeq  As Long                '접수번호
    
    Dim strBldSrc  As String              '혈액번호(혈액원0
    Dim strBldYY   As String              '혈액번호(년도)
    Dim lngBldNo   As Long                '혈액번호(일련번호)
    Dim lngRstSeq  As Long                '결과Seq
    Dim strCompCd  As String              '혈액제재코드
    Dim strXmethod As String              '검사방법
    
    Dim strSTEP1   As String              '테스트1
    Dim strSTEP2   As String              '테스트2
    Dim strSTEP3   As String              '테스트3
    Dim strSTEP4   As String              '테스트4
    Dim strRstVal  As String              '검사결과
    
    Dim strSpcYY   As String              '검체번호(년도)
    Dim lngSpcNo   As Long                '검체번호(일련)
    
    
    Dim strVfyDt   As String              '검사일(일반)
    Dim strVfyTm   As String              '검사시간(일반)
    Dim strVfyId   As String
    Dim strTestchk As String              '검사여부(1:검사한다,0:검사없이 Assign):col=23
    Dim strStat    As String              '응급인경우
    Dim strStatDt  As String              '검사일
    Dim strStatTm  As String              '검사시간
    Dim strStatID  As String              '검사자
    Dim strRmk As String                  'Comment
    
    '2001-11-12 추가
    '검사결과가 Not 인경우, cancelfg/dt/tm/id 를 update 해준다.
    Dim strCancelFg As String
    Dim strCancelDt As String
    Dim strCancelTm As String
    Dim strCancelId As String
    
    'Tag변수
    Dim TagBldno    As String
    Dim TagCompoNm  As String
    Dim TagABO      As String
    Dim TagVolumn   As String
    Dim strSTEP11   As String
    Dim strSTEP22   As String
    Dim strSTEP33   As String
    Dim strSTEP44   As String
    Dim strDetail   As String
    
    '날짜변수
    Dim strThisDate As String
    Dim strThisTime As String

    Dim SSQL       As String
    Dim ii         As Integer
    
    'irr변수,Filter 변수
    Dim strirr     As String
    Dim strFilter  As String
    
    
    strThisDate = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strThisTime = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
    'X매칭 검사방법
    strXmethod = cboMethod.ListIndex
    
    If strXmethod = "3" Then
        strSpcYY = ""
        lngSpcNo = 0
        strTestchk = "1"
    Else
        strSpcYY = UCase(Mid(SpcNum, 1, 2))
        lngSpcNo = Val(Mid(SpcNum, 4))
        strTestchk = "0"
    End If

On Error GoTo XM_Result_Save_Error
    
    DBConn.BeginTrans
    
    
    strAccDt = medGetP(txtSpcNO, 1, "-")
    lngAccSeq = Val(medGetP(txtSpcNO, 2, "-"))
    
    With objXM
        lngRstSeq = .Get_RstSeq(strAccDt, lngAccSeq)
    End With
    
    With tblBlood
        For ii = 1 To .DataRowCnt
            .Row = ii:
            
            .Col = TblColumn1.tcOK
            If .CellType = CellTypeStaticText Then
                 strRstVal = IIf(.value = "√", 1, "")
            Else
                 strRstVal = IIf(.value = True, 1, "")
            End If
            '2001-11-12 추가 ----
            strCancelFg = ""
            strCancelDt = ""
            strCancelTm = ""
            strCancelId = ""
            '-------------------
 
            If strRstVal = "" Then
                .Col = TblColumn1.tcNot
                If .CellType = CellTypeStaticText Then
                    strRstVal = IIf(.value = "√", "0", "")
                Else
                    strRstVal = IIf(.value = True, "0", "")
                End If
                If strRstVal = "0" Then
                    '2001-11-12 추가 ------------------------------------------
                    strCancelFg = "1"
                    strCancelDt = Format(GetSystemDate, CS_DateDbFormat)
                    strCancelTm = Format(GetSystemDate, CS_TimeDbFormat)
                    strCancelId = ObjMyUser.EmpId
                    '----------------------------------------------------------
                End If
            End If
               
            .Col = TblColumn1.tcSTAT:   ' strStat = IIf(.value = True, 1, "")
            If .CellType = CellTypeStaticText Then
                 strStat = IIf(.value = "√", 1, "")
            Else
                 strStat = IIf(.value = True, 1, "")
            End If
            
            If strStat = "1" Then If strRstVal = "" Then strRstVal = ""
                 
            .Col = TblColumn1.tcRMK:     strRmk = .value
            
            .Col = TblColumn1.tcCOMPOCD: strCompCd = .value
            .Col = TblColumn1.tcBldNo:   TagBldno = .value
            
            strBldSrc = medGetP(TagBldno, 1, "-"): strBldYY = medGetP(TagBldno, 2, "-"): lngBldNo = medGetP(TagBldno, 3, "-")
            
            'Tag 변수
            '-------------------------------------------------------------------
            TagBldno = Mid(TagBldno, 1, 6) & Format(Mid(TagBldno, 7), "00000#")
            .Col = TblColumn1.tcABO:     TagABO = .value
            .Col = TblColumn1.tcVol:     TagVolumn = .value
            .Col = TblColumn1.tcABBRNM:  TagCompoNm = .value
            '-------------------------------------------------------------------
            .Col = TblColumn1.tcIRR:     strirr = .value
            .Col = TblColumn1.tcFilter:  strFilter = IIf(.value = True, 1, "")
               
            strSTEP1 = "": strSTEP2 = "": strSTEP3 = "": strSTEP4 = ""
            Select Case Test_Step
                Case 1:
                    .Col = TblColumn1.tcSTEP1: strSTEP1 = .value
                Case 2:
                    .Col = TblColumn1.tcSTEP1: strSTEP1 = .value
                    .Col = TblColumn1.tcSTEP2: strSTEP2 = .value
                Case 3:
                    .Col = TblColumn1.tcSTEP1: strSTEP1 = .value
                    .Col = TblColumn1.tcSTEP2: strSTEP2 = .value
                    .Col = TblColumn1.tcSTEP3: strSTEP3 = .value
                Case 4:
                    .Col = TblColumn1.tcSTEP1: strSTEP1 = .value
                    .Col = TblColumn1.tcSTEP2: strSTEP2 = .value
                    .Col = TblColumn1.tcSTEP3: strSTEP3 = .value
                    .Col = TblColumn1.tcSTEP4: strSTEP4 = .value
            End Select
               
            If strStat = "1" Then
                .Col = TblColumn1.tcASSIGN
                If .value = "1" Then
                    strStatDt = strThisDate
                    strStatTm = strThisTime
                    strStatID = ObjMyUser.EmpId
                    strVfyDt = strThisDate
                    strVfyTm = strThisTime
                    strVfyId = ObjMyUser.EmpId
                Else
                    strVfyDt = ""
                    strVfyTm = ""
                    strVfyId = ""
                    strStatDt = strThisDate
                    strStatTm = strThisTime
                    strStatID = ObjMyUser.EmpId
                End If
            Else
                strVfyDt = strThisDate
                strVfyTm = strThisTime
                strVfyId = ObjMyUser.EmpId
                strStat = ""
                strStatDt = ""
                strStatTm = ""
                strStatID = ""
            End If
            strStat = strStat & COL_DIV & strFilter
            
            .Col = TblColumn1.tcRESULTFG
            If .value = "1" Then
                '----------------------------------------------------
                '응급 검체 Update,결과없이 입력되 검사결과 등록UPDATE
                '----------------------------------------------------
                .Col = TblColumn1.tcASSIGN
                If .value = "1" Then
                    Dim lngseq As Long
                    
                    .Col = TblColumn1.tcRSTSEQ
                    lngseq = Val(.value)
                    SSQL = objXM.SetUpdateBBS302(strAccDt, lngAccSeq, lngseq, strSTEP1, strSTEP2, strSTEP3, strSTEP4, strVfyDt, strRstVal, strVfyTm, _
                                                 strVfyId, strRmk, strStat, strStatDt, strStatTm, strStatID)
                    DBConn.Execute SSQL
                    '결과없이 저장된거 재저장하기위해서..
                    .Col = TblColumn1.tcNORSV
                    If .value = "1" Then
                        If strRstVal = "1" Or strStat = "1" Then
                            '혈액의 ASSIGN 상태로 UPDATE
                            SSQL = objXM.Update_BBS401(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN)
                            DBConn.Execute SSQL
                            'IRRADIATOIN 등록
                            If strirr = "1" Then
                                SSQL = objXM.SetBBS401_IRRADD(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN, strirr, strThisDate, ObjMyUser.EmpId)
                            Else
                                SSQL = objXM.SetBBS401_IRRADD(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN, "", "", "")
                            End If
                            DBConn.Execute SSQL
                            '처방별 ASSIGN COUNT등록
                            SSQL = objXM.Insert_BBS203(strAccDt, lngAccSeq)
                            DBConn.Execute SSQL
                        End If
                    Else
                    
                    End If
                Else
                    '--------------------------------------------------------------------------------------
                    'strTestchk="0" 은 Method가 검사를 하는 경우로 결과테이블 작성시 검사결과까지 저장한다.
                    '--------------------------------------------------------------------------------------
                    If strTestchk = "0" Then
                        '2001-11-12 추가---------------------------------------------------------------------------------
                        If strCancelFg = "1" Then
                            '검사를 수행하여 저장하는 경우(검사결과가 NOT 인 경우)
                            SSQL = objXM.Insert_BBS302NotOk(strAccDt, lngAccSeq, lngRstSeq, _
                                                       strBldSrc, strBldYY, lngBldNo, strCompCd, strXmethod, _
                                                       "", strSTEP1, strSTEP2, strSTEP3, strSTEP4, strRstVal, _
                                                       strSpcYY, lngSpcNo, strVfyDt, strVfyTm, strVfyId, _
                                                       strStat, strStatDt, strStatTm, strStatID, strRmk, strCancelDt, _
                                                       strCancelTm, strCancelId)
                        '------------------------------------------------------------------------------------------------
                        Else
                            '검사를 수행하여 저장하는 경우(검사결과가 OK 인 경우)
                            SSQL = objXM.Insert_BBS302(strAccDt, lngAccSeq, lngRstSeq, _
                                                       strBldSrc, strBldYY, lngBldNo, strCompCd, strXmethod, _
                                                       "", strSTEP1, strSTEP2, strSTEP3, strSTEP4, strRstVal, _
                                                       strSpcYY, lngSpcNo, strVfyDt, strVfyTm, strVfyId, _
                                                       strStat, strStatDt, strStatTm, strStatID, strRmk)
                        End If
                    Else
                    '------------------------------------------------------------------------------------------------
                    'strTestchk<>"0" 은 Method가 검사를 하지않는 경우로 결과테이블 작성시 검사결과는 저장하지 않는다.
                    '------------------------------------------------------------------------------------------------
                        SSQL = objXM.Insert_NotestBBS302(strAccDt, CStr(lngAccSeq), strBldSrc, strBldYY, lngBldNo, _
                                                         strCompCd, strThisDate, strThisTime, ObjMyUser.EmpId, strRmk)
                    End If
                    
                    DBConn.Execute SSQL
                    '------------------------------------------------------------------
                    '판정이 Ok인것,응급인것, 혈액입고내역에 ASSIGN 상태로 Update 해준다.
                    '------------------------------------------------------------------
                    If strRstVal = "1" Or strStat = "1" Then
                        SSQL = objXM.Update_BBS401(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN)
                        DBConn.Execute SSQL
                    End If
                    '------------------------------------------------------------------------------------------------
                    'strrstval="1": 검사결과가 Ok인경우,strstat="1": 응급인 경우,strtestchk="1":검사를 하지 않는경우
                    '2번:Assign인경우 혈액입고내역의 stscd를 assign 상태로 update해준다.
                    '3번:처방별 Assign수량을 update해준다.
                    '------------------------------------------------------------------------------------------------
                    If strRstVal = "1" Or strStat = "1" Or strTestchk = "1" Then

                        '2번
                        If strirr = "1" Then
                            SSQL = objXM.SetBBS401_IRRADD(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN, strirr, strThisDate, ObjMyUser.EmpId)
                        Else
                            SSQL = objXM.SetBBS401_IRRADD(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN, "", "", "")
                        End If
                        DBConn.Execute SSQL
                        '3번
                        SSQL = objXM.Insert_BBS203(strAccDt, lngAccSeq)
                        DBConn.Execute SSQL
                    End If
                    '---------------------------------------------------------------------
                    '처방과 관련된 테이블을 update 해준다.(처방바디,처방헤더,처방접수내역)
                    '---------------------------------------------------------------------
                    SSQL = objXM.Update_OrderStatus(strPtid, strOrdDt, lngOrdNo)
                    DBConn.Execute SSQL
                    
                    SSQL = objXM.Update_OrderStatus(strPtid, strOrdDt, lngOrdNo, lngOrdseq)
                    DBConn.Execute SSQL
                    
                    SSQL = objXM.Update_OrderStatus(strPtid, strOrdDt, lngOrdNo, -99)
                    If SSQL <> "" Then DBConn.Execute SSQL
                    
                    SSQL = objXM.Update_BBS202(medGetP(txtSpcNO, 1, "-"), Val(medGetP(txtSpcNO, 2, "-")))
                    DBConn.Execute SSQL
                    
                    '------------
                    '혈액 Tag출력
                    '------------
                    '-- 주민번호 --> 상세결과 추가 By M.G.Choi 2007.07.02
                    .Col = TblColumn1.tcSTEP1: strSTEP11 = IIf(.value = "1", "S(O)", "S(X)")
                    .Col = TblColumn1.tcSTEP2: strSTEP22 = IIf(.value = "1", "B(O)", "B(X)")
                    .Col = TblColumn1.tcSTEP3: strSTEP33 = IIf(.value = "1", "37(O)", "37(X)")
                    .Col = TblColumn1.tcSTEP4: strSTEP44 = IIf(.value = "1", "C(O)", "C(X)")
                    strDetail = strSTEP11 & strSTEP22 & strSTEP33 & strSTEP44
                    
                    RePrint = False
                    Call TagPrint(TagBldno, TagCompoNm, TagABO, TagVolumn, strirr, strDetail)
                    
                    lngRstSeq = lngRstSeq + 1
                End If
            Else
                '------------------------------------------------------------
                '결과등록후 IRRADITION 등록을 추가로 설정할 경우.(2001/07/12)
                '------------------------------------------------------------
                If strirr = "1" Then
                    SSQL = objXM.SetBBS401_IRRADD(strBldSrc, strBldYY, lngBldNo, strCompCd, BBSBloodStatus.stsASSIGN, strirr, strThisDate, ObjMyUser.EmpId)
                    DBConn.Execute SSQL
                End If
            End If
            
            '---------------------------------------------------------------------------------------------------------------
            '2001/07/23
            '2인의 환자에게 혈액이 대기 상태인경우
            'Assign 되는 혈액의 해당환자에 해당하는 정보만 남겨두고 나머지 환자의 정보는 삭제한다.
            If strRstVal = "1" Or strStat = "1" Or strTestchk = "1" Then
            
                Dim RS As Recordset
                Set RS = objXM.GetAssignReadyBlood(strBldSrc, strBldYY, CStr(lngBldNo), strCompCd)
                If Not RS.EOF Then
                    Do Until RS.EOF
                        SSQL = objXM.DelAssignReadyBlood(RS.Fields("workarea").value & "", RS.Fields("accdt").value & "", _
                                                         RS.Fields("accseq").value & "", RS.Fields("rstseq").value & "")
                        DBConn.Execute SSQL
                        RS.MoveNext
                    Loop
                End If
                Set RS = Nothing
            End If
            '---------------------------------------------------------------------------------------------------------------
        Next
    End With

    DBConn.CommitTrans
    Insert_Sql = True
    MsgBox "정상적으로 결과등록 처리되었습니다.", vbInformation + vbOKOnly, "Cross_Matching 결과등록"
    Set objXM = Nothing
    Exit Function
    
XM_Result_Save_Error:
    
    If Insert_Sql = False Then
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    Set objXM = Nothing
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub lblaboapply_Click()
    Dim objSql As clsCrossMatching
    Dim strTmp As String
    Dim SSQL   As String
    
    strTmp = MsgBox("혈액형을 등록하시겠습니까?", vbInformation + vbYesNo, "혈액형 등록")
    
    If strTmp = vbNo Then
        fraABO.Visible = False
        Exit Sub
    End If
    
    If lblABO.Caption <> "" Then
        strTmp = MsgBox("이미혈액형이 등록되어 있습니다." & vbCrLf & " 수정하시겠습니까?", vbInformation + vbYesNo, "혈액형수정")
        If strTmp = vbNo Then
            fraABO.Visible = False
            Exit Sub
        End If
    End If
    
    If txtCABO = "" Then
        MsgBox "Cell ABO를 입력하세요", vbInformation + vbOKOnly, "혈액형입력"
        fraABO.Visible = False
        Exit Sub
    End If
    
    If txtRH = "" Then
        MsgBox "RH를 입력하세요.", vbInformation + vbOKOnly, "RH입력"
        fraABO.Visible = False
        Exit Sub
    End If
    
    On Error GoTo ABO_SAVE_ERROR
    DBConn.BeginTrans
    
    Set objSql = New clsCrossMatching
    
    SSQL = objSql.DeleteABO(lblPtId.Caption)
    DBConn.Execute SSQL
    
    SSQL = objSql.InsertABO(lblPtId.Caption, txtCABO.Text, txtSABO.Text, txtRH.Text)
    DBConn.Execute SSQL
    
    DBConn.CommitTrans
    
    lblABO.Caption = txtCABO.Text
    If txtSABO.Text <> "" Then lblABO.Caption = lblABO.Caption & "(" & txtSABO.Text & ")"
    lblABO.Caption = lblABO.Caption & txtRH.Text
    fraABO.Visible = False
    chkABO.value = 0
    Exit Sub
    
ABO_SAVE_ERROR:
    DBConn.RollbackTrans
    fraABO.Visible = False
    chkABO.value = 0
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub lblabocancel_Click()
    fraABO.Visible = False
    chkABO.value = 0
End Sub

Private Sub lblApply_Click()
    Dim Step(3) As String
    Dim ii      As Integer
    Dim ResultFg As Boolean
    
    For ii = 1 To lstResult.ListCount
        Step(ii - 1) = lstResult.Selected(ii - 1)
        If Step(ii - 1) = False Then
            ResultFg = True
        End If
    Next
    With tblBlood
        .Row = fraList.tag
        .Col = TblColumn1.tcSTEP1: .value = IIf(Step(0) = True, "1", "0")
        .Col = TblColumn1.tcSTEP2: .value = IIf(Step(1) = True, "1", "0")
        .Col = TblColumn1.tcSTEP3: .value = IIf(Step(2) = True, "1", "0")
        .Col = TblColumn1.tcSTEP4: .value = IIf(Step(3) = True, "1", "0")
        .Col = TblColumn1.tcDTAILRST: .value = ""
        
'모두 선택해제하고 Apply하면 Not으로 할것인지 결과입력안함으로 할것인지 여부를 물어봐
'2005/02/22 추가예정(이놈은 나중에 추가)

'        If Step(0) = False Or Step(1) = False Or Step(2) = False Or Step(3) = False Then
        If Step(0) = False And Step(1) = False And Step(2) = False And Step(3) = False Then
            onPgm = True
            .Col = TblColumn1.tcOK: .value = False
            .Col = TblColumn1.tcNot: .value = True
            onPgm = False
        Else
            onPgm = True
            .Col = TblColumn1.tcOK: .value = True
            .Col = TblColumn1.tcNot: .value = False
            onPgm = False
        End If
    End With
    
    txtBldNo.SetFocus
    fraList.Visible = False
End Sub

Private Sub lblCancel_Click()
    fraList.Visible = False
    txtBldNo.SetFocus
End Sub

Private Sub lstResult_ItemCheck(Item As Integer)
'아래에 있는 넘 선택하면 윗놈은 자동 선택되도록.. '
    Dim i As Integer
    
    If Item = 0 Then Exit Sub
    For i = 0 To Item - 1
        lstResult.Selected(i) = True
    Next
End Sub

'Private Sub LisLabel9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbLeftButton Then
'        fraList.Drag
'    End If
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
        '혈액 삭제
            With tblBlood
                .Row = .ActiveRow
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                InPutNo = InPutNo - 1
            End With
    End Select
End Sub

'Private Sub lblRmkfg_Click()
'    frmXMRemark.sPtid = lblPtId.Caption
'    frmXMRemark.rmk = lblrmk.Caption
'    frmXMRemark.Show
'End Sub
'
'Private Sub lblRmkFg_DblClick()
'    frmXMRemark.sPtid = lblPtId.Caption
'    frmXMRemark.rmk = lblrmk.Caption
'    frmXMRemark.Show
'End Sub

'Private Sub mnuDelete_Click()
''혈액 삭제
'    With tblBlood
'        .Row = .ActiveRow
'        .Action = ActionDeleteRow
'        .MaxRows = .MaxRows - 1
'        InPutNo = InPutNo - 1
'    End With
'End Sub

Private Sub tabData_Click()
   ' fradata(tabData.SelectedItem.Index - 1).ZOrder 0
    If tabData.SelectedItem.Index = 1 Then
        LisLabel3.Caption = "검사정보"
        fradata(1).ZOrder 0
    ElseIf tabData.SelectedItem.Index = 2 Then
        LisLabel3.Caption = "검체정보"
        fradata(0).ZOrder 0
    ElseIf tabData.SelectedItem.Index = 3 Then
        LisLabel3.Caption = "최근수혈정보"
        fradata(2).ZOrder 0
    End If
    
End Sub

Private Sub tblBlood_Click(ByVal Col As Long, ByVal Row As Long)
    Dim StepResult(3) As String
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean
    Dim ii As Integer
    
    If Row < 1 Then
        cmdTagPrint.Enabled = False
    Else
        cmdTagPrint.Enabled = True
    End If
    
    With tblBlood
        .Row = Row
        .Col = TblColumn1.tcRESULTFG
        '검사결과 입력 대기 인것...
        If .value = "1" Then
            '상세결과등록 인거....
            Select Case Col
                Case TblColumn1.tcDTAILRST
                   If txtComment.Visible = True Then
                        txtComment.Visible = False
                   End If
                   If Row > 8 Then
                        fraList.Top = 6395
                        fraList.Left = 10035
                   Else
                       Ret = .GetCellPos(TblColumn1.tcVFYDT, Row, X, Y, Wdt, Hgt)
                       Y = Y + Hgt
                       If .Height - Y < fraList.Height Or Y < 0 Then
                          Ret = .GetCellPos(TblColumn1.tcVFYDT, Row, X, Y, Wdt, Hgt)
                          fraList.Top = .Top + Y - fraList.Height + medMain.picMain.Height + 950
    
                          fraList.Left = .Left + X
                       Else
                          fraList.Left = .Left + X
                          fraList.Top = .Top + Y
                       End If
                   End If
                   
                   .Col = TblColumn1.tcBldNo: txtBloodNo = .value
                   .Col = TblColumn1.tcCOMPONM: txtCompcdnm = .value

                   .Col = TblColumn1.tcSTEP1: StepResult(0) = .value
                   .Col = TblColumn1.tcSTEP2: StepResult(1) = .value
                   .Col = TblColumn1.tcSTEP3: StepResult(2) = .value
                   .Col = TblColumn1.tcSTEP4: StepResult(3) = .value
                    For ii = 1 To lstResult.ListCount
                        lstResult.Selected(ii - 1) = IIf(StepResult(ii - 1) = "1", True, False)
                    Next
                    fraList.tag = Row
                    fraList.Visible = True
                Case TblColumn1.tcCMTBTN
                   If fraList.Visible = True Then
                        fraList.Visible = False
                    
                   End If
                   Ret = .GetCellPos(TblColumn1.tcSPCNO, Row, X, Y, Wdt, Hgt)
                   If Row <> .DataRowCnt Then
                        Y = Y + Hgt
                   Else
                        Y = Y ' + 200
                   End If
                   
                   If .Height - Y < txtComment.Height Or Y < 0 Then
                          Ret = .GetCellPos(TblColumn1.tcSPCNO, Row, X, Y, Wdt, Hgt)
                          txtComment.Top = .Top + Y - txtComment.Height + medMain.picMain.Height + 950
                          txtComment.Left = .Left + X

                   Else
                      txtComment.Left = .Left + X
                      txtComment.Top = .Top + Y
                   End If
                   .Col = TblColumn1.tcRMK
                   txtComment.Text = .value
                   txtComment.tag = Row
                   txtComment.Visible = True
                   txtComment.SetFocus
            End Select
        End If
    End With
    
    If Row = 0 And Col = TblColumn1.tcOK Then
        With tblBlood
            For ii = 1 To .MaxRows
                .Row = ii
                .Col = TblColumn1.tcOK
                If .CellType = CellTypeCheckBox Then .value = IIf(.value = 0, 1, 0)
            Next
        End With
    ElseIf Row = 0 And Col = TblColumn1.tcIRR Then
        With tblBlood
            
            For ii = 1 To .MaxRows
                .Row = ii
                .Col = TblColumn1.tcIRR
                If .CellType = CellTypeCheckBox Then .value = IIf(.value = 0, 1, 0)
            Next
        End With
    ElseIf Row = 0 And Col = TblColumn1.tcSTAT Then
        With tblBlood
            
            For ii = 1 To .MaxRows
                .Row = ii
                .Col = TblColumn1.tcSTAT
                If .CellType = CellTypeCheckBox Then .value = IIf(.value = 0, 1, 0)
            Next
        End With
    ElseIf Row = 0 And Col = TblColumn1.tcCMTBTN Then
        Dim strComment   As String
        Dim strCommentFg As String
        Dim strCFG        As String
        
        With tblBlood
            .Row = 1
            .Col = TblColumn1.tcCMTBTN:   strCommentFg = .value
            .Col = TblColumn1.tcRMK:      strComment = .value
            .Col = TblColumn1.tcRESULTFG: strCFG = .value
            If .value = "" Then Exit Sub
            For ii = 1 To .MaxRows
                .Row = ii
                .Col = TblColumn1.tcCMTBTN:   .value = strCommentFg
                .Col = TblColumn1.tcRMK:      .value = strComment
                .Col = TblColumn1.tcRESULTFG: .value = strCFG
            Next
            
        End With
        
    End If
End Sub


'Private Sub tblBlood_DragDrop(Source As Control, X As Single, Y As Single)
'    If Source = fraList Then
'        fraList.Left = X
'        fraList.Top = Y
'    End If
'End Sub

Private Sub tblBlood_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'마우스 오른쪽 버튼 클릭시 해당 라인의 Delete 기능 수행.
    If Row < 1 Then Exit Sub
    If blnStat = True Then Exit Sub
    
    Dim strTmp As String
    
    With tblBlood
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
        .Col = TblColumn1.tcRESULTFG
        If .value = "1" Then
            Set objPop = New clsPopupMenu
            With objPop
                .AddMenu MENU_DEL, "DELETE"
                .PopupMenus Me.hwnd
            End With
            Set objPop = Nothing
'            Set mnuPopup = frmControls.mnuPopup
'            Set mnuDelete = frmControls.mnuSub
'            mnuDelete.Caption = "Delete"
'
'            PopupMenu mnuPopup
'
'            Set mnuPopup = Nothing
'            Set mnuDelete = Nothing
        End If
    End With
End Sub

Private Function GetMaxRow() As Long
'    With tblResult
'        For GetMaxRow = 1 To .MaxRows
'            .Row = GetMaxRow
'            .Col = 2
'            If .value = "" Then
'                GetMaxRow = GetMaxRow - 1
'                Exit Function
'            End If
'        Next GetMaxRow
'    End With
End Function

Private Function GetBldNo() As String
    '입력된 혈액번호를 ##-##-#양식으로 반환한다.
    If chkBar.value = 1 Then
        GetBldNo = Mid(txtBldNo.Text, 1, 2) & "-" & Mid(txtBldNo.Text, 3, 2) & "-" & Mid(txtBldNo.Text, 5, 6)
    Else
        GetBldNo = txtBldNo.Text
    End If
End Function

Private Sub tblXM_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then
        cmdTagPrint.Enabled = False
    Else
        cmdTagPrint.Enabled = True
    End If
End Sub

Private Sub tblBlood_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim Step(3)  As String
    Dim strBldNo As String
    Dim strRmk   As String
    Dim strCoNm  As String
    
    With tblBlood
        If Row < 1 Then Exit Sub
        .Row = Row
        .Col = TblColumn1.tcRMK
        'If .value = "" Then
            Call .SetTextTipAppearance("굴림체", 10, False, False, &HEEFDF2, vbBlack)
            .Col = TblColumn1.tcBldNo:   strBldNo = .value
            .Col = TblColumn1.tcCOMPONM: strCoNm = .value
            .Col = TblColumn1.tcRMK:     strRmk = .value
            .Row = 0
            .Col = TblColumn1.tcSTEP1: Step(0) = .value
            .Col = TblColumn1.tcSTEP2: Step(1) = .value
            .Col = TblColumn1.tcSTEP3: Step(2) = .value
            .Col = TblColumn1.tcSTEP4: Step(3) = .value
            .Row = Row
            .Col = TblColumn1.tcSTEP1: Step(0) = Step(0) & IIf(.value = "1", "(Ok)", "(Not)")
            .Col = TblColumn1.tcSTEP2: Step(1) = Step(1) & IIf(.value = "1", "(Ok)", "(Not)")
            .Col = TblColumn1.tcSTEP3: Step(2) = Step(2) & IIf(.value = "1", "(Ok)", "(Not)")
            .Col = TblColumn1.tcSTEP4: Step(3) = Step(3) & IIf(.value = "1", "(Ok)", "(Not)")
            MultiLine = 1
            TipWidth = 7000
            TipText = vbNewLine & " 혈액번호 : " & strBldNo & vbNewLine & " Component: " & strCoNm & vbNewLine & _
                     " 상세결과 : " & Step(0) & "," & Step(1) & "," & Step(2) & "," & Step(3) & vbNewLine & _
                     " Comment  : " & strRmk & vbNewLine
            ShowTip = True
            
        'End If
    End With
End Sub



Private Sub tblOrder_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strAccDt As String
    Dim strRmk   As String
    Dim strComp  As String
    Dim strTest  As String
    
    With tblOrder
        .Row = Row
        .Col = 2: strTest = " 처방명   : " & .value
        .Col = 10: strAccDt = " 접수번호 : " & IIf(.value = "-0", "", .value)
        .Col = 11: strRmk = " Comment  : " & .value
        .Col = 7: strComp = " 혈액제제 : " & .value
        
        Call .SetTextTipAppearance("굴림체", 10, False, False, &HEEFDF2, vbBlack)
        MultiLine = 1
        TipWidth = 5000
        .Col = 15
        If .value = "Z" Then
            TipText = vbNewLine & strTest & vbNewLine
        Else
            .Col = 7
            TipText = vbNewLine & strAccDt & vbNewLine & strTest & _
                      vbNewLine & " 혈액제제 : " & .value & vbNewLine & strRmk & _
                      vbNewLine
        End If
        ShowTip = True
    End With
End Sub

Private Sub txtBldNo_Change()
    If chkBar.value = 1 Then Exit Sub
    Dim lngLen As Long
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_GotFocus()
    txtBldNo.SelStart = 0
    txtBldNo.SelLength = Len(txtBldNo)
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strBldNo As String
    Dim Row As Long
    
    If KeyCode = vbKeyReturn Then
        If chkBar.value = 1 Then
            If Len(txtBldNo.Text) < 6 Then
                txtBldNo = ""
                Exit Sub
            End If
        Else
            If Len(txtBldNo.Text) < 7 Then
                txtBldNo = ""
                Exit Sub
            End If
        End If
        strBldNo = GetBldNo
        'Row = GetMaxRow + 1
        
        strBNum = Replace(strBldNo, "-", "")
        '재결과등록인경우
        
        If CheckExist(strBNum, strBldNo) = False Then
            MsgBox "해당 혈액이 존재하지 않습니다.", vbExclamation
        End If
        
'        Call TblBloodInfomation(strBNum, strComponent, strBldNo)
            
        txtBldNo.SelStart = 0
        txtBldNo.SelLength = Len(txtBldNo)
    End If

'    Call SpreadCellBorder(tblBlood)
End Sub

Private Function CheckExist(ByVal vBldNo As String, ByVal pBldNo As String) As Boolean
    Dim objPopup As clsPopUpList
    Dim RS As Recordset
    Dim strSQL As String
    Dim strSrc As String
    Dim strYY As String
    Dim strNo As String
    
    strSrc = Mid(vBldNo, 1, 2)
    strYY = Mid(vBldNo, 3, 2)
    strNo = Mid(vBldNo, 5)
    
    strSQL = " SELECT a.compocd,b.abbrnm, a.abo,a.rh FROM S2BBS401 a, s2bbs006 b"
    strSQL = strSQL & " WHERE " & DBW("a.bldsrc=", strSrc)
    strSQL = strSQL & " AND " & DBW("a.bldyy=", strYY)
    strSQL = strSQL & " AND " & DBW("a.bldno=", strNo)
    strSQL = strSQL & " AND a.compocd=b.compocd "
    
    Set RS = New Recordset
    
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        CheckExist = False
    Else
        CheckExist = True
        '용량, 제제, 혈액형
        '제제가 다르면 용량도 다른건데...
        '혈액형은 TblBloodInfomation에서 체크하고 있고...
        
        If RS.RecordCount = 1 Then
            If RS.Fields("compocd").value & "" <> strComponent Then
                If MsgBox("수혈처방의 제제와 Assign할 혈액의 제제가 서로 다릅니다." & vbNewLine & vbNewLine & _
                          "이 혈액은 용량뿐만 아니라 혈액 종류가 다를 수도 있습니다." & vbNewLine & vbNewLine & vbNewLine & _
                          "이 혈액을 Assign하시겠습니까?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
                    Call TblBloodInfomation(vBldNo, RS.Fields("compocd").value & "", pBldNo)
                End If
            Else
                Call TblBloodInfomation(vBldNo, RS.Fields("compocd").value & "", pBldNo)
            End If
        Else
            Set objPopup = New clsPopUpList
            
            With objPopup
                .Recordset = RS
                    
                .ColumnHeaderText = "제제코드;약어;혈액형;rh"
                .HideColumnHeaders = True
                .SelectByClick = True
                .HideSearchTool = True
                .ColumnHeaderWidth = "374.7402;780.0945;329.9528;299.9055"
                .FormHeight = 1095
                .FormWidth = 2250
                .FormCaption = "제제선택"
                .LoadPopUp
                
                If .SelectedItems(0) <> "" Then
                    If .SelectedItems(0) <> strComponent Then
                        If MsgBox("수혈처방의 제제와 Assign할 혈액의 제제가 서로 다릅니다." & vbNewLine & vbNewLine & _
                                  "이 혈액은 용량뿐만 아니라 혈액 종류가 다를 수도 있습니다." & vbNewLine & vbNewLine & vbNewLine & _
                                  "이 혈액을 Assign하시겠습니까?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
                            Call TblBloodInfomation(vBldNo, .SelectedItems(0), pBldNo)
                        End If
                    Else
                        Call TblBloodInfomation(vBldNo, .SelectedItems(0), pBldNo)
                    End If
                End If
            End With
                                
            Set objPopup = Nothing
        End If
    End If
    
    Set RS = Nothing
End Function

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If txtSpcNO = "" Then Exit Sub
    
    If chkBar.value = 1 Then Exit Sub
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtCABO_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtComment_KeyDown(KeyCode As Integer, Shift As Integer)
    If Val(txtComment.tag) < 1 Or Val(txtComment.tag) > tblBlood.MaxRows Then Exit Sub
    If KeyCode = vbKeyReturn Then
        With tblBlood
            .Row = txtComment.tag
            .Col = TblColumn1.tcRESULTFG
            If .value = "1" Then
                .Col = TblColumn1.tcRMK
                .value = txtComment.Text
                If .value <> "" Then
                    .Col = TblColumn1.tcCMTBTN
                    .value = "Y"
                End If
            End If
        End With
        txtComment.Visible = False
    End If
End Sub

Private Sub txtLabelCnt_Change()
    If Trim(txtLabelCnt.Text) <> "" Then
        If IsNumeric(txtLabelCnt.Text) = False Then
            txtLabelCnt.Text = "2"
        End If
    End If
End Sub

Private Sub txtLabelCnt_LostFocus()
    If Trim(txtLabelCnt.Text) = "" Then
        txtLabelCnt.Text = "2"
    End If
End Sub

Private Sub txtSABO_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpcNo_GotFocus()
    txtSpcNO.tag = txtSpcNO
    
    txtSpcNO.SelStart = 0
    txtSpcNO.SelLength = Len(txtSpcNO)
End Sub

Private Sub txtSpcNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTmp As String
    
    strTmp = Mid(Format(GetSystemDate, "YYYY"), 1, 2)
    If KeyCode = vbKeyReturn Then
        If txtSpcNO.Text <> "" Then
            txtSpcNO.Text = strTmp & medGetP(txtSpcNO, 1, "-") & "-" & medGetP(txtSpcNO, 2, "-")
            txtSpcNO.tag = txtSpcNO.Text
            InPutNo = 0
            txtSpcNoLostFocus
        End If
    End If
End Sub

Private Sub txtSpcNo_Change()
    Dim lngLen As Long

    If lngAccDt = 4 Then
        With txtSpcNO
            lngLen = Len(Trim(.Text))
            If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    Else
        With txtSpcNO
            lngLen = Len(Trim(.Text))
            If lngLen = lngAccDt Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    End If
    
End Sub

Private Sub txtSpcNo_KeyPress(KeyAscii As Integer)
'접수번호 형식(####-##)
    If lngAccDt = 4 Then
        If Len(txtSpcNO) <> lngAccDt - 2 Then
            If KeyAscii = vbKeyInsert Then KeyAscii = 0
        End If
        
        If KeyAscii = vbKeyBack Then
            With txtSpcNO
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 2)
                    .SelStart = Len(.Text)
                    KeyAscii = 0
                End If
            End With
        End If
    Else
        If Len(txtSpcNO) <> lngAccDt Then
            If KeyAscii = vbKeyInsert Then KeyAscii = 0
        End If
        
        If KeyAscii = vbKeyBack Then
            With txtSpcNO
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 2)
                    .SelStart = Len(.Text)
                    KeyAscii = 0
                End If
            End With
        End If
    End If
    
End Sub

Private Sub txtSpcNo_LostFocus()
    If txtSpcNO = "" Then
        Clear
        tblOrder.MaxRows = 0
    Else
        If txtSpcNO.tag <> txtSpcNO Then
            txtSpcNO.tag = txtSpcNO
            InPutNo = 0
            Call txtSpcNoLostFocus
        End If
        txtBldNo.SetFocus
    End If
End Sub

Private Sub txtSpcNoLostFocus()
'접수번호를 가지고 정보를 찾는다.
        
    blnStat = False
    txtBldNo.Text = ""
    txtBldNo.Enabled = True
    tblBlood.MaxRows = 0
    tblSpc.MaxRows = 0
    tblOrder.MaxRows = 0
    
    tabData.Tabs.Item(1).Selected = True
    
    fradata(1).ZOrder 0
    LisLabel3.Caption = "검사정보"
    
    Me.MousePointer = 11
    If Find_Order(txtSpcNO) = False Then
        Call Clear
        txtSpcNO.SetFocus
        tblOrder.MaxRows = 0
    Else
'        SendKeys "{TAB}"
        If txtBldNo.Enabled Then txtBldNo.SetFocus
    End If
    
'    Call SpreadCellBorder(tblOrder)
'    Call SpreadCellBorder(tblBlood)
    
    Me.MousePointer = 0
    
End Sub

Private Sub DetailSearch(Ptid As String, OrdDt As String)
'혈액형,부작용,감염정보,상병코드,상병을 조회한다.
    Dim ObjABO As New clsABO
    
    Dim objinfection As New clsInfection
    Dim objReaction As New clsReaction
    
    With ObjABO
        .Ptid = Ptid
        If .GetABO = True Then
            lblABO.Caption = .ABO & .Rh
        Else
            lblABO.Caption = ""
        End If
    End With
    With objinfection
        .Ptid = Ptid
        .GetInfection
        If .Infection = True Then
            lblInfection.Visible = True
        Else
            lblInfection.Visible = False
        End If
    End With
    
    With objReaction
        .Ptid = Ptid
        If .GetReaction = True Then
            lblReaction.Visible = .Reaction
        Else
            lblReaction.Visible = False
        End If
    End With
    
    
    Set objReaction = Nothing
    Set objinfection = Nothing
    Set ObjABO = Nothing

End Sub

Private Function Find_Order(ByVal AccdtSeq As String) As Boolean
'--------------------------------------------------------------
'접수번호를 가지고 결과등록에 필요한 정보를 보여준다.(처방정보)
'--------------------------------------------------------------
    Dim objProBar  As New clsProgress
    Dim objCollect As clsQueryOrder
    Dim objXM      As New clsCrossMatching
    Dim RS         As Recordset
    Dim strAccDt   As String               '접수일자
    Dim lngAccSeq  As Long                 '접수번호
    Dim strOrdCd   As String
    Dim strTmp     As String
    Dim lngOrdCnt  As Integer
    Dim strReason  As String
    
    Dim ii         As Integer
    
    strAccDt = Mid(AccdtSeq, 1, lngAccDt)
    lngAccSeq = Val(Mid(AccdtSeq, lngAccDt + 2))
    
    With objXM
        Set RS = .Get_XM_Blood_List(strAccDt, lngAccSeq)
    End With
    
    
'    Set objProBar.MyForm = Me
'    Set objProBar.StatusBar = medMain.stsBar
    objProBar.Container = MainFrm.stsBar
    objProBar.Max = 100
    For ii = 1 To 20
        objProBar.value = ii
    Next

    '----------------------------------------
    '해당 처방일자의 처방을 가지고 온다......
    '----------------------------------------
    Dim FirstChk As Boolean
    
    Dim RealOrdno As Long
    
    Dim jj       As Integer
    
    
    With tblOrder
'        RS.MoveFirst
        
        .MaxRows = RS.RecordCount
        If Not RS.EOF Then
            Set objCollect = New clsQueryOrder
            
            Do Until RS.EOF
                jj = jj + 1
                .Row = jj
                If FirstChk = False Then
                    strPtid = RS.Fields("ptid").value & ""
                    strOrdDt = RS.Fields("orddt").value & ""
                    strWardID = RS.Fields("wardid").value & ""
                    strDeptCd = RS.Fields("deptcd").value & ""
                    FirstChk = True
                End If
                
                RealOrdno = Val(RS.Fields("ordno").value & "")
                .Col = 2: .value = RS.Fields("testnm").value & ""
                
                .Col = 3: .value = RS.Fields("unitqty").value & ""
                .ForeColor = DCM_Magenta
                .FontBold = True
                
                .Col = 4: .value = Format(RS.Fields("orddt").value & "", "####-##-##")
                .Col = 5: .value = Format(RS.Fields("reqdt").value & "", "####-##-##")
                strReason = objCollect.GetTransReason(strPtid, strOrdDt, CStr(RealOrdno))
                .Col = 6: .value = strReason
                
                strTmp = objXM.Get_BCNm(RS.Fields("ordcd").value & "")
                .Col = 7: .value = medGetP(strTmp, 2, COL_DIV): .ForeColor = DCM_LightBlue
                
                .Col = 8: .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", ""): .ForeColor = DCM_LightRed
                .Col = 10: .value = RS.Fields("accdt").value & "" & "-" & RS.Fields("accseq").value & ""
                
                '실제 접수번호..
                If .value = txtSpcNO.Text Then
                    lngOrdNo = Val(RS.Fields("ordno").value & "")
                    strComponent = RS.Fields("compocd").value & ""
                    strComponentNm = medGetP(strTmp, 2, COL_DIV)
                    lngUnitQty = Val(RS.Fields("unitqty").value & "")
                    lngOrdseq = Val(RS.Fields("ordseq").value & "")
                End If
                .Col = 11: .value = RS.Fields("mesg").value & ""
                If .value <> "" Then
                    .Col = 9: .value = "Y": .ForeColor = DCM_LightRed
                End If
                
                .Col = 12: .value = Val(RS.Fields("xmethod").value & "")
                .Col = 13: .value = medGetP(strTmp, 1, COL_DIV)
                .Col = 14: .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", ""): .ForeColor = DCM_LightRed
                .Col = 15: .value = RS.Fields("orddiv").value & ""
                
                RS.MoveNext
            Loop
            
            Call DetailSearch(strPtid, strOrdDt)
            
            Dim objDisease As New clsDisease
            
            With objDisease
                .Ptid = strPtid
'                .OrdDt = strOrdDt
'온승호
'2010년5월14일
'날짜포멧수정
                .OrdDt = Format(strOrdDt, "####-##-##")
                .ordno = CStr(RealOrdno)
                If .GetDisease = True Then
                    lblSickCd.Caption = .DiseaseCd      '상병코드
                    lblSick.Caption = .DiseaseNm        '상병명
                Else
                    lblSickCd.Caption = ""
                    lblSick.Caption = ""
                End If
            End With
            
            Set objDisease = Nothing
        Else
            MsgBox "접수번호에 해당하는 정보가 없습니다." & vbNewLine & _
                   "확인후 등록하세요.", vbCritical + vbOKOnly, Me.Caption
            Me.MousePointer = 0
            Set RS = Nothing
            Set objXM = Nothing
            Exit Function
        End If
    End With
    Set RS = Nothing
            
    
    Call CurrentAccDtDiv
    
    For ii = 21 To 40
        objProBar.value = ii
    Next
    
    '--------------
    '환자정보를 Get
    '--------------
    Call Find_PtInFo(strPtid, strOrdDt, lngOrdNo)
    
    '-------------------
    '환자결과 Remark Get
    '-------------------
    Call Find_PtRemark(strPtid)
    
    '----------------------------
    '기존의 검사결과를 가지고온다
    '----------------------------
    For ii = 41 To 70
        objProBar.value = ii
    Next
    tblBlood.MaxRows = 0
    Call ResultHistory(strAccDt, CStr(lngAccSeq))

    Call LastTransInfo(strPtid)
    
    '----------------------------------------------------------------
    '혈액의 준비사항을 봐서, 혈액이 모두 준비되었으면 메시지를 보낸다
    '----------------------------------------------------------------
    For ii = 71 To 100
        objProBar.value = ii
    Next
    Call GetTestInformation
    Call LookUpAssignBloodCount(strAccDt, CStr(lngAccSeq))
    
    Find_Order = True
    
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    
    Set objXM = Nothing
    Set objProBar = Nothing
    Set objCollect = Nothing
End Function

Private Sub LastTransInfo(ByVal Ptid As String)
    Dim objSql As New clsCrossMatching
    Dim SSQL   As String
    Dim RS     As Recordset
    
    lblLastBldNo.Caption = "": lblLastComp.Caption = "": lblTransDt.Caption = ""
    
    SSQL = objSql.LastTransInfo(Ptid)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        lblLastBldNo.Caption = RS.Fields("bldsrc").value & "" & "-" & _
                               RS.Fields("bldyy").value & "" & "-" & _
                               Format(RS.Fields("bldno").value & "", "000000") & "   (" & RS.Fields("abo").value & "" & RS.Fields("rh").value & "" & ")"
        lblLastComp.Caption = RS.Fields("componm").value & ""
        lblTransDt.Caption = Format(RS.Fields("deliverydt").value & "", "0###-##-##") & "   " & _
                             Format(RS.Fields("deliverytm").value & "", "0#:##:##")
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub GetTestInformation()
    Dim objSql As New clsCrossMatching
    Dim RS     As Recordset
    Dim SSQL   As String
    Dim ii     As Integer
    
    'medClearTable tblTest
    tblTest.MaxRows = 0
    
    SSQL = objSql.TestResultXM(strPtid)
    If SSQL <> "" Then
    Set RS = New Recordset
    RS.Open SSQL, DBConn
        If Not RS.EOF Then
            With tblTest
                .MaxRows = RS.RecordCount
                 Do Until RS.EOF
                    ii = ii + 1
                    .Row = ii
                    .Col = 1: .value = RS.Fields("workarea").value & "" & "-" & Mid(RS.Fields("accdt").value & "", 3) & "-" & RS.Fields("accseq").value & ""
                    .Col = 2: .value = RS.Fields("abbrnm10").value & ""
                    .Col = 3: .value = RS.Fields("RstCdNm").value & ""
                                        
                    'Abnormal 결과인 경우 붉게 표시
                    .Row2 = ii + 1
                    .COL2 = 3
                    .BlockMode = True
                    If InStr(UCase(.value), "P") > 0 Then
                        .ForeColor = vbRed
                        .Font.Bold = True
                    Else
                        .ForeColor = vbBlack
                        .Font.Bold = False
                    End If
                    .BlockMode = False
                                        
                    .Col = 4: .value = RS.Fields("rstunit").value & ""
                    RS.MoveNext
                Loop
         End With
        End If
        Set RS = Nothing
    End If
    Set objSql = Nothing
End Sub

Private Sub Find_PtRemark(ByVal Ptid As String)
    Dim objSql As New clsCrossMatching
    
    lblrmk.Caption = objSql.GetptidRmk(Ptid)
    
    If lblrmk.Caption <> "" Then
        cmdRmk.Caption = "Y"
        cmdRmk.Visible = True
'    Else
'        cmdRmk.Caption = ""
'        cmdRmk.Visible = False
    End If
    Set objSql = Nothing
End Sub

Private Sub tblOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    With tblOrder
        .Row = Row
        .Col = 15
        If .value = "Z" Then Exit Sub
        Call Clear
        .Col = 10: txtSpcNO = .value
        Call txtSpcNoLostFocus
        txtBldNo.SetFocus
    End With
End Sub

Private Sub CurrentAccDtDiv()
    Dim ii As Integer
    
    With tblOrder
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1: .value = ""
        Next
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 10
            If .value = txtSpcNO Then
                .Col = 1:  .value = CurrentSelected: .ForeColor = DCM_LightRed
                .Col = 12: cboMethod.ListIndex = Val(.value)
                Exit For
            End If
        Next
    End With
End Sub
Private Function LookUpAssignBloodCount(ByVal accdt As String, ByVal accseq As String)
    Dim objXM As clsCrossMatching
    Dim A_Cnt As Long   'Assign수량
    Dim C_Cnt As Long   'Assign Cancel 수량
    Dim O_Cnt As Long   '출고수량
    Dim R_Cnt As Long   '반환수량
    Dim X_Cnt As Long   '폐기수량
    Dim T_Cnt As Long   '총Assign 수량
    Dim unitqty As Long


    With tblOrder
        .Row = 1
        .Col = 3
         unitqty = Val(.value)
    End With
    
    Set objXM = New clsCrossMatching
    With objXM
        .Assign_Cnt accdt, accseq
        A_Cnt = .AssignCnt
        C_Cnt = .CancelCnt
        O_Cnt = .OutCnt
        R_Cnt = .RetCnt
        X_Cnt = .ExpCnt
    End With
    Set objXM = Nothing
    
    T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt

    If T_Cnt >= unitqty Then
        lblLog.Visible = True
    Else
        lblLog.Visible = False
    End If
End Function

Private Function ResultHistory(ByVal accdt As String, ByVal accseq As String)
    Dim objXM      As clsCrossMatching
    Dim DrRS       As New Recordset
    Dim DrRsOut    As New Recordset
    Dim strCompocd As String
    Dim strCompoNm As String
    Dim strBldNo   As String
    Dim spcyy      As String
    Dim spcno      As String
    Dim ii         As Integer
    Dim jj         As Integer
    
    Set objXM = New clsCrossMatching
    Set DrRS = New Recordset
    Set DrRsOut = New Recordset
    '---------------------------------------------
    '처방에대해서 이미 결과등록 History를 보여준다
    '---------------------------------------------
    
    DrRS.Open objXM.Get_Collect_AssignList(accdt, accseq), DBConn
    If DrRS.EOF = False Then
        
        
        With tblBlood
            .MaxRows = 0
            Do Until DrRS.EOF = True
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                strBldNo = Trim(DrRS.Fields("bldsrc").value & "") & "-" & _
                           Trim(DrRS.Fields("bldyy").value & "") & "-" & _
                           Format(DrRS.Fields("bldno").value & "", "00000#")
                strCompoNm = DrRS.Fields("field1").value & ""
                strCompocd = DrRS.Fields("compocd").value & ""
                .Col = TblColumn1.tcBldNo:   .value = strBldNo
                .Col = TblColumn1.tcCOMPONM: .value = strCompoNm
                .Col = TblColumn1.tcABO:     .value = DrRS.Fields("abo").value & "" & DrRS.Fields("rh").value & ""
                .Col = TblColumn1.tcVol:     .value = CLng(DrRS.Fields("volumn").value & "")
                
                Select Case DrRS.Fields("stat").value & ""
                    Case 1:    .Col = TblColumn1.tcSTAT:  .value = "1"
                               .Col = TblColumn1.tcVFYDT: .value = Format(DrRS.Fields("statdt").value & "", "####-##-##")
                               .Col = TblColumn1.tcASSIGN:       .value = "1"
                               .Col = TblColumn1.tcVFYNM: .value = GetEmpNm(DrRS.Fields("statid").value & "")
                               .Col = TblColumn1.tcVfyTm: .value = Format$(Mid$(DrRS.Fields("stattm").value & "", 1, 4), "0#:#")
                    Case Else: .Col = TblColumn1.tcSTAT:  .value = "0"
                               .Col = TblColumn1.tcVFYDT: .value = Format(DrRS.Fields("vfydt").value & "", "####-##-##")
                               .Col = TblColumn1.tcVFYNM: .value = DrRS.Fields("empnm").value & ""
                               .Col = TblColumn1.tcVfyTm: .value = Format$(Mid$(DrRS.Fields("vfytm").value & "", 1, 4), "0#:0#")
                End Select
                
                Select Case DrRS.Fields("rstv").value & ""
                    '이미 검사판정이 난경우(OK)
                    Case 1: .Col = TblColumn1.tcOK:    .value = "1"
                            .Col = TblColumn1.tcSTEP1: .value = DrRS.Fields("step1").value & ""
                            .Col = TblColumn1.tcSTEP2: .value = DrRS.Fields("step2").value & ""
                            .Col = TblColumn1.tcSTEP3: .value = DrRS.Fields("step3").value & ""
                            .Col = TblColumn1.tcSTEP4: .value = DrRS.Fields("step4").value & ""
                            .Col = TblColumn1.tcVFYNM: .value = DrRS.Fields("empnm").value & ""
                    '검사판정(NOT)
                    Case 0: .Col = TblColumn1.tcNot: .value = "1"
                            .Col = TblColumn1.tcSTEP1: .value = DrRS.Fields("step1").value & ""
                            .Col = TblColumn1.tcSTEP2: .value = DrRS.Fields("step2").value & ""
                            .Col = TblColumn1.tcSTEP3: .value = DrRS.Fields("step3").value & ""
                            .Col = TblColumn1.tcSTEP4: .value = DrRS.Fields("step4").value & ""
                            .Col = TblColumn1.tcVFYNM: .value = DrRS.Fields("empnm").value & ""
                End Select
                
                
                'irradiation 처리여부......
                .Col = TblColumn1.tcIRR:    .value = IIf(DrRS.Fields("irrfg").value & "" = "1", "1", "")
                .Col = TblColumn1.tcFilter: .value = IIf(DrRS.Fields("filterfg").value & "" = "1", "1", "")
                
                .Col = TblColumn1.tcRMK: .value = DrRS.Fields("rmk").value & ""
                If Trim(.value) <> "" Then
                    .Col = TblColumn1.tcCMTBTN: .value = "Y": .ForeColor = vbRed
                End If
                    
                .Col = TblColumn1.tcSPCNO: .value = DrRS.Fields("spcyy").value & "-" & _
                                                    Format(DrRS.Fields("spcno").value & "", "#########")
                
                
                '혈액의 상태를 보여주자-------------------------------------------------------
                If DrRS.Fields("cancelfg").value & "" = "1" Then
                    .Col = TblColumn1.tcDELIVERYDT: .value = ""
                    .Col = TblColumn1.tcSTATUS:     .value = "취소"
                    .Col = TblColumn1.tcFLAG:       .value = "1"
                    .Col = TblColumn1.tcDUP:        .value = Replace(strBldNo, "-", "") & COL_DIV & strCompocd
                    .Col = TblColumn1.tcASSIGN:      .value = "0"
                ElseIf DrRS.Fields("norstfg").value & "" = "1" Then
                    .Col = TblColumn1.tcDELIVERYDT: .value = ""
                    .Col = TblColumn1.tcSTATUS:      .value = "PHER"
                Else
                    Select Case objXM.Get_Blood_Status(accdt, accseq, DrRS.Fields("rstseq").value & "")
                        Case BBSBloodStatus.stsASSIGN
                            .Col = TblColumn1.tcDELIVERYDT:  .value = ""
                            If DrRS.Fields("rstv").value & "" = "1" Then 'Or DrRS.Fields("rstv").value & "" = "" Then
                                .Col = TblColumn1.tcSTATUS:      .value = "A"
                                .Col = TblColumn1.tcASSIGN:      .value = "1"
                            ElseIf DrRS.Fields("rstv").value & "" = "0" Then
                                .Col = TblColumn1.tcSTATUS:      .value = "Not"
                            ElseIf DrRS.Fields("rstv").value & "" = "" And DrRS.Fields("stat").value & "" = "" Then
                                .Col = TblColumn1.tcASSIGN:      .value = "1"
                                .Col = TblColumn1.tcDUP:         .value = Replace(strBldNo, "-", "") & COL_DIV & strCompocd
                                '결과없이 입력된거 표시(update위해서)
                                .Col = TblColumn1.tcNORSV:       .value = "1"
                               ' .Col = TBLCOLUMN1.tcASSIGN:       .value = "1"
                            End If
                        Case BBSBloodStatus.stsDELIVERY
                            Set DrRsOut = Nothing
                            Set DrRsOut = New Recordset
                            DrRsOut.Open objXM.Get_Delivery(accdt, accseq, DrRS.Fields("rstseq").value & ""), DBConn
                            If DrRsOut.RecordCount > 0 Then
                                .Col = TblColumn1.tcDELIVERYDT: .value = Format(DrRsOut.Fields("deliverydt").value & "", "####-##-##")
                            End If
                            Set DrRsOut = Nothing
                            .Col = TblColumn1.tcSTATUS:      .value = "출고"
                            .Col = TblColumn1.tcASSIGN:      .value = "1"
                        Case BBSBloodStatus.stsRETURN
                            Set DrRsOut = Nothing
                            Set DrRsOut = New Recordset
                            DrRsOut.Open objXM.Get_Delivery(accdt, accseq, DrRS.Fields("rstseq").value & ""), DBConn
                            If DrRsOut.RecordCount > 0 Then
                                .Col = TblColumn1.tcDELIVERYDT: .value = Format(DrRsOut.Fields("deliverydt").value & "", "####-##-##")
                            End If
                            Set DrRsOut = Nothing
                            .Col = TblColumn1.tcSTATUS:       .value = "반환"
                            .Col = TblColumn1.tcFLAG:        .value = "1"
                            
                            .Col = TblColumn1.tcDUP: .value = Replace(strBldNo, "-", "") & COL_DIV & strCompocd
                        Case BBSBloodStatus.stsEXPIRE
                            Set DrRsOut = Nothing
                            Set DrRsOut = New Recordset
                            DrRsOut.Open objXM.Get_Delivery(accdt, accseq, DrRS.Fields("rstseq").value & ""), DBConn
                            If DrRsOut.RecordCount > 0 Then
                                .Col = TblColumn1.tcDELIVERYDT: .value = Format(DrRsOut.Fields("deliverydt").value & "", "####-##-##")
                            End If
                            Set DrRsOut = Nothing
                            .Col = TblColumn1.tcSTATUS:       .value = "폐기"
                            .Col = TblColumn1.tcASSIGN: .value = ""
                        Case BBSBloodStatus.stsBAG
                            Set DrRsOut = Nothing
                            Set DrRsOut = New Recordset
                            DrRsOut.Open objXM.Get_Delivery(accdt, accseq, DrRS.Fields("rstseq").value & ""), DBConn
                            If DrRsOut.RecordCount > 0 Then
                                .Col = TblColumn1.tcDELIVERYDT: .value = Format(DrRsOut.Fields("deliverydt").value & "", "####-##-##")
                            End If
                            Set DrRsOut = Nothing
                            .Col = TblColumn1.tcSTATUS:       .value = "회수"
                        Case Else
                            .Col = TblColumn1.tcDELIVERYDT:  .value = ""
                            .Col = TblColumn1.tcSTATUS:       .value = ""
                    End Select
                End If
                .Col = TblColumn1.tcABBRNM:  .value = DrRS.Fields("abbrnm").value & ""
                .Col = TblColumn1.tcCOMPOCD: .value = strCompocd
                .Col = TblColumn1.tcRSTSEQ:  .value = DrRS.Fields("rstseq").value & ""
                
                DrRS.MoveNext
            Loop
            '----------------------------------------
            '응급인거 재결과등록을 위해서 색으로 구분
            '----------------------------------------
            Dim OkTF  As Boolean
            Dim NotTF As Boolean
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = TblColumn1.tcSTAT
                OkTF = False: NotTF = False
                If .value = "1" Then
                    '응급이지만 판정이 난거........
                    .Col = TblColumn1.tcOK
                    If .value = True Then
                        For jj = 1 To .MaxCols
                            .Col = jj
                            .ForeColor = vbBlack
                        Next
                        .Col = TblColumn1.tcRESULTFG: .value = "0"
                    Else
                        OkTF = True
                    End If
                    
                    .Col = TblColumn1.tcNot
                    If .value = "1" Then
                        For jj = 1 To .MaxCols
                            .Col = jj
                            .ForeColor = vbBlack
                        Next
                    Else
                        NotTF = True
                    End If
                    
                    If OkTF = True And NotTF = True Then
                        For jj = 1 To .MaxCols
                            .Col = jj
                            .ForeColor = vbRed
                        Next
                        
                        
                        .Col = TblColumn1.tcRESULTFG: .value = "1"
                    End If
                Else
                    .Col = TblColumn1.tcOK:
                    If .value = True Then
                        .Col = TblColumn1.tcRESULTFG: .value = "0"
                        .Col = TblColumn1.tcFLAG
'                        If .value <> "1" Then
'                            .Col = TblColumn1.tcASSIGN:   .value = "1"
'                        End If
                        For jj = 1 To .MaxCols
                            .Col = jj
                            .ForeColor = vbBlack
                        Next
                    Else
                        .Col = TblColumn1.tcNot
                        If .value = True Then
                            For jj = 1 To .MaxCols
                                .Col = jj
                                .ForeColor = vbBlack
                            Next
                            .Col = TblColumn1.tcRESULTFG: .value = "0"
                            .Col = TblColumn1.tcASSIGN:   .value = "0"
                        Else
                            .Col = TblColumn1.tcRESULTFG: .value = "1"
                            For jj = 1 To .MaxCols
                                .Col = jj
                                .ForeColor = vbBlue
                            Next
                        End If
                    End If
                    
                    For jj = 1 To .MaxCols
                        .Col = jj
                        .ForeColor = vbBlack
                    Next
                End If
            Next
            '----------------------------------------------------
            '판정이 난거는 결과등록을 못하게 Lock을 걸자.........
            '----------------------------------------------------
            For ii = 1 To .DataRowCnt
                .Row = ii
                
                .Col = TblColumn1.tcRESULTFG
                If .value <> "1" Then
                    .Row = ii: .Row2 = ii
                    .Col = TblColumn1.tcSTAT: .COL2 = TblColumn1.tcIRR
                    .BlockMode = True
                    .Lock = True
                    .CellType = CellTypeStaticText
                    .BlockMode = False
                    
                    For jj = TblColumn1.tcSTAT To TblColumn1.tcIRR
                        .Row = ii
                        .Col = jj
                        If jj = TblColumn1.tcIRR Then
                            If .value = "1" Then
                                .value = "√": .ForeColor = vbRed: .TypeHAlign = TypeHAlignCenter
                            Else
                                .CellType = CellTypeCheckBox:      .TypeHAlign = TypeHAlignCenter: .Lock = False
                            End If
                        Else
                            If .value = "1" Then
                                .value = "√": .ForeColor = vbRed
                                .TypeHAlign = TypeHAlignCenter
                            Else
                                .value = ""
                            End If
                        End If
                    Next jj
                    
                Else
                    .Col = TblColumn1.tcSTAT:
                    If .value = "1" Then
                        .CellType = CellTypeStaticText
                        .value = IIf(.value = "1", "√", ""): .ForeColor = DCM_LightRed
                        .TypeHAlign = TypeHAlignCenter
                    End If
                    
                    .Col = TblColumn1.tcIRR: If .value = "1" Then .Lock = True
                End If
                
                .Col = TblColumn1.tcASSIGN
                If .value = "1" Then
                    InPutNo = InPutNo + 1
                    .Col = TblColumn1.tcNo: .value = InPutNo
                Else
                    .Col = TblColumn1.tcNo: .value = "**"
                End If
                .Col = TblColumn1.tcCMTBTN:
                If .value = "Y" Then .ForeColor = vbRed
            Next
            
            '처방제제와 Assign한 제제가 다른 경우 표시
            Dim strOrdComponm As String
            
            With tblOrder
                For ii = 1 To .DataRowCnt
                    .Row = ii
                    .Col = 1
                    If .value = CurrentSelected Then
                        .Col = 7: strOrdComponm = .value
                        Exit For
                    End If
                Next
            End With
            
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = TblColumn1.tcCOMPONM
                strCompoNm = .value
                
                If strCompoNm <> strOrdComponm Then
                    .ForeColor = DCM_Magenta
                    .FontBold = True
                Else
                    .ForeColor = vbBlack
                    .FontBold = False
                End If
            Next
            
            .ReDraw = True
        End With
    Else
        blnStat = False
        txtBldNo.Enabled = True
    End If
    
    Set DrRsOut = Nothing
    Set DrRS = Nothing
    Set objXM = Nothing
End Function


Private Function Find_PtInFo(ByVal Ptid As String, ByVal OrdDt As String, ByVal ordno As Long)
'환자와 검체정보를 조회한다.
    Dim objXM    As New clsCrossMatching
    Dim DrRS     As New Recordset
    Dim objSql   As clsGetSqlStatement
    Dim strTmp   As String
    Dim Timechk  As Long
    Dim ii       As Integer: ii = 0
    Dim KeepOur  As Long
    
    Dim objQuery As New clsQueryOrder
    
    KeepOur = objQuery.GetKeepHour
    
    Set objQuery = Nothing
    
    
    With objXM
'        .setDbConn DBConn
        strTmp = .Get_PtInfo(Ptid, OrdDt, ordno)
        If strTmp <> "" Then
            lblPtId.Caption = Ptid
            lblPtNm.Caption = medGetP(strTmp, 1, COL_DIV)
            lblSexAge.Caption = medGetP(strTmp, 2, COL_DIV)
            lblDeptNm.Caption = medGetP(strTmp, 3, COL_DIV)
            lblWardNm.Caption = medGetP(strTmp, 4, COL_DIV)
            strSSN = Mid(medGetP(strTmp, 5, COL_DIV), 1, 6) & "-" & Mid(medGetP(strTmp, 5, COL_DIV), 7)
            If medGetP(strSSN, 2, "-") <> "" Then
                strSSN = medGetP(strSSN, 1, "-") & "-" & Mid(medGetP(strSSN, 2, "-"), 1, 4) & "xxx"
            End If
            
                
        End If
        Set DrRS = .Get_SpcInfo(Ptid, OrdDt)
    End With
    
    With tblSpc
        'medClearTable tblSpc
        tblSpc.MaxRows = 0
        If DrRS.EOF = False Then
            .MaxRows = DrRS.RecordCount
            
            Set objSql = New clsGetSqlStatement
'            objSql.setDbConn DBConn
            
            Timechk = objSql.Spc_TimeChk(Ptid)
            If Timechk > KeepOur Then
                lblAddChk.ForeColor = vbRed
                lblAddChk.Caption = "검체채취 경과시간: " & Timechk & " 시간"
            Else
                lblAddChk.ForeColor = vbBlue
                lblAddChk.Caption = "검체채취 경과시간: " & Timechk & " 시간"
            End If
            
            Set objSql = Nothing
            
            Do Until DrRS.EOF = True
                ii = ii + 1
                .Row = ii
                .Col = 1: SpcNum = DrRS.Fields("spcyy").value & "" & "-" & DrRS.Fields("spcno").value & ""
                          .value = UCase(SpcNum)
                .Col = 2: .value = DrRS.Fields("storeleg").value & "" & _
                                   "(" & DrRS.Fields("storerno").value & "" & _
                                   "," & DrRS.Fields("storecno").value & "" & ")"
                .Col = 3: .value = Format(DrRS.Fields("coldt").value & "", "####-##-##")
                .Col = 4: .value = IIf(DrRS.Fields("expfg").value & "" = "1", "폐기", "") & "(" & IIf(DrRS.Fields("addfg").value & "" = "1", "추.검", "") & ")"
                DrRS.MoveNext
            Loop
        Else
            lblAddChk.Caption = "검체가 존재하지 않습니다."
        End If
    End With
    
    Set DrRS = Nothing
    Set objXM = Nothing
End Function

Private Function BloodDupChk(ByVal pBldNo As String) As Boolean
'중복값을 체크한다.(true:dup)
    Dim ii As Integer
    
    With tblBlood
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn1.tcRESULTFG
            If .value = "1" Then
                .Col = TblColumn1.tcDUP
                If .value = pBldNo Then
                    BloodDupChk = True
                    MsgBox "이미 결과등록대기중인 혈액입니다.", vbInformation + vbOKOnly, "XM결과등록"
                    Exit Function
                End If
            End If
        Next
    End With
    
End Function

Private Function UnitQtyChk() As Boolean
    Dim UnitCnt As Integer
    Dim ii      As Integer
    
    With tblBlood
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = TblColumn1.tcASSIGN
            If .value = "1" Then
                UnitCnt = UnitCnt + 1
                
            Else
                .Col = TblColumn1.tcRESULTFG
                If .value = "1" Then
                    UnitCnt = UnitCnt + 1
                End If
            End If
        Next
    End With
    
    If UnitCnt >= lngUnitQty Then
        UnitQtyChk = True
        MsgBox "이미 처방에대한 혈액이 결과등록 대기중입니다.", vbInformation + vbOKOnly, "결과등록"
    End If
    
End Function

Private Sub TblBloodInfomation(ByVal BloodNum As String, ByVal compcd As String, ByVal BldNo As String)
    Dim objXM       As clsCrossMatching
    Dim strABO      As String
    Dim strTmp      As String
    Dim strBloodTmp As String
    Dim ii          As Integer
    
    
    If UnitQtyChk = True Then Exit Sub
    Set objXM = New clsCrossMatching
    
    '반환여부 확인처리
    strBloodTmp = objXM.Get_BloodStsCD(BloodNum, compcd, ObjSysInfo.BuildingCd)
    If strBloodTmp = CStr(BBSBloodStatus.stsRETURN) Then
        strTmp = MsgBox("반환처리되었던 혈액입니다. 계속 진행하시겠습니까?", vbInformation + vbYesNo)
        If strTmp = vbNo Then
            Set objXM = Nothing
            Exit Sub
        End If
    End If
    
    If objXM.Get_BloodINfo(BloodNum, compcd, ObjMyUser.EmpId, ObjSysInfo.BuildingCd, lblPtId.Caption) = False Then
        Set objXM = Nothing
        Exit Sub
    End If
    
    If Len(lblABO.Caption) > 3 Then
        strABO = medGetP(lblABO.Caption, 1, "(") & medGetP(lblABO.Caption, 2, ")")
    Else
        strABO = lblABO.Caption
    End If
    
    '----------
    '혈액형비교
    '----------
    If strABO <> medGetP(objXM.strTmp, 1, vbTab) Then
        strTmp = MsgBox("환자 혈액형과 혈액의 혈액형이 동일하지 않습니다." & vbCrLf & "결과등록을 계속진행하시겠습니까?", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption)
        If strTmp = vbNo Then
            Set objXM = Nothing
            Exit Sub
        End If
    End If
    '-----------------------
    '헌혈부적격 판정여부체크
    '-----------------------
    
    Dim strCompoNm As String
    
    With tblOrder
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 1
            If .value = CurrentSelected Then
                .Col = 7: strCompoNm = .value
                Exit For
            End If
        Next
    End With
    
    With tblBlood
        If BloodDupChk(medGetP(BldNo, 1, "-") & medGetP(BldNo, 2, "-") & Format(medGetP(BldNo, 3, "-"), "00000#") & COL_DIV & compcd) = True Then Exit Sub
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = TblColumn1.tcABO:     .value = medGetP(objXM.strTmp, 1, vbTab)
        .Col = TblColumn1.tcVol:     .value = medGetP(objXM.strTmp, 2, vbTab)
        .Col = TblColumn1.tcBldNo:   .value = medGetP(BldNo, 1, "-") & "-" & medGetP(BldNo, 2, "-") & "-" & Format(medGetP(BldNo, 3, "-"), "00000#")

        .Col = TblColumn1.tcOK: .value = 1
        
        .Col = TblColumn1.tcCOMPONM: .value = GetCompoNm(compcd) 'strComponentNm
        .Col = TblColumn1.tcABBRNM:  .value = medGetP(Get_CompNm(compcd), 1, COL_DIV)
        
        .Col = TblColumn1.tcIRR: .value = IIf(medGetP(objXM.strTmp, 4, vbTab) = "1", "1", "0")
        If .value = "1" Then .Lock = True
        
        InPutNo = InPutNo + 1
        .Col = TblColumn1.tcNo: .value = InPutNo

        .Col = TblColumn1.tcSPCNO:    .value = SpcNum '& "(" & Mid(SpcNum, 1, 2) & Mid(SpcNum, 4) & ")"
        .Col = TblColumn1.tcVFYNM:    .value = medGetP(objXM.strTmp, 3, vbTab)          '검사자
        .Col = TblColumn1.tcVFYDT:    .Text = Format(GetSystemDate, "yyyy-MM-dd")   '검사일
        .Col = TblColumn1.tcCOMPOCD:  .value = compcd                                   '혈액제제코드
        .Col = TblColumn1.tcRESULTFG: .value = "1"
        .Col = TblColumn1.tcDUP:      .value = medGetP(BldNo, 1, "-") & medGetP(BldNo, 2, "-") & Format(medGetP(BldNo, 3, "-"), "00000#") & COL_DIV & compcd                   '중복체크위해
        
        For ii = 1 To .DataColCnt
            .Col = ii
            .ForeColor = vbBlue
        Next
        
        '처방제제와 Assign한 제제가 다른경우에 다른 색깔로.. 변경...
        .Row = .MaxRows
        .Col = TblColumn1.tcCOMPONM
        If .value <> strCompoNm Then
            .ForeColor = DCM_Magenta
            .FontBold = True
        Else
            .ForeColor = vbBlack
            .FontBold = False
        End If
        
        'Irradation 처방의 경우 자동으로 IRR 에 체크해준다.
        Call SetIRR(.Row)
    End With
    
    
    Set objXM = Nothing
End Sub

Private Function GetCompoNm(ByVal vCompoCd As String)
    Dim RS As Recordset
    Dim strSQL As String

    strSQL = " SELECT * FROM s2bbs006"
    strSQL = strSQL & " WHERE " & DBW("compocd=", vCompoCd)

    Set RS = New Recordset

    RS.Open strSQL, DBConn

    If RS.EOF = False Then
        GetCompoNm = RS.Fields("componm").value & ""
    End If

    Set RS = Nothing
End Function
'
'Private Function GetCompoAbbrNm(ByVal vCompoCd As String)
'    Dim Rs As Recordset
'    Dim strSQL As String
'
'    strSQL = " SELECT * FROM s2bbs006"
'    strSQL = strSQL & " WHERE " & DBW("compocd=", vCompoCd)
'
'    Set Rs = New Recordset
'
'    Rs.Open strSQL, DBConn
'
'    If Rs.EOF = False Then
'        GetCompoAbbrNm = Rs.Fields("abbrnm").value & ""
'    End If
'
'    Set Rs = Nothing
'End Function

Private Sub SetIRR(ByVal vRow As Long)
    Dim strSQL As String
    Dim RS As Recordset
    Dim vIrrFg As Variant
    
    strSQL = " select irradfg from " & T_LAB102
    strSQL = strSQL & " where " & DBW("workarea=", C_WORKAREA)
    strSQL = strSQL & " and " & DBW("accdt=", medGetP(txtSpcNO.Text, 1, "-"))
    strSQL = strSQL & " and " & DBW("accseq=", medGetP(txtSpcNO.Text, 2, "-"))
    
    Set RS = New Recordset
    
    RS.Open strSQL, DBConn
    
    If RS.EOF = False Then
        Call tblBlood.GetText(TblColumn1.tcIRR, vRow, vIrrFg)
        If vIrrFg <> "1" Then
            Call tblBlood.SetText(TblColumn1.tcIRR, vRow, IIf(RS.Fields("irradfg").value & "" = "1", 1, 0))
        End If
    End If
    
    Set RS = Nothing
End Sub

Private Sub tblBlood_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If onPgm = True Then Exit Sub
    
    Dim Step(3) As String
    Dim ii      As Integer
    Dim BloodNo As String
    Dim componm As String
    
    Dim sValue As Boolean
    
    If Col = TblColumn1.tcIRR Then Exit Sub
    
    sValue = False
    With tblBlood
        .Row = Row
        .Col = TblColumn1.tcRESULTFG
        If .value <> "1" Then Exit Sub
        Select Case Col
            Case TblColumn1.tcNot
                '응급일때....는 OK, Not가 선택되면 않된다.
                .Col = TblColumn1.tcSTAT
                If .CellType = CellTypeStaticText Then sValue = IIf(.value = "√", True, False)
                If sValue = True Then
                    .Col = TblColumn1.tcBldNo
                    If .ForeColor = vbRed Then
                        .Col = Col: onPgm = True
                        If .value = True Then
                            .Col = TblColumn1.tcSTEP1: .value = "0"
                            .Col = TblColumn1.tcSTEP2: .value = "0"
                            .Col = TblColumn1.tcSTEP3: .value = "0"
                            .Col = TblColumn1.tcSTEP4: .value = "0"
                            .Col = TblColumn1.tcOK: .value = False
                        Else
                            .Col = TblColumn1.tcOK
                            If .value = True Then
                                .Col = Col
                                .value = True
                            End If
                            
                        End If
                        onPgm = False
                    Else
                        .Col = Col
                        onPgm = True
                        If .value = True Then
                            .value = False
                        Else
                            .value = True
                        End If
                        onPgm = False
                    End If
                Else
                    .Col = Col
                    If .value = True Then
                        .Col = TblColumn1.tcSTEP1: .value = "0"
                        .Col = TblColumn1.tcSTEP2: .value = "0"
                        .Col = TblColumn1.tcSTEP3: .value = "0"
                        .Col = TblColumn1.tcSTEP4: .value = "0"
                        .Col = TblColumn1.tcOK
                        If .value = True Then
                            onPgm = True
                            .Col = TblColumn1.tcOK: .value = False
                            onPgm = False
                        End If
                    End If
                End If
            Case TblColumn1.tcSTAT
                .Col = Col
                If .CellType = CellTypeStaticText Then sValue = IIf(.value = "√", True, False)
                If sValue = True Then
                    .Col = TblColumn1.tcBldNo
                    If .ForeColor = vbRed Then
                                            
                    Else
                        onPgm = True
                        .Col = TblColumn1.tcOK: .value = False
                        .Col = TblColumn1.tcNot: .value = False
                        onPgm = False
                    End If
                    
                    .Col = TblColumn1.tcSTEP1:  .value = ""
                    .Col = TblColumn1.tcSTEP2:  .value = ""
                    .Col = TblColumn1.tcSTEP3:  .value = ""
                    .Col = TblColumn1.tcSTEP4:  .value = ""
                Else
                
                End If
            Case TblColumn1.tcOK
                .Col = TblColumn1.tcBldNo
                If .ForeColor = vbRed Then
                    .Col = Col
                    If .value = True Then
                        onPgm = True
                        .Col = TblColumn1.tcNot: .value = False
                        .Col = TblColumn1.tcSTEP1:  .value = "1"
                        .Col = TblColumn1.tcSTEP2:  .value = "1"
                        .Col = TblColumn1.tcSTEP3:  .value = "1"
                        .Col = TblColumn1.tcSTEP4:  .value = "1"
                        onPgm = False
                    End If
                Else
                    .Col = TblColumn1.tcSTAT
                    If .CellType = CellTypeStaticText Then sValue = IIf(.value = "√", True, False)
                    If sValue = True Then
                        onPgm = True
                        .Col = TblColumn1.tcOK: .value = False
                        .Col = TblColumn1.tcNot: .value = False
                        .Col = TblColumn1.tcSTEP1:  .value = ""
                        .Col = TblColumn1.tcSTEP2:  .value = ""
                        .Col = TblColumn1.tcSTEP3:  .value = ""
                        .Col = TblColumn1.tcSTEP4:  .value = ""
                        onPgm = False
                    Else
                        .Col = TblColumn1.tcNot: .value = False
                        .Col = TblColumn1.tcSTEP1:  .value = "1"
                        .Col = TblColumn1.tcSTEP2:  .value = "1"
                        .Col = TblColumn1.tcSTEP3:  .value = "1"
                        .Col = TblColumn1.tcSTEP4:  .value = "1"
                    End If
                End If
        End Select
    End With
End Sub

Private Sub SetOkNot(ByVal Row As Long)
'    Dim i As Long
'    Dim strOkNot As String
'    Dim Col2 As Long
'
'    Select Case Test_Step
'        Case 1: Col2 = TBLCOLUMN.tcSTEP1
'        Case 2: Col2 = TBLCOLUMN.tcSTEP2
'        Case 3: Col2 = TBLCOLUMN.tcSTEP3
'        Case 4: Col2 = TBLCOLUMN.tcSTEP4
'    End Select
'
'    With tblResult
'        .Row = Row
'        For i = TBLCOLUMN.tcSTEP1 To Col2
'            .Col = i
'            If .value = 0 Then
'                .Col = TBLCOLUMN.TcJudge: .value = "Not"
'                           .ForeColor = RGB(255, 0, 0)
'                Exit Sub
'            End If
'        Next i
'        .Col = TBLCOLUMN.TcJudge: .value = "Ok"
'                   .ForeColor = RGB(0, 0, 255)
'    End With
End Sub


Private Function TagPrint(ByVal BloodNo As String, ByVal componm As String, ByVal ABO As String, ByVal Volumn As String, _
                            Optional ByVal Rt As String = "", Optional ByVal DetailRst As String = "")
'-------------
'혈액 Tag 출력
'-------------
    Dim aryContent(1 To 14)
    Dim ii          As Integer
    Dim WardDept    As String
    Dim vfydt       As String
    Dim VFYTM       As String
    Dim Ptid        As String
    Dim ptnm        As String
    Dim strTmp      As String
    Dim iCnt        As Integer
    
    Ptid = lblPtId.Caption
    ptnm = lblPtNm.Caption
    WardDept = strWardID
    
    
    vfydt = Format$(Now, PRESENTDATE_FORMAT)
    VFYTM = Format$(Now, PRESENTTIME_FORMAT)
    vfydt = Mid(vfydt, 3, 2) & "-" & Mid(vfydt, 5, 2) & "-" & Mid(vfydt, 7) & " " & Format(Mid(VFYTM, 1, 4), "0#:##")
    '2001-12-26 수정
    '출력내용추가 : 성별/나이,공혈자혈액형,주민번호,검사자,출고준비일,출고일시간,수령자,응급/비응급여부

    aryContent(1) = Ptid:           aryContent(2) = ptnm:
    
    aryContent(3) = lblWardNm.Caption
    
    If aryContent(3) <> "" Then
        aryContent(3) = aryContent(3) & "-" & lblDeptNm.Caption
        
        If lblDeptNm.Caption = "응급의학과" Then
            aryContent(3) = "EM" & "-" & lblDeptNm.Caption
        End If
    Else
        aryContent(3) = lblDeptNm.Caption
    End If
    
    aryContent(4) = ABO:            aryContent(5) = lblABO.Caption:         aryContent(6) = BloodNo:
    aryContent(7) = Volumn:         aryContent(8) = vfydt:                  aryContent(9) = ObjSysInfo.EmpNm:
    aryContent(10) = DetailRst: 'strSSN:
    aryContent(11) = lblSexAge.Caption
    
    strTmp = "M"
    If Trim(medGetP(aryContent(11), 2, "/")) = "여" Then strTmp = "F"
    strTmp = Trim(medGetP(aryContent(11), 1, "/")) & "/" & strTmp
    aryContent(11) = strTmp
    
    If InStr(1, aryContent(5), "(") > 0 Then
        aryContent(5) = medGetP(aryContent(5), 1, "(") & medGetP(aryContent(5), 2, ")")
    End If
    
    With tblOrder
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1
            If .value = CurrentSelected Then
                .Col = 8
                aryContent(12) = IIf(.value <> "", "1", "")
                If RePrint = True Then
                    tblBlood.Row = tblBlood.ActiveRow
                    tblBlood.Col = TblColumn1.tcVFYDT: aryContent(8) = tblBlood.value
                    tblBlood.Col = TblColumn1.tcVfyTm: aryContent(8) = aryContent(8) & " " & tblBlood.value
                    tblBlood.Col = TblColumn1.tcVFYNM: aryContent(9) = tblBlood.value
                End If
                Exit For
            End If
        Next
    End With
    aryContent(13) = componm
    aryContent(14) = IIf(Rt = "1", "1", "")
    
    '** 추가 사용자 출력장수 설정 By M.G.Choi 2007.12.04
    iCnt = Trim(txtLabelCnt.Text)
    
    For iCnt = 1 To iCnt
        BloodLabel_Print aryContent()
    Next
    '-----------------------------------------------------
    
End Function
Private Sub cmdTagPrint_Click()
    Dim componm  As String
    Dim Volumn   As String
    Dim ABO      As String
    Dim BloodNum As String
    Dim Rt       As String
    Dim strSTEP1 As String
    Dim strSTEP2 As String
    Dim strSTEP3 As String
    Dim strSTEP4 As String
    Dim strDetail   As String
    
    With tblBlood
        If .DataRowCnt < 1 Then Exit Sub
        RePrint = True
        .Row = .ActiveRow
        .Col = TblColumn1.tcSTATUS
        If .value = "A" Or .value = "출고" Then
            .Col = TblColumn1.tcBldNo: BloodNum = .value
                                       BloodNum = Mid(BloodNum, 1, 6) & Format(Mid(BloodNum, 7), "000000")
            .Col = TblColumn1.tcABBRNM: componm = .value
            .Col = TblColumn1.tcABO: ABO = .value
            .Col = TblColumn1.tcVol: Volumn = .value
            .Col = TblColumn1.tcIRR: Rt = IIf(.Text <> "", "1", "")
            
            '-- 주민번호 --> 상세결과 추가 By M.G.Choi 2007.07.02
            .Col = TblColumn1.tcSTEP1: strSTEP1 = IIf(.value = "1", "S(O)", "S(X)")
            .Col = TblColumn1.tcSTEP2: strSTEP2 = IIf(.value = "1", "B(O)", "B(X)")
            .Col = TblColumn1.tcSTEP3: strSTEP3 = IIf(.value = "1", "37(O)", "37(X)")
            .Col = TblColumn1.tcSTEP4: strSTEP4 = IIf(.value = "1", "C(O)", "C(X)")
            strDetail = strSTEP1 & strSTEP2 & strSTEP3 & strSTEP4
            
            Call TagPrint(BloodNum, componm, ABO, Volumn, Rt, strDetail)
        Else
            MsgBox "Tag 재출력 대상이 아닙니다.", vbInformation + vbOKOnly, "Tag 재출력"
        End If
    End With
End Sub

Public Sub ClickQueryButton()
    Call txtSpcNoLostFocus
End Sub

Private Sub P_PrtSet()
    Printer.Font = "굴림체"
    Printer.FontSize = 10
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
End Sub

Private Sub UpDown1_DownClick()
    
    If CInt(Trim(txtLabelCnt.Text)) <= 1 Then
        txtLabelCnt.Text = "1"
        Exit Sub
    End If
    
    If Trim(txtLabelCnt.Text) = "" Then
        Exit Sub
    End If
    
    If CInt(Trim(txtLabelCnt.Text)) < 1 Then
        txtLabelCnt.Text = 1
    Else
        txtLabelCnt.Text = CInt(txtLabelCnt.Text) - 1
    End If
    
End Sub

Private Sub UpDown1_UpClick()
    
    If CInt(Trim(txtLabelCnt.Text)) >= 9 Then
        txtLabelCnt.Text = "9"
        Exit Sub
    End If
    
    If Trim(txtLabelCnt.Text) = "" Then
        Exit Sub
    End If
    
    If CInt(Trim(txtLabelCnt.Text)) >= 9 Then
        txtLabelCnt.Text = "9"
    Else
        txtLabelCnt.Text = CInt(txtLabelCnt.Text) + 1
    End If
    
End Sub
