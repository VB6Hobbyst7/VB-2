VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS201_B 
   BackColor       =   &H00DBE6E6&
   Caption         =   "CrossMatching Result Register"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   ClipControls    =   0   'False
   Icon            =   "frmBBS201_B.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.TextBox txtStart 
      Height          =   300
      Left            =   9555
      TabIndex        =   45
      Top             =   2925
      Width           =   450
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   44
      Top             =   8535
      Width           =   1320
   End
   Begin VB.TextBox txtBldNo 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7350
      MaxLength       =   12
      TabIndex        =   43
      Top             =   2940
      Width           =   2205
   End
   Begin VB.CommandButton cmdTagPrint 
      BackColor       =   &H00F4F0F2&
      Enabled         =   0   'False
      Height          =   345
      Left            =   7365
      Picture         =   "frmBBS201_B.frx":000C
      Style           =   1  '그래픽
      TabIndex        =   42
      ToolTipText     =   "혈액Tag 재출력"
      Top             =   90
      Width           =   345
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   315
      Left            =   2895
      TabIndex        =   39
      Top             =   45
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   62455811
      CurrentDate     =   37063
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   480
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   38
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   37
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   4260
      _ExtentX        =   7514
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
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   4350
      TabIndex        =   34
      Top             =   45
      Width           =   10290
      _ExtentX        =   18150
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
      Caption         =   "CrossMatching 처방 리스트"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblOdList 
      Height          =   2460
      Left            =   4335
      TabIndex        =   35
      Top             =   420
      Width           =   10305
      _Version        =   196608
      _ExtentX        =   18177
      _ExtentY        =   4339
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15265518
      GridColor       =   16703181
      GridShowVert    =   0   'False
      MaxCols         =   20
      MaxRows         =   10
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS201_B.frx":053E
      TextTip         =   2
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   4350
      TabIndex        =   36
      Top             =   2925
      Width           =   10290
      _ExtentX        =   18150
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
      Caption         =   "CrossMatching 등록 리스트"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblBldList 
      Height          =   5085
      Left            =   4350
      TabIndex        =   41
      Top             =   3285
      Width           =   10275
      _Version        =   196608
      _ExtentX        =   18124
      _ExtentY        =   8969
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15265518
      GridColor       =   16703181
      GridShowVert    =   0   'False
      MaxCols         =   20
      MaxRows         =   18
      OperationMode   =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS201_B.frx":0E94
      TextTip         =   2
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   8040
      Left            =   75
      TabIndex        =   2
      Top             =   330
      Width           =   4275
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3795
         Width           =   960
         _ExtentX        =   1693
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   4140
         Width           =   960
         _ExtentX        =   1693
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
         Caption         =   "검체위치"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   4485
         Width           =   960
         _ExtentX        =   1693
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
         Caption         =   "경과시간"
         Appearance      =   0
      End
      Begin VB.Frame fraABO 
         BorderStyle     =   0  '없음
         Height          =   2355
         Left            =   1620
         TabIndex        =   22
         Top             =   1230
         Width           =   2415
         Begin VB.TextBox txtCABO 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   25
            Top             =   855
            Width           =   1110
         End
         Begin VB.TextBox txtSABO 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1215
            Width           =   1110
         End
         Begin VB.TextBox txtRH 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   23
            Top             =   1575
            Width           =   1110
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   26
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
            Caption         =   "혈액형 입력"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblaboptnm 
            Height          =   300
            Left            =   1230
            TabIndex        =   27
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
            TabIndex        =   33
            Tag             =   "103"
            Top             =   900
            Width           =   690
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
            TabIndex        =   32
            Tag             =   "103"
            Top             =   540
            Width           =   585
         End
         Begin VB.Label lblaboapply 
            AutoSize        =   -1  'True
            Caption         =   "적용"
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
            Left            =   630
            TabIndex        =   31
            Top             =   2085
            Width           =   405
         End
         Begin VB.Label lblabocancel 
            AutoSize        =   -1  'True
            Caption         =   "취소"
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
            Left            =   1590
            TabIndex        =   30
            Top             =   2085
            Width           =   405
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
            TabIndex        =   29
            Tag             =   "103"
            Top             =   1260
            Width           =   945
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
            TabIndex        =   28
            Tag             =   "103"
            Top             =   1590
            Width           =   225
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2115
         Width           =   960
         _ExtentX        =   1693
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
         Caption         =   "상병"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1620
         Width           =   960
         _ExtentX        =   1693
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
         Index           =   6
         Left            =   60
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1275
         Width           =   960
         _ExtentX        =   1693
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
         Index           =   4
         Left            =   60
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   930
         Width           =   960
         _ExtentX        =   1693
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
         Caption         =   "성별나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   585
         Width           =   960
         _ExtentX        =   1693
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
         Caption         =   "성명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
         _ExtentX        =   1693
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
      Begin VB.CommandButton cmdPrep 
         BackColor       =   &H00F4F0F2&
         Caption         =   "환자별PrepOrder"
         Height          =   705
         Left            =   2670
         Style           =   1  '그래픽
         TabIndex        =   47
         ToolTipText     =   "최대로"
         Top             =   3750
         Width           =   1530
      End
      Begin VB.CommandButton cmdRmk 
         BackColor       =   &H00F4F0F2&
         Caption         =   "환자별특이사항등록"
         Height          =   780
         Left            =   2310
         Style           =   1  '그래픽
         TabIndex        =   20
         ToolTipText     =   "최대로"
         Top             =   1650
         Width           =   1845
      End
      Begin VB.CheckBox chkABO 
         BackColor       =   &H00DBE6E6&
         Caption         =   "혈액형등록"
         Height          =   315
         Left            =   2310
         TabIndex        =   19
         Top             =   990
         Value           =   1  '확인
         Width           =   1455
      End
      Begin DRcontrol1.DrLabel lblTime 
         Height          =   315
         Left            =   1035
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4485
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblSpcPos 
         Height          =   315
         Left            =   1035
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4140
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblSpcNo 
         Height          =   315
         Left            =   1035
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3795
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
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
         ItemData        =   "frmBBS201_B.frx":2666
         Left            =   2130
         List            =   "frmBBS201_B.frx":2676
         Locked          =   -1  'True
         Style           =   1  '단순 콤보
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2985
         Width           =   2085
      End
      Begin VB.TextBox txtDisNm 
         BorderStyle     =   0  '없음
         Height          =   465
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmBBS201_B.frx":26A0
         Top             =   2475
         Width           =   4035
      End
      Begin DRcontrol1.DrLabel lblDisCd 
         Height          =   315
         Left            =   1035
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2115
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblDeptNm 
         Height          =   315
         Left            =   1035
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblWard 
         Height          =   315
         Left            =   1035
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblSexAge 
         Height          =   315
         Left            =   1035
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   930
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblPtNm 
         Height          =   315
         Left            =   1035
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   585
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "이상대"
      End
      Begin VB.TextBox txtPtid 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1110
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2985
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
         BackColor       =   9083801
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
         Caption         =   " 검사방법"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   330
         Index           =   1
         Left            =   60
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3390
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   582
         BackColor       =   9083801
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
         Caption         =   "검체정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   345
         Index           =   2
         Left            =   105
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4950
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   609
         BackColor       =   9083801
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
         Caption         =   "관련 검사 정보"
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblTest 
         Height          =   2565
         Left            =   105
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5385
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
         _ExtentY        =   4524
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
         MaxRows         =   10
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS201_B.frx":26A7
      End
      Begin MedControls1.LisLabel lblRmkFg 
         Height          =   285
         Left            =   3270
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1305
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
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
         Caption         =   "Y"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblrmk 
         Height          =   300
         Left            =   30
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   -60
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   2295
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1275
         Width           =   960
         _ExtentX        =   1693
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
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2265
         TabIndex        =   9
         Top             =   405
         Width           =   1830
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   720
         Left            =   2265
         TabIndex        =   10
         Top             =   210
         Width           =   1920
      End
   End
   Begin VB.Frame fraPrep 
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   5010
      Left            =   105
      TabIndex        =   46
      Top             =   3315
      Visible         =   0   'False
      Width           =   4260
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Refre(&R)"
         Height          =   480
         Left            =   2460
         Style           =   1  '그래픽
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   495
         Width           =   855
      End
      Begin VB.CommandButton CmdClose 
         BackColor       =   &H00F4F0F2&
         Caption         =   "닫기(&T)"
         Height          =   480
         Left            =   3330
         Style           =   1  '그래픽
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   495
         Width           =   855
      End
      Begin FPSpread.vaSpread tblPrep 
         Height          =   3780
         Left            =   60
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1110
         Width           =   4170
         _Version        =   196608
         _ExtentX        =   7355
         _ExtentY        =   6668
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
         MaxCols         =   6
         MaxRows         =   10
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS201_B.frx":2B8B
      End
      Begin MSComCtl2.DTPicker dtpPre 
         Height          =   315
         Left            =   1035
         TabIndex        =   49
         Top             =   585
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62455811
         CurrentDate     =   37063
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "조회기준일"
         Height          =   180
         Index           =   9
         Left            =   75
         TabIndex        =   52
         Tag             =   "40304"
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "3일 이전 처방까지 조회됩니다.............."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   195
         TabIndex        =   50
         Top             =   180
         Width           =   3735
      End
      Begin VB.Label Label4 
         BackColor       =   &H000040C0&
         Height          =   345
         Left            =   30
         TabIndex        =   51
         Top             =   90
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmBBS201_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sORDDT As String
Private onPgm  As Boolean
'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_MFY& = 1

Public Sub ClickQueryButton()
    Call QueryPt
End Sub

Private Sub Form_Clear()
    'txtPtid.Text = ""
    lblPtNm.Caption = ""
    lblABO.Caption = ""
    lblSexAge.Caption = ""
    lblWard.Caption = ""
    lblDeptNm.Caption = ""
    lblSpcNo.Caption = ""
    lblSpcPos.Caption = ""
    lblTime.Caption = ""
    lblDisCd.Caption = ""
    txtDisNm.Text = ""
    fraABO.Visible = False
    lblaboptnm.Caption = ""
    txtSABO.Text = ""
    txtCABO.Text = ""
    txtRH.Text = ""
    lblRmkFg.Caption = ""
    chkABO.value = 0
    
    txtBldNo.Text = ""
    onPgm = False
    Call medClearTable(tblOdList)
    Call medClearTable(tblBldList)
    
End Sub


Private Sub cmdRefresh_Click()
    Call QueryPrep
End Sub
Private Sub cmdClose_Click()
    fraPrep.Visible = False
    txtBldNo.SetFocus
    'Call QueryPrep

End Sub

Private Sub cmdPrep_Click()
    '준비오더보여주는 버튼
    fraPrep.Visible = True
    cmdRefresh.SetFocus
    DoEvents
    medClearTable tblPrep
    dtpPre.value = GetSystemDate
    Call cmdRefresh_Click
'    QueryPrep
End Sub
Private Sub QueryPrep()
    '준비오더 보여주기
    Dim RS   As Recordset
    Dim SSQL As String
    Dim sFrDt As String
    Dim sToDt As String
    Dim sPtid As String
    
    
    
    sPtid = txtPtid.Text
    
    
    If sPtid = "" Then Exit Sub
   
    
    sFrDt = Format(DateAdd("d", -2, dtpPre.value), CS_DateDbFormat)
    sToDt = Format(dtpPre.value, CS_DateDbFormat)
    
    SSQL = " select a.orddt,a.testcd,a.statfg,a.dcfg,a.serial,b.abbrnm10 " & _
           " from s2bbs_prep a," & T_BBS001 & " b" & _
           " where " & _
                      DBW("ptid=", sPtid) & _
           " and " & DBW("orddt>=", sFrDt) & _
           " and " & DBW("orddt<=", sToDt) & _
           " and a.testcd=b.testcd"
          
          
    medClearTable tblPrep
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblPrep
            .ReDraw = False
            Do Until RS.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                '처방일
                .Col = 1: .value = Format(RS.Fields("orddt").value & "", "####-##-##")
                '처방코드
                .Col = 2: .value = RS.Fields("abbrnm10").value & ""
                '응급
                .Col = 3: .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", "")
                          If .value = "Y" Then .ForeColor = DCM_LightRed: .FontBold = True
                '수량
                .Col = 4: .value = "1"
                'DC
                .Col = 5: .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", "")
                          If .value = "Y" Then .ForeColor = DCM_LightRed: .FontBold = True
                'serial
                .Col = 6: .value = RS.Fields("serial").value & ""
                
                RS.MoveNext
            Loop
            If .MaxRows < 12 Then .MaxRows = 12
            
            .ReDraw = True
        End With
    End If
    Set RS = Nothing
End Sub




Private Sub cmdTagPrint_Click()
    Dim componm  As String
    Dim Volumn   As String
    Dim ABO      As String
    Dim BloodNum As String
    Dim Rt       As String
    
    Dim ordno    As String
    Dim orddt   As String
    
    With tblOdList
        If .DataRowCnt < 1 Then Exit Sub
        .Row = .ActiveRow
        .Col = 9
        If .value = "ASSIGN" Or .value = "출고" Then
            .Col = 1: orddt = Trim(Replace(.value, "-", ""))
            .Col = 7: BloodNum = .value
            .Col = 3: componm = .value
            .Col = 4: ABO = .value
            .Col = 5: Volumn = .value & "cc"
            .Col = 15: ordno = Trim(.value)
'            .Col = 16: ORDSEQ = Trim(.value)

    
            Call TagPrint(BloodNum, componm, ABO, Volumn, orddt, ordno)
        Else
            MsgBox "Tag 재출력 대상이 아닙니다.", vbInformation + vbOKOnly, "Tag 재출력"
        End If
    End With

End Sub

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
    
    SSQL = objSql.DeleteABO(txtPtid.Text)
    DBConn.Execute SSQL
    
    SSQL = objSql.InsertABO(txtPtid.Text, txtCABO.Text, txtSABO.Text, txtRH.Text)
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

'Private Sub mnuDelete_Click()
''혈액 삭제
'    With tblBldList
'        .Row = .ActiveRow
'        .Col = 4: .value = ""
'        .Col = 5: .value = ""
'        .Col = 8: .value = 0
'        .Col = 7: .value = ""
'        .Col = 11: .value = ""
'        .Col = 12: .value = ""
'        .Col = 16: .value = ""
'        txtBldNo.SelStart = 0
'        txtBldNo.SelLength = Len(txtBldNo.Text)
'
'        '.Action = ActionDeleteRow
''        .MaxRows = .MaxRows - 1
''        InPutNo = InPutNo - 1
'    End With
'End Sub

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
    End If
End Sub

Private Sub cmdClear_Click()
    Call Form_Clear
    txtPtid.Text = "": txtPtid.SetFocus
    Call ICSPatientMark
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRmk_Click()
    If txtPtid.Text = "" Then Exit Sub
    frmXMRemark.sPtid = txtPtid.Text
    frmXMRemark.rmk = lblrmk.Caption
    frmXMRemark.Show 1

End Sub

Private Sub dtpDt_CloseUp()
    txtPtid.Text = Format(txtPtid.Text, "000000000")
    Call QueryPt
End Sub

Private Sub Form_Load()
    dtpDt.value = GetSystemDate
    Call Form_Clear
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_MFY
        '혈액 삭제
            With tblBldList
                .Row = .ActiveRow
                .Col = 4: .value = ""
                .Col = 5: .value = ""
                .Col = 8: .value = 0
                .Col = 7: .value = ""
                .Col = 11: .value = ""
                .Col = 12: .value = ""
                .Col = 16: .value = ""
                txtBldNo.SelStart = 0
                txtBldNo.SelLength = Len(txtBldNo.Text)
        
                '.Action = ActionDeleteRow
        '        .MaxRows = .MaxRows - 1
        '        InPutNo = InPutNo - 1
            End With
    End Select
End Sub

Private Sub tblBldList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If onPgm = True Then Exit Sub
    
    Dim Step(3) As String
    Dim ii      As Integer
    Dim BloodNo As String
    Dim componm As String
    
    Dim sValue As Boolean
    
    If Row = 0 Then Exit Sub
    'If Col <> 8 Or Col <> 9 Or Col <> 10 Then Exit Sub
    If Col < 8 Or Col > 10 Then Exit Sub
    sValue = False
    With tblBldList
        .Row = Row
        .Col = 7
        If .value = "" Then Exit Sub
        .Col = Col
        
        Select Case Col
            'Not
            Case 9
                .Col = Col: onPgm = True
                If .value = True Then
                    .Col = 8: .value = False
                End If
                onPgm = False

            Case 8
                .Col = Col
                If .value = True Then
                    onPgm = True
                    .Col = 9: .value = False
                    onPgm = False
                End If
    
        End Select
    End With
End Sub

Private Sub tblBldList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'마우스 오른쪽 버튼 클릭시 해당 라인의 Delete 기능 수행.
    If Row < 1 Then Exit Sub
    
    Dim strTmp As String
    
    With tblBldList
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
        .Col = 7
        If .value <> "" Then
            Set objPop = New clsPopupMenu
            With objPop
                .AddMenu MENU_MFY, "수정"
                .PopupMenus Me.hwnd
            End With
            Set objPop = Nothing
'            Set mnuPopup = frmControls.mnuPopup
'            Set mnuDelete = frmControls.mnuSub
'            mnuDelete.Caption = "수정"
'
'            PopupMenu mnuPopup
'
'            Set mnuPopup = Nothing
'            Set mnuDelete = Nothing
        End If
    End With
End Sub

Private Sub txtPtId_GotFocus()
    txtPtid.tag = txtPtid
    
    txtPtid.SelStart = 0
    txtPtid.SelLength = Len(txtPtid)
    Exit Sub
End Sub
Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txtPtId_LostFocus()
    If txtPtid = "" Then Exit Sub
    If txtPtid.tag = txtPtid Then Exit Sub
    
    Dim ii      As Integer
    Dim tmpid   As String
    
    For ii = 1 To BBS_PTID_LENGTH
        tmpid = tmpid & "0"
    Next
    
    
    txtPtid.Text = Format(txtPtid.Text, tmpid)
    Call ICSPatientMark(txtPtid.Text, enICSNum.BBS_ALL)
    
    
    Call QueryPt
End Sub
Private Sub QueryPt()
    Dim strTmp As String
    Call Form_Clear
    '혈액형구하기
    DoEvents
    Call GetABO(txtPtid.Text)
    
    strTmp = GetOrdDt(txtPtid)
    If strTmp <> "" Then
        '상병구하기
        Call GetDisease(txtPtid.Text, medGetP(strTmp, 1, COL_DIV), medGetP(strTmp, 2, COL_DIV))
        '환자정보
        Call GetPt(txtPtid.Text, medGetP(strTmp, 1, COL_DIV), medGetP(strTmp, 2, COL_DIV))
    Else
        txtPtid.Text = "": txtPtid.SetFocus
        Exit Sub
    End If
    '환자별 리마크
    DoEvents
    Call Find_PtRemark(txtPtid)
    '관련검사항목 구하기
    DoEvents
    Call GetTestInformation(txtPtid)
    '처방항목 구하기
    DoEvents
    Call QueryBlood
    txtStart.Text = 1
End Sub
Private Function GetOrdDt(ByVal PtId As String) As String
    Dim SSQL  As String
    Dim sFrDt As String
    Dim sToDt As String
    Dim RS    As Recordset
    
    sToDt = Format(dtpDt.value, CS_DateDbFormat)
    sFrDt = Format(DateAdd("d", -2, dtpDt.value), CS_DateDbFormat)
    
    SSQL = " SELECT MAX(ORDDT) AS ORDDT FROM  " & T_LAB101 & " A" & _
           " WHERE " & DBW("A.PTID=", txtPtid.Text) & _
           " AND   " & DBW("A.ORDDT>=", sFrDt) & _
           " AND   " & DBW("A.ORDDT<=", sToDt) & _
           " AND   " & DBW("ORDDIV=", C_WORKAREA) & _
           " AND EXISTS(SELECT * FROM " & T_LAB102 & " B" & _
           "            WHERE " & _
           "                  " & DBW("A.PTID=", txtPtid.Text) & _
           "            AND   " & DBW("A.ORDDIV=", C_WORKAREA) & _
           "            AND   " & DBW("A.ORDDT>=", sFrDt) & _
           "            AND   " & DBW("A.ORDDT<=", sToDt) & _
           "            AND   " & DBW("B.ocsordno>", "0") & _
           "            AND   (B.DCFG='' OR B.DCFG IS NULL)" & _
           "            AND   A.PTID =B.PTID AND A.ORDDT=B.ORDDT AND A.ORDNO=B.ORDNO)" '& _
           " ORDER BY ORDDT"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        GetOrdDt = RS.Fields("ORDDT").value & ""
        
        SSQL = " SELECT MAX(ORDNO)  AS ORDNO FROM " & T_LAB101 & " A" & _
               " WHERE " & DBW("A.PTID=", txtPtid.Text) & _
               " AND   " & DBW("A.ORDDT=", GetOrdDt) & _
               " AND   " & DBW("ORDDIV=", C_WORKAREA) & _
               " AND EXISTS(SELECT * FROM " & T_LAB102 & " B" & _
               "            WHERE " & _
               "                  " & DBW("A.PTID=", txtPtid.Text) & _
               "            AND   " & DBW("A.ORDDIV=", C_WORKAREA) & _
               "            AND   " & DBW("A.ORDDT>=", sFrDt) & _
               "            AND   " & DBW("A.ORDDT<=", sToDt) & _
               "            AND   " & DBW("B.ocsordno>", "0") & _
               "            AND   (B.DCFG='' OR B.DCFG IS NULL)" & _
               "            AND   A.PTID =B.PTID AND A.ORDDT=B.ORDDT AND A.ORDNO=B.ORDNO)" '& _
               " ORDER BY ORDDT"
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        If Not RS.EOF Then
            GetOrdDt = GetOrdDt & COL_DIV & RS.Fields("ORDNO").value & ""
        End If
        
    End If
    
    Set RS = Nothing

End Function
Private Sub Find_PtRemark(ByVal PtId As String)
    Dim objSql As New clsCrossMatching
    
    lblrmk.Caption = objSql.GetptidRmk(PtId)
    
    If lblrmk.Caption <> "" Then
        lblRmkFg.Caption = "Y"
    Else
        lblRmkFg.Caption = ""
    End If
    Set objSql = Nothing
End Sub

Private Sub GetABO(ByVal PtId As String)
'혈액형,부작용,감염정보,상병코드,상병을 조회한다.
    Dim ObjABO As New clsABO
    
    With ObjABO
        .PtId = PtId
        If .GetABO = True Then
            lblABO.Caption = .ABO & .Rh
        Else
            lblABO.Caption = ""
        End If
    End With
    Set ObjABO = Nothing
    
End Sub

Private Sub GetDisease(ByVal PtId As String, ByVal orddt As String, ByVal ordno As String)
    '상병정보를 가지고 온다.
    Dim objDisease As New clsDisease
    
    With objDisease
        .PtId = PtId
        .orddt = orddt
        .ordno = CStr(ordno)
        If .GetDisease = True Then
            lblDisCd.Caption = .DiseaseCd      '상병코드
            txtDisNm.Text = .DiseaseNm         '상병명
        Else
            lblDisCd.Caption = ""
            txtDisNm.Text = ""
        End If
    End With
    
    Set objDisease = Nothing
End Sub
Private Sub GetPt(ByVal PtId As String, ByVal orddt As String, ByVal ordno As Long)
'환자와 검체정보를 조회한다.
    Dim objXM    As New clsCrossMatching
    Dim objSql   As New clsGetSqlStatement
    Dim DrRS     As Recordset
    Dim strTmp   As String
    Dim Timechk  As Long
    Dim ii       As Integer: ii = 0
    
    With objXM
        strTmp = .Get_PtInfo(PtId, orddt, ordno)
        If strTmp <> "" Then
            lblPtNm.Caption = medGetP(strTmp, 1, COL_DIV)
            lblSexAge.Caption = medGetP(strTmp, 2, COL_DIV)
            lblDeptNm.Caption = medGetP(strTmp, 3, COL_DIV)
            lblWard.Caption = medGetP(strTmp, 4, COL_DIV)
        End If
        
    End With
    
    Set DrRS = objXM.Get_SpcInfo(PtId, orddt)
            
    Timechk = objSql.Spc_TimeChk(PtId)
    lblTime.ForeColor = vbRed
    lblTime.Caption = "검체채취 경과시간: " & Timechk & " 시간"
            
    If Not DrRS.EOF Then
        lblSpcNo.Caption = DrRS.Fields("spcyy").value & "" & "-" & DrRS.Fields("spcno").value & ""
                  
        lblSpcPos.Caption = UCase(DrRS.Fields("storeleg").value & "") & _
                           "(" & DrRS.Fields("storerno").value & "" & _
                           "," & DrRS.Fields("storecno").value & "" & ")"
    Else
        lblTime.Caption = "검체가 존재하지 않습니다."
    End If
    
    Set objSql = Nothing
    Set DrRS = Nothing
    Set objXM = Nothing
End Sub

Private Sub GetTestInformation(ByVal strPtid As String)
    Dim objSql As New clsCrossMatching
    Dim RS     As Recordset
    Dim SSQL   As String
    Dim ii     As Integer
    
    Call medClearTable(tblTest)
    
    SSQL = objSql.TestResultXM(strPtid)
    If SSQL <> "" Then
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        If Not RS.EOF Then
            With tblTest
                If RS.RecordCount < 8 Then
                    .MaxRows = 8
                Else
                    .MaxRows = RS.RecordCount
                End If
                Do Until RS.EOF
                    ii = ii + 1
                    .Row = ii
                    .Col = 1: .value = RS.Fields("workarea").value & "" & "-" & RS.Fields("accdt").value & "" & "-" & RS.Fields("accseq").value & ""
                    .Col = 2: .value = RS.Fields("abbrnm10").value & ""
                    .Col = 3: .value = RS.Fields("rstcd").value & ""
                    .Col = 4: .value = RS.Fields("rstunit").value & ""
                    RS.MoveNext
               Loop
            End With
        End If
        Set RS = Nothing
    End If
    Set objSql = Nothing
End Sub

Private Sub QueryBlood()
    Dim SSQL  As String
    Dim sFrDt As String
    Dim sToDt As String
    Dim RS    As Recordset
    
    Dim ii As Integer
    Dim sAccDt As String
    Dim sAccSeq As String
    Dim strTmp  As String
    Dim blnFG   As Boolean
    
    sToDt = Format(dtpDt.value, CS_DateDbFormat)
    sFrDt = Format(DateAdd("d", -2, dtpDt.value), CS_DateDbFormat)

    SSQL = " SELECT A.WORKAREA,A.ACCDT,A.ACCSEQ,A.ORDCD,A.ORDNO,A.ORDDT,A.ORDSEQ,B.ABBRNM10 AS TESTNM,B.COMPOCD,C.ABBRNM AS COMPONM "
    SSQL = SSQL & _
          " FROM " & T_BBS006 & " C," & T_BBS001 & " B," & T_LAB102 & " A," & T_LAB101 & " Q"
    SSQL = SSQL & " WHERE " & _
                           DBW("Q.PTID=", txtPtid.Text) & _
                  " AND " & DBW("Q.ORDDT>=", sFrDt) & _
                  " AND " & DBW("Q.ORDDT<=", sToDt) & _
                  " AND " & DBW("Q.ORDDIV=", C_WORKAREA) & _
                  " AND " & DBW("A.STSCD>=", BBSOrderStatus.stsACCESS) & _
                  " AND " & DBW("A.ocsordno>", "0") & _
                  " AND (A.DCFG='' OR A.DCFG IS NULL)" & _
                  " AND Q.PTID=A.PTID AND Q.ORDDT=A.ORDDT AND Q.ORDNO=A.ORDNO" & _
                  " AND A.ORDCD=B.TESTCD" & _
                  " AND B.COMPOCD=C.COMPOCD" & _
                  " ORDER BY ORDDT DESC"
                  
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    With tblOdList
        If Not RS.EOF Then
            .ReDraw = False
            If RS.RecordCount < 7 Then
                .MaxRows = 7
            Else
                .MaxRows = RS.RecordCount
            End If
            ii = 1
            Do Until RS.EOF
                .Row = ii
                sAccDt = RS.Fields("ACCDT").value & ""
                sAccSeq = RS.Fields("ACCSEQ").value & ""
                If blnFG = False Then
                    sORDDT = RS.Fields("orddt").value & ""
                    blnFG = True
                End If
                .Col = 1: .value = Format(RS.Fields("ORDDT").value & "", "0#-##-##")
                .Col = 2: .value = RS.Fields("TESTNM").value & ""
                .Col = 3: .value = RS.Fields("COMPONM").value & ""
                .Col = 6: .value = sAccDt & "-" & sAccSeq: .ForeColor = DCM_LightRed
                .Col = 15: .value = RS.Fields("ORDNO").value & ""
                .Col = 16: .value = RS.Fields("ORDSEQ").value & ""
                .Col = 17: .value = RS.Fields("ORDCD").value & ""
                .Col = 18: .value = RS.Fields("COMPOCD").value & ""
                '결과,혈액번호,보고일시,보고자,RT,ABORH,VOLUMN,상태,출고일시,출고자
                strTmp = DetailBlood(sAccDt, sAccSeq)
                If strTmp <> "" Then
                    .Col = 4: .value = medGetP(strTmp, 6, COL_DIV)
                    .Col = 5: .value = medGetP(strTmp, 7, COL_DIV)
                    .Col = 7: .value = medGetP(strTmp, 2, COL_DIV)
                    .Col = 8: .value = medGetP(strTmp, 1, COL_DIV): .ForeColor = DCM_LightBlue
                    .Col = 9: .value = medGetP(strTmp, 8, COL_DIV)
                    
                    .Col = 10: .value = medGetP(strTmp, 4, COL_DIV)
                    .Col = 11: .value = medGetP(strTmp, 3, COL_DIV)
                    .Col = 12: .value = medGetP(strTmp, 10, COL_DIV)
                    .Col = 13: .value = medGetP(strTmp, 9, COL_DIV)
                    .Col = 14: .value = IIf(medGetP(strTmp, 5, COL_DIV) = "1", "Y", ""): .ForeColor = DCM_LightRed
                End If
                
                ii = ii + 1
                RS.MoveNext
            Loop
            
            .ReDraw = True
            Call BloodRegister(sORDDT)
        End If
    End With
    txtBldNo.SetFocus
End Sub
Private Sub tblOdList_Click(ByVal Col As Long, ByVal Row As Long)
    With tblOdList
        .Row = Row
        .Col = 1
        Call BloodRegister(Replace(.value, "-", ""))
        .Col = 9
        If .value = "ASSIGN" Or .value = "출고" Then
            cmdTagPrint.Enabled = True
        Else
            cmdTagPrint.Enabled = False
        End If
    End With
    
End Sub

Public Sub BloodRegister(ByVal orddt As String)
    Dim SSQL As String
    Dim RS   As Recordset
    Dim ii   As Integer
    
    SSQL = " SELECT A.WORKAREA,A.ACCDT,A.ACCSEQ,A.ORDCD,A.ORDNO,A.ORDDT,A.ORDSEQ,B.ABBRNM10 AS TESTNM,B.COMPOCD,C.ABBRNM AS COMPONM "
    SSQL = SSQL & _
          " FROM " & T_BBS006 & " C," & T_BBS001 & " B," & T_LAB102 & " A," & T_LAB101 & " Q"
    SSQL = SSQL & " WHERE " & _
                           DBW("Q.PTID=", txtPtid.Text) & _
                  " AND " & DBW("Q.ORDDT=", orddt) & _
                  " AND " & DBW("Q.ORDDIV=", C_WORKAREA) & _
                  " AND " & DBW("A.STSCD>=", BBSOrderStatus.stsACCESS) & _
                  " AND " & DBW("A.ocsordno>", "0") & _
                  " AND (A.DCFG='' OR A.DCFG IS NULL)" & _
                  " AND NOT EXISTS(SELECT * FROM " & T_BBS302 & " D" & " WHERE " & _
                                            DBW("Q.PTID=", txtPtid.Text) & _
                  "                 AND " & DBW("Q.ORDDT=", orddt) & _
                  "                 AND " & DBW("Q.ORDDIV=", C_WORKAREA) & _
                  "                 AND " & DBW("A.STSCD>=", BBSOrderStatus.stsACCESS) & _
                  "                 AND " & DBW("A.ocsordno>", "0") & _
                  "                 AND (A.DCFG='' OR A.DCFG IS NULL)" & _
                  "                 AND (D.CANCELFG IS NULL OR D.CANCELFG ='')" & _
                  "                 AND Q.PTID=A.PTID AND Q.ORDDT=A.ORDDT AND Q.ORDNO =A.ORDNO" & _
                  "                 AND D.WORKAREA=A.WORKAREA AND D.ACCDT=A.ACCDT AND D.ACCSEQ=A.ACCSEQ)" & _
                  " AND Q.PTID=A.PTID AND Q.ORDDT=A.ORDDT AND Q.ORDNO=A.ORDNO" & _
                  " AND A.ORDCD=B.TESTCD" & _
                  " AND B.COMPOCD=C.COMPOCD" & _
                  " ORDER BY ORDDT DESC"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    With tblBldList
        If Not RS.EOF Then
            If RS.RecordCount < 18 Then
                .MaxRows = 18
            Else
                .MaxRows = RS.RecordCount
            End If
            Do Until RS.EOF
                ii = ii + 1
                .Row = ii
                .Col = 1: .value = Format(RS.Fields("ORDDT").value & "", "0#-##-##")
                .Col = 2: .value = RS.Fields("TESTNM").value & ""
                .Col = 3: .value = RS.Fields("COMPONM").value & ""
                .Col = 6: .value = RS.Fields("ACCDT").value & "" & "-" & RS.Fields("ACCSEQ").value & "": .ForeColor = DCM_LightRed
                .Col = 17: .value = RS.Fields("COMPOCD").value & ""
                .Col = 18: .value = RS.Fields("ORDCD").value & ""
                .Col = 19: .value = RS.Fields("ORDSEQ").value & ""
                .Col = 20: .value = RS.Fields("ORDNO").value & ""
                
            
                RS.MoveNext
            Loop
            
        End If
    End With
    Set RS = Nothing
End Sub
Public Function DetailBlood(ByVal accdt As String, ByVal accseq As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    Dim strTmp As String
    
    SSQL = " SELECT B.IRRFG,A.COMPOCD,A.RSTV,A.BLDSRC,A.BLDYY,A.BLDNO,A.VFYDT,A.VFYTM,A.VFYID,A.CANCELFG,B.STSCD,B.ABO,B.RH,B.VOLUMN"
    SSQL = SSQL & " FROM " & T_BBS401 & " B," & T_BBS302 & " A"
    SSQL = SSQL & " WHERE " & _
                    DBW("A.WORKAREA=", C_WORKAREA) & _
           " AND " & DBW("A.ACCDT=", accdt) & _
           " AND " & DBW("A.ACCSEQ=", accseq) & _
           " AND A.BLDSRC=B.BLDSRC AND A.BLDYY=B.BLDYY AND A.BLDNO=B.BLDNO AND A.COMPOCD=B.COMPOCD" & _
           " ORDER BY rstseq desc"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Select Case RS.Fields("RSTV").value & ""
            Case "1": strTmp = "OK"
            Case Else: strTmp = "NOT"
        End Select
        
        strTmp = strTmp & COL_DIV & _
               RS.Fields("BLDSRC").value & "" & "-" & _
               RS.Fields("BLDYY").value & "" & "-" & _
               Format(RS.Fields("BLDNO").value & "", "000000") & COL_DIV & _
               Format(RS.Fields("VFYDT").value & "", "0###-##-##") & " " & Format(Mid(RS.Fields("VFYTM").value & "", 1, 4) & "", "0#:##") & COL_DIV & _
               GetEmpNm(RS.Fields("VFYID").value & "") & COL_DIV & RS.Fields("IRRFG").value & "" & COL_DIV & _
               RS.Fields("ABO").value & "" & RS.Fields("RH").value & "" & COL_DIV & RS.Fields("VOLUMN").value & "" & COL_DIV
        
        Select Case RS.Fields("CANCELFG").value & ""
            Case "1"
                strTmp = strTmp & "ASSIGN 취소" & COL_DIV
            Case Else
                Select Case RS.Fields("STSCD").value & ""
                    Case "3"
                        strTmp = strTmp & "출고" & COL_DIV
                        SSQL = "SELECT B.DELIVERYDT,B.DELIVERYTM,B.RCVID FROM " & T_BBS402 & " B," & T_BBS302 & " A"
                        SSQL = SSQL & " WHERE " & _
                                          DBW("A.BLDSRC=", RS.Fields("BLDSRC").value & "") & _
                               " AND  " & DBW("A.BLDYY=", RS.Fields("BLDYY").value & "") & _
                               " AND  " & DBW("A.BLDNO=", RS.Fields("BLDNO").value & "") & _
                               " AND  " & DBW("A.COMPOCD=", RS.Fields("COMPOCD").value & "") & _
                               " AND A.BLDSRC=B.BLDSRC AND A.BLDYY=B.BLDYY AND A.BLDNO=B.BLDNO AND A.COMPOCD=B.COMPOCD" & _
                               " AND  (B.EXPFG is null or B.EXPFG='')"

                        Set RS = New Recordset
                        RS.Open SSQL, DBConn
                        If Not RS.EOF Then
                            strTmp = strTmp & Format(RS.Fields("DELIVERYDT").value & "", "0###-##-##") & " " & Format(Mid(RS.Fields("DELIVERYTM").value & "", 1, 4) & "", "0#:##") & COL_DIV & _
                                     GetEmpNm(RS.Fields("RCVID").value & "")

                        End If
                    Case "4"
                        strTmp = strTmp & "폐기" & COL_DIV
                        
                    Case "2"
                        strTmp = strTmp & "ASSIGN" & COL_DIV
                    Case "1"
                        strTmp = strTmp & "반환" & COL_DIV
                End Select
        End Select
    End If
    DetailBlood = strTmp
    
    Set RS = Nothing

End Function

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim ii      As Integer
        Dim RowNo   As Long
        Dim CompoCd As String
    
        With tblBldList
            If Val(txtStart.Text) > .DataRowCnt Or Val(txtStart.Text) = 0 Then
                MsgBox "시작번호를 수정하세요"
                Exit Sub
            End If
            For ii = Val(txtStart.Text) To .DataRowCnt
                .Row = ii
                .Col = 7
                If .value = "" Then
                    RowNo = .Row
                    .Col = 17: CompoCd = .value
                    Exit For
                End If
            Next
        End With
        If RowNo < 1 Then Exit Sub
        txtBldNo.Text = Mid(txtBldNo, 1, Len(txtBldNo.Text) - 2)
        Call TblBloodINPUT(Trim(txtBldNo.Text), CompoCd, RowNo)
        txtBldNo.SelStart = 0
        txtBldNo.SelLength = Len(txtBldNo.Text)
        
       
    End If
End Sub


Private Function TblBloodINPUT(ByVal BloodNum As String, ByVal CompoCd As String, ByVal RowNo As Long)
    Dim objXM  As clsCrossMatching
    Dim strABO As String
    Dim strTmp As String
    Dim ii     As Integer

    
    Set objXM = New clsCrossMatching
    
    
    If objXM.Get_BloodINfo(BloodNum, CompoCd, ObjMyUser.EmpId, ObjSysInfo.BuildingCd, txtPtid.Text) = False Then
        Set objXM = Nothing
        Exit Function
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
        strTmp = MsgBox("환자 혈액형과 혈액의 혈액형이 동일하지 않습니다." & vbCrLf & "결과등록을 계속진행하시겠습니까?", vbInformation + vbYesNo, Me.Caption)
        If strTmp = vbNo Then
            Set objXM = Nothing
            Exit Function
        End If
    End If
    '-----------------------
    '헌혈부적격 판정여부체크
    '-----------------------
    
    
    With tblBldList
        If BloodDupChk(BloodNum & COL_DIV & CompoCd) = True Then Exit Function
        .Row = RowNo
        .Col = 16: .value = BloodNum & COL_DIV & CompoCd
        '혈액형
        .Col = 4:     .value = medGetP(objXM.strTmp, 1, vbTab)
        '용량
        .Col = 5:     .value = medGetP(objXM.strTmp, 2, vbTab)
        '혈액번호
        .Col = 7:  .value = Mid(BloodNum, 1, 2) & "-" & Mid(BloodNum, 3, 2) & "-" & Format(Mid(BloodNum, 5), "00000#")
        'OK
        .Col = 8: .value = 1
        
        'RT
        .Col = 10: .value = IIf(medGetP(objXM.strTmp, 4, vbTab) = "1", "1", "0")
        If .value = "1" Then .Lock = True
        
        .Col = 11: .value = Format(GetSystemDate, "YYYY-MM-DD")
        .Col = 12: .value = ObjSysInfo.EmpNm
        
    End With
    
    Set objXM = Nothing
End Function
Private Function BloodDupChk(ByVal pBldNo As String) As Boolean
'중복값을 체크한다.(true:dup)
    Dim ii As Integer
    
    With tblBldList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 16
            If .value = pBldNo Then
                BloodDupChk = True
                MsgBox "이미 결과등록대기중인 혈액입니다.", vbInformation + vbOKOnly, "XM결과등록"
                Exit Function
            End If
        Next
    End With
    
End Function
Private Function TagPrint(ByVal BloodNo As String, ByVal componm As String, ByVal ABO As String, ByVal Volumn As String, ByVal orddt As String, ByVal ordno As String, Optional ByVal Rt As String = "")
'-------------
'혈액 Tag 출력
'-------------

    Dim WardDept As String
    Dim vfydt    As String
    Dim PtId     As String
    Dim ptnm     As String
    Dim RS       As Recordset
    Dim SSQL     As String
    
    PtId = txtPtid.Text
    ptnm = lblPtNm.Caption
    
    SSQL = " SELECT WARDID ,DEPTCD,hosilid FROM " & T_LAB101 & _
           " WHERE " & _
                     DBW("PTID=", PtId) & _
           " AND " & DBW("ORDDT=", orddt) & _
           " AND " & DBW("ORDNO=", ordno)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        
        WardDept = RS.Fields("WARDID").value & ""
        If WardDept = "" Then
            WardDept = RS.Fields("DEPTCD").value & ""
        Else
            WardDept = WardDept & "-" & RS.Fields("hosilid").value & ""
        End If
    End If
    
    Set RS = Nothing
    
    
    vfydt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    vfydt = Mid(vfydt, 3, 2) & "-" & Mid(vfydt, 5, 2) & "-" & Mid(vfydt, 7)

    '2001-12-26 수정

    Dim objBar As New clsBarcode
'    Set objBar.MyDB = dbconn
    Set objBar.TableInfo = New clsTables

    objBar.ProjectCd = "BAG"
    objBar.GetBarConfig
    '2001-11-27수정 :
    '혈액라벨의 출력장수 조절 (BLOOD_LABEL_CNT)
    objBar.BloodLabel_PrintOut ptnm, PtId, WardDept, componm, Volumn, ABO, BloodNo, vfydt, BLOOD_LABEL_CNT
    Set objBar = Nothing


End Function
Private Sub cmdSave_Click()
    Dim ii As Integer
    Dim SaveFg As Boolean
    
    
    With tblBldList
        If .DataRowCnt = 0 Then Exit Sub
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 7
            If .value <> "" Then
                SaveFg = True
                Exit For
            End If
        
        Next
        
    End With
    
    If SaveFg = False Then Exit Sub
    
    If BLOOD_SAVE = True Then
        MsgBox "저장되었습니다.", vbInformation + vbOKOnly, "결과저장"
        txtPtid.SelStart = 0
        txtPtid.SelLength = Len(txtPtid.Text)
        txtBldNo.Text = ""
        Call medClearTable(tblBldList)
        Call QueryBlood
    End If
    txtStart.Text = 1
End Sub
Private Function BLOOD_SAVE_HOLD(ByVal qWorkArea As String, ByVal qAccDT As String, ByVal qAccseq As String)
    'RUN_CHECK
    Dim OCSORDNO   As String
    Dim RS         As Recordset
    Dim RunRs      As Recordset

    Dim V_HOSPNO   As String
    Dim V_BSDATE   As String
    Dim V_ORDDATE  As String
    Dim V_DRCODE   As String
    Dim V_PUMMOK   As String
    Dim V_DEPT     As String
    Dim V_INOUT    As String
    Dim SSQL       As String
    
    On Error GoTo HOLD_ERROR
    
    DBConn.BeginTrans
    
'    STRORDDATE = GetSystemDate
    
    SSQL = " select a.ocsordno, b.volumn from " & T_LAB102 & " a," & T_BBS001 & " b " _
         & "  where " & DBW("workarea=", qWorkArea) _
         & "    and " & DBW("accdt=", qAccDT) _
         & "    and " & DBW("accseq=", qAccseq) _
         & "    and a.ordcd=b.testcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If RS.EOF = False Then
        OCSORDNO = CStr(RS.Fields("ocsordno").value & "")
    End If
    Set RS = Nothing
    
    
    If OCSORDNO <> "" Or OCSORDNO <> "0" Then
        '-- 관련정보 불러오기(Parameter : 처방번호)
        SSQL = "SELECT * " & _
               "From TO_GUMSA_LAB " & _
               "WHERE serial = " & Val(OCSORDNO) & " " & _
               "AND bstime = (SELECT MAX(bstime) FROM TO_GUMSA_LAB WHERE serial =" & Val(OCSORDNO) & ")"
        
        Set RunRs = New Recordset
        RunRs.Open SSQL, DBConn
        If RunRs.EOF = False Then
            V_HOSPNO = RunRs.Fields("hospno").value & ""
            V_BSDATE = RunRs.Fields("bsdate").value & ""
            V_ORDDATE = Format(RunRs.Fields("oddate").value & "", "YYYYMMDDHHMMSS")
            V_DRCODE = RunRs.Fields("drcode").value & ""
            V_PUMMOK = RunRs.Fields("pummok").value & ""
            V_DEPT = RunRs.Fields("dept").value & ""
            V_INOUT = RunRs.Fields("inout").value & ""
       
            SSQL = "INSERT INTO TO_RUNCHK(hospno,bsdate,dept,drcode,pummok,oddate,inout,runchk) " _
                 & " VALUES(" _
                 & "'" & V_HOSPNO & "','" & V_BSDATE & "'," _
                 & "'" & V_DEPT & "','" & V_DRCODE & "'," _
                 & "'" & V_PUMMOK & "',TO_DATE('" & V_ORDDATE & "','YYYYMMDDHH24MISS')," _
                 & "" & V_INOUT & ",'1')"
            DBConn.Execute (SSQL)
        End If
    End If
    
    DBConn.CommitTrans
    Set RS = Nothing
    Set RunRs = Nothing
    Exit Function
    
HOLD_ERROR:
    DBConn.RollbackTrans
    Set RS = Nothing
    Set RunRs = Nothing
    
End Function
Private Function BLOOD_SAVE() As Boolean
'Cross-Matching 결과내역 작성
    Dim WorkArea As String, accdt As String, accseq As String, RSTSEQ As String, BldSrc As String
    Dim BldYY As String, BldNo As String, CompoCd As String, RSTV As String, spcyy As String
    Dim spcno As String, vfydt As String, VFYTM As String, VFYID As String, STAT As String
    Dim STATDT As String, STATTM As String, STATID As String, Rt As String
    Dim PtId As String, ordno As String, orddt As String
    Dim BARBLDNUM As String, componm As String, ABO As String, VOL As String, Ordseq As String
    
    Dim ii         As Integer
    
    WorkArea = C_WORKAREA
    vfydt = Format(GetSystemDate, "YYYYMMDD")
    VFYTM = Format(GetSystemDate, "HHMMSS")
    VFYID = ObjSysInfo.EmpId
    spcyy = medGetP(lblSpcNo.Caption, 1, "-")
    spcno = medGetP(lblSpcNo.Caption, 2, "-")
    PtId = txtPtid.Text
    
    On Error GoTo SAVE_ERROR
    
    With tblBldList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 7
            If .value <> "" Then
                .Col = 6:  accdt = Trim(medGetP(.value, 1, "-")):
                           accseq = Trim(medGetP(.value, 2, "-"))
                           RSTSEQ = Get_RstSeq(accdt, accseq)
                .Col = 7:  BldSrc = Trim(medGetP(.value, 1, "-"))
                           BldYY = Trim(medGetP(.value, 2, "-"))
                           BldNo = Trim(medGetP(.value, 3, "-"))
                           BARBLDNUM = Trim(.value)
                           
                .Col = 17: CompoCd = Trim(.value)
                .Col = 8:  RSTV = IIf(.value = 1, "1", "0")
                .Col = 14:
                If .value = 1 Then
                    STAT = "1"
                    STATDT = vfydt: STATTM = VFYTM: STATID = VFYID
                Else
                    STAT = "": STATDT = "": STATTM = "": STATID = ""
                End If
                
                .Col = 9:  Rt = IIf(.value = 1, "1", "")
                .Col = 1:  orddt = Replace(.value, "-", "")
                .Col = 20: ordno = Trim(.value)
                .Col = 3:  componm = Trim(.value)
                .Col = 4:  ABO = Trim(.value)
                .Col = 5:  VOL = Trim(.value)
                .Col = 19: Ordseq = Trim(.value)
                If RSTV = "1" Then
                    '결과등록
                    If SetBBSSAVE(WorkArea, accdt, accseq, RSTSEQ, BldSrc, BldYY, BldNo, CompoCd, RSTV, spcyy, spcno, _
                                  vfydt, VFYTM, VFYID, STAT, STATDT, STATTM, STATID, PtId, orddt, ordno, Rt, Ordseq) = False Then GoTo SAVE_ERROR
                    
                    
                    '바코드 출력
                    Call TagPrint(BARBLDNUM, componm, ABO, VOL, orddt, ordno, Rt)
                    
                    'RUNCHECK
                    Call BLOOD_SAVE_HOLD(WorkArea, accdt, accseq)
                
                End If
            End If
        Next
    End With
    BLOOD_SAVE = True
    Exit Function
    
SAVE_ERROR:
    MsgBox "저장중에 오류발생입니다.", vbInformation + vbOKOnly, "결과등록"
    
End Function



Private Function SetBBSSAVE(ByVal WorkArea As String, ByVal accdt As String, ByVal accseq As String, _
                           ByVal RSTSEQ As String, ByVal BldSrc As String, ByVal BldYY As String, _
                           ByVal BldNo As String, ByVal CompoCd As String, ByVal RSTV As String, _
                           ByVal spcyy As String, ByVal spcno As String, ByVal vfydt As String, _
                           ByVal VFYTM As String, ByVal VFYID As String, ByVal STAT As String, _
                           ByVal STATDT As String, ByVal STATTM As String, ByVal STATID As String, _
                           ByVal PtId As String, ByVal orddt As String, ByVal ordno As String, ByVal Rt As String, ByVal Ordseq As String) As Boolean


    Dim objXM As New clsCrossMatching
    Dim SSQL As String
    
    On Error GoTo Result_Save_Error
    
    DBConn.BeginTrans
    
        
    '결과내역 저장
    SSQL = " INSERT INTO " & T_BBS302 & "(WORKAREA,ACCDT,ACCSEQ,RSTSEQ,BLDSRC,BLDYY,BLDNO,COMPOCD," & _
           "               RSTV,SPCYY,SPCNO,VFYDT,VFYTM,VFYID,STAT,STATDT,STATTM,STATID,STEP1,STEP2,STEP3,STEP4) VALUES (" & _
           DBV("WORKAREA", WorkArea, 1) & DBV("ACCDT", accdt, 1) & DBV("ACCSEQ", accseq, 1) & DBV("RSTSEQ", RSTSEQ, 1) & _
           DBV("BLDSRC", BldSrc, 1) & DBV("BLDYY", BldYY, 1) & DBV("BLDNO", BldNo, 1) & DBV("COMPOCD", CompoCd, 1) & _
           DBV("RSTV", RSTV, 1) & DBV("SPCYY", spcyy, 1) & DBV("SPCNO", spcno, 1) & DBV("VFYDT", vfydt, 1) & DBV("VFYTM", VFYTM, 1) & _
           DBV("VFYID", VFYID, 1) & DBV("STAT", STAT, 1) & DBV("STATDT", STATDT, 1) & DBV("STATTM", STATTM, 1) & DBV("STATID", STATID, 1) & _
           DBV("STEP1", "1", 1) & DBV("STEP2", "1", 1) & DBV("STEP3", "1", 1) & DBV("STEP4", "1") & ")"

    DBConn.Execute SSQL
    
    '혈액입고내역 업데이트
    SSQL = SETBBS401(BldSrc, BldYY, BldNo, CompoCd, BBSBloodStatus.stsASSIGN)

    DBConn.Execute SSQL
    
    'IRRADIATOIN 등록
    If Rt = "1" Then
        SSQL = SetBBS401_IRRADD(BldSrc, BldYY, BldNo, CompoCd)
        DBConn.Execute SSQL
    End If
    '처방별 ASSIGN COUNT등록
    SSQL = SETBBS203(accdt, accseq)
    DBConn.Execute SSQL
    
    '---------------------------------------------------------------------
    '처방과 관련된 테이블을 update 해준다.(처방바디,처방헤더,처방접수내역)
    '---------------------------------------------------------------------
    SSQL = Update_OrderStatus(PtId, orddt, ordno)
    DBConn.Execute SSQL
    
    SSQL = Update_OrderStatus(PtId, orddt, ordno, Ordseq)
    DBConn.Execute SSQL
    
    SSQL = Update_BBS202(accdt, accseq)
    DBConn.Execute SSQL

    DBConn.CommitTrans
    
    SetBBSSAVE = True
    Exit Function
    
Result_Save_Error:
    
    If SetBBSSAVE = False Then
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If

End Function
Private Function Update_BBS202(ByVal accdt As String, ByVal accseq As String)
'처방 접수내역의 stscd를 진행중으로 바꾸자
    Update_BBS202 = " update " & T_BBS202 & " set " & _
                                              DBW("stscd", "3", 2) & _
                    " where " & _
                            "     " & DBW("workarea", C_WORKAREA, 2) & _
                            " and " & DBW("accdt", accdt, 2) & _
                            " and " & DBW("accseq", accseq, 2)
                   
End Function
Private Function Update_OrderStatus(ByVal PtId As String, ByVal orddt As String, ByVal ordno As String, _
                                   Optional ByVal Ordseq As String = "")
'결과등록시 처방부분(LAB101,LAB102)의 status를 진행상태로 바꿔준다.(frmBBS102에서 사용함)
'donefg,stscd를 진행중(3)으로 Update 해준다.
    If Ordseq = "" Then
        Update_OrderStatus = " update " & T_LAB101 & " set " & DBW("donefg", "3", 2) & _
                             " where " & DBW("ptid", PtId, 2) & " and " & DBW("orddt", orddt, 2) & " and " & DBW("ordno", ordno, 2)
    Else
        Update_OrderStatus = " update " & T_LAB102 & _
                             " set " & DBW("donefg", "3", 3) & _
                                       DBW("stscd", "3", 2) & _
                             " where" & _
                                    "     " & DBW("ptid", PtId, 2) & _
                                    " and " & DBW("orddt", orddt, 2) & _
                                    " and " & DBW("ordno", ordno, 2) & _
                                    " and " & DBW("ordseq", Ordseq, 2)
    End If

End Function

Private Function SETBBS203(ByVal accdt As String, ByVal accseq As String) As String

'결과등록시 등록하고자 하는 혈액의 갯수만큼 Insert or Update 해준다.
'테이블에 해당환자의 존재여부를 먼저 체크한다.
'이미 존재하면, Update 해준다.
    Dim RS        As New Recordset
    Dim SSQL      As String
    Dim AssignCnt As Long
    Dim Cnt       As Long: Cnt = 1
    
    SSQL = " select assigncnt from " & T_BBS203 & _
           " where " & _
                   "      " & DBW("workarea", C_WORKAREA, 2) & _
                   "  and " & DBW("accdt", accdt, 2) & _
                   "  and " & DBW("accseq", accseq, 2)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If RS.EOF = True Then
        SETBBS203 = " insert into " & T_BBS203 & "(workarea,accdt,accseq,assigncnt,deliverycnt," & _
                                                 " retcnt,expcnt,bagcnt,assigncancelcnt)" & _
                       " values(" & _
                                DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("assigncnt", Cnt, 1) & _
                                DBV("deliverycnt", "", 1) & DBV("retcnt", "", 1) & DBV("expcnt", "", 1) & DBV("bagcnt", "", 1) & DBV("assigncancelcnt", "") & ")"
    Else
        AssignCnt = Cnt + RS.Fields("assigncnt").value & ""
        SETBBS203 = " update " & T_BBS203 & " set " & DBW("assigncnt", AssignCnt, 2) & _
                        " where " & _
                                "      " & DBW("workarea", C_WORKAREA, 2) & _
                                "  and " & DBW("accdt", accdt, 2) & _
                                "  and " & DBW("accseq", accseq, 2)
 
    End If
    Set RS = Nothing
End Function
Private Function SETBBS401(ByVal BldSrc As String, ByVal BldYY As String, ByVal BldNo As String, _
                              ByVal compcd As String, ByVal stscd As String) As String
'혈액입고내역(BBS401)에 stscd를 Assign 상태로 바꾸어준다.
'조건 bldsrc,bldyy,bldno,compcd
    
    SETBBS401 = " update " & T_BBS401 & " set " & DBW("stscd", stscd, 2) & _
                    " where " & _
                                      DBW("bldsrc", BldSrc, 2) & _
                            " and " & DBW("bldyy", BldYY, 2) & _
                            " and " & DBW("bldno", BldNo, 2) & _
                            " and " & DBW("compocd", compcd, 2)
End Function
Private Function SetBBS401_IRRADD(ByVal BldSrc As String, ByVal BldYY As String, ByVal BldNo As Long, ByVal compcd As String) As String
'혈액입고내역(BBS401)에 stscd를 Assign 상태로 바꾸어준다.
'조건 bldsrc,bldyy,bldno,compcd

    SetBBS401_IRRADD = " update " & T_BBS401 & " set " & _
                                  DBW("irrfg=", "1", 1) & _
                                  DBW("irrdt=", Format(GetSystemDate, PRESENTDATE_FORMAT), 1) & _
                                  DBW("irrid=", ObjSysInfo.EmpId, 1) & _
                                  DBW("irrtm=", Format(GetSystemDate, PRESENTTIME_FORMAT)) & _
                       " where " & _
                                         DBW("bldsrc", BldSrc, 2) & _
                               " and " & DBW("bldyy", BldYY, 2) & _
                               " and " & DBW("bldno", BldNo, 2) & _
                               " and " & DBW("compocd", compcd, 2)
End Function
Private Function Get_RstSeq(ByVal accdt As String, ByVal accseq As String) As String

'bbs302에 저장시 결과 Seq를 구해온다.(rstseq)
    Dim RS   As Recordset
    Dim SSQL As String
    
    SSQL = " select max(rstseq) as maxrstseq " & _
           " from " & T_BBS302 & _
           " where " & DBW("workarea", C_WORKAREA, 2) & _
           " and   " & DBW("accdt", accdt, 2) & _
           " and   " & DBW("accseq", accseq, 2)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If RS.EOF = False Then
        If IsNull(RS.Fields("maxrstseq").value) = True Then
            Get_RstSeq = 1
        Else
            Get_RstSeq = Val(RS.Fields("maxrstseq").value & "") + 1
        End If
    Else
        Get_RstSeq = 1
    End If
    
    Set RS = Nothing
    
End Function






