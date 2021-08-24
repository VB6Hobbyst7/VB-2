VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmItemCode 
   Caption         =   "ItemCode 등록화면"
   ClientHeight    =   7935
   ClientLeft      =   810
   ClientTop       =   3135
   ClientWidth     =   11865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11865
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel2 
      Height          =   660
      Left            =   4410
      TabIndex        =   38
      Top             =   45
      Width           =   7485
      _Version        =   65536
      _ExtentX        =   13203
      _ExtentY        =   1164
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin MSForms.CommandButton cmdSelectName 
         Height          =   495
         Left            =   6075
         TabIndex        =   77
         Top             =   90
         Width           =   1320
         Caption         =   "코드찾기"
         PicturePosition =   327683
         Size            =   "2328;873"
         Picture         =   "frmItemCode.frx":0000
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   495
         Left            =   4680
         TabIndex        =   71
         Top             =   90
         Width           =   1275
         Caption         =   "화면정리"
         PicturePosition =   327683
         Size            =   "2249;873"
         Picture         =   "frmItemCode.frx":031A
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   495
         Left            =   3420
         TabIndex        =   70
         Top             =   90
         Width           =   1275
         Caption         =   "삭제확인"
         PicturePosition =   327683
         Size            =   "2249;873"
         Picture         =   "frmItemCode.frx":1AAC
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsert 
         Height          =   495
         Left            =   2160
         TabIndex        =   69
         Top             =   90
         Width           =   1275
         Caption         =   "입력확인"
         PicturePosition =   327683
         Size            =   "2249;873"
         Picture         =   "frmItemCode.frx":2386
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQry 
         Height          =   495
         Left            =   105
         TabIndex        =   39
         Top             =   90
         Width           =   1845
         Caption         =   "  조회▼"
         PicturePosition =   327683
         Size            =   "3254;873"
         Picture         =   "frmItemCode.frx":3B48
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   7860
      Left            =   90
      TabIndex        =   35
      Top             =   45
      Width           =   4245
      _Version        =   65536
      _ExtentX        =   7488
      _ExtentY        =   13864
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtCodeGb 
         Height          =   270
         Left            =   3645
         MaxLength       =   2
         TabIndex        =   87
         Top             =   4680
         Width           =   420
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   330
         Left            =   3015
         TabIndex        =   79
         Top             =   6030
         Visible         =   0   'False
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "interface 때문에 쓴대요"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   8
         Begin VB.TextBox txtGeomjan1 
            Height          =   270
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   83
            Top             =   45
            Width           =   630
         End
         Begin VB.TextBox txtGeomjan2 
            Height          =   270
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   82
            Top             =   315
            Width           =   630
         End
         Begin VB.TextBox txtGeomjan3 
            Height          =   270
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   81
            Top             =   585
            Width           =   630
         End
         Begin Threed.SSCommand cmdHelpJan 
            Height          =   270
            Index           =   0
            Left            =   2130
            TabIndex        =   80
            Top             =   45
            Width           =   210
            _Version        =   65536
            _ExtentX        =   370
            _ExtentY        =   476
            _StockProps     =   78
            Caption         =   "&3"
            BevelWidth      =   1
         End
         Begin Threed.SSCommand cmdHelpJan 
            Height          =   270
            Index           =   1
            Left            =   2130
            TabIndex        =   84
            Top             =   315
            Width           =   210
            _Version        =   65536
            _ExtentX        =   370
            _ExtentY        =   476
            _StockProps     =   78
            Caption         =   "&4"
            BevelWidth      =   1
         End
         Begin Threed.SSCommand cmdHelpJan 
            Height          =   270
            Index           =   2
            Left            =   2130
            TabIndex        =   85
            Top             =   585
            Width           =   210
            _Version        =   65536
            _ExtentX        =   370
            _ExtentY        =   476
            _StockProps     =   78
            Caption         =   "&5"
            BevelWidth      =   1
         End
         Begin VB.Label Label13 
            Caption         =   "검사장비1,2,3"
            Height          =   195
            Left            =   225
            TabIndex        =   86
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.TextBox txtChunit 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   1395
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   7290
         Width           =   495
      End
      Begin VB.ComboBox cmbDeltaQc 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmItemCode.frx":4422
         Left            =   3285
         List            =   "frmItemCode.frx":4435
         Style           =   2  '드롭다운 목록
         TabIndex        =   76
         Top             =   5250
         Width           =   465
      End
      Begin VB.ComboBox cmbExGB 
         Height          =   300
         ItemData        =   "frmItemCode.frx":4448
         Left            =   2835
         List            =   "frmItemCode.frx":4452
         Style           =   2  '드롭다운 목록
         TabIndex        =   75
         Top             =   4365
         Width           =   1230
      End
      Begin VB.CheckBox chkBarGb 
         Caption         =   "관리항목"
         Height          =   195
         Left            =   2880
         TabIndex        =   74
         ToolTipText     =   "BarCode Label을 따로 관리하는것을 Check하세요!"
         Top             =   7020
         Width           =   1095
      End
      Begin VB.TextBox txtSlipNo 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtItemCode 
         Height          =   270
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   1
         Top             =   480
         Width           =   1395
      End
      Begin VB.ComboBox cmbGeomsaGb 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmItemCode.frx":4469
         Left            =   1395
         List            =   "frmItemCode.frx":4473
         TabIndex        =   18
         Top             =   4365
         Width           =   1410
      End
      Begin VB.TextBox txtOrderCode 
         Height          =   270
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   25
         Top             =   5790
         Width           =   1440
      End
      Begin VB.TextBox txtPanicmax 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   2145
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   5520
         Width           =   690
      End
      Begin VB.TextBox txtDeltaMax 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   2145
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   5250
         Width           =   690
      End
      Begin VB.TextBox txtCgcmt 
         Height          =   270
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   17
         Top             =   4095
         Width           =   1395
      End
      Begin VB.TextBox txtChcmt 
         Height          =   270
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   16
         Top             =   3825
         Width           =   1395
      End
      Begin VB.TextBox txtDiffEnd 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   2085
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3555
         Width           =   705
      End
      Begin VB.TextBox txtBarText 
         Height          =   300
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   29
         Top             =   6975
         Width           =   1455
      End
      Begin VB.ComboBox cmbBottle 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   1395
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   2385
         Width           =   2490
      End
      Begin Threed.SSCommand cmdNormal 
         Height          =   315
         Left            =   3060
         TabIndex        =   44
         Top             =   3240
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "참조치"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdHelpSuga 
         Height          =   270
         Left            =   2820
         TabIndex        =   43
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   476
         _StockProps     =   78
         Caption         =   "&S"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdRset 
         Height          =   285
         Left            =   2580
         TabIndex        =   42
         Top             =   2100
         Width           =   225
         _Version        =   65536
         _ExtentX        =   406
         _ExtentY        =   512
         _StockProps     =   78
         Caption         =   "&C"
         BevelWidth      =   1
      End
      Begin VB.ComboBox cmbGeomsaW2 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmItemCode.frx":4489
         Left            =   1980
         List            =   "frmItemCode.frx":44A2
         Style           =   2  '드롭다운 목록
         TabIndex        =   8
         Top             =   2100
         Width           =   615
      End
      Begin VB.ComboBox cmbGeomsaW1 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmItemCode.frx":44C2
         Left            =   1395
         List            =   "frmItemCode.frx":44DB
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   2100
         Width           =   615
      End
      Begin Threed.SSCommand cmdHgeomch 
         Height          =   270
         Index           =   1
         Left            =   3480
         TabIndex        =   41
         Top             =   2700
         Width           =   210
         _Version        =   65536
         _ExtentX        =   370
         _ExtentY        =   476
         _StockProps     =   78
         Caption         =   "&2"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdHgeomch 
         Height          =   270
         Index           =   0
         Left            =   2340
         TabIndex        =   40
         Top             =   2700
         Width           =   210
         _Version        =   65536
         _ExtentX        =   370
         _ExtentY        =   476
         _StockProps     =   78
         Caption         =   "&1"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtItemNM 
         Height          =   270
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   2
         Top             =   750
         Width           =   2655
      End
      Begin VB.TextBox txtItemKO 
         Height          =   270
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1020
         Width           =   2655
      End
      Begin VB.TextBox txtYageo 
         Height          =   270
         Left            =   1395
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1290
         Width           =   1695
      End
      Begin VB.TextBox txtSugacd 
         Height          =   270
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1560
         Width           =   1395
      End
      Begin VB.TextBox txtGeomsaTm 
         Height          =   270
         Left            =   1395
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1830
         Width           =   795
      End
      Begin VB.TextBox txtChwhYg 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   270
         Left            =   900
         MaxLength       =   3
         TabIndex        =   33
         Top             =   2400
         Width           =   480
      End
      Begin VB.TextBox txtGeomchc1 
         Height          =   270
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   10
         Top             =   2700
         Width           =   975
      End
      Begin VB.TextBox txtGeomchc2 
         Height          =   270
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   11
         Top             =   2700
         Width           =   975
      End
      Begin VB.TextBox txtDanwi 
         Height          =   270
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   12
         Top             =   2985
         Width           =   855
      End
      Begin VB.ComboBox cmbResultW 
         Height          =   300
         ItemData        =   "frmItemCode.frx":44FB
         Left            =   1395
         List            =   "frmItemCode.frx":450B
         TabIndex        =   13
         Top             =   3240
         Width           =   1635
      End
      Begin VB.TextBox txtDiffStart 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   1395
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3555
         Width           =   675
      End
      Begin VB.TextBox txtGeomsaAb 
         Height          =   270
         Left            =   1395
         MaxLength       =   1
         TabIndex        =   19
         Top             =   4665
         Width           =   735
      End
      Begin VB.TextBox txtDeltaMin 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   1395
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   5250
         Width           =   735
      End
      Begin VB.TextBox txtPanicmin 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Left            =   1395
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   5520
         Width           =   735
      End
      Begin VB.ComboBox cmbCheck 
         Height          =   300
         ItemData        =   "frmItemCode.frx":4546
         Left            =   1395
         List            =   "frmItemCode.frx":4550
         TabIndex        =   27
         Top             =   6360
         Width           =   1455
      End
      Begin VB.ComboBox cmbInput 
         Height          =   300
         ItemData        =   "frmItemCode.frx":456B
         Left            =   1395
         List            =   "frmItemCode.frx":4575
         TabIndex        =   28
         Top             =   6660
         Width           =   1455
      End
      Begin VB.TextBox txtSlipName 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   32
         Top             =   120
         Width           =   2055
      End
      Begin VB.ComboBox cmbRoutine 
         Height          =   300
         ItemData        =   "frmItemCode.frx":4591
         Left            =   1395
         List            =   "frmItemCode.frx":459B
         TabIndex        =   26
         Top             =   6060
         Width           =   1455
      End
      Begin Threed.SSCommand cmdHelpSlip 
         Height          =   315
         Left            =   1380
         TabIndex        =   31
         Top             =   120
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "&H"
         BevelWidth      =   1
      End
      Begin MSComCtl2.DTPicker dtCodate 
         Height          =   315
         Left            =   1395
         TabIndex        =   20
         Top             =   4935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36299
      End
      Begin VB.Label Label11 
         Caption         =   "코드구분"
         Height          =   195
         Left            =   2835
         TabIndex        =   88
         Top             =   4725
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "BarCode매수"
         Height          =   195
         Left            =   90
         TabIndex        =   78
         Top             =   7335
         Width           =   1155
      End
      Begin VB.Label Label37 
         Caption         =   "GBinput"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   6690
         Width           =   1125
      End
      Begin VB.Label Label36 
         Caption         =   "Check유무"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   6420
         Width           =   1275
      End
      Begin VB.Label Label35 
         Caption         =   "Routine구분"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   6135
         Width           =   1275
      End
      Begin VB.Label Label34 
         Caption         =   "OrderCode"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   5865
         Width           =   1275
      End
      Begin VB.Label Label32 
         Caption         =   "Panic(min-max)"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   5580
         Width           =   1275
      End
      Begin VB.Label Label29 
         Caption         =   "Delta/min-max"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   5295
         Width           =   1275
      End
      Begin VB.Label Label28 
         Caption         =   "등록일"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   5025
         Width           =   1185
      End
      Begin VB.Label Label27 
         Caption         =   "검사방법"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   4740
         Width           =   1140
      End
      Begin VB.Label Label26 
         Caption         =   "검사구분"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   4425
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "참고Comment"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   4140
         Width           =   1140
      End
      Begin VB.Label Label24 
         Caption         =   "채취Comment"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   3855
         Width           =   1185
      End
      Begin VB.Label Label22 
         Caption         =   "Diff.c/Diffmax"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   3585
         Width           =   1275
      End
      Begin VB.Label Label17 
         Caption         =   "결과형태"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   3345
         Width           =   1155
      End
      Begin VB.Label Label16 
         Caption         =   "단위"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   3075
         Width           =   1155
      End
      Begin VB.Label Label10 
         Caption         =   "검체1,2"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   2730
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "채혈용기"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   2445
         Width           =   750
      End
      Begin VB.Label Label7 
         Caption         =   "미검사일1,2"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   2175
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "검사소요시간"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "수가코드"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "약어"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Item한글명"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1095
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Item영문명"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   825
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "ItemCode"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "BarCodeText"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   7020
         Width           =   1125
      End
      Begin VB.Label Label31 
         Caption         =   "QC:"
         Height          =   195
         Left            =   2970
         TabIndex        =   37
         Top             =   5310
         Width           =   330
      End
      Begin VB.Label labSlipTitle 
         Caption         =   "검사종류"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   180
         Width           =   795
      End
   End
   Begin Threed.SSPanel panelPg 
      Height          =   300
      Left            =   6075
      TabIndex        =   72
      Top             =   720
      Width           =   5805
      _Version        =   65536
      _ExtentX        =   10239
      _ExtentY        =   529
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   4
      Begin MSComctlLib.ProgressBar pgbSelect 
         Height          =   210
         Left            =   45
         TabIndex        =   73
         Top             =   45
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin FPSpreadADO.fpSpread ssItem 
      Height          =   6945
      Left            =   4365
      TabIndex        =   34
      Top             =   1035
      Width           =   7530
      _Version        =   196608
      _ExtentX        =   13282
      _ExtentY        =   12250
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   34
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmItemCode.frx":45B8
      UserResize      =   1
      VisibleCols     =   34
      VisibleRows     =   500
      Appearance      =   2
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
   Begin VB.Menu mnuPr 
      Caption         =   "출력"
   End
End
Attribute VB_Name = "frmItemCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function isOrderCode(ByVal siTemCode As String) As Integer
    Dim adoODr      As ADODB.Recordset
    
    sMsg = ""
    isOrderCode = False
    
    strSql = ""
    strSql = strSql & " SELECT ItemCd, SLipno, OrderCode, Ordername"
    strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_OrderCode"
    strSql = strSql & " WHERE  iTemCd = '" & Trim(siTemCode) & "'"
    
    If False = adoSetOpen(strSql, adoODr) Then Exit Function
    
    isOrderCode = True
    sMsg = ""
    sMsg = "OrderCode 가 존재합니다." & vbCrLf
    sMsg = sMsg & "OCS OrderSLip = " & adoODr.Fields("SLipno").Value & "" & vbCrLf
    sMsg = sMsg & "OCS OrderCode = " & adoODr.Fields("OrderCode").Value & "" & vbCrLf
    sMsg = sMsg & "OCS OrderName = " & adoODr.Fields("OrderName").Value & "" & vbCrLf
    sMsg = sMsg & "OCS ItemCode  = " & adoODr.Fields("ItemCD").Value & "" & vbCrLf
    sMsg = sMsg & " " & vbCrLf
    sMsg = sMsg & "위 사항을 전산실에 연락하여 제거하셔야 합니다."
    
    
    
End Function

Private Sub cmbBottle_Click()
    
    If cmbBottle.ListIndex = -1 Then Exit Sub
    txtChwhYg.Text = Left(cmbBottle.Text, 4)
    
End Sub

Private Sub cmbGeomsaGb_Click()
    
    If cmbGeomsaGb.ListIndex = -1 Then Exit Sub
    
    
    If Left(cmbGeomsaGb.Text, 1) = "W" Then
        cmbExGB.ListIndex = 0      '첫번째 1. SCL 자동 Check
    Else
        cmbExGB.ListIndex = -1
    End If
    
    
End Sub

Private Sub cmdClear_Click()
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then Me.Controls(i).Text = ""
        If TypeOf Me.Controls(i) Is VB.ComboBox Then Me.Controls(i).ListIndex = -1
        If TypeOf Me.Controls(i) Is DTPicker Then Me.Controls(i).Value = Dual_Date_Get("yyyy-MM-dd")
    Next
    
    
    If vbYes = MsgBox("오른쪽의 Spread Data 도 Clear 하시겠습니까?", vbYesNo _
                       + vbQuestion + vbDefaultButton2, _
                      "화면정리MessageBox") Then GoSub ssItem_Spread_Clear
    
    mdiMain.stbMain.Panels(1).Text = ""
    cmdHelpSlip.SetFocus
    Exit Sub
    
    
ssItem_Spread_Clear:
    ssItem.Row = 1
    ssItem.Row2 = ssItem.DataRowCnt
    ssItem.Col = 1
    ssItem.Col2 = ssItem.DataColCnt
    ssItem.BlockMode = True
    ssItem.Action = ActionClear
    ssItem.BlockMode = False
    ssItem.MaxRows = 25
    Return
    
End Sub

Private Sub cmdDelete_Click()
    Dim siTemCode       As String
    
    
    If Trim(txtSlipno.Text) = "" Then
        MsgBox "삭제할 SLIP 구분을 먼저 선택하세요!"
        Exit Sub
    End If
    
    If Trim(txtItemCode.Text) = "" Then
        MsgBox "삭제할 Item Code 가 없습니다!.."
        Exit Sub
    End If
    
    siTemCode = Trim(txtSlipno.Text) & Trim(txtItemCode.Text)

    
    strSql = ""
    strSql = strSql & " SELECT Codeky"
    strSql = strSql & " FROM   TWEXAM_ITEMML"
    strSql = strSql & " WHERE  Codeky = '" & Trim(siTemCode) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "Database 에 삭제할ItemCode 가 없습니다. 확인바람"
        Exit Sub
    Else
        Call adoSetClose(adoSet)
        
        
        GoSub Delete_ItemCode
        GoSub Set_DeleteItem_Color_Change
        
        If isOrderCode(Trim(siTemCode)) Then
            MsgBox sMsg, vbOKOnly, "OrderCode 삭제를....."
        End If
        
        'Call ClearForm(Me)
        cmdHelpSlip.SetFocus
    End If
    
    Exit Sub
    
    
Set_DeleteItem_Color_Change:
    For i = 1 To ssItem.DataRowCnt
        ssItem.Row = i
        ssItem.Col = 1
        If Trim(siTemCode) = Trim(ssItem.Text) Then
            ssItem.Text = "*" & ssItem.Text
            ssItem.Row = i: ssItem.Row2 = i
            ssItem.Col = 1: ssItem.Col2 = ssItem.MaxCols
            ssItem.BlockMode = True
            ssItem.ForeColor = RGB(255, 0, 0)
            ssItem.BlockMode = False
        End If
    Next
    
    
    Return
    
    
Delete_ItemCode:
    sMsg = siTemCode & " 를 삭제하시겠습니까?"
    If vbNo = MsgBox(sMsg, vbYesNo + vbQuestion + vbDefaultButton2, "삭제확인 Box") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TWEXAM_ITEMML"
    strSql = strSql & " WHERE  Codeky = '" & Trim(siTemCode) & "'"
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = "삭제가 되었습니다!.."
    Else
        mdiMain.stbMain.Panels(1).Text = "어떠한 오류로 인하여 삭제되지 않았습니다!"
    End If
    
    Return
    
End Sub

Private Sub cmdHelpJan_Click(Index As Integer)
    
    Select Case Index
        Case 0: hWndReturn = txtGeomjan1.hwnd
        Case 1: hWndReturn = txtGeomjan2.hwnd
        Case 2: hWndReturn = txtGeomjan3.hwnd
    End Select
    
    frmQryJangbi.Show vbModal
    
    
End Sub

Private Sub cmdHelpSlip_Click()
    
    mdiMain.stbMain.Panels(1).Text = ""
    
    gCallWin = 1
    
    frmSlipQry.Show vbModal
    If Trim(txtSLipName.Text) <> "" Then
        txtItemCode.SetFocus
    End If
    
End Sub

Private Sub cmdHelpSuga_Click()
    
    hWndReturn = txtSugacd.hwnd
    frmQrySuga.Show vbModal
    
End Sub

Private Sub cmdHgeomch_Click(Index As Integer)
    
    
    Select Case Index
        Case 0: hWndReturn = txtGeomchc1.hwnd
        Case 1: hWndReturn = txtGeomchc2.hwnd
    End Select
    
    frmQryGeom.Show vbModal
    
    
End Sub

Private Sub cmdInsert_Click()
    Dim siTemCode       As String
    Dim sBarGB          As String
    Dim sExGb           As String
    
    
    On Error GoTo DBPut_Error_Code
    
    If Trim(txtSlipno.Text) = "" Then
        MsgBox "입력할 SLIP 구분을 먼저 선택하세요!"
        Exit Sub
    End If
    
    If Trim(txtItemCode.Text) = "" Then
        MsgBox "입력할 Item Code 가 없습니다!.."
        Exit Sub
    End If
    
    siTemCode = Trim(txtSlipno.Text) & Trim(txtItemCode.Text)
    
    If chkBarGb.Value = "1" Then       'BarCode 를 따로 관리하는 항목 Check
        sBarGB = "1"
    Else
        sBarGB = ""
    End If
    
    If Left(cmbGeomsaGb.Text, 1) = "W" Then
        sExGb = Left(cmbExGB.Text, 1)
    Else
        sExGb = ""
    End If
    
    strSql = ""
    strSql = strSql & " SELECT Codeky"
    strSql = strSql & " FROM   TWEXAM_ITEMML"
    strSql = strSql & " WHERE  Codeky = '" & Trim(siTemCode) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub Main_Insert_Sub
    Else
        Call adoSetClose(adoSet)
        GoSub Main_Update_Sub
    End If

    GoSub Form_Reset
    
    Exit Sub
    
'/_____________________________________________________________________________
Main_Insert_Sub:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_ITEMML"
    strSql = strSql & "       (     Codeky,    Itemnm,    ItemKo,    Yageo,     Sugacd,    GeomsaTm, "
    strSql = strSql & "             GeomsaW1,  GeomsaW2,  ChwhYg,    GeomChc1,  GeomChc2,  ChUnit,   "
    strSql = strSql & "             GeomJan1,  GeomJan2,  GeomJan3,  Danwi,     ResultW,             "
    strSql = strSql & "             DiffCount, MaxDiffc, "
    strSql = strSql & "             ChComment, CgComment, GeomsaGb,  GeomsaAb,  Codate,              "
    strSql = strSql & "             DeltaMin,  DeltaMax,  DeltaQC,   PanicMin,  PanicMax,  OrderCD,  "
    strSql = strSql & "             GbRoutine, GbCheck,   GbInput,   Codegu,    BarText,   BarGb,     EXGB   )"
    strSql = strSql & " VALUES('" & siTemCode & "',"
    strSql = strSql & "        '" & Quot_Conv(RTrim(txtItemNM.Text)) & "',"
    strSql = strSql & "        '" & Quot_Conv(RTrim(txtItemKO.Text)) & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtYageo.Text)) & "',"
    strSql = strSql & "        '" & Trim(txtSugacd.Text) & "',"
    strSql = strSql & "         " & Val(txtGeomsaTm.Text) & ","
    strSql = strSql & "        '" & Trim(cmbGeomsaW1.Text) & "',"
    strSql = strSql & "        '" & Trim(cmbGeomsaW2.Text) & "',"
    strSql = strSql & "        '" & Trim(txtChwhYg.Text) & "',"
    strSql = strSql & "        '" & Trim(txtGeomchc1.Text) & "',"
    strSql = strSql & "        '" & Trim(txtGeomchc2.Text) & "',"
    strSql = strSql & "        '" & Trim(txtChunit.Text) & "',"
    strSql = strSql & "        '" & Trim(txtGeomjan1.Text) & "',"
    strSql = strSql & "        '" & Trim(txtGeomjan2.Text) & "',"
    strSql = strSql & "        '" & Trim(txtGeomjan3.Text) & "',"
    strSql = strSql & "        '" & Trim(txtDanwi.Text) & "',"
    strSql = strSql & "        '" & Left(Trim(cmbResultW.Text), 1) & "',"
    strSql = strSql & "         " & Val(txtDiffStart.Text) & ","
    strSql = strSql & "         " & Val(txtDiffEnd.Text) & ","
    strSql = strSql & "        '" & Trim(txtChcmt.Text) & "',"
    strSql = strSql & "        '" & Trim(txtCgcmt.Text) & "',"
    strSql = strSql & "        '" & Left(Trim(cmbGeomsaGb.Text), 1) & "',"
    strSql = strSql & "        '" & Left(Trim(txtGeomsaAb.Text), 1) & "',"
    strSql = strSql & "        TO_DATE('" & Format(dtCodate.Value, "yyyy-MM-dd") & "','YYYY-MM-DD'),"
    strSql = strSql & "         " & Val(txtDeltaMin.Text) & ","
    strSql = strSql & "         " & Val(txtDeltaMax.Text) & ","
    strSql = strSql & "        '" & Trim(cmbDeltaQc.Text) & "',"
    strSql = strSql & "         " & Val(txtPanicmin.Text) & ","
    strSql = strSql & "         " & Val(txtPanicmax.Text) & ","
    strSql = strSql & "        '" & Trim(txtOrderCode.Text) & "',"
    strSql = strSql & "        '" & Left(Trim(cmbRoutine.Text), 1) & "',"
    strSql = strSql & "        '" & Left(Trim(cmbCheck.Text), 1) & "',"
    strSql = strSql & "        '" & Left(Trim(cmbInput.Text), 1) & "',"
    strSql = strSql & "        '" & Trim(txtCodeGb.Text) & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtBarText.Text)) & "',"
    strSql = strSql & "        '" & sBarGB & "',"
    strSql = strSql & "        '" & sExGb & "')"
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = siTemCode & " 가 입력 되었습니다!"
    Else
        mdiMain.stbMain.Panels(1).Text = siTemCode & " 가 어떤오류로 인하여 입력되지 않았습니다!."
        Exit Sub
    End If
    
    Return
    
    
    
Main_Update_Sub:
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_ITEMML"
    strSql = strSql & " SET     Itemnm     =  '" & Quot_Conv(RTrim(txtItemNM.Text)) & "',"
    strSql = strSql & "         ItemKo     =  '" & Quot_Conv(RTrim(txtItemKO.Text)) & "',"
    strSql = strSql & "         Yageo      =  '" & Quot_Conv(Trim(txtYageo.Text)) & "',"
    strSql = strSql & "         Sugacd     =  '" & Trim(txtSugacd.Text) & "',"
    strSql = strSql & "         GeomsaTm   =   " & Val((txtGeomsaTm.Text)) & ","
    strSql = strSql & "         GeomsaW1   =  '" & Trim(cmbGeomsaW1.Text) & "',"
    strSql = strSql & "         GeomsaW2   =  '" & Trim(cmbGeomsaW2.Text) & "',"
    strSql = strSql & "         ChwhYg     =  '" & Trim(txtChwhYg.Text) & "',"
    strSql = strSql & "         GeomChc1   =  '" & Trim(txtGeomchc1.Text) & "',"
    strSql = strSql & "         GeomChc2   =  '" & Trim(txtGeomchc2.Text) & "',"
    strSql = strSql & "         ChUnit     =  '" & Trim(txtChunit.Text) & "',"
    strSql = strSql & "         GeomJan1   =  '" & Trim(txtGeomjan1.Text) & "',"
    strSql = strSql & "         GeomJan2   =  '" & Trim(txtGeomjan2.Text) & "',"
    strSql = strSql & "         GeomJan3   =  '" & Trim(txtGeomjan3.Text) & "',"
    strSql = strSql & "         Danwi      =  '" & Trim(txtDanwi.Text) & "',"
    strSql = strSql & "         ResultW    =  '" & Left(Trim(cmbResultW.Text), 1) & "',"
    strSql = strSql & "         DiffCount  =   " & Val(txtDiffStart.Text) & ","
    strSql = strSql & "         MaxDiffc   =   " & Val(txtDiffEnd.Text) & ","
    strSql = strSql & "         ChComment  =  '" & Trim(txtChcmt.Text) & "',"
    strSql = strSql & "         CgComment  =  '" & Trim(txtCgcmt.Text) & "',"
    strSql = strSql & "         GeomsaGb   =  '" & Left(Trim(cmbGeomsaGb.Text), 1) & "',"
    strSql = strSql & "         GeomsaAb   =  '" & Left(Trim(txtGeomsaAb.Text), 1) & "',"
    strSql = strSql & "         Codate     =  TO_DATE('" & Format(dtCodate.Value, "yyyy-MM-dd") & "','YYYY-MM-DD'),"
    strSql = strSql & "         DeltaMin   =   " & Val(txtDeltaMin.Text) & ","
    strSql = strSql & "         DeltaMax   =   " & Val(txtDeltaMax.Text) & ","
    strSql = strSql & "         DeltaQC    =  '" & Trim(cmbDeltaQc.Text) & "',"
    strSql = strSql & "         PanicMin   =   " & Val(txtPanicmin.Text) & ","
    strSql = strSql & "         PanicMax   =   " & Val(txtPanicmax.Text) & ","
    strSql = strSql & "         OrderCD    =  '" & Trim(txtOrderCode.Text) & "',"
    strSql = strSql & "         GbRoutine  =  '" & Left(Trim(cmbRoutine.Text), 1) & "',"
    strSql = strSql & "         GbCheck    =  '" & Left(Trim(cmbCheck.Text), 1) & "',"
    strSql = strSql & "         GbInput    =  '" & Left(Trim(cmbInput.Text), 1) & "',"
    strSql = strSql & "         Codegu     =  '" & Trim(txtCodeGb.Text) & "',"
    strSql = strSql & "         BarText    =  '" & Quot_Conv(Trim(txtBarText.Text)) & "',"
    strSql = strSql & "         BarGb      =  '" & sBarGB & "',"
    strSql = strSql & "         EXGB       =  '" & sExGb & "'"
    strSql = strSql & " WHERE   Codeky  =  '" & Trim(siTemCode) & "'"
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = siTemCode & " 가 수정 되었습니다!"
    Else
        mdiMain.stbMain.Panels(1).Text = siTemCode & " 가 어떤오류로 인하여 수정되지 않았습니다!."
        Exit Sub
    End If
    
        
    Return
    
Form_Reset:
    txtSlipno.Tag = txtSlipno.Text
    txtSLipName.Tag = txtSLipName.Text
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then
            Me.Controls(i).Text = ""
        ElseIf TypeOf Me.Controls(i) Is ComboBox Then
            If Me.Controls(i).Style = vbComboDropdownList Then
                Me.Controls(i).ListIndex = -1
            Else
                Me.Controls(i).Text = ""
            End If
        End If
    Next
    txtSlipno.Text = txtSlipno.Tag
    txtSLipName.Text = txtSLipName.Tag
    Return
    
DBPut_Error_Code:
    MsgBox Err.Description
    Exit Sub
    Return
End Sub

Private Sub cmdNormal_Click()
    frmNormalEdit.Show vbModal
    
End Sub

Private Sub cmdQry_Click()
    
    If Trim(txtSlipno.Text) = "" Then
        GoSub ItemForm_Clear
        gCallWin = 1
        frmSlipQry.Show vbModal
    End If
    
    Screen.MousePointer = vbHourglass
    
    mdiMain.stbMain.Panels(1).Text = ""
    GoSub ITEMCode_Get_Proc
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
'/______________________________________________
ItemForm_Clear:
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then Me.Controls(i).Text = ""
        If TypeOf Me.Controls(i) Is VB.ComboBox Then Me.Controls(i).ListIndex = -1
        If TypeOf Me.Controls(i) Is DTPicker Then Me.Controls(i).Value = Dual_Date_Get("yyyy-MM-dd")
    Next
    
    
    mdiMain.stbMain.Panels(1).Text = ""
    ssItem.Row = 1
    ssItem.Row2 = ssItem.DataRowCnt
    ssItem.Col = 1
    ssItem.Col2 = ssItem.DataColCnt
    ssItem.BlockMode = True
    ssItem.Action = ActionClear
    ssItem.BlockMode = False
    ssItem.MaxRows = 25
    Return


ITEMCode_Get_Proc:
    strSql = ""
    strSql = strSql & " SELECT a.*, TO_CHAR(a.Codate, 'yyyy-MM-dd') Codate"
    strSql = strSql & " FROM   TWEXAM_ITEMML a"
    strSql = strSql & " WHERE  SUBSTR(a.CODEKY, 1, 2) = '" & Trim(txtSlipno.Text) & "'"
    strSql = strSql & " ORDER  BY Codeky"
    
    ssItem.ReDraw = False
    ssItem.MaxRows = 0:
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssItem.MaxRows = adoSet.RecordCount
    
    panelPg.Visible = True
    pgbSelect.Max = adoSet.RecordCount
    pgbSelect.Min = 0
    pgbSelect.Value = 0
    
    DoEvents
    pgbSelect.ZOrder 0
    Do Until adoSet.EOF
        ssItem.Row = ssItem.DataRowCnt + 1
        ssItem.Col = 1:  ssItem.Text = adoSet.Fields("Codeky").Value & ""
        ssItem.Col = 2:  ssItem.Text = adoSet.Fields("Itemnm").Value & ""
        ssItem.Col = 3:  ssItem.Text = adoSet.Fields("Itemko").Value & ""
        ssItem.Col = 4:  ssItem.Text = adoSet.Fields("Yageo").Value & ""
        ssItem.Col = 5:  ssItem.Text = adoSet.Fields("Sugacd").Value & ""
        ssItem.Col = 6:  ssItem.Text = adoSet.Fields("GeomsaTm").Value & ""
        ssItem.Col = 7:  ssItem.Text = adoSet.Fields("GeomsaW1").Value & ""
        ssItem.Col = 8:  ssItem.Text = adoSet.Fields("GeomsaW2").Value & ""
        ssItem.Col = 9:  ssItem.Text = adoSet.Fields("ChwhYg").Value & ""
        ssItem.Col = 10: ssItem.Text = adoSet.Fields("GeomchC1").Value & ""
        ssItem.Col = 11: ssItem.Text = adoSet.Fields("GeomchC2").Value & ""
        ssItem.Col = 12: ssItem.Text = adoSet.Fields("ChUnit").Value & ""
        ssItem.Col = 13: ssItem.Text = adoSet.Fields("GeomJan1").Value & ""
        ssItem.Col = 14: ssItem.Text = adoSet.Fields("GeomJan2").Value & ""
        ssItem.Col = 15: ssItem.Text = adoSet.Fields("GeomJan3").Value & ""
        ssItem.Col = 16: ssItem.Text = adoSet.Fields("Danwi").Value & ""
        ssItem.Col = 17: ssItem.Text = adoSet.Fields("ResultW").Value & ""
        
        ssItem.Col = 18: ssItem.Text = adoSet.Fields("DiffCount").Value & ""
        ssItem.Col = 19: ssItem.Text = adoSet.Fields("maxdiffc").Value & ""
        ssItem.Col = 20: ssItem.Text = adoSet.Fields("ChComment").Value & ""
        ssItem.Col = 21: ssItem.Text = adoSet.Fields("CgComment").Value & ""
        ssItem.Col = 22: ssItem.Text = adoSet.Fields("GeomsaGb").Value & ""
        ssItem.Col = 23: ssItem.Text = adoSet.Fields("GeomsaAb").Value & ""
        If Not IsNull(adoSet.Fields("Codate").Value) Then
            ssItem.Col = 24: ssItem.Text = adoSet.Fields("Codate").Value
        End If
        ssItem.Col = 25: ssItem.Text = adoSet.Fields("DeltaMin").Value & ""
        ssItem.Col = 26: ssItem.Text = adoSet.Fields("DeltaMax").Value & ""
        ssItem.Col = 27: ssItem.Text = adoSet.Fields("DeltaQC").Value & ""
        ssItem.Col = 28: ssItem.Text = adoSet.Fields("PanicMin").Value & ""
        ssItem.Col = 29: ssItem.Text = adoSet.Fields("PanicMax").Value & ""
        ssItem.Col = 30: ssItem.Text = adoSet.Fields("OrderCD").Value & ""
        ssItem.Col = 31: ssItem.Text = adoSet.Fields("GbRoutine").Value & ""
        ssItem.Col = 32: ssItem.Text = adoSet.Fields("GbCheck").Value & ""
        ssItem.Col = 33: ssItem.Text = adoSet.Fields("GbInput").Value & ""
        ssItem.Col = 34: ssItem.Text = adoSet.Fields("BarText").Value & ""
        pgbSelect.Value = pgbSelect.Value + 1
        panelPg.Caption = " " & pgbSelect.Value & "/" & adoSet.RecordCount
        adoSet.MoveNext
    Loop
    ssItem.ReDraw = True
    'panelPg.Visible = False
    Call adoSetClose(adoSet)
    
    ssItem.Row = 1
    ssItem.Col = 16
    ssItem.Action = SS_ACTION_GOTO_CELL
    Return
    
End Sub

Private Sub cmdRset_Click()
    cmbGeomsaW1.ListIndex = -1
    cmbGeomsaW2.ListIndex = -1
End Sub

Private Sub cmdSelectName_Click()
    
    frmSelectName.Show vbModal
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub Form_Load()
    
    pgbSelect.Value = 0
    
    GoSub Get_SampleBottle_Data
    Exit Sub
    
Get_SampleBottle_Data:
    '검체용기 코드 Select
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '88'"
    strSql = strSql & " ORDER  BY Codeky"
    
    cmbBottle.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbBottle.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                          Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    cmbBottle.AddItem ""
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiMain.stbMain.Panels(1).Text = ""
End Sub

Private Sub mnuPr_Click()
    
    If Trim(txtSlipno.Text) = "" Then
        MsgBox "SLip을 먼저 선택하세요!........."
        Exit Sub
    End If
    
    
    frmPrItem.Show vbModal
    
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub ssItem_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sItemCD     As String
    Dim sUsa        As String
    Dim sKor        As String
    
    
    If Col = 2 Then
        ssItem.Row = Row
        ssItem.Col = 1: sItemCD = Trim(ssItem.Text)
        sUsa = "": sKor = ""
        GoSub Get_ItemName
        mdiMain.stbMain.Panels(1).Text = "검사명(한) : " & sKor
    End If
    Exit Sub
    
Get_ItemName:
    strSql = ""
    strSql = strSql & " SELECT Itemnm, Itemko"
    strSql = strSql & " FROM   TWEXAM_ITEMML"
    strSql = strSql & " WHERE  CODEKY = '" & sItemCD & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    sUsa = Trim(adoSet.Fields("Itemnm").Value & "")
    sKor = Trim(adoSet.Fields("ItemKo").Value & "")
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub ssItem_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then Exit Sub
    
    ssItem.Row = Row
    ssItem.Col = 1
    If Trim(ssItem.Text) = "" Then Exit Sub
    If Left(ssItem.Text, 1) = "*" Then
        MsgBox "이미 삭제된 Code 입니다", vbInformation
        Exit Sub
    End If
    
    txtItemCode.Text = Mid(ssItem.Text, 3, Len(ssItem.Text) - 2)
    
    Call txtItemCode_KeyDown(vbKeyReturn, 0)
    
End Sub

Private Sub txtCgcmt_GotFocus()
    txtCgcmt.SelStart = 0
    txtCgcmt.SelLength = Len(txtCgcmt.Text)
    
End Sub

Private Sub txtChcmt_GotFocus()
    txtChcmt.SelStart = 0
    txtChcmt.SelLength = Len(txtChcmt.Text)
    
End Sub

Private Sub txtChunit_GotFocus()
    
    txtChunit.Alignment = 0
    txtChunit.SelStart = 0
    txtChunit.SelLength = Len(txtChunit.Text)
        

End Sub

Private Sub txtChunit_LostFocus()
    txtChunit.Text = Format(txtChunit.Text, "###0.0")
    txtChunit.Alignment = 1
End Sub

Private Sub txtChwhYg_GotFocus()
    txtChwhYg.SelStart = 0
    txtChwhYg.SelLength = Len(txtChwhYg.Text)

End Sub


Private Sub txtDanwi_GotFocus()
    txtDanwi.SelStart = 0
    txtDanwi.SelLength = Len(txtDanwi.Text)
    
    
End Sub

Private Sub txtDeltaMax_GotFocus()
    
    txtDeltaMax.Alignment = 0
    txtDeltaMax.SelStart = 0
    txtDeltaMax.SelLength = Len(txtDeltaMax.Text)
    
End Sub

Private Sub txtDeltaMax_LostFocus()
    
    txtDeltaMax.Alignment = 1
    
End Sub

Private Sub txtDeltaMin_GotFocus()
    
    txtDeltaMin.Alignment = 0
    txtDeltaMin.SelStart = 0
    txtDeltaMin.SelLength = Len(txtDeltaMin.Text)
    
End Sub

Private Sub txtDeltaMin_LostFocus()
    txtDeltaMin.Text = Format(txtDeltaMin.Text, "######0.00")
    txtDeltaMin.Alignment = 1
    
End Sub


Private Sub txtDiffEnd_GotFocus()
    
    txtDiffEnd.Alignment = 0
    txtDiffEnd.SelStart = 0
    txtDiffEnd.SelLength = Len(txtDiffEnd.Text)

End Sub

Private Sub txtDiffEnd_LostFocus()
    
    txtDiffEnd.Alignment = 1

End Sub

Private Sub txtDiffStart_GotFocus()
    
    txtDiffStart.Alignment = 0
    txtDiffStart.SelStart = 0
    txtDiffStart.SelLength = Len(txtDiffStart.Text)
    
End Sub

Private Sub txtDiffStart_LostFocus()

    txtDiffStart.Alignment = 1

End Sub

Private Sub txtGeomchc1_GotFocus()
    
    txtGeomchc1.SelStart = 0
    txtGeomchc1.SelLength = Len(txtGeomchc1.Text)
    
    If Trim(txtGeomchc1.Text) <> "" Then
        GoSub Get_SampleName
    End If
    Exit Sub
    

Get_SampleName:
    mdiMain.stbMain.Panels(1).Text = ""
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_Sample"
    strSql = strSql & " WHERE  CODE  =  '" & Trim(txtGeomchc1.Text) & "'"
        
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    mdiMain.stbMain.Panels(1).Text = Trim(adoSet.Fields("Codenm").Value & "")
    Call adoSetClose(adoSet)
    
    Return
    
    
    
End Sub

Private Sub txtGeomchc1_LostFocus()
    
    mdiMain.stbMain.Panels(1).Text = ""
    
End Sub

Private Sub txtGeomchc2_GotFocus()
    
    txtGeomchc2.SelStart = 0
    txtGeomchc2.SelLength = Len(txtGeomchc2.Text)

End Sub

Private Sub txtGeomjan1_GotFocus()
    
    txtGeomjan1.SelLength = Len(txtGeomjan1.Text)
    txtGeomjan1.SelStart = 0
    

End Sub

Private Sub txtGeomjan2_GotFocus()
    
    txtGeomjan2.SelStart = 0
    txtGeomjan2.SelLength = Len(txtGeomjan2.Text)
    
    
End Sub

Private Sub txtGeomjan3_GotFocus()
    
    txtGeomjan3.SelStart = 0
    txtGeomjan3.SelLength = Len(txtGeomjan3.Text)

End Sub

Private Sub txtGeomsaAb_GotFocus()
    
    txtGeomsaAb.SelStart = 0
    txtGeomsaAb.SelLength = Len(txtGeomsaAb.Text)
    
End Sub

Private Sub txtGeomsaTm_GotFocus()
    
    txtGeomsaTm.SelStart = 0
    txtGeomsaTm.SelLength = Len(txtGeomsaTm.Text)
    
End Sub


Private Sub txtItemCode_GotFocus()
    
    txtItemCode.SelStart = 0
    txtItemCode.SelLength = Len(txtItemCode.Text)
    
End Sub

Public Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        mdiMain.stbMain.Panels(1).Text = ""
        GoSub Check_Slip_Gubun
        GoSub Get_ItemCodeData
    End If
    Exit Sub
    
'/____________________________________________

Check_Slip_Gubun:
    If Trim(txtSlipno.Text) = "" Then
        MsgBox "처리할 SlipNumber 가 없습니다!.."
        frmSlipQry.Show vbModal
    End If
    Return
    
    
Get_ItemCodeData:
    Dim siTemCode       As String
    
    siTemCode = Trim(txtSlipno.Text) & Trim(txtItemCode.Text)
    
    strSql = ""
    strSql = strSql & " SELECT a.*, "
    strSql = strSql & "        TO_CHAR(a.Codate,'YYYY-MM-DD') Codate"
    strSql = strSql & " FROM   TWEXAM_ITEMML a"
    strSql = strSql & " WHERE  Codeky = '" & Trim(siTemCode) & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtItemNM.Text = RTrim(adoSet.Fields("itemNm").Value & "")
    txtItemKO.Text = Trim(adoSet.Fields("itemKo").Value & "")
    txtYageo.Text = Trim(adoSet.Fields("Yageo").Value & "")
    txtSugacd.Text = Trim(adoSet.Fields("Sugacd").Value & "")
    txtGeomsaTm.Text = Trim(adoSet.Fields("GeomsaTm").Value & "")
    
    Select Case Trim(adoSet.Fields("Geomsaw1").Value & "")
        Case "월": cmbGeomsaW1.ListIndex = 0
        Case "화": cmbGeomsaW1.ListIndex = 1
        Case "수": cmbGeomsaW1.ListIndex = 2
        Case "목": cmbGeomsaW1.ListIndex = 3
        Case "금": cmbGeomsaW1.ListIndex = 4
        Case "토": cmbGeomsaW1.ListIndex = 5
        Case "일": cmbGeomsaW1.ListIndex = 6
        Case Else: cmbGeomsaW1.ListIndex = -1
    End Select
    Select Case Trim(adoSet.Fields("Geomsaw2").Value & "")
        Case "월": cmbGeomsaW2.ListIndex = 0
        Case "화": cmbGeomsaW2.ListIndex = 1
        Case "수": cmbGeomsaW2.ListIndex = 2
        Case "목": cmbGeomsaW2.ListIndex = 3
        Case "금": cmbGeomsaW2.ListIndex = 4
        Case "토": cmbGeomsaW2.ListIndex = 5
        Case "일": cmbGeomsaW2.ListIndex = 6
        Case Else: cmbGeomsaW2.ListIndex = -1
    End Select
    
    'txtGeomsaW1.Text = Trim(adoSet.Fields("GeomsaW1").Value & "")
    'txtGeomsaW2.Text = Trim(adoSet.Fields("GeomsaW2").Value & "")
    
    
    
    txtChwhYg.Text = Trim(adoSet.Fields("ChwhYg").Value & "")
    If Trim(txtChwhYg.Text) <> "" Then
        Call SetComboBox(cmbBottle, Trim(txtChwhYg.Text), 4)
    Else
        cmbBottle.ListIndex = -1
    End If
        
    txtGeomchc1.Text = Trim(adoSet.Fields("Geomchc1").Value & "")
    txtGeomchc2.Text = Trim(adoSet.Fields("Geomchc2").Value & "")
    txtChunit.Text = Trim(adoSet.Fields("Chunit").Value & "")
    txtGeomjan1.Text = Trim(adoSet.Fields("Geomjan1").Value & "")
    txtGeomjan2.Text = Trim(adoSet.Fields("Geomjan2").Value & "")
    txtGeomjan3.Text = Trim(adoSet.Fields("Geomjan3").Value & "")
    txtDanwi.Text = Trim(adoSet.Fields("Danwi").Value & "")
    
    Select Case Trim(adoSet.Fields("Resultw").Value & "")
        Case "C": cmbResultW.ListIndex = 0
        Case "N": cmbResultW.ListIndex = 1
        Case "D": cmbResultW.ListIndex = 2
        Case "B": cmbResultW.ListIndex = 3
        Case Else: cmbResultW.ListIndex = -1
    End Select
    
    txtDiffStart.Text = Trim(adoSet.Fields("DiffCount").Value & "")
    txtDiffEnd.Text = Trim(adoSet.Fields("MaxDiffc").Value & "")
    txtChcmt.Text = Trim(adoSet.Fields("ChComment").Value & "")
    txtCgcmt.Text = Trim(adoSet.Fields("CgComment").Value & "")
    
    Select Case Trim(adoSet.Fields("GeomsaGb").Value & "")
        Case "J": cmbGeomsaGb.ListIndex = 0
        Case "W": cmbGeomsaGb.ListIndex = 1
        Case Else: cmbGeomsaGb.ListIndex = -1
    End Select
    
    Call SetComboBox(cmbExGB, adoSet.Fields("ExGb").Value & "", 1)
    
    txtGeomsaAb.Text = Trim(adoSet.Fields("GeomsaAb").Value & "")
    
    If Not IsNull(adoSet.Fields("Codate").Value) Then
        dtCodate.Value = adoSet.Fields("Codate").Value
    Else
        dtCodate.Value = Format(Now, "yyyy-MM-dd")
    End If
    
    txtDeltaMin.Text = Trim(adoSet.Fields("DeltaMin").Value & "")
    txtDeltaMax.Text = Trim(adoSet.Fields("DeltaMax").Value & "")
    
    Select Case adoSet.Fields("DeltaQc").Value & ""
        Case "1": cmbDeltaQc.ListIndex = 0
        Case "2": cmbDeltaQc.ListIndex = 1
        Case "3": cmbDeltaQc.ListIndex = 2
        Case "4": cmbDeltaQc.ListIndex = 3
        Case Else: cmbDeltaQc.ListIndex = -1
    End Select
    
    'txtDeltaQc.Text = Trim(adoSet.Fields("DeltaQc").Value & "")
    txtPanicmin.Text = Trim(adoSet.Fields("PanicMin").Value & "")
    txtPanicmax.Text = Trim(adoSet.Fields("PanicMax").Value & "")
    txtOrderCode.Text = Trim(adoSet.Fields("OrderCD").Value & "")
    
    Select Case Trim(adoSet.Fields("gbRoutine").Value & "")
        Case "O": cmbRoutine.ListIndex = 0
        Case "I": cmbRoutine.ListIndex = 1
        Case Else: cmbRoutine.ListIndex = -1
    End Select
    
    Select Case Trim(adoSet.Fields("GbCheck").Value & "")
        Case "O": cmbCheck.ListIndex = 0
        Case "I": cmbCheck.ListIndex = 1
        Case Else: cmbCheck.ListIndex = -1
    End Select
        
    Select Case Trim(adoSet.Fields("GbInput").Value & "")
        Case "O": cmbInput.ListIndex = 0
        Case "I": cmbInput.ListIndex = 1
        Case Else: cmbInput.ListIndex = -1
    End Select
    
    txtBarText.Text = Format(adoSet.Fields("BarText").Value & "")
    
    If adoSet.Fields("BarGb").Value & "" = "1" Then
        chkBarGb.Value = "1"
    Else
        chkBarGb.Value = "0"
    End If
    
    Select Case adoSet.Fields("ExGb").Value & ""
        Case "1": cmbExGB.ListIndex = 0        'SCL
        Case "2": cmbExGB.ListIndex = 1        '녹십자
        Case Else: cmbExGB.ListIndex = -1
    End Select
    
    txtCodeGb.Text = adoSet.Fields("Codegu").Value & ""
    
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub txtItemKO_GotFocus()
    txtItemKO.SelStart = 0
    txtItemKO.SelLength = Len(txtItemKO.Text)

End Sub

Private Sub txtItemNM_GotFocus()
    txtItemNM.SelStart = 0
    txtItemNM.SelLength = Len(txtItemNM.Text)
    
End Sub


Private Sub txtOrderCode_GotFocus()
    txtOrderCode.SelStart = 0
    txtOrderCode.SelLength = Len(txtOrderCode.Text)
    
End Sub

Private Sub txtPanicmax_GotFocus()
    
    txtPanicmax.Alignment = 0
    txtPanicmax.SelStart = 0
    txtPanicmax.SelLength = Len(txtPanicmax.Text)

End Sub

Private Sub txtPanicmax_LostFocus()
    txtPanicmax.Text = Format(txtPanicmax.Text, "######0.00")
    txtPanicmax.Alignment = 1
    
End Sub

Private Sub txtPanicmin_GotFocus()
    
    txtPanicmin.Alignment = 0
    txtPanicmin.SelStart = 0
    txtPanicmin.SelLength = Len(txtPanicmin.Text)
    
End Sub

Private Sub txtPanicmin_LostFocus()
    
    txtPanicmin.Text = Format(txtPanicmin.Text, "######0.00")
    txtPanicmin.Alignment = 1
    
End Sub

Private Sub txtSlipNo_GotFocus()
    txtSlipno.SelStart = 0
    txtSlipno.SelLength = Len(txtSlipno.Text)
    
End Sub

Private Sub txtSlipNo_LostFocus()
    
    mdiMain.stbMain.Panels(1).Text = ""
    txtSlipno.Tag = txtSlipno.Text
    txtSLipName.Tag = txtSLipName.Text
    Call ClearForm(Me)
    txtSlipno.Text = txtSlipno.Tag
    txtSLipName.Text = txtSLipName.Tag
    
    GoSub Get_SLIPName_Data
    
    Exit Sub
    
    
Get_SLIPName_Data:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky = '" & Trim(txtSlipno.Text) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        Return
    End If
    
    txtSLipName.Text = adoSet.Fields("Codenm").Value & ""
    Call adoSetClose(adoSet)
    txtItemCode.SetFocus
    
    Return
    
    
End Sub

Private Sub txtSugacd_GotFocus()
    txtSugacd.SelStart = 0
    txtSugacd.SelLength = Len(txtSugacd.Text)
    
End Sub

Private Sub txtYageo_GotFocus()
    txtYageo.SelStart = 0
    txtYageo.SelLength = Len(txtYageo.Text)
    
End Sub
