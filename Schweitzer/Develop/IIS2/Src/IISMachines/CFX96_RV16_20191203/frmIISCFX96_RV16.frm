VERSION 5.00
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D74ED2A2-3650-4720-93BC-FDDD8DCBC769}#1.0#0"; "Han2EngOCX.ocx"
Begin VB.Form frmIISCFX96_RV16 
   BackColor       =   &H00DBE6E6&
   Caption         =   "CFX96"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtBarNo 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   2340
      TabIndex        =   49
      Top             =   150
      Width           =   1455
   End
   Begin VB.OptionButton optTest 
      BackColor       =   &H00DBE6E6&
      Caption         =   "RB5"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   5550
      TabIndex        =   47
      Top             =   210
      Width           =   795
   End
   Begin VB.OptionButton optTest 
      BackColor       =   &H00DBE6E6&
      Caption         =   "RV16"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   46
      Top             =   210
      Value           =   -1  'True
      Width           =   915
   End
   Begin HAN2ENGOCXLib.Han2EngOCX Han2Eng 
      Height          =   315
      Left            =   15900
      TabIndex        =   45
      Top             =   5370
      Visible         =   0   'False
      Width           =   405
      _Version        =   65536
      _ExtentX        =   714
      _ExtentY        =   556
      _StockProps     =   0
   End
   Begin VB.FileListBox FileCFX96 
      Height          =   870
      Left            =   15780
      Pattern         =   "*.csv"
      TabIndex        =   39
      Top             =   4320
      Visible         =   0   'False
      Width           =   2805
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3270
      Left            =   105
      TabIndex        =   22
      Top             =   915
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   5768
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
      MaxCols         =   6
      MaxRows         =   10
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":0000
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "결과지출력"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17490
      Style           =   1  '그래픽
      TabIndex        =   37
      Top             =   6090
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker dtpFrDate 
      Height          =   315
      Left            =   1140
      TabIndex        =   35
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   36372481
      CurrentDate     =   40270
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조회"
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
      Left            =   3900
      TabIndex        =   34
      Top             =   540
      Width           =   855
   End
   Begin VB.CommandButton cmdGetRslt 
      BackColor       =   &H00DBE6E6&
      Caption         =   "결과 얻기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10290
      Style           =   1  '그래픽
      TabIndex        =   33
      Top             =   8567
      Width           =   1185
   End
   Begin VB.Timer tmrResult 
      Left            =   6600
      Top             =   8520
   End
   Begin VB.TextBox txtFileNm 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   16260
      TabIndex        =   31
      Top             =   1560
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSComDlg.CommonDialog AlloFile 
      Left            =   7080
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtWorkNo 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      MaxLength       =   9
      TabIndex        =   29
      Top             =   540
      Width           =   615
   End
   Begin VB.CommandButton cmdMakeWS 
      BackColor       =   &H00DBE6E6&
      Caption         =   "리스트 생성"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16290
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   6090
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   107
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   661
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
      Caption         =   "■ 환자정보"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
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
      Height          =   495
      Left            =   12698
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11483
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin VB.TextBox txtBarNo1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8940
      TabIndex        =   0
      Text            =   "123456789011"
      Top             =   8190
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13913
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8567
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1290
      Left            =   6548
      TabIndex        =   5
      Top             =   407
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   6
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "00000001"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   7
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "이상대"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStatFg 
         Height          =   315
         Left            =   6795
         TabIndex        =   8
         Top             =   165
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "응급"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   315
         Left            =   1245
         TabIndex        =   9
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "이상대 아기"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   10
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "수술실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   6795
         TabIndex        =   11
         Top             =   525
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "Blood"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   1245
         TabIndex        =   12
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "남자 / 29"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   13
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   "65병동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPnlNm 
         Height          =   315
         Left            =   6795
         TabIndex        =   41
         Top             =   900
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   12648447
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "PANEL :"
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
         Index           =   5
         Left            =   5760
         TabIndex        =   42
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
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
         TabIndex        =   21
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급여부 :"
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
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "병  동 : "
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
         TabIndex        =   19
         Top             =   975
         Width           =   810
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "진료과 :"
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
         TabIndex        =   18
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방의 :"
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
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
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
         TabIndex        =   16
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "이     름 :"
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
         TabIndex        =   15
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "환  자 ID :"
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
         TabIndex        =   14
         Top             =   240
         Width           =   990
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   7665
      Top             =   8430
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4110
      Left            =   105
      TabIndex        =   23
      Top             =   4725
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   7250
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
      MaxCols         =   14
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":056D
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   105
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   105
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Caption         =   "■ 검사대상 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   405
      Left            =   105
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4305
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "■ 검사완료 리스트"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6690
      Left            =   6555
      TabIndex        =   26
      Top             =   1710
      Width           =   8580
      _Version        =   393216
      _ExtentX        =   15134
      _ExtentY        =   11800
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
      MaxCols         =   8
      MaxRows         =   22
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":0E13
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   315
      Left            =   2520
      TabIndex        =   36
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   36372481
      CurrentDate     =   40270
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   18870
      TabIndex        =   38
      Top             =   4410
      Visible         =   0   'False
      Width           =   675
      _Version        =   393216
      _ExtentX        =   1191
      _ExtentY        =   1191
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":1506
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   8310
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread vasINPrint 
      Height          =   2685
      Left            =   7110
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   3915
      _Version        =   393216
      _ExtentX        =   6906
      _ExtentY        =   4736
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   76
      Protect         =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":174F
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vasFDPrint 
      Height          =   2685
      Left            =   11070
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   3915
      _Version        =   393216
      _ExtentX        =   6906
      _ExtentY        =   4736
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   76
      Protect         =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":C5BA
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vasExcel 
      Height          =   2055
      Left            =   15600
      TabIndex        =   40
      Top             =   2190
      Visible         =   0   'False
      Width           =   4875
      _Version        =   393216
      _ExtentX        =   8599
      _ExtentY        =   3625
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmIISCFX96_RV16.frx":17A01
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "▶ 접수일자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   50
      Top             =   630
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사명"
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
      Index           =   1
      Left            =   3840
      TabIndex        =   48
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "파일명 : "
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
      Left            =   15390
      TabIndex        =   32
      Top             =   1635
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Start WN"
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
      Index           =   0
      Left            =   4800
      TabIndex        =   30
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BN"
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
      Left            =   1980
      TabIndex        =   27
      Top             =   210
      Width           =   285
   End
End
Attribute VB_Name = "frmIISCFX96_RV16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISCFX96_RV16.frm
'   작성자  : 오세원
'   내  용  : CFX96_RV16 장비폼
'   작성일  : 2015-08-25
'   버  전  : 1.0.0
'   병  원  :
'       1. 전주예수병원
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady의 Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
End Enum

'## tblComplete의 Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccSexAge = 6
    ccDoctNm = 7:       ccDeptNm = 8
    ccWardNm = 9:       ccStatFg = 10
    ccSpcNm = 11:       ccQcFg = 12
    ccSendCnt = 13:     ccResult = 14
End Enum

'## tblResult의 Column Enum
Private Enum TResultEnum
    ccTestNm = 1
    ccEqpResult = 2
    ccLISResult = 3
    ccUnit = 4
    ccHLDiv = 5
    ccDPDiv = 6
    ccRef = 7
    ccClass = 8
    '-- 2015.08.28 추가
    ccIntBase = 9
End Enum

'## Clear Enum
Private Enum ClearEnum
    ccAll = 1
    ccLabel = 2
End Enum

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '인터페이스 클래스
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '팝업메뉴
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors           '인터페이스 에러 컬렉션
'Private mOrder      As clsIISIntOrder           '오더정보 클래스

Private mEqpCd  As String   '장비코드
Private mEqpKey As String   '장비키

'임시 - 데모용
Dim strTransData    As Variant

Dim AdoCn           As ADODB.Connection
Dim AdoRS           As ADODB.Recordset

Private Const AM As Variant = -0.696669715055768
Private Const AN As Variant = 3.57268581287044
Private Const BM As Variant = -0.58804147563697
Private Const BN As Variant = 3.68672614195763
Private Const CM As Variant = -0.424608277908711
Private Const CN As Variant = 3.85830192881032

Private Const VarA As Variant = 0
Private Const VarD As Variant = 150



Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Private Sub OpenExcel()

    Dim strFile As String
    Dim i, iCnt As Integer
    Dim strTemp As String
    Dim varTmp  As Variant
    Dim xlApp As New Excel.Application
    Dim xlSheet As Excel.Worksheet
    Dim strPath As String
    Dim strDestFile As String
    
'MsgBox "11"

    AlloFile.DialogTitle = "엑셀파일 열기"
    AlloFile.InitDir = GetCFX96Config("ExportPath")
    AlloFile.ShowOpen
    
    If Len(AlloFile.FileName) > 0 Then
        xlApp.Workbooks.Open AlloFile.FileName
        strPath = AlloFile.FileName
    Else
        Exit Sub
    End If
    
'MsgBox "12"
    
    Set xlSheet = xlApp.Worksheets("Real-Q MTB & NTM Kit")

'MsgBox "13"
    
    With vasExcel
        .Action = ActionClear
        For iCnt = 5 To .MaxRows
            For i = 1 To 10
                'If xlSheet.Cells(iCnt, i) <> "" Then
                If Trim(xlSheet.Cells(iCnt, 1)) = "" Then
                    'xlApp.Workbooks.Close
                    'xlApp.Quit

                    'Set xlSheet = Nothing
                    GoTo RST
                End If
                
                vasExcel.SetText i, iCnt - 4, Trim(xlSheet.Cells(iCnt, i))
            Next
        Next iCnt
    End With

'MsgBox "14"

RST:
   ' xlApp.Workbooks(strPath).Close
   ' xlApp.Quit
    
    Set xlSheet = Nothing
    
''    '대상 파일 이름을 정의
''    strDestFile = App.path & "\Log\" & Format(Now, "yyyymmdd-hhmm")
''    '원본을 대상에 복사
''    FileCopy strPath, strDestFile
''
''    Kill strSrcfile
    'FileCFX96.Refresh

End Sub
    
Private Sub OpenExcel_New_RD5()

    Dim y As Boolean, z As Boolean
    Dim Var As Variant
    Dim x As Integer, listcount As Integer, handle As Integer
    Dim List(10) As String

    '- 엑셀 파일 경로 찾기
    AlloFile.Filter = "엑셀파일|*.xls"  ' '파일 이름
    AlloFile.InitDir = GetCFX96Config("ExportPath")             '파일 경로
    AlloFile.Action = 1                       '나도 솔직히 이거 잘모르겠다 꼭1아니여도 됨 0은 안됨?

     '- 파일경로로 spread에 넣기 이건 퍼옴
    x = vasExcel.IsExcelFile(AlloFile.FileName) '파일이 있으면 x는 1이됨
    
    '0 = 1번째 tab
    '1 = 2번째 tab
    
    '-- RV16
'    If optTest(0).Value = True Then
'        listcount = 1
'    '-- RB5
'    Else
'        listcount = 0
'    End If
    
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = vasExcel.GetExcelSheetList(AlloFile.FileName, List, listcount, App.Path & "\tmp.txt", handle, True) '엑셀시트를 잘불러오면y는 뜨루
        If y = True Then
            'MsgBox "Got sheet list.", , "Status"
            
            z = vasExcel.ImportExcelSheet(handle, 0)     'z가 임뽈뜨 잘되면 뜨루?
        
            If z = True Then
                'vasExcel.Visible = True
                'vasExcel.ZOrder 0
            Else
                MsgBox "Import did not succeed.", , "Result"
            End If
        Else
            MsgBox "Cannot return information for Excel file.", , "Result"
        End If
    Else
        MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
    End If



'    ' Check if file is an Excel file and set result to x
'    x = fpSpread1.IsExcelFile("C:\Samples\Files\quotes.XLS")
'
'    ' If file is Excel file, tell user, import sheet
'    ' list, and set result to y
'    If x = 1 Then
'        MsgBox "File is an Excel file.", , "File Type"
'        y = fpSpread1.GetExcelSheetList("C:\Samples\Files\quotes.XLS", List, listcount, "C:\Samples\Files\ILOGFILE.TXT", handle, True)
'        ' If received sheet list, tell user, import file,
'        ' and set result to z
'        If y = True Then
'            MsgBox "Got sheet list.", , "Status"
'            z = fpSpread1.ImportExcelSheet(handle, 0)
'            ' Tell user result based on T/F value of z
'            If z = True Then
'                MsgBox "Import complete.", , "Result"
'            Else
'                MsgBox "Import did not succeed.", , "Result"
'            End If
'        Else
'            ' Tell user cannot obtain sheet list
'            MsgBox "Cannot return information for Excel file.", , "Result"
'        End If
'    Else
'        ' Tell user file is not Excel file or is locked
'        MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
'    End If
    

'    Dim x As Boolean
'    ' Export Excel file and set result to x
'    x = fpSpread1.ExportToExcel("C:\Samples\FILE.XLS", "Test Sheet 1", "C:\Samples\LOGFILE.TXT")
'    ' Display result to user based on T/F value of x
'    If x = True Then
'        MsgBox "Export complete.", , "Result"
'    Else
'        MsgBox "Export did not succeed.", , "Result"
'    End If

End Sub

    
Private Sub OpenExcel_New_RV16()

    Dim y As Boolean, z As Boolean
    Dim Var As Variant
    Dim x As Integer, listcount As Integer, handle As Integer
    Dim List(10) As String

    '- 엑셀 파일 경로 찾기
    AlloFile.Filter = "엑셀파일|*.xls"  ' '파일 이름
    AlloFile.InitDir = GetCFX96Config("ExportPath")             '파일 경로
    AlloFile.Action = 1                       '나도 솔직히 이거 잘모르겠다 꼭1아니여도 됨 0은 안됨?

     '- 파일경로로 spread에 넣기 이건 퍼옴
    x = vasExcel.IsExcelFile(AlloFile.FileName) '파일이 있으면 x는 1이됨
    
    '0 = 1번째 tab
    '1 = 2번째 tab
    
    '-- RV16
'    If optTest(0).Value = True Then
        listcount = 1
'    '-- RB5
'    Else
'        listcount = 0
'    End If
    
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = vasExcel.GetExcelSheetList(AlloFile.FileName, List, listcount, App.Path & "\tmp.txt", handle, True) '엑셀시트를 잘불러오면y는 뜨루
        If y = True Then
            'MsgBox "Got sheet list.", , "Status"
            
            z = vasExcel.ImportExcelSheet(handle, 1)     'z가 임뽈뜨 잘되면 뜨루?
        
            If z = True Then
                'vasExcel.Visible = True
                'vasExcel.ZOrder 0
            Else
                MsgBox "Import did not succeed.", , "Result"
            End If
        Else
            MsgBox "Cannot return information for Excel file.", , "Result"
        End If
    Else
        MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
    End If



End Sub

'엑셀 파일을 그리드에 넣기
Private Sub Excel_Open()
    Dim xlApp   As New Excel.Application
    Dim XLappWS As Worksheet
    Dim lngSCnt As Long
    Dim lngSColCnt(100) As Long
    Dim dummy       As String
    Dim strRowData  As Variant
    Dim lngRowCnt   As Long
    Dim chk_str     As String
    Dim dummy_max   As Long
    Dim lngTotColCnt   As Long
    Dim lngTotRowCnt   As Long
    Dim i, j, k     As Long



    lngTotColCnt = 0
    lngTotRowCnt = 0
    
    
    '엑셀 열기
    AlloFile.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls"
    AlloFile.Action = 1
    
    
    If AlloFile.FileTitle = "" Then
        Exit Sub
    End If
    
    xlApp.Workbooks.Open (Trim(AlloFile.FileName))
    
    lngSCnt = xlApp.Worksheets.Count
    
    '-- 전체 워크시트 불러오기와서 '임시.txt' 파일로 저장
    For i = 1 To lngSCnt
        Set XLappWS = xlApp.Worksheets(i)
        XLappWS.Activate
        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
        xlApp.DisplayAlerts = False
    
        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
'        xlApp.ActiveWorkbook.SaveAs "C:\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
        xlApp.ActiveWorkbook.SaveAs "C:\" & Trim(i) + 10 & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
        
        
        'XLappWS.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>엑셀 2000용
        'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>엑셀 2003용
    Next i
    
    xlApp.Quit
    Set XLappWS = Nothing
    Set xlApp = Nothing
    
    '-- 전체 엑셀의 MAX cols값 추출
    dummy_max = 0
    For i = 1 To lngSCnt
        If lngSColCnt(i) >= dummy_max Then
            dummy_max = lngSColCnt(i)
        End If
    Next i
    lngTotColCnt = dummy_max
    
    '전체 row값 추출
    For i = 1 To lngSCnt
'''        Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(Trim(dummy)) > 0 Then
                lngTotRowCnt = lngTotRowCnt + 1
            End If
        Wend
        Close #1
    Next i
    
    '-- 그리드 초기화
    vasExcel.MaxRows = 0
    vasExcel.MaxRows = lngTotRowCnt
    vasExcel.MaxCols = lngTotColCnt
    
    '-- 그리드에 출력
    For i = 1 To lngSCnt
        '''Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(chk_str) > 0 Then
                lngRowCnt = lngRowCnt + 1
                For j = 0 To UBound(strRowData)
                    Call vasExcel.SetText(j + 1, lngRowCnt, CStr(strRowData(j)))
                Next j
            End If
        Wend
        Close #1
    Next i

'    Call SpreadSheetSort(vasExcel, 6, 2)
'    With vasExcel
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 2: .Row2 = .DataRowCnt
'        .SortBy = 0
'        .SortKey(1) = 6       '정렬키 열번호
'        .SortKey(2) = 2       '정렬키 열번호
'
'        .SortKeyOrder(1) = SortKeyOrderAscending
'        .SortKeyOrder(2) = SortKeyOrderAscending
'
'        .Action = ActionSort
'    End With


'Dim SortKeys, SortKeyOrder As Variant
'
'    SortKeys = Array(6, 2)
'    SortKeyOrder = Array(6, 2)
'    ' Sort data in first five columns and rows by column 1 and 3
'    vasExcel.Sort 6, 2, 2, vasExcel.MaxRows, SS_SORT_BY_ROW, SortKeys, SortKeyOrder

End Sub


Private Sub cmdGetRslt_Click()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strBuffer   As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim intCnt      As Integer
    Dim varTmp      As Variant

    Dim fName As String
    Dim Buf() As Byte
    Dim r As Long


    '-- RV16
    If optTest(0).Value = True Then
        Call OpenExcel_New_RV16
    
        strtmpBuf = ""
            
        Call EditRcvData
    '-- RB5
    Else
        Call OpenExcel_New_RD5
    
        strtmpBuf = ""
            
        Call EditRcvData_RB5
    End If
    
    Call mIntLib.ClearBuffer

Exit Sub

ErrRoutine:

    
End Sub


Public Function ReadFileBinary(ByVal strFileName As String) As String

On Error GoTo errHandler
    Dim fsT, tFilePath As String

    Set fsT = CreateObject("ADODB.Stream")

    fsT.Type = 2
 
    fsT.Charset = "utf-8"
 
    fsT.Open
'    fsT.Type = adTypeBinary
'    fsT.Type = adTypeText
    fsT.LoadFromFile strFileName
   
    Dim strText As String
    strText = ""
    Do Until fsT.EOS
        strText = strText & fsT.ReadText(adReadLine) & vbLf     ' 줄바꿈 추가
    Loop
 

    fsT.Close

    ReadFileBinary = strText
    GoTo finish
 
errHandler:

    MsgBox (Err.Description)
    Exit Function

finish:

End Function

Private Sub cmdMakeWS_Click()
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
'    Dim objResult   As clsIISResult     '결과내역 클래스
'    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
'    Dim mLogOn As clsIISLogOn

'    Dim strAlloFile As String
    Dim lngFIleNum  As Long
    Dim strInFo     As String
'    Dim strOldInFo  As String

    
    Dim iCnt As Integer
    Dim varTmp As Variant
    Dim strBarNo As String
    Dim strLabNo As String
    Dim strFNm   As String
    Dim strLNm   As String
    Dim strBirth As String
    Dim strSex   As String
    Dim strAge   As String
    
    Dim intCnt As Integer
    
    Screen.MousePointer = 11
    
    
    With AlloFile
        .CancelError = True
        .FileName = GetCFX96Config("ImportPath") & "\" & Trim(txtFileNm.Text) & ".csv"
        If Len(Dir(.FileName)) Then Kill .FileName

        lngFIleNum = FreeFile

        Open .FileName For Append As #lngFIleNum

'Col1;Col2;Col3;Col4;Col5;Col6;Col7
'BioSciTec / Panel / Rev. 001;Sample 1;123;Goethe;Wolfgang;Male;1759.11.10
'BioSciTec / Panel / Rev. 001;Sample 2;234;Schiller;Friedrich;Male;1749.08.28
'BioSciTec / Panel / Rev. 001;Sample 3;345;Kant;Immanuel;Male;1804.02.12


        
        'MEDIWISS / RIDA Panel 1 KO / Rev. 006
        'MEDIWISS / RIDA Panel 1 KO / Rev. 006
        
        'strInFo = "Col1;Col2;Col3;Col4;Col5;Col6;Col7" & vbCrLf
        
        vasExcel.Action = ActionClear
        Call vasExcel.SetText(1, 1, "RowIndexNumber")
        Call vasExcel.SetText(2, 1, "Test")
        Call vasExcel.SetText(3, 1, "Sample_ID")
        Call vasExcel.SetText(4, 1, "Patient_ID")
        Call vasExcel.SetText(5, 1, "Patient_Last_Name")
        Call vasExcel.SetText(6, 1, "Patient_First_Name")
        Call vasExcel.SetText(7, 1, "Patient_Title")
        Call vasExcel.SetText(8, 1, "Patient_Sex")
        Call vasExcel.SetText(9, 1, "Patient_Date_of_Birth")
        
        strInFo = "RowIndexNumber,Test,Sample_ID,Patient_ID,Patient_Last_Name,Patient_First_Name,Patient_Title,Patient_Sex,Patient_Date_of_Birth,Patient_Street,Patient_Zip_Code,Patient_City,Patient_Country,Patient_Phone,Patient_Fax,Patient_email,Company_Name,Company_Street,Company_Zip_Code,Company_City,Company_Trade_Register,Company_Country,Company_Legal_Form,Company_Tax_ID,Company_Phone,Company_Fax,Company_email,Sample_Date_Of_Receipt,Sample_Source,Sample_Type,Type_of_strip,Control,SubstanceFamily,Tray_ID,Well_no,Connected_with,Company_Website,Sample_Date_Of_Sampling,Custom_ID,Comments,Custom0,Custom1,Custom2,Custom3,Custom4,Custom5,Custom6,Custom7,Custom8,Custom9,Membrane,Patient_Nationality" & vbCr '& vbLf
'1,MEDIWISS / RIDA Panel 1 KO / Rev. 006,1,4,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
'2,MEDIWISS / RIDA Panel 2 KO / Rev. 004,2,5,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
'3,MEDIWISS / RIDA Panel 1 KO / Rev. 006,1,4,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
'4,MEDIWISS / RIDA Panel 2 KO / Rev. 004,2,5,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,

        intCnt = 1
        For iCnt = 1 To tblReady.DataRowCnt
            tblReady.GetText 2, iCnt, varTmp
            strBarNo = varTmp
            tblReady.GetText 4, iCnt, varTmp
            strLabNo = varTmp
            tblReady.GetText 5, iCnt, varTmp
            strFNm = Mid(varTmp, 1, 1)
            strLNm = Mid(varTmp, 2)
            strFNm = Han2Eng.HanToEng(strFNm)
            strLNm = Han2Eng.HanToEng(strLNm)
            tblReady.GetText 6, iCnt, varTmp
            varTmp = Split(varTmp, "|")
            strSex = varTmp(0)
            If strSex = "M" Then
                strSex = "Male"
            Else
                strSex = "FeMale"
            End If
            
            strAge = varTmp(1)
            If Len(strAge) = 13 Then
                If Mid(strAge, 7, 1) = "1" Or Mid(strAge, 7, 1) = "2" Then
                    strAge = "19" & Mid(strAge, 1, 6)
                ElseIf Mid(strAge, 7, 1) = "3" Or Mid(strAge, 7, 1) = "4" Then
                    strAge = "20" & Mid(strAge, 1, 6)
                Else
                    strAge = ""
                End If
            Else
                strAge = ""
            End If
            
            If strAge <> "" Then
                strAge = Mid(strAge, 1, 4) & "." & Mid(strAge, 5, 2) & "." & Mid(strAge, 7, 2)
            End If
            
            tblReady.GetText 1, iCnt, varTmp
            If varTmp = "IN" Then
                strInFo = strInFo & CStr(intCnt) & "," & "MEDIWISS / Panel 30 KO Inhalant A / Rev. 006," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr '& vbLf
                If iCnt = tblReady.DataRowCnt Then
                    strInFo = strInFo & CStr(intCnt + 1) & "," & "MEDIWISS / Panel 30 KO Inhalant B / Rev. 007," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge
                Else
                    strInFo = strInFo & CStr(intCnt + 1) & "," & "MEDIWISS / Panel 30 KO Inhalant B / Rev. 007," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr '& vbLf
                End If
            ElseIf varTmp = "FD" Then
                strInFo = strInFo & CStr(intCnt) & "," & "MEDIWISS / Panel 30 KO Food  A / Rev. 004," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr '& vbLf
                If iCnt = tblReady.DataRowCnt Then
                    strInFo = strInFo & CStr(intCnt + 1) & "," & "MEDIWISS / Panel 30 KO Food  B / Rev. 007," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge
                Else
                    strInFo = strInFo & CStr(intCnt + 1) & "," & "MEDIWISS / Panel 30 KO Food  B / Rev. 007," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr '& vbLf
                End If
            End If
            
'            Call vasExcel.SetText(1, intCnt, CStr(intCnt - 1))
'            Call vasExcel.SetText(1, intCnt + 1, CStr(intCnt))
'            If varTmp = "IN" Then
'                Call vasExcel.SetText(2, intCnt, "MEDIWISS / RIDA Panel 1 KO / Rev. 006")
'                Call vasExcel.SetText(2, intCnt + 1, "MEDIWISS / RIDA Panel 2 KO / Rev. 004")
'            ElseIf varTmp = "FD" Then
'                Call vasExcel.SetText(2, intCnt, "MEDIWISS / RIDA Panel 3 KO / Rev. 004")
'                Call vasExcel.SetText(2, intCnt + 1, "MEDIWISS / RIDA Panel 4 KO / Rev. 003")
'            End If
'
'            Call vasExcel.SetText(3, intCnt, strBarNo)
'            Call vasExcel.SetText(3, intCnt + 1, strBarNo)
'            Call vasExcel.SetText(4, intCnt, strLabNo)
'            Call vasExcel.SetText(4, intCnt + 1, strLabNo)
'
'            'Call vasExcel.SetText(5, iCnt + 1, "Patient_Last_Name")
'            'Call vasExcel.SetText(6, iCnt + 1, "Patient_First_Name")
'            'Call vasExcel.SetText(7, iCnt + 1, "Patient_Title")
'            If strSex = "M" Then
'                Call vasExcel.SetText(8, intCnt, "Male")
'                Call vasExcel.SetText(8, intCnt + 1, "Male")
'            Else
'                Call vasExcel.SetText(8, intCnt, "FeMale")
'                Call vasExcel.SetText(8, intCnt + 1, "FeMale")
'            End If
            intCnt = intCnt + 2
            
            'Call vasExcel.SetText(9, iCnt + 1, "Patient_Date_of_Birth")
            
            

'            tblReady.GetText 6, iCnt, varTmp
'            strInFo = varTmp
'            If strInFo = "1" Then
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 1KO v80 UK.TST" & ";"
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 2KO v80 UK.TST" & ";"
'            ElseIf strInFo = "3" Then
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 3KO v80 UK.TST" & ";"
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 4KO v80 UK.TST" & ";"
'            Else    '9
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 1KO v80 UK.TST" & ";"
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 2KO v80 UK.TST" & ";"
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 3KO v80 UK.TST" & ";"
'                Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & strBarNo & ";Panel 4KO v80 UK.TST" & ";"
'            End If
        Next
       ' Call vasExcel.ExportToExcel(.FileName, txtFileNm.Text, "c:\1.txt")
        Print #lngFIleNum, strInFo
        Close #lngFIleNum
    End With
    
    MsgBox "워크리스트 생성이 완료되었습니다", vbInformation + vbOKOnly, Me.Caption
    Screen.MousePointer = 0

End Sub

Private Function GetCFX96Config(ByVal strConfigNm As String) As String

Dim strFileName As String
Dim strReturnedString As String

    strFileName = App.Path & "\CFX96_RV16.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CFX96", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetCFX96Config = strReturnedString
    
End Function

Private Sub cmdPrint_Click()
    Dim strQcFg     As String   'QC유무
    Dim strResult   As String   'LIS 결과
    Dim strTemp     As String
    Dim i           As Long
    
    Dim strPrtData(100) As String
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intCnt      As Integer
    Dim strPanel    As String
    Dim strValue    As String
    Dim strClass    As String
    Dim iDestRow    As Integer
    
    Erase strPrtData
    intCnt = 0
    
    With vasINPrint
        Call .SetText(3, 2, ""): Call .SetText(7, 2, ""): Call .SetText(11, 2, "")
        Call .SetText(3, 3, ""): Call .SetText(7, 3, ""): Call .SetText(11, 3, "")
        Call .SetText(3, 4, ""): Call .SetText(7, 4, ""): Call .SetText(11, 4, "")
        For i = 6 To 65
            Call .SetText(4, i, ""): Call .SetText(7, i, ""): Call .SetText(10, i, ""): Call .SetText(12, i, "")
        Next
    End With
    
    With vasFDPrint
        Call .SetText(3, 2, ""): Call .SetText(7, 2, ""): Call .SetText(11, 2, "")
        Call .SetText(3, 3, ""): Call .SetText(7, 3, ""): Call .SetText(11, 3, "")
        Call .SetText(3, 4, ""): Call .SetText(7, 4, ""): Call .SetText(11, 4, "")
        For i = 6 To 65
            Call .SetText(4, i, ""): Call .SetText(7, i, ""): Call .SetText(10, i, ""): Call .SetText(12, i, "")
        Next
    End With
    
    
    If lblPtId.Caption <> "" And lblPnlNm.Caption <> "" Then
        strPrtData(0) = lblPtId.Caption
        strPrtData(1) = lblName.Caption
        strPrtData(2) = Format(Now, "yyyy-mm-dd")
        strPrtData(3) = mGetP(lblSexAge.Caption, 2, "/")
        'strPrtData(4) = IIf(mGetP(lblSexAge.Caption, 1, "/") = "M", "남자", "여자")
        strPrtData(4) = IIf(InStr(lblSexAge.Caption, "M") > 0, "남자", "여자")
        strPrtData(5) = lblPnlNm.Caption
        strPrtData(6) = lblDeptNm.Caption
        strPrtData(7) = "최영환"
        strPrtData(8) = "이춘희"
        strPanel = IIf(Trim(lblPnlNm.Caption) = "INHALANT", "IN", "FD")
        strPanel = IIf(Trim(lblPnlNm.Caption) = "FOOD", "FD", "IN")
        
        For intRow = 1 To tblComplete.DataRowCnt
            tblComplete.Row = intRow
            tblComplete.Col = TCompleteEnum.ccPtId
            If Trim(tblComplete.Text) = Trim(lblPtId.Caption) Then
                tblComplete.Col = TCompleteEnum.ccNo
                If Trim(tblComplete.Text) = strPanel Then
                    intCnt = intCnt + 1
                    If strPanel = "IN" Then
                        With vasINPrint
                            If intCnt = 1 Then
                                Call .SetText(3, 2, strPrtData(0)): Call .SetText(7, 2, strPrtData(3)): Call .SetText(11, 2, strPrtData(6))
                                Call .SetText(3, 3, strPrtData(1)): Call .SetText(7, 3, strPrtData(4)): Call .SetText(11, 3, strPrtData(7))
                                Call .SetText(3, 4, strPrtData(2)): Call .SetText(7, 4, lblPnlNm.Caption): Call .SetText(11, 4, strPrtData(8))
                            End If
                            For i = TCompleteEnum.ccResult To tblComplete.DataColCnt
                                tblComplete.Col = i:   strTemp = tblComplete.Text
                                Select Case mGetP(strTemp, TResultEnum.ccIntBase, DIV)
                                    ' 1KO
                                    Case "IN|IgE":      iDestRow = 6
                                    Case "IN|F14":      iDestRow = 7
                                    Case "IN|F2":       iDestRow = 8
                                    Case "IN|F1":       iDestRow = 9
                                    Case "IN|F23":      iDestRow = 10
                                    Case "IN|F24":      iDestRow = 11
                                    Case "IN|F95":      iDestRow = 12
                                    Case "IN|T35":      iDestRow = 13
                                    Case "IN|T15":      iDestRow = 14
                                    Case "IN|T2_T3":    iDestRow = 15
                                    Case "IN|T12":      iDestRow = 16
                                    Case "IN|F17":      iDestRow = 17
                                    Case "IN|T24":      iDestRow = 18
                                    Case "IN|T7":       iDestRow = 19
                                    Case "IN|T14":      iDestRow = 20
                                    Case "IN|T1_T11":   iDestRow = 21
                                    Case "IN|G2":       iDestRow = 22
                                    Case "IN|G3":       iDestRow = 23
                                    Case "IN|G6":       iDestRow = 24
                                    Case "IN|G12":      iDestRow = 25
                                    Case "IN|w12":      iDestRow = 26
                                    Case "IN|I1":       iDestRow = 27
                                    Case "IN|D72":      iDestRow = 28
                                    Case "IN|G9":       iDestRow = 29
                                    Case "IN|T225":     iDestRow = 30
                                    Case "IN|F244":     iDestRow = 31
                                    Case "IN|CCDx":     iDestRow = 32
                                    Case "IN|F84":      iDestRow = 33
                                    Case "IN|F313":     iDestRow = 34
                                    Case "IN|I81":      iDestRow = 35
                                    '2KO
                                    Case "IN|W14":      iDestRow = 36
                                    Case "IN|W11":      iDestRow = 37
                                    Case "IN|W8":       iDestRow = 38
                                    Case "IN|W6":       iDestRow = 39
                                    Case "IN|W2":       iDestRow = 40
                                    Case "IN|M6":       iDestRow = 41
                                    Case "IN|M3":       iDestRow = 42
                                    Case "IN|M2":       iDestRow = 43
                                    Case "IN|M1":       iDestRow = 44
                                    Case "IN|E1":       iDestRow = 45
                                    Case "IN|E5":       iDestRow = 46
                                    Case "IN|I6":       iDestRow = 47
                                    Case "IN|HX":       iDestRow = 48
                                    Case "IN|D2":       iDestRow = 49
                                    Case "IN|D1":       iDestRow = 50
                                    Case "IN|G1":       iDestRow = 51
                                    Case "IN|G7":       iDestRow = 52
                                    Case "IN|T16":      iDestRow = 53
                                    Case "IN|W7":       iDestRow = 54
                                    Case "IN|W22sc":    iDestRow = 55
                                    Case "IN|F206":     iDestRow = 56
                                    Case "IN|F299":     iDestRow = 57
                                    Case "IN|Fx21":     iDestRow = 58
                                    Case "IN|F92":      iDestRow = 59
                                    Case "IN|F37":      iDestRow = 60
                                    Case "IN|F91":      iDestRow = 61
                                    Case "IN|F35":      iDestRow = 62
                                    Case "IN|F93":      iDestRow = 63
                                    Case "IN|K82":      iDestRow = 64
                                    Case "IN|E81":      iDestRow = 65
                                End Select
                                
                                strValue = ""
                                strClass = ""
                                If Trim(strTemp) <> "" Then
                                    strValue = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strValue = Mid(strValue, InStr(strValue, "(") + 1)
                                    strValue = Mid(strValue, 1, Len(strValue) - 1)
                                    strClass = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strClass = mGetP(strClass, 1, " ")
                                End If
                                Call .SetText(10, iDestRow, strValue)
                                Call .SetText(12, iDestRow, strClass)
                                If IsNumeric(strClass) Then
                                    If CCur(strClass) >= 2 Then
                                        Call .SetText(4, iDestRow, "*")
                                        Call .SetText(7, iDestRow, "*")
                                    End If
                                End If
                            Next i
                        End With
                    ElseIf strPanel = "FD" Then
                        With vasFDPrint
                                
                            If intCnt = 1 Then
                                Call .SetText(3, 2, strPrtData(0)): Call .SetText(7, 2, strPrtData(3)): Call .SetText(11, 2, strPrtData(6))
                                Call .SetText(3, 3, strPrtData(1)): Call .SetText(7, 3, strPrtData(4)): Call .SetText(11, 3, strPrtData(7))
                                Call .SetText(3, 4, strPrtData(2)): Call .SetText(7, 4, lblPnlNm.Caption): Call .SetText(11, 4, strPrtData(8))
                            End If
                            For i = TCompleteEnum.ccResult To tblComplete.DataColCnt
                                tblComplete.Col = i:   strTemp = tblComplete.Text
                                Select Case mGetP(strTemp, TResultEnum.ccIntBase, DIV)
                                    ' 3KO
                                    Case "FD|IgE":      iDestRow = 6
                                    Case "FD|F14":      iDestRow = 7
                                    Case "FD|F2":       iDestRow = 8
                                    Case "FD|F81":      iDestRow = 9
                                    Case "FD|F1":       iDestRow = 10
                                    Case "FD|F23":      iDestRow = 11
                                    Case "FD|F24":      iDestRow = 12
                                    Case "FD|F40":      iDestRow = 13
                                    Case "FD|F3":       iDestRow = 14
                                    Case "FD|F41":      iDestRow = 15
                                    Case "FD|F26":      iDestRow = 16
                                    Case "FD|F83":      iDestRow = 17
                                    Case "FD|F27":      iDestRow = 18
                                    Case "FD|Fx10":     iDestRow = 19
                                    Case "FD|F95":      iDestRow = 20
                                    Case "FD|F4":       iDestRow = 21
                                    Case "FD|F9":       iDestRow = 22
                                    Case "FD|F6":       iDestRow = 23
                                    Case "FD|F47":      iDestRow = 24
                                    Case "FD|F48":      iDestRow = 25
                                    Case "FD|F13":      iDestRow = 26
                                    Case "FD|I1":       iDestRow = 27
                                    Case "FD|D72":      iDestRow = 28
                                    Case "FD|G9":       iDestRow = 29
                                    Case "FD|T225":     iDestRow = 30
                                    Case "FD|F244":     iDestRow = 31
                                    Case "FD|CCDx":     iDestRow = 32
                                    Case "FD|F84":      iDestRow = 33
                                    Case "FD|F313":     iDestRow = 34
                                    Case "FD|I81":      iDestRow = 35
                                    ' 4KO
                                    Case "FD|F45":      iDestRow = 36
                                    Case "FD|T2_T3":    iDestRow = 37
                                    Case "FD|T7":       iDestRow = 38
                                    Case "FD|G12":      iDestRow = 39
                                    Case "FD|W6":       iDestRow = 40
                                    Case "FD|W2":       iDestRow = 41
                                    Case "FD|M6":       iDestRow = 42
                                    Case "FD|M3":       iDestRow = 43
                                    Case "FD|M2":       iDestRow = 44
                                    Case "FD|E1":       iDestRow = 45
                                    Case "FD|E5":       iDestRow = 46
                                    Case "FD|I6":       iDestRow = 47
                                    Case "FD|HX":       iDestRow = 48
                                    Case "FD|D2":       iDestRow = 49
                                    Case "FD|D1":       iDestRow = 50
                                    Case "FD|F11":      iDestRow = 51
                                    Case "FD|F25":      iDestRow = 52
                                    Case "FD|M5":       iDestRow = 53
                                    Case "FD|D70":      iDestRow = 54
                                    Case "FD|W22sc":    iDestRow = 55
                                    Case "FD|F206":     iDestRow = 56
                                    Case "FD|F299":     iDestRow = 57
                                    Case "FD|Fx21":     iDestRow = 58
                                    Case "FD|F92":      iDestRow = 59
                                    Case "FD|F37":      iDestRow = 60
                                    Case "FD|F91":      iDestRow = 61
                                    Case "FD|F35":      iDestRow = 62
                                    Case "FD|F93":      iDestRow = 63
                                    Case "FD|K82":      iDestRow = 64
                                    Case "FD|E81":      iDestRow = 65
                                End Select
                                
                                strValue = ""
                                strClass = ""
                                If Trim(strTemp) <> "" Then
                                    strValue = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strValue = Mid(strValue, InStr(strValue, "(") + 1)
                                    strValue = Mid(strValue, 1, Len(strValue) - 1)
                                    strClass = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strClass = mGetP(strClass, 1, " ")
                                End If
                                Call .SetText(10, iDestRow, strValue)
                                Call .SetText(12, iDestRow, strClass)
                                If IsNumeric(strClass) Then
                                    If CCur(strClass) >= 2 Then
                                        Call .SetText(4, iDestRow, "*")
                                        Call .SetText(7, iDestRow, "*")
                                    End If
                                End If
                            Next i
                        End With
                    
                    End If
                End If
            End If
            If intCnt = 2 Then
                Exit For
            End If
        Next
        
        If intCnt = 2 Then
            If strPanel = "IN" Then
                vasINPrint.PrintOrientation = PrintOrientationPortrait '세로출력
                vasINPrint.PrintBorder = False
                vasINPrint.Action = 13
            ElseIf strPanel = "FD" Then
                vasFDPrint.PrintOrientation = PrintOrientationPortrait '세로출력
                vasINPrint.PrintBorder = False
                vasFDPrint.Action = 13
            End If
        End If
    
    End If

End Sub


'Private Sub mnuSaveExL_Click()
'    Dim sFile As String, FileNamed As String
'    Dim irow As Long, icol As Long
'
'With AlloFile
'        .DialogTitle = "Save as Excel"
'        .FileName = ""
'        .CancelError = False
'
'        .Filter = "Text Files (*.xls)|*.xls"
'        .ShowSave
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'    End With
'    If InStr(sFile, ".xls") = 0 Then
'    sFile = sFile & ".xls"
'
'End If
'
'With tblexcel 'myExcelFile
'
'    FileNamed = sFile
'    .CreateFile FileNamed
'
'    .PrintGridLines = True
'
'    .SetFont "Arial", 10, xlsNoFormat              'font0
'    .SetFont "Arial", 10, xlsBold                  'font1
'    .SetFont "Arial", 10, xlsBold + xlsUnderline   'font2
'    .SetFont "Courier", 12, xlsItalic              'font3
'
'    For irow = 1 To GrdSheet.Rows - 1
'        GrdSheet.Row = irow
'        For icol = 1 To GrdSheet.Cols - 1
'            GrdSheet.Col = icol
'            .SetFont GrdSheet.CellFontName, 10, xlsNoFormat
'            .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, irow, icol, GrdSheet.Text
'            If Formula(GrdSheet.Row, GrdSheet.Col) > "" Then
'                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsHidden, irow, icol, Formula(GrdSheet.Row, GrdSheet.Col)
'            End If
'        Next
'    Next
'
'    .CloseFile
'    Close
'
'
'End With
'
'Close
'
'    MsgBox "Excel BIFF Spreadsheet created." & vbCrLf & "Filename: " & FileNamed, vbInformation + vbOKOnly, "Excel Class"
'End Sub

''Private Sub cmdExcel_Click()
''    Call Excel_Save
''End Sub
'Private Sub Excel_Save()
'
'    If tblResult.MaxRows = 1 Then Exit Sub
'    AlloFile.FileName = ""
'
'    AlloFile.Filter = "Excel(*.xls)|*.xls"
'    AlloFile.ShowSave
'
'    If AlloFile.FileName <> "" Then
'
'       Call Excel(AlloFile.FileName)
'    End If
'
'End Sub
'Private Sub Excel(File_Name As String)
'     Dim Ex_App As Object
''     dim Ex_Book  as
'
'    On Error GoTo Err
'
'     Screen.MousePointer = vbHourglass
'
'     Set Ex_App = CreateObject("Excel.Application")
''     Set Ex_Book = Ex_App.Workbooks.Add(1)
''     Set Ex_Sheet = Ex_Book.Worksheets(1)
''
''     Ex_App.ScreenUpdating = False
''     Ex_App.DisplayAlerts = False
'
'     Call EXCEL_DRAW
'
'
'     Ex_App.DisplayAlerts = True
'     Ex_App.ScreenUpdating = True
'
''     Ex_Book.SaveAs File_Name
'
''     Ex_Book.Close
'
'     MsgBox "Excel Complete!!"
'
'     Screen.MousePointer = vbDefault
'     Ex_App.Quit
''     Set Ex_Sheet = Nothing
''     Set Ex_Book = Nothing
''     Set Ex_App = Nothing
'
'     Exit Sub
'
'Err:
'
'    MsgBox "Excel Cancel!!"
'
'    Screen.MousePointer = vbDefault
''    Ex_App.DisplayAlerts = False
''    Ex_Book.Close
''    Ex_App.Quit
''
''     Set Ex_Sheet = Nothing
''     Set Ex_Book = Nothing
''     Set Ex_App = Nothing
'End Sub

'Private Sub EXCEL_DRAW()
'
'    Dim Title As Variant
'    Dim location As Variant
'    Dim COL_Location() As Variant
'    Dim i As Integer
'
'
'
'
'    With Ex_App
'
'        Title = Array("KIDO IP LIST", _
'                       "구분", "종류", "USE IP", "USE ID", "USE NAME", "PASSWORD", "USE MAIL", "GROUP ID", "GROUP PWD")
'
'         location = Array("A1:I2", _
'                          "A3:A3", "B3:B3", "C3:C3", "D3:D3", "E3:E3", _
'                          "F3:F3", "G3:G3", "H3:H3", "I3:I3")
'
'
'         COL_Location = Array("A", _
'                                "B", "C", "D", "E", "F", "G", "H", "I")
'
'
'         '셀 사이즈
'         .Range("A1").ColumnWidth = 18
'         .Range("B1").ColumnWidth = 18
'         .Range("C1").ColumnWidth = 18
'         .Range("D1").ColumnWidth = 18
'         .Range("E1").ColumnWidth = 18
'         .Range("F1").ColumnWidth = 18
'         .Range("G1").ColumnWidth = 18
'         .Range("H1").ColumnWidth = 18
'         .Range("I1").ColumnWidth = 18
'
'          ' title 과 셀 합치기
'         For i = 0 To UBound(location)
'            Call Cell_Draw(location(i), Title, i)
'         Next i
'
'        With .Range("A2:i2")
'           .Interior.Color = vbYellow
'        End With
'
'         'DATA 보이기 (Reason별 변경건)
'         For i = 1 To grdM.Rows - 1
'
'            ' 셀 병합 <구분>
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If i = 1 Then
'                First_Value = ""
'                Last_Value = ""
'
'                first_row = 0
'                last_row = 0
'
'                First_Value = tblResult.TextMatrix(i, tblResult.ColIndex("COM")) ''해당 LINE
'                f = i
'                L = 0
'            End If
'
'            If i > 0 Then
'
'                If First_Value = tblResult.TextMatrix(i, tblResult.ColIndex("COM")) Then
'
'                    first_row = f
'                    L = L + 1
'                    last_row = L
'
'
'                    '초기화
'                    f = i
'                    L = i
'                    First_Value = tblResult.TextMatrix(f, tblResult.ColIndex("COM"))
'
'                End If
'            End If
'
'
'          .Cells(i + 3, "A") = tblResult.TextMatrix(i, tblResult.ColIndex("COM"))
'          .Cells(i + 3, "B") = tblResult.TextMatrix(i, tblResult.ColIndex("KIND"))
'          .Cells(i + 3, "C") = tblResult.TextMatrix(i, tblResult.ColIndex("USEIP"))
'          .Cells(i + 3, "D") = tblResult.TextMatrix(i, tblResult.ColIndex("USEID"))
'          .Cells(i + 3, "E") = tblResult.TextMatrix(i, tblResult.ColIndex("USENAME"))
'          .Cells(i + 3, "F") = tblResult.TextMatrix(i, tblResult.ColIndex("PASSWORD"))
'          .Cells(i + 3, "G") = tblResult.TextMatrix(i, tblResult.ColIndex("EMAIL"))
'          .Cells(i + 3, "H") = tblResult.TextMatrix(i, tblResult.ColIndex("GROUPID"))
'          .Cells(i + 3, "I") = tblResult.TextMatrix(i, tblResult.ColIndex("GROUPPWD"))
'         Next
'
'
'        '기본 FONT, FONTSIZE 정하기
'        With .Range("A4:i" & CStr(i + 3))
'            .Font.Name = "돋움체"
'            .Font.Size = "9"
'        End With
'
'        '기본정렬하기
'        With .Range("A4:i" & CStr(i + 3))
'            .VerticalAlignment = xlCenter
'            .HorizontalAlignment = xlCenter
'        End With
'
''        .Range("A2:i2").Activate
''        .ActiveWindow.FreezePanes = True
'
'        End With
'      End Sub
'      '셀을 채우자
'Private Sub Cell_Draw(Location_Name As Variant, HeadName As Variant, ARR_xl As Integer)
'    With Ex_App
'
'        With .Range(Location_Name)
'              .Select
'              .VerticalAlignment = xlCenter
'              .WrapText = False
'              .Orientation = 0
'              .AddIndent = True
'              .ShrinkToFit = True
'              .MergeCells = True
'              .Value = HeadName(ARR_xl)
'              Select Case ARR_xl
'              Case 0
'                    .HorizontalAlignment = xlCenter
'              Case 2 ' right
'                    .HorizontalAlignment = xlCenter
'
'         End With
'
'
'         With .Selection.Font
'               .Name = "돋움체"
'               Select Case ARR_xl
'               Case 0
'                .Size = 20
'                .Underline = xlUnderlineStyleSingle  ' 문자밑에 밑줄긋기
'               Case Else
'
'                .Size = 8
'               End Select
'               .FontStyle = "Bold"
'
'         End With
'
'   End With
'End Sub

'Private Sub cmdSearch_Click()
'
''    Call GetOrderByAccNo
'
'    Dim Rs          As ADODB.Recordset
'    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
'    Dim strEqpCd    As String           '장비코드
'    Dim strFromDt   As String           'From Date
'    Dim strToDt     As String           'To Date
'    Dim strKey      As String           'Spread의 키(SpcYy+SpcNo)
'    Dim strSpcYy    As String           '바코드번호(연도)
'    Dim strSpcNo    As String           '바코드번호(순번)
'    Dim strTemp     As String
'    Dim vBarNo      As String
'
'    'strEqpCd = Trim$(txtEqpCd.Text)
'    strFromDt = Format$(dtpFrDate.Value, "YYYYMMDD")
'    strToDt = Format$(dtpToDate.Value, "YYYYMMDD")
'    'If strEqpCd = "" Then
'    '    MsgBox "장비를 선택하세요.", vbInformation, "정보"
'    '    Exit Sub
'    'End If
'
'    Me.MousePointer = vbHourglass
'    Call mTblClear(tblReady)
'
'On Error GoTo Errors
'    Set objAccInfo = New clsIISAccInfo
'    Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, strFromDt, strToDt)
'
'    If Not (Rs.BOF Or Rs.EOF) Then
'        With tblReady
'            tblReady.Visible = False
'            Do Until Rs.EOF
'                strSpcYy = Rs.Fields("SPCYY").Value
'                strSpcNo = Rs.Fields("SPCNO").Value
'
'                vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
'                vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
'
'                Call GetOrder(vBarNo)
''''                strSpcYy = Rs.Fields("SPCYY").Value
''''                strSpcNo = Rs.Fields("SPCNO").Value
''''                strKey = strSpcYy & strSpcNo
''''                If strTemp <> strKey Then
''''                    '## 다른 바코드번호 일때는 모든정보 표시
''''                    If .MaxRows <= .DataRowCnt Then
''''                        .MaxRows = .MaxRows + 1
''''                        .Row = .MaxRows
''''                    Else
''''                        .Row = .DataRowCnt + 1
''''                    End If
''''
''''                    .Col = TReadyEnum.ccNo:      .Value = .Row
''''                    .Col = TReadyEnum.ccPtId:    .Value = Rs.Fields("PTID").Value & ""
''''                    .Col = TReadyEnum.ccName:    .Value = Rs.Fields("NAME").Value & ""
''''                    .Col = TReadyEnum.ccAccNo:   .Value = Rs.Fields("WORKAREA").Value & "-" & _
''''                                                          Mid$(Rs.Fields("ACCDT").Value, 3) & "-" & _
''''                                                          Rs.Fields("ACCSEQ").Value
''''
''''                    vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
''''                    vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
''''
''''                    .Col = TReadyEnum.ccBarNo:   .Value = vBarNo
'''''                    .Col = TReadyEnum.ccSexAge:  .Value = Rs.Fields("SEX").Value & "" & "/" & _
'''''                                                          mGetAge(Mid$(Rs.Fields("SSN").Value & "", 1, 6))
'''''                    .Col = TReadyEnum.ccStatFg:  .Value = IIf(Rs.Fields("STATFG").Value & "" = "1", "Y", "")
'''''                    .Col = TReadyEnum.ccWardId:  .Value = Rs.Fields("WARDID").Value & ""
'''''                    .Col = TReadyEnum.ccDept:    .Value = Rs.Fields("DEPTCD").Value & ""
'''''                    .Col = TReadyEnum.ccSpcNm:   .Value = Rs.Fields("SPCNM").Value & ""
'''''                    .Col = TReadyEnum.ccTestNms: .Value = Rs.Fields("TESTNM").Value & ""
'''''                    .Col = TReadyEnum.ccRcvNm:   .Value = Rs.Fields("RCVNM").Value & ""
'''''                    .Col = TReadyEnum.ccRcvDt:   .Value = Format$(Rs.Fields("RCVDT").Value & "", "####-##-##") & " " & _
''''                                                          Mid$(Rs.Fields("RCVTM").Value & "", 1, 2) & ":" & _
''''                                                          Mid$(Rs.Fields("RCVTM").Value & "", 3, 2)
''''
''''                    '## 1.2.3:  (2005-06-14)
''''                    '   - 처방리마크를 조회하도록 수정
'''''                    .Col = TReadyEnum.ccRmk:     .Value = Rs.Fields("MESG").Value & ""
''''                    strTemp = strKey
''''                Else
''''                    '## 같은 바코드번호 일때는 검사명만 표시
'''''                    .Col = TReadyEnum.ccTestNms
'''''                    .Value = .Value & "," & Rs.Fields("TESTNM").Value & ""
''''                End If
'                Rs.MoveNext
'            Loop
'            tblReady.Visible = True
'
''            lblCnt.Caption = CStr(.DataRowCnt)
'        End With
'    End If
'
'    Rs.Close
'    Set Rs = Nothing
'    Set objAccInfo = Nothing
'    Me.MousePointer = vbDefault
'    Exit Sub
'
'Errors:
'    Set Rs = Nothing
'    Set objAccInfo = Nothing
'    Me.MousePointer = vbDefault
'    MsgBox Err.Description, vbCritical, "오류"
'End Sub

Private Sub cmdSearch_Click()
    Dim varTemp As Variant
    Dim i       As Integer
    Dim strBarcode  As String
    Dim intRow  As Integer
    Dim blnSame As Boolean
    
    varTemp = Get_NewResult
    varTemp = Split(varTemp, "/")
    blnSame = False
    
    For i = 0 To UBound(varTemp)
        If varTemp(i) <> "" Then
            strBarcode = varTemp(i)
            With tblReady
                For intRow = 1 To .DataRowCnt
                    .Row = intRow
                    .Col = TReadyEnum.ccNo
                    If Trim(.Text) <> "SEND" Then
                        .Col = TReadyEnum.ccBarNo
                        If Mid(Trim(.Text), 1, 11) = Mid(strBarcode, 1, 11) Then
                            'strBarcode = strBarcode & CheckDisit(strBarcode)
                            'Call GetOrder(strBarcode)
                            blnSame = True
                        Else
                            blnSame = False
                        End If
                    End If
                Next
                
                If blnSame = False Then
                    'strBarcode = strBarcode & CheckDisit(strBarcode)
                    Call GetOrder(strBarcode)
                End If
                
            End With
            
        End If
    Next
End Sub

'   Newest Result Recordset
Public Function Get_NewResult() As String

    Dim Rs          As ADODB.Recordset
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim strEqpCd    As String           '장비코드
    Dim strFromDt   As String           'From Date
    Dim strToDt     As String           'To Date
    Dim strKey      As String           'Spread의 키(SpcYy+SpcNo)
    Dim strSpcYy    As String           '바코드번호(연도)
    Dim strSpcNo    As String           '바코드번호(순번)
    Dim strTemp     As String
    Dim varTemp     As Variant
    Dim i           As Integer
    Dim strTestCd   As String
    
    strFromDt = Format$(dtpFrDate.Value, "YYYYMMDD")
    strToDt = Format$(dtpToDate.Value, "YYYYMMDD")
    
    Me.MousePointer = vbHourglass
    'Call mTblClear(tblReady)
    strTemp = ""
    
On Error GoTo Errors
    Set objAccInfo = New clsIISAccInfo
    'Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, strFromDt, strToDt)
    If optTest(0).Value = True Then
        strTestCd = "'RV1201','RV1202','RV12022','RV1203','RV1204','RV1205','RV1206','RV1207','RV1208','RV1209','RV1210','RV1211','RV121101','RV121102','RV121103','RV1212','RV1213','RV1214','RV1215','RV1216'"  'RV16
    Else
        strTestCd = "'PNBPCR1','PNBPCR2','PNBPCR3','PNBPCR4','PNBPCR5'"  'RB5
    End If
    
    Set Rs = objAccInfo.GetTargetSpcs_BUNJA(mEqpCd, strFromDt, strToDt, strTestCd)
    If Not (Rs.BOF Or Rs.EOF) Then
        With tblReady
            Do Until Rs.EOF
                strSpcYy = Rs.Fields("SPCYY").Value
                strSpcNo = Rs.Fields("SPCNO").Value
        
                strKey = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
                
                'strKey = strSpcYy & strSpcNo
                If strTemp <> strKey Then
                    '## 다른 바코드번호 일때는 모든정보 표시
                    varTemp = varTemp & strKey & "/"
                    
                Else
                    '## 같은 바코드번호 일때는 검사명만 표시
'                    .Col = TReadyEnum.ccTestNms
'                    .Value = .Value & "," & Rs.Fields("TESTNM").Value & ""
                End If
                strTemp = strKey
                Rs.MoveNext
            Loop
            
        End With
    End If

    Rs.Close
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    
    Get_NewResult = varTemp
    
    Exit Function
    
Errors:
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "오류"
End Function

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
'    Me.MDIActiveX.WindowState = ccMaximize
End Sub

'   Access DB Connect
Public Function Set_DbConnect_Jet() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean
    Dim strSrcfile  As String
    Dim strDestFile As String

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError


'    FileCFX96.Path = "C:\RAPID\EXPORT\"

'    DB_Name = "C:\Program Files\LG Life Sciences\AdvanSure AlloScan 2\DB\DBresults.mdb"   '   MDB Full Path & File Name
'    UserName = "admin"  '   MDB User Name (Default = 'admin')
'    Password = "reader_admin"  '   MDB Pass Word
'
'    If (DB_Name = "") Or (UserName = "") Then
'        Set_DbConnect_Jet = False
'        Set AdoCn = Nothing
'        Exit Function
'    End If
'
'    With AdoCn
'        .ConnectionTimeout = 25
'        .CursorLocation = adUseClient
'        .Provider = "Microsoft.Jet.OLEDB.4.0"
'        .Properties("Mode").Value = adModeReadWrite
'        .Properties("Persist Security Info").Value = False
'        .Properties("Data Source").Value = DB_Name
'        .Properties("User ID").Value = UserName
'        .Properties("Jet OLEDB:Database Password").Value = Password
'        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
'        .Open
'    End With
'
'    Set_DbConnect_Jet = True
    
 Exit Function

ConnectError:
    '   오류처리
    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn.State <> adStateOpen Then
        Set_DbConnect_Jet = False
        Set AdoCn = Nothing
    End If

End Function



Private Sub Form_Load()
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    'Call GetEqpComm
    DoEvents
    
    txtFileNm.Text = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
'    2015-08-21_11-46-00
    
    dtpFrDate.Value = Now
    dtpToDate.Value = Now
    
    mIntLib.Phase = 1               '전송상태 초기화
    
    txtWorkNo.Text = 1
    'FileCFX96.Path = "C:\RAPID\EXPORT\"
    
    FileCFX96.Path = GetCFX96Config("ExportPath") & "\"
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Deactivate()
'    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISCFX96_RV16 = Nothing
End Sub



Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '에러폼 표시 클래스
    
    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing
    
    '## 에러가 없으면 버튼색깔 원래대로, 있으면 계속 빨강색
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
    
    '## 1.0.1: 이상대(2005-02-22)
    '   - Alarm창이 닫힌후 포커스를 txtBarNo로 이동
    txtBarNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll
    
'    txtBarNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    
'    With MSComm
'        '## 이미 포트가 열린경우
'        If .PortOpen Then
'            If MsgBox(Me.Caption & " 장비와 연결되어 있습니다." & vbNewLine & vbNewLine & _
'                      Me.Caption & " 인터페이스를 종료하시겠습니까?", vbYesNo + vbCritical) = vbYes Then
'                Unload Me
'            End If
'        End If
'
'    End With
    
    Unload Me
End Sub



Private Sub tmrResult_Timer()
'    Get_SearchList
End Sub

Private Sub Get_SearchList()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    Dim objBuffer    As clsIISBuffer     '버퍼 클래스
    
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 BarNO
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim strIntResult As String   '수신한 검사결과
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   'LIS 검사결과
    Dim strClass     As String   'Class 판정결과
    Dim strTemp      As String
    Dim i            As Long
'    Dim intRow       As Integer
    Dim blnSameBar   As Boolean
    Dim intCnt As Integer
    
    Dim y, Y1, Y2, Y3, X1, X2
    
    Set objIntNms = mIntLib.IntNms
    

    Set AdoRS = Get_ResultList

    If Not AdoRS.BOF Then
'        tblComplete.MaxRows = AdoRS.RecordCount
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
'           Call EditRcvData
            strBarNo = AdoRS.Fields("PATIENTID").Value
            Set objIntInfo = New clsIISIntInfo
            With objIntInfo
                .BarNo = strBarNo
                .SpcPos = strWorkNo
            End With

            For intCnt = 1 To 42
                If intCnt = 1 Then
                    If UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    ElseIf UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    End If
                ElseIf intCnt = 22 Then
                    If UCase(Trim(AdoRS.Fields("STRIPPANEL_B").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    ElseIf UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    End If
                End If
                
                
                strIntBase = Mid(strIntBase, 1, 2) & Format(intCnt, "00")
                
                If intCnt <= 21 Then
                    strResult = "BANDVAL_A" & intCnt + 1
                Else
                    strResult = "BANDVAL_B" & intCnt - 20
                End If
                
                strIntResult = AdoRS.Fields(strResult).Value
                
'                On Error Resume Next
                If Mid(strIntBase, 1, 2) = "FD" Then
                    Select Case intCnt
                    Case 9, 16, 22, 26, 27, 30, 37, 40 '-- 함수A
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        X1 = (Y3 - AN) / AM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 1, 3, 4, 6, 7, 11, 12, 15, 18, 20, 24, 39 '-- 함수B
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        'Y3 = Log(Y2)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        X1 = (Y3 - BN) / BM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 2, 5, 8, 10, 13, 14, 17, 19, 21, 23, 25, 28, 29, 31, 32, 33, 34, 35, 36, 38, 41, 42 '-- 함수C
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        'Y3 = Log(Y2)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        
                        X1 = (Y3 - CN) / CM
                        X2 = Exp(X1)
                        strIntResult = X2
                    End Select
                    

                ElseIf Mid(strIntBase, 1, 2) = "IN" Then
                    Select Case intCnt
                    Case 13, 18, 21, 22, 24, 25, 26, 29, 37, 39, 40 '-- 함수A
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - AN) / AM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 1, 3, 5, 6, 7, 9, 12, 14, 17, 19, 30, 38 '-- 함수B
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - BN) / BM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 2, 4, 8, 10, 11, 15, 16, 20, 23, 27, 28, 31, 32, 33, 34, 35, 36, 41, 42 '-- 함수C
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - CN) / CM
                        X2 = Exp(X1)
                        strIntResult = X2
                    End Select
                End If

                If strIntBase = "FD01" Or strIntBase = "IN01" Then
                    If strIntResult > 100 Then
                        strIntResult = ">100"
                        strClass = "P 증가" '증가됨
                    ElseIf strIntResult <= 100 Then
                        strIntResult = "=<100"
                        strClass = "N 정상" '정상
                    End If
                Else
                    If strIntResult < 0.35 Then
                        strClass = "0"
                        '-- 2010.04.07 단대 성백달 선생님 요구사항
                        '-- Class 0 은 "0.00" 으로 치환한다.
                        'strIntResult = Format(strIntResult, "#0.#0")
                        strIntResult = "0.00"
                    ElseIf strIntResult >= 0.35 And strIntResult < 0.7 Then
                        strClass = "1" '"*"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 0.7 And strIntResult < 3.5 Then
                        strClass = "2" '"**"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 3.5 And strIntResult < 17.5 Then
                        strClass = "3" '"***"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 17.5 And strIntResult < 50 Then
                        strClass = "4" '"****"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 50 And strIntResult < 100 Then
                        strClass = "5" '"*****"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 100 Then
                        strClass = "6" '"******"
                        strIntResult = ">=100"
                    Else
                        strIntResult = "0.00"
                        strClass = ""
                    End If
                End If
                If strIntResult = "0." Then strIntResult = "0.00"

                strIntResult = strIntResult & "  " & strClass
                strResult = strIntResult
'                If strIntBase = "FD01" Or strIntBase = "IN01" Then
'                '    strIntResult = strIntResult & " " & strClass
'                Else
'                    strIntResult = strIntResult & " " & strClass
'                End If
'                x >= 100
'                50 =< x <100
'                17.5 =< x < 50
'                3.5 =< x < 17.5
'                0.7 =< x < 3.5
'                0.35 =< x < 0.7
'                x < 0.35

                '## 결과값에 "?", ">", "<" 포함되어 있으면 에러로 표시
                'If IsNumeric(strIntResult) Then
                '    strResult = strIntResult
                'Else
                '    strResult = IISERROR
                'End If

                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                         strIntResult, strResult, strClass)
                End If
            Next
            
            Call SaveServer(objIntInfo)
            
            intRow = intRow + 1
            AdoRS.MoveNext
        Loop
    End If
    
    Set AdoRS = Nothing
    tmrResult.Enabled = False
End Sub

'   Result List Recordset
Public Function Get_ResultList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM RESULTS "
'    strSql = strSql & " WHERE EXAMDATE = '" & Format(Now, "yyyy-mm-01") & "' "
'    strSql = strSql & " WHERE PATIENTID = '10000189895' "
    
    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_ResultList = AdoRS
        'blnRS = True
    Else
        Set Get_ResultList = Nothing
        'blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
'    blnRS = False

End Function

'   Record Set Open
Public Function Get_Recordset(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                             ByVal AdoRS As ADODB.Recordset, _
                             Optional Call_Name As String, _
                             Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                             Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                             Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean

On Error GoTo DBOpenRsError
    
    With AdoRS
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    
    Get_Recordset = True

Exit Function

DBOpenRsError:
    Set AdoRS = Nothing
    Get_Recordset = False

End Function

Private Sub txtBarNo_GotFocus()
    With txtBarNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub txtBarNo_KeyDown(KeyCode As Integer, Shift As Integer)
    '## 해당 바코드번호에 대한 오더정보 조회
    If KeyCode = vbKeyReturn Then
        Me.MousePointer = vbHourglass
        Call GetOrder(Trim(txtBarNo.Text))
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub txtBarNo_KeyPress(KeyAscii As Integer)
    '## 숫자만 입력
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub tblReady_Click(ByVal Col As Long, ByVal Row As Long)
'''    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
'''    Dim vBarNo      As Variant          'Spread의 바코드번호
'''    Dim strSpcYy    As String           '검체연도
'''    Dim lngSpcNo    As Long             '검체번호
'''
'''    If Row = 0 Then Exit Sub
'''
''''    Call CtlClear(ccLabel)
''''    Call mTblClear(tblResult)
'''    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
'''    If vBarNo = "" Then Exit Sub
'''
''''    Call GetOrder(vBarNo)
'''
'''    Set objAccInfo = mIntLib.GetAccInfo(vBarNo)
'''    If Not (objAccInfo Is Nothing) Then
'''        '## tblReady, tblresult, Label에 정보표시
''''        Call SetReady(objAccInfo)
'''        Call SetLabel(objAccInfo)
'''        Call SetResult(objAccInfo)
'''
'''
''''        Call SetOrderWS(objAccInfo)
'''
'''        Set objAccInfo = Nothing
'''    End If
'''    txtBarNo.Text = "": txtBarNo.SetFocus
'''
''''    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
''''    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
''''    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
'''
'''
'''    '## tblResult, Label에 정보표시
''''    Call SetLabel(objAccInfo)
''''    Call SetResult(objAccInfo)
'''
'''    Set objAccInfo = Nothing

    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim strSpcYy    As String           '검체연도
    Dim lngSpcNo    As Long             '검체번호
    
    If Row = 0 Then Exit Sub
    
    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
    If vBarNo = "" Then Exit Sub
    
    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
    
    '## tblResult, Label에 정보표시
    Call SetLabel(objAccInfo)
    Call SetResult(objAccInfo)
    
    Set objAccInfo = Nothing


End Sub

Private Sub tblReady_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Set mPopup = New clsIISPopup
    With mPopup
        .AddMenu DELETE, "Delete"
        .AddMenu DELETEALL, "Delete All"
        .PopupMenus Me.hWnd
    End With
    Set mPopup = Nothing
End Sub

Private Sub tblComplete_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strQcFg     As String   'QC유무
    Dim strResult   As String   'LIS 결과
    Dim strTemp     As String
    Dim i           As Long
    
    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    With tblComplete
        .Row = Row
        
        .Col = TCompleteEnum.ccQcFg:    strQcFg = .Text
        If strQcFg = "0" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
            '-- 추가
            .Col = TCompleteEnum.ccNo:      lblPnlNm.Caption = IIf(.Text = "IN", "INHALANT", "FOOD")

        ElseIf strQcFg = "1" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
        End If
        
        For i = TCompleteEnum.ccResult To .DataColCnt
            .Col = i:   strTemp = .Text
            
            '## 1.0.3: 이상대(2005-06-24)
            '   - 화면표시 버그수정
            If tblResult.MaxRows <= tblResult.DataRowCnt Then
                tblResult.MaxRows = tblResult.MaxRows + 1
                tblResult.Row = tblResult.MaxRows
            Else
                tblResult.Row = tblResult.DataRowCnt + 1
            End If
            
            tblResult.Col = TResultEnum.ccTestNm:       tblResult.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
            tblResult.Col = TResultEnum.ccUnit:         tblResult.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
            tblResult.Col = TResultEnum.ccHLDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
            tblResult.Col = TResultEnum.ccDPDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
            tblResult.Col = TResultEnum.ccRef:          tblResult.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
            tblResult.Col = TResultEnum.ccEqpResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
            tblResult.Col = TResultEnum.ccClass:        tblResult.Text = mGetP(strTemp, TResultEnum.ccClass, DIV)
            tblResult.Col = TResultEnum.ccLISResult
                strResult = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                tblResult.Text = strResult
                If strResult = IISERROR Then
                    tblResult.ForeColor = vbRed
                Else
                    tblResult.ForeColor = vbBlack
                End If
        Next i
        
        'Call AllergyPrint(Col, Row)
        
    End With

End Sub



'Private Sub AllergyPrint(ByVal sPtid As Long, ByVal sPanel As Long)
'    Dim strPanel    As String
'    Dim strBarNo    As String
'    Dim strResult   As String   'LIS 결과
'    Dim strTemp     As String
'    Dim i           As Long
'    Dim j           As Integer
'
'
'    'Call mTblClear(vasPrint)
'    With tblComplete
'        .Row = Row
'        .Col = TCompleteEnum.ccNo:    strPanel = .Text
'        If strPanel = "IN" Then
'            strPanel = "INHALANT"
'        ElseIf strPanel = "FD" Then
'            strPanel = "FOOD"
'        End If
'        .Col = TCompleteEnum.ccBarNo:    strBarNo = .Text
'        If strBarNo <> "" Then
'            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
'            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
'            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
'            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
'            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
'            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
'            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
'            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
'        End If
'
'        '-- 알러지 Information
'        Call vasPrint.SetText(1, 4, "ID:")
'        Call vasPrint.SetText(2, 4, strBarNo)
'        Call vasPrint.SetText(3, 4, "AGE:")
'        Call vasPrint.SetText(4, 4, mGetP(lblSexAge.Caption, 2, "/"))
'        Call vasPrint.SetText(5, 4, "의뢰과:")
'        Call vasPrint.SetText(6, 4, lblWardNm.Caption)
'
'        Call vasPrint.SetText(1, 5, "성명:")
'        Call vasPrint.SetText(2, 5, lblName.Caption)
'        Call vasPrint.SetText(3, 5, "SEX:")
'        Call vasPrint.SetText(4, 5, IIf(mGetP(lblSexAge.Caption, 1, "/") = "M", "남자", "여자"))
'        Call vasPrint.SetText(5, 5, "검사자:")
'        Call vasPrint.SetText(6, 5, lblDoctNm.Caption)
'
'        Call vasPrint.SetText(1, 6, "검사일:")
'        Call vasPrint.SetText(2, 6, lblName.Caption)
'        Call vasPrint.SetText(3, 6, "PANEL:")
'        Call vasPrint.SetText(4, 6, strPanel)
'        Call vasPrint.SetText(5, 6, "확인자:")
'        Call vasPrint.SetText(6, 6, lblDoctNm.Caption)
'
'        '-- 알러지 Panel
'        If strPanel = "INHALANT" Then
'            Call vasPrint.SetText(3, 10, "INHALANT 1KO PANEL")
'        ElseIf strPanel = "FOOD" Then
'            Call vasPrint.SetText(3, 10, "FOOD 3KO PANEL")
'        End If
'
'        '-- 알러지 결과
'        Call vasPrint.SetText(2, 12, "Control")
'        Call vasPrint.SetText(3, 12, "[POSITIVE]")
'        j = 1
'        For i = TCompleteEnum.ccResult To .DataColCnt
'            .Col = i:   strTemp = .Text
'            Call vasPrint.SetText(1, j + 13, CStr(j))
'            Call vasPrint.SetText(2, j + 13, mGetP(strTemp, TResultEnum.ccTestNm, DIV))
'            Call vasPrint.SetText(3, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(4, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(5, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(6, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            j = j + 1
'            'vasPrint.Col = TResultEnum.ccTestNm:       vasPrint.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
'            'vasPrint.Col = TResultEnum.ccUnit:         vasPrint.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
'            'vasPrint.Col = TResultEnum.ccHLDiv:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
'            'vasPrint.Col = TResultEnum.ccDPDiv:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
'            'vasPrint.Col = TResultEnum.ccRef:          vasPrint.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
'            'vasPrint.Col = TResultEnum.ccEqpResult:    vasPrint.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
'            'vasPrint.Col = TResultEnum.ccClass:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccClass, DIV)
'            'vasPrint.Col = TResultEnum.ccLISResult
'        Next i
'    End With
'
'End Sub

Private Sub MSComm_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    
    '임시 - 데모용
    If strTransData <> "" Then GoTo TransRst
    
    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long
            
            Buffer = MSComm.Input
'임시 - 데모용
TransRst:
            Buffer = strTransData
            
            Call mIntLib.WriteLog(Buffer, ccEqp)
            
            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case mIntLib.Phase
                    Case 1
                        'Debug.Print "----> BufChar " & BufChar
                        Select Case Asc(BufChar)
                            Case 5  'ENQ()
                                '-- 임시 - 데모용 리마크 시작
                                'Debug.Print "----> ENQ 수신"
'                                MSComm.Output = Chr(6) 'ACK()
                                '-- 임시 - 데모용 리마크 끝
                                mIntLib.Phase = 2
                                mIntLib.BufCnt = 1
                        End Select
                    Case 2
                        Select Case Asc(BufChar)
                            Case 2
                                'Debug.Print "----> STX 수신"
                                mIntLib.ClearBuffer
                            Case 23 'ETB()
                                'Debug.Print "----> ETB "
'                                mIntLib.phase = 3
                            Case 3  'ETX()
                                'Debug.Print "----> ETX "
                                mIntLib.Phase = 3
                            Case Asc(vbCr)
                                'Debug.Print "----> vbCr "
                                '-- 임시 - 데모용 리마크 시작
'                                MSComm.Output = Chr(6)
                                '-- 임시 - 데모용 리마크 끝
                                mIntLib.BufCnt = mIntLib.BufCnt + 1
                                mIntLib.Phase = 3
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
                        End Select
                    Case 3
                        Select Case Asc(BufChar)
                            Case 3 'ETX
                                'Debug.Print "----> ETX "
'                                mIntLib.Phase = 4
                            Case 2
                                'Debug.Print "----> STX "
                                mIntLib.Phase = 2
                            Case 4
                                'Debug.Print "----> EOT "
                                Call EditRcvData
                                '-- 임시 - 데모용 시작
                                mIntLib.Phase = 1
                                '-- 임시 - 데모용 끝
                                
                                '-- 임시 - 데모용 리마크 시작
                                'Debug.Print "----> EditData "
'                                If mIntLib.State = "Q" Then
'                                    MSComm.Output = Chr(5)
'                                    mIntLib.SndPhase = 1
'                                    mIntLib.FrameNo = 1
'                                End If
''                                MSComm.Output = Chr(6) 'ACK()
'                                mIntLib.Phase = 4
                                '-- 임시 - 데모용 리마크 끝
                        End Select
                    Case 4
                        Select Case Asc(BufChar)
                            Case 6
                                'Debug.Print "----> ACK "
                                If mIntLib.State = "Q" Then
                                    'Call SendOrdData
                                    'Debug.Print "----> SendData "
                                End If
                            Case 5
                                'Debug.Print "----> EOT "
                                MSComm.Output = Chr(6)
                                mIntLib.Phase = 2
                            Case 21
                                'Debug.Print "----> NAK "
                                MSComm.Output = Chr(5)
                            Case 4
                                'Debug.Print "----> EOT "
                                mIntLib.Phase = 1
                        End Select
                End Select
            Next i
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
        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
        StatusBar.Panels(2).Text = ERMsg$
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    Dim objBuffer    As clsIISBuffer     '버퍼 클래스
    
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 BarNO
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim strIntResult As String   '수신한 검사결과
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   'LIS 검사결과
    Dim strTemp      As String
    Dim i            As Long
    Dim blnSameBar   As Boolean
    Dim strTest      As String
    Dim strClass     As String
    Dim varTmp       As Variant
    Dim intCnt       As Integer
    Dim strPart      As String
    Dim strState     As String
    
    Dim intRow       As Integer
    Dim intCol       As Integer
    
    strState = ""
    
    
    With vasExcel
        For intRow = 2 To .DataRowCnt Step 6
            Set objIntNms = mIntLib.IntNms
            
            For intCol = 1 To 21 '.DataColCnt
                '-- Sample 번호로 받음
                If intCol = 1 Then
                    For i = 1 To tblReady.DataRowCnt
                        tblReady.Row = i
                        tblReady.Col = TReadyEnum.ccNo
                        .Row = intRow
                        .Col = intCol
                        If Val(tblReady.Text) = Val(.Text) Then
                            tblReady.Col = TReadyEnum.ccBarNo
                            strBarNo = tblReady.Text
                            Set objIntInfo = New clsIISIntInfo
                            With objIntInfo
                                .BarNo = strBarNo
                                '.SpcPos = strPart
                                Exit For
                            End With
                        End If
                    Next
                    If strBarNo = "" Then
                        Exit Sub
                    End If
                Else
                    'If intCol >= 6 Then
                    If intCol = 6 Or intCol = 8 Or intCol = 10 Or intCol = 12 Or intCol = 14 Or intCol = 16 Or intCol = 18 Or intCol = 20 Then
                        '-- 첫 두줄
                        .Row = intRow
                        .Col = intCol
                        strIntBase = Trim(.Text)
                        .Row = intRow + 1
                        strResult = Trim(.Text)
                        
                        If strResult = "-" Then
                            strResult = "Negative"
                        ElseIf strResult = "+" Then
                            'strResult = "Positive"
                            strResult = "P"
                        End If
                        
                        If strResult <> "" And strIntBase <> "C(t)" Then
                            If objIntNms.ExistIntBase(strIntBase) Then
                                Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, "")
                                strState = "R"
                            End If
                        End If
                        
                        '-- 두번째 두줄
                        .Row = intRow + 2
                        .Col = intCol
                        strIntBase = Trim(.Text)
                        .Row = intRow + 3
                        strResult = Trim(.Text)
                        
                        If strResult = "-" Then
                            strResult = "Negative"
                        ElseIf strResult = "+" Then
                            'strResult = "Positive"
                            strResult = "P"
                        End If
                        
                        If strResult <> "" And strIntBase <> "C(t)" Then
                            If objIntNms.ExistIntBase(strIntBase) Then
                                Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, "")
                                strState = "R"
                            End If
                        End If
                        
                        '-- 세번째 두줄
                        .Row = intRow + 4
                        .Col = intCol
                        strIntBase = Trim(.Text)
                        .Row = intRow + 5
                        strResult = Trim(.Text)
                        
                        If strResult = "-" Then
                            strResult = "Negative"
                        ElseIf strResult = "+" Then
                            'strResult = "Positive"
                            strResult = "P"
                        End If
                        
                        If strResult <> "" And strIntBase <> "C(t)" Then
                            If objIntNms.ExistIntBase(strIntBase) Then
                                Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, "")
                                strState = "R"
                            End If
                        End If
                        
                    End If
                End If
            Next
            
            strBarNo = ""
            
            
            If strState = "R" Then
                Call SaveServer(objIntInfo)
            End If
            
            Set objIntNms = Nothing
            Set objBuffer = Nothing

        Next
        
    End With
    
'    Set objIntNms = Nothing
'    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_RB5()
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    Dim objBuffer    As clsIISBuffer     '버퍼 클래스
    
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 BarNO
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim strIntResult As String   '수신한 검사결과
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   'LIS 검사결과
    Dim strTemp      As String
    Dim i            As Long
    Dim blnSameBar   As Boolean
    Dim strTest      As String
    Dim strClass     As String
    Dim varTmp       As Variant
    Dim intCnt       As Integer
    Dim strPart      As String
    Dim strState     As String
    
    Dim intRow       As Integer
    Dim intRow2      As Integer
    Dim intCol       As Integer
    
    strState = ""
    
    Set objIntNms = mIntLib.IntNms
    
    With vasExcel
        For intRow = 2 To .DataRowCnt Step 6
            For intCol = 1 To 5
                '-- Sample 번호로 받음
                If intCol = 1 Then
                    For i = 1 To tblReady.DataRowCnt
                        tblReady.Row = i
                        tblReady.Col = TReadyEnum.ccNo
                        .Row = intRow
                        .Col = intCol
                        If Val(tblReady.Text) = Val(.Text) Then
                            tblReady.Col = TReadyEnum.ccBarNo
                            strBarNo = tblReady.Text
                            Set objIntInfo = New clsIISIntInfo
                            With objIntInfo
                                .BarNo = strBarNo
                                Exit For
                            End With
                        End If
                    Next
                    If strBarNo = "" Then
                        Exit Sub
                    End If
                '-- 채널, 결과
                ElseIf intCol = 5 Then
                    For intRow2 = 0 To 4
                        '-- 채널
                        .Row = intRow + intRow2
                        .Col = intCol
                        strIntBase = Trim(.Text)
                        '-- 결과
                        .Col = intCol + 1
                        strResult = Trim(.Text)
                        
                        If strResult = "-" Then
                            strResult = "Negative"
                        ElseIf strResult = "+" Then
                            'strResult = "Positive"
                            strResult = "P"
                        End If
                        
                        If strResult <> "" And strIntBase <> "C(t)" Then
                            If objIntNms.ExistIntBase(strIntBase) Then
                                Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, "")
                                strState = "R"
                            End If
                        End If
                    Next
                End If
            Next
            
            strBarNo = ""
            
            If strState = "R" Then
                Call SaveServer(objIntInfo)
            End If
        Next
        
    End With
    
    Set objIntNms = Nothing
    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과판정, 결과저장, 화면표시
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim vBarNo      As Variant 'Spread의 바코드번호
    Dim strBarNo    As String  '바코드번호
    Dim strSpcYy    As String  '검체연도
    Dim lngSpcNo    As Long    '검체번호
    Dim i           As Long
    
    Me.MousePointer = vbHourglass
    
    strBarNo = pIntInfo.BarNo
    
    '## 결과판정
    If mIntLib.CheckResult(pIntInfo) = -1 Then
        '## 접수정보가 없을때 결과표시
        Call SetComplete1(pIntInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        '## 접수정보가 있을때 결과표시
        strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
        
        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing
        
        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        StatusBar.Panels(2).Text = "검체번호:" & strBarNo & " 를 정상적으로 결과저장 했습니다."
    End If
    
    '## tblReady에서 전송된 검체삭제
    If mIntLib.BarPos = ccPC Then
        With tblReady
            For i = 1 To .DataRowCnt
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                If CStr(vBarNo) = strBarNo Then
                    Call .DeleteRows(i, 1)
                    Exit For
                End If
            Next i
        End With
    End If

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    
    If pBarNo = "" Then Exit Sub
    
    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label에 정보표시
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)
        
        
'        Call SetOrderWS(objAccInfo)
        
        Set objAccInfo = Nothing
    End If
    'txtBarNo.Text = "": txtBarNo.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
'Private Sub GetOrderByAccNo()
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
'    Dim Rs   As Recordset
'    Dim vBarNo      As Variant  'Spread의 바코드번호
'    Dim strSpcYy    As String   '검체연도
'    Dim lngSpcNo    As Long     '검체번호
'
'    Set Rs = New Recordset
'    Set objAccInfo = New clsIISAccInfo
'    'rsBarcode = objAccInfo.GetTargetSpcs(mEqpCd, Format(dtpFrDate.Value, "yyyymmdd"), Format(dtpToDate.Value, "yyyymmdd"))
'
''    If pBarNo = "" Then Exit Sub
'
'    Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, Format(dtpFrDate.Value, "yyyymmdd"), Format(dtpToDate.Value, "yyyymmdd"))
'    Do Until Rs.EOF
'        strSpcYy = Rs.Fields("SPCYY").Value
'        strSpcYy = Mid$(strSpcYy, 1, SPCYYLEN)
'        lngSpcNo = Rs.Fields("SPCNO").Value
'        lngSpcNo = Mid$(lngSpcNo, 1, SPCNOLEN)
'
'        vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(lngSpcNo, String$(SPCNOLEN, "0"))
'
'        vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
'
'        Set objAccInfo = mIntLib.GetAccInfo(vBarNo)
'
'        If Not (objAccInfo Is Nothing) Then
'            '## tblReady, tblresult, Label에 정보표시
'            Call SetReady(objAccInfo)
'            Call SetLabel(objAccInfo)
'            Call SetResult(objAccInfo)
'
'
'            Call SetOrderWS(objAccInfo)
'
'            Set objAccInfo = Nothing
'        End If
'        Rs.MoveNext
'    Loop
'
'    Rs.Close
'    Set Rs = Nothing
''    txtBarNo.Text = "": txtBarNo.SetFocus
'
'End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblReady에 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    Dim lngWorkNo As Long   'WorkNo
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim blnFood     As Boolean
    Dim blnIn       As Boolean
    
    blnFood = False
    blnIn = False
    
    
    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        
        '.Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
        .Col = TReadyEnum.ccNo:     .Text = CStr(txtWorkNo.Text)
        txtWorkNo.Text = txtWorkNo.Text + 1
        .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## 일반검체
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
        End If
        
        .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
        
        For Each objResult In pAccInfo.Results
            If InStr(UCase(objResult.TestNm), "FOOD") > 0 Then
                .Col = TReadyEnum.ccNo:     .Text = "FD"
                blnFood = True
                Exit For
            ElseIf InStr(UCase(objResult.TestNm), "INHALANT") > 0 Then
                .Col = TReadyEnum.ccNo:     .Text = "IN"
                blnIn = True
                Exit For
            End If
        Next
        
        '-- 다른 처방이 있는지 확인(Food면 IN을 IN이면 Food를 찾는다.
        If blnFood = True Then
            For Each objResult In pAccInfo.Results
                If InStr(UCase(objResult.TestNm), "INHALANT") > 0 Then
                    If .MaxRows <= .DataRowCnt Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    
                    .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
                    .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
                    .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
                    
                    If pAccInfo.QcFg = "0" Then         '## 일반검체
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
                    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
                    End If
                    
                    .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
                    .Col = TReadyEnum.ccNo:     .Text = "IN"
                    Exit For
                End If
            Next
        End If
        
        If blnIn = True Then
            For Each objResult In pAccInfo.Results
                If InStr(UCase(objResult.TestNm), "FOOD") > 0 Then
                    If .MaxRows <= .DataRowCnt Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    
                    .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
                    .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
                    .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
                    
                    If pAccInfo.QcFg = "0" Then         '## 일반검체
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
                    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
                    End If
                    
                    .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
                    .Col = TReadyEnum.ccNo:     .Text = "FD"
                    Exit For
                End If
            Next
        End If

        
        Call .SetActiveCell(1, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblComplete에 정보표시 (접수정보가 없을때)
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pIntInfo As clsIISIntInfo)
    Dim objIntResult As clsIISIntResult     '인터페이스 결과 클래스
    Dim i            As Long
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = pIntInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pIntInfo.BarNo
        .Col = TCompleteEnum.ccSendCnt: .Text = pIntInfo.IntResults.Count
        
        For Each objIntResult In pIntInfo.IntResults
            If .MaxCols <= .DataColCnt Then
                .MaxCols = .MaxCols + 1
            End If
            .Col = TCompleteEnum.ccResult + i
            .ColHidden = True
            
            '## 1.0.4: 이상대(2005-06-29)
            '   - 장비결과란에 LIS결과가 표시되는 버그수정
            .Text = objIntResult.IntNm & DIV & objIntResult.IntResult & DIV & DIV & DIV & DIV & DIV & DIV & objIntResult.Info
            i = i + 1
        Next
        Set objIntResult = Nothing
        
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblComplete에 정보표시 (접수정보가 있을때)
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim i           As Long
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## 일반검체
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.PtId
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.Name
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
            .Col = TCompleteEnum.ccDoctNm:  .Text = pAccInfo.OrdDoctNm
            .Col = TCompleteEnum.ccDeptNm:  .Text = pAccInfo.DeptNm
            .Col = TCompleteEnum.ccWardNm:  .Text = pAccInfo.WardNm
            .Col = TCompleteEnum.ccStatFg:  .Text = IIf(pAccInfo.StatFg = "1", "Y", "N")
            .Col = TCompleteEnum.ccSpcNm:   .Text = pAccInfo.SpcNm
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt
            
            For Each objResult In pAccInfo.Results
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                If objResult.IntResult <> "" Then
                    .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                            DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & _
                            DIV & mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal) & DIV & DIV & objResult.IntNm.IntBase & DIV
                            '-- 2015.08.28추가 & objResult.IntNm.IntBase & DIV
                    i = i + 1
                End If
            Next
            Set objResult = Nothing
        ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.CtrlCd
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.LevelCd
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.LotNo
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt
            
            For Each objQCResult In pAccInfo.QCResults
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                
                .Text = objQCResult.IntNm.IntNm & DIV & objQCResult.IntResult & DIV & _
                        objQCResult.RstCd & DIV & objQCResult.Unit & DIV & objQCResult.RADiv & _
                        DIV & DIV & DIV
                i = i + 1
            Next
            Set objQCResult = Nothing
        End If
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblResult 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetResult(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    
    Call mTblClear(tblResult)
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = TResultEnum.ccTestNm:    .Text = objResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objResult.RstCd
                .Col = TResultEnum.ccUnit:      .Text = objResult.Unit
                .Col = TResultEnum.ccHLDiv:     .Text = objResult.HLDiv
                .Col = TResultEnum.ccDPDiv:     .Text = objResult.DPDiv
                .Col = TResultEnum.ccRef:       .Text = mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal)
            End With
        Next
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
            End With
        Next
    End If
    
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : AlloScan 워크리스트 생성 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetOrderWS(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim mLogOn As clsIISLogOn
    
    Dim strAlloFile As String
    Dim lngFIleNum  As Long
    Dim strInFo     As String
    Dim strOldInFo  As String
    Dim blnNewFlag  As Boolean
    
    Set mLogOn = New clsIISLogOn
    
    
'    If pAccInfo.QcFg = "0" Then         '## 일반검체
            With AlloFile
                .CancelError = True
                .FileName = "C:\RAPID\IMPORT\" & Trim(txtFileNm.Text) & ".asc"
'                .ShowSave
                If Len(Dir(.FileName)) Then
                    Kill .FileName
                    blnNewFlag = True
                Else
                    blnNewFlag = False
                End If
                
                lngFIleNum = FreeFile
                
                Open .FileName For Append As #lngFIleNum
                If blnNewFlag = False Then
'[JOBLIST]
'JOBName;1/01/2009;S01-45329578234;Panel 2KO v80 UK.RDF.TST;Last01;First01;1/2/2008;MALE;
'JOBName;1/01/2009;S02-64325984359;Panel 2KO v80 UK.RDF.TST;Last01;First01;12/31/2007;MALE;
'JOBName;1/01/2009;S03-76534954562;Panel 2KO v80 UK.RDF.TST;Last03;First03;1/2/2008;MALE;
'JOBName;1/01/2009;S04-64632134814;Panel 2KO v80 UK.RDF.TST;Last04;First04;1/2/2008;MALE;
'JOBName;1/01/2009;S05-16324873488;Panel 2KO v80 UK.RDF.TST;Last05;First05;1/2/2008;MALE;
'[EOF]
                    
                    
                    Print #lngFIleNum, "[JOBLIST]"
'                    Print #lngFIleNum, "JOBName;" & Format("m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";" & "Panel 2KO v80 UK.RDF.TST"
                End If

                For Each objResult In pAccInfo.Results
                    If InStr(objResult.TestNm, "Food") > 0 Then
                        strInFo = "1"
                    ElseIf InStr(objResult.TestNm, "Inhalant") > 0 Then
                        strInFo = "2"
                    End If

                    If strOldInFo <> strInFo Then
                        If strInFo = "1" Then
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 1KO v80 UK.TST"
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 2KO v80 UK.TST"
                        Else
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 3KO v80 UK.TST"
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 4KO v80 UK.TST"
                        End If
                    End If

                    strOldInFo = strInFo
                Next
            End With
            
        Print #lngFIleNum, "[EOF]"
        Close #lngFIleNum

'    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
'        For Each objQCResult In pAccInfo.QCResults
'            With tblResult
'                If .MaxRows <= .DataRowCnt Then
'                    .MaxRows = .MaxRows + 1
'                    .Row = .MaxRows
'                Else
'                    .Row = .DataRowCnt + 1
'                End If
'
'                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
'                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
'                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
'                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
'            End With
'        Next
'    End If
    
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Label에 환자정보, 접수정보 표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetLabel(ByVal pAccInfo As clsIISAccInfo)
    Call CtlClear(ccLabel)
    
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        Call LabelShow("0")
        lblPtId.Caption = pAccInfo.PtId
        lblName.Caption = pAccInfo.Name
        lblSexAge.Caption = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
        lblDoctNm.Caption = pAccInfo.OrdDoctNm
        lblDeptNm.Caption = pAccInfo.DeptNm
        lblWardNm.Caption = pAccInfo.WardNm
        lblStatFg.Caption = IIf(pAccInfo.StatFg = "1", "Y", "N")
        lblSpcNm.Caption = pAccInfo.SpcNm
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        Call LabelShow("1")
        lblPtId.Caption = pAccInfo.CtrlCd
        lblName.Caption = pAccInfo.LevelCd
        lblSexAge.Caption = pAccInfo.LotNo
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검체종류에 따라 Label을 다르게 표시
'   인수 :
'       - pQcFg : 0(일반검체), 1(QC검체)
'-----------------------------------------------------------------------------'
Private Sub LabelShow(ByVal pQcFg As String)
    Dim i As Long
    
    If pQcFg = "0" Then         '## 일반검체
        lblControl.Caption = "환  자 ID :"
        lblLevel.Caption = "이     름 :"
        lblLotNo.Caption = "성별/나이 :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = True
        Next i
        
        lblDoctNm.Visible = True:   lblDeptNm.Visible = True
        lblWardNm.Visible = True:   lblStatFg.Visible = True
        lblSpcNm.Visible = True
    ElseIf pQcFg = "1" Then     '## QC검체
        lblControl.Caption = "Control :"
        lblLevel.Caption = "Level   :"
        lblLotNo.Caption = "Lot No  :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = False
        Next i
        
        lblDoctNm.Visible = False:   lblDeptNm.Visible = False
        lblWardNm.Visible = False:   lblStatFg.Visible = False
        lblSpcNm.Visible = False
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비설정 정보조회, 포트 Open
'-----------------------------------------------------------------------------'
Private Sub GetEqpComm()
    Dim objComm     As clsIISEqpComm    '통신설정 클래스
    Dim strErrMsg   As String           '에러메시지
    
    '## 통신설정 정보조회
    Set objComm = mIntLib.GetEqpComm
    If objComm Is Nothing Then Exit Sub
    
    With objComm
        MSComm.CommPort = .Port
        MSComm.Settings = .GetSettings
    End With
    Set objComm = Nothing

On Error GoTo Errors
    '## 포트 Open
    With MSComm
        '## 이미 포트가 열린경우
        If .PortOpen Then
            strErrMsg = mEqpCd & " 장비의 통신포트가 이미 열려있습니다."
            Error.SetLog App.EXEName, "frmIISABL835", "GetEqpComm", strErrMsg, Now
            Call mIntLib_EqpError("E004")
            Exit Sub
        End If
        
        .RThreshold = 1
        .SThreshold = 1
        .RTSEnable = True
        .PortOpen = True
    End With
    
    '## 보관일이 지난데이터 삭제
    Call mIntLib.DelHistoryData
    Exit Sub
    
Errors:
    '## 다른 장치에서 포트를 사용하는 경우
    If Err.Number = 8005 Then
        strErrMsg = mEqpCd & " 장비에 설정된 포트가 이미 사용중입니다."
        Error.SetLog App.EXEName, "frmIISABL835", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear(Optional ByVal pFlag As ClearEnum = ccAll)
    lblPtId.Caption = "":       lblName.Caption = ""
    lblSexAge.Caption = "":     lblDoctNm.Caption = ""
    lblDeptNm.Caption = "":     lblWardNm.Caption = ""
    lblStatFg.Caption = "":     lblSpcNm.Caption = ""
    lblPnlNm.Caption = ""
    If pFlag = ccAll Then
        txtBarNo.Text = "":         Call mTblClear(tblResult)
        Call mTblClear(tblReady):   Call mTblClear(tblComplete)
    End If
End Sub

'------------------------------------------------------------------'
'   기능 : 장비설정 관련 에러처리
'------------------------------------------------------------------'
Private Sub mIntLib_EqpError(ByVal pCode As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey)
End Sub

'------------------------------------------------------------------'
'   기능 : 검체관련 에러처리1
'------------------------------------------------------------------'
Private Sub mIntLib_SpcError(ByVal pCode As String, ByVal pBarNo As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo)
End Sub

'------------------------------------------------------------------'
'   기능 : 검체관련 에러처리2
'------------------------------------------------------------------'
Private Sub mIntLib_SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo, pPtId, pName)
End Sub

'------------------------------------------------------------------'
'   기능 : Popup 메뉴 Click 이벤트
'------------------------------------------------------------------'
Private Sub mPopup_Click(ByVal vMenuID As Long)
    Dim vBarNo      As Variant  'Spread의 바코드번호
    Dim strSpcYy    As String   '검체연도
    Dim lngSpcNo    As Long     '검체번호
    
    Select Case vMenuID
        Case DELETE     '## Delete
            With tblReady
                Call .GetText(TReadyEnum.ccBarNo, .ActiveRow, vBarNo)
                If vBarNo <> "" Then
                    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
                    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
                    Call mIntLib.AccInfos.Remove(strSpcYy, lngSpcNo)
                    Call .DeleteRows(.ActiveRow, 1)
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

Private Sub txtFileNm_DblClick()
    If vasINPrint.Visible = True Then
        vasINPrint.Visible = False
    Else
        vasINPrint.Visible = True
    End If
End Sub

Private Sub txtWorkNo_GotFocus()
    With txtWorkNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWorkNo_KeyPress(KeyAscii As Integer)
    '## 숫자만 입력
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
