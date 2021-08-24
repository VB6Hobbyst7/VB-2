VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   "Coapresta 2000"
   ClientHeight    =   11835
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   15315
   ScaleWidth      =   28560
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   28500
      TabIndex        =   37
      Top             =   525
      Width           =   28560
      Begin VB.Frame fraWork 
         Height          =   735
         Left            =   90
         TabIndex        =   52
         Top             =   -30
         Width           =   16185
         Begin VB.OptionButton optBW 
            Caption         =   "장비"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   69
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optBW 
            Caption         =   "PC"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   68
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtBarNum 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1740
            TabIndex        =   0
            Top             =   240
            Width           =   2205
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "ComOpen"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   12390
            TabIndex        =   65
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "선택저장"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   9990
            TabIndex        =   59
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   11190
            TabIndex        =   61
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdExcelExport 
            Caption         =   "Excel"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   14160
            TabIndex        =   62
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdRsltSearch 
            Caption         =   "결과조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   8820
            TabIndex        =   60
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdPatDelete 
            Caption         =   "선택제외"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   12960
            TabIndex        =   56
            Top             =   180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "워크조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   15360
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   1155
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   345
            Left            =   6930
            TabIndex        =   53
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   132775937
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   345
            Left            =   5070
            TabIndex        =   54
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   132775937
            CurrentDate     =   40248
         End
         Begin VB.Label Label2 
            Caption         =   "바코드"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1020
            TabIndex        =   67
            Top             =   330
            Width           =   645
         End
         Begin VB.Label Label20 
            Caption         =   "결과일자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   58
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label12 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "-"
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
            Height          =   195
            Left            =   6750
            TabIndex        =   57
            Top             =   300
            Width           =   105
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9945
      Left            =   90
      TabIndex        =   27
      Top             =   1380
      Width           =   16245
      Begin VB.CommandButton cmdSL 
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkWAll 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   690
         TabIndex        =   29
         Top             =   270
         Width           =   225
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   9645
         Left            =   90
         TabIndex        =   31
         Top             =   210
         Width           =   9345
         _Version        =   393216
         _ExtentX        =   16484
         _ExtentY        =   17013
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   16
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   16
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":14F5
      End
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   9480
         TabIndex        =   32
         Top             =   120
         Width           =   6675
         Begin VB.CommandButton cmdBarInput 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   38
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label8 
            Caption         =   "바코드번호 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   450
            TabIndex        =   66
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   36
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   5130
            TabIndex        =   35
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label6 
            Caption         =   "환자명 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   34
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1905
            TabIndex        =   33
            Top             =   240
            Width           =   1485
         End
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   9090
         Left            =   9480
         TabIndex        =   30
         Top             =   750
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   16034
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":20CF
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   28500
      TabIndex        =   19
      Top             =   0
      Width           =   28560
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1050
         TabIndex        =   25
         Top             =   90
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   132775936
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Coapresta 2000"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   4
         Left            =   3780
         TabIndex        =   44
         Top             =   90
         Width           =   2160
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
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
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Coapresta 2000"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   3810
         TabIndex        =   23
         Top             =   120
         Width           =   2160
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   13560
         Picture         =   "frmInterface.frx":2738
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   14715
         Picture         =   "frmInterface.frx":2CC2
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15870
         Picture         =   "frmInterface.frx":324C
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         Height          =   195
         Index           =   0
         Left            =   13050
         TabIndex        =   22
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         Height          =   195
         Left            =   14235
         TabIndex        =   21
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         Height          =   195
         Left            =   15360
         TabIndex        =   20
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9945
      Left            =   16560
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   5565
      Begin VB.TextBox txtTest 
         Height          =   405
         Left            =   3390
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   5790
         Width           =   1275
      End
      Begin VB.CommandButton Command16 
         Caption         =   "전송테스트"
         Height          =   435
         Left            =   3930
         TabIndex        =   63
         Top             =   4590
         Width           =   1215
      End
      Begin VB.ComboBox cboChk 
         Height          =   315
         ItemData        =   "frmInterface.frx":37D6
         Left            =   3450
         List            =   "frmInterface.frx":37E3
         TabIndex        =   51
         Top             =   4080
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CheckBox chkSaveAll 
         Caption         =   "저장포함"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   50
         Top             =   3690
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Timer tmrConnect 
         Left            =   4260
         Top             =   6510
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "장비"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1110
         TabIndex        =   46
         Top             =   5160
         Width           =   735
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "수정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1890
         TabIndex        =   45
         Top             =   5160
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   180
         TabIndex        =   42
         Top             =   8190
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '평면
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   43
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   39
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   40
            Top             =   1290
            Width           =   2760
            _Version        =   393216
            _ExtentX        =   4868
            _ExtentY        =   1826
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
            MaxCols         =   9
            SpreadDesigner  =   "frmInterface.frx":37F9
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   2715
            _Version        =   393216
            _ExtentX        =   4789
            _ExtentY        =   1720
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
            SpreadDesigner  =   "frmInterface.frx":5280
         End
      End
      Begin VB.CheckBox chkBar 
         Caption         =   "BARCODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   3090
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   3210
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   945
         Left            =   120
         TabIndex        =   16
         Top             =   2250
         Width           =   1665
         _Version        =   393216
         _ExtentX        =   2937
         _ExtentY        =   1667
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
         SpreadDesigner  =   "frmInterface.frx":54A6
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   945
         Left            =   1860
         TabIndex        =   3
         Top             =   1290
         Width           =   2535
         _Version        =   393216
         _ExtentX        =   4471
         _ExtentY        =   1667
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
         SpreadDesigner  =   "frmInterface.frx":56CC
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         Picture         =   "frmInterface.frx":58F2
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   4710
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   15
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   9
         Top             =   3240
         Width           =   1665
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   2730
         TabIndex        =   8
         Top             =   3690
         Width           =   645
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2070
         TabIndex        =   7
         Top             =   3705
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   6
         Top             =   3840
         Width           =   1635
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   1980
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   3180
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   1860
         TabIndex        =   4
         Top             =   2310
         Width           =   2835
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2220
            Top             =   300
         End
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1740
            Top             =   300
         End
         Begin MSCommLib.MSComm comEqp 
            Left            =   90
            Top             =   210
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            RThreshold      =   1
            RTSEnable       =   -1  'True
            EOFEnable       =   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   1140
            Top             =   180
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":5E7C
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":6416
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":69B0
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":6F4A
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":77DC
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7936
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7A90
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1720
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
         SpreadDesigner  =   "frmInterface.frx":7BEA
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1035
         Left            =   1860
         TabIndex        =   11
         Top             =   240
         Width           =   2505
         _Version        =   393216
         _ExtentX        =   4419
         _ExtentY        =   1826
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
         SpreadDesigner  =   "frmInterface.frx":7E10
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   1260
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1720
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
         SpreadDesigner  =   "frmInterface.frx":8036
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1425
         Left            =   3630
         TabIndex        =   70
         Top             =   6840
         Visible         =   0   'False
         Width           =   2115
         _Version        =   393216
         _ExtentX        =   3731
         _ExtentY        =   2514
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
         SpreadDesigner  =   "frmInterface.frx":825C
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과적용"
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
         Height          =   180
         Left            =   240
         TabIndex        =   49
         Top             =   5250
         Width           =   780
      End
      Begin VB.Label lblSaveSeq 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "99999"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   2790
         TabIndex        =   48
         Top             =   5250
         Width           =   615
      End
      Begin VB.Label lblExamDate 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "20160202"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   3600
         TabIndex        =   47
         Top             =   5250
         Width           =   1005
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1980
         TabIndex        =   18
         Top             =   4680
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2880
         TabIndex        =   14
         Top             =   4650
         Width           =   465
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3390
         TabIndex        =   13
         Top             =   4650
         Width           =   435
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   14910
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13582
            MinWidth        =   13582
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2018-01-02"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "오후 5:38"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnPrint 
         Caption         =   "인쇄"
         Begin VB.Menu MnPrintLand 
            Caption         =   "가로인쇄"
         End
         Begin VB.Menu MnPrintPort 
            Caption         =   "세로인쇄"
         End
      End
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "Send"
      Begin VB.Menu MnTransAuto 
         Caption         =   "Auto"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "Manual"
      End
   End
   Begin VB.Menu MnMode 
      Caption         =   "Mode"
      Visible         =   0   'False
      Begin VB.Menu MnModeBarcode 
         Caption         =   "Barcode"
      End
      Begin VB.Menu MnModeWorkList 
         Caption         =   "WorkList"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const colSpecNo = 0     '미사용
'Const colCheckBox = 1
'Const colSAVESEQ = 2    '저장순번(날짜별)
'Const colEXAMDATE = 3   '검사일자
'Const colHOSPDATE = 4   '병원접수일자
'Const colBARCODE = 5
'Const colCHARTNO = 6
'Const colPID = 7        '병록번호(내원번호)
'Const colINOUT = 8      '입원/외래
'Const colDISKNO = 9
'Const colPOSNO = 10
'Const colPNAME = 11
'Const colPSEX = 12
'Const colPAGE = 13
'Const colOCNT = 14
'Const colRCNT = 15
'Const colState = 16

'sendflag
'0: Order
'1: Result
'2: Trans
'vasres, vasrres colum
'Const colEQUIPCODE = 1
'Const colEXAMCODE = 2
'Const colEXAMNAME = 3
'Const colMachResult = 4
'Const colRESULT = 5
'Const colSeq = 6
'Const colFLAG = 7
'Const colSubCode = 8

Dim gRow As Long

Dim gsBarCode       As String
Dim gsSampleType    As String
Dim gsPID           As String
Dim gsRackNo        As String
Dim gsPosNo         As String
Dim gsResDateTime   As String
Dim gsSeqNo         As String
Dim gsExamCode      As String
Dim gsExamName      As String
Dim gsOrder         As String
Dim gsResult        As String
Dim gsFlag          As String

Dim gMT             As String
Dim gComState       As Long
Dim gErrState       As Long

Dim strBuffer       As String
Dim strORQN         As String


'===============================
Const SPCLEN As Integer = 10

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""
Const GS  As String = ""


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================
Dim AckOn   As Boolean
Dim Sample_Seq  As String
Dim aMod    As String
Dim iIID    As String



Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub chkWAll_Click()
    Dim iRow As Long
    
    With vasID
        If chkWAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 1
            Next iRow
        ElseIf chkWAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 0
            Next iRow
        End If
    End With
    
End Sub

Private Sub cmdBarInput_Click()
    If cmdBarInput.Caption = "+" Then
        cmdBarInput.Caption = "-"
        txtBarNum.Visible = True
        txtBarNum.SetFocus
    Else
        cmdBarInput.Caption = "+"
        txtBarNum.Visible = False
    End If
End Sub


Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
    Set xlBook = xlapp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlapp.Quit


End Sub

Private Sub cmdExcelExport_Click()

    Dim iRow As Integer
    Dim j As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    Dim blnWrite As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    vasPrint.MaxRows = vasID.MaxRows
    vasPrint.MaxCols = vasID.MaxCols
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
            
        If vasID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasID.MaxCols
                    SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, j
                Next
            End If
            
            For j = 1 To vasID.MaxCols
                SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, j
            Next
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasPrint
        MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
    End If
    
End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            Res = SaveTransDataW(lRow)
        
            If Res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, lRow, colBARCODE)) & "' "
                
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdOpen_Click()
    Dim strOutput As String
    
    strOutput = "0F" & "1" & ETX
    strOutput = STX & strOutput & GetChkSum(strOutput)
    
    'strOutput = "0F" & "1" & ETX
    strOutput = "0F1A"
    
    
    
'    strOutput = "031123451234567890123456201705092153590201    123.402    987.4"
'    strOutput = STX & strOutput & GetChkSum(strOutput)
    
    comEqp.Output = strOutput
    
    
    SetRawData "[Tx]" & strOutput
'comEqp.Output = RcvBuffer
    
End Sub

Private Sub cmdPatDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
    
End Sub


Private Sub cmdRsltSearch_Click()
    Dim iRow As Long
    Dim strDate As String
    Dim strSaveSeq As String
    Dim strChart As String
    Dim RS          As ADODB.Recordset
    Dim i As Integer
    Dim blnSame As Boolean
    Dim intCol As Integer
    
    
    ClearSpread vasID
    ClearSpread vasRes

    vasID.MaxRows = 0
    vasRes.MaxRows = 0
          
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE "
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    strSaveSeq = GetText(vasID, i, colSAVESEQ)
                    
                    If Trim(RS("접수일자")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    
                    If blnSame = True Then
                        For intCol = colState + 1 To vasID.MaxCols
                            If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                                If Trim(RS.Fields("REFFLAG")) = "H" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbRed
                                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbBlue
                                End If
                                Exit For
                            End If
                        Next
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1

                    SetText vasID, "0", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colPOSNO
                    
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                        Case "0": SetText vasID, "에러", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 201, 112
                        Case "1": SetText vasID, "결과", .MaxRows, colState
                        Case "2": SetText vasID, "완료", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                        Case "3": SetText vasID, "수정", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 245, 112
                    End Select
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                            If Trim(RS.Fields("REFFLAG")) = "H" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbRed
                            ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbBlue
                            End If
                            Exit For
                        End If
                    Next

                End If

                blnSame = False

            End With

            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    vasID.RowHeight(-1) = 12
    
End Sub

Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, REQ_DT AS 접수일자" & vbCrLf
    SQL = SQL & ", QC_BAR_NO AS 바코드번호, LOT_NO AS 차트번호, REQ_SEQ AS 내원번호, '입원' AS 입외" & vbCrLf
    SQL = SQL & ", '' AS R, '' AS P, REQ_SEQ AS 이름, '남자' AS 성별, REQ_SEQ AS 나이, ITEM_CD AS ITEM " & vbCrLf
    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
    SQL = SQL & " WHERE 1=1 " & vbCrLf
    If pBarNo <> "" Then
        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
    Else
        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
    End If
    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & " ORDER BY 접수일자, 바코드번호, 차트번호, 내원번호"
    
'    If pBarNo <> "" Then
'        Res = GetDBSelectVas(gServer, SQL, vasID, vasID.MaxRows + 1)
'    Else
'        Res = GetDBSelectVas(gServer, SQL, vasID)
'    End If
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub



Private Sub GetWorkList_DADESOFT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
'''          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, '' AS 접수일자" & vbCrLf
'''    SQL = SQL & ", '' AS 바코드번호, '' AS 차트번호, '' AS 내원번호, '' AS 입외" & vbCrLf
'''    SQL = SQL & ", '' AS R, '' AS P, '' AS 이름, '' AS 성별, '' AS 나이, '' AS ITEM " & vbCrLf
'''    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'''    SQL = SQL & " WHERE 1=1 " & vbCrLf
'''    If pBarNo <> "" Then
'''        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
'''    Else
'''        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
'''    End If
'''    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
'''    SQL = SQL & " ORDER BY 접수일자, 바코드번호, 차트번호, 내원번호"
    
          SQL = " SELECT DISTINCT '1', '' AS SN, '' AS 결과일시, J.접수일자 AS 접수일자," & vbCrLf
    SQL = SQL & "        L.검체번호 AS 바코드번호, A.챠트번호 AS 차트번호, J.접수번호 AS 내원번호,'입원' AS 입외, " & vbCrLf
    SQL = SQL & "        J.진료검사ID AS R, L.진료지원ID AS P,  A.환자이름 AS 이름, A.환자성별 AS 성별, A.환자나이  AS 나이, L.처방코드 + L.서브코드 AS ITEM " & vbCrLf
    SQL = SQL & "   FROM TB_진료검사 L " & vbCrLf
    SQL = SQL & "  INNER JOIN TB_진료지원 J ON (L.진료지원ID=J.진료지원ID) " & vbCrLf
    SQL = SQL & "  INNER JOIN TB_진료일반 A ON (J.진료일자=A.진료일자 AND J.챠트번호=A.챠트번호 AND J.진료번호=A.진료번호) " & vbCrLf
    SQL = SQL & "  Where 1 = 1 " & vbCrLf
    SQL = SQL & "    AND J.접수일자 Between '" & pFrDt & "' and '" & pToDt & "'" & vbCrLf
    SQL = SQL & "    AND L.검사종류 = '" & gDept_Code & "'" & vbCrLf
    SQL = SQL & "    AND L.검사상태 < 5 " & vbCrLf
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "  AND (L.검사결과 = '' OR L.검사결과 IS NULL)"
    End If
    SQL = SQL & "  ORDER BY J.접수일자, J.접수번호"
    
    
'          SQL = " SELECT DISTINCT '1', '' AS SN, '' AS 결과일시, L.접수일자 AS 접수일자," & vbCrLf
'    SQL = SQL & "        L.검체번호 AS 바코드번호, L.챠트번호 AS 차트번호, '55555' AS 내원번호,'입원' AS 입외, " & vbCrLf
'    SQL = SQL & "        L.진료검사ID AS R, L.진료지원ID AS P,  '홍길동' AS 이름, '남자' AS 성별, '35'  AS 나이, L.처방코드 + L.서브코드 AS ITEM " & vbCrLf
'    SQL = SQL & "   FROM TB_진료검사 L " & vbCrLf
'    SQL = SQL & "  Where 1 = 1 " & vbCrLf
'    SQL = SQL & "    AND L.접수일자 Between convert(datetime,'" & pFrDt & "') and convert(datetime,'" & pToDt & "')" & vbCrLf
'    SQL = SQL & "    AND L.검사종류 = '" & gDept_Code & "'" & vbCrLf
'    SQL = SQL & "    AND L.검사상태 < 5 " & vbCrLf
'    If chkSaveAll.Value = "0" Then
'        SQL = SQL & "  AND (검사결과 = '' OR 검사결과 IS NULL)"
'    End If
'    SQL = SQL & "  ORDER BY L.접수일자"
    
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    
                    SetText vasID, Trim(RS.Fields("R")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("P")) & "", .MaxRows, colPOSNO

                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_TWIN(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
'             SQL = "Select C.SPECNO , C.SNAME, C.DEPTCODE, DECODE(C.GBIO,'I','입 원 ','O','외 래 ') as GBIO, B.EXAMNAME,  B.MASTERCODE, B.CHANNEL "
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, B.JOBDATE AS 접수일자" & vbCrLf
    SQL = SQL & ",       C.SPECNO AS 바코드번호, C.PTNO AS 차트번호, C.JOBNO AS 내원번호, DECODE(C.GBIO,'I','입원','O','외래') AS 입외" & vbCrLf
    SQL = SQL & ", '' AS R, '' AS P, C.SNAME AS 이름, C.SEX AS 성별, C.AGE AS 나이, A.MASTERCODE AS ITEM " & vbCrLf
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A,"
    SQL = SQL & "       TW_HSP_OCS.TWEXAM_MASTER  B,"
    SQL = SQL & "       TW_HSP_OCS.TWEXAM_SPECMST C"
    SQL = SQL & " Where B.JOBDATE BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf '작업일자
    SQL = SQL & "   And B.EQUCODE1 = '" & gEquipCode & "'" & vbCrLf                     ' 장비코드
    SQL = SQL & "   AND C.STATUS   = '3' " & vbCrLf                                     ' 검사상태
    SQL = SQL & "   And (C.SPECNO  = A.SPECNO) " & vbCrLf
    SQL = SQL & "   And (A.MASTERCODE = B.MASTERCODE)"
    SQL = SQL & " ORDER BY 접수일자, 바코드번호, 차트번호, 내원번호"

    SetRawData "[Sql]" & SQL

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub GetWorkList_BIT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- BIT
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, SUBSTRING(O.OCMACPDTM,1,8) AS 접수일자," & vbCrLf
    SQL = SQL & "        R.RESSPMNUM AS 바코드번호, O.OCMCHTNUM AS 차트번호,R.RESOCMNUM AS 내원번호, '' AS 입외," & vbCrLf
    SQL = SQL & "        '' AS R, '' AS P, P.PBSPATNAM AS 이름, P.PBSSEXTYP AS 성별,'' AS 나이, '' AS ITEM" & vbCrLf
    SQL = SQL & "   FROM RESINF AS R, OCMINF AS O, PBSINF AS P, LABMST AS E, ODRINF AS W" & vbCrLf
    SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'" & vbCrLf
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCrLf
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = W.ODROCMNUM" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD = W.ODRCOD" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD = E.LABCOD" & vbCrLf
    '-- 저장미포함
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCrLf         '--  'I':중간 'F' 완료"
        SQL = SQL & "   AND W.ODRDELFLG = 'N'" & vbCrLf
        SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCrLf
    End If
    SQL = SQL & " ORDER BY 접수일시, 차트번호, 내원번호"

    Call SetSQLData("워크조회", SQL)
    
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_BITDLL(ByVal pBarNo As String)
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strBarCode  As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    Dim iNumber     As Integer
    Dim strJumin    As String
    Dim strSex      As String
    Dim strAge      As String
    
    If LabResultObject.GetLabStandByListByBarCode(pBarNo) Then
        With vasID
            .ReDraw = False
            For iNumber = 0 To LabResultObject.GetRowCount - 1
                For i = 1 To .DataRowCnt
                    'strDate = GetText(vasID, i, colHOSPDATE)
                    strBarCode = GetText(vasID, i, colBARCODE)
                    If LabResultObject.GetDataValue(iNumber, "RSBBARCOD") = strBarCode Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If LabResultObject.GetDataValue(iNumber, "RESLABCOD") = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Mid(LabResultObject.GetDataValue(iNumber, "LBQACPDTM"), 1, 8), .MaxRows, colHOSPDATE
                    SetText vasID, LabResultObject.GetDataValue(iNumber, "RSBBARCOD"), .MaxRows, colBARCODE
                    SetText vasID, Trim(LabResultObject.GetDataValue(iNumber, "LBQCHTNUM")), .MaxRows, colCHARTNO
                    SetText vasID, LabResultObject.GetDataValue(iNumber, "LBQPATNAM"), .MaxRows, colPNAME
                    strJumin = LabResultObject.GetDataValue(iNumber, "PBSRESNUM")
                    '9912132185519
                    If strJumin <> "" And Len(strJumin) >= 12 Then
                        If Mid(strJumin, 7, 1) = "1" Then
                            strSex = "M"
                            strAge = Format(Now, "yyyy") - Val("19" & Mid(strJumin, 1, 2))
                        ElseIf Mid(strJumin, 7, 1) = "2" Then
                            strSex = "F"
                            strAge = Format(Now, "yyyy") - Val("19" & Mid(strJumin, 1, 2))
                        ElseIf Mid(strJumin, 7, 1) = "3" Then
                            strSex = "M"
                            strAge = Format(Now, "yyyy") - Val("20" & Mid(strJumin, 1, 2))
                        ElseIf Mid(strJumin, 7, 1) = "4" Then
                            strSex = "F"
                            strAge = Format(Now, "yyyy") - Val("20" & Mid(strJumin, 1, 2))
                        End If
                        SetText vasID, strSex, .MaxRows, colPSEX
                        SetText vasID, strAge, .MaxRows, colPAGE
                        
                        With mOrder
                            .BarNo = pBarNo
                            .LBQACPNUM = Trim(LabResultObject.GetDataValue(iNumber, "LBQACPNUM"))
                            .LBQOCMNUM = Trim(LabResultObject.GetDataValue(iNumber, "LBQOCMNUM"))
                            .RESODRSEQ = Trim(LabResultObject.GetDataValue(iNumber, "RESODRSEQ"))
                            .RESSEQ = Trim(LabResultObject.GetDataValue(iNumber, "RESSEQ"))
                            .RESSUBSEQ = Trim(LabResultObject.GetDataValue(iNumber, "RESSUBSEQ"))
                        End With
                    End If
                End If
            Next

            For iNumber = 0 To LabResultObject.GetRowCount - 1
                vaSpread1.MaxRows = iNumber + 1
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "RESODRSEQ"), iNumber + 1, 1   '처방시퀀스 (OdrSeq)
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "RESLABCOD"), iNumber + 1, 2   '검사코드
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "RESSEQ"), iNumber + 1, 3      '검사접수시퀀스 (ResSeq)
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "RESSUBSEQ"), iNumber + 1, 4   '검사접수서브시퀀스 (ResSubSeq)
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQCHTNUM"), iNumber + 1, 5   '검사실접수번호(챠트번호)
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQOCMNUM"), iNumber + 1, 6   '내원번호
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQPATNAM"), iNumber + 1, 7   '이름
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQDEPCOD"), iNumber + 1, 8   '?
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQEMGYON"), iNumber + 1, 9   'Null
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQACPDTM"), iNumber + 1, 10  '처방일?
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "RESVOLFLG"), iNumber + 1, 11  '상태 (C,F..)
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "RSBBARCOD"), iNumber + 1, 12  '바코드
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "LBQACPNUM"), iNumber + 1, 13  '검사실접수번호(RsbAcpNum)
                SetText vaSpread1, LabResultObject.GetDataValue(iNumber, "PBSRESNUM"), iNumber + 1, 14  '
            Next
        
        End With
        
    End If

End Sub

Private Sub GetWorkList_GINUSDLL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    '-- 지누스
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- 검사대상자 가져오기
                 strRequest = "jobs" + vbTab + "L" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "fr_ymd" + vbTab + pFrDt + vbTab
    strRequest = strRequest & "to_ymd" + vbTab + pToDt + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + "%" + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- 바코드로 검사대상 조회(https://211.172.17.66)
    
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    If UBound(varResponse) > 0 Then
        chkWAll.Value = 1
    Else
        chkWAll.Value = 0
    End If
    
    For i = 0 To UBound(varResponse) - 1
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = UBound(varResponse) - 1
        With vasID
            If .MaxRows = 0 Then
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                
                SetText vasID, "1", intRow, colCheckBox
                SetText vasID, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colHOSPDATE  '-- 접수일자
                SetText vasID, mGetP(varResponse(i), 2, vbTab), intRow, colBARCODE              '-- 바코드번호
                SetText vasID, mGetP(varResponse(i), 6, vbTab), intRow, colPID                  '-- 내원번호
                SetText vasID, mGetP(varResponse(i), 7, vbTab), intRow, colPNAME                '-- 이름
                Select Case mGetP(varResponse(i), 13, vbTab)                                    '-- 입/외
                    Case "O": SetText vasID, "외래", intRow, colINOUT
                    Case "E": SetText vasID, "응급", intRow, colINOUT
                    Case "I": SetText vasID, "입원", intRow, colINOUT
                End Select
                Call SetOrderColor(mGetP(varResponse(i), 2, vbTab), intRow)
            Else
                '-- 같은 바코드 번호가 있는지 체크..
                intRow = GetSameRowNum(Trim(mGetP(varResponse(i), 2, vbTab)))
                If intRow = 0 Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    
                    SetText vasID, "1", intRow, colCheckBox
                    SetText vasID, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colHOSPDATE  '-- 접수일자
                    SetText vasID, mGetP(varResponse(i), 2, vbTab), intRow, colBARCODE              '-- 바코드번호
                    SetText vasID, mGetP(varResponse(i), 6, vbTab), intRow, colPID                  '-- 내원번호
                    SetText vasID, mGetP(varResponse(i), 7, vbTab), intRow, colPNAME                '-- 이름
                    Select Case mGetP(varResponse(i), 13, vbTab)                                    '-- 입/외
                        Case "O": SetText vasID, "외래", intRow, colINOUT
                        Case "E": SetText vasID, "응급", intRow, colINOUT
                        Case "I": SetText vasID, "입원", intRow, colINOUT
                    End Select
                    Call SetOrderColor(mGetP(varResponse(i), 2, vbTab), intRow)
                End If
            End If
        End With
        
        '-- 프로그레스바 진행
        frmProgress.Xprog.Value = i + 1
        DoEvents
        
    Next
    
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub SetOrderColor(ByVal pBarNo As String, ByVal pRow As Integer)
    Dim i       As Integer
    Dim intCol  As Integer
    Dim strItem As String
    
    '-- 지누스
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    
    '-- 검사ITEM 가져오기
                 strRequest = "jobs" + vbTab + "Q" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + pBarNo + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.MCD + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- 바코드로 검사대상 조회(https://211.172.17.66)
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    If UBound(varResponse) > 0 Then
        For i = 0 To UBound(varResponse) - 1
            For intCol = colState + 1 To vasID.MaxCols
                If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colState, 3) Then
                    vasID.Row = pRow
                    vasID.Col = intCol
                    vasID.BackColor = vbYellow
                    '-- 결과저장용 SEQ
                    gArrEquip(intCol - colState, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                    Exit For
                End If
            Next intCol
        Next i
    Else
        SetText vasID, "No Order", pRow, colState
    End If
    
End Sub

Private Sub cmdSearch_Click()
                
    Select Case gOCS
        Case "BIT":         Call GetWorkList_BIT(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "TWIN":        Call GetWorkList_TWIN(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "DADESOFT":    Call GetWorkList_DADESOFT(Format(dtpStartDt.Value, "yyyy-mm-dd"), Format(dtpStopDt.Value, "yyyy-mm-dd"))
        Case "GINUSDLL":    Call GetWorkList_GINUSDLL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "GINUSDB":     Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "BITSMALL":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "BITLARGE":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "MEDICHART":   Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        'Case "BITDLL":      Call GetWorkList_BITDLL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        
    End Select
    
    vasID.RowHeight(-1) = 12
    vasRes.MaxRows = 0
    
End Sub


Private Sub cmdSL_Click()
    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        vasID.Width = 16065
    Else
        cmdSL.Caption = "▶"
        vasID.Width = 9345
    End If

    Call Form_Resize

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If frmInterface.ScaleHeight = 0 Then Exit Sub
    
        
    If cmdSL.Caption = "▶" Then
        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
        vasID.Height = Frame1.Height - 300
        
        Frame1.Width = frmInterface.ScaleWidth - 200
        vasID.Width = frmInterface.ScaleWidth - 7300
        
    
        Frame6.Left = vasID.Width + 300
        vasRes.Height = vasID.Height - 550
        vasRes.Left = Frame6.Left
    Else
        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
        vasID.Height = Frame1.Height - 300
        
        Frame1.Width = frmInterface.ScaleWidth - 200
        vasID.Width = frmInterface.ScaleWidth - 300
    
        'Frame6.Left = frmInterface.ScaleWidth - vasID.Width
        'vasRes.Height = vasID.Height - 550
        'vasRes.Left = Frame6.Left
    
    End If
    
    Picture2.Width = Frame1.Width
    
    StatusBar1.Panels(3).Width = Frame1.Width - 8500
End Sub

Private Sub imgPort_DblClick()
    
    '-- 개발시에만 Remark 풀어서 테스트진행
    If FrmHideControl.Visible = True Then
        Me.Width = 16545
        FrmHideControl.Visible = False
    Else
        Me.Width = 22000
        FrmHideControl.Visible = True
    End If

End Sub

Private Sub Label1_DblClick(Index As Integer)
    If vaSpread1.Visible = True Then
        vaSpread1.Visible = False
    Else
        vaSpread1.Visible = True
    End If
End Sub

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    lblSaveSeq.Caption = ""
    lblExamDate.Caption = ""
End Sub

Private Sub Command16_Click()
    
    strBuffer = ":N1    80 81                 00620141422      15 1   7.0  2   4.1  3   0.5  4   4.5  5    34  6    20  7   417  8   239  9    97 14    85 15    14 16   0.7 18    93 19      T54     1 "
    
    strBuffer = txtTest.Text
    
    Call comEqp_OnComm
        

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
On Error GoTo RST

    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    'Me.Height = 11520
    Me.Width = 16545
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    cmdIFClear_Click
    lblclear_Click
    
    GetSetup
    
    If gSave = "True" Then
        chkMode.Caption = "Auto"
        MnTransAuto.Checked = True
        MnTransManual.Checked = False
        chkMode.Value = 1
    Else
        chkMode.Caption = "Manual"
        MnTransAuto.Checked = False
        MnTransManual.Checked = True
        chkMode.Value = 0
    End If
    
'    If gIFMode = "Barcode" Then
'        'fraBar.Visible = True
'        fraWork.Visible = False
'
'        chkMode.Caption = "Barcode"
'        MnModeBarcode.Checked = True
'        MnModeWorkList.Checked = False
'        chkBar.Value = 1
'    Else
'        'fraBar.Visible = False
'        fraWork.Visible = True
'
'        chkMode.Caption = "WorkList"
'        MnModeBarcode.Checked = False
'        MnModeWorkList.Checked = True
'        chkBar.Value = 0
'    End If
    
'    fraWork.Visible = False
    
'    If gScreen = "통합" Then
'        cmdSL.Caption = "◀"
'        vasID.Width = 14595
'    Else
'        cmdSL.Caption = "▶"
'        vasID.Width = 7725
'    End If
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    cboChk.ListIndex = 0
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If comEqp.PortOpen Then
        frmInterface.StatusBar1.Panels(2).Text = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        frmInterface.StatusBar1.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    End If

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    
    
    Set LabResultObject = New BITLabResultInterface.BITLabResultInterface
    
    '''초기화시에 아래의 함수를 반드시 호출하여 주어야 한다.
    Call LabResultObject.InitializeServer

    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
'    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    Call cmdSL_Click
    
    '-- test
'    vasID.MaxRows = 10


                    'STX        Beta (02/H) 1
                    'AA         Command code (“0F”)   2
                    'B          Instrument No.  1
                    'ETX        (03/H)  2
                    'WW         Checksum    1
'
'    comEqp.Output = STX & "0F" & "1" & ETX & "AA"
    
    tmrConnect.Interval = 50000
    tmrConnect.Enabled = True
    
    Exit Sub
    
RST:
    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    Else
        If (MsgBox(Err.Description & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    End If
    
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
           ' .TypeEditCharSet = TypeEditCharSetAlphanumeric
           ' .TypeEditCharCase = TypeEditCharCaseSetUpper
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 12
        Next
    End With
    
End Sub


Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  seqno * 10 "
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 7)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasCode.DataRowCnt
        If i = 1 Then
            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & "'"
        End If
        
        gArrEquip(i, 1) = i
        For j = 1 To 6
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
'    DisConnect_Server

    '''프로그램 종료시에 아래의 함수를 반드시 호출하여 주어야 한다.
    Call LabResultObject.FinalizeServer
    Set LabResultObject = Nothing
    
    DisConnect_Local
    
    Unload Me
    
    End
    
End Sub

Private Sub MnExamConfig_Click()
    'frmTestSet.Show
    frmTestSet.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnModeBarcode_Click()
    chkMode.Caption = "Barcode"
    MnModeBarcode.Checked = True
    MnModeWorkList.Checked = False
    chkBar.Value = 1
    
    gIFMode = "Barcode"
    Call WritePrivateProfileString("config", "IFMode", gIFMode, App.Path & "\Interface.ini")
 
End Sub

Private Sub MnModeWorkList_Click()
    chkMode.Caption = "WorkList"
    MnModeBarcode.Checked = False
    MnModeWorkList.Checked = True
    chkBar.Value = 0

    gIFMode = "WorkList"
    Call WritePrivateProfileString("config", "IFMode", gIFMode, App.Path & "\Interface.ini")

End Sub

Private Sub MnPrintLand_Click()

    vasID.PrintOrientation = PrintOrientationLandscape '가로출력
    vasID.Action = 13

End Sub

Private Sub MnPrintPort_Click()

    vasID.PrintOrientation = PrintOrientationPortrait '세로출력
    vasID.Action = 13

End Sub

'Private Sub MnScr1_Click()
'    MnScr1.Checked = True
'    MnScr2.Checked = False
'
'    gScreen = "분리"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub
'
'Private Sub MnScr2_Click()
'    MnScr1.Checked = False
'    MnScr2.Checked = True
'
'    gScreen = "통합"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1

    gSave = "True"
    Call WritePrivateProfileString("config", "AutoSave", gSave, App.Path & "\Interface.ini")

End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
    
    gSave = "False"
    Call WritePrivateProfileString("config", "AutoSave", gSave, App.Path & "\Interface.ini")

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '송신할 데이터
    
    Select Case intSndPhase
        Case 1  '## Order
'            STX        Beta (02/H) 1
'            AA         Command code (“0A”)                                           2
'            B          Instrument No.                                                  1
'            FFFFF      Host request No.                                                5
'            L          Request type                                                    1
'            M          Automatic Re-Analysis, Reflection Flag                          1
'            N          Priority Measurement Flag                                       1
'            C          Rack position                                                   1
'            DD         Rack No.                                                        2
'            EE         Sample pos                                                      2
'            aa-a       Sample ID                                                       16
'            P          Vessel type                                                     1
'            Q          Sample type                                                     1
'            GG         Number of tests (if 0, no request, below hhR repetition none)   2
'            hh         Test No number of tests GG cycle repetition                     2 --
'            R          Sampling condition max 20 cycles                                1 --
'            ETX        (03/H)                                                          1
'            WW         Checksum                                                        2

            '0A100000000101011701180073      2102020040bf
            
            strOutput = "0A"                                                            '2
            strOutput = strOutput & mOrder.Inst                                         '1
            strOutput = strOutput & "00000"                                             '5
            strOutput = strOutput & "0"                                                 '1
            strOutput = strOutput & "0"                                                 '1
            strOutput = strOutput & "0"                                                 '1
            strOutput = strOutput & mOrder.RackNo                                       '1
            strOutput = strOutput & mOrder.TubePos                                      '2
            strOutput = strOutput & mOrder.Seq                                          '2
            strOutput = strOutput & mOrder.BarNo & Space(16 - Len(mOrder.BarNo))        '16
            strOutput = strOutput & "2"                                                 '1
            strOutput = strOutput & "1"                                                 '1
            strOutput = strOutput & Format(mOrder.Count, "00")                          '2
            strOutput = strOutput & mOrder.Order                                        '020040
            strOutput = strOutput & ETX
            
            intSndPhase = 2
            
        Case 2  '## EOT
            strState = ""
            intSndPhase = 1
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput)
    comEqp.Output = strOutput
    'Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
    
End Function

'-- 지금날짜와 검사일자 비교한다
Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function

Private Sub tmrConnect_Timer()
    SQL = "Select SYSDATE FROM DUAL "
    
    Res = GetDBSelectColumn(gServer, SQL)
    
    dtpToday = Date
    'dtpStartDt = Date
    dtpStopDt = Date
        
            
End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    'Call Sleep(1000)
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) ' & GetChkSum(strSndMsg) & vbCr
    comEqp.Output = strSndMsg & vbCrLf
    
    'SetRawData "[Tx]" & strSndMsg & vbCrLf
    Debug.Print "[SndMore]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

Private Sub comEqp_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    
    Dim Buffer      As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim strOutput As String
    
    tmrConnect.Enabled = False
    
    Select Case comEqp.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            Buffer = comEqp.Input
            Call SetRawData(Buffer)
            
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
        
            
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case ENQ
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder
                        End If
                    Case vbCr
                    Case vbLf
                    Case STX
                        strBuffer = BufChar

                    Case ETX
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                        
                        Call EditRcvDataASTM
                        strBuffer = ""
                        
                        dtpToday = Date
                        
                    Case EOT
                        If strState = "Q" Then
                            comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                            
                            intSndPhase = 1
                        End If
                        strBuffer = ""
                        
                    Case Else
                        strBuffer = strBuffer & BufChar

                End Select
            Next i
        
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
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

    tmrConnect.Interval = 50000
    tmrConnect.Enabled = True

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, 표시, 검사오더만들기
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    '-- 장비수신정보 표시
    Call SetText(vasID, pBarNo, intRow, colBARCODE)             '-- 바코드
    Call SetText(vasID, mOrder.RackNo, intRow, colDISKNO)       '-- Rack
    Call SetText(vasID, mOrder.TubePos, intRow, colPOSNO)       '-- Pos
    
    '-- 환자정보 표시
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 가져오기
    Call GetSampleInfoW_BITDLL(intRow)
    
    '-- 바코드번호에 해당하는 검사코드 가져오기
    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)

    '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
    strItems = GetEquipExamCode_CP2000(gEquip, pBarNo, intRow)

    '-- 검사채널로 장비오더 만들기
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = strItems
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    
    '-- 진행상태(Order) 표시
    Call SetText(vasID, "Order", intRow, colState)
    
    '-- 현재 Row
    gRow = intRow

End Sub




'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    Dim strCBC As String
    Dim strDiff As String
    
    GetGetEquipExamCode_XN1000 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    SetRawData "[sBarcode]" & sBarcode
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    strCBC = ""
    strDiff = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'NRBC%는 오더를 안준다
'            If Trim(gReadBuf(i)) <> "NRBC%" Then
'                strExamCode = strExamCode & "^^^^" & Trim(gReadBuf(i)) & "\"
'            End If
            
            
            If Trim(gReadBuf(i)) = "WBC" Or Trim(gReadBuf(i)) = "RBC" Or Trim(gReadBuf(i)) = "HGB" Or _
                Trim(gReadBuf(i)) = "HCT" Or Trim(gReadBuf(i)) = "MCV" Or Trim(gReadBuf(i)) = "MCH" Or Trim(gReadBuf(i)) = "MCHC" Or _
                Trim(gReadBuf(i)) = "PLT" Or Trim(gReadBuf(i)) = "RDW-SD" Or Trim(gReadBuf(i)) = "RDW-CV" Or Trim(gReadBuf(i)) = "PDW" Or _
                Trim(gReadBuf(i)) = "MPV" Or Trim(gReadBuf(i)) = "P-LCR" Or Trim(gReadBuf(i)) = "PCT" Or Trim(gReadBuf(i)) = "NRBC#" Or Trim(gReadBuf(i)) = "NRBC%" Then
                
                strCBC = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
                
            End If

            If Trim(gReadBuf(i)) = "NEUT#" Or Trim(gReadBuf(i)) = "LYMPH#" Or Trim(gReadBuf(i)) = "MONO#" Or Trim(gReadBuf(i)) = "EO#" Or Trim(gReadBuf(i)) = "BASO#" Or _
                Trim(gReadBuf(i)) = "NEUT%" Or Trim(gReadBuf(i)) = "LYMPH%" Or Trim(gReadBuf(i)) = "MONO%" Or Trim(gReadBuf(i)) = "EO%" Or Trim(gReadBuf(i)) = "BASO%" Or _
                Trim(gReadBuf(i)) = "IG#" Or Trim(gReadBuf(i)) = "IG%" Then
               
                '-- ^^^^LYMPH#\가 두개인 이유는 ETB 를 장비에서 인식하지 못하기 떄문..(그 자리가 230)
                strDiff = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                
            End If
        Else
            Exit For
        End If
    Next

    strExamCode = strCBC & strDiff
    
    '-- 오더가 없을 경우 CBC만 검사하도록 한다.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
    GetGetEquipExamCode_XN1000 = strExamCode
    
End Function

'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetEquipExamCode_CP2000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    Dim strIntBase  As String           '장비기준 검사명
    Dim strTemp     As String
    
    GetEquipExamCode_CP2000 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    SQL = SQL & " ORDER BY EQUIPCODE "
    
    Res = GetDBSelectRow(gLocal, SQL)
    'strExamCode = ""

    mOrder.Count = 0
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strIntBase = Mid(Trim(gReadBuf(i)), 1, 2) & "0"
            If strIntBase <> strTemp Then
                mOrder.Count = mOrder.Count + 1
                '020040
                strTemp = strTemp & strIntBase
            End If
        End If
    Next
        
    GetEquipExamCode_CP2000 = strTemp
    
    
End Function

'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strTestDt   As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
'    If IsNumeric(pBarNo) And Len(pBarNo) = 10 Then
'        For i = 1 To vasID.DataRowCnt
'            If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
'                intRow = i
'                Exit For
'            End If
'        Next i
'    Else
'        For i = 1 To vasID.DataRowCnt
'            If Trim(GetText(vasID, i, colCheckBox)) = "1" Then
'                pBarNo = Trim(GetText(vasID, i, colBARCODE))
'                intRow = i
'                Exit For
'            End If
'        Next i
'    End If
    
    
    If optBW(0).Value = True Then
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
                intRow = i
                Exit For
            End If
        Next i
    Else
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colCheckBox)) = "1" Then
                pBarNo = Trim(GetText(vasID, i, colBARCODE))
                intRow = i
                Exit For
            End If
        Next i
    End If
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    '-- 장비수신정보 표시
    Call SetText(vasID, "1", intRow, colCheckBox)
    Call SetText(vasID, pBarNo, intRow, colBARCODE)
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 서버테이블에서 가져와 표시(for 워크리스트)  '6,7,8,9
    Call GetSampleInfoW_BITDLL(intRow)
    
    '-- 바코드번호에 존재하는 검사코드 가져오기(인수 : 장비코드,챠트번호,접수일,내원번호,검진번호)
    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

Private Sub SendMessage_1200(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim R As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & RS _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & RS _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & RS & ETX
            
        Case "SMP_ORD"
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    comEqp.Output = sSendData
    
SendMessage_Error:
    If Err <> 0 Then
'        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub

Private Sub GetaModiIID(ByVal sMsg As String)

    Dim tmpData()   As String
    
    '<STX>SYS_READY<FS><RS>aMOD<GS>1265<GS><GS><GS><FS>iIID
    '<GS>12345<GS><GS><GS><FS>aDATE<GS>20Jan2004<GS><GS><GS>
    '<FS>aTIME<GS>13:35:32<GS><GS><GS><FS>iOID<GS>3<GS><GS><GS><FS>
    '<ETX>{chksum}<EOT>

    tmpData() = Split(sMsg, GS)
    
    'aMod
    aMod = Trim(tmpData(1))
    
    'iIID
    iIID = Trim(tmpData(5))

End Sub


Private Function ConvertDateType(ByVal sDate As String) As String
    On Error GoTo ErrRtn
    
    Dim kk%
    Dim sTmp$
    Dim tmpYYYY$, tmpMM$, tmpDD$
    
    ConvertDateType = sDate
    
    tmpYYYY = Right(sDate, 4)
    sDate = Mid(sDate, 1, Len(sDate) - 4)
    
    For kk = 1 To Len(sDate)
        sTmp = Mid(sDate, kk, 1)
        If IsNumeric(sTmp) Then
            tmpDD = tmpDD & sTmp
        Else
            tmpMM = tmpMM & sTmp
        End If
    Next kk
    
    sTmp = tmpDD & Space(1) & tmpMM & Space(1) & tmpYYYY
    
    ConvertDateType = Format(sTmp, "YYYYMMDD")
    
ErrRtn:
    If Err <> 0 Then
        'RaiseEvent DispMsg("ConvertDateType - " & Err.Description)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strTestCnt   As String
    Dim strInstNo    As String   '수신한
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    Dim strPTAPTT    As String
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim strIDRecord  As String   '수신한 Identifyer Record
    Dim Pos1         As Long
    Dim Pos2         As Long
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strFlags        As String
    
    Dim strRcvMsg2      As String
    Dim strRcvMsg3      As String
    Dim strRcvMsg7      As String
    
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    Dim RcvBuffer As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
    
    Dim strCmdcode$
    Dim strOutput$
    Dim strTemp$
    
    Dim strIntB(3)   As String   '수신한 장비기준 검사명
    Dim strRSLT(3)    As String   '수신한 결과(정성)
    
    
    RcvBuffer = strBuffer
    'RcvBuffer = strRecvData(1)
    strCmdcode = Mid(RcvBuffer, 2, 2)
    
    Select Case strCmdcode
        Case "0F"   'Line Check
                    '0F1A
                    'STX        Beta (02/H) 1
                    'AA         Command code (“0F”)   2
                    'B          Instrument No.  1
                    'ETX        (03/H)  2
                    'WW         Checksum    1
                    
'                    strInstNo = Mid(RcvBuffer, 4, 1)
'                    strOutput = "0F" & strInstNo & ETX
'                    strOutput = STX & strOutput & GetChkSum(strOutput)
'                    comEqp.Output = strOutput
                    comEqp.Output = RcvBuffer
    
        Case "02"   '오더요청
                    'Analysis Request Demands
                    '031123451234567890123456201705092153590201    123.402    987.4D0
                    'STX         Beta (02/H)                1
                    'AA          Command code (“02”)      2
                    'B           Instrument No.             1
                    'C           Rack position              1
                    'DD          Rack No.                   2
                    'EE          Sample pos                 2
                    'aa-a        Sample ID                  16
                    'ETX         (03/H)                     1
                    'WW          Checksum                   2
                    
                    '021101011705260046:5    77
                    ' 02 1 1 01 01 1705260046:5    77
                    '021101021705250136      48
                    
                    strInstNo = Mid(RcvBuffer, 4, 1)
                    strRackNo = Mid(RcvBuffer, 5, 1)
                    strTubePos = Mid(RcvBuffer, 6, 2)
                    strSeq = Mid(RcvBuffer, 8, 2)
                    strBarNo = Mid(RcvBuffer, 10, 16)
                    strBarNo = Trim(strBarNo)
                    
                    With mOrder
                        .Inst = strInstNo
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .BarNo = strBarNo
                    End With
                    
                    If strBarNo = "" Then Exit Sub
                    
                    Call GetOrder(strBarNo)
                    
                    strState = "Q"
                    
                    
        Case "03", "04"     '검사결과
                    'Results Data Transmission (Batch : 03, Real Time : 04)
                    'STX  AA  B  FFFFF  aaaaaaaaaaaaaaaa  bbbb  cc  dd  ee  ff  gg  GG hh  ii  H  j IIIII  hh  ii  H  j  IIIII  Repetition ETX  WW
                    '                                                                  1-------------------
                    '                                                                                      2------------------- (MAX 4 CYCLES)
                    'STX        Beta (02/H)                                     1
                    'AA         Command code (“03”)                           2
                    'B          Instrument No.                                  1
                    'FFFFF      Host request No.                                5
                    'aa-a       Sample ID                                       16
                    'bbbb       Measurement data and time, year                 4
                    'cc         Month                                           2
                    'dd         Day                                             2
                    'ee         Hour                                            2
                    'ff         Minute                                          2
                    'gg         Second                                          2
                    'GG         Number of tests                                 2
                    'hh         Test No.                                        2
                    'ii         Judging number of tests GG cycle repetition     2
                    'H          Number of results  max. 20 cycles               1
                    'j          Unit code H cycle repetition                    1
                    'IIIII      Results max. 4 cycles                           5
                    'ETX        (03/H)                                          1
                    'WW         Checksum                                        2
    
                    strBarNo = Trim(Mid(RcvBuffer, 10, 16))
                    strTestCnt = Trim(Mid(RcvBuffer, 40, 2))
                    strPTAPTT = Trim(Mid(RcvBuffer, 42, 2))
'                    If strPTAPTT = "02" Then        'PT
'
'                    ElseIf strPTAPTT = "04" Then    'APTT
'                        strIntB(0) = ""
'                        strRslt(0) = ""
'                    End If
                    
                    If Trim(strBarNo) = "" Then Exit Sub
                    
                    With mResult
                        .BarNo = strBarNo
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    End With
                    
                    Call SetPatInfo(strBarNo)
                    
                    
                    If strTestCnt = "1" Then
                        If strPTAPTT = "02" Then
                            strTemp = Mid(RcvBuffer, 47)
                            strIntB(0) = strPTAPTT & Mid(strTemp, 1, 1)
                            strRSLT(0) = Trim(Mid(strTemp, 2, 5))
                            
                            strTemp = Mid(RcvBuffer, 47 + 6)
                            strIntB(1) = strPTAPTT & Mid(strTemp, 1, 1)
                            strRSLT(1) = Trim(Mid(strTemp, 2, 5))
                            
                            strTemp = Mid(RcvBuffer, 47 + 12)
                            strIntB(2) = strPTAPTT & Mid(strTemp, 1, 1)
                            strRSLT(2) = Trim(Mid(strTemp, 2, 5))
                            
                            
                        ElseIf strPTAPTT = "04" Then
                            strTemp = Mid(RcvBuffer, 47)
                            strIntB(0) = strPTAPTT & Mid(strTemp, 1, 1)
                            strRSLT(0) = Trim(Mid(strTemp, 2, 5))
                        End If
                    Else    '(MAX 4 CYCLES)
                        strTemp = Mid(RcvBuffer, 47)
                        strIntB(0) = strPTAPTT & Mid(strTemp, 1, 1)
                        strRSLT(0) = Trim(Mid(strTemp, 2, 5))
                        
                        strTemp = Mid(RcvBuffer, 47 + 6)
                        strIntB(1) = strPTAPTT & Mid(strTemp, 1, 1)
                        strRSLT(1) = Trim(Mid(strTemp, 2, 5))
                        
                        strTemp = Mid(RcvBuffer, 47 + 12)
                        strIntB(2) = strPTAPTT & Mid(strTemp, 1, 1)
                        strRSLT(2) = Trim(Mid(strTemp, 2, 5))
                        
                        strTemp = Mid(RcvBuffer, 47 + 23)
                        strIntB(3) = "04" & Mid(strTemp, 1, 1)
                        strRSLT(3) = Trim(Mid(strTemp, 2, 5))
                    End If
                    
'                    strTemp = Mid(RcvBuffer, 47)
                    'For intCnt = 47 To Len(RcvBuffer) Step 6
                        
                    If InStr(gOrderExam, "B1540A") > 0 Then
                    
                    Else
                        strRSLT(0) = strRSLT(0) & "(INR:" & strRSLT(2) & ")"
                    End If
                    
                    For i = 0 To UBound(strIntB)
                        If strIntB(i) <> "" Then
                            strIntBase = strIntB(i)
                            strResult = strRSLT(i)
        
                            'strTemp = Mid(strTemp, 7)
    
                            SetRawData "[결과]" & strIntBase & "," & strResult
                    
                            If strResult <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                                SQL = SQL & "  FROM EQPMASTER"
                                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                                SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                                
                                Res = GetDBSelectColumn(gLocal, SQL)
                                
                                '-- 오더 있을 경우
                                If Res > 0 Then
                                    lsExamCode = Trim(gReadBuf(0))
                                    lsExamName = Trim(gReadBuf(1))
                                    lsSeqNo = Trim(gReadBuf(2))
                                    
                                    lsResRow = vasRes.DataRowCnt + 1
                                    If vasRes.MaxRows < lsResRow Then
                                        vasRes.MaxRows = lsResRow
                                    End If
                                    
                                    '소수점 처리, 결과 형태 처리
                                    lsEquipRes = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                    lsResult_Buff = strResult
                                    
                                    '-- Work List
                                    SetText vasID, "Result", gRow, colState                 '11 진행상태
                                    
                                    '-- 결과저장용 seq
                                    For intCol = colState + 1 To vasID.MaxCols
                                        If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                            SetText vasID, strResult, gRow, intCol
                                            SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            
                                            Exit For
                                        End If
                                    Next
                        
                        
                                    '-- 결과 List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes, strComm, lsResRow, colFLAG                    'Flag
                                    'SetText vasRes, strComm, lsResRow, colSUBCODE                    'subcode
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                                
                                    lsResult_Buff = ""
                                    
                                    strState = "R"
                                    
                                '-- 오더 없을 경우
                                Else
                                
                                          SQL = "Select examcode, examname, seqno "
                                    SQL = SQL & "  From EQPMASTER"
                                    SQL = SQL & " Where equipno = '" & gEquip & "' "
                                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                                    Res = GetDBSelectColumn(gLocal, SQL)
                                                            
                                    If Res > 0 Then
                                        lsExamCode = Trim(gReadBuf(0))
                                        lsExamName = Trim(gReadBuf(1))
                                        lsSeqNo = Trim(gReadBuf(2))
                                        
                                        lsResRow = vasRes.DataRowCnt + 1
                                        If vasRes.MaxRows < lsResRow Then
                                            vasRes.MaxRows = lsResRow
                                        End If
                                        
                                        '소수점 처리, 결과 형태 처리
                                        lsEquipRes = strResult
                                        strResult = SetResult(strResult, strIntBase)
                                        lsResult_Buff = strResult
                                        
                                        '-- Work List
                                        SetText vasID, "Result", gRow, colState                 '진행상태
                                        
                                        '-- 결과저장용 seq
                                        For intCol = colState + 1 To vasID.MaxCols
                                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                                SetText vasID, strResult, gRow, intCol
                                                SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                                
                                                Exit For
                                            End If
                                        Next
                        
                                        '-- 결과 List
                                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                        SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                        '-- 로컬 저장
                                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                                        lsResult_Buff = ""
                                        'strState = "R"
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    vasRes.RowHeight(-1) = 14
                            
                    '## DB에 결과저장
                    If MnTransAuto.Checked = True And strState = "R" Then
                        Res = SaveTransDataW(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                            SetText vasID, "Failed", gRow, colState
                        Else
                            '-- 저장 성공
                            SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                            SetText vasID, "Trans", gRow, colState
                            SetText vasID, "0", gRow, colCheckBox
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(vasID, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(vasID, gRow, colSAVESEQ)) & vbCrLf
                            
                            Res = SendQuery(gLocal, SQL)
                            If Res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                        End If
                        strState = ""
                    End If
    End Select
    
End Sub


Function SetResult(asResult As String, asEquipCode As String)
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "select resprec, reflow, refhigh from EQPMASTER where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If IsNumeric(gReadBuf(0)) = True Then
        sPoint = CInt(gReadBuf(0))
        sResType = ""
        For i = 0 To sPoint
            If i = 0 Then
                sResType = "#0"
            ElseIf i = 1 Then
                sResType = sResType & ".0"
            Else
                sResType = sResType & "0"
            End If
        Next
        
        sResult = Format(sEquipRes, sResType)
    Else
        sResult = sEquipRes
    End If
    
''    If IsNumeric(gReadBuf(1)) = True Then
''        sLVal = gReadBuf(1)
''        If CCur(sLVal) > CCur(sEquipRes) Then
''            sResFlag = "H"
''        End If
''    End If
''
''    If IsNumeric(gReadBuf(2)) = True Then
''        sHVal = gReadBuf(2)
''        If CCur(sHVal) < CCur(sEquipRes) Then
''            sResFlag = ">"
''        End If
''    End If
    
    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
        sLVal = gReadBuf(1)
        sHVal = gReadBuf(2)
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
        End If
    End If
    
    gsFlag = sResFlag
    SetResult = sResult
    
End Function

' asRow1 = Work List
' asRow2 = 결과 List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    'sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    'strSaveSeq = getMaxTestNum(Mid(sExamDate, 1, 8))

    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
    SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
    SQL = SQL & ", EXAMDATE"                        '검사일자"
    SQL = SQL & ", HOSPDATE"                        '병원접수일자"
    SQL = SQL & ", EQUIPNO"                         '장비코드"
    SQL = SQL & ", BARCODE" & vbCrLf
    SQL = SQL & ", EQUIPCODE"                       '검사채널"
    SQL = SQL & ", EXAMCODE"                        '병원검사코드"
    SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
    SQL = SQL & ", EXAMNAME"
    SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
    SQL = SQL & ", SAMPLETYPE"                      '검체유형"
    SQL = SQL & ", DISKNO"
    SQL = SQL & ", POSNO"
    SQL = SQL & ", EQUIPRESULT"                     '장비결과"
    SQL = SQL & ", RESULT" & vbCrLf                 '소수점적용결과"
    SQL = SQL & ", REFFLAG"
    SQL = SQL & ", REFVALUE"
    SQL = SQL & ", CHARTNO"
    SQL = SQL & ", PID"                             '병록번호(내원번호)"
    SQL = SQL & ", PNAME" & vbCrLf
    SQL = SQL & ", PSEX"
    SQL = SQL & ", PAGE"
    SQL = SQL & ", PJUMIN"
    SQL = SQL & ", PANICVALUE"
    SQL = SQL & ", DELTAVALUE" & vbCrLf
    SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
    SQL = SQL & ", SENDDATE"
    SQL = SQL & ", EXAMUID"
    SQL = SQL & ", HOSPITAL)" & vbCrLf
    SQL = SQL & " VALUES (" & vbCrLf
'    SQL = SQL & strSaveSeq
    SQL = SQL & Trim(GetText(vasID, asRow1, colSAVESEQ))
    SQL = SQL & ",'" & sExamDate
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colHOSPDATE))
    SQL = SQL & "','" & gEquip
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colBARCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEQUIPCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSUBCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMNAME))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSeq))
    SQL = SQL & "',''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDISKNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPOSNO))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
    SQL = SQL & "',''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPNAME))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'0'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- 결과업데이트
          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If Trim(gReadBuf(0)) = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(gReadBuf(0)) + 1
        End If
    End If
    
    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If
    
End Function

Private Sub Var_Clear()
    
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""

End Sub



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "검사자를 입력해주세요."
    lblUser.Caption = InputBox(sMsg, "검사자 입력")

End Sub

Private Sub txtBarNum_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarNum) Then
            StatusBar1.Panels(3).Text = "바코드번호는 숫자만 입력이 가능합니다."
            txtBarNum = ""
            Exit Sub
        End If
        
        If Len(txtBarNum) <> 10 Then
            StatusBar1.Panels(3).Text = "바코드 자릿수를 확인하세요"
            txtBarNum = ""
            Exit Sub
        End If
        
        If Trim(txtBarNum) <> "" Then
            Call GetWorkList_BITDLL(Trim(txtBarNum))
        End If
        vasID.RowHeight(-1) = 12
        txtBarNum.Text = ""
    End If
    
End Sub


Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    If BlockRow <= 0 Then
        Exit Sub
    End If
    
    For i = BlockRow To BlockRow2
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 0 Then
            vasID.Value = 1
        Else
            vasID.Value = 0
        End If
    Next i
    
End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim RS          As ADODB.Recordset
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
    lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    
    If lblSaveSeq.Caption = "" Then
        Exit Sub
    End If
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        vasRes.MaxRows = 0
        Do Until RS.EOF
            With vasRes
                .MaxRows = .MaxRows + 1
                SetText vasRes, "0", .MaxRows, colCheckBox
                SetText vasRes, Trim(RS.Fields("EQUIPCODE")) & "", .MaxRows, colEQUIPCODE
                SetText vasRes, Trim(RS.Fields("EXAMCODE")) & "", .MaxRows, colEXAMCODE
                SetText vasRes, Trim(RS.Fields("EXAMNAME")) & "", .MaxRows, colEXAMNAME
                SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMachResult
                SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSeq
                SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colFLAG
                SetText vasRes, Trim(RS.Fields("EXAMSUBCODE")) & "", .MaxRows, colSUBCODE
                
                If Trim(RS.Fields("REFFLAG")) = "H" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbRed
                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbBlue
                End If
           
            End With
            RS.MoveNext
        Loop
    End If
    vasRes.RowHeight(-1) = 12
    
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow    As Long
    Dim iCol    As Long
    Dim lsID    As String
    Dim lsTime  As String
    Dim lsPid   As String
    Dim lsSeq   As String
    Dim i       As Integer
    Dim strResult As String
    Dim blnModify As Boolean
    
    blnModify = False
    
    iRow = vasID.ActiveRow
    iCol = vasID.ActiveCol

    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        If iCol > colState Then
            Exit Sub
        End If
        lsID = Trim(GetText(vasID, iRow, colBARCODE))
        lsPid = Trim(GetText(vasID, iRow, colPID))
        lsSeq = Trim(GetText(vasID, iRow, colSAVESEQ))

'        If lsID = "" Or lsPid = "" Or lsSeq = "" Then
'            Exit Sub
'        End If
        If lsSeq = "" Then
            Exit Sub
        End If

        If MsgBox(lsSeq & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        Res = SendQuery(gLocal, SQL)

        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            '-- 바뀐 바코드로 환자정보 불러오기
            Call GetSampleInfoW_BITDLL(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- 바코드 번호가 이전과 틀리다면 업데이트
            'If lsID <> lblChangeBar.Caption Then
            If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Trim(GetText(vasID, iRow, colHOSPDATE)) & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                'SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True

            End If
        Else
            Exit Sub
            vasID.Row = iRow
            vasID.Col = colState
            If Trim(vasID.Text) = "" Then
                Exit Sub
            End If

            '-- 결과만 수정했을 경우의 업데이트는 Delete >> Insert 순으로 한다.
            '-- Delete
                  SQL = "DELETE FROM PATRESULT "
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
            SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, iRow, colBARCODE)) & "' "

            Res = SendQuery(gLocal, SQL)
                
            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

            '-- Insert
            For i = colState + 1 To vasID.MaxCols
                vasID.Row = iRow
                vasID.Col = i
                If Trim(vasID.Text) <> "" Then
                    '-- 결과 소수점 적용
                    strResult = SetResult(Trim(GetText(vasID, iRow, i)), gArrEquip(i - colState, 2))
                    '-- H/L 일때 색표시
                    If gsFlag = "L" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbBlue
                    ElseIf gsFlag = "H" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbRed
                    End If
                    vasID.Text = strResult

                    SQL = ""
                    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
                    SQL = SQL & "SAVESEQ, EXAMDATE, HOSPDATE, EQUIPNO, BARCODE" & vbCrLf
                    SQL = SQL & ", EQUIPCODE, EXAMCODE, EXAMSUBCODE, EXAMNAME, SEQNO" & vbCrLf
                    SQL = SQL & ", SAMPLETYPE, DISKNO, POSNO, EQUIPRESULT, RESULT" & vbCrLf
                    SQL = SQL & ", REFFLAG, REFVALUE, CHARTNO, PID, PNAME" & vbCrLf
                    SQL = SQL & ", PSEX, PAGE, PJUMIN, PANICVALUE, DELTAVALUE" & vbCrLf
                    SQL = SQL & ", SENDFLAG, SENDDATE, EXAMUID, HOSPITAL)" & vbCrLf
                    SQL = SQL & " VALUES (" & vbCrLf
                    SQL = SQL & Trim(GetText(vasID, iRow, colSAVESEQ))
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colEXAMDATE))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colHOSPDATE))
                    SQL = SQL & "','" & gEquip
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colBARCODE))
                    'equipcode , examcode, examname, resprec, seqno
                    SQL = SQL & "','" & gArrEquip(i - colState, 2) 'Trim(GetText(vasRes, asRow2, colEQUIPCODE))
                    SQL = SQL & "','" & gArrEquip(i - colState, 3) 'Trim(GetText(vasRes, asRow2, colEXAMCODE))
                    SQL = SQL & "','"                              'Trim(GetText(vasRes, asRow2, colSubCode))
                    SQL = SQL & "','" & gArrEquip(i - colState, 4) 'Trim(GetText(vasRes, asRow2, colEXAMNAME))
                    SQL = SQL & "','" & gArrEquip(i - colState, 6) 'Trim(GetText(vasRes, asRow2, colSeq))
                    SQL = SQL & "',''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colDISKNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPOSNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colMachResult))
                    SQL = SQL & "','" & strResult 'Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colRESULT))
                    SQL = SQL & "','" & gsFlag & "'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colCHARTNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPID))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPNAME))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPSEX))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPAGE))
                    SQL = SQL & "',''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'3'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & gIFUser
                    SQL = SQL & "','')"

                    Res = SendQuery(gLocal, SQL)
                    SetText vasID, "수정", iRow, colState

                End If
            Next
            blnModify = True
        End If
        'SetText vasID, "수정", iRow, colState

    End If
    
'    If blnModify = True Then
'        Call cmdRsltSearch_Click
'    End If
    
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub

        vasID_Click colBARCODE, lRow
    End If
End Sub


Private Sub vasRes_KeyPress(KeyAscii As Integer)
    Dim strResult   As String
    
    With vasRes
        If KeyAscii = 13 And .ActiveCol = colRESULT And lblBarcode(0).Caption <> "" Then
            '-- 결과 소수점 적용
            strResult = SetResult(Trim(GetText(vasRes, .ActiveRow, colRESULT)), Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)))
            .Col = colRESULT
            .Text = strResult
            '-- H/L 일때 색표시
            If gsFlag = "L" Then
                vasRes.Row = .ActiveRow
                vasRes.Col = colRESULT
                vasRes.ForeColor = vbBlue
            ElseIf gsFlag = "H" Then
                vasRes.Row = .ActiveRow
                vasRes.Col = colRESULT
                vasRes.ForeColor = vbRed
            End If
            
            SetText vasRes, gsFlag, .ActiveRow, colFLAG
            
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT " & vbCrLf
            SQL = SQL & "   SET RESULT  ='" & strResult & "', " & vbCrLf
            SQL = SQL & "       REFFLAG    = '" & gsFlag & "' " & vbCrLf
            SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(lblExamDate.Caption) & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ   = " & lblSaveSeq.Caption & vbCrLf
            SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, .ActiveRow, colEXAMCODE)) & "' " & vbCrLf
            SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)) & "' " & vbCrLf

            Res = SendQuery(gLocal, SQL)

            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

        End If
    End With

End Sub

