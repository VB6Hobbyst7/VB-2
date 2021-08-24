VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{D74ED2A2-3650-4720-93BC-FDDD8DCBC769}#1.0#0"; "Han2EngOCX.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   "INPROVE"
   ClientHeight    =   10605
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   19740
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
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   10605
   ScaleWidth      =   19740
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkWAll 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   570
      TabIndex        =   40
      Top             =   840
      Width           =   225
   End
   Begin VB.CommandButton cmdSL 
      Caption         =   "▶"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   37
      Top             =   810
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   11970
      TabIndex        =   30
      Top             =   690
      Width           =   7635
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "접수번호 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "이    름 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   35
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblPtest 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "1234567890ab"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1200
         TabIndex        =   34
         Top             =   480
         Width           =   5565
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "검 사 명 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblPname 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "1234567890ab"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblPtID 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   1230
         TabIndex        =   31
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   19680
      TabIndex        =   17
      Top             =   0
      Width           =   19740
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   7950
         TabIndex        =   60
         Top             =   0
         Width           =   11625
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "화면지우기"
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
            Left            =   8820
            TabIndex        =   66
            Top             =   150
            Width           =   1155
         End
         Begin VB.CommandButton cmdExcelExport 
            Caption         =   "엑셀"
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
            Left            =   6600
            TabIndex        =   65
            Top             =   150
            Width           =   1100
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
            Height          =   405
            Left            =   5490
            TabIndex        =   64
            Top             =   150
            Width           =   1100
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
            Height          =   405
            Left            =   9990
            TabIndex        =   63
            Top             =   150
            Width           =   1305
         End
         Begin VB.TextBox txtSeq 
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
            Height          =   360
            Left            =   660
            TabIndex        =   62
            Text            =   "1"
            Top             =   180
            Width           =   525
         End
         Begin VB.CommandButton cmdAppend 
            Caption         =   "적용"
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
            Left            =   7710
            TabIndex        =   61
            Top             =   150
            Width           =   1100
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   345
            Left            =   3960
            TabIndex        =   67
            Top             =   210
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   128057345
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   345
            Left            =   2310
            TabIndex        =   68
            Top             =   210
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   128057345
            CurrentDate     =   40248
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
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
            Height          =   195
            Left            =   1350
            TabIndex        =   71
            Top             =   270
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
            Left            =   3810
            TabIndex        =   70
            Top             =   300
            Width           =   105
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "검사순번"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            TabIndex        =   69
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   60
         TabIndex        =   47
         Top             =   0
         Width           =   7935
         Begin VB.CommandButton cmdOrder 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Caption         =   "오더전송"
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
            Left            =   5310
            Style           =   1  '그래픽
            TabIndex        =   56
            Top             =   150
            Width           =   1245
         End
         Begin VB.CommandButton cmdResult 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Caption         =   "결과받기"
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
            Left            =   6570
            Style           =   1  '그래픽
            TabIndex        =   55
            Top             =   150
            Width           =   1245
         End
         Begin VB.TextBox txtCnt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   285
            Left            =   1140
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   240
            Width           =   1005
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
            Left            =   10290
            TabIndex        =   53
            Top             =   210
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.ComboBox cboChk 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmInterface.frx":14F5
            Left            =   11760
            List            =   "frmInterface.frx":1502
            TabIndex        =   52
            Top             =   210
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtStopNum 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   11100
            TabIndex        =   51
            Top             =   210
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtStartNum 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   10380
            TabIndex        =   50
            Top             =   210
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "PET 조회"
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
            Left            =   2790
            TabIndex        =   49
            Top             =   150
            Width           =   1245
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
            Height          =   405
            Left            =   4050
            TabIndex        =   48
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "건"
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
            Left            =   2280
            TabIndex        =   59
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFFFF&
            Caption         =   "조회건수"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   58
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Height          =   195
            Left            =   10950
            TabIndex        =   57
            Top             =   300
            Visible         =   0   'False
            Width           =   165
         End
      End
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   9765
      Left            =   60
      TabIndex        =   38
      Top             =   780
      Width           =   11895
      _Version        =   393216
      _ExtentX        =   20981
      _ExtentY        =   17224
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   16
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   17
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":151A
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   9075
      Left            =   11970
      TabIndex        =   39
      Top             =   1470
      Width           =   7605
      _Version        =   393216
      _ExtentX        =   13414
      _ExtentY        =   16007
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
      SpreadDesigner  =   "frmInterface.frx":1F76
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   8265
      Left            =   12660
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   9585
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   7140
         TabIndex        =   75
         Top             =   3450
         Visible         =   0   'False
         Width           =   1875
      End
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
         Left            =   6840
         TabIndex        =   74
         Top             =   3510
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.FileListBox FileB4C 
         Height          =   285
         Left            =   7050
         Pattern         =   "*.csv"
         TabIndex        =   46
         Top             =   5370
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.TextBox txtTest 
         Height          =   1785
         Left            =   5190
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.CommandButton Command16 
         Caption         =   "전송테스트"
         Height          =   435
         Left            =   5310
         TabIndex        =   41
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
      End
      Begin FPSpread.vaSpread vasExcel 
         Height          =   855
         Left            =   30
         TabIndex        =   29
         Top             =   3090
         Width           =   1935
         _Version        =   393216
         _ExtentX        =   3413
         _ExtentY        =   1508
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
         SpreadDesigner  =   "frmInterface.frx":256C
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   150
         Top             =   4650
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   3570
         TabIndex        =   22
         Top             =   7110
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '평면
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   23
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   19
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   20
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
            SpreadDesigner  =   "frmInterface.frx":27B1
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   21
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
            SpreadDesigner  =   "frmInterface.frx":2C3A
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
         TabIndex        =   18
         Top             =   3210
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   120
         TabIndex        =   14
         Top             =   2250
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1614
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
         SpreadDesigner  =   "frmInterface.frx":2E7F
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   945
         Left            =   1860
         TabIndex        =   1
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
         SpreadDesigner  =   "frmInterface.frx":30C4
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         Picture         =   "frmInterface.frx":3309
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   4710
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   330
         TabIndex        =   13
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
         TabIndex        =   7
         Top             =   3240
         Width           =   1665
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   2730
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   3705
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   3180
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   1860
         TabIndex        =   2
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
                  Picture         =   "frmInterface.frx":3893
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3E2D
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":43C7
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4961
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":51F3
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":534D
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":54A7
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   975
         Left            =   120
         TabIndex        =   8
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
         SpreadDesigner  =   "frmInterface.frx":5601
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1035
         Left            =   1860
         TabIndex        =   9
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
         SpreadDesigner  =   "frmInterface.frx":5846
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   975
         Left            =   120
         TabIndex        =   10
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
         SpreadDesigner  =   "frmInterface.frx":5A8B
      End
      Begin HAN2ENGOCXLib.Han2EngOCX Han2Eng 
         Height          =   225
         Left            =   5280
         TabIndex        =   43
         Top             =   3270
         Visible         =   0   'False
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   397
         _StockProps     =   0
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   5790
         TabIndex        =   44
         Top             =   4320
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
         Format          =   128057344
         CurrentDate     =   40457
      End
      Begin MSComDlg.CommonDialog AllergyFile 
         Left            =   6240
         Top             =   5310
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "INPROVE"
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
         Left            =   8100
         TabIndex        =   77
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "INPROVE"
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
         Left            =   8070
         TabIndex        =   76
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label lblBarcode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "12345"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Index           =   0
         Left            =   8130
         TabIndex        =   73
         Top             =   240
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000008&
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   7860
         TabIndex        =   72
         Top             =   2880
         Width           =   1155
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
         Left            =   4860
         TabIndex        =   45
         Top             =   4380
         Width           =   780
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   16
         Top             =   4680
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2880
         TabIndex        =   12
         Top             =   4650
         Width           =   465
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3390
         TabIndex        =   11
         Top             =   4650
         Width           =   435
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnPrint 
         Caption         =   "인쇄"
         Visible         =   0   'False
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
         Checked         =   -1  'True
      End
      Begin VB.Menu MnModeWorkList 
         Caption         =   "WorkList"
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
'Const colDOB = 8      '입원/외래
'Const colBREED = 9
'Const colASSAYNM = 10
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

Private Sub cmdAppend_Click()
    Dim intRow As Integer
    Dim strResult   As String
    
    With vasRes
        For intRow = 1 To .DataRowCnt
            '-- 결과 소수점 적용
            strResult = SetResult(Trim(GetText(vasRes, intRow, colMachResult)), Trim(GetText(vasRes, intRow, colEQUIPCODE)))
            .Col = colMachResult
            .Text = strResult
            
            If strResult <> "" Then
                If .BackColor = "7405514" Then
                    SetLocalDB vasID.ActiveRow, intRow, "1"
                    .Row = intRow
                    .Row2 = intRow
                    .Col = 1
                    .Col2 = colSUBCODE
                    .BlockMode = True
                    .BackColor = vbWhite
                    .BlockMode = False
                Else
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT " & vbCrLf
                    SQL = SQL & "   SET EQUIPRESULT ='" & strResult & "', " & vbCrLf
                    SQL = SQL & "       REFFLAG    = '" & gsFlag & "' " & vbCrLf
                    SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
                    SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(lblExamDate.Caption) & "' " & vbCrLf
                    SQL = SQL & "   AND SAVESEQ   = " & lblSaveSeq.Caption & vbCrLf
                    SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
                    SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, intRow, colEXAMCODE)) & "' " & vbCrLf
                    SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, intRow, colEQUIPCODE)) & "' " & vbCrLf
                    
                    Res = SendQuery(gLocal, SQL)
        
                    If Res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                End If
            End If
            
            '-- 결과 소수점 적용
            strResult = Trim(GetText(vasRes, intRow, colRESULT))
            .Col = colRESULT
            .Text = strResult
            
            If strResult <> "" Then
                If .BackColor = "7405514" Then
                    SetLocalDB vasID.ActiveRow, intRow, "1"
                    .Row = intRow
                    .Row2 = intRow
                    .Col = 1
                    .Col2 = colSUBCODE
                    .BlockMode = True
                    .BackColor = vbWhite
                    .BlockMode = False
                Else
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT " & vbCrLf
                    SQL = SQL & "   SET RESULT  ='" & strResult & "', " & vbCrLf
                    SQL = SQL & "       REFFLAG    = '" & gsFlag & "' " & vbCrLf
                    SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
                    SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(lblExamDate.Caption) & "' " & vbCrLf
                    SQL = SQL & "   AND SAVESEQ   = " & lblSaveSeq.Caption & vbCrLf
                    SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
                    SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, intRow, colEXAMCODE)) & "' " & vbCrLf
                    SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, intRow, colEQUIPCODE)) & "' " & vbCrLf
                    
                    Res = SendQuery(gLocal, SQL)
        
                    If Res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                End If
            End If
        Next
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
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlApp.Quit


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
        MsgBox "저장할 자료가 없습니다.", vbCritical + vbOKOnly, Me.Caption
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
'    StatusBar1.Panels(3).Text = ""
    txtSeq = "1"
    lblPtest.Caption = ""
    
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

Private Sub cmdOrder_Click()
    Dim intRow      As Integer
    Dim STM         As ADODB.Stream
    Dim blnSendXml  As Boolean
    Dim strHeader   As String
    Dim strBody     As String
    Dim strFileNm   As String
    Dim strAssayNm  As String
    
    Dim lngFIleNum  As Long
    Dim strInFo     As String
    Dim intCnt      As Integer
    Dim iCnt        As Integer
    Dim varTmp      As Variant
    
    Dim strBarNo    As String
    Dim strLabNo    As String
    Dim strFNm      As String
    Dim strLNm      As String
    Dim strSex      As String
    Dim strAge      As String
    Dim i           As Integer
    
    Screen.MousePointer = 11
    
    With AllergyFile
        .CancelError = True
        .Filename = gAssayNM.OrderPath & "\" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".csv"
    
        If Len(Dir(.Filename)) Then Kill .Filename

        lngFIleNum = FreeFile

        Open .Filename For Append As #lngFIleNum
    
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
    
        strInFo = "RowIndexNumber,Test,Sample_ID,Patient_ID,Patient_Last_Name,Patient_First_Name,Patient_Title,Patient_Sex,Patient_Date_of_Birth,Patient_Street,Patient_Zip_Code,Patient_City,Patient_Country,Patient_Phone,Patient_Fax,Patient_email,Company_Name,Company_Street,Company_Zip_Code,Company_City,Company_Trade_Register,Company_Country,Company_Legal_Form,Company_Tax_ID,Company_Phone,Company_Fax,Company_email,Sample_Date_Of_Receipt,Sample_Source,Sample_Type,Type_of_strip,Control,SubstanceFamily,Tray_ID,Well_no,Connected_with,Company_Website,Sample_Date_Of_Sampling,Custom_ID,Comments,Custom0,Custom1,Custom2,Custom3,Custom4,Custom5,Custom6,Custom7,Custom8,Custom9,Membrane,Patient_Nationality" & vbCr
        '1,MEDIWISS / RIDA Panel 1 KO / Rev. 006,1,4,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
        '2,MEDIWISS / RIDA Panel 2 KO / Rev. 004,2,5,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
        '3,MEDIWISS / RIDA Panel 1 KO / Rev. 006,1,4,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
        '4,MEDIWISS / RIDA Panel 2 KO / Rev. 004,2,5,,,,,,,,,,,,,,,,,,,,,,,,,,,PatientStrip,,IgE,,0,,,,,,,,,,,,,,,,,
    
        intCnt = 1
        With vasID
            For intRow = 1 To vasID.DataRowCnt
                .Row = intRow
                .Col = colCheckBox
                If .Value = "1" Then
                    '.GetText colBARCODE, intRow, varTmp: strBarNo = varTmp
                    .GetText colCHARTNO, intRow, varTmp: strBarNo = varTmp
                    .GetText colPID, intRow, varTmp:     strLabNo = varTmp
                    .GetText colPNAME, intRow, varTmp:   strFNm = varTmp
                    
                    strFNm = Han2Eng.HanToEng(strFNm)

                    .GetText colPSEX, intRow, varTmp
                    If Len(varTmp) = 2 Then
                        strSex = Mid(varTmp, 2, 1)
                    Else
                        strSex = varTmp
                    End If
                    
                    If strSex = "M" Then
                        strSex = "Male"
                    Else
                        strSex = "FeMale"
                    End If
                    .GetText colPAGE, intRow, varTmp:   strAge = varTmp
                    strAge = Han2Eng.HanToEng(strAge)
                    .GetText colASSAYNM, intRow, varTmp
                                        
                    '1.     POBALL™ Food Advanced Test
                    '2.     POBALL™ Food Intensive Test
                    '3.     POBALL™ Premium Test
                    '4.     POBALL™ Inhalant Advanced Test
                    '5.     POBALL™ Premium Intensive Test

                    strLabNo = strFNm
                    
                    If varTmp = "POBALL™ Premium Test(127종)" Then
                            strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.IN1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.IN2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                    ElseIf varTmp = "POBALL™ Premium Intensive Test(127종)" Then
                        If txtSeq.Text = "1" Then
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.IN1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.IN2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        ElseIf txtSeq.Text = "2" Then
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        End If
                    ElseIf varTmp = "POBALL™ Basic Test(54종)" Then
                            strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                    ElseIf varTmp = "POBALL™ Food Advanced Test" Then
                            strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                    ElseIf varTmp = "POBALL™ Food Intensive Test(108종)" Then
                        If txtSeq.Text = "1" Then
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            'strInFo = strInFo & CStr(intCnt + 2) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            'strInFo = strInFo & CStr(intCnt + 3) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        ElseIf txtSeq.Text = "2" Then
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        End If
                    ElseIf varTmp = "POBALL™ Inhalant Advanced Test" Then
                            strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.IN1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.IN2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                    ElseIf varTmp = "POBALL™ Health Care Test" Then
                        If txtSeq.Text = "1" Then
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        ElseIf txtSeq.Text = "2" Then
                                strInFo = strInFo & CStr(intCnt) & "," & gAssayNM.FD1 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                            strInFo = strInFo & CStr(intCnt + 1) & "," & gAssayNM.FD2 & "," & strBarNo & "," & strLabNo & "," & strLNm & "," & strFNm & ",," & strSex & "," & strAge & vbCr
                        End If
                    End If
                    
                    intCnt = intCnt + 2
                
                    .SetText colCheckBox, intRow, "0"
                    .SetText colState, intRow, "오더전송"
                    SetBackColor vasID, intRow, intRow, colState, colState, 202, 255, 112
                End If
            Next
        End With
        '-- 끝에 줄바꿈 제거
        If strInFo <> "" Then
            strInFo = Mid(strInFo, 1, Len(strInFo) - 2)
        End If
        Print #lngFIleNum, strInFo
        Close #lngFIleNum
    End With
    
    If intCnt > 1 Then
        MsgBox "워크리스트 생성이 완료되었습니다", vbInformation + vbOKOnly, Me.Caption
    End If
    Screen.MousePointer = 0
    
    
    
End Sub


'''Private Sub MakeXML(ByVal vasIDRow As Integer)
''''선택전송
'''    'Dim vasIDRow As Integer
'''    Dim vasResRow As Integer
'''    Dim iRow As Integer
'''    Dim liRet As Integer
'''    Dim FindFile As String
'''
''''    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
''''        Exit Sub
''''    End If
'''
'''    With frmInterface
'''        For vasResRow = 1 To .vasTemp.DataRowCnt
'''            .vasID.Row = vasIDRow
'''            .vasID.Col = 1
'''            If .vasID.Value = 1 Then
'''                .vasTemp.Row = vasResRow
'''                liRet = -1
'''                If Trim(GetText(.vasTemp, vasResRow, 3)) <> "" Then
'''                    liRet = Make_XML(vasResRow)
'''                End If
'''
'''                If liRet = 1 Then
'''                    SetBackColor .vasID, vasIDRow, vasIDRow, colCheckBox, 12, 202, 255, 112
'''                    'SetText vasList, "전송완료", vasIDRow, colState
'''
'''                    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
'''                    If FindFile <> "" Then
'''                        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '전송완료가 됐을때 파일지우기
'''                    End If
'''
'''                          SQL = " Update pat_res Set "
'''                    SQL = SQL & " TransYN = '2', "
'''                    SQL = SQL & " TransDt = '" & Format(Now, "yyyymmdd") & "' "
'''                    .vasID.Row = vasIDRow: .vasID.Col = 4
'''                    SQL = SQL & " Where ChartNo  = '" & Trim(.vasID.Text) & "' "
'''                    .vasID.Row = vasIDRow: .vasID.Col = 12
'''                    SQL = SQL & "   and ExamID   = '" & Trim(.vasID.Text) & "' "
'''                    .vasID.Row = vasIDRow: .vasID.Col = 10
'''                    SQL = SQL & "   and CommDate = '" & Trim(.vasID.Text) & "'"
'''                    Res = SendQuery(gLocal, SQL)
'''
'''                Else
'''                    SetBackColor .vasID, vasIDRow, vasIDRow, colCheckBox, 12, 255, 0, 0
'''                    'SetText vasID, "실패", vasIDRow, colState
'''                End If
'''                '.vasID.Col = 1
'''                '.vasID.Value = "0"
'''            Else
'''
'''            End If
'''        Next
'''    End With
'''
'''    If XmlTxtHead = "" Then
'''        '<?xml version="1.0" encoding="utf-8"?>
'''        '<GROUP NAME="WORKLIST" TYPE="HOST" VERSION="00.00.00.00">
'''
'''        XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
'''                     "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare검사정보>"
'''    End If
'''
'''    If XmlTxtTail = "" Then
'''        XmlTxtTail = "</UBCare검사정보>"
'''    End If
'''
''''    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
'''    SaveXMLFile XMLAllTxt
'''
'''End Sub


'Public Sub SaveXMLFile(argSQL As String, Optional argFlag As Integer = 0)
''argSQL의 내용을 파일로 저장
'    Dim FilNum, FilNum1
'    Dim FindFile As String
'    Dim TxtString1 As String
'    Dim AllString1 As String
'    Dim i As Long
'
'    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out.xml")
'
'
'    If FindFile <> "" Then
''        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
'        FilNum1 = FreeFile
'        Open "C:\UBCare\SINAI\IF\ExamIF_out.xml" For Input As FilNum1
'
'        Do While Not EOF(FilNum1)
'            Input #FilNum1, TxtString1
'            Line Input #FilNum1, TxtString1
'            AllString1 = AllString1 & TxtString1
'        Loop
'
'        Close #FilNum1
'        i = InStr(1, AllString1, "</UBCare검사정보>")
'        XmlBody = Mid(AllString1, 1, i - 1)
'        argSQL = XmlBody & argSQL & XmlTxtTail
'        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
'    Else
'        argSQL = XmlTxtHead & argSQL & XmlTxtTail
'    End If
'
''    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
'
'    FilNum = FreeFile
'
'
'    If argFlag = 0 Then
'        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Output As FilNum
'    Else
'        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Append As FilNum
'    End If
'    Print #FilNum, argSQL
'    Close FilNum
'    argSQL = ""
'End Sub

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


Private Sub cmdResult_Click()
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
    
    Dim iCnt As Integer
    
    
    Screen.MousePointer = 11
    
'    StatusBar1.Panels(3).Text = ""
    iCnt = 0
    
    If OpenExcelCSV = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    
End Sub

Public Function ReadFileBinary(ByVal strFileName As String) As String

On Error GoTo errHandler
    Dim fsT, tFilePath As String

    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")

    'Specify stream type - we want To save text/string data.
    fsT.Type = 2
 
    'Specify charset For the source text data.
    fsT.Charset = "utf-8"
 
    'Open the stream And write binary data To the object
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


Private Function OpenExcelCSV() As Boolean
    Dim strFile As String
    Dim i, iCnt As Integer
    Dim strTemp As String
    Dim varTmp  As Variant
    Dim xlApp As New Excel.Application
    Dim xlSheet As Excel.Worksheet
    Dim strPath As String
    Dim strDestFile As String
    Dim varDestNm   As Variant


    Dim strBuffer As String
    Dim strBuf    As String
    Dim lngBufLen As Long
    Dim BufChar   As String
'    Dim iCnt      As Long

    OpenExcelCSV = False

    AllergyFile.DialogTitle = "엑셀파일 열기"
    'AllergyFile.Filter = "Excel Files (*.xlsx)|*.xls|All Files (*.*)|*.*"
    AllergyFile.InitDir = gAssayNM.ResultPath
    AllergyFile.ShowOpen
    iCnt = 0
    strBuffer = ""


    If Len(AllergyFile.Filename) > 0 Then
        Open AllergyFile.Filename For Input As #3

        strBuffer = ""
'        Do While Not EOF(3)
'            strBuf = strBuf & Input(1, #3)
'        Loop

        strBuf = ReadFileBinary(AllergyFile.Filename)
        
        Close #3
        
        lngBufLen = Len(strBuf)


        For i = 1 To lngBufLen
            BufChar = Mid$(strBuf, i, 1)
            Select Case BufChar
                Case LF
                    iCnt = iCnt + 1
                    ReDim Preserve strRecvData(iCnt)
                    strRecvData(iCnt) = strBuffer
                    strBuffer = ""

                Case Else
                    strBuffer = strBuffer & BufChar
            End Select
        Next i
        
        Call EditRcvDataINPROVE_CSV
        
    Else
        Exit Function
    End If

    Exit Function

Rst:

End Function



Private Function OpenExcel() As Boolean
    Dim strFile As String
    Dim i, iCnt As Integer
    Dim strTemp As String
    Dim varTmp  As Variant
    Dim xlApp As New Excel.Application
    Dim xlSheet As Excel.Worksheet
    Dim strPath As String
    Dim strDestFile As String
    
    OpenExcel = False

    AllergyFile.DialogTitle = "엑셀파일 열기"
    'AllergyFile.Filter = "Excel Files (*.xlsx)|*.xls|All Files (*.*)|*.*"
    AllergyFile.InitDir = gAssayNM.ResultPath
    AllergyFile.ShowOpen
    
    
    
    If Len(AllergyFile.Filename) > 0 Then
        xlApp.Workbooks.Open AllergyFile.Filename
        strPath = AllergyFile.Filename
    Else
        Exit Function
    End If
    Set xlSheet = xlApp.Worksheets("sheet1")

    With vasExcel
        .Action = ActionClear
        For iCnt = 1 To .MaxRows
            For i = 1 To .MaxCols
                'If xlSheet.Cells(iCnt, i) <> "" Then
                If Trim(Format(xlSheet.Cells(iCnt, 1), "####-##")) = "" Then
                    xlApp.Workbooks.Close
                    xlApp.Quit

                    Set xlSheet = Nothing
                    OpenExcel = True
                    GoTo Rst
                End If
                
                vasExcel.SetText i, iCnt, Trim(xlSheet.Cells(iCnt, i))
            Next
        Next iCnt
    End With

Rst:
    xlApp.Quit
    
    Set xlSheet = Nothing

End Function

Private Sub Open_CSV(Path As String)

    Dim DB As ADODB.Connection
    Dim connCSV As ADODB.Connection ' CSV 커넥션
    Dim commCSV As ADODB.Command    ' CSV 명령어
    Dim rsCSV As ADODB.Recordset    ' CSV 레코드셋

    Dim PNAME  As String
    
    
'    blsRsresult.Visible = False
'    lblCount = ""

    
    Set connCSV = New ADODB.Connection
    Set commCSV = New ADODB.Command
    Set rsCSV = New ADODB.Recordset
'
'
'Private Sub CSV_Link()
'
'Dim cn As ADODB.Connection
'Dim rs As ADODB.Recordset
'Dim mySql As String
'
'Set cn = New ADODB.Connection
'cn.ConnectionString = "Previder=MSDASQL;DNS=cnText"
'cn.Open
'
'Set rs = New ADODB.Recordset
'mySql = "Select * From [Test.csv]"
'rs.Open mySql, cn, sdOpenStatic
'
'Set DataGrid1.DataSource = rs
'
'End Sub
'
'
'
'
'
' ------------------------------------------------------------------------
'
'Run-time error '-214746259(80004005)':
'
'
'
'[Microsoft][ODBC 드라이버관리자] 원본 이름이 없고 기본 드라이버를
'
'지정하지 않았습니다.
'
'------------------------------------------------------------------------------------
'
'하지만 ODBC 에는 cnText 을 *.txt, *.csv 형식으로 이미 만든 상태 입니다.
'
'
'
'어떻게 해결하죠???

    
    
    
    
    On Error GoTo Err

    Path = "C:\프로젝트\PnV\Allergen\Export\"
    PNAME = "12345.csv"
    
    
    
'connCSV.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" _
'                        "Data Source=Path & pname &
'      + 'Extended Properties="Excel 12.0 xml;HDR=YES";'

connCSV.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & Path & PNAME & " ;" & _
               "Extended Properties=""Excel 8.0;HDR=NO;"""

 
 connCSV.Open

  commCSV.CommandText = "Select * from [Sheet1$];"
  Set rsCSV = commCSV.Execute

 
    
    
    
    
    
    
    
    
    
    
    connCSV.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0; " _
                            & " DSN=csvText;"

    
'    connCSV.ConnectionString = "Previder=MSDASQL;DSN=csvText"
    
    connCSV.Open

    

    commCSV.ActiveConnection = connCSV

    '----------------------------------------------

    ' 스프레드의 첫번째 행은 체크로 정해져 있으므로

    ' select 1, * 로 검색한다.

    '----------------------------------------------

    commCSV.CommandText = "SELECT 1,* FROM " & PNAME    'Authors.csv

   ' rsCSV.Open commCSV
    
    Set rsCSV = commCSV.Execute

   ' cmdSQL.CommandText = argSQL
   ' Set RS = cmdSQL.Execute
  
    
          
          
          
    If rsCSV.RecordCount = 0 Then

        MsgBox "데이터가 없습니다."

        rsCSV.Close

        Set rsCSV = Nothing

        Set commCSV = Nothing

        Set connCSV = Nothing

 

        Exit Sub

    End If

    Set vasExcel.DataSource = rsCSV

    vasExcel.MaxRows = rsCSV.RecordCount

    vasExcel.MaxCols = vasExcel.DataColCnt

    Set vasExcel.DataSource = Nothing

    

    rsCSV.Close

        

    Set rsCSV = Nothing

    Set commCSV = Nothing

    Set connCSV = Nothing

            
'
'    blsBot.Caption = xlfile & " ☞ " & fName
'
'
'
'    ' 스프레드 열이름 및 정리대상열 콤보박스 초기화
'
'    Call SpreadSetting
'
'
'
'    cboColno.ListIndex = 0
'
'    cboFilterDiv.ListIndex = 0
'
'
'
'    aRow = 1: aCol = 1
'
'    fpSpread1.Col = 1
'
'    fpSpread1.Row = 1

    

    Exit Sub

Err:

    Set rsCSV = Nothing

    Set commCSV = Nothing

    Set connCSV = Nothing

    MsgBox Err.Number & " " & Err.Description, vbCritical

End Sub

Private Function XMLFileOpen(strPath) As String
    Dim myXml As New MSXML2.DOMDocument60
    Dim node1 As IXMLDOMNode
    Dim node2 As IXMLDOMNode
    Dim strMSG As String
    Dim objElem As IXMLDOMNodeList
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim l       As Integer
    Dim m       As Integer
    Dim n       As Integer
    Dim blnEdit As Boolean
    
    
   ' On Error Resume Next
    
    n = 0
    myXml.async = False
    If myXml.Load(strPath) = True Then
        Set objElem = myXml.selectNodes("GROUP//GROUP")

        For i = 0 To objElem.Length - 1
            For j = 0 To objElem.Item(i).childNodes.Length - 1
                For k = 0 To objElem.Item(i).childNodes(j).childNodes.Length - 1
                    If objElem.Item(i).childNodes(j).Attributes(k).nodeName = "TYPE" Then
                        If objElem.Item(i).childNodes(j).Attributes(k).nodeValue = "Patient" Then
                            blnEdit = True
                        Else
                            blnEdit = False
                        End If
                    End If
                    
                    If blnEdit = True And objElem.Item(i).childNodes(j).Attributes(k).nodeName = "ID" Then
                        ReDim Preserve strRecvData(n)
                        strRecvData(n) = strRecvData(n) & "P|" & objElem.Item(i).childNodes(j).Attributes(k).nodeValue
                        n = n + 1
                        Exit For
                        'Set node1 = myXml.selectSingleNode("GROUP//GROUP//GROUP//GROUP//GROUP//PARAM")
                        'Debug.Print node1.XML
                    End If
                    'Call EditRcvDataAPEX
                    
                Next
                
            Next
        Next
        Set myXml = Nothing
    Else
      MsgBox "읽기에러", vbCritical
    End If
    
    For i = 0 To UBound(strRecvData)
        Debug.Print strRecvData(i)
    Next
    
End Function


Private Function f_subSet_XMLWorkList(ByVal strXML As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    Dim varTmp  As Variant
    Dim j       As Integer
    Dim k As Integer
    
    Dim blnAppend1  As Boolean
    Dim blnAppend2  As Boolean
    Dim varTmp1    As Variant
    Dim strTest   As String
    Dim strResult As String
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    j = 0
    blnAppend1 = False
    blnAppend2 = False
    
    '-- 오더파일명과 경로를 지정한다.
    strPath = strXML

    
    '1라인씩 가져오기 MSDN내용
    Dim TextLine
    Open strPath For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' 파일을 닫습니다
 
    intIdx = 0
    lngBufLen = Len(strBuffer)
        

    
    'strBuffer = Replace(strBuffer, Chr(9), "")
    varTmp = Split(strBuffer, "</GROUP>")

    Erase strRecvData

'    blnSameRecord = True
    
    For i = 0 To UBound(varTmp)
        'Debug.Print varTmp(i)
        If InStr(varTmp(i), """Patient""") > 0 Then 'blnAppend1 = False And
            strTmp = Mid(varTmp(i), InStr(varTmp(i), """Patient""") + 14)
            ReDim Preserve strRecvData(j)
            strRecvData(j) = strRecvData(j) & "P|" & mGetP(strTmp, 1, """")
            j = j + 1
            blnAppend1 = True
        End If
        
        If InStr(varTmp(i), """Assay""") > 0 Then 'blnAppend1 = True And
            strTmp = Mid(varTmp(i), InStr(varTmp(i), """Assay""") + 12)
            ReDim Preserve strRecvData(j)
            
'            If gAssayNM.INHALANT = mGetP(strTmp, 1, """") Then
'                strRecvData(j) = strRecvData(j) & "O|INHALANT"
'            ElseIf gAssayNM.FOOD = mGetP(strTmp, 1, """") Then
'                strRecvData(j) = strRecvData(j) & "O|FOOD"
'            ElseIf gAssayNM.ATOPY = mGetP(strTmp, 1, """") Then
'                strRecvData(j) = strRecvData(j) & "O|ATOPY"
'            Else
'                f_subSet_XMLWorkList = ""
'                Exit Function
'            End If
            
            j = j + 1
            blnAppend2 = True
            
            'varTmp1 = Split(varTmp(i), "</PARAM>")
        End If
        
        
        If InStr(varTmp(i), """Blot""") > 0 Then 'blnAppend1 = True And blnAppend2 = True And
            varTmp1 = Split(varTmp(i), "</PARAM>")
            strTest = ""
            strResult = ""
            For k = 0 To UBound(varTmp1)
                If InStr(varTmp1(k), """Code""") > 0 Then
                    strTest = Mid(varTmp1(k), InStr(varTmp1(k), """Code""") + 7)
                End If
                
                If strTest <> "" And InStr(varTmp1(k), """QntResult""") > 0 Then
                    strResult = Mid(varTmp1(k), InStr(varTmp1(k), """QntResult""") + 12)
                    
                    strResult = Replace(strResult, "&lt;", "<")
                    strResult = Replace(strResult, "&gt;", ">")
                    ReDim Preserve strRecvData(j)
                    strRecvData(j) = strRecvData(j) & "R|" & strTest & "^" & strResult
                    j = j + 1
                    strTest = ""
                    strResult = ""
                End If
            Next
        End If
    Next
    
    Screen.MousePointer = 0

    Exit Function
        
ErrorTrap:
    
'    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function


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
'    StatusBar1.Panels(3).Text = ""
          
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG, HOSPITAL " & vbCrLf
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
                    
                    'If Trim(RS("접수일자")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("바코드번호")) = strChart Then
                    If Trim(RS("EXAMDATE")) = GetText(vasID, i, colEXAMDATE) And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("바코드번호")) = strChart Then
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
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colBREED
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colDOB
                    SetText vasID, Trim(RS.Fields("HOSPITAL")) & "", .MaxRows, colASSAYNM
                    
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
    Else
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
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
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
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
                    
                    SetText vasID, Trim(RS.Fields("R")) & "", .MaxRows, colBREED
                    SetText vasID, Trim(RS.Fields("P")) & "", .MaxRows, colASSAYNM

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
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
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
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
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
    SQL = SQL & "   FROM DRBITPACK..RESINF AS R, DRBITPACK..OCMINF AS O, DRBITPACK..PBSINF AS P, DRBITPACK..LABMST AS E, DRBITPACK..ODRINF AS W" & vbCrLf
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
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub GetWorkList_NTL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    Dim sqlRet      As Integer
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
'    StatusBar1.Panels(3).Text = ""
    
    '   @StatustIndex         tinyint,          '+#13#10+    // 필수입력: 0-전체, 1-미완료, 2-완료
    '   @WorkListCode         varchar(50),      '+#13#10+    // 필수입력: 워크리스트코드
    '   @BeginDate         smalldatetime,       '+#13#10+    // 필수입력: 조회일-시작
    '   @EndDate         smalldatetime,         '+#13#10+    // 필수입력: 조회일-끝
    '   @BeginNo         int,                   '+#13#10+    // 선택입력: 접수번호 시작 (기본값 : 0)
    '   @EndNo         int,                     '+#13#10+    // 선택입력: 접수번호 종료 (기본값 : 0)
    '   @TestCodes         varchar(200)         '+#13#10+    // 선택입력: 검사코드

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute("Exec interface_GetPatientResultList02 '" & cboChk.ListIndex & "','" & gWKCD & "','" & Format(dtpStartDt.Value, "yyyy-mm-dd") & "','" & Format(dtpStopDt.Value, "yyyy-mm-dd") & "'," & Val(txtStartNum.Text) & "," & Val(txtStopNum.Text) & ",''", sqlRet)
    
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
                    strChart = GetText(vasID, i, colPID)
                    If Trim(RS("LabRegDate")) = strDate And Format(Trim(RS.Fields("LabRegDate")), "yymmdd") & PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0") = strChart Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("LabRegDate")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Format(Trim(RS.Fields("LabRegDate")), "yymmdd") & PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0"), .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("PatientChartNo")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("LabRegNo")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("PatientName")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("CompanyCode")) & "", .MaxRows, colBREED
                    SetText vasID, Trim(RS.Fields("PatientBirthDay")) & "", .MaxRows, colASSAYNM    '-- 생년월일
                    SetText vasID, Trim(RS.Fields("PatientSex")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("PatientAge")) & "", .MaxRows, colPAGE
                    Select Case Trim(RS.Fields("OrderCode")) & ""
                        Case "62800":   SetText vasID, "INHALANT", .MaxRows, colDOB
                        Case "62700":   SetText vasID, "FOOD", .MaxRows, colDOB
                        Case "62500":   SetText vasID, "ATOPY", .MaxRows, colDOB
                    End Select
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
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub GetWorkList_PNV(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    Dim sqlRet      As Integer
    
    Dim sURL        As String
    Dim sHeader     As String
    Dim sRcvData    As String
    Dim sBody       As String
    Dim varRcvData  As Variant
    Dim varPetData  As Variant
    Dim strPetData  As String
    
    Dim strAssayNm  As String
    
'On Error Resume Next

    '-- 오더조회
    '워크리스트 :
    '   URL : /api/selectLabList
    '   param1 = startIndex : int
    '   param2 = cutCount : int
    

    '결과저장 :
    '   URL : /api/insertLabResult
    '   param1 = id             : int
    '   param2 = LabResultList  : array
    '   param3 = SerialNo       : String    검사결과코드(B1, B2 ….)
    '   param4 = Type           : Int       결과        (1 : 음성, 2 : 양성)
    '   param5 = Value          : String    결과 값

    
    sURL = PnVAPI.APIURL & PnVAPI.APIOrdPath
    'sURL = "https://dev.pnv.co.kr/PnV_Lab/api/selectLabList"
    
              sHeader = "X-LAB-SECURITY: AGASGBggFASVfg42ASFV5255GGSAVNJJKPQOWDKVM4fiFHDoWFmqSGHYASDksapqmdm2DASFASfyomsFASGAS==" & vbCrLf
    sHeader = sHeader & "X-LAB-Client: allergy" & vbCrLf
    sHeader = sHeader & "X-LAB-MACHINE: TEST001" & vbCrLf
    sHeader = sHeader & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    
    If txtCnt.Text = "" Then
        txtCnt.Text = "100"
    End If
    
    sBody = "startIndex=1&cutCount=" & txtCnt.Text
    

    sRcvData = OpenURLWithIE2(sURL, sHeader, sBody, Inet1)
    
    SetRawData "[sRcvData ]" & sRcvData
    
'    Debug.Print sRcvData
    sRcvData = Replace(sRcvData, """", "")
    sRcvData = Replace(sRcvData, "}", "")
    sRcvData = Replace(sRcvData, "]", "")
    varRcvData = Split(sRcvData, "{")
    
    vasID.MaxRows = 0
    intRow = 0
    
    If varRcvData(1) = "" Then
        vasID.MaxRows = 0
        intRow = 0
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        Exit Sub
    End If
    
    
    'vasID.MaxRows = UBound(varRcvData) - 1
    
    blnSame = False
    vasID.ReDraw = False
'    StatusBar1.Panels(3).Text = ""
        
        For iCnt = 2 To UBound(varRcvData)
            '0 : id:104,
            '1 : petGender:0,
            '2 : petBirthDay:2016-02-01,
            '3 : petSpecies:0,      '종(개/고양이 ??)
            '4 : number:201610088,  '접수번호(=챠트번호)
            '5 : petName:test,
            '6 : date:2016-03-15,
            '7 : petBreed:11        '품종(푸들/시베리안허스키 ??)
            '8 : petNumber:124      '등록번호},
            
            '0 : id:102,
            '1 : petGender:2,
            '2 : petBirthDay:2016-03-01,
            '3 : petSpecies:0,
            '4 : name:POBALL Food,
            '5 : number:201610086,
            '6 : petName:111,
            '7 : date:2016-03-16,'
            '8 : petBreed:111,
            '9 : petNumber:,


            'err_code    String  에러 코드
            'err_msg String  에러 메시지
            'totalCount  Int 총 카운트 수
            'startIndex  Int 리스트 인덱스
            'cutCount    Int 리스트에 나타낼 수
            'labList Array   워크 리스트
            
            'id  Int 검사 아이디
            'number  String  접수번호
            'date    String  신청일 yyyy-MM-dd
            'petName String  반려동물 이름
            'petBirthDay String  반려동물 생년월일 yyyy-MM-dd
            'petNumber   String  반려동물 등록번호
            'petGender   Int 반려동물 성별(0 : F, 1 : NF, 2 : M, 3 : NM, 4 : 모름)
            'petSpecies  Int 반려동물 종(0 : 개, 1 : 고양이)
            'petBreed    String  반려동물 품종
            'name    String  검사명

'Debug.Print varRcvData(iCnt)
            
            varPetData = Split(varRcvData(iCnt), ",")
            If UBound(varPetData) < 8 Then
                Exit Sub
            End If
            
            strAssayNm = Trim(mGetP(varPetData(5), 2, ":"))
            If strAssayNm <> "" And InStr(strAssayNm, "POBALL™") > 0 Then
                With vasID
                    .ReDraw = False
                    For i = 1 To .DataRowCnt
                        strDate = GetText(vasID, i, colHOSPDATE)
                        strChart = GetText(vasID, i, colPID)
                        If Len(mGetP(varPetData(0), 2, ":")) = 5 Then
                            If mGetP(varPetData(8), 2, ":") = strDate And Format(mGetP(varPetData(8), 2, ":"), "yymmdd") & PedLeftStr(mGetP(varPetData(0), 2, ":"), 5, "0") = strChart Then
                                blnSame = True
                            End If
                        Else
                            If mGetP(varPetData(8), 2, ":") = strDate And Format(mGetP(varPetData(8), 2, ":"), "yymmdd") & PedLeftStr(mGetP(varPetData(0), 2, ":"), 6, "0") = strChart Then
                                blnSame = True
                            End If
                        End If
                    Next
                    
                    If blnSame = False Then
                        .MaxRows = .MaxRows + 1
                        SetText vasID, "1", .MaxRows, colCheckBox
                        SetText vasID, Trim(mGetP(varPetData(8), 2, ":")), .MaxRows, colHOSPDATE
                        If Len(mGetP(varPetData(0), 2, ":")) = 5 Then
                            SetText vasID, Format(mGetP(varPetData(8), 2, ":"), "yymmdd") & PedLeftStr(mGetP(varPetData(0), 2, ":"), 5, "0"), .MaxRows, colBARCODE
                        Else
                            SetText vasID, Format(mGetP(varPetData(8), 2, ":"), "yymmdd") & PedLeftStr(mGetP(varPetData(0), 2, ":"), 6, "0"), .MaxRows, colBARCODE
                        End If
                        SetText vasID, Trim(mGetP(varPetData(6), 2, ":")), .MaxRows, colCHARTNO
                        'SetText vasID, Trim(mGetP(varPetData(9), 2, ":")), .MaxRows, colPID
                        SetText vasID, Trim(mGetP(varPetData(0), 2, ":")), .MaxRows, colPID
                        SetText vasID, Trim(mGetP(varPetData(7), 2, ":")), .MaxRows, colPNAME
                        SetText vasID, Trim(mGetP(varPetData(9), 2, ":")), .MaxRows, colBREED       '-- 품종
                        SetText vasID, Trim(mGetP(varPetData(2), 2, ":")), .MaxRows, colDOB         '-- 생년월일
                        SetText vasID, strAssayNm, .MaxRows, colASSAYNM                             '-- 검사구분(명)
                                        
                        Select Case Trim(mGetP(varPetData(1), 2, ":"))
                            Case "1": SetText vasID, "F", .MaxRows, colPSEX
                            Case "2": SetText vasID, "FS", .MaxRows, colPSEX
                            Case "3": SetText vasID, "M", .MaxRows, colPSEX
                            Case "4": SetText vasID, "MN", .MaxRows, colPSEX
                            Case "5": SetText vasID, "모름", .MaxRows, colPSEX
                            
                        End Select
                        
                        Select Case Trim(mGetP(varPetData(4), 2, ":")) & ""
                            Case "1":   SetText vasID, "개", .MaxRows, colPAGE
                            Case "2":   SetText vasID, "고양이", .MaxRows, colPAGE
                        End Select
                    End If
                    
''                    For i = 1 To .DataRowCnt
''                        strDate = GetText(vasID, i, colHOSPDATE)
''                        strChart = GetText(vasID, i, colPID)
''                        If mGetP(varPetData(7), 2, ":") = strDate And Format(mGetP(varPetData(7), 2, ":"), "yymmdd") & PedLeftStr(mGetP(varPetData(0), 2, ":"), 5, "0") = strChart Then 'PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0")
''                            blnSame = True
''                        End If
''                    Next
''
''                    If blnSame = False Then
''                        .MaxRows = .MaxRows + 1
''                        SetText vasID, "1", .MaxRows, colCheckBox
''                        SetText vasID, Trim(mGetP(varPetData(7), 2, ":")), .MaxRows, colHOSPDATE
''                        SetText vasID, Format(mGetP(varPetData(7), 2, ":"), "yymmdd") & PedLeftStr(mGetP(varPetData(0), 2, ":"), 5, "0"), .MaxRows, colBARCODE
''                        SetText vasID, Trim(mGetP(varPetData(5), 2, ":")), .MaxRows, colCHARTNO
''                        'SetText vasID, Trim(mGetP(varPetData(9), 2, ":")), .MaxRows, colPID
''                        SetText vasID, Trim(mGetP(varPetData(0), 2, ":")), .MaxRows, colPID
''                        SetText vasID, Trim(mGetP(varPetData(6), 2, ":")), .MaxRows, colPNAME
''                        SetText vasID, Trim(mGetP(varPetData(8), 2, ":")), .MaxRows, colBREED       '-- 품종
''                        SetText vasID, Trim(mGetP(varPetData(2), 2, ":")), .MaxRows, colDOB         '-- 생년월일
''                        SetText vasID, strAssayNm, .MaxRows, colASSAYNM                             '-- 검사구분(명)
''
''                        Select Case Trim(mGetP(varPetData(1), 2, ":"))
''                            Case "1": SetText vasID, "F", .MaxRows, colPSEX
''                            Case "2": SetText vasID, "FS", .MaxRows, colPSEX
''                            Case "3": SetText vasID, "M", .MaxRows, colPSEX
''                            Case "4": SetText vasID, "MN", .MaxRows, colPSEX
''                            Case "5": SetText vasID, "모름", .MaxRows, colPSEX
''
''                        End Select
''
''                        Select Case Trim(mGetP(varPetData(3), 2, ":")) & ""
''                            Case "1":   SetText vasID, "개", .MaxRows, colPAGE
''                            Case "2":   SetText vasID, "고양이", .MaxRows, colPAGE
''                        End Select
''                    End If
                    blnSame = False
                End With
            End If
            
            '-- 프로그레스바 진행
    '        frmProgress.Xprog.Value = iCnt
    '        DoEvents
            
            'RS.MoveNext
        'Loop
        Next
        chkWAll.Value = "1"
    'Else
    '    StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
    '    chkWAll.Value = "0"
    'End If
    
    'RS.Close
    '-- 프로그레스바 닫기
    'Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
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
                    Case "O": SetText vasID, "외래", intRow, colDOB
                    Case "E": SetText vasID, "응급", intRow, colDOB
                    Case "I": SetText vasID, "입원", intRow, colDOB
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
                        Case "O": SetText vasID, "외래", intRow, colDOB
                        Case "E": SetText vasID, "응급", intRow, colDOB
                        Case "I": SetText vasID, "입원", intRow, colDOB
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
'        Case "BIT":         Call GetWorkList_BIT(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "TWIN":        Call GetWorkList_TWIN(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "DADESOFT":    Call GetWorkList_DADESOFT(Format(dtpStartDt.Value, "yyyy-mm-dd"), Format(dtpStopDt.Value, "yyyy-mm-dd"))
'        Case "GINUSDLL":    Call GetWorkList_GINUSDLL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "GINUSDB":     Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "BITSMALL":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "BITLARGE":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "MEDICHART":   Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'        Case "NTL":         Call GetWorkList_NTL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "PNV":         Call GetWorkList_PNV(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    End Select
    
    vasID.RowHeight(-1) = 12
    vasRes.MaxRows = 0
    
End Sub


Private Sub cmdSL_Click()
    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        vasID.Width = 20655 '19605
    Else
        cmdSL.Caption = "▶"
        vasID.Width = 11985
    End If

End Sub


Private Sub Form_Resize()
    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    'Me.Top = 0

    vasID.Width = Me.ScaleWidth - vasRes.Width - 200
    vasID.Height = Me.ScaleHeight - picHeader.Height - 200
    vasRes.Left = vasID.Left + vasID.Width + 50
    vasRes.Height = Me.ScaleHeight - picHeader.Height - fraPatInfo.Height - 100

    fraPatInfo.Left = vasID.Left + vasID.Width + 50
    fraPatInfo.Height = Me.ScaleHeight - picHeader.Height - vasRes.Height - 100

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
    
On Error GoTo Err

    If App.PrevInstance Then
        End
    End If
    
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
    
    If gIFMode = "Barcode" Then
        'fraBar.Visible = True
        'fraWork.Visible = False
    
        chkMode.Caption = "Barcode"
        MnModeBarcode.Checked = True
        MnModeWorkList.Checked = False
        chkBar.Value = 1
    Else
        'fraBar.Visible = False
'        fraWork.Visible = True
    
        chkMode.Caption = "WorkList"
        MnModeBarcode.Checked = False
        MnModeWorkList.Checked = True
        chkBar.Value = 0
    End If
    
'    If gScreen = "통합" Then
'        cmdSL.Caption = "◀"
'        vasID.Width = 14595
'    Else
'        cmdSL.Caption = "▶"
'        vasID.Width = 7725
'    End If
    
'    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    cboChk.ListIndex = 0
    
'    comEqp.CommPort = gSetup.gPort
'    comEqp.RTSEnable = gSetup.gRTSEnable
'    comEqp.DTREnable = gSetup.gDTREnable
'    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
'
'    If comEqp.PortOpen = False Then
'        comEqp.PortOpen = True
'    End If
'
'    If comEqp.PortOpen Then
'        frmInterface.StatusBar1.Panels(2).Text = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
'        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'    Else
'        frmInterface.StatusBar1.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
'        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'    End If

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
'    -- osw 추가
'    For i = 1 To 1
'        If Not Connect_PRServer Then
'            MsgBox "연결되지 않았습니다."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
            cn_Server_Flag = True
'        End If
'    Next
    
    '-- osw 추가
'    For i = 1 To 1
'        If Not Connect_DRServer Then
'            MsgBox "연결되지 않았습니다."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
            cn_Server_Flag = True
'        End If
'    Next
    
    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -999), "yyyymmdd")
    
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
    
   ' Call cmdSL_Click
    
    '-- test
'    vasID.MaxRows = 10
    
    FileB4C.Path = gAssayNM.ResultPath & "\"
    
    Exit Sub

Err:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
    
    '경로에러
    If Err.Number = 76 Then
        Resume Next
    Else
        End
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
            .TypeEditCharSet = TypeEditCharSetAlphanumeric
            .TypeEditCharCase = TypeEditCharCaseSetUpper
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 6
        Next
    End With
    
End Sub


Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno, gubun, examtype " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by gubun, seqno * 10 "
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 9)
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
        For j = 1 To 7
            'Debug.Print Trim(GetText(vasCode, i, j))
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
    MsgBox "통신포트 사용 안함", vbOKOnly + vbInformation, Me.Caption
    'frmConfig.Show
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
    
    '-- ASTM TYPE별 Define 해야함.
    '-- ASTM TYPE = Standard
    Select Case intSndPhase
        Case 1  '## Header
            'strOutput = intFrameNo & "H|\^&||| XN-10^00-14^15097^^^^AP795756||||||||E1394-97" & vbCr & ETX
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||||^^|||U|||||^||||||||||||^^^" & vbCr & ETX
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            
        Case 4  '## Order
            If mOrder.NoOrder = True Then
                    
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## 최초 보낼때
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                Else                        '## 남은 문자열이 있을때
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                End If
                
            End If
            
            intFrameNo = intFrameNo + 1
            
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = EOT '""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
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

Private Sub tmrReceive_Timer()
    
    'imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    'tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    'imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    'tmrSend.Enabled = False

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
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

'    If txtTest.Text <> "" And strBuffer = txtTest.Text Then
'        Buffer = strBuffer
'        GoTo Rst
'    End If
    
    Select Case comEqp.CommEvent
        Case comEvReceive

'            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
'            If tmrReceive.Enabled = False Then
'                tmrReceive.Enabled = True
'            Else
'                tmrReceive.Enabled = False
'                tmrReceive.Enabled = True
'            End If

            Buffer = comEqp.Input
'Rst:
            SetRawData "[Rx]" & Buffer
'            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
            'Debug.Print Buffer

            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case intPhase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                intBufCnt = 1
                                Erase strRecvData
                                ReDim Preserve strRecvData(intBufCnt)
                                intPhase = 2
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                            Case ACK
                                '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                                '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")
                                
                                DoEvents
                                
                                If strState = "Q" Then Call SendOrder
                        
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                            Case STX
                                intBufCnt = 1
                                Erase strRecvData
                                ReDim Preserve strRecvData(intBufCnt)
                            Case ETB
                                blnIsETB = True
                                intPhase = 3
                            Case ETX
                                intBufCnt = intBufCnt + 1
                                ReDim Preserve strRecvData(intBufCnt)
                                intPhase = 3
                            Case vbCr, vbLf
                            Case EOT
                                intPhase = 1
                            Case Else
                                If blnIsETB = False Then
                                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                                Else
                                    blnIsETB = False
                                End If
                        End Select
                    Case 3      '## Transfer Phase
                        Select Case BufChar
                            Case vbCr
                            Case vbLf
                                intPhase = 4
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
                            Case EOT
                                '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                                '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")

                                DoEvents
                                
                                Call EditRcvDataASTM
                                
                                If strState = "Q" Then
                                    intSndPhase = 1
                                    intFrameNo = 1
                                    comEqp.Output = ENQ
                                    SetRawData "[Tx]" & ENQ
                                End If
                                
                                intPhase = 1
                        End Select
                End Select
            Next i
            
        Case comEvSend
'            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
'            If tmrSend.Enabled = False Then
'                tmrSend.Enabled = True
'            Else
'                tmrSend.Enabled = False
'                tmrSend.Enabled = True
'            End If
        
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
    Call SetText(vasID, mOrder.RackNo, intRow, colBREED)       '-- Rack
    Call SetText(vasID, mOrder.TubePos, intRow, colASSAYNM)       '-- Pos
    
    '-- 환자정보 표시
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 가져오기
    Call GetSampleInfoW(intRow)
    
    '-- 바코드번호에 해당하는 검사코드 가져오기
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)

    '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
    strItems = GetGetEquipExamCode_XN1000(gEquip, pBarNo, intRow)

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
    
    For i = 1 To vasID.DataRowCnt
        'If Trim(GetText(vasID, i, colBARCODE)) = pBarNo And UCase(GetText(vasID, i, colASSAYNM)) = UCase(mResult.MnmCd) Then
        If Trim(GetText(vasID, i, colCHARTNO)) = pBarNo And UCase(GetText(vasID, i, colASSAYNM)) = UCase(mResult.MnmCd) Then
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
    Call SetText(vasID, "1", intRow, colCheckBox)
    Call SetText(vasID, pBarNo, intRow, colBARCODE)
'    Call SetText(vasID, pBarNo, intRow, colCHARTNO)
    
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    Call SetText(vasID, mResult.MnmCd, intRow, colASSAYNM)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 서버테이블에서 가져와 표시(for 워크리스트)  '6,7,8,9
    'Call GetSampleInfoW_NTL(intRow)
    
    '-- 바코드번호에 존재하는 검사코드 가져오기(인수 : 장비코드,챠트번호,접수일,내원번호,검진번호)
    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    
    '-- 현재 Row
    gRow = intRow
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
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
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
        
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "O"    '## Order
                strBarNo = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^"))
                'strRackNo = mGetP(strTemp1, 1, "^")
                'strTubePos = mGetP(strTemp1, 2, "^")
                
                If strBarNo = "" Then Exit Sub
                
                With mResult
                    .BarNo = strBarNo
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                End With
                                
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
                '-- 오른쪽 결과화면 초기화
                vasRes.MaxRows = 0
                
            Case "R"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                strResult = Trim(mGetP(strRcvBuf, 4, "|"))
'                If InStr(strTemp2, "^") > 0 Then
'                    '## 정성결과 저장
'                    strResult = mGetP(strTemp2, 2, "^")
'                Else
'                    '## 정량결과 저장
'                    strResult = strTemp2
'                End If
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 14
                
            Case "C"    '## Comment
                '## Abnormal 결과일때 Comment 저장
'                If strFlag <> "N" Then
'                    strTemp1 = mGetP(strRcvBuf, 4, "|")
'                    strComm = "[Flag]: " & mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'                End If
                
            Case "L"    '## Terminator
                '-- HCT%
                strIntBase = "HCT%"
                strResult = "%"
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                    End If
                End If
                
                '-- LUC
                strIntBase = "LUC"
                strResult = "0.0"
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                    End If
                End If
                
                '-- Diff
                strIntBase = "Diff"
                strResult = "100"
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                    End If
                End If
                
                
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
    Next

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAPEX()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    
    Dim strFIntBase   As String   '수신한 장비기준 검사명
    Dim strFResult    As String   '수신한 결과(정성)
    
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim strGubun  As String
    
    For intCnt = 0 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Order
                strBarNo = Trim(mGetP(strRcvBuf, 2, "|"))
                'strRackNo = mGetP(strTemp1, 1, "^")
                'strTubePos = mGetP(strTemp1, 2, "^")
                
                'If strBarNo = "" Then Exit Sub
                
            
            Case "O"    '## Order
                strGubun = Trim(mGetP(strRcvBuf, 2, "|"))
                                
                With mResult
                    .BarNo = strBarNo
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .MnmNm = strGubun
                End With
                                
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    'Exit Sub
                End If
                
                strState = "O"
                
                '-- 오른쪽 결과화면 초기화
                vasRes.MaxRows = 0

                
            Case "R"    '## Result
                strFIntBase = ""
                strFResult = ""
                '## 장비기준 검사명, 결과, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                'Debug.Print strIntBase
                strFIntBase = strIntBase
                strFResult = strResult
                
                'code fx1(Fish mix) -> f313(멸치)-1.00, f254(가자미)-0.99, f413(명태)-0.98
                'If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx1" And strResult <> "" and  IsNumeric(strResult)  Then
                If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx1" And strResult <> "" Then
                    For i = 1 To 3
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "f313"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "f254"
                                strResult = strResult * 0.99
                            ElseIf i = 3 Then
                                strIntBase = "f413"
                                strResult = strResult * 0.98
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "f313"
                            ElseIf i = 2 Then
                                strIntBase = "f254"
                            ElseIf i = 3 Then
                                strIntBase = "f413"
                            End If
                        End If
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult

                End If
                
                'code fx4(Nut mix) -> f20(아몬드)-1.00, f253(잣)-0.99, w204(해바라기씨)-0.98
                'If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx4" And strResult <> "" And IsNumeric(strResult) Then
                If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx4" And strResult <> "" Then
                    For i = 1 To 3
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "f20"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "f253"
                                strResult = strResult * 0.99
                            ElseIf i = 3 Then
                                strIntBase = "w204"
                                strResult = strResult * 0.98
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "f20"
                            ElseIf i = 2 Then
                                strIntBase = "f253"
                            ElseIf i = 3 Then
                                strIntBase = "w204"
                            End If
                        End If
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                
                'code w71/e73(생쥐/쥐) -> w71(생쥐)-1.00, e73(쥐)-0.99
                'If strBarNo <> "" And mResult.MnmNm = "INHALANT" And strIntBase = "w71/e73" And strResult <> "" And IsNumeric(strResult) Then
                If strBarNo <> "" And mResult.MnmNm = "INHALANT" And strIntBase = "e71/e73" And strResult <> "" Then
                    For i = 1 To 2
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "e71"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "e73"
                                strResult = strResult * 0.99
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "e71"
                            ElseIf i = 2 Then
                                strIntBase = "e73"
                            End If
                        End If
                        
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 8), lsResRow, colSUBCODE               'subcode
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
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                
                'code gx1(grass mix) -> g1(향기풀)-1.00, g3(오리새)-1.01, g7(갈대)-0.99, g9(외겨이삭)-0.98
                'If strBarNo = "" And mResult.MnmNm = "INHALANT" And strIntBase = "gx1" And strResult <> "" And IsNumeric(strResult) Then
                If strBarNo <> "" And mResult.MnmNm = "INHALANT" And strIntBase = "gx1" And strResult <> "" Then
                    For i = 1 To 4
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "g1"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "g3"
                                strResult = strResult * 1.01
                            ElseIf i = 3 Then
                                strIntBase = "g7"
                                strResult = strResult * 0.99
                            ElseIf i = 4 Then
                                strIntBase = "g9"
                                strResult = strResult * 0.98
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "g1"
                            ElseIf i = 2 Then
                                strIntBase = "g3"
                            ElseIf i = 3 Then
                                strIntBase = "g7"
                            ElseIf i = 4 Then
                                strIntBase = "g9"
                            End If
                        End If
                        
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                                
                'If strResult <> "" And Len(strIntBase) <= 6 Then
                If strResult <> "" And Len(strIntBase) > 0 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                            If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                SetText vasID, strResult, gRow, intCol
                                'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                        SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                    SetText vasID, strResult, gRow, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    'SetText vasRes, gArrEquip(intCol - colState, 8), lsResRow, colSUBCODE               'subcode
                                    SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 10
                
            Case "L"    '## Terminator
                
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
    Next

End Sub



'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataINPROVE()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    
    Dim strFIntBase   As String   '수신한 장비기준 검사명
    Dim strFResult    As String   '수신한 결과(정성)
    
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim strGubun  As String
    Dim varTmp As Variant
    Dim strClass As String
    Dim strAssay As String
    Dim blnSame  As Boolean
    Dim j        As Integer
    
    For i = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(i)
        varTmp = strRcvBuf
        varTmp = Split(varTmp, "|")
        
        For intCnt = 0 To UBound(varTmp) - 1
            If intCnt = 1 Then
                strGubun = varTmp(intCnt)
                If strGubun = gAssayNM.PR1 Then
                    strAssay = "POBALL PRENIUM"
                ElseIf strGubun = gAssayNM.PR2 Then
                    strAssay = "POBALL PRENIUM"
                ElseIf strGubun = gAssayNM.FD1 Then
                    strAssay = "POBALL FOOD"
                ElseIf strGubun = gAssayNM.FD2 Then
                    strAssay = "POBALL FOOD"
                ElseIf strGubun = gAssayNM.FA1 Then
                    strAssay = "POBALL FOOD_ADVANCED"
                ElseIf strGubun = gAssayNM.FA2 Then
                    strAssay = "POBALL FOOD_ADVANCED"
                ElseIf strGubun = gAssayNM.FI1 Then
                    strAssay = "POBALL FOOD_INTENSIVE"
                ElseIf strGubun = gAssayNM.FI2 Then
                    strAssay = "POBALL FOOD_INTENSIVE"
                ElseIf strGubun = gAssayNM.IN1 Then
                    strAssay = "POBALL INHALANT"
                ElseIf strGubun = gAssayNM.IN2 Then
                    strAssay = "POBALL INHALANT"
                ElseIf strGubun = gAssayNM.IA1 Then
                    strAssay = "POBALL INHALANT_ADVANCED"
                ElseIf strGubun = gAssayNM.IA2 Then
                    strAssay = "POBALL INHALANT_ADVANCED"
                Else
                    strAssay = strGubun
                End If
            End If
        
            If intCnt = 8 Then
                strBarNo = varTmp(intCnt)
                'strBarNo = "16031600101"
                                
                blnSame = False
                
                For j = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, j, colBARCODE)) = strBarNo Then
                        If Trim(GetText(vasID, j, colState)) = "" Then
                            blnSame = False
                        Else
                            blnSame = True
                        End If
                        Exit For
                    End If
                Next
                                
                If blnSame = False Then
                    With mResult
                        .BarNo = strBarNo
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                        .MnmNm = strGubun
                    End With
                                    
                    Call SetPatInfo(strBarNo)
                    
                    vasRes.MaxRows = 0
                End If
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
                '-- 오른쪽 결과화면 초기화
               ' vasRes.MaxRows = 0
                
            End If
        
            If intCnt > 8 And (InStr(intCnt, 4) > 0 Or InStr(intCnt, 9) > 0) Then
                strIntBase = varTmp(intCnt)

'                If UCase(strIntBase) = "HX" Then
'                    strIntBase = "HX"
'                End If
'
'                If strIntBase = "Total_IgE" Then
'                    strIntBase = "IgE"
'                End If
                'strIntBase = strGubun & strIntBase
                
                
                strIntResult = varTmp(intCnt + 1)
                strResult = strIntResult
                
                '-- Class 추가
                If strIntBase = "FD|IgE" Or strIntBase = "IN|IgE" Then
                    If strIntResult = "<100" Then
                        strClass = "1.0"
                    ElseIf strIntResult = "100-200" Then
                        strClass = "2.0"
                    ElseIf strIntResult = ">200" Then
                        strClass = "3.0"
                    End If
                Else
                    strClass = varTmp(intCnt + 2)
                    strClass = Format(strClass, "0.0")
                End If

                '-- 결과값에 Class값을 포함
                'strResult = strClass & " Class" & "(" & strResult & ")"

                If strResult <> "" And Len(strIntBase) > 0 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND GUBUN = '" & strAssay & "'"
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
                            If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                SetText vasID, strResult, gRow, intCol
                                'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                        'SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
                                If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                    SetText vasID, strResult, gRow, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    'SetText vasRes, gArrEquip(intCol - colState, 8), lsResRow, colSUBCODE               'subcode
                                    SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
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
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 10
            End If
        Next
        
        strGubun = ""
        strBarNo = ""
        
        '## DB에 결과저장
        If MnTransAuto.Checked = True And strState = "R" Then
'''            '-- 수기입력항목 조회
'''            '장비코드, 검사코드, 검사명, 결과, 순번
'''                  SQL = "SELECT a.EQUIPCODE, a.EXAMCODE, a.EXAMNAME, a.EQUIPRESULT, a.RESULT, a.SEQNO, a.REFFLAG, a.EXAMSUBCODE " & vbCrLf
'''            SQL = SQL & "  FROM PATRESULT a, EQPMASTER b " & vbCrLf
'''            SQL = SQL & " WHERE a.EQUIPNO = '" & gEquip & "'" & vbCrLf
'''            SQL = SQL & "   AND a.EQUIPNO = b.EQUIPNO " & vbCrLf
'''            SQL = SQL & "   AND b.EXAMTYPE = '수기'" & vbCrLf
'''            SQL = SQL & "   AND b.GUBUN = '" & strASSAYNM & "' " & vbCrLf
'''            SQL = SQL & "   AND (a.RESULT = '' OR a.RESULT IS NULL) " & vbCrLf
'''            SQL = SQL & "   AND a.SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
'''            SQL = SQL & "   AND a.BARCODE = '" & lsID & "' " & vbCrLf
'''            SQL = SQL & "   AND a.DISKNO = '" & Trim(GetText(vasID, Row, colDOB)) & "' " & vbCrLf
'''            SQL = SQL & " GROUP BY a.SEQNO, a.EQUIPCODE, a.EXAMCODE, a.EXAMNAME, a.EQUIPRESULT, a.RESULT, a.SEQNO, a.REFFLAG, a.EXAMSUBCODE "
'''            SQL = SQL & " ORDER BY a.SEQNO * 10"
'''
'''            Set RS = cn.Execute(SQL, , 1)
'''            If Not RS.EOF = True And Not RS.BOF = True Then
'''                vasRes.MaxRows = 0
'''                Do Until RS.EOF
'''                    With vasRes
'''                        .MaxRows = .MaxRows + 1
'''                        SetText vasRes, "0", .MaxRows, colCheckBox
'''                        SetText vasRes, Trim(RS.Fields("EQUIPCODE")) & "", .MaxRows, colEQUIPCODE
'''                        SetText vasRes, Trim(RS.Fields("EXAMCODE")) & "", .MaxRows, colEXAMCODE
'''                        SetText vasRes, Trim(RS.Fields("EXAMNAME")) & "", .MaxRows, colEXAMNAME
'''                        SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMachResult
'''                        SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
'''                        SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSeq
'''                        SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colFLAG
'''                        SetText vasRes, Trim(RS.Fields("EXAMSUBCODE")) & "", .MaxRows, colSUBCODE
'''
'''                        If Trim(RS.Fields("REFFLAG")) = "H" Then
'''                            .Row = .MaxRows
'''                            .Col = colRESULT
'''                            .ForeColor = vbRed
'''                        ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
'''                            .Row = .MaxRows
'''                            .Col = colRESULT
'''                            .ForeColor = vbBlue
'''                        End If
'''
'''                    End With
'''                    RS.MoveNext
'''                Loop
'''            End If
'''
'''            RS.Close
        
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
    
    Next

End Sub



'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataINPROVE_CSV()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    
    Dim strFIntBase   As String   '수신한 장비기준 검사명
    Dim strFResult    As String   '수신한 결과(정성)
    
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim strGubun  As String
    Dim varTmp As Variant
    Dim strClass As String
    Dim strAssay As String
    Dim blnSame  As Boolean
    Dim j        As Integer
    Dim strAssayNm As String
    
    For i = 2 To UBound(strRecvData)
        strRcvBuf = strRecvData(i)
        'Debug.Print strRcvBuf
        
        strRcvBuf = Replace(strRcvBuf, "레몬,라임,오렌지", "레몬 라임 오렌지")
        strRcvBuf = Replace(strRcvBuf, "Lemon, Orange, Lime", "Lemon Orange Lime")
        strRcvBuf = Replace(strRcvBuf, "Oak, White", "Oak  White")
        strRcvBuf = Replace(strRcvBuf, "Cedar, japan", "Cedar  japan")
        'strRcvBuf = Replace(strRcvBuf, ",", " ")

        varTmp = strRcvBuf
        varTmp = Split(varTmp, ",")
        
        For intCnt = 0 To UBound(varTmp) - 1
            If intCnt = 1 Then
                strGubun = varTmp(intCnt)
            End If
        
            If intCnt = 8 Then
                strBarNo = varTmp(intCnt)
                'strBarNo = "201610121"
                                
                blnSame = False
                
                For j = 1 To vasID.DataRowCnt
                    'If Trim(GetText(vasID, j, colBARCODE)) = strBarNo Then
                    If Trim(GetText(vasID, j, colCHARTNO)) = strBarNo Then
                        'If Trim(GetText(vasID, j, colASSAYNM)) = strGubun Then
                        
                            If Trim(GetText(vasID, j, colState)) = "" Then
                                blnSame = False
                            Else
                                blnSame = True
                            End If
                            strAssay = Trim(GetText(vasID, j, colASSAYNM))
                            Exit For
                        'End If
                    End If
                Next
'                1.     POBALL™ Food Advanced Test
'                2.     POBALL™ Food Intensive Test
'                3.     POBALL™ Premium Test
'                4.     POBALL™ Inhalant Advanced Test
'                5.     POBALL™ Premium Intensive Test
                                
                Select Case strAssay
                    Case "POBALL™ Premium Test(127종)":            strAssayNm = "PREMIUM"
                    Case "POBALL™ Premium Intensive Test(127종)":  strAssayNm = "PREMIUM_I"
                    Case "POBALL™ Food Advanced Test":             strAssayNm = "FOOD_A"
                    Case "POBALL™ Basic Test(54종)":               strAssayNm = "FOOD_A"
                    Case "POBALL™ Food Intensive Test(108종)":     strAssayNm = "FOOD_I"
                    Case "POBALL™ Inhalant Advanced Test":         strAssayNm = "INHALANT_A"
                    Case "POBALL™ Health Care Test":               strAssayNm = "HEALTH_C"
                    Case Else:                                      strAssayNm = ""
                End Select
            

'                If blnSame = False Then
'                    vasID.MaxRows = vasID.MaxRows + 1
'                    intRow = vasID.MaxRows
'                    gRow = intRow
'                End If
                                
                If blnSame = False Then
                    With mResult
                        .BarNo = strBarNo
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                        .MnmNm = strAssayNm
                        .MnmCd = strAssay
                    End With
                                    
                    Call SetPatInfo(strBarNo)
                    
                    vasRes.MaxRows = 0
                End If
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
                '-- 오른쪽 결과화면 초기화
               ' vasRes.MaxRows = 0
                
            End If
        
            'If intCnt > 8 And (InStr(intCnt, 4) > 0 Or InStr(intCnt, 9) > 0) Then
            If intCnt > 8 And ((intCnt + 1) Mod 5) = 0 Then
            'If intCnt > 9 And (intCnt Mod 5) = 0 Then
                strIntBase = varTmp(intCnt)


                Debug.Print strBarNo & ">>" & intCnt & ":" & strIntBase
'                If UCase(strIntBase) = "HX" Then
'                    strIntBase = "HX"
'                End If
'
'                If strIntBase = "Total_IgE" Then
'                    strIntBase = "IgE"
'                End If
                'strIntBase = strGubun & strIntBase
                
                
                ' ***************************************
                'intCnt + 1 : IU/ml
                'intCnt + 2 : Class
                ' ***************************************
                
                strIntResult = varTmp(intCnt + 1)
                strResult = strIntResult
                
                '-- Class 추가
                If strIntBase = "FD|IgE" Or strIntBase = "IN|IgE" Then
                    If strIntResult = "<100" Then
                        strClass = "1.0"
                    ElseIf strIntResult = "100-200" Then
                        strClass = "2.0"
                    ElseIf strIntResult = ">200" Then
                        strClass = "3.0"
                    End If
                Else
                    If Len(strIntBase) < 10 Then
                        strClass = varTmp(intCnt + 2)
                        strClass = Format(strClass, "0.0")
                    End If
                End If

                '-- 결과값에 Class값을 포함
                'strResult = strClass & " Class" & "(" & strResult & ")"
                'strResult = strClass & "(" & strResult & ")"

                If strResult <> "" And Len(strIntBase) > 0 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND GUBUN = '" & strAssayNm & "'"
                    'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                    
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
                            If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                SetText vasID, strResult, gRow, intCol
                                'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                Exit For
                            End If
                        Next
                        
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        'SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        'SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        
                        SetText vasRes, strResult, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strClass, lsResRow, colRESULT          '결과
                        
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
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
                        SQL = SQL & "   AND GUBUN = '" & strAssayNm & "'"
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
                                If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                    SetText vasID, strResult, gRow, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    'SetText vasRes, gArrEquip(intCol - colState, 8), lsResRow, colSUBCODE               'subcode
                                    SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                            'SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                            'SetText vasRes, strResult, lsResRow, colRESULT          '결과
                            
                            SetText vasRes, strResult, lsResRow, colMachResult     '장비결과
                            SetText vasRes, strClass, lsResRow, colRESULT          '결과
                            
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 10
            End If
        Next
        
        strGubun = ""
        strBarNo = ""
        
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
    
    Next

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAPEX_Front()
    Dim intCnt       As Integer
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strGubun     As String
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    
    For intCnt = 0 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Order
                strBarNo = Trim(mGetP(strRcvBuf, 2, "|"))
                
            Case "O"    '## Order
                strGubun = Trim(mGetP(strRcvBuf, 2, "|"))
                                
            Case "R"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
                    
                End If
        
        End Select
    Next

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
    Dim RS          As ADODB.Recordset
    Dim blnUpdate   As Boolean
    Dim intUpCnt    As Integer
    Dim strUpData   As String
    Dim strUpData1  As String
    Dim strGubuns   As String
    Dim varGubuns   As Variant
    Dim intCnt      As Integer
    Dim intCnt1     As Integer
    Dim strChannel  As String
    Dim strGubun    As String
    Dim strAssayNm  As String
    Dim strDOB      As String
    
    blnUpdate = False
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    'sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
    strChannel = Trim(GetText(vasRes, asRow2, colEQUIPCODE))
    strGubun = Trim(GetText(vasID, asRow1, colDOB))
    strAssayNm = Trim(GetText(vasID, asRow1, colASSAYNM))
    
    'If strChannel = "I1" Then Stop
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'" & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, asRow1, colDOB)) & "'"
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    '-- 같은 바코드에 (구분은 틀려도 됨) 같은 채널의 결과가 있는지 확인한다.
    '-- Select
          SQL = " SELECT RESULT,EQUIPRESULT, DISKNO " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "    AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ))
    SQL = SQL & "    AND DISKNO =  '" & strGubun & "' "
    SQL = SQL & "    AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
    SQL = SQL & "    AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'"
    SQL = SQL & "    AND SENDFLAG <> '2'"
    Set RS = cn.Execute(SQL, , 1)
    strUpData = ""
    strGubuns = ""
    intUpCnt = 0
        
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            If IsNumeric(RS.Fields("RESULT")) Then
                blnUpdate = True
                strUpData = Trim(GetText(vasRes, asRow2, colRESULT))
                strUpData1 = Trim(GetText(vasRes, asRow2, colMachResult))
                If strUpData > RS.Fields("RESULT") Then
                    'strUpData = SetResult(strUpData, strChannel)
                    strDOB = RS.Fields("DISKNO") & ""
                Else
                    strUpData = ""
                End If
            End If
            RS.MoveNext
        Loop
    End If
            
    '-- 만약 같은 결과가 있으면 그결과와 지금결과를 비교해서
    '   지금 결과가 크면 업데이트
    '   기존 결과가 크면 SKIP.
    '-- Update
    If blnUpdate = True Then
        If strUpData <> "" Then
                  SQL = "UPDATE PATRESULT Set "
            SQL = SQL & " EQUIPRESULT = '" & strUpData1 & "'," & vbCrLf
            SQL = SQL & " RESULT = '" & strUpData & "'" & vbCrLf
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "  AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
            SQL = SQL & "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
            SQL = SQL & "  AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf
            SQL = SQL & "  AND DISKNO = '" & strDOB & "'"
            Res = SendQuery(gLocal, SQL)
            
            If Res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        End If
    Else
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
        'SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colBREED))
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDOB))
        SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colBREED))
        SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
        If strUpData <> "" Then
            SQL = SQL & "','" & strUpData
        Else
            SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
        End If
        SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
        SQL = SQL & "',''"
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO))
        SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
        SQL = SQL & "','" & Mid(Trim(GetText(vasID, asRow1, colPNAME)), 1, 20)
        SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))
        SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
        SQL = SQL & "',''"
        SQL = SQL & ",''"
        SQL = SQL & ",''"
        SQL = SQL & ",'1'"
        SQL = SQL & ",''"
        SQL = SQL & ",'" & gIFUser
        SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colASSAYNM)) & "')"
        
        Res = SendQuery(gLocal, SQL)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
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

    txtStartNum.Text = "0"
    txtStopNum.Text = "0"
    txtCnt = ""

End Sub



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "검사자를 입력해주세요."
    lblUser.Caption = InputBox(sMsg, "검사자 입력")

End Sub

Private Sub txtBarNum_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarNum) Then
'            StatusBar1.Panels(3).Text = "바코드번호는 숫자만 입력이 가능합니다."
            txtBarNum = ""
            Exit Sub
        End If
        
        If Len(txtBarNum) <> 12 Then
'            StatusBar1.Panels(3).Text = "바코드 자릿수를 확인하세요"
            txtBarNum = ""
            Exit Sub
        End If
        
        If Trim(txtBarNum) <> "" Then
            Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
        End If
        vasID.RowHeight(-1) = 12
        txtBarNum.Text = ""
    End If
    
End Sub


Private Sub txtCnt_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        Call cmdSearch_Click
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
    Dim strAssayNm  As String
    Dim strEqpCode  As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
    lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPtID.Caption = Trim(GetText(vasID, Row, colCHARTNO))
    
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    lblPtest.Caption = Trim(GetText(vasID, Row, colASSAYNM))
        
'POBALL™ Food_Advanced
'POBALL™ Food_Intensive
'POBALL™ Inhalant_Advanced
'POBALL™ Premium
'POBALL™ Premium_Intensive

'                1.     POBALL™ Food Advanced Test
'                2.     POBALL™ Food Intensive Test
'                3.     POBALL™ Premium Test
'                4.     POBALL™ Inhalant Advanced Test
'                5.     POBALL™ Premium Intensive Test


    Select Case GetText(vasID, Row, colASSAYNM)
        Case "POBALL™ Premium Test":               strAssayNm = "PREMIUM"
        Case "POBALL™ Premium Intensive Test":     strAssayNm = "PREMIUM_I"
        Case "POBALL™ Food Advanced Test":         strAssayNm = "FOOD_A"
        Case "POBALL™ Basic Test":                 strAssayNm = "FOOD_A"
        Case "POBALL™ Food Intensive Test":        strAssayNm = "FOOD_I"
        Case "POBALL™ Inhalant Advanced Test":     strAssayNm = "INHALANT_A"
        Case "POBALL™ Health Care Test":           strAssayNm = "HEALTH_C"
        Case Else:                                  strAssayNm = ""
    End Select
    
    If lblSaveSeq.Caption = "" Then
        Exit Sub
    End If
    
'''    '-- 수기입력항목 조회
'''    '장비코드, 검사코드, 검사명, 결과, 순번
'''          SQL = "SELECT a.EQUIPCODE, a.EXAMCODE, a.EXAMNAME, a.EQUIPRESULT, a.RESULT, a.SEQNO, a.REFFLAG, a.EXAMSUBCODE " & vbCrLf
'''    SQL = SQL & "  FROM PATRESULT a, EQPMASTER b " & vbCrLf
'''    SQL = SQL & " WHERE a.EQUIPNO = '" & gEquip & "'" & vbCrLf
'''    SQL = SQL & "   AND a.EQUIPNO = b.EQUIPNO " & vbCrLf
'''    SQL = SQL & "   AND b.EXAMTYPE = '수기'" & vbCrLf
'''    SQL = SQL & "   AND b.GUBUN = '" & strASSAYNM & "' " & vbCrLf
'''    SQL = SQL & "   AND (a.RESULT = '' OR a.RESULT IS NULL) " & vbCrLf
'''    SQL = SQL & "   AND a.SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
'''    SQL = SQL & "   AND a.BARCODE = '" & lsID & "' " & vbCrLf
'''    SQL = SQL & "   AND a.DISKNO = '" & Trim(GetText(vasID, Row, colDOB)) & "' " & vbCrLf
'''    SQL = SQL & " GROUP BY a.SEQNO, a.EQUIPCODE, a.EXAMCODE, a.EXAMNAME, a.EQUIPRESULT, a.RESULT, a.SEQNO, a.REFFLAG, a.EXAMSUBCODE "
'''    SQL = SQL & " ORDER BY a.SEQNO * 10"
'''
'''    Set RS = cn.Execute(SQL, , 1)
'''    If Not RS.EOF = True And Not RS.BOF = True Then
'''        vasRes.MaxRows = 0
'''        Do Until RS.EOF
'''            With vasRes
'''                .MaxRows = .MaxRows + 1
'''                SetText vasRes, "0", .MaxRows, colCheckBox
'''                SetText vasRes, Trim(RS.Fields("EQUIPCODE")) & "", .MaxRows, colEQUIPCODE
'''                SetText vasRes, Trim(RS.Fields("EXAMCODE")) & "", .MaxRows, colEXAMCODE
'''                SetText vasRes, Trim(RS.Fields("EXAMNAME")) & "", .MaxRows, colEXAMNAME
'''                SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMachResult
'''                SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
'''                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSeq
'''                SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colFLAG
'''                SetText vasRes, Trim(RS.Fields("EXAMSUBCODE")) & "", .MaxRows, colSUBCODE
'''
'''                If Trim(RS.Fields("REFFLAG")) = "H" Then
'''                    .Row = .MaxRows
'''                    .Col = colRESULT
'''                    .ForeColor = vbRed
'''                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
'''                    .Row = .MaxRows
'''                    .Col = colRESULT
'''                    .ForeColor = vbBlue
'''                End If
'''
'''            End With
'''            RS.MoveNext
'''        Loop
'''    End If
'''
'''    RS.Close
    
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(Trim(GetText(vasID, Row, colEXAMDATE)), 1, 8) & "' " & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, Row, colDOB)) & "' " & vbCrLf
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
                
                strEqpCode = strEqpCode & "'" & Trim(RS.Fields("EQUIPCODE")) & "',"
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
    
    '-- 결과가 입력되지 않은 수기입력항목 조회
    If strEqpCode <> "" Then
        strEqpCode = Mid(strEqpCode, 1, Len(strEqpCode) - 1)
    Else
        Exit Sub
    End If
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, '' as EQUIPRESULT, '' as RESULT, SEQNO, '' as REFFLAG, '' as EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND EXAMTYPE = '수기'" & vbCrLf
    SQL = SQL & "   AND GUBUN = '" & strAssayNm & "' " & vbCrLf
    SQL = SQL & "   AND NOT EQUIPCODE IN (" & strEqpCode & ") " & vbCrLf
    SQL = SQL & " ORDER BY SEQNO * 10"

    Set RS = cn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
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
                
                SetBackColor vasRes, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
            End With
            RS.MoveNext
        Loop
    End If

    RS.Close
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
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' "
        'SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        Res = SendQuery(gLocal, SQL)

        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        vasID.MaxRows = vasID.MaxRows - 1
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            '-- 바뀐 바코드로 환자정보 불러오기
            Call GetSampleInfoW_NTL(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- 바코드 번호가 이전과 틀리다면 업데이트
            'If lsID <> lblChangeBar.Caption Then
            If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "yyyymmdd") & "' " & vbCrLf
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

                'SetRawData "[SQL]" & SQL
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
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colBREED))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colASSAYNM))
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


Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
    
'    Call SpreadSheetSort(vasRes, Col)

End Sub

Private Sub vasRes_KeyPress(KeyAscii As Integer)
    Dim strResult   As String
    
    With vasRes
        If KeyAscii = 13 And lblBarcode(0).Caption <> "" Then
            ' IU/ml
            If .ActiveCol = colMachResult Then
                '-- 결과 소수점 적용
                strResult = SetResult(Trim(GetText(vasRes, .ActiveRow, colMachResult)), Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)))
                .Col = colMachResult
                .Text = strResult
                
                If .BackColor = "7405514" Then
                    SetLocalDB vasID.ActiveRow, .ActiveRow, "1"
                    .Row = .ActiveRow
                    .Row2 = .ActiveRow
                    .Col = 1
                    .Col2 = colSUBCODE
                    .BlockMode = True
                    .BackColor = vbWhite
                    .BlockMode = False
                Else
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT " & vbCrLf
                    SQL = SQL & "   SET EQUIPRESULT ='" & strResult & "', " & vbCrLf
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
                
            ' Class
            ElseIf .ActiveCol = colRESULT Then
                '-- 결과 소수점 적용
                'strResult = SetResult(Trim(GetText(vasRes, .ActiveRow, colRESULT)), Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)))
                strResult = Trim(GetText(vasRes, .ActiveRow, colRESULT))
                .Col = colRESULT
                .Text = strResult
'                '-- H/L 일때 색표시
'                If gsFlag = "L" Then
'                    vasRes.Row = .ActiveRow
'                    vasRes.Col = colRESULT
'                    vasRes.ForeColor = vbBlue
'                ElseIf gsFlag = "H" Then
'                    vasRes.Row = .ActiveRow
'                    vasRes.Col = colRESULT
'                    vasRes.ForeColor = vbRed
'                End If
'
'                SetText vasRes, gsFlag, .ActiveRow, colFLAG
                
                If .BackColor = "7405514" Then
                    SetLocalDB vasID.ActiveRow, .ActiveRow, "1"
                    .Row = .ActiveRow
                    .Row2 = .ActiveRow
                    .Col = 1
                    .Col2 = colSUBCODE
                    .BlockMode = True
                    .BackColor = vbWhite
                    .BlockMode = False
                Else
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
            End If
        End If
    End With

End Sub

