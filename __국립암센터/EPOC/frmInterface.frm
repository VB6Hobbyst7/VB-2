VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " EPOC 인터페이스"
   ClientHeight    =   12000
   ClientLeft      =   -15
   ClientTop       =   570
   ClientWidth     =   18660
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
   ScaleHeight     =   12000
   ScaleWidth      =   18660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame FrmTempBox 
      Caption         =   "TempBox"
      Height          =   6585
      Left            =   14280
      TabIndex        =   49
      Top             =   2100
      Visible         =   0   'False
      Width           =   9165
      Begin VB.TextBox Text_Today 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3360
         TabIndex        =   75
         Text            =   "2002/02/18"
         Top             =   5280
         Width           =   2595
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "SEND"
         Height          =   345
         Left            =   1200
         TabIndex        =   70
         Top             =   4380
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   4350
         Width           =   5115
      End
      Begin VB.CommandButton cmdCancelState 
         BackColor       =   &H00C0E0FF&
         Caption         =   "검증취소"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7170
         Style           =   1  '그래픽
         TabIndex        =   68
         Top             =   2400
         Width           =   1305
      End
      Begin VB.ComboBox cmbDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1380
         TabIndex        =   65
         Text            =   "마취통증의학과"
         Top             =   3240
         Width           =   2595
      End
      Begin VB.ComboBox cmbDr 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5100
         TabIndex        =   64
         Top             =   3240
         Width           =   2625
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8460
         Top             =   210
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7980
         Top             =   210
      End
      Begin VB.CommandButton cmdQC 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6030
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton Command14 
         Caption         =   "사용자변경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1650
         TabIndex        =   56
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdResCall 
         Caption         =   "QC 결과전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   55
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   435
         Left            =   90
         TabIndex        =   54
         Top             =   1050
         Width           =   2325
      End
      Begin VB.CommandButton Command_setup 
         Caption         =   "코드설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2310
         TabIndex        =   53
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_close 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3420
         TabIndex        =   52
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_Config 
         Caption         =   "통신설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1200
         TabIndex        =   51
         Top             =   240
         Width           =   1065
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
         Height          =   585
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   240
         Value           =   1  '확인
         Width           =   1065
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   4770
         Top             =   270
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
               Picture         =   "frmInterface.frx":14F5
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":1A8F
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":2029
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":25C3
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":2E55
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":2FAF
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":3109
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   945
         Left            =   720
         TabIndex        =   111
         Top             =   5220
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
         SpreadDesigner  =   "frmInterface.frx":3263
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         Height          =   195
         Left            =   3450
         TabIndex        =   73
         Top             =   2640
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         Height          =   195
         Left            =   2325
         TabIndex        =   72
         Top             =   2640
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   71
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   3960
         Picture         =   "frmInterface.frx":34AA
         Top             =   2610
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   2805
         Picture         =   "frmInterface.frx":3A34
         Top             =   2610
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   1650
         Picture         =   "frmInterface.frx":3FBE
         Top             =   2610
         Width           =   240
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "처방부서"
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
         Left            =   420
         TabIndex        =   67
         Top             =   3300
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "처방의사"
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
         Index           =   5
         Left            =   4200
         TabIndex        =   66
         Top             =   3300
         Width           =   780
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   63
         Top             =   1830
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   11625
      Width           =   18660
      _ExtentX        =   32914
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19226
            MinWidth        =   19226
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2019-07-12"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 2:40"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   10245
      Left            =   60
      TabIndex        =   98
      Top             =   1320
      Width           =   18495
      Begin FPSpread.vaSpread vasID 
         Height          =   9975
         Left            =   90
         TabIndex        =   101
         Top             =   180
         Width           =   11535
         _Version        =   393216
         _ExtentX        =   20346
         _ExtentY        =   17595
         _StockProps     =   64
         ButtonDrawMode  =   4
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
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":4548
      End
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   11700
         TabIndex        =   102
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
            TabIndex        =   104
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
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
            Left            =   1890
            TabIndex        =   103
            Top             =   150
            Visible         =   0   'False
            Width           =   1875
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
            Left            =   480
            TabIndex        =   109
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1905
            TabIndex        =   108
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label4 
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
            TabIndex        =   107
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   5130
            TabIndex        =   106
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   105
            Top             =   720
            Width           =   1155
         End
      End
      Begin VB.CheckBox chkWAll 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   690
         TabIndex        =   100
         Top             =   270
         Width           =   225
      End
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
         TabIndex        =   99
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   9390
         Left            =   11700
         TabIndex        =   110
         Top             =   750
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   16563
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
         SpreadDesigner  =   "frmInterface.frx":51B1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9165
      Left            =   6570
      TabIndex        =   1
      Top             =   4410
      Visible         =   0   'False
      Width           =   15195
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   660
         TabIndex        =   45
         Top             =   300
         Width           =   195
      End
      Begin FPSpread.vaSpread vasID1 
         Height          =   8805
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   7965
         _Version        =   393216
         _ExtentX        =   14049
         _ExtentY        =   15531
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":584D
         UserResize      =   2
      End
      Begin FPSpread.vaSpread vasRes1 
         Height          =   8805
         Left            =   8250
         TabIndex        =   47
         Top             =   240
         Width           =   6795
         _Version        =   393216
         _ExtentX        =   11986
         _ExtentY        =   15531
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":9B4B
      End
      Begin VB.Label Label2 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   270
         TabIndex        =   58
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.Frame fraWork 
      Height          =   705
      Left            =   60
      TabIndex        =   80
      Top             =   570
      Width           =   10125
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
         ItemData        =   "frmInterface.frx":D951
         Left            =   10230
         List            =   "frmInterface.frx":D95E
         TabIndex        =   91
         Top             =   240
         Visible         =   0   'False
         Width           =   975
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
         Height          =   495
         Left            =   10230
         TabIndex        =   90
         Top             =   150
         Width           =   1185
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
         Height          =   495
         Left            =   10230
         TabIndex        =   89
         Top             =   150
         Width           =   1185
      End
      Begin VB.TextBox txtRack 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10290
         TabIndex        =   88
         Text            =   "5"
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtPos 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10650
         TabIndex        =   87
         Text            =   "A"
         Top             =   180
         Width           =   315
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
         Left            =   13260
         TabIndex        =   86
         Top             =   180
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdOrder 
         Appearance      =   0  '평면
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
         Height          =   495
         Left            =   11160
         TabIndex        =   85
         Top             =   150
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtTimer 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5850
         TabIndex        =   84
         Text            =   "60"
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdReadEpoc 
         Appearance      =   0  '평면
         Caption         =   "결과읽기"
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
         Left            =   4620
         TabIndex        =   83
         Top             =   180
         Width           =   1155
      End
      Begin VB.CheckBox chkSearch 
         Caption         =   "결과조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7050
         TabIndex        =   82
         Top             =   210
         Value           =   1  '확인
         Width           =   1245
      End
      Begin VB.CheckBox chkSave 
         Caption         =   "저장포함"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8430
         TabIndex        =   81
         Top             =   210
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   345
         Left            =   2820
         TabIndex        =   92
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   21233665
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   345
         Left            =   1170
         TabIndex        =   93
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   21233665
         CurrentDate     =   40248
      End
      Begin VB.Label Label20 
         Caption         =   "조회일자"
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
         Left            =   180
         TabIndex        =   95
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
         Left            =   2640
         TabIndex        =   94
         Top             =   330
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdRsltSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "결과조회"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10620
      Style           =   1  '그래픽
      TabIndex        =   79
      Top             =   660
      Width           =   1305
   End
   Begin VB.CommandButton cmdIFClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "화면정리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   13260
      Style           =   1  '그래픽
      TabIndex        =   78
      Top             =   660
      Width           =   1305
   End
   Begin VB.CommandButton cmdIFTrans 
      BackColor       =   &H00C0FFFF&
      Caption         =   "결과저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11940
      Style           =   1  '그래픽
      TabIndex        =   77
      Top             =   660
      Width           =   1305
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   14580
      Style           =   1  '그래픽
      TabIndex        =   76
      Top             =   660
      Width           =   1305
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   18600
      TabIndex        =   61
      Top             =   0
      Width           =   18660
      Begin VB.Timer tmrResult 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6540
         Top             =   0
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1290
         TabIndex        =   74
         Top             =   60
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
         Format          =   21233664
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "epoc EDM"
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
         Left            =   4140
         TabIndex        =   96
         Top             =   60
         Width           =   1380
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
         Index           =   3
         Left            =   240
         TabIndex        =   62
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "epoc EDM"
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
         Left            =   4140
         TabIndex        =   97
         Top             =   90
         Width           =   1380
      End
   End
   Begin VB.Frame FrmUseControl 
      Caption         =   "UseControl"
      Height          =   975
      Left            =   14220
      TabIndex        =   48
      Top             =   8880
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox Text4 
         Height          =   675
         Left            =   1050
         TabIndex        =   60
         Top             =   150
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   435
         Left            =   5460
         TabIndex        =   59
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   150
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   4455
      Left            =   540
      TabIndex        =   2
      Top             =   5820
      Visible         =   0   'False
      Width           =   13095
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1575
         Left            =   3540
         TabIndex        =   30
         Top             =   2400
         Width           =   9285
         Begin VB.TextBox txtEquipID 
            Height          =   345
            Left            =   3600
            TabIndex        =   41
            Text            =   "10"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Rack Pos"
            Height          =   375
            Left            =   7560
            TabIndex        =   40
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command10 
            Caption         =   "결과입력"
            Height          =   375
            Left            =   5880
            TabIndex        =   39
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEquipCode 
            Height          =   345
            Left            =   1710
            TabIndex        =   38
            Text            =   "0ADVI120"
            Top             =   1125
            Width           =   1875
         End
         Begin VB.CommandButton Command9 
            Caption         =   "장비ID조회"
            Height          =   375
            Left            =   60
            TabIndex        =   37
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command8 
            Caption         =   "미검사상세목록"
            Height          =   375
            Left            =   5010
            TabIndex        =   36
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command7 
            Caption         =   "미검사목록"
            Height          =   375
            Left            =   3360
            TabIndex        =   35
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command6 
            Caption         =   "검사상세목록"
            Height          =   375
            Left            =   1710
            TabIndex        =   34
            Top             =   690
            Width           =   1635
         End
         Begin VB.TextBox txtID 
            Height          =   345
            Left            =   6660
            TabIndex        =   33
            Text            =   "05111000003"
            Top             =   720
            Width           =   1875
         End
         Begin VB.CommandButton Command5 
            Caption         =   "검사목록"
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command4 
            Caption         =   "서버시간"
            Height          =   375
            Left            =   60
            TabIndex        =   31
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   1920
            TabIndex        =   43
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   3150
            TabIndex        =   42
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   210
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3360
         Width           =   945
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
         Height          =   615
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   27
         Top             =   1950
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   285
         Left            =   60
         TabIndex        =   26
         Top             =   555
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   240
         TabIndex        =   25
         Top             =   1380
         Width           =   3045
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   60
         TabIndex        =   18
         Top             =   3780
         Visible         =   0   'False
         Width           =   3675
         Begin VB.TextBox txtEnd 
            Alignment       =   1  '오른쪽 맞춤
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
            Height          =   315
            Left            =   1950
            TabIndex        =   21
            Top             =   180
            Width           =   885
         End
         Begin VB.TextBox txtStart 
            Alignment       =   1  '오른쪽 맞춤
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
            Height          =   315
            Left            =   630
            TabIndex        =   20
            Top             =   180
            Width           =   885
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "삭제"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   19
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "번호"
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
            Left            =   60
            TabIndex        =   23
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " - "
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
            Left            =   1530
            TabIndex        =   22
            Top             =   240
            Width           =   360
         End
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
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   1875
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   10260
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   16
         Top             =   1950
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   1260
         Picture         =   "frmInterface.frx":D976
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   2010
         Picture         =   "frmInterface.frx":DAA5
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   285
         Left            =   1710
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Top             =   570
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   240
         TabIndex        =   10
         Top             =   2850
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton cmdResSave 
         Caption         =   "결과저장"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5970
         TabIndex        =   9
         Top             =   1500
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 작성"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   2355
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   285
         Left            =   1710
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1125
         Left            =   10740
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":DBD7
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   3450
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":120A0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1125
         Left            =   7110
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":122E7
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   5295
         TabIndex        =   24
         Top             =   240
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":1252E
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   8925
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":12775
      End
      Begin VB.Label lblMT 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
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
         Left            =   9750
         TabIndex        =   44
         Top             =   2370
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "파일"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
      Begin VB.Menu MnTest 
         Caption         =   "테스트"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "설정"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
         Visible         =   0   'False
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "자동"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "수동"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const colCHECKBOX = 1
'Const colBARCODE = 2
'Const colSeqNo = 3
'Const colReceno = 4
'Const colRack = 5
'Const colPos = 6
'Const colPID = 7
'Const colPNAME = 8
'Const colPSEX = 9
'Const colPAGE = 10
'Const colPJumin = 11
'Const colState = 12

Const colOrd = 13
Const colRes = 14
Const colDate = 15
Const colTime = 16
Const colTestType = 17

Const colEQUIPCODE = 1
Const colEXAMCODE = 2
Const colEXAMNAME = 3
Const colRESULT = 4
Const colSeq = 5
Const colRCheck = 6

'2004/10/21 이상은
'Const colRefLow = 7
Const colResult1 = 7

Const colRefHigh = 8

Dim gRow As Long

Dim gsBarCode As String
Dim gsSampleType As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim strState        As String


Function LRC(ByVal asData As String) As String
    Dim i As Integer
    Dim a
    
    a = Asc(Left(asData, 1))
    
    For i = 2 To Len(asData)
        a = a Xor Asc(Mid(asData, i, 1))
    Next i
    
    If a = 3 Then a = 127
    
    LRC = Chr(a)
End Function

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If

End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub cmdCall_Click()
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
          
          'SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
          SQL = " SELECT '', SAVESEQ, EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    'SQL = SQL & "    AND SAVESEQ > 0 " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE,SENDFLAG DESC "
    
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

                    SetText vasID, "0", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("INOUT")) & "", .MaxRows, colINOUT
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

Private Sub cmdCancelState_Click()
    Dim i As Integer
    Dim lsBarcode As String
    Dim lsDTID As String
    Dim lsStr As String
    
    
    lsDTID = gDoctor(cmbDr.ListIndex).WKPERS_ID
    
    For i = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 1 Then
            lsBarcode = Trim(GetText(vasID, i, colBARCODE))
            lsStr = "<NewDataSet>"
            lsStr = lsStr & "    <Table>"
            lsStr = lsStr & "        <QID><![CDATA[PG_SRL.SLP91_U08]]></QID>"
            lsStr = lsStr & "        <QTYPE><![CDATA[Package]]></QTYPE>"
            lsStr = lsStr & "        <USERID><![CDATA[LIA]]></USERID>"
            lsStr = lsStr & "        <EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
            lsStr = lsStr & "        <TABLENAME><![CDATA[]]></TABLENAME>"
            lsStr = lsStr & "        <P0><![CDATA[" & lsBarcode & "]]></P0>"
            lsStr = lsStr & "        <P1><![CDATA[" & lsDTID & "]]></P1>"
            lsStr = lsStr & "        <P2><![CDATA[]]></P2>"
            lsStr = lsStr & "        <P3><![CDATA[]]></P3>"
            lsStr = lsStr & "    </Table>"
            lsStr = lsStr & "</NewDataSet>"
            
            Online_Result_Qry_Conf_Cancel lsStr
            
        End If
    Next
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    Dim lsPid As String
    Dim lsReceNo1 As String
    Dim lsReceNo2 As String
    
    Dim sStart As String
    Dim send As String
    
    sStart = Trim(txtStart.Text)
    send = Trim(txtEnd.Text)
    
    If sStart <> "" And send <> "" Then
        For lRow = sStart To send
            lsPid = Trim(GetText(vasID, lRow, 5))
            lsReceNo1 = Trim(GetText(vasID, lRow, 11))
            lsReceNo2 = Trim(GetText(vasID, lRow, 12))
            SQL = "Delete from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                  "  and equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and pid = '" & lsPid & "' " & vbCrLf & _
                  "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                  "  and receno1 = '" & lsReceNo2 & "' "
            res = SendQuery(gLocal, SQL)
            
            DeleteRow vasID, lRow, lRow
        Next lRow
    Else
        lRow = 1
        Do While lRow <= vasID.DataRowCnt
            vasID.Row = lRow
            vasID.Col = 1
            If vasID.Value = 1 Then
                lsPid = Trim(GetText(vasID, lRow, 5))
                lsReceNo1 = Trim(GetText(vasID, lRow, 11))
                lsReceNo2 = Trim(GetText(vasID, lRow, 12))
                SQL = "Delete from pat_res " & vbCrLf & _
                      "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and pid = '" & lsPid & "' " & vbCrLf & _
                      "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                      "  and receno1 = '" & lsReceNo2 & "' "
                res = SendQuery(gLocal, SQL)
                
                DeleteRow vasID, lRow, lRow
            Else
                lRow = lRow + 1
            End If
        Loop
    End If
    
    MsgBox "삭제 완료"
    chkAll.Value = 0
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
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
    
    txtRack.Text = "5"
    txtPos.Text = "A"
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            res = SaveTransDataW(lRow)
        
            If res = -1 Then
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
                
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
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

Private Sub cmdQC_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdReadEpoc_Click()
    txtTimer.Text = "1"
End Sub

Private Sub cmdResCall_Click()
'    frmResult.Show 0
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    gRow = 0
    
    txtRack.Text = "5"
    txtPos.Text = "A"
    
    
End Sub

Private Sub cmdResSave_Click()
'    Proc_Result txtBarcode
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            res = SaveTransDataW(lRow)
        
            If res = -1 Then
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
                
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
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
          
          'SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
          SQL = " SELECT '', SAVESEQ, EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    'SQL = SQL & "    AND SAVESEQ > 0 " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE,SENDFLAG DESC "
    
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

                    SetText vasID, "0", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("INOUT")) & "", .MaxRows, colINOUT
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

Private Sub cmdSL_Click()
    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        vasID.Width = 18285 '18075 '15225
    Else
        cmdSL.Caption = "▶"
        vasID.Width = 11535 '11355 '8475
    End If

End Sub

Private Sub cmdTest_Click()

    Call ABL500(Mid(Text5.Text, 2))
    Text5.Text = ""
    
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
End Sub

Private Sub Command_close_Click()
    Unload Me
End Sub

Private Sub Command_config_Click()
    frmConfig.Show 1
End Sub


Private Sub Command_setup_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub


Private Sub Command13_Click()
    Dim i As Integer
    
    SQL = "select item_code, item_name, m_stype_code, disp_seq, m_item_code from tbl_item"
    res = db_select_Vas(gLocal_1, SQL, vaSpread1)
    
    SQL = "delete from equipexam"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vaSpread1.DataRowCnt
    
        SQL = "insert into equipexam(equipno, examcode, equipcode, examname, examtype, seqno, resprec, examflag) " & vbCrLf & _
              "values('C064','" & Trim(GetText(vaSpread1, i, 1)) & "','" & Trim(GetText(vaSpread1, i, 5)) & "','" & Trim(GetText(vaSpread1, i, 2)) & "','" & Trim(GetText(vaSpread1, i, 3)) & "','" & Trim(GetText(vaSpread1, i, 4)) & "','1','1')"
        res = SendQuery(gLocal, SQL)
    Next
    
End Sub

Private Sub Command14_Click()
'    frmUserChange.Show 0
    
End Sub

Private Sub Command15_Click()
    
    Online_XML gXml_S07, "10010700001"
'    vasID.MaxRows = 1
'    SetText vasID, "10010700001", 1, colBarCode
'
'    Get_Sample_Info 1
'
End Sub

Private Sub Command16_Click()
    Dim lsChar As String
    Dim i As Long
    Dim strRece As String
    
'''    strRece = "<NewDataSet><Table>"
'''            strRece = strRece & vbCrLf & "<QID><![CDATA[PG_SRL.SLP91_U06]]></QID>"
'''            strRece = strRece & vbCrLf & "<QTYPE><![CDATA[Package]]></QTYPE>"
'''            strRece = strRece & vbCrLf & "<USERID><![CDATA[LIA]]></USERID>"
'''            strRece = strRece & vbCrLf & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
'''            strRece = strRece & vbCrLf & "<TABLENAME><![CDATA[]]></TABLENAME>"
'''            strRece = strRece & vbCrLf & "<P0><![CDATA[EM014]]></P0>"
'''            strRece = strRece & vbCrLf & "<P1><![CDATA[" & gsPID & "]]></P1>"
'''            strRece = strRece & vbCrLf & "<P2><![CDATA[" & Format(Date, "yyyymmdd") & " " & Format(Time, "hhmmss") & "]]></P2>"
'''            strRece = strRece & vbCrLf & "<P3><![CDATA[SVAN]]></P3>"
'''            strRece = strRece & vbCrLf & "<P4><![CDATA[" & gDoctor(cmbDr.ListIndex).WKPERS_ID & "]]></P4>"
'''            strRece = strRece & vbCrLf & "<P5><![CDATA[X0010]]></P5>"
'''            strRece = strRece & vbCrLf & "<P6><![CDATA[10120]]></P6>"
'''            strRece = strRece & vbCrLf & "<P7><![CDATA[]]></P7>"
'''            strRece = strRece & vbCrLf & "<P8><![CDATA[]]></P8>"
'''            strRece = strRece & vbCrLf & "<P9><![CDATA[]]></P9>"
'''            strRece = strRece & vbCrLf & "<P10><![CDATA[]]></P10>"
'''            strRece = strRece & vbCrLf & "</Table></NewDataSet>"
'''
'''            Online_Param gXml_U06, strRece
'''            gsBarCode = gEMRBarcode
'''
    
    
'    For i = 1 To Len(Text4.Text)
'
'        lsChar = Mid(Text4.Text, i, 1)
'
'        Select Case lsChar
'        Case chrENQ
'            SaveData "[RX]" & lsChar
'            MSComm1.Output = chrACK
'            SaveData "[TX]" & chrACK
'            txtData.Text = ""
'            gRow = -1
'        Case chrSTX
'            txtData.Text = lsChar
'        Case chrLF
'            txtData.Text = txtData.Text & lsChar
'            SaveData "[RX]" & txtData.Text
'            ABL500 Mid(txtData.Text, 2)
'            MSComm1.Output = chrACK
'            SaveData "[TX]" & chrACK
'            txtData.Text = ""
'
'        Case Else
'            txtData.Text = txtData.Text & lsChar
'        End Select
'
'    Next
    ABL500 Text4.Text
    Text4.Text = ""
End Sub



Private Sub Command3_Click()
    SQL = "CREATE INDEX resindex1 ON pat_res (examdate,equipno,barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex1 created"
    Else
        MsgBox "resindex1 failed"
    End If
    SQL = "CREATE INDEX resindex2 ON pat_res (examdate,equipno,barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex2 created"
    Else
        MsgBox "resindex2 failed"
    End If
    
    SQL = "CREATE INDEX resindex3 ON pat_res (barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex3 created"
    Else
        MsgBox "resindex3 failed"
    End If
    
    SQL = "CREATE INDEX resindex4 ON pat_res (barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex4 created"
    Else
        MsgBox "resindex4 failed"
    End If
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon

    'cmdReset_Click
    
    GetSetup
   ' gEquipCode = "EPOC-4"
   ' gIFUser = "11783"
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    
    SocketsInitialize
    gIP = GetTheIP
    SocketsCleanup
    
    StatusBar1.Panels(1) = gIP
    
    
    cmdIFClear_Click
    'lblclear_Click
        
        
    'MSComm1.CommPort = gSetup.gPort
'    MSComm1.RTSEnable = gSetup.gRTSEnable
'    MSComm1.DTREnable = gSetup.gDTREnable
    'MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

'    If MSComm1.PortOpen = False Then
'        MSComm1.PortOpen = True
'    End If
'
'    If MSComm1.PortOpen Then
'        frmInterface.StatusBar1.Panels(2).Text = "COM" & MSComm1.CommPort & " 포트에 연결 되었습니다"
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
    
    '-- EDM
    For i = 1 To 1
        If Not Connect_DRServer Then
            MsgBox "EDM에 연결되지 않았습니다."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
'    If Not Connect_Local_1 Then
'        MsgBox "연결되지 않았습니다."
'
'    End If
    
        
'    If Not Connect_Server Then
'        MsgBox "연결되지 않았습니다."
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If

'    cn_Server_Flag = dce_setenv("client.env", "", "")
'
'    Dim i_equip_cd$
'    Dim machine_id$(), equip_cd$(), equip_nm$()
'
'    'i_equip_cd = gEquip
'    res = sl_sel_machine_id(gEquip, machine_id(), equip_cd(), equip_nm())
'    If res > 0 Then
'        gEquipID = machine_id(0)
'    End If
'    SQL = "update equipexam set equipno = 'C063' where equipno = '" & gEquip & "'"
'    res = SendQuery(gLocal, SQL)

    'Text_Today = Format(CDate(GetDateFull), "yyyy/mm/dd")

    GetExamCode
        
    SetExamCode
        
    tmrResult.Interval = 1000
    tmrResult.Enabled = True

    Call cmdSL_Click
        
        
'    sDate = Format(DateAdd("y", CDate(Text_Today.Text), -365), "yyyymmdd")
'
'    SQL = "delete from pat_res where examdate < '" & sDate & "'"
'    res = SendQuery(gLocal, SQL)
    
    Online_XML gXml_S26, "SVAN"
'    For i = 0 To giIndex
'        cmbDr.AddItem gDoctor(i).WKPERS_NM, i
'    Next
'
'    cmbDr.ListIndex = 0
    
'    Dim strBuf
'
'    Text5.Text = ""
'             strBuf = "1H|\^&|||ABL80^300855||||||||1|20180405083448" & vbCrLf
'    strBuf = strBuf & "7F" & vbCrLf
'    strBuf = strBuf & "2O|1||33328971^33328971|||20180503000852||||ANONYMOUS|||250279^02||QCLevel^C8002" & vbCrLf
'    strBuf = strBuf & "7F" & vbCrLf
'    strBuf = strBuf & "3L||QC #^34|||20180503000852||||ANONYMOUS|||250279^02||QCLevel^C8002" & vbCrLf
'    strBuf = strBuf & "" & vbCrLf
'    strBuf = strBuf & ""
'
'    Text5.Text = strBuf

End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
Dim lsBarcode As String
Dim lsPid As String
Dim lsReceNo As String
Dim sRes As String


    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    
    lsBarcode = Trim(GetText(vasID, asRow, colBARCODE))   '샘플 바코드 번호
    
'    If Trim(lsbarcode) = "" Then: Exit Function
    sRes = Online_XML(gXml_S03, lsBarcode)
'    If sRes = 1 Then
        SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
        SetText vasID, gPat_Info_Select.PT_NM, asRow, colPNAME
        SetText vasID, gPat_Info_Select.Sex, asRow, colPSEX
        SetText vasID, gPat_Info_Select.Age, asRow, colPAGE
'        SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colSeqNo
        SetText vasID, Format(gPat_Info_Select.ACPT_DTETM, "yyyymmdd"), asRow, colDate
'        SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colReceno

'        vasID.RowHeight(asRow) = 20
        
        Get_Sample_Info = 1
'    End If
End Function

Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    
    If vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        End If
    Next i

    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            '.TypeEditCharSet = TypeEditCharSetAlphanumeric
            '.TypeEditCharCase = TypeEditCharCaseSetUpper
            
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 9
            
            'cboTest.AddItem gArrEquip(i + 1, 4) & Space(20) & "|" & gArrEquip(i + 1, 3)
        Next
        
'        cboTest.ListIndex = 0
    End With
    
End Sub



Function GetExamCode() As Integer
'    Dim i, j As Long
'
'    ClearSpread vasTemp
'    GetExamCode = -1
'
'    SQL = "Select equipcode, examcode, examname, reflow, refhigh " & vbCrLf & _
'          "From equipexam " & vbCrLf & _
'          "Where equipno = '" & gEquip & "' " & vbCrLf & _
'          "order by  examcode "
'    res = db_select_Vas(gLocal, SQL, vasTemp)
'    If res > 0 Then
'        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 6)
'    Else
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    For i = 1 To vasTemp.DataRowCnt
'        gArrEquip(i, 1) = i
'        For j = 1 To 5
'            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
'        Next j
'    Next i
'
'    GetExamCode = 1

    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  seqno * 10 "
    res = GetDBSelectVas(gLocal, SQL, vasCode)
    If res > 0 Then
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
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    
    Unload Me
    
    End
    
End Sub

Private Sub MDButton3_Click()

End Sub

Private Sub Label1_DblClick(Index As Integer)

    If chkSave.Enabled = True Then
        chkSave.Enabled = False
    Else
        chkSave.Enabled = True
    End If
    
End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show 1
End Sub

Private Sub MnTest_Click()

    If cmdTest.Visible = False Then
        cmdTest.Visible = True
        Text5.Visible = True
    Else
        Text5.Text = ""
        cmdTest.Visible = False
        Text5.Visible = False
    End If
    
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1
    
End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
End Sub

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    
    Select Case MSComm1.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
    End Select
    
    lsChar = MSComm1.Input
    
    StatusBar1.Panels(2) = lsChar
    
    Select Case lsChar
    Case chrENQ
        SaveData "[RX]" & lsChar
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        txtData.Text = ""
        gRow = -1
    Case chrSTX
        txtData.Text = lsChar
    Case chrLF
        txtData.Text = txtData.Text & lsChar
        SaveData "[RX]" & txtData.Text
        ABL500 Mid(txtData.Text, 2)
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        txtData.Text = ""
        
    Case Else
        txtData.Text = txtData.Text & lsChar
    End Select
    
End Sub

Sub ABL500(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPid As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsExamCode As String
    Dim lsResult As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    
    Dim sSampleType As String
    Dim sLotNo As String
    Dim sLevel As String
    
    Dim rv As Integer
    Dim vTemp As String
    Dim qOrdDate As String
    Dim qQMCode As String
    Dim qOrdSeqNo As String
    Dim qEquipCode As String
    Dim qSpcCode As String
    Dim qExamCode As String
    Dim qSetYN As String
    Dim qLotNo As String
    Dim qRoomCode As String
    Dim qQCType As String
    Dim qEditID As String
    Dim qEditIP As String
    Dim qTransStr As String
    Dim strRece As String

'〓 구분자
'io_erryn in out varchar2
'io_errmsg in out varchar2
'io_spcidlist in out varchar2
'io_outcnt in out varchar2
'io_rowcnt in out varchar2

    If asData = "" Then
        Exit Sub
    End If

    
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
        Var_Clear
        
        iCnt = 0
        
        For i = 1 To Len(asData)
            If Mid(asData, i, 1) = "|" Then
                iCnt = iCnt + 1

                Select Case iCnt
                    Case 13
                        gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                End Select
            End If
        Next i
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "O" Then
        sTmp = Trim(ResultTbl(4))      'Sample구분
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            gsSampleType = Mid(sTmp, 1, i - 1)
            If gsSampleType = "Sample #" Then
                gsBarCode = gsSampleType & Mid(sTmp, i + 1)
                gsSampleType = "P"
                
            ElseIf gsSampleType = "QC #" Then
                gsSampleType = "Q"
                sLotNo = Trim(ResultTbl(16)) 'lotno
                i = InStr(1, sLotNo, "")
                If i > 0 Then
                    sLotNo = Mid(sLotNo, 1, i - 1)
                End If
                i = InStr(1, sLotNo, "^")
                If i > 0 Then
'                    sLevel = Mid(sLotNo, 1, i - 1)
'                    sLotNo = Mid(sLotNo, i + 1)
                    sLotNo = Mid(sLotNo, 1, i - 1)
                    
                    sLotNo = Trim(mGetP(mGetP(asData, 16, "|"), 2, "^"))
                    sLotNo = Replace(sLotNo, vbCr, "")
                    sLotNo = Replace(sLotNo, vbLf, "")
                    sLotNo = Replace(sLotNo, chrETX, "")
                    
                    If Len(sLotNo) <> 5 Then
                        sLotNo = Mid(sLotNo, 1, 5)
                    End If
                End If
            Else
                gsSampleType = "C"
            End If
        Else
            gsSampleType = "P"
        End If

        If gsSampleType = "Q" Then  'QC 자동접수
            Online_TLA gXml_S11, gQCEquip, sLotNo
            qOrdDate = ""
            qQMCode = ""
            qOrdSeqNo = ""
            qEquipCode = ""
            qSpcCode = ""
            qExamCode = ""
            qSetYN = ""
            qLotNo = ""
            qRoomCode = ""
            qQCType = ""
            qEditID = ""
            qEditIP = ""
            qQCType = "I"
            qEditID = "10738"
            qEditIP = gIP
            
            For X = 0 To giIndex
                qOrdDate = qOrdDate & gQC_Select(X).ORDDATE & "〓"
                qQMCode = qQMCode & gQC_Select(X).QMCODE & "〓"
                qOrdSeqNo = qOrdSeqNo & gQC_Select(X).ORDSEQNO & "〓"
                qEquipCode = qEquipCode & gQC_Select(X).EQIPCODE & "〓"
                qSpcCode = qSpcCode & gQC_Select(X).SPCCODE & "〓"
                qExamCode = qExamCode & gQC_Select(X).EXAMCODE & "〓"
                qSetYN = qSetYN & gQC_Select(X).SETEXYN & "〓"
                qLotNo = qLotNo & gQC_Select(X).LOTNO & "〓"
                qRoomCode = qRoomCode & gQC_Select(X).ROOMCODE & "〓"
'                If X = 0 Then
'                    qOrdDate = gQC_Select(X).ORDDATE
'                    qQMCode = gQC_Select(X).QMCODE
'                    qOrdSeqNo = gQC_Select(X).ORDSEQNO
'                    qEquipCode = gQC_Select(X).EQIPCODE
'                    qSpcCode = gQC_Select(X).SPCCODE
'                    qExamCode = gQC_Select(X).EXAMCODE
'                    qSetYN = gQC_Select(X).SETEXYN
'                    qLotNo = gQC_Select(X).LOTNO
'                    qRoomCode = gQC_Select(X).ROOMCODE
'
'                Else
'                    qOrdDate = qOrdDate & "〓" & gQC_Select(X).ORDDATE
'                    qQMCode = qQMCode & "〓" & gQC_Select(X).QMCODE
'                    qOrdSeqNo = qOrdSeqNo & "〓" & gQC_Select(X).ORDSEQNO
'                    qEquipCode = qEquipCode & "〓" & gQC_Select(X).EQIPCODE
'                    qSpcCode = qSpcCode & "〓" & gQC_Select(X).SPCCODE
'                    qExamCode = qExamCode & "〓" & gQC_Select(X).EXAMCODE
'                    qSetYN = qSetYN & "〓" & gQC_Select(X).SETEXYN
'                    qLotNo = qLotNo & "〓" & gQC_Select(X).LOTNO
'                    qRoomCode = qRoomCode & "〓" & gQC_Select(X).ROOMCODE
'
'                End If
            Next
            
               
            qTransStr = "<Table>" & vbCrLf & _
                        "<QID><![CDATA[PG_SRL.SLP91_U03]]></QID>" & vbCrLf & _
                        "<QTYPE><![CDATA[Package]]></QTYPE>" & vbCrLf & _
                        "<USERID><![CDATA[LIA]]></USERID>" & vbCrLf & _
                        "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & vbCrLf & _
                        "<TABLENAME><![CDATA[]]></TABLENAME>" & vbCrLf & _
                        "<P0><![CDATA[" & qOrdDate & "]]></P0>" & vbCrLf & _
                        "<P1><![CDATA[" & qQMCode & "]]></P1>" & vbCrLf & _
                        "<P2><![CDATA[" & qOrdSeqNo & "]]></P2>" & vbCrLf & _
                        "<P3><![CDATA[" & qEquipCode & "]]></P3>" & vbCrLf & _
                        "<P4><![CDATA[" & qSpcCode & "]]></P4>" & vbCrLf & _
                        "<P5><![CDATA[" & qExamCode & "]]></P5>" & vbCrLf & _
                        "<P6><![CDATA[" & qSetYN & "]]></P6>" & vbCrLf & _
                        "<P7><![CDATA[" & qLotNo & "]]></P7>" & vbCrLf & _
                        "<P8><![CDATA[" & qRoomCode & "]]></P8>" & vbCrLf & _
                        "<P9><![CDATA[" & qQCType & "]]></P9>" & vbCrLf & _
                        "<P10><![CDATA[" & qEditID & "]]></P10>" & vbCrLf & _
                        "<P11><![CDATA[" & qEditIP & "]]></P11>" & vbCrLf & _
                        "<P12><![CDATA[]]></P12>" & vbCrLf & _
                        "<P13><![CDATA[]]></P13>" & vbCrLf & _
                        "<P14><![CDATA[]]></P14>" & vbCrLf & _
                        "<P15><![CDATA[]]></P15>" & vbCrLf & _
                        "<P16><![CDATA[]]></P16>" & vbCrLf & _
                        "</Table>"
    
            qTransStr = "<NewDataSet>" & qTransStr & "</NewDataSet>"
    
            Online_Param gXml_U03, qTransStr
            
            gsBarCode = gQC_Rece(2)
        
        ElseIf gsSampleType = "C" Then  'Cal
        
        Else
            
            strRece = "<NewDataSet><Table>"
            strRece = strRece & vbCrLf & "<QID><![CDATA[PG_SRL.SLP91_U06]]></QID>"
            strRece = strRece & vbCrLf & "<QTYPE><![CDATA[Package]]></QTYPE>"
            strRece = strRece & vbCrLf & "<USERID><![CDATA[LIA]]></USERID>"
            strRece = strRece & vbCrLf & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
            strRece = strRece & vbCrLf & "<TABLENAME><![CDATA[]]></TABLENAME>"
            strRece = strRece & vbCrLf & "<P0><![CDATA[EM014]]></P0>"
            strRece = strRece & vbCrLf & "<P1><![CDATA[" & gsPID & "]]></P1>"
            strRece = strRece & vbCrLf & "<P2><![CDATA[" & Format(Date, "yyyymmdd") & " " & Format(Time, "hhmmss") & "]]></P2>"
            strRece = strRece & vbCrLf & "<P3><![CDATA[SVAN]]></P3>"
            strRece = strRece & vbCrLf & "<P4><![CDATA[" & gDoctor(cmbDr.ListIndex).WKPERS_ID & "]]></P4>"
            strRece = strRece & vbCrLf & "<P5><![CDATA[X0010]]></P5>"
            strRece = strRece & vbCrLf & "<P6><![CDATA[10120]]></P6>"
            strRece = strRece & vbCrLf & "<P7><![CDATA[]]></P7>"
            strRece = strRece & vbCrLf & "<P8><![CDATA[]]></P8>"
            strRece = strRece & vbCrLf & "<P9><![CDATA[]]></P9>"
            strRece = strRece & vbCrLf & "<P10><![CDATA[]]></P10>"
            strRece = strRece & vbCrLf & "</Table></NewDataSet>"
            
            Online_Param gXml_U06, strRece
            gsBarCode = gEMRBarcode
            
        End If
        
        gOrderExam = ""
        
        gRow = -1
        For i = 1 To vasID.DataRowCnt
'            If sSampleType = "P" Then
                If Trim(GetText(vasID, i, colBARCODE)) = gsBarCode Then
                    gRow = i
                    Exit For
                End If
'            ElseIf sSampleType = "Q" Then

'            End If
        Next i
        
        If gRow < 0 Then
            gRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < gRow Then
                vasID.MaxRows = gRow
            End If
        End If
        
        SetText vasID, gsBarCode, gRow, colBARCODE
        
        vasActiveCell vasID, gRow, colBARCODE
        ClearSpread vasRes
        
        '샘플정보 가져오기
        If gsSampleType = "Q" Then
            SetText vasID, sLotNo, gRow, colPID
            SetText vasID, "QC", gRow, colPNAME
            
            Online_XML gXml_S07, gsBarCode
        ElseIf gsSampleType = "C" Then
            SetText vasID, "CAL", gRow, colPNAME
        
        Else
            If Trim(GetText(vasID, gRow, colPID)) = "" And Len(Trim(GetText(vasID, gRow, colBARCODE))) = 11 Then
                '검체번호에 대한 검사목록 전체조회 (상세조회)
                Get_Sample_Info gRow

            End If
            Online_XML gXml_S07, gsBarCode
        End If
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "P") Then          'Test Order Record
        sTmp = Trim(ResultTbl(4))      '검체번호
        i = InStr(1, sTmp, Chr(13))
        gsPID = ""
        If i > 0 Then
            gsPID = Mid(sTmp, 1, i - 1)
        Else
            gsPID = sTmp
        End If
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "L" Then
        gOrderExam = ""
        If MnTransAuto.Checked = True Then
        
            res = Insert_Data(gRow)
            
            If res = -1 Then
                SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", gRow, colState
            Else
               
                SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", gRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
                    
        End If
    
    End If
    

    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsTestID = Left(sTmp, i - 1)    '장비코드

        sTmp = ResultTbl(4)
        lsResult = Trim(sTmp)           '결과
        
        
        gsResDateTime = ResultTbl(10)    'result time
        
        ClearSpread vasTemp
        
        SQL = "Select examcode, examname From equipexam" & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "And equipcode = '" & lsTestID & "'  "
        If Mid(gsBarCode, 5, 1) <> "9" Then
            SQL = SQL & "and examcode like 'X%' "
        Else
            SQL = SQL & "and examcode like 'L%' "
        End If
        
        res = db_select_Vas(gLocal, SQL, vasTemp)
            
        If vasTemp.DataRowCnt > 0 Then
            k = -1
            If vasTemp.DataRowCnt > 1 And vasResTemp.DataRowCnt > 0 Then
                For j = 1 To vasTemp.DataRowCnt
                    k = -1
                    For X = 1 To vasResTemp.DataRowCnt
                        If Trim(GetText(vasResTemp, X, 1)) = Trim(GetText(vasTemp, j, 1)) Then
                            k = j
                            Exit For
                        End If
                    Next X
                    If k > 0 Then
                        Exit For
                    End If
                Next j
            End If
            
            If k < 1 Then
                k = 1
            Else
                vasTemp.MaxRows = k
            End If
            
            For j = k To vasTemp.DataRowCnt
'                i = i + 1
                
'                If IsNumeric(lsTestID) = True And IsNumeric(lsResult) = True Then
                    i = vasRes.DataRowCnt + 1
                    
                    If i > vasRes.MaxRows Then
                        vasRes.MaxRows = i
                    End If
                    
                    If i > 0 Then
                        lsExamCode = Trim(GetText(vasTemp, j, 1))
                        
                        '숫자만 디스플레이 하기
                        If IsNumeric(lsResult) = False Then
                            For ii = 1 To Len(lsResult)
                                If Mid(lsResult, ii, 1) = "?" Then
                                    lsResult = Mid(lsResult, ii + 1)
                                    
                                    Exit For
                                End If
                            Next ii
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        
                        lsEquipRes = lsResult
                        lsResult = SetResult(lsResult, lsTestID)
                        
                        SetText vasRes, lsTestID, i, colEQUIPCODE                           '장비코드
                        SetText vasRes, Trim(GetText(vasTemp, j, 1)), i, colEXAMCODE        '검사코드
                        SetText vasRes, Trim(GetText(vasTemp, j, 2)), i, colEXAMNAME        '검사명
                        
                        SetText vasRes, lsResult, i, colRESULT                              '변환결과
                        SetText vasRes, lsEquipRes, i, colResult1                           '장비결과
                        
                        Save_Local_One_1 gRow, i, "A"
                    End If
               
'                End If
            Next j
        End If
        
    End If
    
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
    
    If IsNumeric(sEquipRes) = False Then
        Exit Function
    End If
    
    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' "
    res = db_select_Col(gLocal, SQL)
    
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
    
'    If IsNumeric(gReadBuf(1)) = True Then
'        sLVal = gReadBuf(1)
'        If CCur(sLVal) > CCur(sEquipRes) Then
'            sResFlag = "<"
'        End If
'    End If
'
'    If IsNumeric(gReadBuf(2)) = True Then
'        sHVal = gReadBuf(2)
'        If CCur(sHVal) < CCur(sEquipRes) Then
'            sResFlag = ">"
'        End If
'    End If
    
    sResult = sResFlag & sResult
    SetResult = sResult
    
End Function

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull
    
'    If UCase(Left(Trim(GetText(vasID, asRow1, colPJumin)), 1)) = "F" Then
''        Save_Local_QC Trim(Text_Today.Text) & " " & Format(Time, "hh:nn:ss"), _
'                      Trim(GetText(vasID, asRow1, colBarcode)), _
'                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
'                      Trim(GetText(vasRes, asRow2, colResult)), _
'                      Trim(GetText(vasRes, asRow2, colResult1))
'        'Save_Local_QC Trim(Text_Today.Text) & " " & Trim(GetText(vasID, asRow1, colPID)), _
'                      Trim(GetText(vasID, asRow1, colBarcode)), _
'                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
'                      Trim(GetText(vasRes, asRow2, colResult)), _
'                      Trim(GetText(vasRes, asRow2, colResult1))
'        Exit Function
'    End If

    sCnt = ""
    If Trim(GetText(vasRes, asRow2, colEQUIPCODE)) = "" Then Exit Function
    
    SQL = "select count(*) from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
    res = db_select_Col(gLocal, SQL)
    sCnt = Trim(gReadBuf(0))
    If res = -1 Then
        SaveQuery SQL, 1
        Exit Function
    End If
    
    If Not IsNumeric(sCnt) Then
        sCnt = "0"
    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAGE)) Then
        SetText vasID, "0", asRow1, colPAGE
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If

    If sCnt = "0" Then
        'SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
              "pid, pname, jumin, page, psex, resdate, receno, " & _
              "equipcode, examcode, result, result1, sendflag, examname, " & _
              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colBARCODE)) & "', '" & Trim(GetText(vasID, asRow1, colSeqNo)) & "'," & _
              "'" & Trim(GetText(vasID, asRow1, colRack)) & "', '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colPNAME)) & "', '" & Trim(GetText(vasID, asRow1, colPJumin)) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colPAGE)) & "', '" & Trim(GetText(vasID, asRow1, colPSEX)) & "', " & _
              "'" & sExamDate & "', '" & Trim(GetText(vasID, asRow1, colReceno)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "', '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "',  " & _
              "'" & Trim(GetText(vasRes, asRow2, colRESULT)) & "', '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colEXAMNAME)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "',  " & _
              "'" & Trim(GetText(vasID, asRow1, colOrd)) & "', '" & Trim(GetText(vasID, asRow1, colRes)) & "', '" & Trim(GetText(vasID, asRow1, colDate)) & "') "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    Else
       ' SQL = " Update pat_res Set " & vbCrLf & _
              " diskno = '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
              " posno  = '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf & _
              " result = '" & Trim(GetText(vasRes, asRow2, colRESULT)) & "', " & vbCrLf & _
              " result1 = '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', " & vbCrLf & _
              " refflag = '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', " & vbCrLf & _
              " refvalue = '" & Trim(GetText(vasID, asRow1, colOrd)) & "', " & vbCrLf & _
              " panicvalue = '" & Trim(GetText(vasID, asRow1, colRes)) & "', " & vbCrLf & _
              " resdate = '" & sExamDate & "' " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
              " And equipcode = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "' " & vbCrLf & _
              " And examcode = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "' "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

    End If
    
End Function

Function Insert_Data(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim sDpcd, sDate1, sSlip, sItem, sOitp, sWkno As String
    Dim sIDNo, sSmyr, sSmsn, sSms1 As String
    Dim tSmsn As String
    Dim lsExamCode, lsResult As String
    Dim lPanicLow, lPanicHigh As Currency
    Dim lDeltaLow, lDeltaHigh, lDeltaMeth, lDeltaGap
    Dim lsPanic, lsDelta As String
    Dim lsPreDate, lsPreResult As String
    Dim lsNState, lsWState As String
    Dim lStdVal
    Dim lTerm As Long
    Dim lsQCChk As String

    Dim iNone, iDP

    Dim sResDate As String
    Dim sRDate As String
    Dim sRTime As String

    Dim lsID As String

    Dim i, j As Long
    Dim lRow As Long
    Dim lsQCOn As String
    
    Dim sResult As String
    Dim sExamCode As String
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sResStr As String
    Dim sResRow As Long
    Dim sResCnt As String
    Dim sEquipRes As String
    Dim sParam As String
    Dim X As Integer
    Dim sRes As String
    
    Insert_Data = -1

    lsQCOn = ""

    lRow = argSpcRow

    If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Function

    lsID = Trim(GetText(vasID, lRow, colBARCODE))
    sBarcode = ""
    sEquipCode = ""
    sResult = ""
    sExamCode = ""
    
    If lsID = "" Then Exit Function

    ClearSpread vasTemp
    ClearSpread vasTemp1

    iNone = 0
    iDP = 0
    
    gOrderExam = ""
    
'    Online_XML gXml_S07, lsID

    SQL = "Select equipcode, examcode, examname, result, result1 " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and result <> '' "
'    SQL = "Select equipcode, examcode, examname, result, result1 " & vbCrLf & _
'          "from pat_res " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & lsID & "' " & vbCrLf & _
'          "  and examcode in (" & gOrderExam & ") " & vbCrLf & _
'          "  and result <> '' "
    If asSend = 0 Then
'        SQL = SQL & vbCrLf & _
'          "  and sendflag <> 'C' "
    End If
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If vasTemp.DataRowCnt < 1 Then Exit Function

    Save_Raw_Data lsID & " : 서버 결과 전송 시작"
    Save_Raw_Data lsID & " : 장부 정보 가져오기"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" And IsNumeric(Trim(GetText(vasTemp, sResRow, 4))) = True Then
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[LIA]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & Trim(GetText(vasTemp, sResRow, 4)) & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & gEquip & "]]></P4>" & _
                    "<P5><![CDATA[]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
'            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Function
'            End If
        End If
    Next
    
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
    Online_Result_Qry sParam
    
    Insert_Data = 1
    
    sParam = "<NewDataSet><Table>"
    sParam = sParam & vbCrLf & "<QID><![CDATA[PG_SRL.SLP91_U07]]></QID>"
    sParam = sParam & vbCrLf & "<QTYPE><![CDATA[Package]]></QTYPE>"
    sParam = sParam & vbCrLf & "<USERID><![CDATA[LIA]]></USERID>"
    sParam = sParam & vbCrLf & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
    sParam = sParam & vbCrLf & "<TABLENAME><![CDATA[]]></TABLENAME>"
    sParam = sParam & vbCrLf & "<P0><![CDATA[" & lsID & "]]></P0>"
    sParam = sParam & vbCrLf & "<P1><![CDATA[10120]]></P1>"
    sParam = sParam & vbCrLf & "<P2><![CDATA[]]></P2>"
    sParam = sParam & vbCrLf & "<P3><![CDATA[]]></P3></Table></NewDataSet>"
    
    Online_Result_Qry_Conf sParam
    
    Save_Raw_Data lsID & " : 서버 결과 전송 완료!"

    Exit Function

ErrHandle:
    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
    
End Function

Sub Var_Clear()
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

Private Sub Picture1_Click()

End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        LX20 Text2
'        'MsgBox CS(Text2)
'
'        'Hitachi747 Mid(Text2.Text, 2)
'        'txtData = ""
'    End If
End Sub



Private Sub tmrResult_Timer()
    Dim RSH          As ADODB.Recordset
    Dim RSD          As ADODB.Recordset
    
    Dim strFDT As String
    Dim strTDT As String
    Dim pPtID       As String
    Dim pBarNo       As String
    Dim intRow      As Integer
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResRow    As String
    Dim lsEquipRes As String
    Dim lsResult_Buff As String
    Dim strComm    As String
    Dim strEDTM    As String
    Dim strSpcCd   As String
    Dim strTemperature   As String
    Dim strFiO2     As String
    
    Dim intCol     As Integer
    Dim strChartID  As String
    Dim strMaxSeq   As String
    Dim strQCID      As String
    
    Dim i As Integer
    
    txtTimer.Text = txtTimer.Text - 1
    If txtTimer.Text = "0" Then
        
        txtTimer.Enabled = False
        txtTimer.Text = "60"
        
        dtpStopDt.Value = Now
        
        strFDT = Format(dtpStartDt.Value, "yyyy-MM-dd HH:mm:00")
        strTDT = Format(dtpStopDt.Value, "yyyy-MM-dd 23:59:59")
             
        SQL = ""
        SQL = SQL & "Select M.TEST_ID, M.PatientOrLotID, CONVERT(CHAR(19), M.TESTDT, 20) AS EDTM, M.EDMUploadDT, M.OperatorID, M.HostSerialNumber, M.HostAlias, S.LISStatus, S.SendResultMessage,M.Department  " & vbNewLine
        SQL = SQL & "  From Tests M, TestSendLog S " & vbNewLine
        SQL = SQL & " Where M.TestStatus in ('Success', 'iQC') " & vbNewLine ''Incomplete'
        SQL = SQL & "   And M.Test_ID = S.EPOCTestID   " & vbNewLine
        SQL = SQL & "   And M.TestDT between '" & strFDT & "' and '" & strTDT & "'" & vbNewLine
        If gDepartment <> "" Then
            SQL = SQL & "   And M.Department = '" & gDepartment & "'" & vbNewLine
        End If
        
        If gHostSN <> "" Then
            SQL = SQL & "   And M.HostSerialNumber = '" & gHostSN & "'" & vbNewLine
        End If
        
        If chkSave.Value = "0" Then
            SQL = SQL & "   And s.LISStatus < 8 " & vbNewLine
        End If
        SQL = SQL & " Order by EDTM                  "
        
        'Call SetSQLData("EDM_LIST", SQL)

        Set RSH = cn_Ser_Census.Execute(SQL, , 1)
    
        If Not RSH.EOF = True And Not RSH.BOF = True Then
            Do Until RSH.EOF
                vasID.MaxRows = vasID.MaxRows + 1
                '-- 현재 Row
                intRow = vasID.MaxRows
                
                '-- 테스트번호
                pPtID = Trim(RSH.Fields("TEST_ID")) & ""
                '-- 환자번호
                strChartID = Trim(RSH.Fields("PatientOrLotID")) & ""
                strQCID = ""
                '-- 검사시간
                strEDTM = Trim(RSH.Fields("EDTM")) & ""
                
                '-- 검체정보
                SQL = "Select Value From TestAttributes  "
                SQL = SQL & " Where Test_Id = '" & pPtID & "'  "
                SQL = SQL & "   And TestAttrName = 'Sample type' "
                res = GetDBSelectColumn(gServer_Census, SQL)
                strSpcCd = Trim(gReadBuf(0))
                If Trim(strSpcCd) = "Venous" Then
                    '정맥혈
                    strSpcCd = "VBGA(OR)"
                Else
                    '동맥혈
                    strSpcCd = "ABGA(OR)"
                End If

                '-- 체온
                strTemperature = ""
                SQL = "Select Value From TestAttributes  "
                SQL = SQL & " Where Test_Id = '" & pPtID & "'  "
                SQL = SQL & "   And TestAttrName = 'Patient temperature' "
                res = GetDBSelectColumn(gServer_Census, SQL)
                'Call SetSQLData("체온", SQL)
                strTemperature = Trim(gReadBuf(0))

                '-- FiO2
                strFiO2 = ""
                SQL = "Select Value From TestAttributes  "
                SQL = SQL & " Where Test_Id = '" & pPtID & "'  "
                SQL = SQL & "   And TestAttrName = 'FiO2' "
                res = GetDBSelectColumn(gServer_Census, SQL)
                strFiO2 = Trim(gReadBuf(0))
                
                strMaxSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                
                SetText vasID, Format(strEDTM, "yyyymmddhhmmss"), intRow, colEXAMDATE
                SetText vasID, strMaxSeq, intRow, colSAVESEQ
                SetText vasID, strChartID, intRow, colCHARTNO
                SetText vasID, Trim(RSH.Fields("DEPARTMENT")) & "", intRow, colPOSNO
                
                '-- 검사자ID
                mResult.OperatorID = Trim(RSH.Fields("OperatorID")) & ""
                gIFUser = mResult.OperatorID
                
                If strChartID <> "" Then
                    '-- 검사자 정보 서버테이블에서 가져와 표시
                    'For i = 1 To vasID.MaxRows
                        Call GetSampleInfoW_NCC(intRow)
                    'Next
                End If
                
                vasRes.MaxRows = 0
                
                '-- 검사결과
                SQL = ""
                SQL = SQL & "SELECT Result_Type, Analyte, Value, Unit, InRange, ReturnCode "
                SQL = SQL & "  From Results "
                SQL = SQL & " Where TestID = '" & pPtID & "'"
                
                Call SetSQLData("EDM_RSLT", SQL)
                
                Set RSD = cn_Ser_Census.Execute(SQL, , 1)
            
               If Not RSD.EOF = True And Not RSD.BOF = True Then
                    Do Until RSD.EOF
                        strIntBase = Trim(RSD.Fields("Analyte")) & ""
                        strResult = Trim(RSD.Fields("Value")) & ""
                        strComm = Trim(RSD.Fields("InRange")) & ""
                        
                        If strResult <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                            
                            res = GetDBSelectColumn(gLocal, SQL)
                            '-- 오더 있을 경우
                            If res > 0 Then
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
                                
                                
                                '-- 결과저장용 seq
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText vasID, strResult, vasID.MaxRows, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        Exit For
                                    End If
                                Next
                                
                                '-- Work List
                                SetText vasID, "Result", vasID.MaxRows, colState                 '11 진행상태
                                
            
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- 로컬 저장
                                SetLocalDB intRow, lsResRow, "1", lsEquipRes
                                
                                SetText vasID, "1", intRow, colCHECKBOX
                                lsResult_Buff = ""
                                
                                strState = "R"
                            End If
                        End If
                        RSD.MoveNext
                        
                    Loop
                    
                    '-- 체온저장
                    If strTemperature <> "" Then
                        strIntBase = "Pt.temp"
                        strResult = Trim(strTemperature)
                        strComm = ""
                    
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        res = GetDBSelectColumn(gLocal, SQL)
                        '-- 오더 있을 경우
                        If res > 0 Then
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
                            
                            
                            '-- 결과저장용 seq
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strResult, vasID.MaxRows, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- Work List
                            SetText vasID, "Result", vasID.MaxRows, colState                 '11 진행상태
                            
        
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                            SetText vasRes, strResult, lsResRow, colRESULT          '결과
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB vasID.MaxRows, lsResRow, "1", lsEquipRes
                            
                            SetText vasID, "1", vasID.MaxRows, colCHECKBOX
                            lsResult_Buff = ""
                            
                            strState = "R"
                        End If
                    End If
                    
                    '-- Fio2저장
                    If strFiO2 <> "" Then
                        strIntBase = "FiO2"
                        strResult = Trim(strFiO2)
                        strComm = ""
                    
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        res = GetDBSelectColumn(gLocal, SQL)
                        '-- 오더 있을 경우
                        If res > 0 Then
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
                            
                            
                            '-- 결과저장용 seq
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strResult, vasID.MaxRows, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- Work List
                            SetText vasID, "Result", vasID.MaxRows, colState                 '11 진행상태
                            
        
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                            SetText vasRes, strResult, lsResRow, colRESULT          '결과
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB vasID.MaxRows, lsResRow, "1", lsEquipRes
                            
                            SetText vasID, "1", vasID.MaxRows, colCHECKBOX
                            lsResult_Buff = ""
                            
                            strState = "R"
                        End If
                    End If
                    
                    ' EDM - LIS 상태 업데이트
                    SQL = ""
                    SQL = SQL & "UPDATE TestSendLog SET LISStatus = '8' "
                    SQL = SQL & " WHERE EPOCTestID = '" & pPtID & "'"
                    
                    Call SetSQLData("EDM_UPDATE", SQL)
                    
                    res = SendQuery(gServer_Census, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                    '-- 서버저장
                    If MnTransAuto.Checked = True And strState = "R" Then
                        Call cmdIFTrans_Click
                    End If
                End If
                RSH.MoveNext
            Loop
        End If
    End If
    
    txtTimer.Enabled = True
    dtpToday.Value = Now
    
End Sub


' asRow1 = Work List
' asRow2 = 결과 List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String

    sExamDate = Format(dtpToday, "yyyymmddhhmmss")

    If Trim(GetText(vasID, asRow1, colSAVESEQ)) = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
          
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
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
    SQL = SQL & ", INOUT"                           '검체코드
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
    SQL = SQL & "','"
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colINOUT))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colDISKNO))
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
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPOSNO))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
'    SQL = SQL & ",'" & gIFUser
'    SQL = SQL & ",'" & mResult.OperatorID
    SQL = SQL & ",'" & gUserID
    SQL = SQL & "','')"
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Private Sub txtEnd_GotFocus()
    SelectFocus txtEnd
End Sub

'Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If IsNumeric(txtEnd) = False Then
'            txtEnd.SetFocus
'            Exit Sub
'        End If
'        cmdSend.SetFocus
'    End If
'End Sub
'
'Private Sub txtHelp_Change()
'
'End Sub

Private Sub txtID_GotFocus()
    SelectFocus txtID
End Sub

Private Sub txtStart_GotFocus()
    SelectFocus txtStart
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtStart) = False Then
            txtStart.SetFocus
            Exit Sub
        End If
        txtEnd.SetFocus
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
    
    If Row = 0 Then
        With vasID
            .Col = 1: .Col2 = .MaxCols
            .Row = 2: .Row2 = .DataRowCnt
            .SortBy = 0
            .SortKey(1) = Col       '정렬키 열번호

            .SortKeyOrder(1) = SortKeyOrderAscending
    
            .Action = ActionSort
        End With
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
    'lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    'lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    'lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    
'    If lblSaveSeq.Caption = "" Then
'        Exit Sub
'    End If
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    'If lblSaveSeq.Caption <> "" Then
    '    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    'End If
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
                SetText vasRes, "0", .MaxRows, colCHECKBOX
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
        If lsID = "" Then
            Exit Sub
        End If

        If MsgBox(lsID & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        res = SendQuery(gLocal, SQL)

        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        
        vasID.MaxRows = vasID.MaxRows - 1
        
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            'Exit Sub
            
            '-- 바뀐 바코드로 환자정보 불러오기
            Call GetSampleInfoW_NCC(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- 바코드 번호가 이전과 틀리다면 업데이트
            'If lsID <> lblChangeBar.Caption Then
            'If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "########") & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                'SetRawData "[SQL]" & SQL
                res = SendQuery(gLocal, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True
        ElseIf iCol = colCHARTNO Then
            'Exit Sub
            
            If Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
                Exit Sub
            End If
            
            '-- 바뀐 바코드로 환자정보 불러오기
            Call GetSampleInfoW_NCC(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- 바코드 번호가 이전과 틀리다면 업데이트
            'If lsID <> lblChangeBar.Caption Then
            'If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "########") & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                'SetRawData "[SQL]" & SQL
                res = SendQuery(gLocal, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True
                
            'End If
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

            res = SendQuery(gLocal, SQL)
                
            If res = -1 Then
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
                    'If gsFlag = "L" Then
                        vasID.Row = iRow
                        vasID.Col = i
                    '    vasID.ForeColor = vbBlue
                    'ElseIf gsFlag = "H" Then
                    '    vasID.Row = iRow
                    '    vasID.Col = i
                    '    vasID.ForeColor = vbRed
                    'End If
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
                    SQL = SQL & "','" & "" & "'"
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

                    res = SendQuery(gLocal, SQL)
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
    
'    With vasRes
'        If KeyAscii = 13 And .ActiveCol = colRESULT And lblBarcode(0).Caption <> "" Then
'            '-- 결과 소수점 적용
'            strResult = SetResult(Trim(GetText(vasRes, .ActiveRow, colRESULT)), Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)))
'            .Col = colRESULT
'            .Text = strResult
'            '-- H/L 일때 색표시
'            'If gsFlag = "L" Then
'                vasRes.Row = .ActiveRow
'                vasRes.Col = colRESULT
'            '    vasRes.ForeColor = vbBlue
'            'ElseIf gsFlag = "H" Then
'            '    vasRes.Row = .ActiveRow
'            '    vasRes.Col = colRESULT
'            '    vasRes.ForeColor = vbRed
'            'End If
'
'            SetText vasRes, "", .ActiveRow, colFLAG
'
'            SQL = ""
'            SQL = SQL & "UPDATE PATRESULT " & vbCrLf
'            SQL = SQL & "   SET RESULT  ='" & strResult & "', " & vbCrLf
'            SQL = SQL & "       REFFLAG    = '" & "" & "' " & vbCrLf
'            SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
'            SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(lblExamDate.Caption) & "' " & vbCrLf
'            SQL = SQL & "   AND SAVESEQ   = " & lblSaveSeq.Caption & vbCrLf
'            SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
'            SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, .ActiveRow, colEXAMCODE)) & "' " & vbCrLf
'            SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)) & "' " & vbCrLf
'
'            res = SendQuery(gLocal, SQL)
'
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Sub
'            End If
'
'        End If
'    End With

End Sub


Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    Dim lCCR, lM_C_ratio, lP_C_ratio As Long
    Dim sCCR, sCrea_S, sCrea_U, sM_ALB_U, sTP_U As String
    
    Dim sResult As String
    Dim sResult1 As String
    
    Dim i As Integer
    
    Dim sTotalVol As String
    
    Dim lsTime As String
    
    vasIDRow = vasID.ActiveRow
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    
    If KeyCode = vbKeyReturn Then

        If vasResCol = colRESULT Then
            
            If Trim(GetText(vasRes, vasResRow, colEQUIPCODE)) = "88888" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colRESULT))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colRESULT)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
            
            ElseIf Trim(GetText(vasRes, vasResRow, colEQUIPCODE)) = "99999" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colRESULT))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colRESULT)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
                
                If IsNumeric(sTotalVol) Then
                    lCCR = -1
                    sCCR = ""
                    sCrea_S = ""
                    sCrea_U = ""
                    sM_ALB_U = ""
                    sTP_U = ""
                    
                    i = 1
                    Do While i <= vasRes.DataRowCnt
                        Select Case Trim(GetText(vasRes, i, colEXAMCODE))
                        Case "L3117", "L3101", "L3102", "L3103"  'Microalbumun(24hr),Na,K,Cl
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 1000, "0.00")
                                SetText vasRes, sResult, i, colRESULT
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                            
                        Case "L3104", "L3106", "L3107", "L3109" 'Ca,Pi,UA,Protein(24hr)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colRESULT
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31094", "L31095" 'Protein 16hr, 8hr
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colRESULT
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31111", "L31112", "L31123", "L3113" 'Creatinie 16hr, 8hr,24hr, BUN(24hr UR)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, "L31123", i, colExamCode
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100 / 1000, "0.00")
                                SetText vasRes, sResult, i, colRESULT
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L3041", "88888"   'Serum Creatinine
                            sCrea_S = Trim(GetText(vasRes, i, colResult1))
                            
                            'Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31121"   'CCR
                            sCCR = Trim(GetText(vasRes, i, colResult1))
                            lCCR = i
                        Case "L31171"   'Microalbumin(random)
                            sM_ALB_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31110"  'Creatinine(random)
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31090"   'Protein(random)
                            sTP_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31172"   'Microalbumin / creatinine (random urine)
                            lM_C_ratio = i
                        Case "L31172"   'protein / creatinie (random)
                            lP_C_ratio = i
                        End Select
                        i = i + 1
                    Loop
                    
                    If lCCR > 0 And lCCR <= vasRes.DataRowCnt And IsNumeric(sCrea_U) = True And IsNumeric(sCrea_S) = True Then
                        sResult = Format(CCur(sCrea_U) * CCur(sTotalVol) / 1440 / CCur(sCrea_S), "0.000")
                        SetText vasRes, sResult, lCCR, colRESULT
                        SetText vasRes, sResult, lCCR, colResult1
                        Save_Local_One_1 vasIDRow, i, "A"
                    End If
                    
'                    If IsNumeric(sM_ALB_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sM_ALB_U) / CCur(sCrea_U), "0.00") * 100
'                        If lM_C_ratio > 0 And lM_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "101", i, colEquipCode
'                            SetText vasRes, "L31172", i, colExamCode
'                            SetText vasRes, "Microalbumin / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
'                    If IsNumeric(sTP_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sTP_U) / CCur(sCrea_U), "0.00") * 1000
'                        If lP_C_ratio > 0 And lP_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "102", i, colEquipCode
'                            SetText vasRes, "L31201", i, colExamCode
'                            SetText vasRes, "Urine Protein / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
                End If
            Else
                
'                If Trim(GetText(vasRes, vasIDRow, colPJumin)) = "F" Then
                
'                    If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colEXAMNAME)) & " 결과를 수정 하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'                        Exit Sub
'                    End If
'
'                    lsTime = Trim(GetText(vasID, vasIDRow, colPID))
'                    If Len(lsTime) = 4 Then
'                    Else
'                        lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'                    End If
'
'                    SQL = "update qc_res set result = '" & sResult & "' " & vbCrLf & _
'                          "where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'                          "  and examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'                          "  and examtime = '" & lsTime & "' " & vbCrLf & _
'                          "  and levelname = '" & Trim(GetText(vasID, vasIDRow, colBARCODE)) & "' " & vbCrLf & _
'                          "  and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEQUIPCODE)) & "' "
'                    res = SendQuery(gLocal, SQL)
'
'                    Exit Sub
'                Else
'
'
'                    sResult = Trim(GetText(vasRes, vasResRow, colRESULT))
'                    If MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!") = vbYes Then
'                        sResult = Trim(GetText(vasRes, vasResRow, colRESULT))
'
'                        SQL = " update pat_res set " & vbCrLf & _
'                              " Result = '" & sResult & "' " & vbCrLf & _
'                              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                              " And equipno = '" & gEquip & "' " & vbCrLf & _
'                              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBARCODE)) & "' " & vbCrLf & _
'                              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEQUIPCODE)) & "' "
'                        res = SendQuery(gLocal, SQL)
'
'                        If res = -1 Then
'                            SaveQuery SQL
'                            Exit Sub
'                        End If
'
'                        'SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
'
'                    End If
'                End If
            End If
            
            
        End If
    ElseIf KeyCode = vbKeyDelete Then
'        If Trim(GetText(vasID, vasIDRow, colPJumin)) = "F" Then
'
'            If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colEXAMNAME)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'                Exit Sub
'            End If
'
'            lsTime = Trim(GetText(vasID, vasIDRow, colPID))
'            If Len(lsTime) = 4 Then
'            Else
'                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'            End If
'
'            SQL = "Delete From qc_res a " & vbCrLf & _
'                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
'                  "  and a.levelname = '" & Trim(GetText(vasID, vasIDRow, colBARCODE)) & "' " & vbCrLf & _
'                  " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEQUIPCODE)) & "' "
'            res = SendQuery(gLocal, SQL)
'
'            Exit Sub
'        End If
'        If MsgBox("해당 환자의 " & Trim(GetText(vasRes, vasResRow, colEXAMNAME)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " Delete From pat_res " & vbCrLf & _
'              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBARCODE)) & "' " & vbCrLf & _
'              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEQUIPCODE)) & "' " & vbCrLf & _
'              " and examcode =  '" & Trim(GetText(vasRes, vasResRow, colEXAMCODE)) & "' "
'        res = SendQuery(gLocal, SQL)
'
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasRes, vasResRow, vasResRow
    
    End If
End Sub

Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
    'sControl = Trim(Left(asBarcode, 2))
    'sLotNo = Trim(Mid(asBarcode, 3))
    sControl = asBarcode
    sRefFlag = ""
    
    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
            If CCur(sRefHigh) < CCur(asRes2) Then
                sRefFlag = "H"
            End If
            If CCur(sRefLow) > CCur(asRes2) Then
                sRefFlag = "L"
            End If
        End If
    End If
    
    sCnt = ""
    SQL = "Select count(*) from qc_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        db_RollBack gLocal
        Exit Function
    End If
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        Exit Function
    End If
    If Not IsNumeric(sCnt) Then sCnt = "0"
    
    If CInt(sCnt) > 0 Then
        SQL = "delete from qc_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
              "  and levelname = '" & sControl & "' " & vbCrLf & _
              "  and equipcode = '" & asExamCode & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            'db_RollBack gLocal
            SaveQuery SQL
            Exit Function
        End If
    End If
    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gLocal
        SaveQuery SQL
        Exit Function
    End If
    
End Function


