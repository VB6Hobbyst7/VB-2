VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15360
   Visible         =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4905
      Top             =   5130
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5310
      Top             =   5130
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   4365
      Top             =   4545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   4350
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   5715
      Top             =   5130
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
            Picture         =   "frmComm.frx":219C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2736
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2CD0
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":326A
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3AFC
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3C56
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3DB0
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   45
      TabIndex        =   1
      Top             =   9015
      Width           =   15315
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6150
         Top             =   150
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   6615
         TabIndex        =   15
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Run"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   7920
         TabIndex        =   16
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Stop"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   9225
         TabIndex        =   17
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   10530
         TabIndex        =   18
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmComm.frx":3F0A
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdEot 
         Height          =   375
         Left            =   90
         TabIndex        =   84
         Top             =   120
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "초기화"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         ImgOutLineSize  =   3
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "작업대기 중.."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2430
         TabIndex        =   6
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   " 상태 :"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   225
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Caption         =   " Communication"
      SubCaption      =   "검사 장비와 통신하여 결과를 저장 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive : "
         Height          =   180
         Left            =   14055
         TabIndex        =   4
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   13020
         TabIndex        =   3
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   11925
         TabIndex        =   2
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14925
         Picture         =   "frmComm.frx":5794
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13695
         Picture         =   "frmComm.frx":5D1E
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12555
         Picture         =   "frmComm.frx":62A8
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8370
      Left            =   60
      TabIndex        =   7
      Top             =   585
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   14764
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "    WorkList     "
      TabPicture(0)   =   "frmComm.frx":6832
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameError"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "spdRstview"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSel(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSel(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "spdResult1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdRackNo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdWorkList1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdStartNo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdSearch"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdWorkList"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkAuto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "SSPanel1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdOrder"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdAppend(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Picture1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "spdWorkList"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdAction_Wide"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkOrdCheck"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtSeq"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "    받은 결과    "
      TabPicture(1)   =   "frmComm.frx":684E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel3"
      Tab(1).Control(1)=   "SSPanel5"
      Tab(1).Control(2)=   "SSPanel6"
      Tab(1).Control(3)=   "SSPanel7"
      Tab(1).Control(4)=   "SSPanel8"
      Tab(1).Control(5)=   "cmdRstQuery"
      Tab(1).Control(6)=   "lvwCuData"
      Tab(1).Control(7)=   "cmdAppend(1)"
      Tab(1).Control(8)=   "SSPanel4"
      Tab(1).Control(9)=   "cmdSel(3)"
      Tab(1).Control(10)=   "cmdSel(2)"
      Tab(1).Control(11)=   "chkServer"
      Tab(1).ControlCount=   12
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   405
         Left            =   6690
         TabIndex        =   85
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox chkOrdCheck 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   83
         Top             =   990
         Width           =   195
      End
      Begin BHButton.BHImageButton cmdAction_Wide 
         Height          =   390
         Left            =   3780
         TabIndex        =   82
         Top             =   900
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   688
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmComm.frx":686A
         PictureAlignment=   5
         Alignment       =   2
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1755
         Left            =   -74985
         TabIndex        =   45
         Top             =   315
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   3096
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   16311512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   1
            ItemData        =   "frmComm.frx":697C
            Left            =   960
            List            =   "frmComm.frx":6989
            Style           =   2  '드롭다운 목록
            TabIndex        =   53
            Top             =   480
            Width           =   1710
         End
         Begin VB.CheckBox chkGonDan 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "NEGATIVE"
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
            Height          =   225
            Left            =   960
            TabIndex        =   52
            Top             =   1110
            Width           =   1365
         End
         Begin VB.CheckBox chkER 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "PANIC"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2850
            TabIndex        =   51
            Top             =   870
            Width           =   1245
         End
         Begin VB.CheckBox chkNoResult 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "H/L"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   1920
            TabIndex        =   50
            Top             =   870
            Width           =   885
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "정상"
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
            Height          =   225
            Left            =   960
            TabIndex        =   49
            Top             =   870
            Width           =   945
         End
         Begin VB.CheckBox chePositive 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "POSITIVE"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2850
            TabIndex        =   48
            Top             =   1110
            Width           =   1245
         End
         Begin VB.TextBox txtKeyword 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2580
            MaxLength       =   40
            TabIndex        =   47
            Top             =   1380
            Width           =   1560
         End
         Begin VB.ComboBox cboKeyword 
            Height          =   300
            ItemData        =   "frmComm.frx":69B3
            Left            =   960
            List            =   "frmComm.frx":69B5
            Style           =   2  '드롭다운 목록
            TabIndex        =   46
            Top             =   1380
            Width           =   1590
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   300
            Left            =   2640
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   90
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
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
            Format          =   21430273
            CurrentDate     =   37112
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   300
            Left            =   960
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   90
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
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
            Format          =   21430273
            CurrentDate     =   37112
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "결과일자 "
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   90
            TabIndex        =   60
            Top             =   165
            Width           =   780
         End
         Begin VB.Label Label11 
            BackColor       =   &H00F8E4D8&
            Caption         =   "-"
            Height          =   285
            Left            =   2490
            TabIndex        =   59
            Top             =   135
            Width           =   195
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "전송구분"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   90
            TabIndex        =   58
            Top             =   540
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "결과구분"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   90
            TabIndex        =   57
            Top             =   900
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "검색기준"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   90
            TabIndex        =   56
            Top             =   1410
            Width           =   720
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5745
         Left            =   -75000
         TabIndex        =   61
         Top             =   2535
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   10134
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   16311512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin FPSpreadADO.fpSpread spdPtList 
            Height          =   5640
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   4110
            _Version        =   524288
            _ExtentX        =   7250
            _ExtentY        =   9948
            _StockProps     =   64
            BackColorStyle  =   1
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
            GrayAreaBackColor=   16777215
            GridColor       =   14737632
            MaxCols         =   4
            Protect         =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmComm.frx":69B7
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   885
         Left            =   -70755
         TabIndex        =   63
         Top             =   315
         Width           =   10965
         _Version        =   65536
         _ExtentX        =   19341
         _ExtentY        =   1561
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   16311512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.TextBox txtSEX 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   8640
            MaxLength       =   100
            TabIndex        =   70
            Text            =   "1234567890"
            Top             =   120
            Width           =   570
         End
         Begin VB.TextBox txtAGE 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   9300
            MaxLength       =   100
            TabIndex        =   69
            Text            =   "1234567890"
            Top             =   120
            Width           =   570
         End
         Begin VB.TextBox txtPtNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   1020
            MaxLength       =   100
            TabIndex        =   68
            Text            =   "1234567890"
            Top             =   120
            Width           =   2280
         End
         Begin VB.TextBox txtSpcNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   1020
            MaxLength       =   100
            TabIndex        =   67
            Text            =   "1234567890"
            Top             =   480
            Width           =   2280
         End
         Begin VB.TextBox txtPtId 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   4800
            MaxLength       =   100
            TabIndex        =   66
            Text            =   "1234567890"
            Top             =   120
            Width           =   2280
         End
         Begin VB.TextBox txtRcvDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   4800
            MaxLength       =   100
            TabIndex        =   65
            Text            =   "1234567890"
            Top             =   480
            Width           =   2280
         End
         Begin VB.TextBox txtResultDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   270
            Left            =   8610
            MaxLength       =   100
            TabIndex        =   64
            Text            =   "1234567890"
            Top             =   480
            Width           =   2280
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "검체번호"
            Height          =   180
            Left            =   120
            TabIndex        =   76
            Top             =   540
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "수검자명 "
            Height          =   180
            Left            =   120
            TabIndex        =   75
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "병록번호"
            Height          =   180
            Left            =   3900
            TabIndex        =   74
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "성별/나이"
            Height          =   180
            Left            =   7680
            TabIndex        =   73
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "접수일자"
            Height          =   180
            Left            =   3900
            TabIndex        =   72
            Top             =   540
            Width           =   720
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "검사일자"
            Height          =   180
            Left            =   7710
            TabIndex        =   71
            Top             =   540
            Width           =   720
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   7065
         Left            =   -70755
         TabIndex        =   77
         Top             =   1215
         Width           =   10965
         _Version        =   65536
         _ExtentX        =   19341
         _ExtentY        =   12462
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   16311512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin FPSpreadADO.fpSpread spdResult2 
            CausesValidation=   0   'False
            Height          =   6960
            Left            =   60
            TabIndex        =   78
            Tag             =   "20001"
            Top             =   60
            Width           =   10860
            _Version        =   524288
            _ExtentX        =   19156
            _ExtentY        =   12277
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            EditEnterAction =   8
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   13290186
            MaxCols         =   11
            MaxRows         =   15
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmComm.frx":92AA
            VisibleCols     =   10
            VisibleRows     =   13
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   435
         Left            =   -74985
         TabIndex        =   79
         Top             =   2085
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   767
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.CheckBox chkAll 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "전체선택"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   1245
         End
         Begin BHButton.BHImageButton cmdFind 
            Height          =   345
            Left            =   3015
            TabIndex        =   81
            Top             =   45
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            Caption         =   "조회"
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImgOutLineSize  =   3
         End
      End
      Begin FPUSpreadADO.fpSpread spdWorkList 
         Height          =   6915
         Left            =   60
         TabIndex        =   41
         Top             =   900
         Width           =   3705
         _Version        =   524288
         _ExtentX        =   6535
         _ExtentY        =   12197
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         MaxRows         =   10
         RowsFrozen      =   1
         SpreadDesigner  =   "frmComm.frx":9C28
      End
      Begin VB.PictureBox Picture1 
         Height          =   30
         Left            =   12270
         ScaleHeight     =   30
         ScaleWidth      =   30
         TabIndex        =   40
         Top             =   1770
         Width           =   30
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   0
         Left            =   13680
         TabIndex        =   25
         Top             =   420
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "결과등록"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdOrder 
         Height          =   375
         Left            =   10305
         TabIndex        =   32
         Top             =   420
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "장비오더전송"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   60
         TabIndex        =   26
         Top             =   390
         Width           =   3705
         _Version        =   65536
         _ExtentX        =   6535
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   16311512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin MSMask.MaskEdBox mskOrdDate1 
            Height          =   300
            Left            =   2475
            TabIndex        =   27
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOrdDate 
            Height          =   300
            Left            =   1170
            TabIndex        =   28
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F8E4D8&
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
            Left            =   2310
            TabIndex        =   30
            Top             =   150
            Width           =   315
         End
         Begin VB.Label Label6 
            BackColor       =   &H00F8E4D8&
            Caption         =   "처방일자 :"
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
            Height          =   225
            Left            =   150
            TabIndex        =   29
            Top             =   150
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  '평면
         Caption         =   "결과자동등록"
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
         Height          =   195
         Left            =   11880
         TabIndex        =   14
         Top             =   540
         Value           =   1  '확인
         Width           =   1620
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   9780
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   375
         Left            =   -69195
         TabIndex        =   21
         Top             =   480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "조회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4830
         Left            =   -67980
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   8520
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -67935
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "서버등록"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   420
         Left            =   60
         TabIndex        =   19
         Top             =   7860
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   741
         Caption         =   "WorkList 등록"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   375
         Left            =   3840
         TabIndex        =   31
         Top             =   420
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "조회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdStartNo 
         Height          =   375
         Left            =   8400
         TabIndex        =   33
         Top             =   60
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "Pos변경"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdWorkList1 
         Height          =   390
         Left            =   90
         TabIndex        =   34
         Top             =   4905
         Visible         =   0   'False
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   688
         Caption         =   "불러오기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdRackNo 
         Height          =   375
         Left            =   6840
         TabIndex        =   35
         Top             =   60
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "Rack변경"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   465
         Left            =   -63525
         TabIndex        =   37
         Top             =   405
         Visible         =   0   'False
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.OptionButton optGae 
            BackColor       =   &H000080FF&
            Caption         =   "계약자/고계약자"
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
            Height          =   285
            Left            =   1620
            TabIndex        =   39
            Top             =   90
            Width           =   1860
         End
         Begin VB.OptionButton optJong 
            BackColor       =   &H000080FF&
            Caption         =   "종합검진"
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
            Height          =   285
            Left            =   270
            TabIndex        =   38
            Top             =   90
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin FPUSpreadADO.fpSpread spdResult1 
         Height          =   7380
         Left            =   3780
         TabIndex        =   43
         Top             =   900
         Width           =   8085
         _Version        =   524288
         _ExtentX        =   14261
         _ExtentY        =   13018
         _StockProps     =   64
         ColsFrozen      =   6
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   10
         RowsFrozen      =   1
         SpreadDesigner  =   "frmComm.frx":ACE8
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   556
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":BDA6
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   556
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":C228
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   11
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":C696
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   12
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":CB18
      End
      Begin Threed.SSCheck chkServer 
         Height          =   165
         Left            =   -65550
         TabIndex        =   36
         Top             =   630
         Visible         =   0   'False
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   291
         _StockProps     =   78
         Caption         =   " SERVER DATA"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FPUSpreadADO.fpSpread spdRstview 
         Height          =   6105
         Left            =   11880
         TabIndex        =   42
         Top             =   900
         Width           =   3285
         _Version        =   524288
         _ExtentX        =   5794
         _ExtentY        =   10769
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   10
         RowsFrozen      =   1
         SpreadDesigner  =   "frmComm.frx":CF86
      End
      Begin Threed.SSFrame FrameError 
         Height          =   1260
         Left            =   11880
         TabIndex        =   22
         Top             =   7020
         Width           =   3300
         _Version        =   65536
         _ExtentX        =   5821
         _ExtentY        =   2222
         _StockProps     =   14
         Caption         =   "::: Event Log "
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtResult 
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1050
            Left            =   570
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   23
            Top             =   330
            Visible         =   0   'False
            Width           =   3150
         End
         Begin VB.ListBox List1 
            Height          =   960
            Left            =   60
            TabIndex        =   24
            Top             =   210
            Width           =   3150
         End
      End
      Begin VB.Label Label19 
         Caption         =   "Number :"
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
         Height          =   225
         Left            =   5670
         TabIndex        =   86
         Top             =   510
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "::::: Happy Call Center : 0505-831-1515 :::::"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   10890
         TabIndex        =   44
         Top             =   90
         Width           =   4290
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "순서"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "등록번호"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "성  명"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "검체번호"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "검체번호"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "상 태"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "검사항목"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""

Const Field_      As String = "|"
Const Repeat_     As String = "\"
Const Component_  As String = "^"
Const Escape_     As String = "&"
Const Slash_      As String = "/"
Dim cntField_     As Integer '|
Dim cntRepeat_    As Integer '\
Dim cntComponent_ As Integer '^
Dim cntEscape_    As Integer '&
Dim cntSlash_     As Integer '/
Dim Patiant_Recevid As Boolean
Dim sStxCheck As Integer
Dim sEtxCheck As Integer
' --------------------------------------------------------------
Private Type typeElecsys2010
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(100)   As String
    Result(100)   As String
    Status(100)   As String
    Rerun(100)    As String
End Type
Dim Elecsys2010 As typeElecsys2010
Dim strOrdLst As String

Dim fElecsys2010(100) As String
Dim fElecsys2010_1(100) As String
Dim SendData(10)     As String
Dim SendCount        As String
Dim Or_Seq           As Integer
Dim SendBuffW           As String
Dim SendBuffT           As String
Dim intRow          As Integer
Dim brStr           As String

Dim fRcvString As String
Dim cntCheckSum      As Integer
Dim ReceiveData      As String

Dim SendFlg          As Boolean
Dim HostOutput       As String

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult

Private mAdoRs      As ADODB.Recordset
Private mAdoRs1     As ADODB.Recordset
Private mAdoRs2     As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String
Private f_strJOB_FLAG   As String
Private f_strOrdList    As String
Private f_intSampleNo   As Integer

Private f_blnWorkList   As Boolean
Private f_lngWork_Row   As Long

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Dim fCellDynCfg(100) As Integer
Dim fCellDynSize(100, 1) As Integer
Dim fChannel() As String
Dim pName   As String
Dim pNo     As String
Dim chkEnq  As Integer

Dim flgETB           As Boolean
Dim flgETX           As Boolean

Private Type SugaMatch
    TestId          As String
    Sugacd          As String
    Testcd          As String
    DecPoint        As Long
End Type

Dim SMatch() As SugaMatch
Dim CountTest As Integer, sErrorFlag As Boolean

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestCd(100) As String
End Type

Private f_typCode() As TYPE_CD

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String

Dim SeqNo As String
Dim RecordChk   As Boolean

Dim G_CLVALU    As String
Dim G_CHVALU    As String
Dim G_EVALUATE  As String
Dim G_PANIC     As String
Dim G_DELTA     As String
Dim strFrameNo  As Integer
Dim OrderCnt As Integer
Dim vRow As Integer
Dim sPatiant_No As String

Dim gTestCode As String
Dim OrderSort_Flag As Integer
Const gItemStartPos As Integer = 10


Private Function f_funGet_ConvertResult(ByVal strRstval As String) As String

    Dim intPos  As Integer
    Dim strTmp1 As String, strTmp2  As String
    
    intPos = InStr(strRstval, "E")
    If intPos > 0 Then
        strTmp1 = Mid$(strRstval, 1, intPos - 1)
        strTmp2 = Mid$(strRstval, intPos + 1)
        
        If Mid$(strTmp2, 1, 1) = "-" Then
            f_funGet_ConvertResult = Round(Val(strTmp1) * (0.1 ^ Val(Mid$(strTmp2, 2))), 2)
        Else
            f_funGet_ConvertResult = Round(Val(strTmp1) * (10 ^ Val(Mid$(strTmp2, 2))), 2)
        End If
    Else
        f_funGet_ConvertResult = strRstval
    End If
    
End Function

Private Function MakeCS(Source As String) As String
    Dim X      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    For X = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, X, 1))
    Next X
    SumCS = Hex(AddCS)
    ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
    ChkCS = ChkCS & Right(SumCS, 1)
    MakeCS = ChkCS
End Function

Public Function GetChkSum(ByVal pMsg As String) As String

    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = LCase(Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2))
    End If

End Function

Private Function f_funGet_SpreadRow(ByVal objSpd As Object, ByVal intCol As Integer, _
                                    ByVal strPara As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For intRow = 1 To .maxrows
            .GetText intCol, intRow, varTmp
            If Trim$(varTmp) = strPara Then
                f_funGet_SpreadRow = intRow
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub f_subGet_JobList(ByVal strKeyno As String, ByRef strOrder As String, _
                             ByRef intOrdCnt As Integer, ByRef strSpec As String, _
                             ByRef strPcFlag As String)

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    strOrder = "":  strPcFlag = "  ":   strSpec = "SE": intOrdCnt = 0
    sqlDoc = "select ORD_CODE, CHART_NO From L3A01" & _
             " where SAMPLE_DATE = '" & Mid$(strKeyno, 1, 8) & "'" & _
             "   and SAMPLE_SEQ  = " & Format(Mid$(strKeyno, 9, 3), "##0") & "" & _
             "   and PART        = '" & Mid$(strKeyno, 12, 2) & "'"
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD, REMARK, AUTOVERIFY from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "ORD_CODE = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            Select Case Trim(adoRS2(2) & "")
                Case "128": strSpec = "PL"
                Case Else:  strSpec = "SE"
            End Select
            
            If Trim(adoRS2("TESTCD_EQP") & "") = "XXX" Then
                strOrder = strOrder + "06A ," + Trim$(adoRS2("AUTOVERIFY") & "") + ",": strPcFlag = "PC"
            Else
                strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + " ," + Trim$(adoRS2("AUTOVERIFY") & "") + ","
            End If
            intOrdCnt = intOrdCnt + 1
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
    
End Sub

Private Sub f_subGet_WorkList(ByRef strOrder As String, ByRef intOrdCnt As Integer, _
                              ByRef strSpec As String, ByRef strPcFlag As String, _
                              ByVal intRow As Integer)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant
    Dim intCol  As Integer
    
    Dim itemX   As ListItems
    
    Set itemX = lvwCuData.ListItems
    
    strOrder = "":  strPcFlag = "  ": strSpec = "SE":   intOrdCnt = 0
    With spdWorkList
        For intCol = 5 To .MaxCols
            .Row = intRow:  .Col = intCol
            If .BackColor = &HC6FEFF Then
                Select Case itemX.Item(intCol - 4).SubItems(11)
                    Case "128": strSpec = "PL"
                    Case Else:  strSpec = "SE"
                End Select
                .GetText intCol, 0, varTmp
                
                If itemX.Item(intCol - 4).tag = "XXX" Then
                    strOrder = strOrder + "06A ," + itemX.Item(intCol - 4).SubItems(10) + ",": strPcFlag = "PC"
                Else
                    strOrder = strOrder + itemX.Item(intCol - 4).tag + " ," & itemX.Item(intCol - 4).SubItems(10) + ","
                End If
                intOrdCnt = intOrdCnt + 1
            End If
        Next
    End With
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
   
End Sub

Private Sub f_subSet_ComCharacter()

    MSG_STX = Chr(COM_STX)
    MSG_ETX = Chr(COM_ETX)
    MSG_ENQ = Chr(COM_ENQ)
    MSG_EOT = Chr(COM_EOT)
    MSG_ACK = Chr(COM_ACK)
    MSG_NAK = Chr(COM_NACK)
    MSG_CR = Chr(COM_CR)
    MSG_LF = Chr(COM_LF)
    MSG_CRLF = Chr(COM_CR) & Chr(COM_LF)
    
End Sub

Private Sub f_subSet_ItemHeader()
    
    '검사코드 테이블
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "검사결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFL", "참고치(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFH", "참고치(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "재검", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Sub f_subSet_ItemComplete(lvw As Listview)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    
    Dim intCol  As Integer
    
    lvw.ColumnHeaders.Clear
    Call lvw.ColumnHeaders.Add(, "EQP_ID", "검체 번호")
    
    intCol = 4
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = '" & INS_CODE & "') AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With lvw
            .Enabled = True
            Set itemH = .ColumnHeaders.Add
            With itemH
                '컬럽 헤더키를 장비검사 코드로
                .Key = COL_KEY & Trim(adoRS.Fields("TESTCD_EQP") & "")
                '컬럽명은 검사 항목 이름
                .Text = Trim(adoRS.Fields("TESTNM") & "")
                '테그는 검사 코드로
                .tag = Trim(adoRS.Fields("TESTCD") & "")
                .Width = 700
                .Alignment = lvwColumnCenter
            End With
            Set itemH = Nothing
        End With
        
        With spdWorkList
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Function f_subSet_WorkList(ByVal strDate As String, ByVal strDate1 As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
     
    sqlDoc = ""
    sqlDoc = sqlDoc & vbCrLf & "SELECT"
    sqlDoc = sqlDoc & vbCrLf & "       L.Ymd AS REQDATE ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.Ymd + '^' + L.ChtNo AS SPCNO ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.ChtNo AS PTID ,"
    sqlDoc = sqlDoc & vbCrLf & "       P.Nm AS PTNM ,"
    sqlDoc = sqlDoc & vbCrLf & "       '' AS SEX,"
    sqlDoc = sqlDoc & vbCrLf & "       P.Age AS AGE ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.NO AS ORDNO "
    sqlDoc = sqlDoc & vbCrLf & "  FROM Lab_List L,"
    sqlDoc = sqlDoc & vbCrLf & "       Person p"
    sqlDoc = sqlDoc & vbCrLf & " Where p.CODE = L.ChtNo"
    sqlDoc = sqlDoc & vbCrLf & "   AND L.Ymd BETWEEN '" & strDate & "' AND '" & strDate1 & "'"
    sqlDoc = sqlDoc & vbCrLf & "   AND (L.KulGa = '' OR L.KulGa IS NULL)"
    sqlDoc = sqlDoc & vbCrLf & "   AND L.CODE IN (" & Mid(gTestCode, 1, Len(gTestCode) - 1) & " ) "
    
    Set AdoRs_SQL = New ADODB.Recordset
  
    AdoRs_SQL.CursorLocation = adUseClient
    AdoRs_SQL.Open sqlDoc, AdoCn_SQL

    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_WorkList = Nothing
        RecordChk = False
    Else
        Set f_subSet_WorkList = AdoRs_SQL
        RecordChk = True
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_ResultList(ByVal strDate As String, ByVal strDate1 As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_ResultList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset

    If optJong = True Then
        If D0COM_CENTERCOD = "10" Then
            sqlDoc = "         SELECT center_code, name, resdnt, health_num, sample_num"
            sqlDoc = sqlDoc & "     , interface_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE interface_day = " + Chr(39) + strDate + Chr(39)
            sqlDoc = sqlDoc & "   AND center_code = " + Chr(39) + D0COM_CENTERCOD + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            
            sqlDoc = sqlDoc & "   AND NOT sample_num = 0"
            sqlDoc = sqlDoc & " ORDER BY sample_num ASC"
        Else
            '각 총국에 종합건강진단 검사
            sqlDoc = "         SELECT center_code, name, resdnt, health_num, serial_num"
            sqlDoc = sqlDoc & "     , health_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE health_day = " + Chr(39) + strDate + Chr(39)
            sqlDoc = sqlDoc & "   AND center_code = " + Chr(39) + D0COM_CENTERCOD + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            
            sqlDoc = sqlDoc & " ORDER BY serial_num ASC"
        End If
    ElseIf optGae = True Then
        If D0COM_CENTERCOD = "10" Then
            sqlDoc = "SELECT DISTINCT center_code, name, resdnt, health_num, sample_num"
            sqlDoc = sqlDoc & "     , interface_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE interface_day = " + Chr(39) + strDate + Chr(39)
            '-- Query 추가 본부일경우 본부/남대문/인천데이터만 조회(종양표시자/갑상선의 경우..)
            sqlDoc = sqlDoc & "   AND center_code in (" + Chr(39) + "10" + Chr(39) + "," + Chr(39) + "12" + Chr(39) + "," + Chr(39) + "14" + Chr(39) + ")"
            sqlDoc = sqlDoc & "   AND SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT sample_num = 0"
            sqlDoc = sqlDoc & " UNION ALL "
                
            sqlDoc = sqlDoc & "SELECT DISTINCT center_code, name, resdnt, health_num, sample_num"
            sqlDoc = sqlDoc & "     , interface_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE interface_day = " + Chr(39) + strDate + Chr(39)
            '-- Query 추가 본부일경우 본부/남대문/인천데이터만 조회(종양표시자/갑상선의 경우..)
            sqlDoc = sqlDoc & "   AND NOT center_code = " + Chr(39) + "10" + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            sqlDoc = sqlDoc & "   AND not center_code in (" + Chr(39) + "15" + Chr(39) + "," + Chr(39) + "16" + Chr(39) + "," + Chr(39) + "17" + Chr(39) + "," + Chr(39) + "18" + Chr(39) + ")"
            sqlDoc = sqlDoc & "   AND NOT sample_num = 0"
            sqlDoc = sqlDoc & " ORDER BY sample_num ASC"
        Else
            '각 총국 계약자서비스
            sqlDoc = "         SELECT DISTINCT center_code, name, resdnt, health_num, serial_num"
            sqlDoc = sqlDoc & "     , health_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE health_day = " + Chr(39) + strDate + Chr(39)
            sqlDoc = sqlDoc & "   AND center_code = " + Chr(39) + D0COM_CENTERCOD + Chr(39)
            sqlDoc = sqlDoc & "   AND SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            sqlDoc = sqlDoc & " ORDER BY serial_num ASC"
        End If
    End If
    
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
   
    If AdoRs_ORACLE.EOF = True Then
        Set f_subSet_ResultList = Nothing
        RecordChk = False
    Else
        Set f_subSet_ResultList = AdoRs_ORACLE
        RecordChk = True
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
                       
        If Len(strBarno) > 8 Then
            sqlDoc = " SELECT a.per_gumjin_date, a.per_gum_num, a.edpscode, a.result, a.send_date, a.per_name " & _
                    " FROM mdck..gumjin_interface a, mdck..bag_interfacecode b " & _
                    " WHERE substring(a.per_gumjin_date,3,8) = '" & Mid(strBarno, 1, 6) & "'" & _
                    " AND a.per_gum_num = '" & Val(Mid(strBarno, 7)) & "' " & _
                    " AND a.result = '' " & _
                    " AND substring(b.kind,1,1) = 'C' " & _
                    " AND a.edpscode=b.meditem " & _
                    " ORDER BY a.per_gumjin_date, a.per_gum_num "
        Else
            sqlDoc = " SELECT a.EnterDate, b.Status, b.waitseqno, b.MAP2SEQNO, b.DispDesc, b.RVALUEKIND, b.NORMLOW, b.NORMHIGH, b.NORMALVALUE, b.RVALUEKIND , " & _
                    " a.ChartNo, b.GumsaKind, c.sujinname, b.status " & _
                    " FROM medicom..WaitPrsnp a, medicom..jun370_resulttb b, medicom..pewprsnp c, medicom..BAGMAP2PREF d " & _
                    " WHERE a.Chartno = '" & strBarno & "' " & _
                    " AND a.WaitSeqNo = b.WaitSeqNo " & _
                    " AND a.status = '1' " & _
                    " AND d.labno = 4 " & _
                    " AND b.jun370no = d.map2seqno " & _
                    " AND b.status = '0' " & _
                    " AND a.chartno = c.chartno " & _
                    " ORDER BY a.chartno "
        End If
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_WorkList_Barcode = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
        Else
            Set f_subSet_WorkList_Barcode = AdoRs_SQL
            RecordChk = True
        End If
    
        Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)

    
End Function

Private Function f_subSet_SearchList(ByVal strBarcode As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
'    gSql = "select IN_CODE from EXAM_TOC Where RE_RCID = '" & strBarcode & "'"
    
    gSql = "select a.IN_CODE,a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where a.RE_RCID = '" & strBarcode & "'" _
            & " And b.HE_UNID = 'HC-46101'" _
            & " And a.RE_RCID = c.RE_RCID" _
            & " And b.JU_PERID = c.JU_PERID " _
            & " And a.EX_INST = '2'" _
            & " And a.IN_CODE like 'SE%'"

    AdoRs_SQL.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_SearchList = Nothing
    Else
        Set f_subSet_SearchList = AdoRs_SQL
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, intRow  As Integer
    Dim intCol3  As Integer
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 10
    intCol3 = 9
    intCol2 = 1
    intRow = 1
    
    With spdPtList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdRstview
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL)) AND (REMARK <> '1') " & _
             " order by OUT_SEQ, TESTCD_EQP"

             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: ReDim fChannel(adoRS.RecordCount)
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        Dim pTestCode As String
        pTestCode = Replace(Trim(adoRS.Fields("TESTCD")), ",", "','")
        
        gTestCode = gTestCode & "'" & pTestCode & "'" & ","
        
        With spdWorkList
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .Col = intCol:  .ColHidden = True
        End With
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
'        With spdRstview
'            If intRow > .maxrows Then
'                intRow = 1
'                intCol2 = intCol2 + 2
'            End If
'
'            .SetText intCol2, intRow, Trim$(adoRS("TESTNM") & "")
'            intRow = intRow + 1
'
'        End With
        
'        With spdResult2
'            If intCol > .MaxCols Then
'                .MaxCols = .MaxCols + 1
'                .ColWidth(intCol) = 7.5
'            End If
'            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
'        End With
        
        fChannel(intCol - 9) = adoRS.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestCd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"
        f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
        f_typCode(intCnt).strTestCd(f_typCode(intCnt).intCnt) = strTmp
        
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With

    txtPtId.Text = ""
    txtPtNm.Text = ""
    txtSpcNo.Text = ""
    txtSEX.Text = ""
    txtAGE.Text = ""
    txtRcvDt.Text = ""
    txtResultDt.Text = ""
Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub
Private Function f_funGet_CODE(ByVal strOrdcd As String) As String

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE = ""
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestCd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

Private Sub chkOrdCheck_Click()
Dim intRow As Integer
    With spdWorkList
        For intRow = 1 To .maxrows
            .Row = intRow
            .Col = 1
            .Text = IIf(.Text = "0", "1", "0")
        Next intRow
    End With
End Sub

Private Sub cmdAction_Wide_Click()
    Call cmdWide_Click
End Sub

Private Sub cmdWide_Click()
    If spdResult1.Width < 11355 Then
        spdResult1.Width = 11380
        'cmdWide.Caption = "축소모드"
    Else
        spdResult1.Width = 8085
        'cmdWide.Caption = "확장모드"
    End If
End Sub

Private Sub cmdEot_Click()
    comEQP.Output = EOT
End Sub

Private Sub cmdFind_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcNo    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Call cmdAction_Click(2)
    
    intRow = 0
    
    sqlDoc = "Select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  From INTERFACE003" & _
             " Where TRANSDT >= '" & Format(dtpFromDate.Value, "YYYYMMDD") & "'" & _
             "   And TRANSDT <= '" & Format(dtpToDate.Value, "YYYYMMDD") & "'" & _
             "   And EQUIPCD = '" & INS_CODE & "'"
             
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   And SERVERGBN = '' "
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
    End If
    ' 정상
    If chkResult.Value = 1 Then
        sqlDoc = sqlDoc & "   And (JUDGE = '' OR JUDGE IS NULL) "
    End If
    ' H/L
    If chkNoResult.Value = 1 Then
        sqlDoc = sqlDoc & "   And JUDGE <> '' "
    End If
    ' PANIC
    If chkER.Value = 1 Then
        sqlDoc = sqlDoc & "   And PANIC <> '' "
    End If
    ' NEGATIVE
    If chkGonDan.Value = 1 Then
        sqlDoc = sqlDoc & "   And UCASE(RSTVAL) = 'NEGATIVE' "
    End If
    ' POSITIVE
    If chePositive.Value = 1 Then
        sqlDoc = sqlDoc & "   And UCASE(RSTVAL) = 'POSITIVE' "
    End If
    ' KEYWORD
    If txtKeyword.Text <> "" Then
        Select Case cboKeyword.ListIndex
            Case 1: sqlDoc = sqlDoc & "   And NAME LIKE '%" & txtKeyword.Text & "%'  "
            Case 2: sqlDoc = sqlDoc & "   And PNO LIKE '%" & txtKeyword.Text & "%'  "
            Case 3: sqlDoc = sqlDoc & "   And SPCNO LIKE '%" & txtKeyword.Text & "%'  "
            Case 4: sqlDoc = sqlDoc & "   And RSTVAL LIKE '%" & txtKeyword.Text & "%'  "
        End Select
    End If
    
    sqlDoc = sqlDoc & " Order By TRANSDT, SPCNO"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        Do While Not adoRS.EOF
            With spdPtList
            If strSpcNo <> Trim$(adoRS("SPCNO") & "") + Trim$(adoRS("PNO") & "") Then
                    intRow = intRow + 1
                    If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                    .SetText 1, intRow, "0"
                    .SetText 2, intRow, Trim$(adoRS("TRANSDT") & "")
                    .SetText 3, intRow, Trim$(adoRS("NAME") & "")
                    .SetText 4, intRow, Trim$(adoRS("SPCNO") & "")
                End If
                strSpcNo = Trim$(adoRS("SPCNO") & "") + Trim$(adoRS("PNO") & "")
            End With
            adoRS.MoveNext
        Loop
    Else
        Call MsgBox("검색조건을 확인하세요.", vbExclamation, "검색조건확인")
    End If
    
    txtKeyword.Text = ""
    adoRS.Close:    Set adoRS = Nothing
End Sub

Private Sub cmdOrder_Click()
    Dim ii As Integer
    Dim chkRackNo As Variant
    Dim chkPos    As Variant
    Dim strMsg As String

    spdResult1.GetText 7, 1, chkRackNo
    spdResult1.GetText 8, 1, chkPos

    strMsg = ":: 오더전송 준비가 되었습니다." & vbCrLf & vbCrLf & " Rack : " & chkRackNo & " / Pos : " & chkPos & "  부터 오더를 전송하시겠습니까.? "
    
    If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
    
        OrderCnt = 0
        
        With spdResult1
            For ii = 1 To .maxrows
                .Col = 1: .Row = ii
                If .Value = 1 Then
                    .Col = 2
                    If Len(Trim(.Text)) > 0 And .BackColor = vbWhite Then
                        comEQP.Output = ENQ
                        OrderCnt = ii
                        .Col = 1: .Text = 0
                        .Col = 7: .BackColor = vbCyan
                        .Col = 8: .BackColor = vbCyan
                        .Col = 9: .BackColor = vbCyan
                         Exit For
                    End If
                End If
            Next
        End With
    Else
        Exit Sub
    End If
        
End Sub

Private Sub cmdRackNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer
Dim fNum1 As Integer, fNum2 As Integer
Dim intRow1 As Integer

AgainInput:
    fNum1 = 1: fNum2 = 0
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                intRow1 = intRow1 + 1
                .Row = sCnt
                .Col = 1
                If .Value >= 1 Then
                    'If .ActiveCol = 3 Then
                        .Col = 7 '.ActiveCol
                        If intRow1 = (30 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
                        fNum2 = fNum2 + 1
                        .Text = Format(Trim((fNum1 + Val(sNo)) - 1), "0")
                        .Col = 8 '.ActiveCol + 1
                        .Text = fNum2
                    'End If
                End If
            Next sCnt
        End With
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestCd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
    Dim strOld  As Integer
    
    chkOrdCheck.Value = 0
       
    '-- WorkList조회
    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text, mskOrdDate1.Text)
    If RecordChk = False Then
            MsgBox "조회 된 대상자가 없습니다." & vbCrLf & "검진일자를 확인하세요.", vbInformation, Me.Caption
        Exit Sub
    End If
    
    
    
    With spdWorkList
        .maxrows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    intRow = 0
    
    Do Until mAdoRs.EOF
        intIdx = 0
        With spdWorkList
            If strKeyno <> mAdoRs.Fields("SPCNO") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 14
                
                .SetText 1, intRow, 0
                .SetText 2, intRow, mAdoRs("REQDATE")
                .SetText 3, intRow, mAdoRs("PTNM")
                .SetText 4, intRow, mAdoRs("PTID")
                .SetText 5, intRow, mAdoRs("SPCNO")
                .SetText 8, intRow, mAdoRs("SEX")
                .SetText 9, intRow, mAdoRs("AGE")
                
            '-- 검사항목조회
                Set mAdoRs1 = New Recordset
                Set mAdoRs1 = f_subSet_TestList(mAdoRs("REQDATE"), mAdoRs("PTID"))
                
                Do Until mAdoRs1.EOF
                    strEqpCd = f_funGet_CODE(mAdoRs1("TESTCODE"))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then .SetText 9 + itemX.Index, intRow, "V"
                    Set itemX = Nothing
                    mAdoRs1.MoveNext
                Loop
            End If
            strKeyno = mAdoRs("SPCNO")
        End With
        intIdx = intIdx + 1
        mAdoRs.MoveNext
    Loop
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)
    
End Sub

Private Function medGetP(ByVal strText As String, _
                        ByVal intPosition As Integer, _
                        ByVal Delimiter As String) As String
                        
Dim intPos1 As Integer, intPos2 As Integer, i As Integer

    intPos1 = 0: intPos2 = 0
    
    ' intPosition 인수가 1인 경우 For문 Skip
    For i = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    ' 해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, strText, Delimiter)
    If intPos2 = 0 Then intPos2 = Len(strText) + 1
    
    medGetP = Mid$(strText, intPos1, intPos2 - intPos1)
    
    Exit Function
    
ReturnNull:
    medGetP = ""
    
End Function


Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

Private Sub cmdACK_Click()

    comEQP.Output = ACK

End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdExit
        Case Else
    End Select
    
    intRow = 0
    
End Sub

Private Sub cmdClear()
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    
    If tabWork.Tab = 0 Then
        With spdRstview
            .maxrows = 1
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
            .RowHeight(-1) = 14
        End With
        
        With spdWorkList
            .UserColAction = UserColActionSort
            .maxrows = 1
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
            .RowHeight(-1) = 14
        End With
        
        With spdResult1
            .maxrows = 1
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BackColor = vbWhite
            .BlockMode = False
            .RowHeight(-1) = 14
        End With

   Else
        txtPtId = ""
        txtPtNm = ""
        txtSpcNo = ""
        txtSEX = ""
        txtAGE = ""
        txtRcvDt = ""
        txtResultDt = ""
'        chkResult.Value = 1
        
        With spdPtList
            .maxrows = 1
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
            .RowHeight(-1) = 14
        End With
        
        With spdResult2
            .maxrows = 1
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
            .RowHeight(-1) = 14
        End With
    End If
End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Private Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
        
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdStop()
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If comEQP.PortOpen Then comEQP.PortOpen = False
    If comEQP.PortOpen Then
        Call ShowMessage("중지 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdAppend_Click(Index As Integer)
   
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
        
    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String, strBarno     As String, strTime      As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestCd    As String, strTestRst   As String
    Dim strName         As String
    Dim strChartNo      As String
    Dim strHealth       As String
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As Object
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    Dim strResult   As String
    Dim WK_SLEKWA   As String
    Dim WK_SJKEY    As String
    Dim strDate     As String
    Dim pOrd_No As String
    Dim pOrd_Seq_No As String
    Dim strREF          As String
    Dim strReceptNo As String
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"
    
'On Error GoTo ErrorRoutine
    
    Me.MousePointer = 11
    
    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If
    
    With objSpd
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp:         strDate = Trim$(varTmp)
            .GetText 3, intRow, varTmp:         strBarno = Trim$(varTmp)
            .GetText 4, intRow, varTmp:         strChartNo = Trim$(varTmp)
            .GetText 5, intRow, varTmp:         strReceptNo = Trim$(varTmp)
            
            .GetText 1, intRow, varTmp
            
            If strChartNo = "" Then Exit For
            
            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 8 To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp:    strResult = Trim(varTmp)
                            strTestCd = itemX.ListSubItems(1)

                            If Len(strTestCd) = 0 And Len(Trim(strResult)) = 0 Then
                                Exit Sub
                            Else
                                Dim pLabCnt As String
                                Dim pLabsCnt As String
                                Dim pResult  As String
                                
                                .Row = intRow: .Col = intCol
                                
                                pOrd_No = medGetP(.CellTag, 1, "|")
                                pOrd_Seq_No = medGetP(.CellTag, 2, "|")
                                strTestCd = medGetP(.CellTag, 3, "|")
                                
                                pResult = Result_Chang(itemX.Text, strResult)
                                
                                Select Case .ForeColor
                                    Case vbRed
                                        strREF = "H"
                                    Case vbBlue
                                        strREF = "L"
                                    Case Else
                                        strREF = ""
                                End Select
                                
                                If Len(pOrd_No) > 0 And Len(pOrd_Seq_No) > 0 And Len(strTestCd) > 0 Then
                                
                                    sqlDoc = ""
                                    sqlDoc = sqlDoc & vbCr & " UPDATE Lab_List"
                                    sqlDoc = sqlDoc & vbCr & "   Set KulGa = '" & strResult & "'"
                                    sqlDoc = sqlDoc & vbCr & " WHERE ChtNo = '" & strChartNo & "'"
                                    sqlDoc = sqlDoc & vbCr & "   AND CODE = '" & strTestCd & "'"
                                    sqlDoc = sqlDoc & vbCr & "   AND Ymd = '" & strDate & "'"
                                    sqlDoc = sqlDoc & vbCr & "   AND (KulGa IS NULL OR KulGa = '')"

                                    Debug.Print sqlDoc
                                        
                                    AdoCn_SQL.Execute sqlDoc
                                
                                End If
                                
                                lblStatus.Caption = "저장 성공!!"
                                .Row = intRow: .Col = 9: .Text = "결과저장"
                                
                                .BackColor = vbRed
                            End If
                        End If
                                                
                        Set itemX = Nothing
                    End If
                    .Row = intRow: .Col = 1: .Text = 0
                Next

                If strErrMsg = "" Then
                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                             " where SPCNO   = '" & strBarno & "'" & _
                             "   and TRANSDT = '" & Replace(txtRcvDt.Text, "-", "") & "'"
                    AdoCn_Jet.Execute sqlDoc
                Else
                    MsgBox strErrMsg, vbInformation, Me.Caption
                End If
            End If
        Next
    End With
    
    Me.MousePointer = 0

End Sub

Private Function f_delta_chk(ByVal WK_SAMPLE As String, ByVal WK_WORKNM As String, ByVal WK_VALUE As String)
Dim WK_CLVALUE As String, WK_CHVALUE As String, WK_MLVALUE As String, WK_MHVALUE As String
Dim WK_FLVALUE As String, WK_FHVALUE As String, WK_DLVALUE As String, WK_DHVALUE As String, WK_PLVALUE As String, WK_PHVALUE As String
Dim WK_METHOD As String, S_GMSEX As String
Dim WK_STAND As Integer

On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_BarcodeSp() As ADODB.Recordset"

    If Mid(WK_VALUE, 1, 1) = "p" Or Mid(WK_VALUE, 1, 1) = "n" Then
        WK_VALUE = ""
    End If

    Set AdoRs_ORACLE = New ADODB.Recordset

    gSql = "SELECT G13_GMSEX " _
         & "  From GUMSA013 " _
         & " WHERE GUMSA013.G13_SAMPLE = '" & WK_SAMPLE & "' "

    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly

    If AdoRs_ORACLE.RecordCount = 0 Then
    Else
        S_GMSEX = AdoRs_ORACLE("G13_GMSEX")
    End If


'    If AdoRs_ORACLE("WK_STAND") = 0 Then
'        MsgBox "작업 품목인  " + WK_WORKNM + " 의 정상치값이 입력되어 있지 않습니다.!!!", "'F_DELTA_CHECK MESSAGE!'", vbExclamation
'        Exit Function
'    End If
    AdoRs_ORACLE.Close

    
    gSql = "SELECT NVL(GUMSA010.G10_CLVALU,-99999999) WK_CLVALUE,NVL(GUMSA010.G10_CHVALU,-99999999) WK_CHVALUE, " _
         & "       NVL(GUMSA010.G10_MLVALU,-99999999) WK_MLVALUE,NVL(GUMSA010.G10_MHVALU,-99999999) WK_MHVALUE, " _
         & "       NVL(GUMSA010.G10_FLVALU,-99999999) WK_FLVALUE,NVL(GUMSA010.G10_FHVALU,-99999999) WK_FHVALUE, " _
         & "       NVL(GUMSA010.G10_DLVALU,-99999999) WK_DLVALUE,NVL(GUMSA010.G10_DHVALU,-99999999) WK_DHVALUE, " _
         & "       NVL(GUMSA010.G10_PLVALU,-99999999) WK_PLVALUE,NVL(GUMSA010.G10_PHVALU,-99999999) WK_PHVALUE, " _
         & "       NVL(GUMSA010.G10_METHOD ,' ') WK_METHOD" _
         & "  From GUMSA010" _
         & " WHERE GUMSA010.G10_GMPART = '13' " _
         & "   AND GUMSA010.G10_WORKNM = '" & WK_WORKNM & "'  "
    
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.RecordCount = 0 Then
'        MsgBox "검사 기준값을 정의한 TABLE을 읽지 못했습니다..!!!", vbExclamation, "CAUTION!"
        Exit Function
    End If
    
    If AdoRs_ORACLE.EOF Then
    Else
        WK_CLVALUE = AdoRs_ORACLE("WK_CLVALUE"):    WK_CHVALUE = AdoRs_ORACLE("WK_CHVALUE")
        WK_MLVALUE = AdoRs_ORACLE("WK_MLVALUE"):    WK_MHVALUE = AdoRs_ORACLE("WK_MHVALUE")
        WK_FLVALUE = AdoRs_ORACLE("WK_FLVALUE"):    WK_FHVALUE = AdoRs_ORACLE("WK_FHVALUE")
        WK_DLVALUE = AdoRs_ORACLE("WK_DLVALUE"):    WK_DHVALUE = AdoRs_ORACLE("WK_DHVALUE")
        WK_PLVALUE = AdoRs_ORACLE("WK_PLVALUE"):    WK_PHVALUE = AdoRs_ORACLE("WK_PHVALUE")
        WK_METHOD = AdoRs_ORACLE("WK_METHOD")
        
        '----------------------------------------
        '-- COMMON EVALUATE CHECKING
        '----------------------------------------
        If WK_CLVALUE > -99999999 Then
            G_CLVALU = WK_CLVALUE
            G_CHVALU = WK_CHVALUE
            
            If WK_VALUE = "" Then
                G_EVALUATE = ""
            Else
                If WK_VALUE < WK_CLVALUE Then
                    G_EVALUATE = "L"
                ElseIf WK_VALUE > WK_CHVALUE Then
                    G_EVALUATE = "H"
                Else
                    G_EVALUATE = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- MAN EVALUATE CHECKING
        '----------------------------------------
        If WK_MLVALUE > -99999999 And S_GMSEX = "M" Then
            G_CLVALU = WK_MLVALUE
            G_CHVALU = WK_MHVALUE
            
            If WK_VALUE = "" Then
                G_EVALUATE = ""
            Else
                If WK_VALUE < WK_MLVALUE Then
                    G_EVALUATE = "L"
                ElseIf WK_VALUE > WK_MHVALUE Then
                    G_EVALUATE = "H"
                Else
                    G_EVALUATE = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- FEMALE EVALUATE CHECKING
        '----------------------------------------
        If WK_FLVALUE > -99999999 And S_GMSEX = "F" Then
            G_CLVALU = WK_FLVALUE
            G_CHVALU = WK_FHVALUE
            
            If WK_VALUE = "" Then
                G_EVALUATE = ""
            Else
                If WK_VALUE < WK_FLVALUE Then
                    G_EVALUATE = "L"
                ElseIf WK_VALUE > WK_FHVALUE Then
                    G_EVALUATE = "H"
                Else
                    G_EVALUATE = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- PANIC CHECKING
        '----------------------------------------
        If WK_PLVALUE > -99999999 And S_GMSEX = "F" Then
            
            If WK_VALUE = "" Then
                G_PANIC = ""
            Else
                If WK_VALUE < WK_PLVALUE Then
                    G_PANIC = "L"
                ElseIf WK_VALUE > WK_PHVALUE Then
                    G_PANIC = "H"
                Else
                    G_PANIC = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- DELTA CHECKING
        '----------------------------------------
        If WK_DLVALUE > -99999999 And S_GMSEX = "F" Then
            If WK_VALUE = "" Then
                G_DELTA = ""
            Else
                If WK_VALUE < WK_DLVALUE Then
                    G_DELTA = "L"
                ElseIf WK_VALUE > WK_DHVALUE Then
                    G_DELTA = "H"
                Else
                    G_DELTA = ""
                End If
            End If
        End If
    End If
    
    Set AdoRs_ORACLE = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

End Function

Private Sub F_HEL30203_UPDATE(ByVal WK_SJKEY As String, ByVal WK_WORKNM As String, ByVal WK_VALUE As String)
Dim WK_CNT As Integer
'------------------------------------------------------------
'-- 종합/일반/채용/암환자에 대한 자료 UPDATE (면역혈청 검사)
'------------------------------------------------------------

'-- WK_SJKEY/WK_WORKNM/WK_VALUE

On Error GoTo HEL323_UPDATE

    CallForm = "frmComm - Private Sub F_HEL30203_UPDATE()"

    Set Ado323 = New ADODB.Recordset
    
    gSql = "SELECT COUNT(H33_SJKEY) WK_CNT" _
         & "  From HEL30203 " _
         & " Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    
    Ado323.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If Ado323.RecordCount = 0 Then
            MsgBox "건강검진 환자에 대한 종합검진 면역혈청 자료를 찾지 못했습니다.!!!", vbExclamation, "HEL30203"
            Return
    Else
        WK_CNT = Ado323("WK_CNT")
    End If
    
    Ado323.Close
    Set Ado323 = Nothing
    
    If WK_CNT < 1 Then
        MsgBox "건강검진 환자에 대한 종합검진 면역혈청 자료를 찾지 못했습니다.!!!", vbExclamation, "HEL30203"
        Return
    End If
    
    If WK_WORKNM = "uTSH" Then                         '-- TSH
        gSql = " Update HEL30203 " _
             & "    Set H33_TSHTSH = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
             
    ElseIf WK_WORKNM = "FreeT4" Then                   '-- FREE T4
        gSql = " Update HEL30203 " _
             & "    Set H33_FREET4 = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
       
    ElseIf WK_WORKNM = "HBeAg 정" Then                 '-- B형 간염E항원(HBe-Ag)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        gSql = " Update HEL30203 " _
             & "    Set H33_HBEAG = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HBeAb 정" Then                '-- Anti-HBe  B형 간염E항체(HEe-Ag)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_ANHBE = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HBsAg 정" Then           '-- B형 간염항원(HBs-Ag)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_HBSAG = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
        
    ElseIf WK_WORKNM = "HBsAb 정" Then           '-- B형 간염항체(Anti-HBs)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_ANHBS = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HCV Ab 정" Then         '-- C형 간염항체(HCV-Ab)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_HCVAB = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "VDRL(Quan)" Then          '-- VDRL 매독
        If WK_VALUE = "Non-Reactive" Then
            WK_VALUE = "음성"
        End If
        
        gSql = " Update HEL30203 " _
             & "    Set H33_VDRL = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "VDRL(Qual)" Then          '-- VDRL 매독
        If WK_VALUE = "Non-Reactive" Then
            WK_VALUE = "음성"
        End If
        
        gSql = " Update HEL30203 " _
             & "    Set H33_VDRL = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "AIDS" Then                '-- AIDS
        gSql = " Update HEL30203 " _
             & "    Set H33_AIDS = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HIV(AIDS)" Then               '-- AIDS
        gSql = " Update HEL30203 " _
             & "    Set H33_AIDS = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "AFP(정밀)" Then            '-- AFP
        gSql = " Update HEL30203 " _
             & "    Set H33_AFPAFP = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CEA" Then                  '-- CEA
        gSql = " Update HEL30203 " _
             & "    Set H33_CEACEA = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "PSA" Then                  '-- PSA
        gSql = " Update HEL30203 " _
             & "    Set H33_PSA = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "ASO(Quan)" Then              '-- ASO
        If WK_VALUE >= "220" Then
           WK_VALUE = "양성"
        Else
           WK_VALUE = "음성"
        End If
        gSql = " Update HEL30203 " _
             & "    Set H33_ASO = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "RA(Quan)" Then             '-- RA-FACTOR
        If WK_VALUE >= "20" Then
           WK_VALUE = "양성"
        Else
           WK_VALUE = "음성"
        End If
        gSql = " Update HEL30203 " _
             & "    Set H33_RAFACT = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CRP(Quan)" Then         '-- CRP
        If WK_VALUE >= "0.5" Then
           WK_VALUE = "양성"
        Else
           WK_VALUE = "음성"
        End If
        gSql = " Update HEL30203 " _
             & "    Set H33_CRP = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CA125" Then         '-- CA125
        gSql = " Update HEL30203 " _
             & "    Set H33_CA125 = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CA19-9" Then            '-- CA19-9
        gSql = " Update HEL30203 " _
             & "    Set H33_CA199 = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "THCY" Then          '-- homocysteine
        gSql = " Update HEL30203 " _
             & "    Set H33_HOMO = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
     End If

    AdoCn_ORACLE.Execute (gSql)
        
HEL323_UPDATE:
    Set Ado323 = Nothing
    Call ErrMsgProc(CallForm)

End Sub

Public Function CheckSum_ECi_Tx(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For i = 1 To intValueLength
        intCheck = intCheck + Asc(Mid(strPrmValue, i, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        CheckSum_ECi_Tx = "0" & strCheck
    Else
        CheckSum_ECi_Tx = Right(strCheck, 2)
    End If

End Function

Public Function LenA(strPrmString As String) As Integer

    Dim i                   As Integer
    Dim intStrLen           As Integer
    Dim intAnsiStrLen       As Integer
    Dim strTemp             As String
    
    intStrLen = Len(strPrmString)
    For i = 1 To intStrLen
        strTemp = Mid(strPrmString, i, 1)
        
        Select Case AscW(strTemp)
        Case 0 To 255
            intAnsiStrLen = intAnsiStrLen + 1
        
        Case Else
            intAnsiStrLen = intAnsiStrLen + 2
        
        End Select
    Next
    
    LenA = intAnsiStrLen

End Function

Private Sub cmdENQ_Click()
    
    comEQP.Output = ENQ

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*   생년월일로 나이를 계산                      *
'*   passport_id   :  생년월일 변환대상 data     *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function D0SUB_BIRTHDAY(ByVal PassPort_Id As String) As String

    Dim yy       As String
    Dim age      As Integer

    On Error GoTo D0SUB_BIRTHDAY
    
    Select Case Len(left$(PassPort_Id, 6))
        Case 2, 3, 4, 5
            yy = left$(PassPort_Id, 2) & "-01-01"
            age = DateDiff("yyyy", yy, Now)

        Case 6
            yy = Format$(PassPort_Id, "##-##-##")
            age = DateDiff("yyyy", yy, Now)
    End Select
        
    D0SUB_BIRTHDAY = Trim$(Str$(age))
    
    On Error GoTo 0
    Exit Function
D0SUB_BIRTHDAY:
    Resume Next
        
End Function

Function D0SUB_SETCENTER(para As String) As Variant

    Select Case Trim$(para)
        Case "10": D0SUB_SETCENTER = " 본사 "
        Case "12": D0SUB_SETCENTER = "남대문"
        Case "14": D0SUB_SETCENTER = " 인천 "
        Case "15": D0SUB_SETCENTER = " 대전 "
        Case "16": D0SUB_SETCENTER = " 광주 "
        Case "17": D0SUB_SETCENTER = " 대구 "
        Case "18": D0SUB_SETCENTER = " 부산 "
        Case "본사": D0SUB_SETCENTER = "10"
        Case "남대문": D0SUB_SETCENTER = "12"
        Case "인천": D0SUB_SETCENTER = "14"
        Case "대전": D0SUB_SETCENTER = "15"
        Case "광주": D0SUB_SETCENTER = "16"
        Case "대구": D0SUB_SETCENTER = "17"
        Case "부산": D0SUB_SETCENTER = "18"
    End Select

End Function

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
'    If Index = 2 Or Index = 3 Then
        With spdWorkList
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
'    Else
'        With spdWorkList
'            For intRow = 1 To .maxrows
'                .GetText 2, intRow, varTmp
'                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
'            Next
'        End With
'    End If
    
End Sub

Private Sub cmdStartNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 8:       .Text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Function f_subSet_TestList(ByVal strDate As String, ByVal strBarcode As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    
    sqlDoc = sqlDoc & vbCrLf & "SELECT"
    sqlDoc = sqlDoc & vbCrLf & "       L.Ymd AS REQDATE ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.Ymd + '^' + L.ChtNo AS SPCNO ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.ChtNo AS PTID ,"
    sqlDoc = sqlDoc & vbCrLf & "       P.Nm AS PTNM ,"
    sqlDoc = sqlDoc & vbCrLf & "       '' AS SEX,"
    sqlDoc = sqlDoc & vbCrLf & "       P.Age AS AGE ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.NO AS ORDNO ,"
    sqlDoc = sqlDoc & vbCrLf & "       L.CODE AS TESTCODE"
    sqlDoc = sqlDoc & vbCrLf & "  FROM Lab_List L,"
    sqlDoc = sqlDoc & vbCrLf & "       Person p"
    sqlDoc = sqlDoc & vbCrLf & " Where p.CODE = L.ChtNo"
    sqlDoc = sqlDoc & vbCrLf & "   AND L.Ymd = '" & strDate & "'"
    sqlDoc = sqlDoc & vbCrLf & "   AND L.ChtNo = '" & strBarcode & "'"
    sqlDoc = sqlDoc & vbCrLf & "   AND (L.KulGa = '' OR L.KulGa IS NULL)"
    sqlDoc = sqlDoc & vbCrLf & "   AND L.CODE IN (" & Mid(gTestCode, 1, Len(gTestCode) - 1) & " ) "
        
    AdoRs_SQL.Open sqlDoc, AdoCn_SQL, adOpenStatic, adLockReadOnly
    
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_TestList = Nothing
    Else
        Set f_subSet_TestList = AdoRs_SQL
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Function f_subSet_TestList1(ByVal strBarcode As String, ByVal strOld As Integer, ByVal strSex As Integer)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
   If optJong.Value = True Then
        If strOld >= 40 And strSex = 1 Then
                   gSql = "SELECT * from tb_msmedhsa" & vbLf
            gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA01','ABA02','AFA06')" & vbLf
        Else
                   gSql = "SELECT * from tb_msmedhsa" & vbLf
            gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA01','ABA02')" & vbLf
        End If
   Else
        If Mid(strBarcode, 1, 1) = "7" Then
                    gSql = "SELECT * from tb_msmedhsa" & vbLf
             gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
             gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
             gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
             gSql = gSql & "   AND medi_item_code in ('AFA01','ABA02','ACA01')" & vbLf
         ElseIf Mid(strBarcode, 1, 1) = "1" Then
            If strOld >= 40 And strSex = 1 Then
                       gSql = "SELECT * from tb_msmedhsa" & vbLf
                gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA02','AFA06','ABA01')" & vbLf
            Else
                       gSql = "SELECT * from tb_msmedhsa" & vbLf
                gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA02','ABA01')" & vbLf
            End If
         End If
    End If
    
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.EOF = True Then
        Set f_subSet_TestList1 = Nothing
        RecordChk = False
        Exit Function
    Else
        Set f_subSet_TestList1 = AdoRs_ORACLE
        RecordChk = True
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function
Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String, strDate As String, strCNT As String
    Dim strSex      As String, strOld   As String, strArea   As String, strage As String
    Dim strTmpSex   As Integer
    
    Dim strEqpCd    As String
       
    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    Dim strRack_tmp As String
    
    blnFlag = False

    With spdWorkList
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 3, intRow1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strSPid = Trim$(varTmp)
                .GetText 2, intRow1, varTmp:    strDate = Trim$(varTmp)
                .GetText 5, intRow1, varTmp:    strCNT = Trim$(varTmp)
                .GetText 8, intRow1, varTmp:    strage = Trim$(varTmp)
                .GetText 9, intRow1, varTmp:    strSex = Trim$(varTmp)
                
                .Row = intRow1:
                .Col = 2: .ForeColor = vbRed
                .Col = 3: .ForeColor = vbRed
                .Col = 4: .ForeColor = vbRed
                
                intRow2 = f_funGet_SpreadRow(spdResult1, 4, strSPid)
                
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 4, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 14
                        intRow2 = spdResult1.maxrows
                    End If
                    
                    blnFlag = False
                    Set mAdoRs = f_subSet_TestList(strDate, strSPid)
                    If Len(strSPid) > 0 Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("TESTCODE"))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 9
                                spdResult1.CellTag = mAdoRs("ORDNO") & "|" & mAdoRs("PTID") & "|" & mAdoRs("TESTCODE")
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                        spdResult1.SetText 2, intRow2, strDate
                        spdResult1.SetText 3, intRow2, strSPnm
                        spdResult1.SetText 4, intRow2, strSPid:  'spdResult1.tag = strCNT
                        
                        spdResult1.SetText 5, intRow2, strSex
                        
                        spdResult1.SetText 6, intRow2, Format(Now, "MMDD") & "-" & Format(txtSeq.Text, "00")
                        txtSeq.Text = txtSeq.Text + 1

                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                    
                    .SetText 1, intRow2, ""
                End If
                
                spdResult1.SetText 1, intRow2, "1"

                .SetText 1, intRow1, ""
            End If
        Next
        

    End With
    
    With spdResult1
        iCnt = 1
        .GetText 1, 1, varChk
        .GetText 2, 1, varBar
        varNum = 1
        If Trim(varChk) = "1" And Trim(varBar) <> "" Then
            For iRow = 1 To .maxrows
                strRack_tmp = Format(varNum, "0")
                
                
                .SetText 7, iRow, strRack_tmp
                .SetText 8, iRow, ((iCnt Mod 11) + 1) - 1
                
                
                
                iCnt = iCnt + 1
                If (iCnt Mod 11) = 0 Then
                    varNum = varNum + 1
                    iCnt = 1
                End If
                
                DoEvents
                
            Next
        End If
    End With

End Sub

Private Sub cmdWorkList1_Click()
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcNo    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    sqlDoc = "select * from Worklist" _
              & " where workdate = '" & mskOrdDate.Text & "'"
                 
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdWorkList
            If strSpcNo <> Trim$(adoRS(1) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 14
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(2) & "")
                .SetText 3, intRow, Trim$(adoRS(3) & "")
                .SetText 4, intRow, Trim$(adoRS(1) & "")
                
                spdWorkList.Row = intRow
                spdWorkList.Col = 2
                spdWorkList.BackColor = &HEDD3CD
                spdWorkList.Col = 3
                spdWorkList.BackColor = &HEDD3CD
                spdWorkList.Col = 4
                spdWorkList.BackColor = &HEDD3CD
                
                .BlockMode = True
                .Row = intRow
                .Col = 2
                .BackColor = &HEDD3CD
                .Col = 3
                .BackColor = &HEDD3CD
                .Col = 4
                .BackColor = &HEDD3CD
                .Col = 1
                .Position = PositionUpperLeft
                .Action = ActionActiveCell
                .Action = ActionGotoCell
                 If .maxrows > 1 Then
                    .Row = 1:       .Row2 = .maxrows - 1
                    .Col = 1:       .Col2 = .MaxCols
                    .BackColor = vbWhite
                End If
                .BlockMode = False
            End If
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing

End Sub

Private Sub comEQP_OnComm()
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim brStr As String
    Dim sStxCheck As Integer, sEtxCheck As Integer, sCrcheck As Integer
    Dim com_sTemp As String
    Dim ii As Integer, jj As Integer
    Dim MHead  As String, Pinfo As String
    Dim PatientID As String
    
    Dim Orderoutput As String
    Dim OutPutData  As String
    Dim Rev As Long
    Dim Test_Cd() As String, strPid()    As String, strPnm() As String
    Dim sRow As Integer
    Dim oPatNo As String
    Dim oRackNo As String
    Dim oPosNo As String
    Dim oIdNo As String
    
    Dim adoRS As ADODB.Recordset
    Dim sqlDoc As String
    Dim itemX As ListItem
    Dim strEqpCd1 As String
    
    Dim varTmp  As Variant
    Dim intCol  As Integer
    Dim strLevel() As String
    Dim sDate   As String
    Dim sOCnt   As Integer
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            brStr = ""
            brStr = comEQP.Input
            
            
            Debug.Print "[DATA] " & brStr
            
            txtResult.Text = txtResult.Text + brStr
            
            If left$(brStr, 1) = ENQ Or Mid(brStr, 2) = ENQ Or Mid(brStr, 3) = ENQ Then '광주추가(Mid(brStr, 2) = ENQ)
                Debug.Print "[E411] " & ENQ
                comEQP.Output = ACK
                Debug.Print "[HOST] " & ACK
                fRcvString = ""
                Exit Sub
            End If

            If left$(brStr, 1) = ENQ Or Mid(brStr, 2) = ENQ Or Mid(brStr, 3) = ENQ Or Mid(brStr, 4) = ENQ Or Mid(brStr, 5) = ENQ Then '광주추가(Mid(brStr, 2) = ENQ)
                Debug.Print "[E411] " & ENQ
                comEQP.Output = ACK
                Debug.Print "[HOST] " & ACK
                fRcvString = ""
                Exit Sub
            End If
            
            For ii = 1 To Len(brStr)
                fRcvString = fRcvString + Mid(brStr, ii, 1)
            Next ii
            
            spdResult1.Col = 1: spdResult1.Row = OrderCnt
            If Len(Trim(spdResult1.Text)) > 0 Then
                '-- Ordering
                If left$(brStr, 1) = ACK Then
                    Debug.Print "[E411] " & ACK
                    Select Case SendCount
                        Case 0      'Message Header
                            sDate = Format(Now, "yyyymmddhhmmss")
                            MHead = "1H|\^&|" & sDate & vbCr & ETX
                            comEQP.Output = STX & MHead & MakeCS(MHead) & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & MHead & MakeCS(MHead) & vbCr & vbLf
                            SendCount = SendCount + 1
                            MHead = ""
                        Case 1      'patient information
                        
                            spdResult1.Col = 4: oIdNo = spdResult1.Text
                            
                            Pinfo = "2P|1|" & oIdNo & vbCr & ETX
                            comEQP.Output = STX & Pinfo & MakeCS(Pinfo) & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & Pinfo & MakeCS(Pinfo) & vbCr & vbLf
                            
                            SendCount = SendCount + 1
                            Pinfo = ""
                        Case 2      'Test Order
                            SendCount = SendCount + 1
                            spdResult1.Row = OrderCnt
                            spdResult1.Col = 2
                            
                            oPatNo = spdResult1.Text
                            
                            Dim strOrderKey As String
                                                       
                            spdResult1.Col = 4: oIdNo = spdResult1.Text
                            spdResult1.Col = 6: strOrderKey = spdResult1.Text
                            spdResult1.Col = 7: oRackNo = spdResult1.Text
                            spdResult1.Col = 8: oPosNo = spdResult1.Text
                            
                            With spdResult1
                                sOCnt = 0
                                Orderoutput = ""
                                For intCol = 7 To .MaxCols
                                    spdResult1.GetText intCol, 0, varTmp
                                    If Trim$(varTmp) = "" Then Exit For
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        spdResult1.Col = intCol:
                                        If spdResult1.BackColor = &HC6FEFF Then
                                            sOCnt = sOCnt + 1
                                           
                                            If Orderoutput = "" Then
                                                Orderoutput = "^^^" & Trim(itemX.tag) & "^1"
                                            Else
                                                Orderoutput = Orderoutput & "\^^^" & Trim(itemX.tag) & "^1"
                                            End If

                                        End If
                                     End If
                                    Set itemX = Nothing
                                Next intCol
                            End With
                            
                            Orderoutput = "3O|1|" & strOrderKey & "|" & "^" & oRackNo & "^" & oPosNo & "|" & Orderoutput & "|R||||||A||||Serum" & vbCr & ETX
                            Rem Orderoutput = "3O|1|" & oIdNo & "|" & "^" & oRackNo & "^" & oPosNo & "|" & Orderoutput & "|R||||||A||||Serum" & vbCr & ETX
                            OutPutData = STX & Orderoutput & MakeCS(Orderoutput) & vbCr & vbLf
                            comEQP.Output = OutPutData
                            
                            spdResult1.Col = 9: spdResult1.Text = "오더전송"
                            Debug.Print "[HOST] " & OutPutData
                        Case 3      'Message Terminator
                            SendCount = SendCount + 1
                            
                            Orderoutput = "4L|1|N" & vbCr & ETX
                            OutPutData = STX & Orderoutput & MakeCS(Orderoutput) & vbCr & vbLf
                            
                            comEQP.Output = OutPutData
                            Debug.Print "[HOST] " & OutPutData
                        Case Else
                            comEQP.Output = EOT
                            Debug.Print "[HOST] " & EOT
                            '-- 오더내역이 남아있는지 체크

                            With spdResult1
                                For ii = OrderCnt To .maxrows
                                    .Col = 1: .Row = ii
                                    If .Value = 1 Then
                                        .Col = 2
                                        If .BackColor = vbWhite Then
                                            .Col = 1: .Text = 0

                                            .Col = 7: .BackColor = vbCyan
                                            .Col = 8: .BackColor = vbCyan
                                            .Col = 9: .BackColor = vbCyan
                                            
                                            OrderCnt = OrderCnt + 1
                                            .Col = 2
                                            If Len(Trim(.Text)) > 0 Then
                                                .Row = ii '+ 1
                                                If Len(Trim(.Text)) > 0 Then
                                                    comEQP.Output = ENQ
                                                    Debug.Print "[END_HOST] " & ENQ
                                                End If
                                                SendCount = 0
                                                Exit For
                                            End If
                                        End If
                                    End If
                                    Me.Enabled = True
                                Next
                            End With
                            SendCount = 0
                    End Select
                End If
            End If
                        
            sStxCheck = InStr(fRcvString, STX)
            sEtxCheck = InStr(fRcvString, ETX)
            sCrcheck = InStr(fRcvString, vbCr)
            
            If sStxCheck <> 0 And sEtxCheck <> 0 And sCrcheck <> 0 Then
                Call psDataDefine(fRcvString, fChannel(), spdResult1)
                Debug.Print fRcvString
                fRcvString = ""
            End If
            
        Case comEvSend
        
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) 변경 감지"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) 변경 감지"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) 변경 감지"
        Case comEvRing
            strEVMsg = " 전화 벨이 울리는 중"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) 감지"

        ' 오류 메시지
        Case comBreak
            strERMsg = " 중단 신호 수신"
        Case comCDTO
            strERMsg = " 반송파 검출 시간 초과"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) 시간 초과"
        Case comDCB
            strERMsg = " 포트에 대한 장치 제어 블록(DCB) 검색 중 예기치 못한 오류"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) 시간 초과"
        Case comFrame
            strERMsg = " 프레이밍 오류"
        Case comOverrun
            strERMsg = " 패리티 오류"
        Case comRxOver
            strERMsg = " 수신 버퍼 초과"
        Case comRxParity
            strERMsg = " 패리티 오류"
        Case comTxFull
            strERMsg = " 전송 버퍼에 여유가 없음"
        Case Else
            strERMsg = " 알 수 없는 오류 또는 이벤트"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
        
End Sub

Private Sub ComReceive(ByRef RecData As String)
    
'------------------------------세미양방향일 경우 사용함---------------------------------------------
'---------------------------------------------------------------------------------------------------
Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    Dim intPos0 As Integer, intPos1 As Integer, intPos2 As Integer

    Dim age As String
    Dim i As Integer
    Dim tt As Boolean
    Dim sHead       As String
    Dim sPInfo      As String
    Dim sRtypeId    As String  'Record Type ID(1)
    Dim sSNumber    As String  'Sequence Number(6)
    '---[Specimen ID]----------------------------
    Dim sSampleNo   As String  'Sample No(5)
    Dim sSampleId   As String  'Sample ID(13)
    Dim sSampleType As String  'Sample Type(1)
    Dim sRackId     As String  'Rack Id(5)
    Dim sPositionNo As String  'Position No(1)
    '--------------------------------------------
    Dim sSpecimenID As String  'Specimen ID(2)
    '---[Universal Test Id]----------------------
    Dim sAppCode    As String  'Application Code(3)
    Dim sIdc        As String  'Inc,Dec or Cir(3)
    '--------------------------------------------
    Dim sPriority   As String  'Priority(1)
    Dim sRDateTime  As String  'Requested/Ordered Date and Time
    Dim sSDateTime  As String  'Specimen Collection Date and Time(14)
    Dim sCEndTime   As String  'Collection End Time
    Dim sCvolume    As String  'Collection Volume
    Dim sCId        As String  'Collection Id
    Dim sACode      As String  'Action Code(1)
    Dim sDCode      As String  'Danger Code
    Dim sRcinfo     As String  'Relevant Clinical Information(7)
    Dim sDtSpeR     As String  'Date/Time Specimen Received
    Dim sSpeDesc    As String  'Specimen Descriptor(2)
    Dim sOrderPh    As String  'Ordering Physician
    Dim sPtNum      As String  'Physician's Telephone Number
    Dim sUserF1     As String  'User Field No1(6)
    Dim sUserF2     As String  'User Field No2(104)
    Dim sLaboF1     As String  'Laboratory Field No.1
    Dim sLaboF2     As String  'Laboratory Field No.2
    Dim sDtRr       As String  'Date/Time Result(14)
    Dim sIccs       As String  'Instrument Charge to Computer System
    Dim sIsId       As String  'Instrument Section ID
    Dim sReportT    As String  'Report Types(1)
    Dim ii As Integer
    Dim sTempid As String
    Dim Orderoutput As String
    Dim OutPutData As String
    Dim Testcd As String, sOrderLst As String
    Dim Loop_count As Integer, pDoCount, pChnoCount As Integer
    Dim SEX As String
    Dim intldx As Integer
    Dim sCrLfCheck As Integer
    Dim strOrder As String
        
    Dim SendBuffD           As String           'data
    
    Dim Lencheck
    Dim Specimen            As String

    Static OrgMsg As String

    strRec = RecData
    Print #1, strRec;

    Debug.Print strRec

    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        Select Case strBuff
            Case STX
                    sStxCheck = InStr(strBuff, STX)
            Case ETX

            Case ETB
                    If Mid(f_strBuffer, intIdx, 2) = vbCrLf Then
                        f_strBuffer = left(f_strBuffer, Len(f_strBuffer) - 2) 'Remove CR & LF
                    End If
                    cntCheckSum = cntCheckSum + 1
                    flgETB = True
            Case vbCr

            Case vbLf
                    sCrLfCheck = InStr(strBuff, vbLf)
                    If sCrLfCheck <> 0 And sCrLfCheck <> 0 Then
                        Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
                        comEQP.Output = ACK
                        GoSub ClearReceiveData
                    End If
            Case ENQ
                    comEQP.Output = ACK
            Case ACK
                    Dim varTmp      As Variant
                    Dim intRow      As Integer, intCol  As Integer
                    Dim strBarno    As String, strTest  As String
                    Dim strRack     As String, strCup   As String
                    Dim intCnt1      As Integer
                    Dim itemX       As ListItem

                    With spdResult1
                        For intRow = 1 To .maxrows
                            .Row = intRow
                            .Col = 2
                            If .BackColor = vbWhite Then
                                sAppCode = ""
                                intCnt1 = 0
                                .GetText 2, intRow, varTmp: strBarno = Trim$(varTmp)
                                .GetText 5, intRow, varTmp: strRack = Trim$(varTmp)
                                .GetText 6, intRow, varTmp: strCup = Trim$(varTmp)
                                'sRackId = Format$(strRack, "00")
                                For intCol = 7 To .MaxCols
                                    spdResult1.GetText intCol, 0, varTmp
                                    If Trim$(varTmp) = "" Then Exit For
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        spdResult1.Col = intCol:    'spdResult1.Row = OrderCnt
                                        If spdResult1.BackColor = &HC6FEFF Then
                                            If SendBuffD = "" Then
                                                SendBuffD = "^^^" & Trim(itemX.tag)
                                            Else
                                                SendBuffD = SendBuffD & "\^^^" & Trim(itemX.tag)
                                            End If
                                        End If
                                    End If
                                    Set itemX = Nothing
                                Next intCol
                                If Or_Seq = 5 And SendBuffD <> "" Then
                                    .Row = intRow
                                    .Col = 2: .BackColor = vbCyan
                                    .Col = 3: .BackColor = vbCyan
                                    .Col = 4: .BackColor = vbCyan
                                End If
                                Exit For
                            End If
                        Next intRow
                        
                        If intRow >= .maxrows Then
                            Timer1.Enabled = False
                        End If
                    End With
                    
                    Select Case Or_Seq
                           Case 1   ' Send Header
                                    sSDateTime = Format(Now, "YYYYMMDDHHMMSS")
                                    SendBuffW = Or_Seq & "H|\^&|||Elecsys2010^3.60^9501^H1P1O1R1C1Q1L1M1|||||||P|1|20000617121401" & vbCr & Chr(3)
                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                                    comEQP.Output = SendBuffT
                                    Debug.Print "HOST ==>" & SendBuffT
                                    Or_Seq = Or_Seq + 1

                           Case 2   ' Send Patient Information
                                    SendBuffW = Or_Seq & "P|1||" & strBarno & "||" & vbCr & Chr(3)
                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                                    comEQP.Output = SendBuffT
                                    Debug.Print "HOST ==>" & SendBuffT
                                    Or_Seq = Or_Seq + 1

                           Case 3   ' Send Order Record
                       
                                    SendBuffW = Or_Seq & "O|1|" & strBarno & "||" & SendBuffD & "|||||||A||||||||||||||O" & vbCr & Chr(3)

                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf

                                    comEQP.Output = SendBuffT
                                    strBarno = ""
                                    strOrder = ""
                                    Debug.Print "HOST ==>" & SendBuffT
                                    Or_Seq = Or_Seq + 1

                           Case 4   ' Send Message Terminator
                                    SendBuffW = Or_Seq & "L|1" & vbCr & Chr(3)
                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                                    comEQP.Output = SendBuffT
                                    Debug.Print "HOST ==>" & SendBuffT
                                    Or_Seq = Or_Seq + 1
                           Case 5   ' Send EOT
                                    Or_Seq = 1
                                    SendBuffD = ""
                                    comEQP.Output = EOT
                                    Debug.Print "HOST ==>" & EOT
                                    If intRow < spdResult1.maxrows Then
                                        Timer1.Interval = 200
                                        Timer1.Enabled = True
                                    End If
                    End Select
'
            Case NAK
                    comEQP.Output = ACK
            Case EOT

            Case Else
                    f_strBuffer = f_strBuffer + strBuff
        End Select
    Next

    Exit Sub
ClearReceiveData:
    sStxCheck = 0
    sEtxCheck = 0
    ReceiveData = ""
    cntField_ = 0
    cntRepeat_ = 0
    cntComponent_ = 0
    cntEscape_ = 0
    cntSlash_ = 0
    f_strBuffer = ""
    Return

End Sub

Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brspread As Object)
    Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
    Dim Channel_No  As String       ' 문자형 변수
    Dim Patiant_No  As String       ' 환자번호
    Dim pGrid_Point As Integer      ' 해당 검사자 Point
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    '-------------------------------' 임시 변수들.....
    Dim sDeCnt      As Integer
    Dim pDoCount    As Integer
    Dim Loop_count  As Integer
    Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
    Dim itemX As ListItem
    Dim strRstval       As String
    Dim strRefVal       As String
    Dim FunStr As String
    Dim sqlDoc  As String
    Dim intCol As Integer
    Dim Test_Cd() As String, strPid()    As String, strPnm() As String
    Dim Rev As Long
    Dim ii As Integer
    Dim tmpTstCd As String
    Dim strLevel() As String
    Dim chkPos  As Variant
    Dim strResult As String
    Dim strBarno    As String, strSPid  As String, strSPnm   As String
    Dim strSex      As String, strOld   As String, strArea   As String
    Dim varTmp  As Variant
    Dim strDate As String, strTime  As String, sqlRet   As Integer
    Dim strResultTmp As String
    Dim tmpRefVal   As String
    Dim tmpChrResult As String
    Dim varResult   As Variant
    
    On Error GoTo errDefine
    
    sRstText = brbarcd
    '------------------------------<<< fElecsys2010() 배열 Clear 한다.         >>>----------
    For Loop_count = 1 To 100: fElecsys2010(Loop_count) = "": Next Loop_count
    '------------------------------<<< fElecsys2010() 배열에 구분하여 넣는다.  >>>----------
        
    pDoCount = 0
'    sRstText = Mid(sRstText, STX)
    sRstText = Mid(sRstText, InStr(fRcvString, STX))
    Do While InStr(sRstText, "|") > 0
        pDoCount = pDoCount + 1
        fElecsys2010(pDoCount) = Text_Redefine(sRstText, "|")
        sRstText = Mid$(sRstText, InStr(sRstText, "|") + 1)   ' 구분자가 "|" 이다....
        If pDoCount > 99 Then
            sRstText = ""
            Exit Do
        End If
    Loop
   
    sRstText = ""
    If Mid$(fElecsys2010(1), 3, 1) = "H" Then          ' "H" Head Message Display
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "P" Then      ' "P" Patiant Information Data Process
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "C" Then
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "Q" Then      ' "Q" Patiant Information Data Process
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "O" Then      ' "O" Order Data Process
        comEQP.Output = ACK
        PatientID = fElecsys2010(3)
        pDoCount = 0
        Do While InStr(fElecsys2010(4), "^") > 0
            pDoCount = pDoCount + 1
            Select Case pDoCount
                Case 1:    PatientSeq = Text_Redefine(fElecsys2010(4), "^")
                Case 2:    PatientRack = Text_Redefine(fElecsys2010(4), "^")
                Case 3:    PatientPos = Text_Redefine(fElecsys2010(4), "^")
                Case Else: Exit Do
            End Select
            fElecsys2010(4) = Mid$(fElecsys2010(4), InStr(fElecsys2010(4), "^") + 1)   ' 구분자가 "^" 이다....
        Loop

        Patiant_Recevid = False        ' 환자번호 Flag
        sPatiant_No = fElecsys2010(3)  ' 환자번호
        '-------------------------------------------<<< 해당검사결과와 해당환자를 찿는다.       >>>----------
        With brspread
            For pDoCount = 1 To .maxrows
                .Row = pDoCount: .Col = 6
                If Trim$(.Text) = sPatiant_No Then
                    vRow = pDoCount
                    Patiant_Recevid = True
                    Exit For
                End If
            Next pDoCount
        End With

    ElseIf Mid$(fElecsys2010(1), 3, 1) = "R" Then
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
        If Patiant_Recevid = True Then
            fElecsys2010(3) = medGetP(fElecsys2010(3), 4, "^")

            Channel_No = fElecsys2010(3)
            With spdResult1
                For pDoCount = 10 To .MaxCols
                    .Row = vRow
                    .Col = pDoCount
                    .GetText 2, vRow, varTmp:    strDate = Trim$(varTmp)
                    .GetText 3, vRow, varTmp:    strSPnm = Trim$(varTmp)
                    .GetText 4, vRow, varTmp:    strSPid = Trim$(varTmp)
                    .GetText pDoCount, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Channel_No = itemX.tag Then
                        If Trim(fElecsys2010(4)) <> "" Then
                        
                            strResult = medGetP(Trim(fElecsys2010(4)), 1, "^")
                            strResult = Replace(strResult, "<", "")
                            strResult = Replace(strResult, ">", "")
                            
                            tmpChrResult = Result_Set(Channel_No, strResult) ' 장비채널/장비결과
                            
                            varResult = Split(tmpChrResult, "/")
                            
                            If Len(varResult(1)) > 0 Then
                                strRstval = Trim(varResult(1))
                            Else
                                strRstval = Trim(varResult(0))
                            End If
                                        
                            strRefVal = Trim(varResult(2))
                            
                            .Col = pDoCount:  .Row = vRow
                            
                            If strRefVal = "H" Then
                                .ForeColor = IIf(Trim$(strRefVal) = "H", vbRed, vbBlack)
                            ElseIf strRefVal = "L" Then
                                .ForeColor = IIf(Trim$(strRefVal) = "L", vbBlue, vbBlack)
                            Else
                                 .ForeColor = IIf(Trim$(strRefVal) = "", vbBlack, vbBlack)
                            End If
                            
                        Else
                            .Text = ""
                            strResult = ""
                        End If
                        
                        .Text = strRstval
                        .TypeHAlign = TypeVAlignCenter
                        .TypeVAlign = TypeVAlignCenter
                        
                        .Col = 1: .Text = 1
                        
                        .Col = 7: .BackColor = &HC0FFC0
                        .Col = 8: .BackColor = &HC0FFC0
                        .Col = 9: .BackColor = &HC0FFC0: .Text = "장비결과"
                    End If
                  '  .SetText 1, vRow, 0
                Next pDoCount
            End With
        End If
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "L" Then      ' "L" Data Last
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
        Patiant_Recevid = False                        ' 환자 번호  Flag
        
        If chkAuto.Value = "1" Then
            Call cmdAppend_Click(0)
        End If
    End If
                        
    Exit Sub
errDefine:

End Sub


''''' 2018.05.10 MAJESTIC
''''' 문자결과변화, 소수점처리, H/L 판정
Private Function Result_Set(examcode As String, Result As String) As String
    Dim strRefH As String
    Dim strRefM As String
    Dim strRefL As String
    Dim cRefH As String
    Dim cRefL As String
    Dim strResGubun As String
    Dim strLEquil As String
    Dim strHEquil As String
    Dim i As Integer
    Dim strRespRec As String
    Dim strPointFormat As String
    Dim cRepH As String
    Dim cRepL As String
    Dim strGiho As String
    Dim strResult As String
    Dim strResValue As String
    Dim strRefFlag As String
    Dim strDoc As String
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc As String
    Dim strRstDsp As String
    
    On Error GoTo ErrRes:
    
    Result_Set = ""
    strRefFlag = ""
    
    strResValue = Result
    
    If IsNumeric(strResValue) = False Then
        Result_Set = strResValue & "/" & strResValue & "/" & strRefFlag
        Exit Function
    End If
    
''    PANICH
''    PANICL
''    REFH
''    REFL
''    MREFH
''    MREFL
''    FREFL
''    FREFM
''    UNIT
''    EQP_NM
''    ResultLength
''    RESULT_TYPE
''    RESULT_LOW
''    RESULT_LOW_INT
''    RESULT_LOW_CHR
''    RESULT_HIGH
''    RESULT_HIGH_INT
''    RESULT_HIGH_CHR
''    RESULT_DSP
    
    sqlDoc = ""
    sqlDoc = sqlDoc & "SELECT * "
    sqlDoc = sqlDoc & "  FROM INTERFACE002 "
    sqlDoc = sqlDoc & "  WHERE EQP_CD = '" & INS_CODE & "' AND TESTCD_EQP = '" & examcode & "'"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet

    With adoRS
        If .RecordCount > 0 Then
            cRepL = Trim(.Fields("RESULT_LOW"))          '문자변환 LOW 참고치
            cRepH = Trim(.Fields("RESULT_HIGH"))         '문자변환 HIGH 참고치
            cRefL = Trim(.Fields("REFH"))                'NORMAL참고치상
            cRefH = Trim(.Fields("REFL"))                'NORMAL참고치하
            strRefL = Trim(.Fields("RESULT_LOW_CHR"))    'LOW 변환값
            strRefM = Trim(.Fields("RESULT_MID_CHR"))    'MID 변환값
            strRefH = Trim(.Fields("RESULT_HIGH_CHR"))   'HIGH 변환값
            strLEquil = Trim(.Fields("RESULT_LOW_INT"))  'LOW구분
            strHEquil = Trim(.Fields("RESULT_HIGH_INT")) 'HIGH구분
            strRespRec = Trim(.Fields("ResultLength"))   '소수점
            strResGubun = Trim(.Fields("RESULT_TYPE"))   '결과유헝
            strRstDsp = Trim(.Fields("RESULT_DSP"))
        End If
    End With
    
    adoRS.Close
    
    If IsNumeric(cRefL) = True Then
        If CCur(cRefL) > CCur(strResValue) Then
            strRefFlag = "L"
        End If
    End If
    
    If IsNumeric(cRefH) = True Then
        If CCur(cRefH) < CCur(strResValue) Then
            strRefFlag = "H"
        End If
    End If
    
    If strResGubun = "1" Then '문자
        If IsNumeric(cRepL) = True Then
            If strLEquil = "2" Then
                If CCur(cRepL) >= CCur(strResValue) Then
                    strResult = strRefL
                End If
            Else
                If CCur(cRepL) > CCur(strResValue) Then
                    strResult = strRefL
                End If
            End If
        End If
        
        If IsNumeric(cRepH) = True Then
            If strHEquil = "2" Then
                If CCur(cRepH) <= CCur(strResValue) Then
                    strResult = strRefH
                End If
            Else
                If CCur(cRefH) < CCur(strResValue) Then
                    strResult = strRefH
                End If
            End If
        End If
        
        ' 중간값
        If IsNumeric(cRepL) = True And IsNumeric(cRepH) = True Then
            If strRefM <> "" Then
                If strLEquil = "2" And strHEquil = "2" Then
                    If CCur(cRepL) <= CCur(strResValue) And CCur(cRepH) >= CCur(strResValue) Then
                        strResult = strRefM
                    End If
                ElseIf strLEquil = "2" And strHEquil = "1" Then
                    If CCur(cRepL) <= CCur(strResValue) And CCur(cRepH) > CCur(strResValue) Then
                        strResult = strRefM
                    End If
                ElseIf strLEquil = "1" And strHEquil = "2" Then
                    If CCur(cRepL) < CCur(strResValue) And CCur(cRepH) >= CCur(strResValue) Then
                        strResult = strRefM
                    End If
                Else
                    If CCur(cRepL) < CCur(strResValue) And CCur(cRepH) > CCur(strResValue) Then
                        strResult = strRefM
                    End If
                    
                End If
            Else
                strResult = strResult
            End If
        End If
    End If
       
    If IsNumeric(strRespRec) = True And strRespRec <> "9" Then
        strPointFormat = ""
        For i = 1 To CInt(strRespRec)
            strPointFormat = strPointFormat & "0"
        Next
        
        If strRespRec = "0" Then
            strPointFormat = "#######0" & strPointFormat
        Else
            strPointFormat = "##0." & strPointFormat
        End If
        strResValue = Format(strResValue, strPointFormat)

    Else
        strResValue = strResValue
    End If
    
    If Len(strRstDsp) > 0 Then
        Select Case strRstDsp
            Case 0: strResult = strResult
            Case 1: strResult = strResult & " " & strResValue
            Case 2: strResult = strResult & "(" & strResValue & ")"
            Case 3: strResult = strResValue
            Case 4: strResult = strResValue & " " & strResult
            Case 5: strResult = strResValue & "(" & strResult & ")"
        End Select
    End If
    
    strGiho = ""
    Result_Set = strGiho & strResValue & "/" & strResult & "/" & strRefFlag
    
    Set adoRS = Nothing
    
    Exit Function
    
ErrRes:
    
    Result_Set = strResValue & "/" & strResValue & "/" & strRefFlag
    Exit Function
    
End Function

Private Function Result_Chang(ByVal pTestcd As String, ByVal Result As String) As String
     Dim strResult As String
        Select Case UCase(pTestcd)
            Case "C4802", "C4711", "C4872"   'HBs Ag / HIV / HCV
                If Result <= 1 Then
                    strResult = "Negative"
                ElseIf Result > 1 Then
                    strResult = "Positive"
                End If
            Case "C4812" ' HBs Ab
                If Result <= 10 Then
                    strResult = "Negative"
                ElseIf Result > 10 Then
                    strResult = "Positive"
                End If
    
            Case Else
                strResult = ""
        End Select
        
        Result_Chang = strResult
    
End Function


Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Function f_funAdd_Server(ByVal strBarno As String, ByVal strTestCd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestCd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        strTmp2(1) = INS_CODE
        
        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "0" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
    End If
                                
End Function

Private Function f_funAdd_QcServer(ByVal strBarno As String, ByVal strTestCd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_QcServer = False
    
    strTmp = strTestCd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        strTmp2(1) = INS_CODE
        
        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_QcServer = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
    End If
                                
End Function

Public Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Public Function Text_Change(FSend_Str As String, FCheck_Char As String, FChange_Char As String) As String
Dim Pos_point As Integer
    Do
        Pos_point = InStr(FSend_Str, FCheck_Char)
        If Pos_point < 1 Then
            Exit Do
        ElseIf Pos_point = 1 Then
            FSend_Str = FChange_Char + Mid$(FSend_Str, 2)
        Else
            FSend_Str = Mid$(FSend_Str, 1, Pos_point - 1) + FChange_Char + Mid$(FSend_Str, Pos_point + 1)
        End If
    Loop
    Text_Change = FSend_Str
    
End Function

Private Sub Command1_Click()

   
    Dim Arr()   As Byte
    Dim strDta  As String

   
    strDta = ENQ & STX & "1H|\^&|||ACCESS" & vbCr & ETX
    strDta = strDta & "94" & vbCr & vbLf
    strDta = strDta & STX & "2Q|1|^060104T011||ALL||||||||O" & vbCr & ETX
    strDta = strDta & "32" & vbCr & vbLf
    strDta = strDta & STX & "3L|1|F" & vbCr & ETX
    strDta = strDta & "01" & vbCr & vbLf & EOT
    
    Call ComReceive(strDta)
    
End Sub
Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    cboKeyword.AddItem ""
    cboKeyword.AddItem "수검자명"
    cboKeyword.AddItem "병록번호"
    cboKeyword.AddItem "검체번호"
    cboKeyword.AddItem "결과값"
    
    cboKeyword.ListIndex = 0
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun           ' 실행
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now
    
    mskOrdDate.Text = Format$(Now - 1, "YYYYMMDD")
    mskOrdDate1.Text = Format$(Now, "YYYYMMDD")
    
    Open App.Path + "\" & INS_NAME & ".Log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(1).ListIndex = 2
    tabWork.Tab = 0
    Or_Seq = 1
    strFrameNo = 1
    SendCount = 0
    
    txtSeq.Text = "1"
    
'    If D0COM_CENTERCOD = "10" Then
'        cmdStartNo.Visible = False
'        cmdRackNo.Visible = True
'    Else
'        cmdStartNo.Visible = True
'        cmdRackNo.Visible = False
'    End If
    
End Sub

Private Sub f_subGet_Setting()
    
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subGet_Setting()"
    Set objComSetting = New clsCommon
    
    With objComSetting
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpProperty(INS_CODE)
    End With
    Set objComSetting = Nothing
    
    If mAdoRs Is Nothing Then
        IS_SET = False
        MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
            Set mAdoRs = Nothing
            Exit Sub
        Else
            IS_SET = True
            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")
            
            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
'                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
'                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
'                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
'                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
                .RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
                .SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
            End With
            Call Del_OldData
        End If
    End If
    
    Set mAdoRs = Nothing
Exit Sub

ErrRoutine:
    Set objComSetting = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call cmdStop
    Set Result = Nothing
    
    Close #1

End Sub

Private Sub fpSpread1_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub FrameError_Click()
    txtResult.Visible = True
    List1.Visible = False
End Sub

Private Sub imgPort_DblClick()
    
    If lvwCuData.Visible Then
        lvwCuData.Visible = False
    Else
        lvwCuData.Visible = True
        lvwCuData.ZOrder 0
    End If
    
End Sub

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "인터페이스 내용을 화면에 출력하지 않습니다."
    Else
        COM_MODE = "1"
        ShowMessage "인터페이스 내용을 화면에 출력합니다."
    End If
End Sub

Private Sub mskOrdDate_GotFocus()

    With mskOrdDate
        .SelStart = 8
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
    
End Sub

Private Sub Order_Ready(ByVal ACK As String)

    Static msgIndex As Long
    
    Select Case ACK
        Case Chr(COM_ENQ)
            msgIndex = 1
        Case Chr(COM_ACK)
            msgIndex = msgIndex + 1
        Case Chr(COM_NACK)
            msgIndex = msgIndex
        Case Chr(COM_EOT)
            msgIndex = 7
            Set Order = Nothing
        Case Else
        
    End Select
    
    Select Case msgIndex
        Case 1
            comEQP.Output = Order.MSG_ENQ
        Case 2
            comEQP.Output = Order.MSG_HEADER
        Case 3
            comEQP.Output = Order.MSG_PATIENT
        Case 4
            comEQP.Output = Order.MSG_ORDER
        Case 5
            comEQP.Output = Order.MSG_TERMINATION
        Case 6
            comEQP.Output = Order.MSG_EOT
        Case Else
    End Select
    
End Sub

Private Sub spdPtList_Advance(ByVal AdvanceNext As Boolean)
    Dim intCnt As Integer
    Dim varTmp
    Dim strRstDt As String
    Dim strSpcNo As String
    Dim tmpRS    As ADODB.Recordset
    Dim strSql   As String
    Dim strHL    As String
    
    With spdPtList
        .GetText 2, .ActiveRow, varTmp: strRstDt = Trim(varTmp)
        .GetText 4, .ActiveRow, varTmp: strSpcNo = Trim(varTmp)
        
        strSql = ""
        strSql = strSql & vbLf & " SELECT A.*, B.TESTNM, B.UNIT, B.REFL, B.REFH FROM INTERFACE003 A, INTERFACE002 B"
        strSql = strSql & vbLf & "  WHERE A.TRANSDT = '" & strRstDt & "' "
        strSql = strSql & vbLf & "    AND A.SPCNO = '" & strSpcNo & "' "
        strSql = strSql & vbLf & "    AND A.EqpNum = B.TESTCD_EQP "
        
        ' 정상
        If chkResult.Value = 1 Then
            strSql = strSql & "   And (A.JUDGE = '' OR A.JUDGE IS NULL) "
        End If
        ' H/L
        If chkNoResult.Value = 1 Then
            strSql = strSql & "   And A.JUDGE <> '' "
        End If
        ' PANIC
        If chkER.Value = 1 Then
            strSql = strSql & "   And A.PANIC <> '' "
        End If
        ' NEGATIVE
        If chkGonDan.Value = 1 Then
            strSql = strSql & "   And UCASE(A.RSTVAL) = 'NEGATIVE' "
        End If
        ' POSITIVE
        If chePositive.Value = 1 Then
            strSql = strSql & "   And UCASE(A.RSTVAL) = 'POSITIVE' "
        End If
        
        ' KEYWORD
'        If txtKeyword.text <> "" Then
'            sqlDoc = sqlDoc & "   And NAME LIKE '%" & txtKeyword.text & "%'  "
'        End If
        
        Set tmpRS = New ADODB.Recordset
        
        tmpRS.CursorLocation = adUseClient
        tmpRS.Open strSql, AdoCn_Jet
        
        If tmpRS.RecordCount > 0 Then
            spdResult2.maxrows = tmpRS.RecordCount
            tmpRS.MoveFirst
            For intCnt = 1 To tmpRS.RecordCount
                txtPtNm.Text = tmpRS.Fields("NAME") & ""
                txtSpcNo.Text = tmpRS.Fields("SPCNO") & ""
                txtPtId.Text = tmpRS.Fields("PNO") & ""
                txtResultDt.Text = Format(tmpRS.Fields("TRANSDT") & "", "####-##-##")
                txtSEX.Text = ""
                txtAGE.Text = ""
                
                With spdResult2
                    .SetText 1, intCnt, tmpRS.Fields("TESTNM") & ""
                    .SetText 2, intCnt, tmpRS.Fields("TESTCD") & ""
                    .SetText 3, intCnt, tmpRS.Fields("RSTVAL") & ""
                    Select Case tmpRS.Fields("JUDGE") & ""
                        Case "": strHL = "N": .Col = 6: .Row = intCnt: .Fontbold = False: .ForeColor = vbBlack
                        Case "L": strHL = "L": .Col = 6: .Row = intCnt: .Fontbold = True: .ForeColor = vbBlue
                        Case "H": strHL = "H": .Col = 6: .Row = intCnt: .Fontbold = True: .ForeColor = vbRed
                        Case Else: strHL = "N": .Col = 6: .Row = intCnt: .Fontbold = False: .ForeColor = vbBlack
                    End Select
                    .SetText 6, intCnt, strHL
                    .Col = 7: .Row = intCnt: .Fontbold = True: .ForeColor = vbRed
                    .SetText 7, intCnt, tmpRS.Fields("PANIC") & ""
                    .SetText 8, intCnt, "" 'tmpRS.Fields("DELTA") & ""
                    .SetText 9, intCnt, tmpRS.Fields("REFL") & "" & " ~ " & tmpRS.Fields("REFH") & ""
                    .SetText 10, intCnt, tmpRS.Fields("UNIT") & ""
                    .SetText 11, intCnt, tmpRS.Fields("EQUIPCD") & ""
                End With
                tmpRS.MoveNext
            Next
        End If
    End With
    
    Set tmpRS = Nothing
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    Dim varTmp  As Variant
    Dim pResult   As String
    Dim pTestName As String
    Dim mTestName As Boolean
    
    If Row = 0 Then Exit Sub
    
    With spdRstview
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    mTestName = False
    
    With spdResult1
        For iCnt = gItemStartPos To .MaxCols
            .Row = Row: .Col = iCnt

            If .BackColor = &HC6FEFF Or .Text <> "" Then
                .GetText .Col, .Row, varTmp:         pResult = Trim$(varTmp)
                .GetText .Col, 0, varTmp:           pTestName = Trim$(varTmp)
                
                If spdRstview.maxrows = 1 And mTestName = False Then
                    mTestName = True
                    spdRstview.SetText 1, spdRstview.maxrows, pTestName
                    spdRstview.SetText 2, spdRstview.maxrows, pResult
                Else
                    spdRstview.maxrows = spdRstview.maxrows + 1
                    spdRstview.SetText 1, spdRstview.maxrows, pTestName
                    spdRstview.SetText 2, spdRstview.maxrows, pResult
                End If
            End If
        Next
    End With
End Sub


Private Sub spdResult1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long

    Set oMenu = New cPopupMenu

    lMenuChosen = oMenu.Popup(" ▒ 검사자 삭제")

    Select Case lMenuChosen

        Case 1
            With spdResult1
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End With
    End Select
End Sub

Private Sub SpreadSheetSort(ByRef Spread As vaSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
    Dim intCount As Integer
    Dim strDataField As String
    'SortType
    ' 0 : none
    ' 1 : ascending
    ' 2 : descending

    With Spread
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = Col       '정렬키 열번호

        If SortType = 0 Then
            .SortKeyOrder(1) = SortKeyOrderNone
        ElseIf SortType = 1 Then
            .SortKeyOrder(1) = SortKeyOrderAscending
        ElseIf SortType = 2 Then
            .SortKeyOrder(1) = SortKeyOrderDescending
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending
        End If

        .Action = ActionSort
    End With

End Sub

Private Sub spdWorklist_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim varTmp  As Variant
    

        With spdWorkList
            .GetText 2, Row, varTmp
            
            If Trim$(varTmp) = "" Then Exit Sub
    
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
            cmdWorkList_Click
        End With

    
End Sub

'Private Sub spdWorklist_DblClick(ByVal Col As Long, ByVal Row As Long)
'
'    Dim varTmp  As Variant
'
'    If Row = 0 Then
'        If Col = 1 Then
'            Col = 2
'        End If
'
'        If OrderSort_Flag = 1 Then
'            Call SpreadSheetSort(spdWorkList, Col, 2)
'            OrderSort_Flag = 2
'        Else
'            Call SpreadSheetSort(spdWorkList, Col, 1)
'            OrderSort_Flag = 1
'        End If
'    Else
'        With spdWorkList
'            .GetText 2, Row, varTmp
'            If Trim$(varTmp) = "" Then Exit Sub
'
'            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
'            Call cmdWorkList_Click
'        End With
'    End If
'
'End Sub

Private Sub Timer1_Timer()

    comEQP.Output = ENQ

End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub


Private Sub txtResult_DblClick()
    txtResult.Text = ""
    List1.Text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    List1.Visible = True
End Sub
