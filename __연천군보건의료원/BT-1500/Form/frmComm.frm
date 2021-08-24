VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15375
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '최대화
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8880
      Top             =   5160
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   6480
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7860
      Top             =   5130
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8340
      Top             =   5130
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   7230
      Top             =   5130
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
      Begin VB.Timer tmrWorking 
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1440
         Top             =   60
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   6615
         TabIndex        =   25
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
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   28
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
      Begin VB.Image imgBack 
         BorderStyle     =   1  '단일 고정
         Height          =   1050
         Index           =   0
         Left            =   4710
         Picture         =   "frmComm.frx":5794
         Stretch         =   -1  'True
         Top             =   -210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Image imgLogo 
         Height          =   240
         Index           =   0
         Left            =   4320
         Picture         =   "frmComm.frx":6F67
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "팝업용 ==>"
         Height          =   225
         Index           =   1
         Left            =   3390
         TabIndex        =   75
         Top             =   210
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
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
         Left            =   960
         TabIndex        =   6
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Left            =   210
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
      Width           =   15375
      _ExtentX        =   27120
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
         Left            =   14145
         TabIndex        =   4
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   13110
         TabIndex        =   3
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Index           =   0
         Left            =   12015
         TabIndex        =   2
         Top             =   195
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15015
         Picture         =   "frmComm.frx":74F1
         Top             =   165
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13725
         Picture         =   "frmComm.frx":7A7B
         Top             =   165
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12525
         Picture         =   "frmComm.frx":8005
         Top             =   165
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8370
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   14764
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      TabCaption(0)   =   " ▒    WorkList     "
      TabPicture(0)   =   "frmComm.frx":858F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdReceve"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPosNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOrder"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "spdRstview"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSearch"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAppend(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "SSPanel2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "SSPanel(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "spdResult1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "spdWorklist"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtBarCode"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "pnlCom2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdRequist(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdPrint"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkAuto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtResult"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdRackNo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdStartNo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdWordQuery"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdEot"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdWorkList"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "List1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdNext"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdPrevious"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtSeqNo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "pnlCom"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtSEQ"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cboChk"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   " ▒   받은 결과     "
      TabPicture(1)   =   "frmComm.frx":85AB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkExcel"
      Tab(1).Control(1)=   "cmdExcel"
      Tab(1).Control(2)=   "spdResult2"
      Tab(1).Control(3)=   "cmdRstQuery"
      Tab(1).Control(4)=   "lvwCuData"
      Tab(1).Control(5)=   "cmdAppend(1)"
      Tab(1).Control(6)=   "CommonDialog1"
      Tab(1).Control(7)=   "tblexcel"
      Tab(1).Control(8)=   "SSPanel(0)"
      Tab(1).Control(9)=   "cmdSel(3)"
      Tab(1).Control(10)=   "cmdSel(2)"
      Tab(1).ControlCount=   11
      Begin VB.ComboBox cboChk 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmComm.frx":85C7
         Left            =   5280
         List            =   "frmComm.frx":85D1
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtSEQ 
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
         Height          =   375
         Left            =   7000
         TabIndex        =   87
         Text            =   "123"
         Top             =   460
         Width           =   975
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   4725
         Left            =   1890
         TabIndex        =   32
         Top             =   6540
         Visible         =   0   'False
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   8334
         Bevel           =   1
         Moveble         =   -1  'True
         CloseEnabled    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   34
            Top             =   4020
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Save"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMOutput 
               Height          =   360
               Left            =   1155
               TabIndex        =   36
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear 
               Height          =   360
               Left            =   9450
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput 
               Height          =   360
               Left            =   90
               TabIndex        =   38
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
         Begin VB.TextBox txtCom 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3720
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   33
            Top             =   315
            Width           =   11595
         End
      End
      Begin VB.TextBox txtSeqNo 
         Alignment       =   2  '가운데 맞춤
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
         Height          =   300
         Left            =   12750
         MaxLength       =   12
         TabIndex        =   76
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   750
      End
      Begin BHButton.BHImageButton cmdPrevious 
         Height          =   330
         Left            =   90
         TabIndex        =   56
         Top             =   5400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         Caption         =   "◀"
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
         ForeColor       =   16711680
         BackColor       =   16711680
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdNext 
         Height          =   330
         Left            =   330
         TabIndex        =   57
         Top             =   5400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         Caption         =   "▶"
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
         TransparentPicture=   "frmComm.frx":85E1
         ForeColor       =   16711680
         BackColor       =   255
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin VB.Frame Frame3 
         Height          =   315
         Left            =   90
         TabIndex        =   51
         Top             =   900
         Width           =   555
         Begin Threed.SSCommand cmdSel 
            Height          =   345
            Index           =   1
            Left            =   270
            TabIndex        =   53
            Top             =   0
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   609
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":8A53
         End
         Begin Threed.SSCommand cmdSel 
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   609
            _StockProps     =   78
            ForeColor       =   14735310
            BevelWidth      =   1
            Picture         =   "frmComm.frx":8ED5
         End
      End
      Begin VB.CheckBox chkExcel 
         Appearance      =   0  '평면
         Caption         =   "Excel 생성"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -60960
         TabIndex        =   50
         Top             =   30
         Value           =   1  '확인
         Width           =   1155
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   420
         Left            =   -68910
         TabIndex        =   49
         Top             =   420
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   741
         Caption         =   "Excel 파일 생성 / 출력"
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
      Begin VB.ListBox List1 
         Height          =   2220
         ItemData        =   "frmComm.frx":9343
         Left            =   7950
         List            =   "frmComm.frx":9345
         TabIndex        =   44
         Top             =   6060
         Width           =   7215
      End
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   435
         Left            =   90
         TabIndex        =   29
         Top             =   4890
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   767
         Caption         =   "WorkList 등록"
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
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74910
         TabIndex        =   40
         Top             =   900
         Width           =   15075
         _Version        =   393216
         _ExtentX        =   26591
         _ExtentY        =   12965
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   1
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":9347
         UserResize      =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   375
         Left            =   6300
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   420
         Left            =   -70260
         TabIndex        =   31
         Top             =   420
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
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
         TabIndex        =   20
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
         Left            =   -62355
         TabIndex        =   30
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
      Begin BHButton.BHImageButton cmdEot 
         Height          =   375
         Left            =   12420
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "초기화"
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
      Begin BHButton.BHImageButton cmdWordQuery 
         Height          =   390
         Left            =   9330
         TabIndex        =   41
         Top             =   5400
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   688
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
      Begin BHButton.BHImageButton cmdStartNo 
         Height          =   420
         Left            =   8160
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
         Caption         =   "시작번호변경"
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
         Left            =   11130
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
         Height          =   2040
         Left            =   4920
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   10290
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  '평면
         Caption         =   "Auto Server"
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
         Height          =   225
         Left            =   13650
         TabIndex        =   24
         Top             =   60
         Visible         =   0   'False
         Width           =   1470
      End
      Begin BHButton.BHImageButton cmdPrint 
         Height          =   420
         Left            =   9600
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
         Caption         =   "WorkSheet Print"
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
      Begin BHButton.BHImageButton cmdRequist 
         Height          =   390
         Index           =   2
         Left            =   7950
         TabIndex        =   48
         Top             =   5400
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   688
         Caption         =   "Last Order.."
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -65490
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   5385
         Left            =   8460
         TabIndex        =   10
         Top             =   8520
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   9499
         Bevel           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtCOM2 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4395
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   19
            Top             =   300
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   11
            Top             =   4635
            Width           =   5760
            Begin MSComDlg.CommonDialog cdlFile 
               Left            =   5265
               Top             =   60
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin HSCotrol.CButton cmdChksum 
               Height          =   360
               Left            =   2205
               TabIndex        =   12
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "SUM"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMOutput2 
               Height          =   360
               Left            =   1155
               TabIndex        =   13
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear2 
               Height          =   360
               Left            =   3600
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput2 
               Height          =   360
               Left            =   90
               TabIndex        =   15
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMLoad 
               Height          =   360
               Left            =   4635
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Load"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdACK 
               Height          =   360
               Left            =   3105
               TabIndex        =   17
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ACK"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdENQ 
               Height          =   360
               Left            =   2655
               TabIndex        =   18
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ENQ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   1500
      End
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   3960
         Left            =   90
         TabIndex        =   54
         Top             =   900
         Width           =   4755
         _Version        =   393216
         _ExtentX        =   8387
         _ExtentY        =   6985
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   20
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":98A1
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4425
         Left            =   4890
         TabIndex        =   55
         Top             =   900
         Width           =   10275
         _Version        =   393216
         _ExtentX        =   18124
         _ExtentY        =   7805
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   7
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
         FormulaSync     =   0   'False
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   22
         MaxRows         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735309
         SpreadDesigner  =   "frmComm.frx":9FA7
         UserResize      =   0
         TextTip         =   1
         TextTipDelay    =   1
         CellNoteIndicator=   3
      End
      Begin FPSpread.vaSpread tblexcel 
         Height          =   675
         Left            =   -66240
         TabIndex        =   58
         Top             =   270
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
         SpreadDesigner  =   "frmComm.frx":A7B8
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   465
         Index           =   0
         Left            =   -74910
         TabIndex        =   59
         Top             =   390
         Width           =   4545
         _Version        =   65536
         _ExtentX        =   8017
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
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
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   1
            ItemData        =   "frmComm.frx":A9DE
            Left            =   2640
            List            =   "frmComm.frx":A9E5
            Style           =   2  '드롭다운 목록
            TabIndex        =   64
            Top             =   105
            Width           =   1770
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "frmComm.frx":A9F6
            Left            =   4590
            List            =   "frmComm.frx":A9F8
            TabIndex        =   61
            Top             =   480
            Visible         =   0   'False
            Width           =   1725
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4560
            TabIndex        =   60
            Top             =   450
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtpRsltDay 
            Height          =   315
            Left            =   1290
            TabIndex        =   63
            Top             =   90
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   130220033
            CurrentDate     =   40248
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "검사결과일 :"
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
            Left            =   90
            TabIndex        =   65
            Top             =   150
            Width           =   1125
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   62
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   465
         Index           =   1
         Left            =   90
         TabIndex        =   66
         Top             =   390
         Width           =   5235
         _Version        =   65536
         _ExtentX        =   9234
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
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
         Begin VB.TextBox txtChart 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   3780
            MaxLength       =   12
            TabIndex        =   68
            Top             =   90
            Width           =   1395
         End
         Begin VB.ComboBox cboComNm 
            Height          =   300
            ItemData        =   "frmComm.frx":A9FA
            Left            =   4590
            List            =   "frmComm.frx":A9FC
            TabIndex        =   67
            Top             =   480
            Visible         =   0   'False
            Width           =   1725
         End
         Begin MSMask.MaskEdBox mskOrdtime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Left            =   4560
            TabIndex        =   69
            Top             =   450
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2490
            TabIndex        =   70
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   130220033
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   1080
            TabIndex        =   71
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   130220033
            CurrentDate     =   40248
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   74
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2370
            TabIndex        =   73
            Top             =   150
            Width           =   315
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   90
            TabIndex        =   72
            Top             =   150
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   12090
         TabIndex        =   78
         Top             =   5400
         Visible         =   0   'False
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   767
         _StockProps     =   15
         BackColor       =   14737632
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
         Begin VB.OptionButton optSeq 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "POS"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1260
            TabIndex        =   80
            Top             =   90
            Width           =   645
         End
         Begin VB.OptionButton optBar 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "등록번호"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1950
            TabIndex        =   79
            Top             =   90
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "▒ Search :"
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
            Height          =   255
            Left            =   60
            TabIndex        =   81
            Top             =   120
            Width           =   1485
         End
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   21
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":A9FE
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   22
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":AE80
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   420
         Index           =   0
         Left            =   13620
         TabIndex        =   82
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
         Caption         =   "파일만들기"
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
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   420
         Left            =   5400
         TabIndex        =   83
         Top             =   435
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
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
      Begin FPSpread.vaSpread spdRstview 
         Height          =   2865
         Left            =   90
         TabIndex        =   84
         Top             =   5400
         Width           =   7815
         _Version        =   393216
         _ExtentX        =   13785
         _ExtentY        =   5054
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   6
         MaxRows         =   8
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":B2EE
         UserResize      =   0
      End
      Begin BHButton.BHImageButton cmdOrder 
         Height          =   420
         Left            =   9000
         TabIndex        =   85
         Top             =   435
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   741
         Caption         =   "오더전송"
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
      Begin BHButton.BHImageButton cmdPosNo 
         Height          =   420
         Left            =   8040
         TabIndex        =   86
         Top             =   435
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   741
         Caption         =   "Seq변경"
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
      Begin BHButton.BHImageButton cmdReceve 
         Height          =   420
         Left            =   10080
         TabIndex        =   89
         Top             =   435
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   741
         Caption         =   "결과받기"
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
      Begin VB.Label Label12 
         Caption         =   "▒ Start POS :"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   11490
         TabIndex        =   77
         Top             =   540
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         DrawMode        =   5  '카피 펜이 아님
         X1              =   9480
         X2              =   15180
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Label Label8 
         Caption         =   "● Information List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   7950
         TabIndex        =   46
         Top             =   5790
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "재검/QC :"
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
         Left            =   8040
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag_HQL As String

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
Private tDrive         As String


Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const fs  As String = ""
Const RS  As String = ""
Const ASUB As String = ""



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
Dim sLfCheck  As Integer
Dim sCrcheck  As Integer
' --------------------------------------------------------------
Dim strOrdLst As String

Dim ELEC1010(100)   As String
Dim fELEC1010       As Variant
Dim fELEC1010_1     As Variant
Dim fELEC1010_2     As Variant
Dim fELEC1010_3     As Variant
Dim SendData(10)     As String
Dim SendCount        As String
Dim Or_Seq           As Integer
Dim SendBuffW           As String
Dim SendBuffT           As String
Dim intRow          As Integer
Dim brStr           As String

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


Dim fSens(100)   As String
Dim fCellDynSize(50, 1) As Integer
Dim fChannel() As String
Dim PName   As String
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
    strTestcd(50) As String
End Type

Private f_typCode() As TYPE_CD

Dim RecordChk As Boolean

Dim strGumCd As String
Dim strJinCd As String
Dim fRcvString As String

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String

Dim SeqNo As String
'Dim RecordChk   As Boolean

Dim G_CLVALU    As String
Dim G_CHVALU    As String
Dim G_EVALUATE  As String
Dim G_PANIC     As String
Dim G_DELTA     As String
Dim strFrameNo  As Integer
Dim OrderCnt As Integer
Dim vRow As Integer
Dim sPatiant_No As String

Private Type typeALFA
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(50)   As String
    Result(50)   As String
    Status(50)   As String
    Rerun(50)    As String
End Type

Dim ALFA As typeALFA


Private Type typeXMLData
    Company     As String
    HospCode    As String
    ChartNo     As String
    patname     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLData As typeXMLData


Dim OrderSort_Flag As Integer
Dim gspdResultRow  As Integer
Dim chrCount       As Integer

Dim gRecodeType As String



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

Private Function f_funGet_SpreadRow(ByVal objSpd As vaSpread, ByVal intCol As Integer, _
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
    With spdWorklist
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
        
        With spdWorklist
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Function f_subSet_XMLWorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim strPath   As String
    Dim tstrBuffer As String
    Dim strBuffer As String
    Dim I         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_XMLWorkList() As ADODB.Recordset"
    
    Screen.MousePointer = 11
    
    '-- 오더파일명과 경로를 지정한다.
    
   
    strPath = gIn_Path & "\ExamIF_In.xml"

    Open strPath For Input As #300

    strBuffer = ""
    Do While Not EOF(300)
        Line Input #300, tstrBuffer
        strBuffer = strBuffer & tstrBuffer
    Loop

    Close #300
    
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For I = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, I, 4)
        Else
            BufChar = Mid$(strBuffer, I + 3)
        End If
        
        If BufChar = "<검사>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    
    Next

'    f_subSet_XMLWorkList = Split(strTmp, "</검사>")
    strTmp = Replace(strTmp, "<검사>", ""): strTmp = Replace(strTmp, "</검사>", "|")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<요양기관번호>", ""): strTmp = Replace(strTmp, "</요양기관번호>", ",")
    strTmp = Replace(strTmp, "<차트번호>", ""): strTmp = Replace(strTmp, "</차트번호>", ",")
    strTmp = Replace(strTmp, "<수진자명>", ""): strTmp = Replace(strTmp, "</수진자명>", ",")
    strTmp = Replace(strTmp, "<주민등록번호>", ""): strTmp = Replace(strTmp, "</주민등록번호>", ",")
    strTmp = Replace(strTmp, "<내원번호>", ""): strTmp = Replace(strTmp, "</내원번호>", ",")
    strTmp = Replace(strTmp, "<의뢰일>", ""): strTmp = Replace(strTmp, "</의뢰일>", ",")
    strTmp = Replace(strTmp, "<검사번호>", ""): strTmp = Replace(strTmp, "</검사번호>", ",")
    strTmp = Replace(strTmp, "<검사ID>", ""): strTmp = Replace(strTmp, "</검사ID>", ",")
    strTmp = Replace(strTmp, "<업체검사ID>", ""): strTmp = Replace(strTmp, "</업체검사ID>", ",")
    strTmp = Replace(strTmp, "<검체>", ""): strTmp = Replace(strTmp, "</검체>", ",")
    strTmp = Replace(strTmp, "<결과치>", ""): strTmp = Replace(strTmp, "</결과치>", ",")
    strTmp = Replace(strTmp, "<참조치>", ""): strTmp = Replace(strTmp, "</참조치>", ",")
    strTmp = Replace(strTmp, "<소견>", ""): strTmp = Replace(strTmp, "</소견>", ",")
    strTmp = Replace(strTmp, "<결과일>", ""): strTmp = Replace(strTmp, "</결과일>", ",")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<입원외래구분>", ""): strTmp = Replace(strTmp, "</입원외래구분>", ",")
    
    f_subSet_XMLWorkList = Split(strTmp, "|")
    
    Screen.MousePointer = 0

    Exit Function
        
ErrorTrap:
    
    Screen.MousePointer = 0
    
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_WorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
        Set AdoRs_SQL = New ADODB.Recordset
        
        If cboChk.ListIndex = 0 Then
            sqlDoc = ""
            sqlDoc = "         SELECT A.REQUEST_DATE, A.EXAM_NO, A.PERSON_NAME, A.PERSONAL_ID, B.EXAM_CODE, B.REFER_VALUE, B.RESULT_STYLE"
            sqlDoc = sqlDoc + "  FROM MEDITOLISS..TOTAL A, MEDITOLISS..TOTRES B"
            sqlDoc = sqlDoc + " WHERE  A.REQUEST_DATE between '" & strDate & "' and '" & strDate1 & "'"
            sqlDoc = sqlDoc + "   AND B.EXAM_PART = 'H'"
            sqlDoc = sqlDoc + "   AND B.RESULT_VALUE = ''"
            sqlDoc = sqlDoc + "   AND A.REQUEST_DATE = B.REQUEST_DATE"
            sqlDoc = sqlDoc + "   AND A.EXAM_NO = B.EXAM_NO"
        Else
             sqlDoc = "         Select a.*, b.수진자명,b.챠트번호,b.주민등록번호,  b.주민등록번호 as 성별 from TB_검사항목 a, TB_인적사항 b"
            sqlDoc = sqlDoc & " Where a.진료년+a.진료월+a.진료일 between '" & strDate & "' and '" & strDate1 & "'"
            sqlDoc = sqlDoc & "   And a.진료지원상태 < 5"
            sqlDoc = sqlDoc & "   And a.진료지원상태 <> 5"
'            sqlDoc = sqlDoc & "   and a.처방코드 + a.서브코드 in('B1050','B1040','B1010','B1020','B1060001','B1060002','B1060003','B1060','B1091001','B1091002','B1091003','B1091','B1020001','B1020002','B1020003','ZZ015','ZZ016','ZZ017','ZZ052','ZZ053','ZZ054') "
'            sqlDoc = sqlDoc & "   and 서브코드 in('','001','002','003','004','005','006','007','008','009','010') "
            sqlDoc = sqlDoc & "   and a.처방번호 >= 0"
            sqlDoc = sqlDoc & "   And a.챠트번호 = b.챠트번호"
            sqlDoc = sqlDoc & " Order By a.챠트번호"

        End If
        
        Set AdoRs_SQL = New ADODB.Recordset
        
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_WorkList = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
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

Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String, Optional ByVal strStatus As String, Optional ByVal strName As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
        If Mid(strName, 1, 2) = "검진" Then
            strDate = strBarno
            '-- 검진
            sqlDoc = ""
            sqlDoc = "         SELECT EXAM_CODE, REFER_VALUE, RESULT_STYLE"
            sqlDoc = sqlDoc + "  FROM MEDITOLISS..TOTRES"
            sqlDoc = sqlDoc + " WHERE REQUEST_DATE = '" & strDate & "'"
            sqlDoc = sqlDoc + "   AND EXAM_NO = '" & strStatus & "'"
            sqlDoc = sqlDoc + "   AND EXAM_PART = 'H'"
            sqlDoc = sqlDoc + "   AND RESULT_VALUE = ''"
        Else
        
            stryy = Mid(strBarno, 1, 4)
            strmm = Mid(strBarno, 5, 2)
            strdd = Mid(strBarno, 7, 2)
            
            sqlDoc = ""
            sqlDoc = sqlDoc & "Select 처방코드 + 서브코드 as 처방코드 " & vbCrLf
            sqlDoc = sqlDoc & "  From TB_검사항목 " & vbCrLf
            sqlDoc = sqlDoc & " Where 진료년 = '" & stryy & "'" & vbCrLf
            sqlDoc = sqlDoc & "   and 진료월 = '" & strmm & "'" & vbCrLf
            sqlDoc = sqlDoc & "   and 진료일 = '" & strdd & "'" & vbCrLf
            sqlDoc = sqlDoc & "   and 챠트번호 = '" & strStatus & "'" & vbCrLf
'            sqlDoc = sqlDoc & "   and 처방코드 in ('B1050','B1040','B1010','B1020','B1060001','B1060002','B1060003','B1060','B1091001','B1091002','B1091003','B1091','B1020001','B1020002','B1020003','ZZ015','ZZ016','ZZ017','ZZ052','ZZ053','ZZ054') " & vbCrLf
'            sqlDoc = sqlDoc & "   and 서브코드 in ('','001','002','003','004','005','006','007','007','009','010') " & vbCrLf
            sqlDoc = sqlDoc & "   and 처방번호 >= 0" & vbCrLf
            sqlDoc = sqlDoc & "   and 진료지원상태 < 5 " & vbCrLf


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
    
    Dim intPos1 As Integer
    
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 10
    intCol2 = 1
    intRow = 1
    With spdWorklist
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
'             " order by TESTCD_EQP, TESTCD"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        ReDim fChannel(adoRS.RecordCount)
        strJinCd = ""
        strGumCd = ""
    End If
    Do While Not adoRS.EOF
        If Trim(adoRS.Fields("TESTCD")) <> "" Then
            intPos1 = InStr(Trim(adoRS.Fields("TESTCD")), ",")
            If intPos1 = 0 Then
                strGumCd = strGumCd & "'" & Trim(adoRS.Fields("TESTCD")) & "',"
            Else
                strGumCd = strGumCd & "'" & Mid(Trim(adoRS.Fields("TESTCD")), 1, intPos1 - 1) & "',"
                strJinCd = strJinCd & "" & Mid(Trim(adoRS.Fields("TESTCD")), intPos1 + 1) & ","
            End If
        End If
        
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
        
        With spdWorklist
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
        
        With spdRstview
            If intRow > .maxrows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(adoRS("TESTNM") & "")
            intRow = intRow + 1
            
        End With
        
        With spdResult2
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
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
            f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"
        f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
        f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = strTmp
        
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    If Trim(strGumCd) <> "" Then strGumCd = Mid(strGumCd, 1, Len(strGumCd) - 1)
    If Trim(strJinCd) <> "" Then strJinCd = Mid(strJinCd, 1, Len(strJinCd) - 1)
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With

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
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

Private Function f_subSet_ComList()
    
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap

    CallForm = "clsCommon - Public Function f_subSet_ComList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
        
        sqlDoc = "         SELECT B.COM_CODE, B.COM_NAME " & vbCr
        sqlDoc = sqlDoc & "  FROM MDCK..GUMJIN_INTERFACE A, MDCK..TB_COMPANY B, MDCK..BAG_INTERFACECODE C " & vbCr
        sqlDoc = sqlDoc & " WHERE A.Per_com_Code = B.COM_CODE " & vbCr
        sqlDoc = sqlDoc & "   AND A.per_gumjin_date BETWEEN '" & Format(dtpStartDt, "yyyymmdd") & "' AND '" & Format(dtpStopDt, "yyyymmdd") & "'" & vbCr
        sqlDoc = sqlDoc & "  AND SUBSTRING(C.KIND, 1, 1) = 'C' " & vbCr
        sqlDoc = sqlDoc & "   AND A.EDPSCODE = C.MEDITEM " & vbCr
        sqlDoc = sqlDoc & " GROUP BY B.COM_CODE, B.COM_NAME " & vbCr
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount > 0 Then
            AdoRs_SQL.MoveFirst
            cboComNm.Clear
            cboComNm.AddItem "전체"
            Do Until AdoRs_SQL.EOF
                cboComNm.AddItem AdoRs_SQL.Fields("COM_NAME") & ""
                AdoRs_SQL.MoveNext
            Loop
            cboComNm.ListIndex = 0
        End If
        
        AdoRs_SQL.Close:  Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub cboChk_Click()
    If Trim(cboChk.Text) = "검진" Then
'        cboComNm.Visible = True
'        mskOrdtime.Visible = False
'        Label10.Visible = False
'        Call f_subSet_ComList
    Else
        cboComNm.Visible = False
'        mskOrdtime.Visible = True
'        Label10.Visible = True
        cboComNm.Clear
    End If
End Sub

Private Sub cboComNm_DropDown()
        Call f_subSet_ComList
End Sub


Private Function MakeXMLResult(ByVal m업체 As String, ByVal m요양기관번호 As String, ByVal m차트번호 As String, ByVal m수진자명 As String, ByVal m주민등록번호 As String, ByVal m내원번호 As String, ByVal m의뢰일 As String, ByVal m검사번호 As String, ByVal m검사ID As String, ByVal m업체검사ID As String, ByVal m검체 As String, ByVal m결과치 As String, ByVal m참조치 As String, ByVal m소견 As String, ByVal m결과일 As String, ByVal m입원외래구분 As String, ByVal chkYN As Boolean)

    Dim objXML, objXMLv
    Dim bInFileExist, intID, I
    Dim strPath     As String
    Dim iCnt        As Integer
    Dim strSql      As String
    Dim strFile     As String
    Dim strTerm     As String
    Dim strAppName  As String
    Dim strDBName   As String
    
    Dim rsAuthors As ADODB.Recordset
    Dim Conn As ADODB.Connection
    Dim strIP           As String
    Dim strPORT         As String
    Dim strXSLPath      As String
    Dim strTemp         As String
    Dim strBuffer       As String
    Dim tstrBuffer      As String
    
    strPath = gIn_Path & "\ExamIF_Out.xml"
    strXSLPath = "C:\UBCare\SINAI\IF\Form\"
    
    strFile = strPath

    Set objXML = CreateObject("Microsoft.XMLDom")

    objXML.async = False

    bInFileExist = objXML.Load(strFile)

    If bInFileExist = True Then
        Open strPath For Input As #50
    
        strBuffer = ""
        Do While Not EOF(50)
            Line Input #50, tstrBuffer
            strBuffer = strBuffer & tstrBuffer
        Loop
    
        Close #50
    End If
    
    If InStr(strBuffer, "<?xml version=") = 0 Then
        Call objXML.appendChild(objXML.createProcessingInstruction("xml", "version=""1.0"" encoding=""euc-kr"" "))
        Call objXML.appendChild(objXML.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=" & strXSLPath & "ExamIF_Form_05.xsl"""))
        Call objXML.appendChild(objXML.createElement("UBCare검사정보"))
    End If
    
    Set objXMLv = objXML.createElement("검사")
    
    Call objXMLv.appendChild(objXML.createElement("업체"))
    Call objXMLv.appendChild(objXML.createElement("요양기관번호"))
    Call objXMLv.appendChild(objXML.createElement("차트번호"))
    Call objXMLv.appendChild(objXML.createElement("수진자명"))
    Call objXMLv.appendChild(objXML.createElement("주민등록번호"))
    Call objXMLv.appendChild(objXML.createElement("내원번호"))
    Call objXMLv.appendChild(objXML.createElement("의뢰일"))
    Call objXMLv.appendChild(objXML.createElement("검사번호"))
    Call objXMLv.appendChild(objXML.createElement("검사ID"))
    Call objXMLv.appendChild(objXML.createElement("업체검사ID"))
    Call objXMLv.appendChild(objXML.createElement("검체"))
    Call objXMLv.appendChild(objXML.createElement("결과치"))
    Call objXMLv.appendChild(objXML.createElement("참조치"))
    Call objXMLv.appendChild(objXML.createElement("소견"))
    Call objXMLv.appendChild(objXML.createElement("결과일"))
    Call objXMLv.appendChild(objXML.createElement("입원외래구분"))
    
    objXMLv.childNodes(0).Text = m업체
    objXMLv.childNodes(1).Text = m요양기관번호
    objXMLv.childNodes(2).Text = m차트번호
    objXMLv.childNodes(3).Text = m수진자명
    objXMLv.childNodes(4).Text = m주민등록번호
    objXMLv.childNodes(5).Text = m내원번호
    objXMLv.childNodes(6).Text = m의뢰일
    objXMLv.childNodes(7).Text = m검사번호
    objXMLv.childNodes(8).Text = m검사ID
    objXMLv.childNodes(9).Text = m업체검사ID
    objXMLv.childNodes(10).Text = m검체
    objXMLv.childNodes(11).Text = m결과치
    objXMLv.childNodes(12).Text = m참조치
    objXMLv.childNodes(13).Text = m소견
    objXMLv.childNodes(14).Text = m결과일
    objXMLv.childNodes(15).Text = m입원외래구분
    
    objXML.documentElement.appendChild (objXMLv.CloneNode(True))
    
    objXML.Save (strPath)
    
    Set objXML = Nothing
End Function

Private Sub cmdAppend_Click(Index As Integer)
   
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    Dim strTestnm       As String
    Dim strREF          As String
    Dim strUnit         As String
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String

    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    Dim SaveGbn As Integer
    
    Dim strDate As String
    Dim strBarno As String
    Dim strSPnm As String
    Dim strSPid As String
    Dim strChartNo As String
    Dim strEqpCd As String
    Dim valEqpcd As Variant
    Dim strSeq As String
    
    Dim varXMLTmp   As Variant
    Dim varSEQNO    As Variant
    Dim tmpXML      As Boolean
    
    Dim t업체, t요양기관번호, t차트번호, t수진자명, t주민등록번호, t내원번호, t의뢰일, t검사번호, t검사ID, t업체검사ID, t검체, t결과치, t참조치, t소견, t결과일, t입원외래구분 As String
    
    Dim m검사번호 As String
    Dim m검사코드 As String
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"

On Error GoTo ErrorRoutine

    Me.MousePointer = 11

    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If

    tmpXML = False

    With objSpd
        For intRow = 1 To .maxrows
        
            t업체 = "":                t요양기관번호 = ""
            t차트번호 = "":            t수진자명 = ""
            t주민등록번호 = "":        t내원번호 = ""
            t의뢰일 = "":              t검사번호 = ""
            t검사ID = "":              t업체검사ID = ""
            t검체 = "":                t결과치 = ""
            t참조치 = "":              t소견 = ""
            t결과일 = "":              t입원외래구분 = ""

            .GetText 2, intRow, varTmp:    strDate = Trim$(varTmp)
            .GetText 3, intRow, varTmp:    strBarno = Trim$(varTmp)
            .GetText 4, intRow, varTmp:    strSPnm = Trim$(varTmp)
            .GetText 5, intRow, varTmp:    strSPid = Trim$(varTmp)
            .GetText 6, intRow, varTmp:    strChartNo = Trim$(varTmp)
            .GetText 7, intRow, varTmp:    strSeq = Trim$(varTmp)
            
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
                            .GetText intCol, intRow, varTmp
                            strTestcd = itemX.ListSubItems(1)
                            intPos = InStr(strTestcd, ",")
                            strEqpCd = ""
                        
                            spdResult1.Col = intCol
                            spdResult1.Row = intRow
                            Debug.Print spdResult1.CellNote
                            
                            m검사번호 = spdResult1.CellNote
                            m검사코드 = spdResult1.CellTag

                            strTestcd = Replace(strTestcd, ",", "")
                            varSEQNO = Split(strSeq, "|")
            
                            blnFlag = False
                            
                            '-- 이 부분에 데이터를 편집해서 넣는다.
                            
                            t업체 = "ACK"                     ' 업체
                            t요양기관번호 = "37348388"        ' 요양기관번호
                            t차트번호 = Trim(strChartNo)      ' 차트번호
                            t수진자명 = Trim(strSPnm)         ' 수진자명
                            t주민등록번호 = Trim(strBarno)    ' 주민등록번호
                            t내원번호 = Trim(varSEQNO(1))     ' 내원번호
                            t의뢰일 = Trim(strDate)           ' 의뢰일
                            t검사번호 = Trim(m검사번호)       ' 검사번호
                            t검사ID = m검사코드               ' 검사ID
                            t업체검사ID = ""                  ' 업체검사ID
                            t검체 = ""                        ' 검체
                            t결과치 = Trim(varTmp)            ' 결과치
                            t참조치 = ""                      ' 참조치
                            t소견 = ""                        ' 소견
                            t결과일 = Format(Now, "YYYYMMDD") ' 결과일
                            t입원외래구분 = Trim(varSEQNO(0)) ' 입원외래구분
                            
                            If .BackColor = &HC6FEFF Then
                                Print #100, t업체 & "|" & t요양기관번호 & "|" & t차트번호 & "|" & t수진자명 & "|" & t주민등록번호 & "|" & t내원번호 & "|" & t의뢰일 & "|" & t검사번호 & "|" & t검사ID & "|" & t업체검사ID & "|" & t검체 & "|" & t결과치 & "|" & t참조치 & "|" & t소견 & "|" & t결과일 & "|" & t입원외래구분, Chr(13) + Chr(10);
                                Call MakeXMLResult(t업체, t요양기관번호, t차트번호, t수진자명, t주민등록번호, t내원번호, t의뢰일, t검사번호, t검사ID, t업체검사ID, t검체, t결과치, t참조치, t소견, t결과일, t입원외래구분, tmpXML)
                            End If
                            
                            tmpXML = True
                                    
                            spdResult1.Row = intRow
                            spdResult1.Col = 2: spdResult1.BackColor = vbCyan
                            spdResult1.Col = 3: spdResult1.BackColor = vbCyan
                            spdResult1.Col = 4: spdResult1.BackColor = vbCyan
                            spdResult1.Col = 5: spdResult1.BackColor = vbCyan
                            spdResult1.Col = 6: spdResult1.BackColor = vbCyan
                            spdResult1.Col = 7: spdResult1.BackColor = vbCyan
                            spdResult1.Col = 1: spdResult1.Value = 0
                            
                            If strErrMsg = "" Then
                                sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y' , REFVAL = '" & strDate & "' " & _
                                         " where SPCNO   = '" & strChartNo & "'" & _
                                         "   and TRANSDT = '" & Format(dtpRsltDay.Value, "yyyymmdd") & "'"
                                AdoCn_Jet.Execute sqlDoc
                            Else
                                MsgBox strErrMsg, vbInformation, App.Title
                            End If

                            Set itemX = Nothing
                        End If
                    End If
                Next
            End If
        Next
    End With
    Me.MousePointer = 0
    
    MsgBox "▒ 업로더용 파일을 만들었습니다.. ▒      " & vbCrLf & vbCrLf & "    의사랑 프로그램에서 결과 가져오기 업무를 진행하십시요.  ", vbInformation, App.Title

    Exit Sub
ErrorRoutine:

    Set AdoRs_SQL = Nothing

    Set itemX = Nothing

    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
End Sub


Private Sub cmdReceve_Click()
    gRecodeType = "R"
    comEQP.Output = STX
End Sub

Private Sub cmdEot_Click()
    Call COM_OUTPUT(EOT)
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If spdResult2.DataRowCnt = 0 And spdResult2.DataRowCnt = 0 Then Exit Sub
    
    With spdResult2
        .Row = 0: .Row2 = .maxrows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .maxrows
    End With
 
    With tblexcel
        .maxrows = spdResult2.maxrows + 1
        .MaxCols = spdResult2.MaxCols
        .Row = 1: .Row2 = .maxrows
        .Col = 1: .Col2 = spdResult2.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.Filter = "ExCelFile(*.XLS)|*.XLS"
    CommonDialog1.FileName = REG_INSNAME & "  " & Format(dtpRsltDay, "####-##-##") & " 검사현황대장"
    CommonDialog1.ShowSave

    tblexcel.SaveTabFile (CommonDialog1.FileName)
    
End Sub

Private Sub cmdOrder_Click()
    Dim varTmp      As Variant
    Dim intRow      As Integer, intCol  As Integer
    Dim strBarno    As String, strTest  As String
    Dim strName     As String
    Dim strRack     As String, strCup   As String
    Dim intCnt      As Integer
    Dim itemX       As ListItem
    Dim strOrdList  As String
    Dim TestStatus(32)  As String * 1
    Dim xx              As Integer
    Dim intchannel As Integer
    Dim ssTemp2    As String
    Dim Loop_count, pDoCount As Integer
    
    comEQP.Output = STX
    gRecodeType = "O"
'    Timer2.Enabled = True
    
End Sub

Private Sub cmdPosNo_Click()
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
'                If Trim(sAdd + Val(sNo)) = 14 Then sNo = 0
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If
End Sub

Private Sub cmdPrint_Click()
Dim objclsCommon As New clsCommon

Dim Tmp_Testnm As String
Dim Row_cnt As Integer, Col_cnt As Integer, TmpPrintline As Integer
Dim vTmp As Variant
Dim stragesex As String

Const TmpLine = "────────────────────────────────────────────────────────────"

    If spdResult1.maxrows >= 1 Then
        With objclsCommon
            .PrintText 15, 3, Format(Date, "yyyy/mm/dd") & "  WorkList Report..( " & App.EXEName & " )", "Arial", 12
            
            .PrintText 0.5, 5, TmpLine
            .PrintText 0.5, 6, "순", , 9
            .PrintText 2, 6, "처방일자", , 9
            .PrintText 7, 6, "환자성명", , 9
            .PrintText 12, 6, "병록번호", , 9
            .PrintText 16, 6, "장비검사종목", , 9
            .PrintText 0.5, 7, TmpLine
            
            TmpPrintline = 8
        
        For Row_cnt = 1 To spdResult1.maxrows
            spdResult1.Row = Row_cnt
            
            If (Row_cnt Mod 34) <> 0 Then
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' 순
                spdResult1.Col = 2: .PrintText 2, TmpPrintline, Mid(spdResult1.Text, 3), , 9                    ' 처방일자
                spdResult1.Col = 4: .PrintText 7, TmpPrintline, Trim(spdResult1.Text), 9              ' 검체번호
                spdResult1.Col = 6: .PrintText 12, TmpPrintline, Trim(spdResult1.Text), , 9             ' 이    름
               ' spdResult1.Col = 2: .PrintText 16, TmpPrintline, Trim(spdResult1.text), , 9             ' 병원명
                
                
                For Col_cnt = 8 To spdResult1.MaxCols
            
                    spdResult1.Row = Row_cnt:            spdResult1.Col = Col_cnt
                    
                    If spdResult1.BackColor = &HC6FEFF Then
                        spdResult1.GetText Col_cnt, 0, vTmp
                        Tmp_Testnm = Tmp_Testnm & ", " & vTmp
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 16, TmpPrintline, Mid(Trim(Tmp_Testnm), 2), , 7.5
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
            Else
            
                '-------------------------------------------------------
            
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' 순
                spdResult1.Col = 2: .PrintText 2, TmpPrintline, Mid(spdResult1.Text, 3), , 9                   ' 처방일자
                spdResult1.Col = 4: .PrintText 6, TmpPrintline, Trim(spdResult1.Text), 9              ' 검체번호
                spdResult1.Col = 6: .PrintText 12, TmpPrintline, Trim(spdResult1.Text), , 9             ' 이    름
                
                
                For Col_cnt = 8 To spdResult1.MaxCols
            
                    spdResult1.Row = Row_cnt:            spdResult1.Col = Col_cnt
                    
                    If Trim(spdResult1.Text) <> "" Then
                        spdResult1.GetText Col_cnt, 0, vTmp
                        Tmp_Testnm = Tmp_Testnm & ", " & vTmp
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 16, TmpPrintline, Mid(Trim(Tmp_Testnm), 2), , 7.5
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
                
                '-------------------------------------------------------
            
                    .PrintText 0.5, TmpPrintline, TmpLine
                    .PrintText 1, TmpPrintline + 1, "── Next Report ──", , 9, True
                    Printer.NewPage
                    
                    .PrintText 0.5, 5, TmpLine
                    .PrintText 0.5, 6, "순", , 9
                    .PrintText 2, 6, "접수번호", , 9
                    .PrintText 6, 6, "환자성명", , 9
                    .PrintText 12, 6, "병록번호", , 9
                    .PrintText 16, 6, "처방일자", , 9
                    .PrintText 20, 6, "장비검사종목", , 9
                    .PrintText 0.5, 7, TmpLine
                    
                    TmpPrintline = 9
            End If
        
        Next Row_cnt
        .PrintText 0.5, TmpPrintline, TmpLine
        .PrintText 1, TmpPrintline + 1, "── End of Report ──", , 9, True
        
        End With
        Printer.NewPage
        Printer.EndDoc
        
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 WorkList가 Print되었습니다..       " & vbCrLf & vbCrLf & "다음 작업을 진행하십시요..", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 WorkList가  Load 되어 있지 않습니다..       " & vbCrLf & vbCrLf & "자료를 확인 하십시요..", vbInformation + vbOKOnly, App.Title
    End If
    
    '
    ' 마지막 저장
    '
    spdResult1.SaveTabFile App.Path & "\" & REG_INSNAME & "_Request.txt"

End Sub

Private Sub cmdRackNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer
Dim fNum1 As Integer, fNum2 As Integer
Dim intRow1 As Integer

AgainInput:
    fNum1 = 1: fNum2 = 0
    sNo = InputBox("시작 렉번호를 입력하세요!")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        sNo = UCase(sNo)
        
'        If Asc(sNo) < 65 Or Asc(sNo) > 70 Then
'            MsgBox "a~f까지의 문자만 입력하세요.!", vbCritical
'            GoTo AgainInput
'        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                intRow1 = intRow1 + 1
                .Row = sCnt
                .Col = 1
                If .Value >= 1 Then
                    .Col = 6
                    If intRow1 = (14 * fNum1) + 1 Then
                        fNum1 = fNum1 + 1: fNum2 = 0
                    End If
                    fNum2 = fNum2 + 1
                    .Text = Chr(fNum1 + Asc(sNo) - 1)
                    
                    .Col = 7
                    .Text = fNum2
                End If
            Next sCnt
        End With
    End If

End Sub

Private Sub cmdRequist_Click(Index As Integer)
    Dim ret As Integer
    
    ret = spdResult1.LoadFromFile(App.Path & "\" & REG_INSNAME & "_Request.txt")

End Sub

Private Sub cmdSearch_Click()
    Dim adoRS   As New ADODB.Recordset


    Dim strDoc As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim intRow      As Integer
    Dim pGrid_Point As Integer
    Dim intCnt      As Integer
    Dim strBarno As String
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim strBartmpNo As String
    Dim blt As Boolean
    Dim varXML      As Variant
    Dim varTmp      As Variant
    
    With spdWorklist
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    blt = True
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    Set AdoRs_ORACLE = New ADODB.Recordset
       
    '-- WorkList조회
    Dim strTime As String
    
    strTime = mskOrdtime.Text
    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), strTime)
    
    If UBound(varXML) <= 1 Then
        MsgBox dtpStartDt.Value & "일 에서  " & dtpStopDt.Value & "일까지의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarno = ""

        With spdWorklist
            For intCnt = 0 To UBound(varXML) - 1
                varTmp = Split(varXML(intCnt), ",")
                
                sqlDoc = "Select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
                         "  From INTERFACE003" & _
                         " Where SPCNO = '" & varTmp(2) & "'" & _
                         "   And REFVAL = '" & varTmp(6) & "'"
                
                adoRS.CursorLocation = adUseClient
                adoRS.Open sqlDoc, AdoCn_Jet
                
                If adoRS.RecordCount = 0 Then
                
                    strEqpCd = ""
                    
                    Debug.Print varTmp(8)
                    
                    strEqpCd = f_funGet_CODE(Trim(varTmp(8)))
                    
                   ' Debug.Print XMLData.ExamID
                    
                    If strEqpCd <> "" Then
                        XMLData.Company = varTmp(0)
                        XMLData.HospCode = varTmp(1)
                        XMLData.ChartNo = varTmp(2)
                        XMLData.patname = varTmp(3)
                        XMLData.PatJumin = varTmp(4)
                        XMLData.PatNo = varTmp(5)
                        XMLData.CommDate = varTmp(6)
                        XMLData.ExamNo = varTmp(7)
                        XMLData.ExamID = varTmp(8)
                        XMLData.ComExamID = varTmp(9)
                        XMLData.Specimen = varTmp(10)
                        XMLData.Result = varTmp(11)
                        XMLData.Reference = varTmp(12)
                        XMLData.Remark = varTmp(13)
                        XMLData.RsltDate = varTmp(14)
                        XMLData.IOFlag = varTmp(15)
                        
                        If strBarno <> XMLData.ChartNo Then
                            optBar.Value = True
                            pGrid_Point = SeqSearch(spdWorklist, XMLData.ExamNo, 6)
    
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(spdWorklist, XMLData.ExamNo, 6)
                                If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                            End If
    
                            .SetText 1, pGrid_Point, "0"
                            .SetText 2, pGrid_Point, XMLData.CommDate
                            .SetText 3, pGrid_Point, XMLData.PatJumin
                            .SetText 4, pGrid_Point, XMLData.patname
                            .SetText 6, pGrid_Point, XMLData.ChartNo
                            .SetText 7, pGrid_Point, XMLData.IOFlag & "|" & XMLData.PatNo
                            
                            Dim mSex As String
                            
                            If Mid(XMLData.PatJumin & "", 8, 1) = " " Then
                                mSex = ""
                            Else
                                mSex = Mid(XMLData.PatJumin & "", 8, 1)
                                Select Case mSex
                                    Case 1, 3
                                        .SetText 5, pGrid_Point, "M"
                                    Case 2, 4
                                        .SetText 5, pGrid_Point, "F"
                                End Select
                            End If
                            
                            .Row = pGrid_Point: .Col = 1: .ForeColor = HNC_Black
                                                .Col = 2: .ForeColor = HNC_Black
                                                .Col = 4: .ForeColor = HNC_Black
                                                .Col = 5: .ForeColor = HNC_Black
                                                .Col = 6: .ForeColor = HNC_Black
                        End If
    
                        strEqpCd = f_funGet_CODE(Trim(XMLData.ExamID))
                        
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                        If Not itemX Is Nothing Then
                            spdWorklist.SetText 1, pGrid_Point, "0"
                            
                            spdWorklist.Col = itemX.Index + 9
                            spdWorklist.Row = pGrid_Point
                            spdWorklist.BackColor = &HC6FEFF   '&HC6FEFF
                            spdWorklist.CellTag = XMLData.ExamID
                            spdWorklist.CellNote = XMLData.ExamNo
                            
'                           spdWorklist.text = XMLData.ExamNo
                            
                            
    '                        spdWorklist.Row2 = itemX.Index + 7
    '                        spdWorklist.Col2 = pGrid_Point
    '                        spdWorklist.text = XMLData.ExamNo
                            blt = True
                        End If
                        strBarno = XMLData.ChartNo
                    End If
                End If
                
                adoRS.Close:    Set adoRS = Nothing
                
            Next
        End With
    End If
    
    spdWorklist.Row = 1
    spdWorklist.Col = 1
    spdWorklist.Action = ActionActiveCell
    
    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    Dim strRack_tmp As String

    
    txtChart.ForeColor = &HFFC0C0
    txtChart.Text = "차트번호 입력"
    

    
Exit Sub

ErrorTrap:

    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdACK_Click()
'
'    Call COM_OUTPUT(charCOM_Convert(COM_ACK))
Call COM_OUTPUT(Chr(1))
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
    Dim Colcnt1 As Integer
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    List1.Clear
    txtChart.ForeColor = &HFFC0C0
    txtChart.Text = "차트번호 입력"
    txtSEQ.Text = "1"
    
    With spdWorklist

        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
        .maxrows = 1
        
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

    With spdResult2
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    Dim Rowcnt As Integer
    Dim Colcnt As Integer

    With spdRstview
        For Rowcnt = 1 To 8
            For Colcnt = 2 To 6 Step 2
                .Row = Rowcnt
                .Col = Colcnt
                .BackColor = &HFFFFFF
                .Text = ""
            Next Colcnt
        Next Rowcnt
    End With

    For Colcnt1 = 6 To spdResult1.MaxCols
        spdResult1.Row = 1
        spdResult1.Col = Colcnt1
        spdResult1.CellNote = ""
        spdResult1.CellTag = ""
        
        spdResult1.Text = ""
    Next Colcnt1

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

'Private Sub cmdAppend_Click(Index As Integer)
'
'    Dim adoRS   As New ADODB.Recordset
'    Dim sqlDoc  As String
'
'    Dim varTmp  As Variant, strErrMsg   As String
'    Dim strSampleno()   As String, strBarno     As String, strTime      As String
'    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
'    Dim strTmp1()       As String, strTmp2()    As String
'    Dim intPos          As String, strTestcd    As String, strTestRst   As String
'    Dim strTestnm       As String
'    Dim strRef          As String
'    Dim strUnit         As String
'    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
'
'    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
'    Dim itemX   As ListItem
'    Dim objSpd  As vaSpread
'    Dim sqlRet  As Integer
'    Dim flgSave As Boolean
'    Dim SaveGbn As Integer
'    Dim strDate As String
'
'    CallForm = "frmComm - Private Sub cmdAppend_Click()"
'
'On Error GoTo ErrorRoutine
'
'    Me.MousePointer = 11
'
'    If Index = 0 Then
'        Set objSpd = spdResult1
'    Else
'        Set objSpd = spdResult2
'    End If
'
'    With objSpd
'        For intRow = 1 To .maxrows
'
''            .GetText 2, intRow, varTmp:         strDate = Trim$(varTmp)
''            .GetText 3, intRow, varTmp:         strBarno = Trim$(varTmp)
''            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
'
'            .GetText 2, pGrid_Point, varTmp:   strDate = Trim$(varTmp)
'            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
'            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
'            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
'
'
'            .GetText 1, intRow, varTmp
'
'            If strBarno = "" Then Exit For
'
'            intCnt = 0: Erase strOrdcd: Erase strRstval
'            If Trim$(varTmp) = "1" Then
'                For intCol = 6 To .MaxCols
'                    .GetText intCol, intRow, varTmp
'                    If Trim$(varTmp) <> "" Then
'                        .GetText intCol, 0, varTmp
'                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                        If Not itemX Is Nothing Then
'                            .GetText intCol, intRow, varTmp
'                            strTestcd = itemX.ListSubItems(1)
'                            intPos = InStr(strTestcd, ",")
'                            If intPos > 0 Then
'                                Do While intPos > 0
'
'                                    blnFlag = False
'                                    Set mAdoRs = f_subSet_TestList(strBarno)
'                                    Do Until mAdoRs.EOF
'                                        If mAdoRs("LCODE") = Mid$(strTestcd, 1, intPos - 1) Then blnFlag = True: Exit Do
'                                        mAdoRs.MoveNext
'                                    Loop
'                                    strTestcd = Mid$(strTestcd, intPos + 1)
'                                    intPos = InStr(strTestcd, ",")
'
'                                    AdoCn_ORACLE.BeginTrans
'
'                                    sqlDoc = "insert into lab_result" & _
'                                             "   (RESULTNO, LCODE, LSEQ, LNAME, LRESULT, UNIT, REFV,LTYPE,RESULT_DATE, REPORTER, PT_SEQ) " & _
'                                             " values('" & mAdoRs("resultno") & "', '" & strTestcd & "'," & _
'                                             "       '1', '" & Trim(itemX.ListSubItems(2)) & "'," & _
'                                             "       '" & Trim(varTmp) & "', '" & Trim(itemX.ListSubItems(9)) & "'," & _
'                                             "       '" & Trim(itemX.ListSubItems(8)) & "','0',sysdate," & _
'                                             "       '70001','" & mAdoRs("pt_seq") & "')"
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    sqlDoc = ""
'                                    sqlDoc = sqlDoc & "update ipd_order_date set req_result2 = '*'"
'                                    sqlDoc = sqlDoc & " where patient_no = '" & mAdoRs("patient_no") & "'"
'                                    sqlDoc = sqlDoc & "   and order_date = to_date('" & strDate & "','yyyy-mm-dd')"
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    sqlDoc = ""
'                                    sqlDoc = sqlDoc & "update lab_order set r_flag='1'"
'                                    sqlDoc = sqlDoc & " where resultno = '" & mAdoRs("resultno") & "'"
'                                    sqlDoc = sqlDoc & "   and patient_no = '" & mAdoRs("patient_no") & "'"
'                                    sqlDoc = sqlDoc & "   and lcode = '" & mAdoRs("lcode") & "'"
'                                    sqlDoc = sqlDoc & "   and pt_seq = '" & mAdoRs("pt_seq") & "'"
'
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    AdoCn_ORACLE.CommitTrans
'
'                                    lblStatus.Caption = "저장 성공!!"
'                                    Set adoRS = Nothing:    mAdoRs.Close
'                                Loop
'                            Else
'                                blnFlag = False
'                                Set mAdoRs = f_subSet_TestList(strBarno)
'                                Do Until mAdoRs.EOF
'                                    If Trim(mAdoRs("LCODE")) = strTestcd Then blnFlag = True: Exit Do
'                                    mAdoRs.MoveNext
'                                Loop
'                                If blnFlag Then
'                                    AdoCn_SQL.BeginTrans
'
'                                    If chkAuto.Value = "1" Then
'                                           If Mid(pName, 1, 2) = "검진" Then
'                                               sqlDoc = "Update MDCK..GUMJIN_INTERFACE" & _
'                                                        "   set RESULT = '" & strRstval & "'," & _
'                                                        "       ACT_RETURN_DATE = '" & strDate & "'" & _
'                                                        " where PER_GUMJIN_DATE = '" & strDate2 & "'" & _
'                                                        "   and PER_GUM_NUM = " & pNo & "" & _
'                                                        "   and EDPSCODE = '" & Mid(itemX.text, 1, 4) & "'"
'                                           Else
'                                               sqlDoc = "Update MEDICOM..jun370_resulttb" _
'                                                       & "   Set Result = '" & strRstval & "', status='1'" _
'                                                       & " Where WaitSeqNo = '" & pNo & "'" _
'                                                       & "   and map2seqno = '" & strEqpCd & "'"
'
'                                           End If
'                                           AdoCn_SQL.Execute sqlDoc
'                                    End If
'
'                                    AdoCn_SQL.CommitTrans
'
'                                    lblStatus.Caption = "저장 성공!!"
'                                    Set adoRS = Nothing:    mAdoRs.Close
'                                End If
'                            End If
'                        End If
'
'                        Set itemX = Nothing
'                    End If
'                Next
'                spdResult1.Row = intRow
'                spdResult1.Col = 2
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 3
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 4
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 1: spdResult1.Value = 0
'
'                If strErrMsg = "" Then
'                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
'                             " where SPCNO   = '" & strBarno & "'" & _
'                             "   and TRANSDT = '" & mskRstDate.text & "'"
'                    AdoCn_Jet.Execute sqlDoc
'                Else
'                    MsgBox strErrMsg, vbInformation, Me.Caption
'                End If
'            End If
'        Next
'    End With
'    Me.MousePointer = 0
'    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
'
'    Exit Sub
'ErrorRoutine:
'    Set itemX = Nothing
'
'    Me.MousePointer = 0
'    Call ErrMsgProc(CallForm)
'End Sub

Public Function CheckSum_ECi_Tx(ByVal strPrmValue As String)

    Dim I                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For I = 1 To intValueLength
        intCheck = intCheck + Asc(Mid(strPrmValue, I, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        CheckSum_ECi_Tx = "0" & strCheck
    Else
        CheckSum_ECi_Tx = Right(strCheck, 2)
    End If

End Function

Public Function LenA(strPrmString As String) As Integer

    Dim I                   As Integer
    Dim intStrLen           As Integer
    Dim intAnsiStrLen       As Integer
    Dim strTemp             As String
    
    intStrLen = Len(strPrmString)
    For I = 1 To intStrLen
        strTemp = Mid(strPrmString, I, 1)
        
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
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "Select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  From INTERFACE003" & _
             " Where TRANSDT >= '" & Format(dtpRsltDay.Value, "yyyymmdd") & "'" & _
             "   And EQUIPCD = '" & INS_CODE & "'"
'    If cboRstgbn(1).ListIndex = 0 Then
'        sqlDoc = sqlDoc & "   And SERVERGBN = ''"
'    ElseIf cboRstgbn(1).ListIndex = 1 Then
'        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
'    End If
    sqlDoc = sqlDoc & " Order By SPCNO, TRANSTM"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdResult2
        If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(9) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 14
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(3) & "")
                .SetText 3, intRow, Trim$(adoRS(0) & "")
                .SetText 6, intRow, Trim$(adoRS(8) & "")
                .SetText 7, intRow, Trim$(adoRS(9) & "")
                '.SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(9) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 8
                .SetText intCol, intRow, Trim$(adoRS(4)) & ""
                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
            End If
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    If Index = 2 Or Index = 3 Then
        With spdResult1
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    Else
        With spdWorklist
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    End If
    
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
                .Col = 7:       .Text = Trim(sAdd + Val(sNo))
                txtSeqNo.Text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Function f_subSet_TestList(ByVal strDate As String, ByVal strSeq As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
        Set AdoRs_SQL = New ADODB.Recordset
               
                 sqlDoc = "select EDPSCODE from GUMJIN_INTERFACE"
        sqlDoc = sqlDoc & " where PER_GUMJIN_DATE = '" & strDate & "'"
        sqlDoc = sqlDoc & "   and PER_GUM_NUM = '" & strSeq & "'"
        sqlDoc = sqlDoc & "   and EDPSCODE in ('0208','0226','0225','0227','0207','0206','0205','0221','0222','0223','0224','0209')"
        sqlDoc = sqlDoc & "   and RESULT=''"
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_TestList = Nothing
        Else
            Set f_subSet_TestList = AdoRs_SQL
        End If
    
        Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub cmdWordQuery_Click()
'    On Error GoTo ErrRoutine
'    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
'
'    Dim strKeyno    As String
'    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
'    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
'    Dim strEqpCd    As String
'    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
'    Dim itemX   As ListItem
'
'    '-- WorkList조회
'    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text)
'
'    If RecordChk = False Then
'        Exit Sub
'    End If
'
''    With spdWorkList
''        .maxrows = 14
''        .Col = 1:   .Col2 = .MaxCols
''        .Row = 1:   .Row2 = .maxrows
''        .BlockMode = True
''        .Action = ActionClearText
''        .BlockMode = False
''        .RowHeight(-1) = 12
''    End With
'
'    intRow = 0
'    Do Until mAdoRs.EOF
'        intIdx = 0
'        With spdResult1
'            If strKeyno <> mAdoRs.Fields("EXAM_NO") Then
''                intRow = SeqNullSearch(spdResult1, "", 1)
''                If intRow = "0" Then
''                    .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
''                    intRow = .maxrows
''                Else
''                    intRow = intRow + 1
''                End If
'                intRow = intRow + 1
'                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'
'                .SetText 1, intRow, "1"
'                .SetText 2, intRow, Trim(mAdoRs("REQUEST_DATE")) & ""
'                '.SetText 3, intRow, Trim(cboRegGbn.Text) & ""
'                .SetText 4, intRow, Trim(mAdoRs("PERSON_NAME")) & ""
'                .SetText 5, intRow, Trim(mAdoRs("EXAM_NO")) & ""
'                .SetText 6, intRow, Trim(mAdoRs("CHART_NO")) & ""
'                '.SetText 6, intRow, Trim(mAdoRs("COMPANY_NAME"))
'
'                '-- 검사항목조회
''                blnFlag = False
'                Set mAdoRs1 = f_subSet_TestList(mAdoRs.Fields("EXAM_NO"))
'                If Len(mAdoRs.Fields("EXAM_NO")) > 0 Then
'                    Do Until mAdoRs1.EOF
'                        strEqpCd = f_funGet_CODE(Trim(mAdoRs1("EXAM_CODE")))
'                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                        If Not itemX Is Nothing Then
''                            blnFlag = True
'                            spdResult1.Row = intRow
'                            spdResult1.Col = itemX.Index + 6
'                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
'                            DoEvents
'                        End If
'                        mAdoRs1.MoveNext
'                    Loop
'                End If
'
'            End If
'            strKeyno = mAdoRs("EXAM_NO")
'        End With
'        intIdx = intIdx + 1
'        mAdoRs.MoveNext
'    Loop
'    Exit Sub
'
'ErrRoutine:
'
'    Call ErrMsgProc(CallForm)

End Sub


Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String, strChartNo As String, strSex As String
    Dim strWDate As String
    Dim strEqpCd    As String
    Dim tmpDate     As String
    
    Dim iCnt As Integer

    blnFlag = False
    
    With spdWorklist
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
            
                .GetText 2, intRow1, varTmp:    strWDate = Trim$(varTmp)
                .GetText 3, intRow1, varTmp:    strBarno = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 5, intRow1, varTmp:    strSPid = Trim$(varTmp)
                .GetText 6, intRow1, varTmp:    strChartNo = Trim$(varTmp)
                .GetText 5, intRow1, varTmp:    strSex = Trim$(varTmp)
                
                intRow2 = f_funGet_SpreadRow(spdResult1, 6, strChartNo)
                

                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 3, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 14
                        intRow2 = spdResult1.maxrows
                    End If
                End If
                
                For iCnt = 1 To .MaxCols
                    
                    .GetText iCnt, intRow1, varTmp
                    
                    .Col = iCnt
                    .Row = intRow1
                     
                     Dim tTag As String
                     Dim tTestcd As String
                     
                     tTag = .CellNote
                     tTestcd = .CellTag
                    
                    If .BackColor = &HC6FEFF Then
                        spdResult1.Row = intRow2
                        spdResult1.Col = iCnt
                        spdResult1.BackColor = &HC6FEFF
                        spdResult1.CellTag = tTestcd
                        spdResult1.CellNote = tTag
                        
                        
                        Rem spdResult1.text = spdWorklist.CellTag
                        
                        Rem spdResult1.SetText iCnt, intRow1, tTag
                        
                    End If
                    
                    If iCnt = 1 Then
                    Else
                        If iCnt < 8 Then
                            spdResult1.SetText iCnt, intRow2, varTmp
                        End If
                    End If
                    .Row = intRow1
                    .Col = 1: .ForeColor = HNC_Red
                    .Col = 2: .ForeColor = HNC_Red
                    .Col = 4: .ForeColor = HNC_Red
                    .Col = 5: .ForeColor = HNC_Red
                    .Col = 6: .ForeColor = HNC_Red
                    spdResult1.SetText 8, intRow2, txtSEQ.Text
                    .SetText 1, intRow1, 0
                    
                Next iCnt
                txtSEQ.Text = txtSEQ.Text + 1
            End If
        Next
    End With
                
End Sub


'
'Private Sub cmdWorkList_Click()
'
'    Dim varTmp  As Variant
'    Dim intRow1 As Integer, intRow2 As Integer
'    Dim intIdx  As Integer
'    Dim Rev     As Long
'    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
'    Dim itemX As ListItem
'    Dim blnFlag As Boolean
'    Dim strBarno    As String, strSPid  As String, strSPnm   As String, strChartNo As String, strSex As String
'    Dim strWDate As String
'    Dim strEqpCd    As String
'    Dim tmpDate     As String
'
'    blnFlag = False
'
'    With spdWorklist
'        For intRow1 = 1 To .maxrows
'            .GetText 1, intRow1, varTmp
'            If Trim$(varTmp) = "1" Then
'                .GetText 2, intRow1, varTmp:    strWDate = Trim$(varTmp)
'                .GetText 3, intRow1, varTmp:    strBarno = Trim$(varTmp)
'                .GetText 4, intRow1, varTmp:    strSPnm = Trim$(varTmp)
'                .GetText 5, intRow1, varTmp:    strSPid = Trim$(varTmp)
'                .GetText 5, intRow1, varTmp:    strChartNo = Trim$(varTmp)
'                .GetText 6, intRow1, varTmp:    strSex = Trim$(varTmp)
'
'                '
'                ' WorkLIST 적용여부
'                '
'
'                .Row = intRow1:
'
'                .Col = 1: .ForeColor = HNC_Red
'                .Col = 2: .ForeColor = HNC_Red
'                .Col = 4: .ForeColor = HNC_Red
'                .Col = 5: .ForeColor = HNC_Red
'                .Col = 6: .ForeColor = HNC_Red
'
'                intRow2 = f_funGet_SpreadRow(spdResult1, 6, strSPid)
'                If intRow2 < 1 Then
'                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
'                    If intRow2 < 1 Then
'                        spdResult1.maxrows = spdResult1.maxrows + 1
'                        spdResult1.RowHeight(spdResult1.maxrows) = 13
'                        intRow2 = spdResult1.maxrows
'                    End If
'
'                    blnFlag = False
'
'                    tmpDate = Mid(strWDate, 1, 4) & Mid(strWDate, 6, 2) & Mid(strWDate, 9, 2)
'
''                    Set mAdoRs = f_subSet_WorkList_Barcode(tmpDate, strSPid, strSPnm)
'
'
'                    If cboChk.ListIndex = 0 Then
'                        If RecordChk = True Then
'                            Do Until mAdoRs.EOF
'
'                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("EXAM_CODE")))
'
'                                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                                If Not itemX Is Nothing Then
'                                    blnFlag = True
'                                    spdResult1.Row = intRow2
'                                    spdResult1.Col = itemX.Index + 7
'                                    spdResult1.BackColor = &HC6FEFF '&H80C0FF
'                                    spdResult1.text = " "
'
'                                    DoEvents
'                                End If
'                                mAdoRs.MoveNext
'                            Loop
'                        End If
'                    Else
'                        If RecordChk = True Then
'                            Do Until mAdoRs.EOF
'
'                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("처방코드")))
'
'                                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                                If Not itemX Is Nothing Then
'                                    blnFlag = True
'                                    spdResult1.Row = intRow2
'                                    spdResult1.Col = itemX.Index + 7
'                                    spdResult1.BackColor = &HC6FEFF '&H80C0FF
'                                    spdResult1.text = " "
'
'                                    DoEvents
'                                End If
'                                mAdoRs.MoveNext
'                            Loop
'                        End If
'                    End If
'                    If blnFlag = True Then
'                        Dim tmpSeq As String
'                        tmpSeq = txtSeqNo.text + 1
'
'                        spdResult1.SetText 2, intRow2, strWDate
'                        spdResult1.SetText 4, intRow2, strSPnm
'                        spdResult1.SetText 3, intRow2, strBarno
'                        spdResult1.SetText 5, intRow2, strSex
'                        spdResult1.SetText 6, intRow2, strChartNo
'                        spdResult1.SetText 7, intRow2, tmpSeq
'
'                        spdResult1.Row = intRow2:
'                        spdResult1.Col = 7:
'                        spdResult1.ForeColor = HNC_Red
''                        txtSeqNo.text = tmpSeq
'                    Else
'                        spdResult1.maxrows = spdResult1.maxrows - 1
'                    End If
'                End If
'
'                ' spdResult1.SetText 1, intRow2, "1"
'
'                .SetText 1, intRow1, ""
'                If tmpSeq <> "" Then
'                    txtSeqNo.text = tmpSeq
'                End If
'            End If
'        Next
'    End With
'
'End Sub

'Private Sub Command2_Click()
'                comEQP.Output = ACK
'End Sub

Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim TmpYesno As String
Dim Tmpptno, TmpPtnm As String

    If Row = 0 Then
    
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(spdResult1, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(spdResult1, Col, 1)
            OrderSort_Flag = 1
        End If
        
        Exit Sub
    End If


    If Col = 4 Or Col = 6 Then
        With spdResult1
            .Row = Row
            
            ' 병록번호 불러오기
            .Col = 6
            Tmpptno = .Text
            
            ' 환자이름 불러오기
            .Col = 4
            TmpPtnm = .Text
        End With
        
        If Len(Trim(Tmpptno)) >= 1 And Len(Trim(TmpPtnm)) >= 1 Then
             TmpYesno = MsgBox(Tmpptno & " (  " & TmpPtnm & "  ) " & " 환자를 선택 하셨습니다..    " & vbCrLf & vbCrLf & "검사를 제외 하시겠습니까..??", vbCritical + vbYesNo, App.Title)
        
             If TmpYesno = vbYes Then
                spdResult1.Action = ActionDeleteRow
                spdResult1.maxrows = spdResult1.maxrows - 1
             End If
        End If
    End If
        
End Sub

Private Sub spdResult1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 검사자 추가", "-", " ▒ 검사자 삭제", "-", " ▒ 시작번호수정", "-", " ▒ 서버 저장")

    Select Case lMenuChosen
        Case 1
            With spdResult1
                .maxrows = .maxrows + 1
                .Col = Col
                .Row = Row
                .Action = ActionInsertRow
            End With
        Case 3
            With spdResult1
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End With
        Case 5
            Call cmdStartNo_Click
        Case 7
            Call cmdAppend_Click(0)
    End Select
End Sub


Private Sub ComReceive(ByRef RecData As String)
    Dim sStxCheck As Integer, sEnqCheck As Integer, sEtxCheck As Integer, sLfCheck As Integer, sCrcheck As Integer
    Dim com_sTemp As String, Orderoutput As String, OrderLst As String
    Dim SequenceNo As Integer
    Dim MHead As String, Pinfo As String
    Dim OutPutData As String, sOrderLst As String
    Dim ii As Integer, pDoCount As Integer, Loop_count As Integer
    Dim strBarno As String
    Dim intCnt As Integer
    Dim pGrid_Point  As Integer
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim iCol As Integer, iRow As Integer
    Dim varTmp
    Dim strOrd  As String


    On Error GoTo errOnComm

    fRcvString = fRcvString + RecData
    
    If left(RecData, 1) = STX Or left(RecData, 1) = EOT Then
        comEQP.Output = ACK
        fRcvString = ""
    End If
    
    Debug.Print fRcvString
    
    com_sTemp = fRcvString
    
    For ii = 1 To Len(com_sTemp)
        Select Case Mid(com_sTemp, ii, 1)
            Case NAK:
                SendCount = 0
            Case STX:
            Case EOT:
                Call ReceiveTheData(com_sTemp, fChannel(), spdResult1)
                comEQP.Output = ACK
                fRcvString = ""
            Case vbCr
        End Select
    Next ii

'    If left(RecData, 1) = EOT Or Mid(RecData, 2, 1) = EOT Or _
'                                 Mid(RecData, 3, 1) = EOT Or _
'                                 Mid(RecData, 4, 1) = EOT Or _
'                                 Mid(RecData, 4, 1) = EOT Or _
'                                 Mid(RecData, 5, 1) = EOT Or _
'                                 Mid(RecData, 6, 1) = EOT Then
'        comEQP.Output = STX
'    End If


    If left(RecData, 1) = ACK Then

        Debug.Print RecData
        Dim strSeq As String
        Dim strOrdItem As String
        Dim strSendYN As String
        strOrdItem = ""
        intCnt = 0
        
        With spdResult1
            Dim strOrdCount As Integer
            strOrdCount = 0
            
            For iRow = 1 To .maxrows
                .GetText 9, iRow, varTmp: strSendYN = Trim$(varTmp)
                .GetText 6, iRow, varTmp: strBarno = Trim$(varTmp)
                If strSendYN = "" And strBarno <> "" Then
                    strOrdCount = strOrdCount + 1
                End If
            Next
            
            If gRecodeType = "O" And strOrdCount > 0 Then
                For iRow = 1 To .maxrows
                    .GetText 6, iRow, varTmp: strBarno = Trim$(varTmp)
                    .GetText 8, iRow, varTmp: strSeq = Trim$(varTmp)
                    .GetText 9, iRow, varTmp:
                    
                    If Trim$(varTmp) = "" Then
                        For iCol = 10 To .MaxCols
                            .Row = iRow: .Col = iCol
                            strEqpCd = f_funGet_CODE(Trim(.CellTag))
                            
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                
                                If .BackColor = &HC6FEFF Then
                                    strOrdItem = strOrdItem & SetSpace(itemX.tag, 4, 2)
                                    intCnt = intCnt + 1
                                End If
                            End If
    
                        Next
                          
                        strOrd = ""
                        strOrd = SetSpace(Trim(strBarno), 15) & "TSN"
                        strOrd = strOrd & Format(Trim(strSeq), "0#") & Format(CStr(intCnt), "0#") & strOrdItem
                        strOrd = strOrd & CSum(strOrd) & EOT
                        .SetText 9, iRow, "오더"
                        comEQP.Output = strOrd
                        fRcvString = ""
                        Exit For
                        
                    End If
                Next
            ElseIf gRecodeType = "O" Then
            
                comEQP.Output = "R" & EOT
                
            End If
            
        End With
    End If
    
    Exit Sub
    
errOnComm:
    Call ErrMsgProc("통신에러")
End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
    Dim Arr()       As Byte
    Dim sStxCheck As Integer, sEnqCheck As Integer, sEtxCheck As Integer
    Dim sLfCheck As Integer, sCrcheck As Integer, ii As Integer
    Dim MHead As String, Pinfo As String, OutPutData As String, com_sTemp As String
    Dim strRec  As String, strBuff  As String
    
    Dim intIdx     As Integer, intIdx1    As Integer, intIdx2     As Integer
    Dim strTmp1     As String, strTmp2      As String
    Dim intPos1     As Integer, intPos2     As Integer
    Dim intCnt       As Integer
   
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            strDta = comEQP.Input
            Call ComReceive(strDta)
                                        
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

Private Sub ReceiveTheData(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)

    Dim sTemp       As String
    Dim Channel_No  As String
    Dim Patiant_No  As String
    Dim pGrid_Point As Integer
    Dim Max_Arary_Cnt As Integer
    Dim sDeCnt      As Integer
    Dim pDoCount    As Integer
    Dim Loop_count  As Integer
    Dim sRtn As Integer
    Dim sChannel As String
    Dim sRstText As String
    Dim sRstValue As Single
    Dim sUnit As String
    Dim intIdx As Integer
    Dim strEqpCd As String
    Dim itemX   As ListItem
    Dim sCol As Integer
    Dim varTmp As Variant
    Dim sSeq, strTmp, strBarno, strDate, strTime, strDate1 As String
    Dim intCol As Integer
    Dim strRstval As String, strRefVal  As String
    Dim mstrRstval As String
    Dim mstrDBSEQ As String
    Dim sqlDoc  As String
    Dim iDataChk As Integer
    Dim BooData As Boolean
    Dim HLBOOL As Boolean
    Dim valEqpcd As Variant
    Dim intCnt As Integer
    Dim strTransDT As String, strTransTM As String

    On Error Resume Next
'    On Error GoTo errDefine

    strTransDT = Format(Now, "YYYYMMDD")
    strTransTM = Format(Now, "hhMMss")
    
    sPatiant_No = Right(Trim(Mid(strdata, 1, 15)), 5)
    sSeq = Trim(Mid(strdata, 18, 3))
    

    With spdResult1
        pGrid_Point = SeqSearch(spdResult1, sPatiant_No, 6)

        .GetText 2, pGrid_Point, varTmp:    strDate = Trim$(varTmp)
        .GetText 4, pGrid_Point, varTmp:    PName = Trim$(varTmp)
        .GetText 6, pGrid_Point, varTmp:    strBarno = Trim$(varTmp)
        .GetText 3, pGrid_Point, varTmp:    pNo = Trim$(varTmp)

        strTmp = Mid(strdata, 22)

        Do While Len(strTmp) >= 11
            Channel_No = Trim(left(strTmp, 3))
            strRstval = Trim(Mid(strTmp, 4, 8))
            
            Select Case UCase(Channel_No)
                Case "BUN", "CRE", "HDL", "T-B", "ALB", "T-P"
                    strRstval = Format(strRstval, "####0.0")
                Case Else
                    strRstval = Format(strRstval, "####0")
                
            End Select

            If pGrid_Point > 0 Then
                For intCol = 7 To .MaxCols
                    .GetText intCol, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        If Len(Trim(strRstval)) > 0 Then
                            If Channel_No = itemX.tag Then
                                
                                .SetText intCol, pGrid_Point, strRstval
                                
                                If Len(Trim(itemX.SubItems(8))) <> 0 And Len(Trim(itemX.SubItems(9))) <> 0 Then
                                    If Val(strRstval) < Val(itemX.SubItems(8)) Then
                                        strRefVal = "L"
                                    ElseIf Val(strRstval) > Val(itemX.SubItems(9)) Then
                                        strRefVal = "H"
                                    Else
                                        strRefVal = ""
                                    End If
                                    .Col = intCol:  .Row = pGrid_Point
                                    If strRefVal = "H" Then
                                          .ForeColor = IIf(Trim$(strRefVal) = "H", vbRed, vbBlack)
                                    Else
                                          .ForeColor = IIf(Trim$(strRefVal) = "L", vbBlue, vbBlack)
                                    End If
                                End If
                                
                                
                                
                                If pGrid_Point <> 0 Then
                                  sqlDoc = "Update INTERFACE003" & _
                                           "   set RSTVAL  = '" & mstrRstval & "', REFVAL = '" & strRefVal & "'" & _
                                           " where SPCNO   = '" & strBarno & "'" & _
                                           "   and EQPNUM  = '" & itemX.tag & "'" & _
                                           "   and TRANSDT = '" & strDate & "'" & _
                                           "   and TRANSTM = '" & strTime & "'"
                                  AdoCn_Jet.Execute sqlDoc

                                  If cboChk.ListIndex = 0 Then
                                      sqlDoc = "insert into INTERFACE003(" & _
                                               "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                               "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                               "            '" & strDate1 & "', '" & strTime & "'," & _
                                               "            '" & strRstval & "', '" & strRefVal & "'," & _
                                               "            '" & INS_CODE & "', '', '" & PName & "', '" & pNo & "')"
                                  Else
                                      sqlDoc = "insert into INTERFACE003(" & _
                                               "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                               "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                               "            '" & strTransDT & "', '" & strTransTM & "'," & _
                                               "            '" & strRstval & "', '" & strRefVal & "'," & _
                                               "            '" & INS_CODE & "', '', '" & PName & "', '" & pNo & "')"
                                  End If
                                  Debug.Print sqlDoc
                                  AdoCn_Jet.Execute sqlDoc
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            
            If pGrid_Point <> 0 Then
                spdResult1.Row = pGrid_Point
                spdResult1.Col = 9: spdResult1.BackColor = vbCyan
                spdResult1.Col = 9: spdResult1.Text = "검사"
                spdResult1.Col = 1: spdResult1.Text = "1"
            End If

            strTmp = Mid(strTmp, 12)
        Loop
    End With
'    sRstText = strdata
'    HLBOOL = False
'
'    For Loop_count = 1 To 100: fSens(Loop_count) = "": Next Loop_count
'
'    pDoCount = 0
'    Do While InStr(sRstText, "|") > 0
'        pDoCount = pDoCount + 1
'        fSens(pDoCount) = Text_Redefine(sRstText, "|")
'        sRstText = Mid$(sRstText, InStr(sRstText, "|") + 1)
'        If pDoCount > 99 Then
'            sRstText = ""
'            Exit Do
'        End If
'    Loop
'
'    sRstText = ""
'
'    If Mid$(fSens(1), 3, 1) = "H" Then
'        Flag_HQL = ""
'        Flag_HQL = Flag_HQL & Mid(strdata, 3, 1)
'    ElseIf Mid$(fSens(1), 3, 1) = "O" Then
'        Flag_HQL = ""
'        Flag_HQL = Flag_HQL & Mid(strdata, 3, 1)
'        sPatiant_No = ""
'       Rem  sPatiant_No = Mid(Trim(Text_Redefine(fSens(3), "^")), 1, Len(Trim(Text_Redefine(fSens(3), "^"))))
'    ElseIf Mid$(fSens(1), 3, 1) = "R" Then
'        Dim mType As String
'
'        fSens(3) = Text_Change(fSens(3), "^", "")                  ' channel
'
'        Channel_No = fSens(3)                                         ' channel
'
'        intRow = 0
'        pGrid_Point = 0
'
'        With spdResult1
'            sCol = 8
'
'            pGrid_Point = SeqSearch(spdResult1, sPatiant_No, sCol)
'
'            .GetText 2, pGrid_Point, varTmp:    strDate = Trim$(varTmp)
'            .GetText 4, pGrid_Point, varTmp:    PName = Trim$(varTmp)
'            .GetText 7, pGrid_Point, varTmp:    strBarno = Trim$(varTmp)
'            .GetText 3, pGrid_Point, varTmp:    pNo = Trim$(varTmp)
'
'            If pGrid_Point > 0 Then
'                For intCol = 7 To .MaxCols
'                    strRstval = ""
'                    strEqpCd = ""
'                    .GetText intCol, 0, varTmp
'                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                    If Not itemX Is Nothing Then
'                        For intIdx = 10 To .MaxCols
'                            If Len(Trim(fSens(4))) > 0 Then
'                                If Channel_No = itemX.tag Then
'
'                                    mstrDBSEQ = ""
'                                    mstrRstval = ""
'                                    strRstval = Trim(fSens(4))
'                                    strRefVal = ""
'                                    mstrRstval = strRstval
'
'                                    .Row = pGrid_Point: .Col = intCol
'                                    mstrDBSEQ = .CellTag
'
'                                    strDate1 = Format$(Now, "YYYYMMDD"):
'                                    strTime = Format$(Now, "MMSS")
'                                    .SetText intCol, pGrid_Point, mstrRstval
'                                    .Col = intCol:  .Row = pGrid_Point
'                                    spdResult1.TypeHAlign = TypeHAlignLeft
'                                    spdResult1.TypeVAlign = TypeVAlignCenter
'
'                                    If Len(Trim(itemX.SubItems(8))) <> 0 And Len(Trim(itemX.SubItems(9))) <> 0 Then
'                                        If Val(strRstval) < Val(itemX.SubItems(8)) Then
'                                            strRefVal = "L"
'                                        ElseIf Val(strRstval) > Val(itemX.SubItems(9)) Then
'                                            strRefVal = "H"
'                                        Else
'                                            strRefVal = ""
'                                        End If
'                                        .Col = intCol:  .Row = pGrid_Point
'                                        If strRefVal = "H" Then
'                                              .ForeColor = IIf(Trim$(strRefVal) = "H", vbRed, vbBlack)
'                                        Else
'                                              .ForeColor = IIf(Trim$(strRefVal) = "L", vbBlue, vbBlack)
'                                        End If
'                                    End If
'
'                                    If pGrid_Point <> 0 Then
'                                      sqlDoc = "Update INTERFACE003" & _
'                                               "   set RSTVAL  = '" & mstrRstval & "', REFVAL = '" & strRefVal & "'" & _
'                                               " where SPCNO   = '" & strBarno & "'" & _
'                                               "   and EQPNUM  = '" & itemX.tag & "'" & _
'                                               "   and TRANSDT = '" & strDate & "'" & _
'                                               "   and TRANSTM = '" & strTime & "'"
'                                      AdoCn_Jet.Execute sqlDoc
'
'                                      If cboChk.ListIndex = 0 Then
'                                          sqlDoc = "insert into INTERFACE003(" & _
'                                                   "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'                                                   "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
'                                                   "            '" & strDate1 & "', '" & strTime & "'," & _
'                                                   "            '" & mstrRstval & "', '" & strRefVal & "'," & _
'                                                   "            '" & INS_CODE & "', '', '" & PName & "', '" & pNo & "')"
'                                      Else
'                                          sqlDoc = "insert into INTERFACE003(" & _
'                                                   "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'                                                   "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
'                                                   "            '" & strDate1 & "', '" & strTime & "'," & _
'                                                   "            '" & mstrRstval & "', '" & strRefVal & "'," & _
'                                                   "            '" & INS_CODE & "', '', '" & PName & "', '" & pNo & "')"
'                                      End If
'                                      Debug.Print sqlDoc
'                                      AdoCn_Jet.Execute sqlDoc
'                                    End If
'
'                                End If
'                                Exit For
'                            End If
'                        Next intIdx
'                    End If
'                    Set itemX = Nothing
'                Next
'            End If
'        End With
'    ElseIf Mid$(fSens(1), 3, 1) = "L" Then
'        Flag_HQL = Flag_HQL & Mid(strdata, 3, 1)
'
'        sCol = 8
'
'        pGrid_Point = SeqSearch(spdResult1, sPatiant_No, sCol)
'
'        If pGrid_Point <> 0 Then
'            spdResult1.Row = pGrid_Point
'            spdResult1.Col = 8: spdResult1.BackColor = vbCyan
'            spdResult1.Col = 8: spdResult1.Text = "검사"
'            spdResult1.Col = 1: spdResult1.Text = "1"
'        End If
'
'
'
'        '
'        ' 저장..
'        '
''        If chkAuto.Value = "1" Then
''            Call cmdAppend_Click(0)
''        End If
'
'    End If
'
    Exit Sub
    
errDefine:
    
    Call ErrMsgProc(CallForm)

End Sub

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    Dim strEqpCd As String
    Dim strOrderMsg As String
    Dim itemX   As ListItem
    Dim pGrid_Point As Integer
    Dim varTmp
    Dim strBarno As String
    Dim PName As String
    Dim pNo As String
    Dim pChart As String
    Dim intCol0 As Integer
    Dim intCol As Integer
    Dim strRstval As String
    Dim intIdx As Integer
    Dim TestId As String
    Dim Channel_No As String
    Dim strRefVal As String
    Dim strDate As String
    Dim strTime As String
    Dim sqlDoc As String
    Dim sSeq As String
    Dim intCnt As Integer
    Dim intOrdCnt As Integer
    Dim strTmpBar1 As String
    Dim sCol As Integer
    Dim strTmpDate As String
    
    Dim strTmpBar As String
    
    Dim sqlRet   As Integer
    Dim stryy, strmm, strdd As String
    Dim mResult() As String
    Dim mIcount As Integer
    Dim sPosition As Integer
    Dim strRstHL  As String
    
    Const iTemresultLen = "5"
    
    On Error Resume Next
       
    CallForm = "frmInterface - Private Sub psDataDefine()"
    
    On Error GoTo ErrReceive
                ReceiveData = strdata
                
'<p><n>RBC</n><v>4.23</v><l>3.50</l><h>5.50</h></p>
'<p><n>MCV</n><v>88.5</v><l>75.0</l><h>100.0</h></p>
'<p><n>HCT</n><v>37.5</v><l>35.0</l><h>55.0</h></p>
'<p><n>MCH</n><v>28.1</v><l>25.0</l><h>35.0</h></p>
'<p><n>MCHC</n><v>31.8</v><l>31.0</l><h>38.0</h></p>
'<p><n>RDWR</n><v>13.6</v><l>11.0</l><h>16.0</h></p>
'<p><n>RDWA</n><v>63.0</v><l>30.0</l><h>150.0</h></p>
'<p><n>PLT</n><v>250</v><l>100</l><h>400</h></p>
'<p><n>MPV</n><v>7.0</v><l>8.0</l><h>11.0</h></p>
'<p><n>PCT</n><v>0.17</v><l>0.01</l><h>9.99</h></p>
'<p><n>PDW</n><v>8.8</v><l>0.1</l><h>99.9</h></p>
'<p><n>LPCR</n><v>9.6</v><l>0.1</l><h>99.9</h></p>
'<p><n>HGB</n><v>11.9</v><l>11.5</l><h>16.5</h></p>
'<p><n>WBC</n><v>9.3</v><l>3.5</l><h>10.0</h></p>
'<p><n>LA</n><v>2.2</v><l>0.5</l><h>5.0</h></p>
'<p><n>MA</n><v>0.5</v><l>0.1</l><h>1.5</h></p>
'<p><n>GA</n><v>6.6</v><l>1.2</l><h>8.0</h></p>
'<p><n>LR</n><v>24.2</v><l>15.0</l><h>50.0</h></p>
'<p><n>MR</n><v>4.1</v><l>2.0</l><h>15.0</h></p>
'<p><n>GR</n><v>71.7</v><l>35.0</l><h>80.0</h></p>
'
                
                
                Dim tResult
                mResult = Split(strdata, vbCrLf)
                
                tResult = Split(mResult(14), "|"): ALFA.SID = tResult(4)
                tResult = Split(mResult(14), "|"): ALFA.SampleNo = tResult(1)
                
                List1.AddItem ("▒ ALFA.Sample Position Number : " & Val(ALFA.SID))
               
                strTmpDate = Format(Now, "YYYY")
                strTmpBar = strTmpDate & Mid(Trim(ALFA.SID), 1, 4) & "-" & Mid(Trim(ALFA.SID), 5, 4) & "-" & Mid(Trim(ALFA.SID), 9, 2)

                 tResult = Split(mResult(63), "|"): ALFA.TestId(1) = tResult(2):        ALFA.Result(1) = tResult(4)  ' WBC
                 
                 tResult = Split(mResult(64), "|"): ALFA.TestId(2) = tResult(2):        ALFA.Result(2) = tResult(4)  ' Lymph#
                 tResult = Split(mResult(65), "|"): ALFA.TestId(3) = tResult(2):        ALFA.Result(3) = tResult(4)  ' MID#
                 tResult = Split(mResult(66), "|"): ALFA.TestId(4) = tResult(2):        ALFA.Result(4) = tResult(4)  ' Gran#
                 tResult = Split(mResult(67), "|"): ALFA.TestId(5) = tResult(2):       ALFA.Result(5) = tResult(4)  ' Lymph%
                 tResult = Split(mResult(68), "|"): ALFA.TestId(6) = tResult(2):       ALFA.Result(6) = tResult(4)  ' MID%
                 tResult = Split(mResult(69), "|"): ALFA.TestId(7) = tResult(2):       ALFA.Result(7) = tResult(4)  ' Gran%
                 tResult = Split(mResult(70), "|"): ALFA.TestId(8) = tResult(2):        ALFA.Result(8) = tResult(4)  ' RBC
                 tResult = Split(mResult(71), "|"): ALFA.TestId(9) = tResult(2):        ALFA.Result(9) = tResult(4)  ' HGB
                tResult = Split(mResult(72), "|"): ALFA.TestId(10) = tResult(2):      ALFA.Result(10) = tResult(4)  ' MCHC
                tResult = Split(mResult(73), "|"): ALFA.TestId(11) = tResult(2):      ALFA.Result(11) = tResult(4)  ' MCV
                tResult = Split(mResult(74), "|"): ALFA.TestId(12) = tResult(2):      ALFA.Result(12) = tResult(4)  ' MCH
                tResult = Split(mResult(75), "|"): ALFA.TestId(13) = tResult(2):      ALFA.Result(13) = tResult(4)  ' RDW
                tResult = Split(mResult(76), "|"): ALFA.TestId(14) = tResult(2):      ALFA.Result(14) = tResult(4)  ' HCT
'                tResult = Split(mResult(77), "|"): ALFA.TestId(15) = tResult(2):      ALFA.Result(15) = tResult(4)  ' PLT
'                tResult = Split(mResult(78), "|"): ALFA.TestId(16) = tResult(2):      ALFA.Result(16) = tResult(4)  ' MPV
'                tResult = Split(mResult(79), "|"): ALFA.TestId(17) = tResult(2):      ALFA.Result(17) = tResult(4)  ' PDW
                tResult = Split(mResult(80), "|"): ALFA.TestId(18) = tResult(2):      ALFA.Result(18) = tResult(4)  ' PCT
                tResult = Split(mResult(81), "|"): ALFA.TestId(19) = tResult(2):      ALFA.Result(19) = tResult(4)  ' RDW
                tResult = Split(mResult(82), "|"): ALFA.TestId(20) = tResult(2):      ALFA.Result(20) = tResult(4)  ' RDW
                
                If Len(ALFA.SID) > 0 Then
                
                    With spdResult1
                    
                        Dim strDate2 As String
                        
                        pGrid_Point = Position_Search(spdResult1, ALFA.SID, 6)
                        
                        .GetText 2, pGrid_Point, varTmp:   strDate2 = Trim$(varTmp)
                        .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
                        .GetText 4, pGrid_Point, varTmp:   PName = Trim$(varTmp)
                        .GetText 6, pGrid_Point, varTmp:   pChart = Trim$(varTmp)
                        .GetText 7, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
                        
                        
                        List1.AddItem ("▒ " & pChart & " | " & PName)
                        List1.AddItem ("--------------------------------------------------------------------------------")
                        DoEvents
    
                        If pGrid_Point > 0 Then
                            For intCol = 8 To .MaxCols
                                strRstval = ""
                                .GetText intCol, 0, varTmp
                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                If Not itemX Is Nothing Then
                                   
                                    For intIdx = 1 To .MaxCols
                                    
                                        If Trim(ALFA.TestId(intIdx)) = itemX.tag Then
                                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                            If Not itemX Is Nothing Then
                                                
                                                
                                                strBarno = Mid(strDate2, 1, 4) & Mid(strDate2, 6, 2) & Mid(strDate2, 9, 2)
                                               
                                                strRstval = ALFA.Result(intIdx)
                                                strRefVal = ""
                                                
                                                strDate = Format$(Now, "YYYYMMDD")
                                                strTime = Format$(Now, "HHMMSS")
                                                
                                                strRstHL = Right(Trim(strRstval), 1)
                                                
                                                Select Case strRstHL
                                                    Case "H", "h"
                                                        strRstval = Replace(strRstval, "H", "")
                                                        .Col = intCol: .ForeColor = vbRed
                                                    Case "L", "l"
                                                        strRstval = Replace(strRstval, "L", "")
                                                        .Col = intCol: .ForeColor = vbBlue
                                                    Case Else
                                                End Select
                                                
                                                .SetText intCol, pGrid_Point, Trim(strRstval)

                                                .Col = 2: .BackColor = vbYellow
                                                .Col = 3: .BackColor = vbYellow
                                                .Col = 4: .BackColor = vbYellow
                                                .Col = 5: .BackColor = vbYellow
                                                .Col = 6: .BackColor = vbYellow
                                                .SetText 1, pGrid_Point, "1"
                                                                                                
                                                If Len(itemX.tag) <> 0 Then
                                                
                                                    sqlDoc = "Update INTERFACE003" & _
                                                             "   Set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                                             " where SPCNO   = '" & pChart & "'" & _
                                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                             "   and TRANSDT = '" & strDate & "'" & _
                                                             "   and TRANSTM = '" & strTime & "'"

                                                    AdoCn_Jet.Execute sqlDoc, sqlRet

                                                    '
                                                    ' Update가 안 되었을시.. 신규 였을때..
                                                    '
                                                    If sqlRet = 0 Then
                                                        sqlDoc = "Insert Into INTERFACE003(" & _
                                                                 "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                                 "    Values( '" & pChart & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                                                 "            '" & strDate & "', '" & strTime & "'," & _
                                                                 "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                                 "            '" & INS_CODE & "', '', '" & PName & "', '" & pChart & "')"
    
                                                        AdoCn_Jet.Execute sqlDoc, sqlRet
                                                    End If
                                                End If
                                            Exit For
                                            Set itemX = Nothing
                                        End If
                                        End If
                                    Next intIdx
                                End If
                            Next
                        End If
                    End With
                End If

    
    Exit Sub
    
ErrReceive:

    Set AdoRs_SQL = Nothing

    Set itemX = Nothing

    Call ErrMsgProc(CallForm)

End Sub

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

Private Function f_funAdd_Server(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
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

Private Function f_funAdd_QcServer(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_QcServer = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
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

Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Trim(.Text) = brSeq Then
                    SeqSearch = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        Else
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Trim(.Text) = Trim(brSeq) Then
                    SeqSearch = sCnt 'brSeq
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        End If
    End With

End Function

'
' Position 찾기
'
Private Function Position_Search(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long
    
    '
    ' 초기화
    '
    Position_Search = 0
    
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Val(.Text) = brSeq Then
                    Position_Search = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        Else
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Val(.Text) = Val(brSeq) Then
                    Position_Search = sCnt 'brSeq
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        End If
    End With

End Function


Private Sub Command1_Click()
   
    Dim Arr()   As Byte
    Dim strTmp  As String
   
''    strTmp = strTmp & "D1U0905280000000000001000000S004400460001240039000848202700031800311003870010600507000250000700033004070011500094002170"
'
'    strTmp = strTmp & "D1U1003020000000000007000000S006700486001340043500895002760030820175000890009900812000060000700054004730011900099002380"
'    strTmp = strTmp & "D1U1003020000000000008000000S009500432001190038600894002750030820355006801010400216200650001000020004320010400090001740"
'    strTmp = strTmp & "D1U1003020000000000009000000S007700464001260038900838202720032400256003010006800631000230000500049004190008420080201102"

'    strTmp = strTmp & "D1U1004230000000000000000000S00010000000000000000*0000*0000*000000000*0000*0000*0000*0000*0000*0000*0000*0000*0000*0000"
' strTmp = strTmp & ""
' strTmp = strTmp & "AAAI10P1900000000001800529201017400055002300020030411044545376108032408860287134333019507715315004360000000000000000000001106107725502718301623600000000000000000000000000000000000000000000000000000000010020050080160290480730991311551822042262412502552532522482452352272202122021991971891851761641571511431331261191141101091020970890830770720670580540500490480460460450470480480450430440430410380380380420430420430430460460470470470470470470460450450480510540570610630670710750740770820860910940960960991011041061061061111151211261251261321381401401411401411381381321341301291301291331331361331341281241211151131071010970970950930910820790800780730730710680670670600560540480460420400370340320290280240240220200190190170160140130120130120110100100100100100080080080080080070060050040040030030020010010010010010010020020020020010020010010000010010010010010010010010010010010010010010010010010010010010010010010010010010010010010010000000000000000000000000"
' strTmp = strTmp & "000000000000000000000000000000000000000000000000000010010020020030030040040040040040040050050040040040040040040040050060080100130170210260320390490580"
' strTmp = strTmp & "680810961111261431581751922062202302412482532542552522502462382332262192102031941861801681611551481411351291221201141101071030980940920860840790740720700670640620590570550510480460440410380350320300280260240230220210210200190190180170180170170170170160150140140130130120110110110110100110110100100100090090090080080070070070070070070060060060050050040040040040040030030030030030020020020020020020020020020020020020020020020020010010010010010010010010010010010010010010010010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010010010010020030040050070110150200250310380460550630710800901001101211321441561671781871972042122182252302352402452482512522542542552532522512502492482462442412392362332292252202152102062011971921871821781721671611551501461411371341311291271241221191171151131111091061041021010"
' strTmp = strTmp & "99097094091088086083080077074072070068066063061059057055054052051050049048047045044043042041040038037036036035034033032031030029029028027026026026026025025024024023023022022022022022022021021021021021021020019018018017017017017016016016017016016016016015015015015015015015015015015014014013013012012012012011011011011010010009009009009009009009009009009009009009009009009009009008008007007006006006006005005005005005005005005005005005005005005005005005"


'    strTmp = strTmp & ""
'
'strTmp = strTmp & "<!--:Begin:Chksum:1:-->" & vbCrLf
'strTmp = strTmp & "<!--:Begin:Msg:5:0:-->" & vbCrLf
'strTmp = strTmp & "<sample>" & vbCrLf
'strTmp = strTmp & "<ver>1.1</ver>" & vbCrLf
'strTmp = strTmp & "<instrinfo>" & vbCrLf
'strTmp = strTmp & "<p><n>PRDI</n><v>BM800</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>FIWV</n><v>2.4.4hw</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>SNO</n><v>10668</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>BRND</n><v>S</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>IAPL</n><v>H</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>IID</n></p>" & vbCrLf
'strTmp = strTmp & "</instrinfo>" & vbCrLf
'strTmp = strTmp & "<smpinfo>" & vbCrLf
'strTmp = strTmp & "<p><n>ID</n><v>1489</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>SEQ</n><v>43</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>DATE</n><v>2014-03-14T03:45:46</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>APNU</n><v>1</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>APNA</n><v>BLOOD</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>ASPM</n><v>OT</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>ASPS</n><v>1</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>SORC</n><v>0</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>BLMD</n><v>0</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>BLNK</n><v>0</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>STYP</n><v>0</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RGED</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RGEL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RGEC</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RDLI</n><v>1308-226</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RDPN</n><v>830</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RDED</n><v>2016-08-11</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RLLI</n><v>1301-166</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RLPN</n><v>1375</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RLED</n><v>2016-01-21</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RCLI</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RCPN</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RCED</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RPD</n><v>27</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RPDS</n><v>1</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RPDL</n><v>15</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RPDH</n><v>30</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RPDF</n><v>27</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDDM</n><v>0</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDDP</n><v>45</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDMS</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDMA</n><v>2</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDFB</n><v>0</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDLL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDLH</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDCL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDCH</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WLGL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WLGH</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDIL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDIH</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WDOM</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>XLT</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>CAPL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>CLVL</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>CEXP</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>CEXT</n></p>" & vbCrLf
'strTmp = strTmp & "<p><n>ASWP</n></p>" & vbCrLf
'strTmp = strTmp & "</smpinfo>" & vbCrLf
'strTmp = strTmp & "<smpresults>" & vbCrLf
'strTmp = strTmp & "<p><n>RBC</n><v>4.23</v><l>3.50</l><h>5.50</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>MCV</n><v>88.5</v><l>75.0</l><h>100.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>HCT</n><v>37.5</v><l>35.0</l><h>55.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>MCH</n><v>28.1</v><l>25.0</l><h>35.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>MCHC</n><v>31.8</v><l>31.0</l><h>38.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RDWR</n><v>13.6</v><l>11.0</l><h>16.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>RDWA</n><v>63.0</v><l>30.0</l><h>150.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>PLT</n><v>250</v><l>100</l><h>400</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>MPV</n><v>7.0</v><l>8.0</l><h>11.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>PCT</n><v>0.17</v><l>0.01</l><h>9.99</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>PDW</n><v>8.8</v><l>0.1</l><h>99.9</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>LPCR</n><v>9.6</v><l>0.1</l><h>99.9</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>HGB</n><v>11.9</v><l>11.5</l><h>16.5</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WBC</n><v>9.3</v><l>3.5</l><h>10.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>LA</n><v>2.2</v><l>0.5</l><h>5.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>MA</n><v>0.5</v><l>0.1</l><h>1.5</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>GA</n><v>6.6</v><l>1.2</l><h>8.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>LR</n><v>24.2</v><l>15.0</l><h>50.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>MR</n><v>4.1</v><l>2.0</l><h>15.0</h></p>" & vbCrLf
'strTmp = strTmp & "<p><n>GR</n><v>71.7</v><l>35.0</l><h>80.0</h></p>" & vbCrLf
'strTmp = strTmp & "</smpresults>" & vbCrLf
'strTmp = strTmp & "<tparams>" & vbCrLf
'strTmp = strTmp & "<p><n>RCT</n><v>13060</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>WCT</n><v>9750</v></p>" & vbCrLf
'strTmp = strTmp & "<p><n>aspt</n><v>1082</v></p>" & vbCrLf
'strTmp = strTmp & "</tparams>" & vbCrLf
'strTmp = strTmp & "</sample>" & vbCrLf
'strTmp = strTmp & "<!--:End:Msg:5:0:-->" & vbCrLf
'strTmp = strTmp & "<!--:End:Chksum:1:158:230:-->" & vbCrLf



'
'    strTmp = strTmp & ""
'    strTmp = strTmp & "MEK-6400" & vbCrLf
'    strTmp = strTmp & "18" & vbCrLf
'    strTmp = strTmp & "1024" & vbCrLf
'    strTmp = strTmp & "VENOUS" & vbCrLf
'    strTmp = strTmp & "CBC" & vbCrLf
'    strTmp = strTmp & "1" & vbCrLf
'    strTmp = strTmp & "BLOOD" & vbCrLf
'    strTmp = strTmp & "MMM" & vbCrLf
'    strTmp = strTmp & "2301" & vbCrLf
'    strTmp = strTmp & "                                          " & vbCrLf
'    strTmp = strTmp & "2010" & vbCrLf
'    strTmp = strTmp & "8" & vbCrLf
'    strTmp = strTmp & "27" & vbCrLf
'    strTmp = strTmp & "     " & vbCrLf
'    strTmp = strTmp & "19" & vbCrLf
'    strTmp = strTmp & "7" & vbCrLf
'    strTmp = strTmp & "46" & vbCrLf
'    strTmp = strTmp & "37547     020" & vbCrLf
'    strTmp = strTmp & "7.3H" & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "4.47L" & vbCrLf
'    strTmp = strTmp & "12.8" & vbCrLf
'    strTmp = strTmp & "38.5" & vbCrLf
'    strTmp = strTmp & "86.1" & vbCrLf
'    strTmp = strTmp & "28.6" & vbCrLf
'    strTmp = strTmp & "33.2" & vbCrLf
'    strTmp = strTmp & "19.7H" & vbCrLf
'    strTmp = strTmp & "224" & vbCrLf
'    strTmp = strTmp & "0.09L" & vbCrLf
'    strTmp = strTmp & "4.1L" & vbCrLf
'    strTmp = strTmp & "16.3" & vbCrLf
'    strTmp = strTmp & "22.1" & vbCrLf
'    strTmp = strTmp & "2.8" & vbCrLf
'    strTmp = strTmp & "75.1" & vbCrLf
'    strTmp = strTmp & "1.6" & vbCrLf
'    strTmp = strTmp & "0.2" & vbCrLf
'    strTmp = strTmp & "5.5" & vbCrLf
'    strTmp = strTmp & "      " & vbCrLf
'    strTmp = strTmp & "                                                                                                                                                                " & vbCrLf
'    strTmp = strTmp & "                                                                                       " & vbCrLf
'    strTmp = strTmp & "                                                                                                                                                                                                                                                               " & vbCrLf
'    strTmp = strTmp & ""

    strTmp = "            984TS017 GOT18.0690 GPT11.3320 CRE0.96300 BUN13.9080 LDH187.491 AMY99.5860 U-A3.69000 T-B0.81000 ALP118.827 GGT15.4870HD-D46.4830 GLU88.3540 CHO201.816 PRO7.21400 T-G177.315 ALB4.40300 CRP0.14600180" & EOT
    spdResult1.SetText 6, 1, "984"
    Call ComReceive(strTmp)

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
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun                 ' 실행
    
'    mskRstDate.text = Format$(Now, "YYYYMMDD")
'    mskOrdDate.text = Format$(Now - 1, "YYYYMMDD")
'    mskOrdDate1.text = Format$(Now, "YYYYMMDD")
    mskOrdtime.Text = Format$(Now, "HHMM")
    
    dtpRsltDay.Value = Now
    dtpStartDt.Value = Now
    dtpStopDt.Value = Now
    mskOrdtime.Text = Format$(Now, "HHMM")
    
    Open App.Path + "\" + REG_INSNAME + ".log" For Append As #100

    Print #100, Chr(13) + Chr(10);
    
    Open App.Path + "\ErrorLog\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".sql" For Append As #200

    Print #200, Chr(13) + Chr(10);
   
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    tabWork.Tab = 0
    Or_Seq = 1
    intRow = 0
    chkEnq = 0
    cboChk.ListIndex = 1

'    '-- 2010.03.11 osw 추가 : 검사결과 팝업메세지
'    Set mobjPopups = New PopUpMessages
'    With mobjPopups
'       ' .XPos = Screen.Width / 2
'       ' .YPos = 0
'       ' .PopUpDirection = vbPopDown
'        .ShowDelay = 3000
'        .MovementIndex = 5
'        .ScrollDelay = 30
'
'    End With
'
'    SetupDefaultPopup
    
    COM_MODE = "1"
    
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
    
    Close #100
    Close #200
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

Private Sub imgReceive_DblClick()

    If pnlCom2.Visible = True Then
        pnlCom2.Visible = False
    Else
        pnlCom2.Visible = True
        pnlCom2.ZOrder 0
    End If
    
End Sub

Private Sub imgSend_DblClick()
    
    If pnlCom.Visible = True Then
        pnlCom.Visible = False
    Else
        pnlCom.Visible = True
        pnlCom.ZOrder 0
    
        tmrWorking.Interval = 20000
        tmrWorking.Enabled = True
    End If

End Sub

Private Sub Label6_DblClick()
    If Command1.Visible = False Then
        Command1.Visible = True
    Else
        Command1.Visible = False
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

'Private Sub mskOrdDate_GotFocus()
'
'    With mskOrdDate
'        .SelStart = 8
'        .SelLength = Len(.text)
'    End With
'
'End Sub
'
'
'Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)
'
'    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
'
'End Sub


'Private Sub mskRstDate_GotFocus()
'
'    With mskRstDate
'        .SelStart = 0
'        .SelLength = Len(.text) + 2
'    End With '
'
'End Sub
'
'
'Private Sub mskRstDate_KeyPress(KeyAscii As Integer)
'
'    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
'
'End Sub

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
            Call COM_OUTPUT(Order.MSG_ENQ)
        Case 2
            Call COM_OUTPUT(Order.MSG_HEADER)
        Case 3
            Call COM_OUTPUT(Order.MSG_PATIENT)
        Case 4
            Call COM_OUTPUT(Order.MSG_ORDER)
        Case 5
            Call COM_OUTPUT(Order.MSG_TERMINATION)
        Case 6
            Call COM_OUTPUT(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

Private Sub spdRstview_Click(ByVal Col As Long, ByVal Row As Long)
Dim iCnt, rCnt As Integer
Dim intCol, intRow As Integer
Dim tCol As Integer
Dim iresult As String
'
' 결과 시작 Position
'
Const sResultPos As Integer = 8
    With spdRstview
        For iCnt = 2 To .MaxCols Step 2
            For rCnt = 1 To .maxrows
                .Row = rCnt: .Col = iCnt
                iresult = Trim(.Text)
                
                With spdResult1
                    .Row = gspdResultRow:  .Col = sResultPos + tCol
                    If Len(Trim(iresult)) <> 0 Then
                        .Text = iresult
                    End If
                    DoEvents
                End With
                tCol = tCol + 1
                
            Next rCnt
            rCnt = 0
        Next iCnt
    End With
End Sub

Private Sub spdRstview_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
    Call spdRstview_Click(Row, RowIsLast)
End Sub

'
'
'

Private Sub spdRstview_KeyPress(KeyAscii As Integer)

Dim iCnt, rCnt As Integer
Dim intCol, intRow As Integer
Dim tCol As Integer
Dim iresult As String

'
' 결과 시작 Position
'
Const sResultPos As Integer = 8
     
    ' 처방 존재 유무 확인..
    With spdRstview
        .Row = .ActiveRow: .Col = .ActiveCol
        If .BackColor <> &HC6FEFF And Len(.Text) >= 1 Then
            .Text = ""
            MsgBox "▒ OCS/EMR의 검사 처방이 없는 항목 입니다.." & Space(5), vbOKOnly + vbInformation, App.Title
            spdRstview.SetFocus
            Exit Sub
        End If
    End With
    
    ' Enter Key 유무..
    If KeyAscii = vbKeyReturn Then
    
        If gspdResultRow < 1 Then
            With spdRstview
                .Row = .ActiveRow:  .Col = .ActiveCol
                .Text = ""
            End With
            
            MsgBox "▒ 수정을 원하는 검사 Sample을 선택 후 수정 하십시요.." & Space(5), vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        
        ' 수정된 결과 본 Spread로 옮기기..
        With spdRstview
            For iCnt = 2 To .MaxCols Step 2
                For rCnt = 1 To .maxrows
                    .Row = rCnt: .Col = iCnt
                    iresult = .Text
                    
                    With spdResult1
                        .Row = gspdResultRow:  .Col = sResultPos + tCol
                        If Len(Trim(iresult)) <> 0 Then
                            .Text = iresult
                        End If
                    End With
                    tCol = tCol + 1
                Next rCnt
            Next iCnt
        End With
    End If

End Sub


Private Sub spdRstview_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim objResult As clsResult
   Dim lngCol As Long
   
   If gspdResultRow = 0 Then Exit Sub
   
   If 2280 >= X And X >= 1410 Then
      lngCol = 2
   ElseIf 4125 >= X And X >= 3210 Then
      lngCol = 4
   ElseIf 5055 >= X And X >= 5955 Then
      lngCol = 8
   ElseIf 6885 >= X And X >= 7755 Then
      lngCol = 8
   Else
      lngCol = 9
   End If

   If y < 330 Then Exit Sub

   Select Case lngCol
      Case 2, 4, 6, 8
        spdRstview_TextTipFetch lngCol, gspdResultRow, 1, 6500, "", True
      Case Else
        Exit Sub
   End Select
   
End Sub


Private Sub spdRstview_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    Dim pDate, pPtnm, pPtno, pSex, pPos As String
    
    With spdResult1
        .Row = gspdResultRow
        .Col = 2: pDate = .Text
        .Col = 4: pPtnm = .Text
        .Col = 5: pSex = .Text
        .Col = 6: pPtno = .Text
        .Col = 7: pPos = .Text
    End With
            
    Debug.Print pDate, pPtnm, pPtno, pSex, pPos
            
    With spdRstview
        .Row = Row
        .Col = Col
         MultiLine = 1
         TipWidth = 3000
         .SetTextTipAppearance "굴림체", 9, False, False, &HEEFDF2, vbBlack
         .TextTip = TextTipFloating
         
    
         .SetTextTipAppearance "굴림체", 9, False, False, &HEEFDF2, vbBlue
         
         TipText = "" & vbNewLine & _
                   "   ▒ 처방일자 ; " & pDate & vbNewLine & _
                   "   ▒ 환 자 명 ; " & pPtnm & vbNewLine & _
                   "   ▒ 병록번호 ; " & pPtno & vbNewLine & _
                   "   ▒ 성    별 ; " & pSex & vbNewLine & vbNewLine & _
                   "   ▒ 검사 POS ; " & pPos & vbNewLine
                   
         ShowTip = True
       
    End With
End Sub


Private Sub spdRstview_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 이전 환자", "-", " ▒ 다음 환자")

    Select Case lMenuChosen
        Case 1
            With spdResult1
                Col = .ActiveCol
                Row = .ActiveRow
            End With
            
            If gspdResultRow >= 1 Then
                Call spdResult1_Click(Col, gspdResultRow - 1)
            ElseIf gspdResultRow = 0 Then
                MsgBox "▒ 처음 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
            Else: Exit Sub
            End If
            
        Case 3
            With spdResult1
                Col = .ActiveCol
                Row = .ActiveRow
            End With
            
            If gspdResultRow < spdResult1.maxrows Then
                Call spdResult1_Click(Col, gspdResultRow + 1)
            ElseIf gspdResultRow = spdResult1.maxrows Then
                MsgBox "▒ 마지막 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
            Else: Exit Sub
    
            End If
    
    End Select


End Sub

Private Sub cmdNext_Click()
Dim Col, Row As Integer
    
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow < spdResult1.maxrows Then
        Call spdResult1_Click(Col, gspdResultRow + 1)
    ElseIf gspdResultRow = spdResult1.maxrows Then
        MsgBox "▒ 마지막 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
    Else: Exit Sub
    
    End If
    
End Sub

Private Sub cmdPrevious_Click()
Dim Col, Row As Integer
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow >= 1 Then
        Call spdResult1_Click(Col, gspdResultRow - 1)
    ElseIf gspdResultRow = 0 Then
        MsgBox "▒ 처음 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
    Else: Exit Sub
    End If
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then
        gspdResultRow = 0:        Exit Sub
    Else
        gspdResultRow = Row
    End If
    
    intCol1 = 10
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = intRow1
            spdRstview.Col = intCol2
            spdRstview.BackColor = vbWhite
            
            If .BackColor = &HC6FEFF Then
                spdRstview.BackColor = &HC6FEFF
            Else
                spdRstview.BackColor = &H80000005
            End If
            
            spdRstview.Text = .Text
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstview.maxrows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If

        Next
    End With
    

End Sub

'
' END
'

Private Sub spdResult1_KeyPress(KeyAscii As Integer)

    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    'Debug.Print Col & NewCol & Row & NewRow
       
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            aCOL = .ActiveCol
            aROW = .ActiveRow
            If aCOL = 4 Then
                iCnt = 0
                For iRow = aROW To .maxrows
                    .GetText 1, iRow, varChk
                    .GetText 3, iRow, varBar
                    .GetText aCOL, aROW, varNum
                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
                        .SetText aCOL, iRow, varNum
                        .SetText aCOL + 1, iRow, ((iCnt Mod 40) + 1) + (40 * (varNum - 1))
                        iCnt = iCnt + 1
                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
                    End If
                Next
'            ElseIf aCOL = 5 Then
'                iCnt = 0
'                For iRow = aROW To .maxrows
'                    .GetText 1, iRow, varChk
'                    .GetText 3, iRow, varBar
'                    .GetText aCOL, aROW, varNum
'                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'                        .SetText aCOL, iRow, ((iCnt Mod 40) + varNum) '+ (40 * (varNum - 1))
'                        '.SetText aCOL - 1, iRow, varNum
'                        iCnt = iCnt + 1
'                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
'                    End If
'                Next
            
            End If
        End With
    End If
    
End Sub

Private Sub spdWorklist_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim varTmp  As Variant
    
    If Row = 0 Then
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(spdWorklist, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(spdWorklist, Col, 1)
            OrderSort_Flag = 1
        End If
    Else
        With spdWorklist
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
    
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
            cmdWorkList_Click
        End With
    End If
    
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


Private Sub tabWork_Click(PreviousTab As Integer)
    cboRstgbn(1).ListIndex = 0
'    spdResult2.maxrows = 0
End Sub

Private Sub Timer1_Timer()

    Call COM_OUTPUT(ENQ)
'    Debug.Print ENQ

End Sub

Private Sub Timer2_Timer()
    comEQP.Output = STX
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
    Dim I As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For I = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(I).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - I)) + (70 * (cmdAction.UBound - I)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub

Private Sub tmrWorking_Timer()
    pnlCom.Visible = False
End Sub

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtBarCode_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub txtBarCode_LostFocus()"
    
    Dim varTmp  As Variant, strEqpCd    As String
    Dim intRow  As Integer, intCol  As Integer, blnFlag As Boolean
    Dim strOrdcd() As String, strPid()  As String, strPnm() As String
    Dim strPexzm() As String, strPeqpcd() As String
    Dim strEqcode() As String, strExamname() As String, strAcptno() As String

    Dim itemX   As ListItem
    
    If txtBarCode.Text = "" Then Exit Sub
    
    blnFlag = False
    If KeyAscii = vbKeyReturn Then
        intCol = sl_examdata_select&(txtBarCode.Text, INS_CODE, strEqcode, strExamname, strOrdcd, strPid, strPnm, strAcptno)
        
        For intCol = 0 To UBound(strOrdcd)
            If strOrdcd(intCol) <> "" Then
                strEqpCd = f_funGet_CODE(strOrdcd(intCol))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    If Not blnFlag Then
                        intRow = f_funGet_SpreadRow(spdResult1, 2, txtBarCode.Text)
                        If intRow < 1 Then
                            intRow = f_funGet_SpreadRow(spdResult1, 2, "")
                            If intRow < 1 Then
                                spdResult1.maxrows = spdResult1.maxrows + 1
                                spdResult1.RowHeight(spdResult1.maxrows) = 14
                                intRow = spdWorklist.maxrows
                            End If
                            spdResult1.SetText 2, intRow, txtBarCode.Text
                            spdResult1.SetText 3, intRow, strPnm(0)
                            spdResult1.SetText 4, intRow, strPid(0)
                        End If
                        spdResult1.SetText 1, intRow, "1"
                    End If
                        
                    'spdResult1.SetText itemX.Index + 6, intRow, "V"
                    spdResult1.Col = itemX.Index + 6
                    spdResult1.Row = intRow
                    spdResult1.BackColor = &HC6FEFF
                    
                    blnFlag = True
                End If
            End If
        Next
    
        If Not blnFlag Then MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, App.Title
        
        txtBarCode.Text = "":   txtBarCode.SetFocus
        Exit Sub
    
    End If
    
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Function psDataExists() As Boolean
Dim sCnt As Long
    
    psDataExists = False
    With spdWorklist
        For sCnt = 1 To .maxrows
            .Row = sCnt:    .Col = 2
            If Trim(.Text) = Mid(txtBarCode.Text, 1, 11) Then
                psDataExists = True
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub txtBarCode_LostFocus()

'    Dim intRow      As Integer
'    Dim strOrdcd(1 To 100) As String
'
'    Call sl_spcid_tstcd_select&(txtBarCode.Text, strOrdcd)
'    If strOrdcd(1) = "" Then
'        MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    intRow = f_funGet_SpreadRow(spdWorkList, 2, txtBarCode.Text)
'    If intRow < 1 Then
'        intRow = f_funGet_SpreadRow(spdWorkList, 2, "")
'        If intRow < 1 Then
'            spdWorkList.maxrows = spdWorkList.maxrows + 1
'            spdWorkList.RowHeight(spdWorkList.maxrows) = 13
'            intRow = spdWorkList.maxrows
'        End If
'        spdWorkList.SetText 2, intRow, txtBarCode.Text
'    End If
'    spdWorkList.SetText 1, intRow, "1"
    
End Sub

Private Sub txtChart_GotFocus()
'
' Focus 가졌을 경우
'
    txtChart.ForeColor = &HFF&
    txtChart.Text = ""
End Sub

Private Sub txtChart_LostFocus()
'
' Focus 가 없을 경우
'
    txtChart.ForeColor = &HFFC0C0
    txtChart.Text = "차트번호 입력"
End Sub

Private Sub txtChart_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim intRow2 As Integer
    
    Dim tBlood As Boolean
        
    If Len(Trim(txtChart)) > 0 Then
        If KeyCode = 13 Then
        
          tBlood = False
          
          Rem txtChart = Format(txtChart, "0000000")
          
          intRow2 = f_funGet_SpreadRow(spdWorklist, 6, txtChart)
          
          If intRow2 >= 1 Then
              
              With spdWorklist
                .SetText 1, intRow2, "1"
                cmdWorkList_Click
                txtChart.Text = ""
                tBlood = True
              End With
          End If
          
          If tBlood = False Then
            MsgBox txtChart.Text & " 해당 환자의 처방이 없습니다.     ", vbInformation + vbOKOnly, App.Title
            txtChart.Text = ""
          End If
        
         End If
    End If

End Sub

' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.Text)
End Sub

Private Sub cmdCOMLoad_Click()
    Dim I               As Long
    Dim lngFIleNum      As Long
    Dim strTemp         As String
    Dim strTemp2        As String
    Dim bteBuffer()     As Byte
    
On Error GoTo ErrorRoutine
    
    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowOpen
        lngFIleNum = FreeFile
        
        Open .FileName For Binary Access Read As #lngFIleNum
        
        txtCOM2.Text = ""
        ReDim bteBuffer(LOF(lngFIleNum))
        Get #lngFIleNum, , bteBuffer

        strTemp = StrConv(bteBuffer, vbUnicode)
        txtCOM2.Text = strTemp
                
        Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum
        
End Sub

Private Sub cmdCOMSave_Click()
    Dim lngFIleNum      As Long
    
On Error GoTo ErrorRoutine

    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowSave
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum
        Print #lngFIleNum, _
              Format(Date, "YYYY년 MM월 DD일") & "  "; Time & vbNewLine & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine & _
              txtCom.Text & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine
    Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum

End Sub

Private Sub cmdCOMOutput_Click()
    'Call COM_OUTPUT(StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode))
    Call COM_OUTPUT(charCOM_Convert(txtCom.SelText))
End Sub

Private Sub cmdCOMClear_Click()
    mlngRecLen = 0
    txtCom.Text = ""
End Sub

Private Sub cmdCOMClear2_Click()
    txtCOM2.Text = ""
End Sub

Private Sub cmdCOMInput_Click()

    Dim bytTemp() As Byte
    
    bytTemp = StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode)

    Call ComReceive(txtCom.SelText)
    
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.Text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(txtCOM2.SelText)

End Sub

Private Sub cmdCOMOutput2_Click()
    
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.Text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트


Private Sub txtResult_DblClick()
    txtResult.Text = ""
    List1.Text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    List1.Visible = True
End Sub

