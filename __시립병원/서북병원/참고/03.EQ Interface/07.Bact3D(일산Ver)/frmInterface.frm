VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " BACT-3D Interface "
   ClientHeight    =   10365
   ClientLeft      =   645
   ClientTop       =   675
   ClientWidth     =   16620
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
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   10365
   ScaleWidth      =   16620
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   9930
      Width           =   16620
      _ExtentX        =   29316
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2011-11-24"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 11:49"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
            Text            =   "Service Center (02)6205-1751"
            TextSave        =   "Service Center (02)6205-1751"
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
   Begin VB.Frame Frame7 
      Caption         =   "안쓰는것"
      Height          =   4035
      Left            =   17820
      TabIndex        =   61
      Top             =   5760
      Width           =   2535
      Begin VB.CheckBox chkRAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   225
      End
      Begin VB.CommandButton cmdRClear 
         Caption         =   "화면초기화"
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
         Left            =   1560
         TabIndex        =   72
         Top             =   2640
         Width           =   1395
      End
      Begin VB.CommandButton cmdRTrans 
         Caption         =   "결과수동전송"
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
         TabIndex        =   71
         Top             =   2580
         Width           =   1395
      End
      Begin VB.CommandButton cmdRSch 
         Caption         =   "로컬결과조회"
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
         TabIndex        =   70
         Top             =   3000
         Width           =   1395
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "EXCEL"
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
         TabIndex        =   69
         Top             =   3420
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Frame Frame5 
         Height          =   585
         Left            =   300
         TabIndex        =   62
         Top             =   780
         Width           =   5565
         Begin VB.Label lblRrow 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   67
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Left            =   4155
            TabIndex        =   66
            Top             =   240
            Width           =   1305
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
            Left            =   3105
            TabIndex        =   65
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblBarcode 
            Caption         =   "1234567890ab"
            Height          =   165
            Left            =   1605
            TabIndex        =   64
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label2 
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
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1380
         End
      End
      Begin FPSpread.vaSpread vasRRes 
         Height          =   1350
         Left            =   180
         TabIndex        =   68
         Top             =   1200
         Width           =   4515
         _Version        =   393216
         _ExtentX        =   7964
         _ExtentY        =   2381
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
         MaxCols         =   7
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":058D
      End
      Begin MSComCtl2.DTPicker dtpExamdate 
         Height          =   435
         Left            =   300
         TabIndex        =   91
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40847
      End
   End
   Begin VB.TextBox txtTest 
      Height          =   405
      Left            =   60
      TabIndex        =   56
      Top             =   9600
      Width           =   13485
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   3075
      Left            =   14880
      TabIndex        =   37
      Top             =   5640
      Width           =   3015
      Begin VB.CommandButton cmdLoad 
         Caption         =   "파일 불러오기"
         Height          =   375
         Left            =   1200
         TabIndex        =   41
         Top             =   1020
         Width           =   1725
      End
      Begin VB.CommandButton cmdPatSend 
         Caption         =   "환자 전송"
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
         Left            =   1200
         TabIndex        =   40
         Top             =   600
         Width           =   1185
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 조회"
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
         Left            =   1200
         TabIndex        =   39
         Top             =   180
         Width           =   1725
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   1845
         Left            =   60
         TabIndex        =   38
         Top             =   180
         Width           =   1035
         _Version        =   393216
         _ExtentX        =   1826
         _ExtentY        =   3254
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
         MaxCols         =   7
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":4328
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         Top             =   2100
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40739
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   315
         Left            =   1110
         TabIndex        =   44
         Top             =   2520
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40739
      End
      Begin VB.Label Label3 
         Caption         =   "조회기간 : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   46
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   45
         Top             =   2580
         Width           =   195
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   315
      Left            =   13620
      TabIndex        =   34
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   3495
      Left            =   14880
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   8580
      Begin VB.Timer tmResultCheck 
         Interval        =   10000
         Left            =   120
         Top             =   300
      End
      Begin VB.CommandButton cmdIFTrans 
         Caption         =   "결과수동전송"
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
         Left            =   60
         TabIndex        =   57
         Top             =   960
         Visible         =   0   'False
         Width           =   1395
      End
      Begin FPSpread.vaSpread vasPatList 
         Height          =   780
         Left            =   60
         TabIndex        =   33
         Top             =   2520
         Width           =   1425
         _Version        =   393216
         _ExtentX        =   2514
         _ExtentY        =   1376
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
         SpreadDesigner  =   "frmInterface.frx":80A6
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   1530
         TabIndex        =   26
         Top             =   180
         Width           =   1605
         _Version        =   393216
         _ExtentX        =   2831
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
         SpreadDesigner  =   "frmInterface.frx":82BE
      End
      Begin VB.FileListBox FileBeeBlot 
         Height          =   285
         Left            =   1650
         Pattern         =   "*.txt"
         TabIndex        =   31
         Top             =   2130
         Visible         =   0   'False
         Width           =   2805
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7200
         TabIndex        =   30
         Top             =   315
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
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
         SpreadDesigner  =   "frmInterface.frx":84D6
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   29
         Top             =   1980
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
         Height          =   435
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   23
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   4200
         TabIndex        =   22
         Top             =   2760
         Width           =   1125
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
         Left            =   120
         TabIndex        =   21
         Top             =   735
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   1875
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
         Left            =   3600
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   1320
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   6720
         TabIndex        =   18
         Top             =   1380
         Visible         =   0   'False
         Width           =   1335
         Begin MSCommLib.MSComm MSComm1 
            Left            =   135
            Top             =   270
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            InputLen        =   1
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
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   975
         Left            =   6780
         TabIndex        =   17
         Top             =   240
         Width           =   315
         _Version        =   393216
         _ExtentX        =   556
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
         SpreadDesigner  =   "frmInterface.frx":86EE
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   3195
         TabIndex        =   24
         Top             =   180
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
         SpreadDesigner  =   "frmInterface.frx":8906
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   25
         Top             =   180
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
         SpreadDesigner  =   "frmInterface.frx":8B1E
      End
      Begin FPSpread.vaSpread vasExcelRes 
         Height          =   870
         Left            =   1620
         TabIndex        =   35
         Top             =   2520
         Width           =   2490
         _Version        =   393216
         _ExtentX        =   4392
         _ExtentY        =   1535
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
         SpreadDesigner  =   "frmInterface.frx":8D36
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   690
         Left            =   4860
         TabIndex        =   36
         Top             =   1920
         Width           =   870
         _Version        =   393216
         _ExtentX        =   1535
         _ExtentY        =   1217
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
         SpreadDesigner  =   "frmInterface.frx":8F4E
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   28
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   27
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   2085
      Left            =   14820
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   9465
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   1305
         TabIndex        =   14
         Top             =   270
         Width           =   8160
         _Version        =   393216
         _ExtentX        =   14393
         _ExtentY        =   2725
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
         SpreadDesigner  =   "frmInterface.frx":9166
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1065
         _Version        =   393216
         _ExtentX        =   1879
         _ExtentY        =   2196
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
         SpreadDesigner  =   "frmInterface.frx":ABDF
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   8730
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   15399
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "진행상태"
      TabPicture(0)   =   "frmInterface.frx":ADF7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSPanel61"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ssp6Weeks"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "sspannel1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ssp3Weeks"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSPanel3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "sspPositive"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":AE13
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "cmdExcelSave"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton Command1 
         Caption         =   "통     계"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -62100
         TabIndex        =   90
         Top             =   4680
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton cmdExcelSave 
         Caption         =   "Excel 저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -62100
         TabIndex        =   89
         Top             =   4080
         Width           =   1755
      End
      Begin VB.Frame Frame8 
         Caption         =   "조회"
         Height          =   3615
         Left            =   -62160
         TabIndex        =   74
         Top             =   360
         Width           =   1875
         Begin VB.CheckBox chkNogrow 
            Caption         =   "No grow"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   86
            Top             =   2460
            Width           =   1275
         End
         Begin VB.CheckBox chkPositive 
            Caption         =   "Positive"
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
            Left            =   420
            TabIndex        =   85
            Top             =   2220
            Width           =   1275
         End
         Begin VB.CheckBox chkExam 
            Caption         =   "검사중"
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
            Left            =   120
            TabIndex        =   81
            Top             =   1620
            Width           =   1035
         End
         Begin VB.CheckBox chkResult 
            Caption         =   "검사완료"
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
            Left            =   120
            TabIndex        =   80
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "조회 하기"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   60
            TabIndex        =   79
            Top             =   2940
            Width           =   1755
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   315
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21430275
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   315
            Left            =   120
            TabIndex        =   76
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21430275
            CurrentDate     =   40457
         End
         Begin VB.Label Label11 
            Caption         =   "L"
            Height          =   255
            Left            =   180
            TabIndex        =   88
            Top             =   2520
            Width           =   195
         End
         Begin VB.Label Label10 
            Caption         =   "L"
            Height          =   255
            Left            =   180
            TabIndex        =   87
            Top             =   2160
            Width           =   195
         End
         Begin VB.Label Label9 
            Caption         =   "End."
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label8 
            Caption         =   "Start."
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   8235
         Left            =   -74820
         TabIndex        =   12
         Top             =   360
         Width           =   12600
         Begin VB.CheckBox chkSprSearch 
            BackColor       =   &H0000FF00&
            Caption         =   "Barcode"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  '그래픽
            TabIndex        =   84
            Top             =   180
            Width           =   1215
         End
         Begin VB.CommandButton cmdSprSearch 
            Caption         =   "검색"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3540
            TabIndex        =   83
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox txtSprSearch 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   15
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1500
            MousePointer    =   3  'I-빔
            TabIndex        =   82
            Top             =   180
            Width           =   1935
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7365
            Left            =   180
            TabIndex        =   60
            Top             =   780
            Width           =   12315
            _Version        =   393216
            _ExtentX        =   21722
            _ExtentY        =   12991
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
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
            MaxCols         =   11
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            OperationMode   =   2
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmInterface.frx":AE2F
            UserResize      =   2
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8250
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   12645
         Begin VB.TextBox txtSearch 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   15
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3780
            MousePointer    =   3  'I-빔
            TabIndex        =   94
            Top             =   180
            Width           =   1935
         End
         Begin VB.CommandButton cmdDataSearch 
            Caption         =   "검색"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5820
            TabIndex        =   93
            Top             =   180
            Width           =   1095
         End
         Begin VB.CheckBox chkSprSearch1 
            BackColor       =   &H0000FF00&
            Caption         =   "Barcode"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            Style           =   1  '그래픽
            TabIndex        =   92
            Top             =   180
            Width           =   1215
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   4305
            Left            =   180
            TabIndex        =   59
            Top             =   3840
            Width           =   12375
            _Version        =   393216
            _ExtentX        =   21828
            _ExtentY        =   7594
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
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
            MaxCols         =   11
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            OperationMode   =   2
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmInterface.frx":B7C4
            UserResize      =   2
         End
         Begin VB.CommandButton cmdLocload 
            Caption         =   "불러오기"
            Height          =   375
            Left            =   9780
            TabIndex        =   58
            Top             =   180
            Width           =   1335
         End
         Begin VB.PictureBox Picture1 
            Height          =   75
            Left            =   120
            ScaleHeight     =   15
            ScaleWidth      =   12375
            TabIndex        =   55
            Top             =   3480
            Width           =   12435
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "화면초기화"
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
            Left            =   11160
            TabIndex        =   11
            Top             =   180
            Width           =   1395
         End
         Begin VB.Frame Frame2 
            Caption         =   "Error Log"
            Height          =   1815
            Left            =   540
            TabIndex        =   8
            Top             =   6060
            Visible         =   0   'False
            Width           =   5970
            Begin VB.TextBox txtErrLog 
               Appearance      =   0  '평면
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   9
               Top             =   240
               Width           =   5775
            End
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   10
            Top             =   3960
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   2685
            Left            =   180
            TabIndex        =   42
            Top             =   660
            Width           =   12375
            _Version        =   393216
            _ExtentX        =   21828
            _ExtentY        =   4736
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
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
            MaxCols         =   11
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            OperationMode   =   2
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmInterface.frx":C159
            UserResize      =   2
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   675
            TabIndex        =   32
            Top             =   780
            Width           =   225
         End
         Begin VB.Label Label7 
            Caption         =   "검사 중"
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
            TabIndex        =   48
            Top             =   3600
            Width           =   1005
         End
         Begin VB.Label Label6 
            Caption         =   "검사 완료"
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
            TabIndex        =   47
            Top             =   360
            Width           =   1005
         End
      End
      Begin Threed.SSPanel sspPositive 
         Height          =   1515
         Left            =   12900
         TabIndex        =   49
         Top             =   1140
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2672
         _StockProps     =   15
         Caption         =   "10"
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   30
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   555
         Left            =   12900
         TabIndex        =   50
         Top             =   540
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "Positive"
         ForeColor       =   65535
         BackColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel ssp3Weeks 
         Height          =   1515
         Left            =   12900
         TabIndex        =   51
         Top             =   3720
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2672
         _StockProps     =   15
         Caption         =   "99"
         BackColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   30
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspannel1 
         Height          =   555
         Left            =   12900
         TabIndex        =   52
         Top             =   3120
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "2 Day"
         ForeColor       =   65535
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel ssp6Weeks 
         Height          =   1575
         Left            =   12900
         TabIndex        =   53
         Top             =   6300
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2778
         _StockProps     =   15
         Caption         =   "99"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   30
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel61 
         Height          =   555
         Left            =   12900
         TabIndex        =   54
         Top             =   5700
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "5 Day"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   14700
      _Version        =   65536
      _ExtentX        =   25929
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "     BACT-3D Interface "
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4785
         Picture         =   "frmInterface.frx":CAEE
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   11925
         TabIndex        =   2
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
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
         Format          =   21430272
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   10995
         TabIndex        =   5
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
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
         Height          =   255
         Left            =   5190
         TabIndex        =   4
         Top             =   255
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "메인"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "설정"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
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

'vasid, vasrid colum
'Const colCheckBox = 1
'Const colBarcode = 2
'Const colRack = 3
'Const colPos = 4
'Const colPID = 5
'Const colPName = 6
'Const colSex = 7
'Const colAge = 8
'Const colJumin = 9
'Const colOCnt = 10
'Const colHospital = 11
'Const colState = 12


Const colCheckBox = 1
Const colSpecNo = 2
Const colBarcode = 3
Const colSampleNo = 4
Const colRack = 5
Const colPos = 6
Const colExamCD_MI = 6
Const colPID = 7
Const colBact_Time = 7
Const colPName = 8
Const colBact_result = 8
Const colStartDate = 9
Const colSex = 9
Const colEndDate = 10
Const colAge = 10
Const colOCnt = 11
Const colRCnt = 12
Const colState = 13
Const colExamDate = 14

'Const colA1c = 13
'Const colIFCC = 15
'Const coleAg = 17

'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colResult = 4
Const colSeq = 5
Const colFLAG = 6

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
Dim gsFlag As String

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim gIFCC1 As String
Dim gIFCC2 As String
Dim geAg1 As String
Dim geAg2 As String
Dim gADD_IFCC As String
Dim gADD_eAg As String

Dim strBuffer As String

Public gENQFlag As Integer
Public gNAKFlag As Integer

Public Data_gubun As String '/ 결과를 최종보고 할것인지 중간보고 할것인지 (2: 중간보고 , 3:최종보고  0:보고 하지 않음 )
Public Result_gubun As String '/ 결과를 최종보고 할것인지 중간보고 할것인지 (2: 중간보고 , 3:최종보고)
Dim POS_FLAG    As Integer

'===============================
Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const rs  As String = ""
Const GS  As String = ""

Dim gTimeCnt As Integer '/타이머 떔에 돌림
Dim gDataString         '/mscomm땜에 돌림



Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'/결과 받을때 몇개가 나왔는지 확인
Dim sResult_Flag    As Integer

'Dim mOrder.NoOrder  As Boolean
'Dim mOrder.Order    As String
'Dim mOrder.IsSending As Boolean

'===============================

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.value = 1
        Next iRow
    ElseIf chkAll.value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.value = 0
        Next iRow
    End If
End Sub

Private Sub chkMode_Click()
    If chkMode.value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub


Private Sub chkRAll_Click()
    Dim iRow As Long
    
    If chkRAll.value = 1 Then
        For iRow = 1 To vasRID.DataRowCnt
            vasRID.Row = iRow
            vasRID.Col = 1
            
            vasRID.value = 1
        Next iRow
    ElseIf chkRAll.value = 0 Then
        For iRow = 1 To vasRID.DataRowCnt
            vasRID.Row = iRow
            vasRID.Col = 1
            
            vasRID.value = 0
        Next iRow
    End If
End Sub

Private Sub chkResult_Click()
    If chkResult.value = 0 Then
        chkPositive.value = 0
        chkPositive.Enabled = 0
        chkNogrow.value = 0
        chkNogrow.Enabled = 0
    Else
        chkPositive.Enabled = 1
        chkNogrow.Enabled = 1
    End If
End Sub

Private Sub chkSprSearch_Click()
    If chkSprSearch.Caption = "Barcode" Then
        chkSprSearch.Caption = "환자번호"
        chkSprSearch.BackColor = &HFF8080
    Else
        chkSprSearch.Caption = "Barcode"
        chkSprSearch.BackColor = &HFF00&
    End If
    
End Sub
'
'Private Sub cmdExcel_Click()
'    Dim iRow As Integer
'    Dim j As Integer
'
'    Dim sCurDate As String
'    Dim sSerDate As String
'    Dim sHead As String
'    Dim sFoot As String
'    Dim sFileName As String
'
'    Dim sA1c As String
'    Dim sIFCC As String
'    Dim seAg As String
'
'
'
'    ClearSpread vasPrint
'
'    j = 1
'
'    For iRow = 1 To vasRID.DataRowCnt
'        vasRID.Row = iRow
'        vasRID.Col = 1
'
'        If vasRID.value = 1 Then
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colSpecNo)), j, 1
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colBarcode)), j, 2
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 3
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 4
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 5
'            'SetText vasPrint, Trim(GetText(vasRID, iRow, colHospital)), j, 5
'
'            SQL = "SELECT RESULT " & vbCrLf & _
'                  "FROM PAT_RES " & vbCrLf & _
'                  "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
'                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                  "  AND BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
'                  "  AND PID = '" & Trim(GetText(vasPrint, iRow, 3)) & "' " & vbCrLf & _
'                  "ORDER BY SEQNO"
'            res = db_select_Vas(gLocal, SQL, vasPrintBuf)
'
'            sA1c = GetText(vasPrintBuf, 1, 1)
'            sIFCC = GetText(vasPrintBuf, 2, 1)
'            seAg = GetText(vasPrintBuf, 3, 1)
'
'            ClearSpread vasPrintBuf, 1, 1
'
'            SetText vasPrint, sA1c, j, 7
'            SetText vasPrint, sIFCC, j, 8
'            SetText vasPrint, seAg, j, 9
'
'            '"GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, JUMIN, Hospital, SENDFLAG"
'
''            SetText vasprint, Trim(GetText(vasrid, iRow, vasrid.MaxCols)), j, 8
''            SetText vasprint, Trim(GetText(vasrid, iRow, 10)), j, 9
'
'            j = j + 1
'        End If
'    Next iRow
'
'    If vasPrint.DataRowCnt < 1 Then
'        MsgBox "저장할 자료가 없습니다.", , "알 림"
'        Exit Sub
'    Else
'        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
'        CommonDialog1.ShowSave
'        sFileName = CommonDialog1.FileName
'        SaveExcel sFileName, vasPrint
'
'    End If
'End Sub
Sub SaveExcel(FileName As String, argSpread As vaSpread)

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
    
    xlBook.SaveAs (FileName)
    xlapp.Quit


End Sub

Private Sub chkSprSearch1_Click()
    If chkSprSearch1.Caption = "Barcode" Then
        chkSprSearch1.Caption = "환자번호"
        chkSprSearch1.BackColor = &HFF8080
    Else
        chkSprSearch1.Caption = "Barcode"
        chkSprSearch1.BackColor = &HFF00&
    End If
End Sub

Private Sub cmdDataSearch_Click()
    Dim Search_Flag As Integer
    Dim i As Integer
    
    Search_Flag = -1
    
    If chkSprSearch1.Caption = "Barcode" Then
        For i = 1 To vasID.DataRowCnt
            If GetText(vasID, i, colBarcode) = txtSearch Then
                vasID.Row = i
                vasID.Col = 1
                vasID.Action = ActionGotoCell
                vasID.Action = ActionActiveCell
                Search_Flag = 1
                Exit For
            End If
        Next i
        If Search_Flag = 1 Then Exit Sub
        For i = 1 To vasWorkList.DataRowCnt
            If GetText(vasWorkList, i, colBarcode) = txtSearch Then
                vasWorkList.Row = i
                
                vasWorkList.Col = 1
                vasWorkList.Action = ActionGotoCell
                vasWorkList.Action = ActionActiveCell
                Search_Flag = 1
                Exit For
            End If
        Next i
    Else
        For i = 1 To vasID.DataRowCnt
            If GetText(vasID, i, colBarcode) = txtSearch Then
                vasID.Col = colBarcode
                vasID.SetFocus
                Search_Flag = 1
                Exit For
            End If
        Next i
        If Search_Flag = 1 Then Exit Sub
        For i = 1 To vasWorkList.DataRowCnt
            If GetText(vasWorkList, i, colBarcode) = txtSearch Then
                vasWorkList.Col = colBarcode
                vasWorkList.SetFocus
                Search_Flag = 1
                Exit For
            End If
        Next i
    End If
    
    
    If Search_Flag = -1 Then MsgBox "검색조건을 확인해 주세요"
    
    
End Sub

Private Sub cmdSprSearch_Click()
    Dim Search_Flag As Integer
    Dim i As Integer
    
    Search_Flag = -1
    
    If chkSprSearch.Caption = "Barcode" Then
        For i = 1 To vasRID.DataRowCnt
            If GetText(vasRID, i, colBarcode) = txtSprSearch Then
                vasRID.Row = i
                vasRID.Col = 1
                vasRID.Action = ActionGotoCell
                vasRID.Action = ActionActiveCell
                Search_Flag = 1
                Exit For
            End If
        Next i
    Else
        For i = 1 To vasRID.DataRowCnt
            If GetText(vasRID, i, colBarcode) = txtSprSearch Then
                vasRID.Col = colBarcode
                vasRID.SetFocus
                Search_Flag = 1
                Exit For
            End If
        Next i
    End If
    
    
    If Search_Flag = -1 Then MsgBox "검색조건을 확인해 주세요"
    
    
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Call cmdDataSearch_Click
End Sub

Private Sub cmdExcelSave_Click()
    Dim i As Long
    Dim j As Long
    Dim myExcelFile As New ExcelFile
    Dim FileName    As String
    
    Screen.MousePointer = 11
    
    'EXCEL파일 작성

    FileName = App.Path & "\TEST1.Xls"
    
    With myExcelFile
        .CreateFile (FileName)
        .SetFont "Arial", 10, xlsNoFormat               ' xlsFont0
        .SetFont "Courier", 15, xlsBold                 ' xlsFont1
        .SetFont "Arial", 20, xlsItalic                 ' xlsFont2
        '''.SetColumnWidth 1, 6, 18                        ' 1 컬럼의 넓이를 18 로 정한다.
        .PrintGridLines = True
        

        For i = 1 To vasRID.MaxCols
            vasRID.Col = i
            For j = 0 To vasRID.MaxRows
                vasRID.Row = j
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, j + 1, i, vasRID.Text
            Next j
        Next i

    End With
    Screen.MousePointer = 0
    MsgBox "엑셀파일이 만들어 졌습니다." & vbCrLf & vbCrLf & "EXCEL 파일명 : '" & FileName & "'"
End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasWorkList, 1, vasID.MaxRows, 1, vasWorkList.MaxCols, 0, 0, 0
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    
    ClearSpread vasWorkList
    ClearSpread vasID
    ClearSpread vasRes
    
    vasWorkList.MaxRows = 0
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.value = 1 Then
            res = Insert_Data_ABI7500(lRow)
        
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasID, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.value = 0
        End If
    Next lRow
End Sub

Private Sub cmdLoad_Click()
    Dim xlApp1 As Excel.Application
    Dim xlApp2 As Excel.Application
    Dim xlApp3 As Excel.Application
    Dim xlBook1 As Excel.Workbook
    Dim xlBook2 As Excel.Workbook

    Dim sFilePath As String
    Dim sFileName As String
    Dim sResStart As Boolean
    Dim sVarRes() As String
    Dim sVarCnt As Integer
    Dim sVarPos As String
    Dim sVarRow As Integer
    Dim sVarCol As Integer
    
    Dim sResult As String
    Dim sResCnt As Integer
    Dim sReceDate As String
    Dim sReceNo As String
    Dim sEquipCode As String
    Dim sExamName As String
    Dim sExamDate As String
    Dim sExamCode As String
    Dim sBarcode As String
    Dim sRow As Integer
    Dim sResIU As String
    Dim sResCopy As String
    Dim sEV As String
    Dim sSV As String
    Dim sPos As String
    
    Dim sSex As String
    Dim sAge As String
    
    Dim sMakeFile As String
    Dim i, j As Integer
    Dim X, Y As Integer
    Dim lRow As Integer
    
    Const XL_NOTRUNNING As Long = 429
    
On Error GoTo ErrPnt:
    
    ClearSpread vasExcelRes
    
    'ClearSpread vasID
    
'결과파일 열기 ===========================================================
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Excel Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
    
    CommonDialog1.ShowOpen
    If Err.Number = 32755 Then: Exit Sub
    

    
    If Trim(CommonDialog1.FileName) = "" Then
        Exit Sub
    End If
    sFilePath = CommonDialog1.FileName
    
'    Set xlApp1 = GetObject(, "Excel.Application")
    Set xlApp1 = New Excel.Application
'    xlApp1.Visible = True
    xlApp1.DisplayAlerts = False
    
'    xlApp1.Workbooks.Add
    xlApp1.Workbooks.Open sFilePath
    
    Set xlBook1 = xlApp1.ActiveWorkbook
    sFileName = xlBook1.Name
    
    
    
    For i = 1 To 320
        vasExcelRes.MaxRows = vasExcelRes.MaxRows + 1
        For X = 1 To 14
            vasExcelRes.SetText X, i, xlApp1.Cells(i, X)
        Next
    Next
    
    vasExcelRes.MaxRows = vasExcelRes.DataRowCnt
    
    xlApp1.DisplayAlerts = True
    xlApp1.Workbooks(sFileName).Close
    Set xlApp1 = Nothing
'=========================================================================

'결과넣기 ================================================================
    
    
    
    
    For i = vasExcelRes.DataRowCnt To 25 Step -1
        If GetText(vasExcelRes, i, 2) = "" Then
            Call DeleteRow(vasExcelRes, i, i)
        End If
    Next i
    
    For i = 24 To 1 Step -1
        Call DeleteRow(vasExcelRes, i, i)
    Next i
    
    vasExcelRes.MaxRows = vasExcelRes.DataRowCnt
    
    
    j = 1
    Do While j <= vasExcelRes.DataRowCnt
    
        
        '///// 바코드번호가 같으면 그쪽으로 들어 가기
        lRow = -1
        
        For i = 1 To vasID.DataRowCnt
            If GetText(vasID, i, colBarcode) = GetText(vasExcelRes, j, 2) Or GetText(vasID, i, colPos) = GetText(vasExcelRes, j, 1) Then
                lRow = i
                Exit For
            End If
        Next i
        
        '///// 순서대로
        For i = 1 To vasID.DataRowCnt
            If GetText(vasID, i, colState) = "" Then
                lRow = i
                Exit For
            End If
        Next i
        
    
        If lRow = -1 Then: vasID.MaxRows = vasID.MaxRows + 1: lRow = vasID.MaxRows
        
        'vasResTemp%
        
        If GetText(vasID, lRow, 2) = "" Then
            SetText vasID, GetText(vasExcelRes, j, 2), lRow, colBarcode
            
            If IsNumeric(GetText(vasExcelRes, j, 2)) = True Then
                SetText vasID, GetText(vasExcelRes, j, 2), lRow, colExamDate + 1
            End If
            
        End If
        SetText vasID, GetText(vasExcelRes, j, 1), lRow, colPos
        
        Call Get_Sample_Info(lRow)
        
        sEquipCode = GetText(vasExcelRes, j, 3)
        sResult = GetText(vasExcelRes, j, 5)
        
        If sResult = "Undetermined" Then: sResult = "Negative"
        
        Call EquipExamCode(sEquipCode, GetText(vasID, lRow, colBarcode))
    
        sExamCode = gEquipExamCode
        
        sAge = Trim(GetText(vasID, lRow, colAge))
        If sAge = "" Then: sAge = "0"
        
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EXAMNAME "
        SQL = SQL & vbCrLf & "  FROM EQUIPEXAM "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & sEquipCode & "' "
        res = db_select_Col(gLocal, SQL)
        
        If sExamCode = "" Then
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT EXAMNAME, EXAMCODE "
            SQL = SQL & vbCrLf & "  FROM EQUIPEXAM "
            SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & sEquipCode & "' "
            res = db_select_Col(gLocal, SQL)
            sExamCode = gReadBuf(1)
        End If
        sExamName = gReadBuf(0)
        
        
        '//// 결과 저장부분
        sExamDate = Format(dtpToday, "yyyymmdd")
            
        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
              "WHERE EXAMDATE = '" & sExamDate & "' " & vbCrLf & _
              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "  AND BARCODE = '" & Trim(GetText(vasID, lRow, colBarcode)) & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(sEquipCode) & "'" & vbCrLf & _
              "  and examcode= '" & Trim(sExamCode) & "'"
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        
        SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
              "                    POSNO, PID, PNAME, " & vbCrLf & _
              "                    PSEX, PAGE, EXAMDATE, " & vbCrLf & _
              "                    EQUIPCODE, EXAMCODE, SEQNO, " & vbCrLf & _
              "                    RESULT, EXAMNAME, SENDFLAG, " & vbCrLf & _
              "                    REFFLAG, EQUIPRESULT, RECENO, RESFLAG) " & vbCrLf & _
              "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, lRow, colBarcode)) & "', '" & Trim(GetText(vasID, lRow, colRack)) & "', " & vbCrLf & _
              "       '" & Trim(GetText(vasID, lRow, colPos)) & "', '" & Trim(GetText(vasID, lRow, colPID)) & "', '" & Trim(GetText(vasID, lRow, colPName)) & "', " & vbCrLf & _
              "       '" & Trim(GetText(vasID, lRow, colSex)) & "', " & Trim(sAge) & ", '" & Trim(sExamDate) & "', " & vbCrLf & _
              "       '" & Trim(sEquipCode) & "', '" & Trim(sExamCode) & "', " & vbCrLf & _
              "       '', '" & Trim(sResult) & "', '" & Trim(sExamName) & "', '1', " & vbCrLf & _
              "       '0', '" & Trim(sResult) & "', '" & Trim(GetText(vasID, lRow, colSpecNo)) & "', '') "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        
        SetText vasID, "결과", lRow, colState
        j = j + 1
    Loop
'=========================================================================
    
    ' vasID.Sort colBarcode, 1, colBarcode, vasID.DataRowCnt, SortByRow
     vasSort vasID, colExamDate + 1, colBarcode
    If MnTransAuto.Checked = True Then
        vasID.Row = lRow
        vasID.Col = colSpecNo
        If vasID.value <> "" Then
            vasID.Row = lRow
            vasID.Col = 1
            vasID.value = 1
            
            res = Insert_Data_ABI7500(lRow)
        
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "전송", lRow, colState
                
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.value = 0
        End If
    End If
    
    Exit Sub
ErrPnt:
    
    Exit Sub
End Sub

Private Sub cmdLocload_Click()
    Dim i As Integer
    
    
    '/검사완료-------------------------------------------------------------
    ClearSpread vasID
                   
                   SQL = "SELECT WORKNO, BARCODE, SAMPLENO, POS, EXAMCODE, RESULTTIME, RESULT, STARTDATE, ENDDATE  "
    SQL = SQL & vbCrLf & "  FROM SEND_RESULT "
    SQL = SQL & vbCrLf & " WHERE RESULT <> '' "
    SQL = SQL & vbCrLf & "ORDER BY STARTDATE DESC "
    res = db_select_Vas(gLocal, SQL, vasID, , 2)
    
    
    'SQL = SQL & vbCrLf & ""
    'SQL = SQL & vbCrLf & ""
    'SQL = SQL & vbCrLf & ""
    'SQL = SQL & vbCrLf & ""
    '/----------------------------------------------------------------------
    
    '/검사중-------------------------------------------------------------
    ClearSpread vasWorkList
                   SQL = "SELECT WORKNO, BARCODE, SAMPLENO, POS, EXAMCODE, RESULTTIME, RESULT, STARTDATE, ENDDATE  "
    SQL = SQL & vbCrLf & "  FROM SPEC_RESULT "
    SQL = SQL & vbCrLf & "ORDER BY STARTDATE DESC "
    res = db_select_Vas(gLocal, SQL, vasWorkList, , 2)

    '/----------------------------------------------------------------------
    
    '/카운트 만들기---------------------------------------------------------
                   SQL = "SELECT COUNT(*)  "
    SQL = SQL & vbCrLf & "  FROM SEND_RESULT "
    SQL = SQL & vbCrLf & " WHERE RESULT = 'POSITIVE' "
    res = db_select_Col(gLocal, SQL)
    sspPositive.Caption = gReadBuf(0)
    
                   SQL = "SELECT COUNT(*)  "
    SQL = SQL & vbCrLf & "  FROM SEND_RESULT "
    SQL = SQL & vbCrLf & " WHERE RESULT = 'No Growth for 5 Days' "
    res = db_select_Col(gLocal, SQL)
    ssp6Weeks.Caption = gReadBuf(0)
    
    
                   SQL = "SELECT COUNT(*)  "
    SQL = SQL & vbCrLf & "  FROM SPEC_RESULT "
    SQL = SQL & vbCrLf & " WHERE RESULT = 'No Growth for 2 Days' "
    res = db_select_Col(gLocal, SQL)
    ssp3Weeks.Caption = gReadBuf(0)
    '/----------------------------------------------------------------------
    
    For i = 1 To vasID.DataRowCnt
        If GetText(vasID, i, colBact_result) = "POSITIVE" Then
            SetBackColor vasID, i, i, 1, vasID.MaxCols, 255, 0, 0
        End If
    Next i
    
End Sub

Private Sub cmdRClear_Click()
    Dim i As Integer

'    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasRID, 1, vasRID.MaxRows, 1, vasRID.MaxCols, 0, 0, 0
    SetForeColor vasRRes, 1, vasRRes.MaxRows, 1, vasRRes.MaxCols, 0, 0, 0
    
    vasRID.MaxRows = 0
    vasRRes.MaxRows = 0
    
    dtpStartDate = DateAdd("m", -1, Date)
    dtpEndDate = Date
    
    chkPositive.value = 0
    chkPositive.Enabled = 0
    chkNogrow.value = 0
    chkNogrow.Enabled = 0
    
End Sub



Private Sub cmdRSch_Click()
'    Dim iRow As Long
'
'    ClearSpread vasRID
'    ClearSpread vasRRes
'    Call chkRAll_Click
'
' SQL = "SELECT '', RECENO, BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf & _
'          "FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND SENDFLAG IN ('1', '2') " & vbCrLf & _
'          "GROUP BY BARCODE, RECENO, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG "
'
'    res = db_select_Vas(gLocal, SQL, vasRID)
'
'          '"  AND SENDFLAG IN ('1', '2') "
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    For iRow = 1 To vasRID.DataRowCnt
'        If IsNumeric(Trim(GetText(vasRID, iRow, colBarcode))) = True Then
'             SetText vasRID, Trim(GetText(vasRID, iRow, colBarcode)), iRow, colState + 1
'        End If
'    Next iRow
'
'
'    For iRow = 1 To vasRID.DataRowCnt
'        Select Case Trim(GetText(vasRID, iRow, colState))
'        Case "2"
'            SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
'            SetText vasRID, "완료", iRow, colState
''        Case "0"
''            SetText vasID, "오더", iRow, colState
'        Case "1"
'            SetText vasRID, "결과", iRow, colState
'        End Select
'    Next iRow
'
'    vasSort vasRID, colState + 1, colBarcode
End Sub

Private Sub cmdRTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.value = 1 Then
            res = Insert_Data_ABI7500_R(lRow)
        
            If res = -1 Then
                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasRID, "실패", lRow, colState
            ElseIf res = 0 Then
            
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "완료", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquipCode & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasRID.Row = lRow
            vasRID.Col = 1
            vasRID.value = 0
        End If
    Next lRow
End Sub

Private Sub cmdSearch_Click()
    Dim SQL_SPEC        As String
    Dim SQL_EXAM        As String
    Dim StartDate       As String
    Dim EndDate         As String
    
    ClearSpread vasRID
    SQL_SPEC = ""
    SQL_EXAM = ""
    
    StartDate = Format(dtpStartDate, "yyyymmdd") & "000000"
    EndDate = Format(dtpEndDate, "yyyymmdd") & "000000"
    
    '/검사결과 불러오기
    '/옵션에 따라 불러오기
    
    If chkExam.value = 1 Or (chkResult.value = 0 And chkExam.value = 0) Then '/검사중
                            SQL_SPEC = "SELECT WORKNO, BARCODE, SAMPLENO, POS, PID, RESULTTIME, RESULT, STARTDATE AS S, ENDDATE AS E "
        SQL_SPEC = SQL_SPEC & vbCrLf & "  FROM SPEC_RESULT "
        SQL_SPEC = SQL_SPEC & vbCrLf & " WHERE STARTDATE BETWEEN '" & StartDate & "' AND '" & EndDate & "' "
        'SQL_SPEC = SQL_SPEC & vbCrLf & "ORDER BY ENDDATE DESC "
    End If
    
    If chkResult.value = 1 Or (chkResult.value = 0 And chkExam.value = 0) Then '/검사완료
                            SQL_EXAM = "SELECT WORKNO, BARCODE, SAMPLENO, POS, PID, RESULTTIME, RESULT, STARTDATE AS S, ENDDATE AS E  "
        SQL_EXAM = SQL_EXAM & vbCrLf & "  FROM SEND_RESULT "
        SQL_EXAM = SQL_EXAM & vbCrLf & " WHERE STARTDATE BETWEEN '" & StartDate & "' AND '" & EndDate & "' "
        
        If (chkPositive.value = 1 And chkNogrow.value = 1) Or (chkPositive.value = 0 And chkNogrow.value = 0) Then
            '/둘다 체크를 (안)할때는 전체조회
        ElseIf chkPositive.value = 1 Then
            SQL_EXAM = SQL_EXAM & vbCrLf & "   AND RESULT = 'POSITIVE' "
        ElseIf chkNogrow.value = 1 Then
            SQL_EXAM = SQL_EXAM & vbCrLf & "   AND RESULT <> 'POSITIVE'"
        End If
        
        'SQL_SPEC = SQL_SPEC & vbCrLf & "ORDER BY ENDDATE DESC "
    End If
    
    '/옵션 체크에 따라 쿼리문이 바뀜
    If (chkResult.value = 1 And chkExam.value = 1) Or (chkResult.value = 0 And chkExam.value = 0) Then
        SQL = "" & SQL_SPEC & vbCrLf & " UNION " & SQL_EXAM & " "
        SQL = SQL & vbCrLf & "ORDER BY S DESC "
    ElseIf chkResult.value = 1 Then
        SQL = SQL_EXAM
    ElseIf chkExam.value = 1 Then
        SQL = SQL_SPEC
    End If
    
    res = db_select_Vas(gLocal, SQL, vasRID, , 2)
    
End Sub

Private Sub cmdWorkList_Click()
    Dim i As Integer
    Call GetWorkList(dtpFrDt.value, dtpToDt.value)
    
    For i = 1 To vasWorkList.DataRowCnt
        If GetText(vasWorkList, i, colState) = "1" Then
            SetText vasWorkList, "결과입력", i, colState
        ElseIf GetText(vasWorkList, i, colState) = "0" Then
            SetText vasWorkList, "접수", i, colState
        End If
    Next i
End Sub

Private Sub lblclear_Click()
    lblChangeBar.Caption = ""
    lblBarcode.Caption = ""
    lblChangePID.Caption = ""
    lblPname.Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
       
    For i = 1 To Len(txtTest)
        lsChar = Mid(txtTest, i, 1)

    Select Case lsChar
        Case chrENQ
            txtData = ""
            
            SaveData "[Rx]" & chrENQ
                    
            'MSComm1.Output = chrACK
            SaveData "[Tx]" & chrACK
                    
        Case chrSTX     '자료수신 시작

            txtData.Text = ""
            
        Case chrETX
            txtData.Text = txtData.Text & lsChar

        Case chrLF
            
            txtData.Text = txtData.Text & lsChar
            SaveData "[Rx]" & chrSTX & txtData.Text
            
            Call BACT(txtData)
            txtData = ""
            'MSComm1.Output = chrACK
            SaveData "[Tx]" & chrACK
            
        Case chrEOT     '자료수신 완료
            SaveData "[Rx]" & chrEOT
            txtData.Text = ""
               
        Case Else
            txtData.Text = txtData.Text & lsChar

    End Select
    Next i
    
    txtTest = ""

End Sub

Function BACT(asData As String)
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    
    Dim iCnt As Integer
    Dim iCnt_buff As Integer
    
    Dim lsID As String
    Dim lsPid As String
    Dim lsPName As String
    Dim lsPSex As String
    Dim lsPage As String
    
    Dim lsTestNo As String
    
    Dim lsTestID As String
    Dim lsSubCode As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    
    Dim lsresult_IFCC As String
    Dim lsresult_eAg As String
    
    Dim sSampleType As String
    Dim sLotNo As String

    If asData = "" Then
        Exit Function
    End If
    X = 0
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
        tmResultCheck.Enabled = False
        StatusBar1.Panels.Item(1) = " 장비와 통신중 "
        sResult_Flag = 0
        Var_Clear
        gsSampleType = ""
        Result_gubun = ""
        Data_gubun = ""
        iCnt = 0
        POS_FLAG = 1 '/POSITIVE 결과 나온거 체크(1:기본, -1:양성)
'        For i = 1 To Len(asData)
'            If Mid(asData, i, 1) = "|" Then
'                iCnt = iCnt + 1
'
'                Select Case iCnt
'                    Case 11
'                        gsSampleType = Mid(asData, i + 1, 1)
'                    Case 13
'                End Select
'            End If
'        Next i
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "O" Then
        
        If sResult_Flag = 2 Or sResult_Flag = 1 Then
            If POS_FLAG = 1 Then '/1이면 음성 -1 이면 양성임
                
                For i = 1 To vasID.DataRowCnt
                    If gsBarCode <> "" Then
                        If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                            gRow = i
                            Call Insert_Data(gRow, Result_gubun, Data_gubun)
                            
                            'Call Insert_Data_SE_LAST(gRow, Result_gubun, Data_gubun)
                            
                            'Call Insert_Data_SE(gRow, Result_gubun, Data_gubun)
                            'Exit For
                        End If
                    End If
                Next i
            ElseIf POS_FLAG = -1 Then
                For i = 1 To vasID.DataRowCnt
                    If gsBarCode <> "" Then
                        If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                            gRow = i
'                            Call Save_Local_One_MI(gRow, "", "")
'                            Call Save_Local_One_MI(gRow, "", Now)
                            'DeleteRow vasWorkList, gRow, gRow
                            
                        End If
                    End If
                Next i
            End If
            
        ElseIf sResult_Flag > 2 Then
            If POS_FLAG = 1 Then
                For i = 1 To vasID.DataRowCnt
                    If gsBarCode <> "" Then
                        If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                            gRow = i
                            Call Save_Local_One_MI(gRow, "SAMPLE ERR", "")
                            
                        End If
                    End If
                Next i
            ElseIf POS_FLAG = -1 Then
                
            End If
        End If
        
        '/이전 바코드번호에 대해서 처리후 새 신호 처리---------------------------------------------
        sTmp = Trim(ResultTbl(3))      'Barcode
        
        gsBarCode = sTmp
        POS_FLAG = 1
        sResult_Flag = 0
        '/이전 바코드번호에 대해서 처리후 새 신호 처리---------------------------------------------
        
        
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "P") Then     '없음
        
    End If
    
    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        Result_gubun = ""
        Data_gubun = ""
        lsTestNo = ""
        
        lsTestNo = ResultTbl(2)
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsTestID = Left(sTmp, i - 1)    '장비코드(혐기성,호기성, ETC 뭐 그런거 나옴)
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsEquipCode = Left(sTmp, i - 1) 'Mid(sTmp, i + 1)
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsSubCode = Mid(sTmp, i + 1)    '병번호 나옴
        sTmp = ResultTbl(4)
        lsResult = Trim(sTmp)           '결과
        
        
        
        
        Dim StartDate As String
        Dim EndDate As String
        Dim SamplePos As String
        
        sTmp = ResultTbl(12)
        StartDate = Trim(sTmp)
        sTmp = ResultTbl(13)
        EndDate = Trim(sTmp)
        sTmp = ResultTbl(14)
        SamplePos = Trim(Mid(sTmp, 1, 5))
        
        '////////////////////// 결과가 아닌것은 처리 안함
        If lsTestID <> "BC" Then Exit Function
        
        gRow = -1
        For i = 1 To vasWorkList.DataRowCnt
            If gsBarCode <> "" Then
                If Trim(GetText(vasWorkList, i, colBarcode)) = gsBarCode And Trim(GetText(vasWorkList, i, colRack)) = SamplePos Then
                    gRow = i
                    Exit For
                End If
'           ElseIf sSampleType = "Q" Then

            End If
        Next i
        
        If gRow < 0 Then
            gRow = vasWorkList.DataRowCnt + 1
            If vasWorkList.MaxRows < gRow Then
                vasWorkList.MaxRows = gRow
            End If
        End If
        If sSampleType = "Q" Then Exit Function
        
        If Trim(GetText(vasWorkList, i, colBarcode)) = "" Then
            SetText vasWorkList, gsBarCode, gRow, colBarcode
        Else
        End If
        SetText vasWorkList, gDate, gRow, colExamDate
        SetText vasWorkList, gsPosNo, gRow, colPos
        
        vasActiveCell vasWorkList, gRow, colBarcode
        ClearSpread vasRes
        

        If Trim(GetText(vasWorkList, gRow, colSpecNo)) = "" Then 'And Len(gsBarCode) = 10
            Get_Sample_Info gRow
        End If

        
        
        '/---------오더 찾기-------------------------------------------------------------------
        SQL = "SELECT EXAMCODE "
        SQL = SQL & vbCrLf & " FROM EQUIPEXAM  "
        SQL = SQL & vbCrLf & "WHERE EQUIPCODE = '" & Trim(lsEquipCode) & "' "
        'SQL = SQL & vbCrLf & "  AND SEQNO = " & lsTestNo & " "
        res = db_select_Col(gLocal, SQL)
        
        If gReadBuf(0) = "" Then
            gReadBuf(0) = "''"
        Else
            gReadBuf(0) = "'" & gReadBuf(0) & "' "
        End If
        '/---------오더 찾기--------------------------------------------------------------------
        
        
        '/---------오더 찾기--------------------------------------------------------------------
        SQL = "SELECT EXMN_CD "
        SQL = SQL & vbCrLf & " FROM SPSLHRRST  "
        SQL = SQL & vbCrLf & "WHERE WORK_NO = '" & Trim(GetText(vasWorkList, gRow, colSpecNo)) & "' "
        SQL = SQL & vbCrLf & "  AND EXMN_CD IN (" & gReadBuf(0) & ") "
        res = db_select_Col(gServer, SQL)
        '/---------------------------------------------------------------------------------------
        
        If lsEquipCode = "BPF" Then
            SetText vasWorkList, lsEquipCode, gRow, colExamCD_MI    '검사코드
        Else
            SetText vasWorkList, gReadBuf(0), gRow, colExamCD_MI    '검사코드
        End If
        
        SetText vasWorkList, lsSubCode, gRow, colSampleNo           '병번호
        SetText vasWorkList, SamplePos, gRow, colRack               '위치
        SetText vasWorkList, StartDate, gRow, colStartDate          '시작일자
        
        If lsResult = "*" Then
            SetText vasWorkList, "", gRow, colEndDate       '장비코드
        Else
            SetText vasWorkList, EndDate, gRow, colEndDate       '장비코드
        End If
        StartDate = Format(StartDate, "####-##-## ##:##:##")
        EndDate = Format(EndDate, "####-##-## ##:##:##")
        
        '====================
        ' 2 day - 중간 보고
        ' 5 day - 최종 보고
        ' 담당병리사가 verify해야만 전송
        '====================
        If lsResult <> "*" Then
            sResult_Flag = sResult_Flag + 1
            If lsResult = "-" Then  '/lsResult = "*" Or 결과줃은 뺏음 (2일차)
                '-- 검사시작 : 검사시작시간 = 장비결과시간
                If CStr(DateAdd("h", 24, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then
                    lsResult = ""
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 1 Day No Growth : 검사시작시간 + 24시간 <= 장비결과시간 AND 검사시작시간 + 48시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 24, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 48, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 1day no growth
                    lsResult = "No growth for 1 Days"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 2 Day No Growth : 검사시작시간 + 48시간 <= 장비결과시간 AND 검사시작시간 + 72시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 48, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 72, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 2day no growth
                    lsResult = "No growth for 2 Days"
                    Result_gubun = "2"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 3 Day No Growth : 검사시작시간 + 72시간 <= 장비결과시간 AND 검사시작시간 + 96시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 72, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 96, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 3day no growth
                    lsResult = "No Growth for 3 Days"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 4 Day No Growth : 검사시작시간 + 96시간 <= 장비결과시간 AND 검사시작시간 + 120시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 96, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 120, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 4day no growth
                    lsResult = "No Growth for 4 Day"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 5 Day No Growth : 검사시작시간 + 96시간 <= 장비결과시간 AND 검사시작시간 + 168시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 120, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 168, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then '-- 5day no growth
                    lsResult = "No Growth for 5 Days"
                    Result_gubun = "3"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
    
                '-- 7 Day No Growth : 검사시작시간 + 168시간 <= 장비결과시간
                ElseIf CStr(DateAdd("h", 168, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) Then  '-- 7day no growth
                    lsResult = "No Growth for 7 Days"    '-- 최종
                    SetText vasWorkList, "148", gRow, colBact_Time
                End If
            ElseIf lsResult = "+" Then
                '-- 검사시작 : 검사시작시간 = 장비결과시간
                If CStr(DateAdd("h", 24, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then
                    lsResult = "POSITIVE"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 1 Day No Growth : 검사시작시간 + 24시간 <= 장비결과시간 AND 검사시작시간 + 48시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 24, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 48, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 1day no growth
                    lsResult = "POSITIVE"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 2 Day No Growth : 검사시작시간 + 48시간 <= 장비결과시간 AND 검사시작시간 + 72시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 48, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 72, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 2day no growth
                    lsResult = "POSITIVE"
                    Result_gubun = "0"
                    
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 3 Day No Growth : 검사시작시간 + 72시간 <= 장비결과시간 AND 검사시작시간 + 96시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 72, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 96, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 3day no growth
                    lsResult = "POSITIVE"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 4 Day No Growth : 검사시작시간 + 96시간 <= 장비결과시간 AND 검사시작시간 + 120시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 96, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 120, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then  '-- 4day no growth
                    lsResult = "POSITIVE"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 5 Day No Growth : 검사시작시간 + 96시간 <= 장비결과시간 AND 검사시작시간 + 168시간 > 장비결과시간
                ElseIf CStr(DateAdd("h", 120, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) And CStr(DateAdd("h", 168, StartDate)) > CStr(DateAdd("h", 0, EndDate)) Then '-- 5day no growth
                    lsResult = "POSITIVE" '-- 최종
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                '-- 7 Day No Growth : 검사시작시간 + 168시간 <= 장비결과시간
                ElseIf CStr(DateAdd("h", 168, StartDate)) <= CStr(DateAdd("h", 0, EndDate)) Then  '-- 7day no growth
                    lsResult = "POSITIVE"
                    Result_gubun = "0"
                    SetText vasWorkList, DateDiff("h", StartDate, EndDate), gRow, colBact_Time
                End If
            End If
            SetText vasWorkList, lsResult, gRow, colBact_result         '검사결과
            If lsResult = "POSITIVE" Then
                SetText vasWorkList, "P", gRow, colState
                POS_FLAG = -1
            End If
            Call Save_Local_One_MI(gRow, lsResult, EndDate)
            
            If lsResult = "POSITIVE" Then
                For i = vasWorkList.DataRowCnt To 1 Step -1
                    If Trim(GetText(vasWorkList, i, colBarcode)) = gsBarCode Then
                        Call Save_Local_One_MI(i, Trim(GetText(vasWorkList, i, colBact_result)), EndDate)
                        POS_FLAG = -1
                        'vasWorkList.MaxRows = vasWorkList.MaxRows - 1
                        'Exit For
                    End If
                Next i
            End If
        
        Else
            POS_FLAG = 0
            Call Save_Local_One_MI(gRow, lsResult, "")
            'Call Insert_Data_SE_FIRST(gRow, "", "")
        End If
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "L" Then
        If sResult_Flag = 2 Or sResult_Flag = 1 Then
            If POS_FLAG = 1 Then '/1이면 음성 -1 이면 양성임
                
                For i = 1 To vasID.DataRowCnt
                    If gsBarCode <> "" Then
                        If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                            gRow = i
                            Call Insert_Data(gRow, Result_gubun, Data_gubun)
                            
                            'Call Insert_Data_SE_LAST(gRow, Result_gubun, Data_gubun)
                            
                            'Call Insert_Data_SE(gRow, Result_gubun, Data_gubun)
                            sResult_Flag = 0
                        End If
                    End If
                Next i
                
            End If
        ElseIf sResult_Flag > 2 Then
'            If POS_FLAG = 1 Then
'                For i = 1 To vasID.DataRowCnt
'                    If gsBarCode <> "" Then
'                        If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
'                            gRow = i
'                            Call Save_Local_One_MI(gRow, "SAMPLE ERR", "")
'                            sResult_Flag = 0
'                        End If
'                    End If
'                Next i
'            ElseIf POS_FLAG = -1 Then
'
'            End If
        End If

        tmResultCheck.Enabled = True
        StatusBar1.Panels.Item(1) = " 데이터 체크 시작 "
    End If
'        If Trim(GetText(vasID, gRow, colPName)) <> "" Then
'
'            gOrderExam = ""
'            If MnTransAuto.Checked = True Then
'                res = Insert_Data_PhD(gRow)
'
'                If res = -1 Then
'                    SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
'                    SetText vasID, "Failed", gRow, colState
'                Else
'
'                    SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
'                    SetText vasID, "Trans", gRow, colState
'
'                    SQL = " Update pat_res Set " & vbCrLf & _
'                          " sendflag = '2' " & vbCrLf & _
'                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'                          " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
'                    res = SendQuery(gLocal, SQL)
'                    If res = -1 Then
'                        SaveQuery SQL
'                        Exit Function
'                    End If
'
'                End If
'            Else
'                SetText vasID, "Result", gRow, colState
'            End If
'
'        End If
'
'    End If
    
End Function

Private Sub URISCAN_PRO(asData As String)
    Dim MyVar As String
    Dim MyRet As String
          
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim iRow As Integer
    Dim jRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sResult As String
    
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sDate As String
    
    Dim lsSeq As String
    Dim lsCnt As String
    
    If Trim(asData) = "" Then
        Exit Sub
    End If
    
    MyVar = Trim(asData)
         
    sDate = Format(dtpToday, "yyyymmdd")
    
    i = InStr(1, MyVar, "Date")
    If i > 0 Then
        sDate = Format(CDate(Trim(Mid(MyVar, i + 6, 20))), "yyyy-mm-dd hh:nn:ss")
    End If
    
    i = InStr(1, MyVar, "ID_NO")
    sSeqNo = CStr(CLng(Trim(Mid(MyVar, i + 6, 4))))

    sBarcode = CStr(Trim(Mid(MyVar, i + 11, 12)))
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    llRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, 0)) = sSeqNo Then
            llRow = iRow
            Exit For
        End If
        
        If Trim(GetText(vasID, iRow, 0)) = "" Then
            llRow = iRow
            Exit For
        End If
    Next iRow

    If llRow = -1 Then
        llRow = vasID.DataRowCnt + 1
        If llRow > vasID.MaxRows Then
            vasID.MaxRows = llRow
        End If
    End If
    
    ClearSpread vasRes, 1, 1

    SetText vasID, sSeqNo, llRow, 0
    'SetText vasID, sExamDate, llRow, colDate
    'SetText vasID, sDate, llRow, colTime
    SetText vasID, sBarcode, llRow, colBarcode
    
    '수신중========================================================
    SetText vasID, "수신중", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
    '==============================================================
    
    '샘플의 환자 정보 가져오기
    Get_Sample_Info llRow
    
    '검사코드만큼 Row의 갯수를 설정
    gReadBuf(0) = "0"
    
    SQL = "Select count(examcode) From equipexam" & vbCrLf & _
          " Where equipno = '" & gEquip & "' "
    res = db_select_Col(gLocal, SQL)

    vasRes.MaxRows = Trim(gReadBuf(0))

    
    lsSeq = ""
    lsCnt = ""
        
    
    '결과 잘라 넣기
    j = 0
    For j = 1 To vasRes.MaxRows
        sExamName = Trim(GetText(vasCode, j, 1))
        
        Select Case sExamName
        Case "BLD", "BIL", "URO", "KET", "PRO", "NIT", "GLU", "LEU"
            i = InStr(1, MyVar, Trim(sExamName))
            sResult = Trim(Mid(MyVar, i + 3, 8))

        Case "p.H"
            i = InStr(1, MyVar, "p.H")
            sResult = Trim(Mid(MyVar, i + 3, 14))

        Case "S.G"
            i = InStr(1, MyVar, "S.G")

            If Mid(MyVar, i) = "<=" Or Mid(MyVar, i) = ">=" Then
                sResult = Trim(Mid(MyVar, i + 3, 9))
            Else
                sResult = Trim(Mid(MyVar, i + 3, 12))
            End If
        End Select
        
        Select Case sResult
        Case "-"
            sResult = "Negatvie"
        End Select
        
        ClearSpread vasTemp
        
        SQL = "Select examcode, '', examname From EquipExam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  And EquipCode = '" & Trim(sExamName) & "'"
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        For jRow = 1 To vasTemp.DataRowCnt
            sExamCode = Trim(GetText(vasTemp, jRow, 1))
            sSeqNo = Trim(GetText(vasTemp, jRow, 2))
            sExamName = Trim(GetText(vasTemp, jRow, 3))
        
            SetText vasRes, Trim(sExamName), j, colEquipCode '장비코드
            SetText vasRes, sExamCode, j, colExamCode '검사코드
            SetText vasRes, sExamName, j, colExamName '검사명
            SetText vasRes, Trim(sResult), j, colResult   '검사결과
            SetText vasRes, sSeqNo, j, colSeq        '순번(서브코드)
            Trim (GetText(vasID, llRow, 0))
            Save_Local_One llRow, j, "1", CStr(Trim(sResult))
        Next jRow
    Next j
    gReadBuf(0) = ""
    
    '수신중========================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    If App.PrevInstance Then
        End
    End If

    Me.Left = 0
    Me.Top = 0

    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    sspPositive.Caption = "0"
    ssp3Weeks.Caption = "0"
    ssp6Weeks.Caption = "0"
    
    
    GetSetup
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If

    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    
    '-- osw 추가
     For i = 1 To 3
        If Not Connect_PRServer Then
            cn_cnt = cn_cnt + 1
            If cn_cnt = 3 Then
                If Not Connect_DRServer Then
                    MsgBox "연결되지 않았습니다."
                    cn_Server_Flag = False
                    Exit Sub
                Else
                    cn_Server_Flag = True
                End If
            End If
        Else
            cn_Server_Flag = True
        End If
    Next

    GetExamCode
    
    
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.value), -90), "yyyymmdd")
    
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    POS_FLAG = 1
    dtpFrDt.value = Now
    dtpToDt.value = Now
    sResult_Flag = 0
    Call cmdLocload_Click
    '==============================
    
    
End Sub



Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From equipexam " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasCode)
    If res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 6)
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
        For j = 1 To 5
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
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.value = 1
    
End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.value = 0
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

Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim j As Integer
    Dim rs As ADODB.Recordset
    Dim sSpecNo As String
    Dim Server_date As String
    Dim buff As String
    
    vasWorkList.MaxRows = 0
    
    '-- 로컬 검사코드 찾기
          SQL = "Select distinct examcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    
    res = db_select_Row(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    Server_date = Format(Now, "yyyymmdd")
    buff = "0.7"
    '-- 검사대상자 가져오기
    SQL = "Select distinct SPCM_NO From SPSLHRRST "
    SQL = SQL & vbCrLf & " WHERE RGST_DT BETWEEN SYSDATE - " & CLng(Server_date) - (CLng(Format(pFrDt, "yyyymmdd")) - CCur(buff))
    SQL = SQL & vbCrLf & "                                     AND SYSDATE - " & CLng(Server_date) - CLng(Format(pToDt, "yyyymmdd"))
    SQL = SQL & vbCrLf & "   and exmn_cd in (" & strExamCode & ")"
    SQL = SQL & vbCrLf & "   and rslt_no IS NOT NULL"
          
    Set rs = cn_Ser.Execute(SQL, , 1)
          
    Do Until rs.EOF
        SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(rs.Fields(0)) & "') FROM DUAL "
        res = db_select_Col(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
        SQL = "SELECT PID, PT_NM, SEX, AGE, RSLT_STAT "
        SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
        SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & Trim(rs.Fields(0)) & "' "
        SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
        SQL = SQL & vbCrLf & "  AND RSLT_STAT <= '1' "
        res = db_select_Col(gServer, SQL)
        
        j = j + 1
        vasWorkList.MaxRows = j
        SetText vasWorkList, Trim(rs.Fields(0)), j, colSpecNo     '2
        SetText vasWorkList, sSpecNo, j, colBarcode     '3
        SetText vasWorkList, Trim(gReadBuf(0)), j, colPID    '6
        SetText vasWorkList, Trim(gReadBuf(1)), j, colPName  '7
        SetText vasWorkList, Trim(gReadBuf(2)), j, colSex    '8
        SetText vasWorkList, Trim(gReadBuf(3)), j, colAge    '9
            
        rs.MoveNext
    
    Loop

End Sub

'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i               As Integer
    Dim intRow          As Long
    Dim strItems        As String
    Dim BeforeBarcode   As String
    
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBarcode)) = pBarNo Or Trim(GetText(vasID, i, colSpecNo)) = "" Then
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
    
    BeforeBarcode = Trim(GetText(vasID, intRow, colBarcode))
    
    Call vasID_Click(colBarcode, intRow)
    Call DELETE_LOCAL_ONE(BeforeBarcode, Format(dtpToday, "yyyymmdd"))
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    'Call SetText(vasID, mResult.RackNo, intRow, colRack)       '4
    'Call SetText(vasID, mResult.TubePos, intRow, colPos)         '5
    'Call vasActiveCell(vasID, intRow, colBarcode)
    'Call ClearSpread(vasRes)
    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    For i = 1 To vasRes.DataRowCnt
        Call Save_Local_One(intRow, i, "1", Trim(GetText(vasRes, i, colResult)))
    Next i
    
    
    If GetText(vasID, intRow, colPos) <> "" Then: Call SetText(vasID, "결과", intRow, colState) '3
    'gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)

'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = strItems
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = ""
'    End If
    
    BeforeBarcode = ""
End Sub

Private Function GetPos(ByVal pRowNum As String) As String
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    GetPos = ""
    For i = 1 To vasID.DataRowCnt
        If pRowNum = i Then
            intRow = i
            
            GetPos = GetText(vasID, i, colBarcode) '3
            Exit For
        End If
    
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
'    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
'    Call SetText(vasID, mResult.RackNo, intRow, colRack)       '4
'    Call SetText(vasID, mResult.TubePos, intRow, colPos)         '5
'    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
'    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    gRow = intRow
    

'    gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    

End Function

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
    
'    If IsNumeric(sEquipRes) = False Then
'        Exit Function
'    End If
    
    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
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
    'sResult = sResFlag & sResult
    SetResult = sResult
    
End Function

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer
    
'    SQL = "SELECT COUNT(*) FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'    res = db_select_Col(gLocal, SQL)

    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
'    SQL = "SELECT  MAX(RESCNT) FROM PAT_RES WHERE BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "'"
'    res = db_select_Col(gLocal, SQL)
'    If Trim(gReadBuf(0)) = "" Then
'        RCnt = 1
'    Else
'        RCnt = CCur(gReadBuf(0)) + 1
'    End If
    
    SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
          "POSNO, PID, PNAME, " & vbCrLf & _
          "PSEX, PAGE, " & vbCrLf & _
          "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
          "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, " & _
          "EQUIPRESULT, RECENO ) " & vbCrLf & _
          "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPos)) & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
          "'" & Trim(sExamDate) & "', '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colFLAG)) & "', " & _
          "'" & Trim(asEquipResult) & "', '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "' )"
    res = SendQuery(gLocal, SQL)

    
End Function

Function Save_Local_One_MI(ByVal asRow1 As Long, asEquipResult As String, asExamTime As String)
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer
    
    SQL = "SELECT SENDFLAG "
    SQL = SQL & vbCrLf & "  FROM SPEC_RESULT "
    SQL = SQL & vbCrLf & " WHERE WORKNO  = '" & Trim(GetText(vasWorkList, asRow1, colSpecNo)) & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "' "
    SQL = SQL & vbCrLf & "   AND POS     = '" & Trim(GetText(vasWorkList, asRow1, colRack)) & "' "
    res = db_select_Col(gLocal, SQL)
    
    If POS_FLAG = -1 Or asEquipResult = "POSITIVE" Or asEquipResult = "SAMPLE ERR" Or asEquipResult = "No Growth for 5 Days" Then gReadBuf(0) = "1"
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    If gReadBuf(0) = "" Then
                       SQL = "INSERT INTO SPEC_RESULT"
        SQL = SQL & vbCrLf & "      (WORKNO, BARCODE ,SAMPLENO ,POS  "
        SQL = SQL & vbCrLf & "      ,EXAMCODE ,STARTDATE ,SENDFLAG) "
        SQL = SQL & vbCrLf & "VALUES('" & Trim(GetText(vasWorkList, asRow1, colSpecNo)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasWorkList, asRow1, colSampleNo)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasWorkList, asRow1, colRack)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasWorkList, asRow1, colPos)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasWorkList, asRow1, colStartDate)) & "' "
        SQL = SQL & vbCrLf & "      ,'0') "
        
        res = SendQuery(gLocal, SQL)
        
    ElseIf gReadBuf(0) = "0" Then
                       SQL = "UPDATE SPEC_RESULT SET "
        SQL = SQL & vbCrLf & "       RESULT = '" & Trim(GetText(vasWorkList, asRow1, colBact_result)) & "' "
        SQL = SQL & vbCrLf & "      ,RESULTTIME = '" & Trim(GetText(vasWorkList, asRow1, colBact_Time)) & "' "
        SQL = SQL & vbCrLf & "      ,ENDDATE = '" & Trim(Format(asExamTime, "yyyymmddhhmmss")) & "' "
        SQL = SQL & vbCrLf & "      ,SENDFLAG = '1' "
        SQL = SQL & vbCrLf & " WHERE WORKNO = '" & Trim(GetText(vasWorkList, asRow1, colSpecNo)) & "' "
        SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "' "
        SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(vasWorkList, asRow1, colRack)) & "' "
        res = SendQuery(gLocal, SQL)
        
        Data_gubun = "1"
    ElseIf gReadBuf(0) = "1" Then
        Call vasWorkList_Send(asRow1)
        Data_gubun = "2"
        
                       SQL = "INSERT INTO SEND_RESULT"
        SQL = SQL & vbCrLf & "      (WORKNO, BARCODE ,SAMPLENO ,POS  "
        SQL = SQL & vbCrLf & "      ,EXAMCODE ,STARTDATE ,ENDDATE "
        SQL = SQL & vbCrLf & "      ,RESULTTIME ,RESULT ,SENDFLAG) "
        SQL = SQL & vbCrLf & "VALUES('" & Trim(GetText(vasID, 1, colSpecNo)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colBarcode)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colSampleNo)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colRack)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colPos)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colStartDate)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colEndDate)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colBact_Time)) & "' "
        SQL = SQL & vbCrLf & "      ,'" & Trim(GetText(vasID, 1, colBact_result)) & "' "
        SQL = SQL & vbCrLf & "      ,'0') "
        res = SendQuery(gLocal, SQL)
        
        If asEquipResult = "POSITIVE" Or asEquipResult = "SAMPLE ERR" Then
            Data_gubun = "0"                                                                '/POSITIVE 일때에는 서버에 저장을 하지 않음 (일단 Bact만)
                           SQL = "UPDATE SEND_RESULT SET "
            SQL = SQL & vbCrLf & "       PFLAG = 'P' "
            SQL = SQL & vbCrLf & " WHERE WORKNO = '" & Trim(GetText(vasID, 1, colSpecNo)) & "' "
            SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(GetText(vasID, 1, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(vasID, 1, colPos)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
        
        If asEquipResult = "POSITIVE" Or asEquipResult = "SAMPLE ERR" Then
            
                           SQL = "DELETE FROM SPEC_RESULT "
            SQL = SQL & vbCrLf & " WHERE WORKNO = '" & Trim(GetText(vasID, 1, colSpecNo)) & "' "
            SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(GetText(vasID, 1, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(vasID, 1, colRack)) & "' "
            res = SendQuery(gLocal, SQL)
        Else
                           SQL = "DELETE FROM SPEC_RESULT "
            SQL = SQL & vbCrLf & " WHERE WORKNO = '" & Trim(GetText(vasID, 1, colSpecNo)) & "' "
            SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(GetText(vasID, 1, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(vasID, 1, colRack)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
        
        
    End If
    
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



Private Sub MSComm1_OnComm()
    Dim i As Long
    Dim lsChar As String
       
    dtpToday = Date
    lsChar = MSComm1.Input
    
    Select Case lsChar
        Case chrENQ
            txtData = ""
            
            SaveData "[Rx]" & chrENQ
                    
            MSComm1.Output = chrACK
            SaveData "[Tx]" & chrACK
                    
        Case chrSTX     '자료수신 시작

            txtData.Text = ""
            
        Case chrETX
            txtData.Text = txtData.Text & lsChar

        Case chrLF
            
            txtData.Text = txtData.Text & lsChar
            SaveData "[Rx]" & chrSTX & txtData.Text
            
            BACT txtData
            
            MSComm1.Output = chrACK
            SaveData "[Tx]" & chrACK
            
        Case chrEOT     '자료수신 완료
            SaveData "[Rx]" & chrEOT
            txtData.Text = ""
               
        Case Else
            txtData.Text = txtData.Text & lsChar

    End Select
    
'    Select Case lsChar
'    Case chrENQ
'        SaveData "[RX]" & chrSTX & txtData.Text & lsChar
'
'        MSComm1.Output = chrACK
'        SaveData "[TX]" & chrACK
'
'        txtData.Text = ""
'    Case chrSTX
'        txtData.Text = ""
'        lsChar = ""
''    Case chrLF
''
''
''        MSComm1.Output = chrACK '
'    Case chrEOT
'        lsChar = ""
'        SaveData "[RX]" & lsChar
'
'    Case chrETX
'        SaveData "[RX]" & chrSTX & txtData.Text & lsChar & vbCrLf
'        Call BACT(txtData)
'
'        'SaveData "[RX]" & txtData.Text & lsChar
'        MSComm1.Output = chrACK
'        SaveData "[TX]" & chrACK
'
'    Case Else
'        txtData.Text = txtData.Text & lsChar
'    End Select

End Sub

Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "검사자를 입력해주세요."
    lblUser.Caption = InputBox(sMsg, "검사자 입력")

End Sub

Private Sub tmResultCheck_Timer()
    Dim res_Search  As ADODB.Recordset
    
    Dim sBarcode    As String
    Dim sSampleNo   As String
    
    Dim i           As Integer
    Dim lsRow       As Integer
    
    
    If gTimeCnt = 30 Then
        
        Call cmdIFClear_Click
        Call cmdLocload_Click
        
        SQL = "SELECT * FROM SPEC_RESULT "
        SQL = SQL & vbCrLf & "WHERE STARTDATE <=  '" & Format(DateAdd("h", -48, Now), "YYYYMMDDHHMMSS") & "' "
        SQL = SQL & vbCrLf & "  AND STARTDATE > '" & Format(DateAdd("h", -72, Now), "YYYYMMDDHHMMSS") & "' "
        SQL = SQL & vbCrLf & "  AND (ENDDATE IS NULL OR ENDDATE =  '')  "
        SQL = SQL & vbCrLf & "  AND SENDFLAG = '0' "
        Set res_Search = cn.Execute(SQL)
        Do Until res_Search.EOF
            sBarcode = res_Search.Fields("BARCODE") & ""
            sSampleNo = res_Search.Fields("SAMPLENO") & ""
            
            '///////// 바코드번호와 병번호로 해당 Row  찾기
            lsRow = -1
            For i = 1 To vasWorkList.DataRowCnt
                If Trim(GetText(vasWorkList, i, colBarcode)) = sBarcode And Trim(GetText(vasWorkList, i, colSampleNo)) = sSampleNo Then
                    lsRow = i
                    Exit For
                End If
            Next i
            
            SetText vasWorkList, "No growth for 2 Days", lsRow, colBact_result         '검사결과
            Call Save_Local_One_MI(lsRow, "No growth for 2 Days", Now)
            Call Insert_Data(lsRow, "2", Data_gubun)
            
            '/새로추가했는데 한번 테스트 해봅시다.
            'Call Insert_Data_SE_MIDDLE(lsRow, Result_gubun, Data_gubun)
            '/Call Insert_Data_SE(lsRow, "2", Data_gubun)
            
            res_Search.MoveNext
        Loop
        
        
        gTimeCnt = 0
    Else
        gTimeCnt = gTimeCnt + 1
        StatusBar1.Panels.Item(1) = " 데이터 체크중 (" & gTimeCnt & ") "
    End If
End Sub

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasID.Col = 1
        vasID.Row = i
        If vasID.value = 0 Then
        vasID.value = 1
        Else
        vasID.value = 0
        End If
    Next i
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasID, Row, colPos)) & "' " & vbCrLf & _
          " AND EXAMDATE = '" & Format(dtpToday.value, "yyyymmdd") & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG "
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
End Sub

'Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Long
'    Dim lsID As String
'    Dim lsTime As String
'    Dim lsPid As String
'    Dim i As Integer
'
'    iRow = vasID.ActiveRow
'    If KeyCode = vbKeyDelete Then
'        If iRow < 1 Or iRow > vasID.DataRowCnt Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'        lsPid = Trim(GetText(vasID, iRow, colPID))
'
'        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              " AND PID = '" & lsPid & "' " & vbCrLf & _
'              " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'              " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'              " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasID, iRow, iRow
'        vasRes.MaxRows = 0
'    ElseIf KeyCode = 13 Then
'
'        Get_Sample_Info (iRow)
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO " & vbCrLf & _
'              "  FROM EQUIPEXAM " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " ORDER BY SEQNO "
'
'        res = db_select_Vas(gLocal, SQL, vasTemp)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
'                  "POSNO, PID, PNAME, " & vbCrLf & _
'                  "JUMIN, PSEX, PAGE, " & vbCrLf & _
'                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
'                  "SEQNO, RESULT, EXAMNAME, " & vbCrLf & _
'                  "SENDFLAG, Hospital, refflag) " & vbCrLf & _
'                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, iRow, colBarcode)) & "', '" & Trim(GetText(vasID, iRow, colRack)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colPos)) & "', '" & Trim(GetText(vasID, iRow, colPID)) & "', '" & Trim(GetText(vasID, iRow, colPName)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colJumin)) & "', '" & Trim(GetText(vasID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
'                  "'" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "', '" & Trim(GetText(vasID, 0, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 2)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasTemp, i, 4)) & "', '" & Trim(GetText(vasID, iRow, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 3)) & "', " & vbCrLf & _
'                  "'1', '" & Trim(GetText(vasID, iRow, colHospital)) & "', '" & Trim(GetText(vasID, iRow, colState + (i * 2))) & "')"
'                res = SendQuery(gLocal, SQL)
'            Next i
'
'            SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'                  " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                  " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
'                  " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
'                  " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'                  " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'                  " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'            res = SendQuery(gLocal, SQL)
'
'        ElseIf lsID = lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "UPDATE PAT_RES "
'                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasID, iRow, colState + (i * 2) - 1)) & "' "
'                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, iRow, colBarcode)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasTemp, i, 2)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasID, 0, colState + (i * 2) - 1)) & "' "
'                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' "
'                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' "
'                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'                res = SendQuery(gLocal, SQL)
'            Next i
'        End If
'        SetText vasID, "Result", gRow, colState
'
'    End If
'
'
'End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_Click colBarcode, lRow
    End If
End Sub

'Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
'    Dim sResDateTime As String
'    Dim sControl As String
'    Dim sLotNo As String
'
'    Dim sRefLow As String
'    Dim sRefHigh As String
'    Dim sRefFlag As String
'
'    Dim sCnt As String
'
'    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
'    'sControl = Trim(Left(asBarcode, 2))
'    'sLotNo = Trim(Mid(asBarcode, 3))
'    sControl = asBarcode
'    sRefFlag = ""
'
'    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and levelname = '" & sControl & "' " & vbCrLf & _
'          "  and equipcode = '" & asExamCode & "' "
'    res = db_select_Col(gLocal, SQL)
'    If res > 0 Then
'        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
'            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
'            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
'            If CCur(sRefHigh) < CCur(asRes2) Then
'                sRefFlag = "H"
'            End If
'            If CCur(sRefLow) > CCur(asRes2) Then
'                sRefFlag = "L"
'            End If
'        End If
'    End If
'
'    sCnt = ""
'    SQL = "Select count(*) from qc_res " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
'          "  and levelname = '" & sControl & "' " & vbCrLf & _
'          "  and equipcode = '" & asExamCode & "' "
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res <= 0 Then
'        SaveQuery SQL
'        db_RollBack gLocal
'        Exit Function
'    End If
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res <= 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'    If Not IsNumeric(sCnt) Then sCnt = "0"
'
'    If CInt(sCnt) > 0 Then
'        SQL = "delete from qc_res " & vbCrLf & _
'              "where equipno = '" & gEquip & "' " & vbCrLf & _
'              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
'              "  and levelname = '" & sControl & "' " & vbCrLf & _
'              "  and equipcode = '" & asExamCode & "' "
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            'db_RollBack gLocal
'            SaveQuery SQL
'            Exit Function
'        End If
'    End If
'    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
'          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        'db_RollBack gLocal
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function

Private Sub vasRID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasRID.Col = 1
        vasRID.Row = i
        If vasRID.value = 0 Then
        vasRID.value = 1
        Else
        vasRID.value = 0
        End If
    Next i
End Sub

Private Sub vasRID_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim lsID As String
'    Dim i As Integer
'
'    If Row < 1 Or Row > vasRID.DataRowCnt Then
'        Exit Sub
'    End If
'
'    lsID = Trim(GetText(vasRID, Row, colBarcode))
'    lblChangeBar.Caption = lsID
'    lblBarcode.Caption = lsID
'    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
'    lblPname.Caption = Trim(GetText(vasRID, Row, colPName))
'    lblRrow.Caption = Row
'    'Local에서 불러오기
'    ClearSpread vasRRes
'
'    '장비코드, 검사코드, 검사명, 결과, 순번
'    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, EQUIPRESULT " & vbCrLf & _
'          "FROM PAT_RES " & vbCrLf & _
'          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
'          " AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
'          " AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
'          " AND EXAMDATE = '" & Format(dtpExamDate.value, "yyyymmdd") & "' " & vbCrLf & _
'          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG , EQUIPRESULT"
'
'    res = db_select_Vas(gLocal, SQL, vasRRes)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    vasRRes.MaxRows = vasRRes.DataRowCnt
'
'    For i = 1 To vasRRes.MaxRows
'        If Trim(GetText(vasRRes, i, colFLAG)) = "H" Then
'            SetForeColor vasRRes, i, i, colResult, colResult, 255, 0, 0
'        ElseIf Trim(GetText(vasRRes, i, colFLAG)) = "L" Then
'            SetForeColor vasRRes, i, i, colResult, colResult, 0, 255, 0
'        End If
'    Next i
End Sub

Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Long
'    Dim lsID As String
'    Dim lsTime As String
'    Dim lsPid As String
'    Dim i As Integer
'
'    iRow = vasRID.ActiveRow
'
'    If KeyCode = 13 Then
'
'        Get_Sample_InfoR (iRow)
'
'        lsID = Trim(GetText(vasRID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
'              "FROM PAT_RES " & vbCrLf & _
'              "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              "  AND EXAMDATE = '" & Trim(Format(dtpExamDate.value, "yyyymmdd")) & "' " & vbCrLf & _
'              "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
'
'        res = db_select_Vas(gLocal, SQL, vasTemp)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To vasRRes.DataRowCnt
'                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
'                  "POSNO, PID, PNAME, " & vbCrLf & _
'                  " PSEX, PAGE, " & vbCrLf & _
'                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
'                  "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, RECENO, EQUIPRESULT) " & vbCrLf & _
'                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasRID, iRow, colBarcode)) & "', '" & Trim(GetText(vasRID, iRow, colRack)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasRID, iRow, colPos)) & "', '" & Trim(GetText(vasRID, iRow, colPID)) & "', '" & Trim(GetText(vasRID, iRow, colPName)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasRID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
'                  "'" & Trim(Format(dtpExamDate.value, "yyyymmdd")) & "', '" & Trim(GetText(vasRRes, i, 1)) & "', '" & Trim(GetText(vasRRes, i, 2)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasRRes, i, 5)) & "', '" & Trim(GetText(vasRRes, i, 4)) & "', '" & Trim(GetText(vasRRes, i, 3)) & "', " & vbCrLf & _
'                  "'1', '" & Trim(GetText(vasRRes, i, colFLAG)) & "','" & Trim(GetText(vasRID, iRow, colSpecNo)) & "', '" & Trim(GetText(vasRRes, i, 7)) & "')"
'                res = SendQuery(gLocal, SQL)
'            Next i
'
'                SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                      " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
'                      " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
'                      " AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf & _
'                      " AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf & _
'                      " AND EXAMDATE = '" & Format(dtpExamDate.value, "yyyymmdd") & "' "
'                res = SendQuery(gLocal, SQL)
'        ElseIf lsID = lblChangeBar.Caption Then
'            For i = 1 To vasRRes.DataRowCnt
'                SQL = "UPDATE PAT_RES "
'                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasRRes, i, 4)) & "' "
'                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasRRes, i, 2)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, 1)) & "' "
'                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasRID, iRow, colPID)) & "' "
'                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' "
'                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpExamDate.value, "yyyymmdd") & "' "
'                res = SendQuery(gLocal, SQL)
'            Next i
'        End If
'    ElseIf KeyCode = vbKeyDelete Then
'        If iRow < 1 Or iRow > vasRID.DataRowCnt Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasRID, iRow, colBarcode))
'        lsPid = Trim(GetText(vasRID, iRow, colPID))
'
'        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              " AND PID = '" & lsPid & "' " & vbCrLf & _
'              " AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf & _
'              " AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf & _
'              " AND EXAMDATE = '" & Format(dtpExamDate.value, "yyyymmdd") & "' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasRID, iRow, iRow
'        vasRRes.MaxRows = 0
'
'    End If
End Sub

Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasRID.ActiveRow
        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
            
        vasRID_Click colBarcode, lRow
    End If
End Sub

Private Sub vasRRes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then: vasRID_KeyDown KeyCode, 0
End Sub

Private Sub vasWorkList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasWorkList.Col = 1
        vasWorkList.Row = i
        If vasWorkList.value = 0 Then
        vasWorkList.value = 1
        Else
        vasWorkList.value = 0
        End If
    Next i
End Sub

Function vasWorkList_Send(asRow As Long)
    Dim i As Integer
    Dim WorkList_Row As Long
    WorkList_Row = asRow
    
    With vasWorkList
        Call InsertRow(vasID, 1)

        For i = 1 To .MaxCols
            SetText vasID, GetText(vasWorkList, WorkList_Row, i), 1, i
        Next i
        .DeleteRows WorkList_Row, WorkList_Row
    End With
End Function
