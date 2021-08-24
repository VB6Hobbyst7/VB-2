VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   "APEX XML Editor   [씨젠의료재단 Seegene Medical Foundation]"
   ClientHeight    =   13455
   ClientLeft      =   330
   ClientTop       =   525
   ClientWidth     =   14730
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
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":08CA
   ScaleHeight     =   13455
   ScaleWidth      =   14730
   StartUpPosition =   1  '소유자 가운데
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   18
      Top             =   13050
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1
            MinWidth        =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   18523
            MinWidth        =   18523
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2018-06-28"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "오전 11:07"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   7005
      Left            =   14940
      TabIndex        =   0
      Top             =   1770
      Visible         =   0   'False
      Width           =   8325
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   270
         TabIndex        =   32
         Top             =   5850
         Visible         =   0   'False
         Width           =   6675
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
            TabIndex        =   34
            Top             =   150
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
            Left            =   120
            TabIndex        =   33
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   39
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   5130
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1905
            TabIndex        =   36
            Top             =   240
            Width           =   1485
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
            TabIndex        =   35
            Top             =   240
            Width           =   1380
         End
      End
      Begin VB.TextBox txtSN 
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
         Height          =   360
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   705
      End
      Begin VB.TextBox txtExamDate 
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
         Height          =   360
         Left            =   6270
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   1365
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   120
         TabIndex        =   6
         Top             =   1350
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
         SpreadDesigner  =   "frmInterface.frx":0B4D
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   5
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtTemp 
         Appearance      =   0  '평면
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   1050
         Width           =   675
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   3330
         TabIndex        =   1
         Top             =   3060
         Width           =   1665
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   975
         Left            =   120
         TabIndex        =   3
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
         SpreadDesigner  =   "frmInterface.frx":0DA9
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   2370
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
         SpreadDesigner  =   "frmInterface.frx":1005
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   5490
         TabIndex        =   40
         Top             =   210
         Visible         =   0   'False
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
         Format          =   129761280
         CurrentDate     =   40457
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
         Left            =   4560
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
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
         TabIndex        =   17
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
         Left            =   5580
         TabIndex        =   16
         Top             =   1680
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
         Left            =   6390
         TabIndex        =   15
         Top             =   1680
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
         TabIndex        =   7
         Top             =   4680
         Width           =   825
      End
   End
   Begin VB.PictureBox picMenu2 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   14670
      TabIndex        =   12
      Top             =   525
      Width           =   14730
      Begin VB.Frame fraXML 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   705
         Left            =   90
         TabIndex        =   19
         Top             =   -60
         Width           =   14145
         Begin VB.TextBox txtPart 
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "COMMON"
            Top             =   420
            Width           =   1665
         End
         Begin VB.CommandButton cmdResult 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Caption         =   "결과열기"
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
            Left            =   9120
            Style           =   1  '그래픽
            TabIndex        =   23
            Top             =   150
            Width           =   1575
         End
         Begin VB.CommandButton cmdIFClear 
            BackColor       =   &H00FFFFFF&
            Caption         =   "화면지움"
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
            Left            =   12420
            Style           =   1  '그래픽
            TabIndex        =   22
            Top             =   150
            Width           =   1665
         End
         Begin VB.CommandButton cmdXMLSave 
            BackColor       =   &H00C0FFC0&
            Caption         =   "XML저장"
            Enabled         =   0   'False
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
            Left            =   10770
            Style           =   1  '그래픽
            TabIndex        =   21
            Top             =   150
            Width           =   1575
         End
         Begin VB.TextBox txtBarcode 
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "12121212121212"
            Top             =   420
            Width           =   1725
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "파일경로 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   180
            Width           =   840
         End
         Begin VB.Label lblFileName 
            BackStyle       =   0  '투명
            Caption         =   "wqewqeqwewqeqwewqewqewqeqwewqewqeqwewqewqewqewqewqewqewqewqewe"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1170
            TabIndex        =   29
            Top             =   180
            Width           =   7875
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검체번호 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   450
            Width           =   840
         End
      End
   End
   Begin VB.Frame fraAPEX 
      BackColor       =   &H00FFFFFF&
      Height          =   11805
      Left            =   -30
      TabIndex        =   9
      Top             =   1260
      Width           =   14625
      Begin FPSpread.vaSpread vasID 
         Height          =   11475
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   5685
         _Version        =   393216
         _ExtentX        =   10028
         _ExtentY        =   20241
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
         MaxCols         =   17
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":1261
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   11475
         Left            =   5850
         TabIndex        =   10
         Top             =   180
         Width           =   8625
         _Version        =   393216
         _ExtentX        =   15214
         _ExtentY        =   20241
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
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
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   1
         SpreadDesigner  =   "frmInterface.frx":1F40
      End
   End
   Begin VB.PictureBox picMenu1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14670
      TabIndex        =   8
      Top             =   0
      Width           =   14730
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13020
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdCodeSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "코드설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Visible         =   0   'False
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
      Visible         =   0   'False
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "Send"
      Visible         =   0   'False
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

Dim varXMLData()    As Variant

Dim blnEditMode     As Boolean
Dim strEditNum      As String

Private Sub cmdCodeSet_Click()

    frmTestSet.Show vbModal
    GetExamCode

End Sub

Private Sub cmdExit_Click()

    If MsgBox("종료 하시겠습니까?", vbInformation + vbYesNo, "알림") = vbYes Then
        End
    End If

End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    StatusBar1.Panels(3).Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    vasID.RowHeight(-1) = 14
    
    txtSN.Text = ""
    txtExamDate.Text = ""
    txtBarcode.Text = ""
    txtPart.Text = ""
    
    lblFileName.Caption = ""
    
    cmdXMLSave.Enabled = False
    
    gRow = 0
    
End Sub

Private Sub cmdXMLSave_Click()
    Dim intRow      As Integer
    Dim STM         As ADODB.Stream
    Dim blnEditXml  As Boolean
    Dim strHeader   As String
    Dim strBody     As String
    Dim strFileNm   As String
    Dim strAssayNm  As String
    Dim i           As Integer
    Dim j           As Integer
    
    Dim varEditData As Variant
    Dim strEditData As String
    Dim lngLastNum  As Long
    Dim lngPrevNum  As Long
    Dim varEditNum  As Variant
    Dim strDestFile As String
    
    Dim strResult   As String
    Dim strClass    As String
    
    Dim varFileName As Variant
    
    blnEditXml = False
    
    strHeader = ""
    'strHeader = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    
    For i = 0 To 5
        strHeader = strHeader & Mid(varXMLData(i), InStr(varXMLData(i), "<")) & "</GROUP>"
    Next
    
    strBody = ""
    varEditNum = Split(strEditNum, "|")
                    
    For i = 1 To vasID.MaxRows
        'If vasID.ActiveRow <> i Then
            Call vasID_Click(colBARCODE, i)
        '    DoEvents
        'End If
        
        lngPrevNum = 0
        
        With vasRes
            For intRow = 1 To .MaxRows
                lngLastNum = CCur(GetText(vasRes, intRow, colSNO))
                
                If lngPrevNum <> 0 And lngPrevNum + 1 <> lngLastNum Then
                    strBody = strBody & varXMLData(lngPrevNum + 1) & "</GROUP>"
                End If
                
                blnEditXml = False
                For j = 0 To UBound(varEditNum)
                    If CStr(lngLastNum) = CStr(varEditNum(j)) Then
                        blnEditXml = True
                        Exit For
                    End If
                Next
                .Row = intRow
                .Col = colRESULT
                
                If .BackColor = vbYellow Or blnEditXml = True Then
                    
                    varEditData = ""
                    strEditData = ""
                    
                    varEditData = varXMLData(lngLastNum)
                    varEditData = Split(varEditData, "</PARAM>")
                    
                    For j = 0 To UBound(varEditData)
                        If varEditData(j) <> "" Then
                            If InStr(varEditData(j), "QntResult") > 0 Then
                                strEditData = strEditData & "<PARAM TYPE=""String"" ID=""QntResult"">"
                                
                                strResult = GetText(vasRes, intRow, colRESULT)
                                If InStr(strResult, "<") > 0 Then
                                    strResult = Replace(strResult, "<", "&lt;")
                                End If
                                If InStr(strResult, ">") > 0 Then
                                    strResult = Replace(strResult, ">", "&gt;")
                                End If
                                
                                strEditData = strEditData & strResult
                                strEditData = strEditData & "</PARAM>"
                            ElseIf InStr(varEditData(j), "Result") > 0 Then
                                strEditData = strEditData & "<PARAM TYPE=""String"" ID=""Result"">"
                                
                                strClass = GetText(vasRes, intRow, colCLASS)
                                strEditData = strEditData & strClass
                                strEditData = strEditData & "</PARAM>"
                            Else
                                strEditData = strEditData & varEditData(j) & "</PARAM>"
                            End If
                        End If
                    Next
                    
                    strBody = strBody & strEditData & "</GROUP>"
                Else
                    strBody = strBody & varXMLData(CCur(GetText(vasRes, intRow, colSNO))) & "</GROUP>"
                End If
                lngPrevNum = lngLastNum
            Next
            
            If i < vasID.MaxRows Then
                For j = lngLastNum + 1 To lngLastNum + 8
                    strBody = strBody & varXMLData(j) & "</GROUP>"
                Next
            End If
        End With
    Next
    
    For i = 0 To 5
        strBody = strBody & "</GROUP>"
    Next
    
    If strBody <> "" Then
        '## 기존에 파일이 있으면 삭제
        strFileNm = lblFileName.Caption
    
        If Dir$(strFileNm, vbNormal) <> "" Then
            If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
                MkDir (App.Path & "\Log")
            End If
            
            varFileName = Split(strFileNm, "\")
            
            '기존파일 백업
            strDestFile = App.Path & "\Log\" & varFileName(UBound(varFileName)) '& ".xml"
            '원본을 대상에 복사
            FileCopy strFileNm, strDestFile
            
            Kill strFileNm
        End If
        
         '## 파일오픈
        Set STM = New ADODB.Stream
        
        STM.Open
        STM.Type = adTypeText
        STM.Charset = "utf-8"
        STM.WriteText strHeader & strBody '& "</GROUP>" & vbCrLf
                    
        STM.SaveToFile strFileNm, adSaveCreateNotExist
        STM.Close
        Set STM = Nothing
        
        StatusBar1.Panels(3).Text = "XML 저장"
        
        strEditNum = ""
        'MsgBox "저장완료"
    End If
    
End Sub


Private Sub cmdResult_Click()
    Dim varXML      As Variant
    Dim strXmlName  As String
    Dim i As Integer
    Dim STM         As ADODB.Stream
    
    dtpToday = Date
    
    StatusBar1.Panels(3).Text = ""

    CommonDialog1.InitDir = gAssayNM.ResultPath
    CommonDialog1.Filter = "XML(*.xml)|*.xml"
    CommonDialog1.Action = 1
    
    If CommonDialog1.FileTitle = "" Then
        Exit Sub
    End If
    
    SQL = "delete from PATRESULT "
    Res = SendQuery(gLocal, SQL)
    
    
    Call cmdIFClear_Click

    Screen.MousePointer = 11
    
    DoEvents
    
    cmdXMLSave.Enabled = True
    
    Call cmdIFClear_Click
    
    strXmlName = Trim(CommonDialog1.FileName)
    
    lblFileName.Caption = strXmlName
    
    Call f_subSet_XMLWorkList(strXmlName)
    
'    vasID.Visible = False
'    vasRes.Visible = False
    
    Call EditRcvDataAPEX
    
'    vasID.Visible = True
'    vasRes.Visible = True
        
    vasRes.MaxRows = 0
    
    'Call vasID_Click(colBARCODE, 1)
    
    vasRes.MaxRows = 0
    
    Screen.MousePointer = 0
    
End Sub

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
    Dim varTmp1     As Variant
    Dim strTest     As String
    Dim strResult   As String
    Dim strClass    As String
    
    Dim STM         As ADODB.Stream
    Dim strBarcode  As String
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    j = 0
    blnAppend1 = False
    blnAppend2 = False
    
    '-- 오더파일명과 경로를 지정한다.
    strPath = strXML

    '## 파일오픈
    Set STM = New ADODB.Stream
    
    STM.Open
    STM.LoadFromFile strPath
    STM.Type = adTypeText
    STM.Charset = "utf-8"
    'STM.WriteText strHeader & strBody & "</GROUP>" & vbCrLf
    varTmp = STM.ReadText

    'STM.SaveToFile strFileNm, adSaveCreateNotExist
    STM.Close
    Set STM = Nothing
    
    varTmp = Split(varTmp, "</GROUP>")
    
    ReDim Preserve varXMLData(UBound(varTmp))
    
    Erase strRecvData
        
    frmProgress.Show
    frmProgress.ZOrder 0
    frmProgress.Xprog.Min = 1
    frmProgress.Xprog.Max = UBound(varTmp) + 1
    
    
    For i = 0 To UBound(varTmp)
        frmProgress.Xprog.Value = i + 1
        DoEvents
        
        varXMLData(i) = varTmp(i) '& "</GROUP>"
        If InStr(varTmp(i), """Patient""") > 0 Then
            strTmp = Mid(varTmp(i), InStr(varTmp(i), """Patient""") + 14)
            ReDim Preserve strRecvData(j)
            
            strBarcode = mGetP(strTmp, 1, """")
            strBarcode = Mid(strBarcode, 1, 12)
            
            strRecvData(j) = strRecvData(j) & "P|" & strBarcode
            j = j + 1
            blnAppend1 = True
        End If
        
        If InStr(varTmp(i), """Assay""") > 0 Then
            strTmp = Mid(varTmp(i), InStr(varTmp(i), "Assay") + 11)
            strTmp = mGetP(strTmp, 1, """")
            ReDim Preserve strRecvData(j)
            
            If gAssayNM.INHALANT = strTmp Then
                strRecvData(j) = strRecvData(j) & "O|INHALANT"
            ElseIf gAssayNM.FOOD = strTmp Then
                strRecvData(j) = strRecvData(j) & "O|FOOD"
            ElseIf gAssayNM.ATOPY = strTmp Then
                strRecvData(j) = strRecvData(j) & "O|ATOPY"
            ElseIf gAssayNM.COMMON = strTmp Then
                strRecvData(j) = strRecvData(j) & "O|COMMON"
            Else
                f_subSet_XMLWorkList = ""
                Exit Function
            End If
            
            j = j + 1
            blnAppend2 = True
        End If
        
        
        If InStr(varTmp(i), """Blot""") > 0 Then 'blnAppend1 = True And blnAppend2 = True And
            varTmp1 = Split(varTmp(i), "</PARAM>")
            strTest = ""
            strResult = ""
            For k = 0 To UBound(varTmp1)
                If InStr(varTmp1(k), "Code") > 0 Then
                    strTest = Mid(varTmp1(k), InStr(varTmp1(k), "Code") + 6)
                End If
                
                If strTest <> "" And InStr(varTmp1(k), """QntResult""") > 0 Then
                    strResult = Mid(varTmp1(k), InStr(varTmp1(k), """QntResult""") + 11)
                    strResult = Replace(strResult, "&lt;", "<")
                    strResult = Replace(strResult, "&gt;", ">")
                End If
            
                If strTest <> "" And strResult <> "" And InStr(varTmp1(k), """Result""") > 0 Then
                    strClass = Mid(varTmp1(k), InStr(varTmp1(k), """Result""") + 9)
                    ReDim Preserve strRecvData(j)
                    strRecvData(j) = strRecvData(j) & "R|" & strTest & "^" & strResult & "^" & strClass & "^" & CStr(i)
                    j = j + 1
                    strTest = ""
                    strResult = ""
                    strClass = ""
                End If
            
            Next
        End If
    Next
    
    '-- 프로그레스바 닫기
    Unload frmProgress
    Screen.MousePointer = 0

    Exit Function
        
ErrorTrap:
    
    Screen.MousePointer = 0
End Function


Private Sub lblclear_Click()
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    lblSaveSeq.Caption = ""
    lblExamDate.Caption = ""
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    'Me.Height = 11520
    Me.Height = picMenu1.Height + picMenu2.Height + fraAPEX.Height + StatusBar1.Height + 400
    Me.Width = fraAPEX.Left + fraAPEX.Width + 50
        
    cmdIFClear_Click
    lblclear_Click
    
    GetSetup
    
    If gSave = "True" Then
        MnTransAuto.Checked = True
        MnTransManual.Checked = False
    Else
        MnTransAuto.Checked = False
        MnTransManual.Checked = True
    End If
    
    If gIFMode = "Barcode" Then
        MnModeBarcode.Checked = True
        MnModeWorkList.Checked = False
    Else
        MnModeBarcode.Checked = False
        MnModeWorkList.Checked = True
    End If
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    'dtpStartDt = Date
    'dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -90), "yyyymmdd")
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    SQL = "delete from PATRESULT "
    
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
'    If lblUser.Caption = "" Then
'        Call picLogin_Click
'    End If
    
'    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    blnEditMode = False
   ' Call cmdSL_Click
    
    '-- test
'    vasID.MaxRows = 10
    
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
'            Debug.Print Trim(GetText(vasCode, i, j))
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)

    DisConnect_Local
    Unload Me
    End
End Sub

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
        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo And Trim(GetText(vasID, i, colSAVESEQ)) = mResult.RsltSeq Then
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
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    Call SetText(vasID, mResult.MnmNm, intRow, colINOUT)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAPEX()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strSNO       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    
    Dim strFIntBase  As String   '수신한 장비기준 검사명
    Dim strFResult   As String   '수신한 결과(정성)
    Dim strClass     As String
    
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
    Dim blntIge     As Boolean
    
    frmProgress.Show
    frmProgress.ZOrder 0
    frmProgress.Xprog.Min = 1
    frmProgress.Xprog.Max = UBound(strRecvData) + 1

    blntIge = False
    
    For intCnt = 0 To UBound(strRecvData)
    
        frmProgress.Xprog.Value = intCnt + 1
        DoEvents
        
        strRcvBuf = strRecvData(intCnt)
        'Debug.Print strRcvBuf
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Order
                strBarNo = Trim(mGetP(strRcvBuf, 2, "|"))
                mResult.BarNo = strBarNo
            Case "O"    '## Order
                strGubun = Trim(mGetP(strRcvBuf, 2, "|"))
                With mResult
                    .BarNo = strBarNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .MnmNm = strGubun
                End With
                                
                Call SetPatInfo(strBarNo)
                
                strState = "O"
                
                '-- 오른쪽 결과화면 초기화
                vasRes.MaxRows = 0
                
            Case "R"    '## Result
                strFIntBase = ""
                strFResult = ""
                
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                strClass = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 3, "^"))
                strSNO = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 4, "^"))
                strFIntBase = strIntBase
                
                
                If strIntBase = "tIgE" Then
                    If InStr(strResult, "2000") > 0 Then
                        strResult = "> 1000"
                        strFResult = strResult
                        blntIge = True
                    ElseIf InStr(strResult, "1000") > 0 Then
                        strResult = "> 1000"
                        strFResult = strResult
                    Else
                        strResult = Replace(strResult, ">", "")
                        strFResult = strResult
                    End If
                    cmdXMLSave.Enabled = True
                Else
                    strResult = Replace(strResult, ">", "")
                    strFResult = strResult
                End If
                
                'If strResult <> "" And Len(strIntBase) > 0 Then
                If strResult <> "" And Len(strIntBase) > 0 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
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
'                        For intCol = colState + 1 To vasID.MaxCols
'                            If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
'                                SetText vasID, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
                        
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMACHResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '순번
                        SetText vasRes, strSNO, lsResRow, colSNO                'SNO
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                        
                        If blntIge = True Then
                            vasRes.Row = 1
                            vasRes.Col = colRESULT
                            vasRes.BackColor = vbYellow
                            
                            strEditNum = strEditNum & GetText(vasRes, 1, colSNO) & "|"
                            
                            blntIge = False
                        End If
                        
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
                                    Exit For
                                End If
                            Next
                            
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                            SetText vasRes, lsEquipRes, lsResRow, colMACHResult     '장비결과
                            SetText vasRes, strResult, lsResRow, colRESULT          '결과
                            SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                            SetText vasRes, lsSeqNo, lsResRow, colSEQ               '순번
                            SetText vasRes, strSNO, lsResRow, colSNO                'SNO
                            '-- 로컬 저장
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 14
                
            Case "L"    '## Terminator
                        
        End Select
    Next

    '-- 프로그레스바 닫기
    Unload frmProgress
    
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
    
    SetResult = asResult

    Exit Function
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
          SQL = "SELECT resprec, reflow, refhigh " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE equipcode = '" & sEquipCode & "' " & vbCr
    SQL = SQL & "   AND EQUIPNO = '" & gEquip & "' " & vbCr
    
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
    Dim strGubuns   As String
    Dim varGubuns   As Variant
    Dim intCnt      As Integer
    Dim intCnt1     As Integer
    Dim strChannel  As String
    Dim strGubun    As String
    
    
    blnUpdate = False
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    'sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
    strChannel = Trim(GetText(vasRes, asRow2, colEQUIPCODE))
    strGubun = Trim(GetText(vasID, asRow1, colINOUT))
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'" & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, asRow1, colINOUT)) & "'"
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
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
    SQL = SQL & Trim(GetText(vasID, asRow1, colSAVESEQ))
    SQL = SQL & ",'" & sExamDate & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colHOSPDATE)) & "'"
    SQL = SQL & ",'" & gEquip & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCr
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colEXAMNAME)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colSEQ)) & "'" & vbCr
    SQL = SQL & ",''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colINOUT)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDISKNO)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colMACHResult)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colRESULT)) & "'" & vbCr
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colCLASS)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colSNO)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPID)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPNAME)) & "'" & vbCr
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPSEX)) & "'"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPAGE)) & "'"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",''" & vbCr
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser & "'"
    SQL = SQL & ",'')"
    
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
    
    'Local에서 불러오기
    ClearSpread vasRes
    
'    If blnEditMode = True Then
'        Call cmdXMLSave_Click
'    End If

    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
        
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    
    txtSN.Text = Trim(GetText(vasID, Row, colSAVESEQ))
    txtExamDate.Text = Mid(Trim(GetText(vasID, Row, colEXAMDATE)), 1, 8)
    txtBarcode.Text = lsID
    'txtPart.Text = "[" & Trim(GetText(vasID, Row, colINOUT)) & "]"
    txtPart.Text = Trim(GetText(vasID, Row, colINOUT))
    
    If lblSaveSeq.Caption = "" Then
        Exit Sub
    End If
    
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE, REFVALUE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, Row, colINOUT)) & "' " & vbCrLf
    SQL = SQL & " GROUP BY EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE, REFVALUE  "
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
                SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMACHResult
                SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
                SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colCLASS
                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSEQ
                SetText vasRes, Trim(RS.Fields("REFVALUE")) & "", .MaxRows, colSNO
                
'                If Trim(RS.Fields("REFFLAG")) = "H" Then
'                    .Row = .MaxRows
'                    .Col = colRESULT
'                    .ForeColor = vbRed
'                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
'                    .Row = .MaxRows
'                    .Col = colRESULT
'                    .ForeColor = vbBlue
'                End If
           
            End With
            RS.MoveNext
        Loop
    End If
    
    vasRes.RowHeight(-1) = 14
        
'    vasRes.Row = 1
'    vasRes.Col = colRESULT
'    vasRes.SetFocus
    
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
        blnModify = True
    
    End If
    
    
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub

        vasID_Click colBARCODE, lRow
    End If
End Sub

Private Function GetClass(dblResult As Double) As String
    
    GetClass = ""
    
    Select Case dblResult
        Case Is < 0.35:         GetClass = 0
        Case 0.35 To 0.7:       GetClass = 1
        Case 0.7 To 3.5:        GetClass = 2
        Case 3.5 To 17.5:       GetClass = 3
        Case 17.5 To 50:        GetClass = 4
        Case 50 To 100:         GetClass = 5
        Case Is > 100:          GetClass = 6
    End Select

End Function

Private Sub vasRes_KeyPress(KeyAscii As Integer)
    Dim strIntBase      As String
    Dim strResult       As String
    Dim strClass        As String
    Dim strBarcode      As String
    Dim strActiveNum    As String
    
    With vasRes
        
        If Trim(txtSN.Text) = "" Then
            Exit Sub
        End If
        
        If Trim(txtExamDate.Text) = "" Then
            Exit Sub
        End If
                
        If Trim(txtBarcode.Text) = "" Then
            Exit Sub
        End If
        
        strIntBase = Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE))
        
        If KeyAscii = 13 And .ActiveCol = colRESULT Then
            strResult = GetText(vasRes, .ActiveRow, colRESULT)
            strActiveNum = GetText(vasRes, .ActiveRow, colSNO)
            
            If strIntBase <> "tIgE" Then
                If IsNumeric(strResult) Then
                    strClass = GetClass(CDbl(strResult))
                    Call SetText(vasRes, strClass, .ActiveRow, colCLASS)
                Else
                    If Mid(strResult, 1, 1) = "<" Then
                        strClass = "0"
                        Call SetText(vasRes, strClass, .ActiveRow, colCLASS)
                    End If
                End If
            End If
            
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT " & vbCr
            SQL = SQL & "   SET RESULT      = '" & strResult & "'" & vbCr
            SQL = SQL & "     , REFFLAG     = '" & strClass & "'" & vbCr
            SQL = SQL & " WHERE BARCODE     = '" & Trim(txtBarcode.Text) & "' " & vbCr
            SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(txtExamDate.Text) & "' " & vbCr
            SQL = SQL & "   AND SAVESEQ   = " & Trim(txtSN.Text) & vbCr
            SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCr
            SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, .ActiveRow, colEXAMCODE)) & "' " & vbCr
            SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)) & "' " & vbCr

            Res = SendQuery(gLocal, SQL)

            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

            .Row = .ActiveRow
            .Col = .ActiveCol
            .BackColor = vbYellow
            
            strEditNum = strEditNum & strActiveNum & "|"
            cmdXMLSave.Enabled = True
            blnEditMode = True
        End If
    End With

End Sub


