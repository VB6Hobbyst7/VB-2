VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   " Sysmex XS-1000i Interface Program"
   ClientHeight    =   10635
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15225
   FillColor       =   &H0000FFFF&
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
   ScaleHeight     =   10635
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpread.vaSpread vasOrder 
      Height          =   5970
      Left            =   2610
      TabIndex        =   22
      Top             =   3150
      Visible         =   0   'False
      Width           =   5415
      _Version        =   393216
      _ExtentX        =   9551
      _ExtentY        =   10530
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
      MaxCols         =   10
      SpreadDesigner  =   "frmInterface.frx":0442
   End
   Begin Threed.SSPanel sspRes 
      Height          =   6945
      Left            =   8190
      TabIndex        =   44
      Top             =   2400
      Visible         =   0   'False
      Width           =   6405
      _Version        =   65536
      _ExtentX        =   11298
      _ExtentY        =   12250
      _StockProps     =   15
      Caption         =   "SSPanel3"
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCloseRes 
         Caption         =   "X"
         Height          =   315
         Left            =   6030
         TabIndex        =   46
         Top             =   0
         Width           =   375
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   6255
         Left            =   150
         TabIndex        =   45
         Top             =   540
         Width           =   6135
         _Version        =   393216
         _ExtentX        =   10821
         _ExtentY        =   11033
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   16
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":408F
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test2"
      Height          =   465
      Left            =   13890
      TabIndex        =   43
      Top             =   9630
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test1"
      Height          =   465
      Left            =   12840
      TabIndex        =   30
      Top             =   9630
      Width           =   1035
   End
   Begin VB.TextBox txtBuff 
      Height          =   465
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   9630
      Width           =   12555
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   780
      TabIndex        =   19
      Top             =   1830
      Width           =   195
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   7875
      Left            =   270
      TabIndex        =   18
      Top             =   1740
      Width           =   14655
      _Version        =   393216
      _ExtentX        =   25850
      _ExtentY        =   13891
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   3
      EditEnterAction =   2
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   30
      SpreadDesigner  =   "frmInterface.frx":810B
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "결과 출력"
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
      Height          =   555
      Left            =   240
      TabIndex        =   29
      Top             =   9450
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6300
      Top             =   2370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   4096
      InputLen        =   1
      RThreshold      =   1
      RTSEnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4230
      _Version        =   65536
      _ExtentX        =   7461
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "     XS-1000i  INTERFACE"
      ForeColor       =   16777215
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4470
         Picture         =   "frmInterface.frx":C712
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   1965
      Left            =   1200
      TabIndex        =   21
      Top             =   3720
      Visible         =   0   'False
      Width           =   5295
      _Version        =   393216
      _ExtentX        =   9340
      _ExtentY        =   3466
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
      SpreadDesigner  =   "frmInterface.frx":CFDC
   End
   Begin Threed.SSPanel sspMode 
      Height          =   675
      Left            =   8250
      TabIndex        =   17
      Top             =   150
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "전송모드"
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "새굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin VB.CommandButton cmd_Trans 
      Caption         =   "선택전송"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9390
      TabIndex        =   15
      Top             =   150
      Width           =   1125
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12810
      TabIndex        =   14
      Top             =   150
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   13950
      TabIndex        =   13
      Top             =   150
      Width           =   1125
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "코드설정"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11670
      TabIndex        =   12
      Top             =   150
      Width           =   1125
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "통신설정"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10530
      TabIndex        =   11
      Top             =   150
      Width           =   1125
   End
   Begin VB.TextBox txtUID 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5580
      TabIndex        =   9
      Top             =   510
      Width           =   1515
   End
   Begin VB.TextBox txtToday 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5580
      TabIndex        =   7
      Text            =   "2002/02/18"
      Top             =   150
      Width           =   1515
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   4455
      Left            =   4290
      TabIndex        =   5
      Top             =   2730
      Width           =   3555
      _Version        =   393216
      _ExtentX        =   6271
      _ExtentY        =   7858
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmInterface.frx":D250
   End
   Begin VB.TextBox txtAll 
      Height          =   375
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2610
      Width           =   2055
   End
   Begin VB.TextBox txtDate 
      Height          =   405
      Left            =   5190
      TabIndex        =   4
      Top             =   1950
      Width           =   2325
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   1695
      Left            =   9090
      TabIndex        =   24
      Top             =   4230
      Width           =   2475
      _Version        =   393216
      _ExtentX        =   4366
      _ExtentY        =   2990
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
      SpreadDesigner  =   "frmInterface.frx":117AA
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   1125
      Left            =   2490
      TabIndex        =   28
      Top             =   6150
      Visible         =   0   'False
      Width           =   1875
      _Version        =   393216
      _ExtentX        =   3307
      _ExtentY        =   1984
      _StockProps     =   64
      ColHeaderDisplay=   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   14
      MaxRows         =   50
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SpreadDesigner  =   "frmInterface.frx":11A1E
   End
   Begin FPSpread.vaSpread vasOrderTemp 
      Height          =   6600
      Left            =   960
      TabIndex        =   23
      Top             =   2100
      Visible         =   0   'False
      Width           =   4785
      _Version        =   393216
      _ExtentX        =   8440
      _ExtentY        =   11642
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
      MaxCols         =   2
      SpreadDesigner  =   "frmInterface.frx":12724
   End
   Begin VB.TextBox txtMsg 
      ForeColor       =   &H000000C0&
      Height          =   585
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   7110
      Visible         =   0   'False
      Width           =   6285
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   10260
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   661
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
            TextSave        =   "2009-09-28"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 1:26"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "메디메이트 ☎(051)462-1751"
            TextSave        =   "메디메이트 ☎(051)462-1751"
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
   Begin FPSpread.vaSpread vasOrderBuf 
      Height          =   6360
      Left            =   9180
      TabIndex        =   25
      Top             =   2310
      Visible         =   0   'False
      Width           =   5055
      _Version        =   393216
      _ExtentX        =   8916
      _ExtentY        =   11218
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
      MaxCols         =   2
      SpreadDesigner  =   "frmInterface.frx":162AE
   End
   Begin Threed.SSPanel chkMode 
      Height          =   675
      Left            =   7260
      TabIndex        =   38
      Top             =   150
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "Auto"
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin VB.Frame Frame1 
      Height          =   9300
      Left            =   120
      TabIndex        =   16
      Top             =   930
      Width           =   14970
      Begin VB.CommandButton cmdTest 
         Caption         =   "TEST"
         Height          =   525
         Left            =   13590
         TabIndex        =   42
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   525
         Left            =   8670
         TabIndex        =   41
         Top             =   210
         Width           =   4875
      End
      Begin VB.TextBox TxtBarcode 
         Height          =   315
         Left            =   4020
         TabIndex        =   37
         Top             =   300
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton cmdCall 
         Caption         =   "데이타 불러오기"
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
         Left            =   330
         TabIndex        =   35
         Top             =   240
         Width           =   2265
      End
      Begin VB.ComboBox cboGubun 
         Height          =   315
         ItemData        =   "frmInterface.frx":19E38
         Left            =   1380
         List            =   "frmInterface.frx":19E3A
         TabIndex        =   33
         Top             =   810
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CommandButton cmdListPrint 
         Caption         =   "리스트 출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4620
         TabIndex        =   27
         Top             =   8700
         Visible         =   0   'False
         Width           =   1845
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   60
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "검사종류"
         ForeColor       =   8388736
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.76
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "※ 처방일자를 확인하세요!"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7410
         TabIndex        =   31
         Top             =   9090
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label Label6 
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "※ 검사종류를 확인하세요!  WorkList작성하세요!"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3480
         TabIndex        =   36
         Top             =   855
         Visible         =   0   'False
         Width           =   5280
      End
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1500
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검 사 자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   39
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검 사 자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1380
      TabIndex        =   10
      Top             =   345
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   8
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   195
      Left            =   7545
      TabIndex        =   6
      Top             =   555
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pp"
      Visible         =   0   'False
      Begin VB.Menu subDel 
         Caption         =   "삭제"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'RTS, DTR = False

'vasID
Const colSeq = 0
Const colCheckBox = 1
Const colRack = 2
'Const ColPos = 3
Const ColBarcode = 3
Const colSampleNo = 4
Const colPID = 5
Const colPName = 6
Const colJumin = 7
Const colPSex = 8
Const colPAge = 9
Const colOCnt = 10
Const colRCnt = 11
Const colState = 12
Const colReceNo = 13
Const colReqDate = 14       '접수일자
Const colGubun = 15         '검사종류

'vasRes
Const colEquipExam = 3
Const colExamCode = 4       '검사코드
Const colSubCode = 5        '서브코드
Const colOcsCode = 6
Const colExamName = 7       '검사명
Const colResult = 8         '결과
Const colRCheck = 9         '판정
Const colPCheck = 10
Const colDCheck = 11
Const colUnit = 12
Const colRef = 13
Const colPanic = 14
Const colResult1 = 15
Const colSpcCOde = 16

Dim ConfirmData As String
Dim aCount

Public gRackNo As String        'Rack
Public gPosNo As String         'Pos

Public gBarCode As String
Public gPID As String           '챠트번호
Public gTestID As String        '장비코드
Public gSpecID As String        '검체번호
Public gResult As String
Public gResult1 As String

Public glRow As Long
Public gCount As String
Public gOCnt As Integer
Public gOCnt_1 As Integer
Public gRCnt As Integer
Public gCheck As String

Public gGubun As String         '검진/진료

Dim gsRack As String
Dim gsTube As String
Dim gsBarCode As String
Dim gsPID As String
Dim gsResDateTime As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String

'변수 추가
Dim plExamCode As String
Dim plRSCode As String
Dim plResult As String
Dim plDecision As String

Dim varExamList()

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub chkMode_Click()
    If chkMode.Caption = "Manual" Then
        chkMode.BackColor = &HFF0000
        chkMode.ForeColor = &HFFFFFF
        chkMode.Caption = "Auto"
        SaveSetting "MEDIMATE", "XS1000i", "SendMode", "1"
    Else
        chkMode.BackColor = &H8000&
        chkMode.ForeColor = &HFFFFFF
        chkMode.Caption = "Manual"
        SaveSetting "MEDIMATE", "XS1000i", "SendMode", "0"
    End If
End Sub

Private Sub cmd_Trans_Click()
'선택전송

    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    
    Dim sGubun As String
    
    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

    If (vasID.DataRowCnt < 1) Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    If txtUID = "" Then
        MsgBox "검사자를 입력하세요"
        Exit Sub
    End If
    
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        If vasID.Value = 1 Then
            liRet = -1
            
            If Trim(GetText(vasID, vasIDRow, ColBarcode)) <> "" Then
                liRet = Insert_Data(vasIDRow)
            End If
            
            If liRet = 1 Then
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", vasIDRow, colState
                
                vasID.Row = vasIDRow
                vasID.Col = 1
                
                vasID.Value = 0
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", vasIDRow, colState
            End If
        Else
        
        End If
    Next vasIDRow
    
'    If optGubun(0).Value = True Then
'        db_Commit gServer
'    ElseIf optGubun(1).Value = True Then
'        db_Commit gServer_1
'    End If
    
End Sub

Function ResultDecision(asNo As String, asResult As String, asExamCode As String) As Integer
    Dim lsRef As String
    Dim i As Long
    Dim j As Long
    
    Dim lsHigh, lsLow As String

    Dim iFloat As Integer
    
    Dim sExamCode As String
    Dim sRsCode As String
    
    ResultDecision = -1
    
    plExamCode = ""
    plResult = ""
    plDecision = ""
    
    If asNo = "" Then
        Exit Function
    End If

    lsHigh = ""
    lsLow = ""
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    SQL = "Select IN_CODE, RS_CODE from EXAM_TOC  " & vbCrLf & _
          "where RE_RCID = '" & Trim(asNo) & "' and IN_CODE = '" & Trim(asExamCode) & "' "
    res = db_select_Col(gServer, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    ElseIf res = 0 Then
        ResultDecision = 0
        Exit Function
    End If
    
    ResultDecision = 1
    
    If ResultDecision = 1 Then
        plExamCode = Trim(gReadBuf(0))
        plRSCode = Trim(gReadBuf(1))
        
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        SQL = "Select RS_HIGH, RS_LOW, RS_MIDDLE from RSLT_TCD where IN_CODE = '" & plExamCode & "' and RS_CODE = '" & plRSCode & "'"
        If db_select_Col(gServer, SQL) > 0 Then
            lsHigh = Trim(gReadBuf(0))
            lsLow = Trim(gReadBuf(1))
        
            If IsNumeric(lsLow) Then
                If CCur(lsLow) > CCur(asResult) Then
                    plDecision = "L"
                End If
            End If
            If IsNumeric(lsHigh) Then
                If CCur(lsHigh) < CCur(asResult) Then
                    plDecision = "H"
                End If
            End If
        End If
    End If
End Function

Function Insert_Data(ByVal argSpcRow As Integer) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow As Integer
    Dim jRow As Integer
    Dim i, j, k As Integer
    
    Dim lsBarcode As String
    Dim lsRCID As String
    Dim lsExamCode As String
    Dim lsRsCode As String
    Dim lsRet As String
    Dim lsDecision As String
    
    Dim sCnt As String
    Dim sSegRes As String
    Dim sLymRes As String
    Dim sMonoRes As String
    Dim sDiffRes As String
    
    Dim sParam As String
    Dim sParam1 As String
    Dim lsSpcCode As String
    
    Insert_Data = -1
    
    
    sParam = ""
    
    gCurDate = Format(Date, "yyyymmdd") & Format(Time, "hhnnss")
    
    lsBarcode = Trim(GetText(vasID, argSpcRow, ColBarcode))
    lsRCID = Trim(GetText(vasID, argSpcRow, colSampleNo))
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select b.equipcode, a.INTER_CODE, a.INTER_RESULT, a.INTER_RSCODE, a.INTER_MEMO " & vbCrLf & _
          " From pat_res a, equipexam b" & vbCrLf & _
          " Where a.INTER_DATE = '" & Format(Trim(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          " And a.INTER_GUBUN = '" & gEquip & "' " & vbCrLf & _
          " And a.INTER_SPECIMENID = '" & lsBarcode & "' " & vbCrLf & _
          " And a.INTER_CHAM_ID = '" & lsRCID & "' " & vbCrLf & _
          " And a.INTER_GUBUN = b.equipno " & vbCrLf & _
          " And a.INTER_CODE = b.examcode "
    res = db_select_Vas(gLocal, SQL, vasResTemp)

    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    vasSort vasResTemp, 2
    vasResTemp.MaxRows = vasResTemp.DataRowCnt
    
'    db_BeginTran gServer
    
    'sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13)
    sParam = "MSH|^~\&|HL7|MMS|||" & gCurDate & "||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1 " & Chr(13)
    sParam = sParam & "PID|||" & lsBarcode & "^" & gHPEquip & "^" & gUID & "^^^DefaultDomain^PI" & Chr(13)
    sParam = sParam & "PV1||E|" & gHPID & Chr(13)
    sParam = sParam & "OBR|1||||||" & gCurDate & Chr(13)
    
    sParam1 = sParam
    sParam = ""
    j = 0
    '서버로 결과값 저장하기
    i = 1
    For i = 1 To vasResTemp.DataRowCnt
        lsExamCode = Trim(GetText(vasResTemp, i, 2))
        lsRsCode = Trim(GetText(vasResTemp, i, 4))
        lsRet = Trim(GetText(vasResTemp, i, 3))
        lsSpcCode = Trim(GetText(vasResTemp, i, 5))
        lsDecision = ""
        
        If lsSpcCode = "" Then
            lsSpcCode = "ST"
        End If
        
        '검사코드,참고치 체크
        'res = ResultDecision(lsRCID, lsRet, lsExamCode)
        
        If lsExamCode <> "" And lsRet <> "" Then
            j = j + 1
            sParam = sParam & "OBX|" & CStr(j) & "|" & lsSpcCode & "|" & lsExamCode & "||" & lsRet & "||||||R" & Chr(13)
            
            'If j >= 21 Then Exit Do
        End If
        
    Next i
    
    'sParam = Chr(11) & sParam1 & sParam & Chr(12) & Chr(13)
    sParam = Chr(11) & sParam1 & sParam
    
    Save_Raw_Data sParam
    
    res = SendResult(sParam)
    
    
'    k = i
'    j = 0
'    sParam = ""
'    For i = k To vasResTemp.DataRowCnt
'        lsExamCode = Trim(GetText(vasResTemp, i, 2))
'        lsRsCode = Trim(GetText(vasResTemp, i, 4))
'        lsRet = Trim(GetText(vasResTemp, i, 3))
'        lsSpcCode = Trim(GetText(vasResTemp, i, 5))
'        lsDecision = ""
'
'        If lsSpcCode = "" Then
'            lsSpcCode = "ST"
'        End If
'
'        '검사코드,참고치 체크
'        'res = ResultDecision(lsRCID, lsRet, lsExamCode)
'
'        If lsExamCode <> "" And lsRet <> "" Then
'            j = j + 1
'            sParam = sParam & "OBX|" & CStr(j) & "|" & lsSpcCode & "|" & lsExamCode & "||" & lsRet & "||||||R" & Chr(13)
'
'        End If
'    Next i
'
'    If sParam <> "" Then
'        sParam = Chr(11) & sParam1 & sParam & Chr(12) & Chr(13)
'
'        Save_Raw_Data sParam
'
'        res = SendResult(sParam)
'    End If
    If res > 0 Then
'    db_Commit gServer
    
        SQL = " Update pat_res Set " & vbCrLf & _
              " INTER_SENDFLAG = '1' " & vbCrLf & _
              " Where INTER_GUBUN = '" & gEquip & "' " & vbCrLf & _
              " And INTER_SPECIMENID = '" & lsBarcode & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
              
        Insert_Data = 1
    Else
        Insert_Data = res
    End If
End Function


Private Sub cmdCall_Click()
    Dim i As Integer
    Dim j, k As Integer
    
    ClearSpread vasID
    
    SQL = "select INTER_RACKPOS, inter_SPECIMENID,inter_cham_id, inter_pid, inter_pname,'', inter_psex, " & vbCrLf & _
          "inter_page,'','', inter_sendflag from pat_res where inter_date = '" & Format(txtToday, "yyyymmdd") & "' " & vbCrLf & _
          "group by  INTER_RACKPOS, inter_SPECIMENID,inter_cham_id, inter_pid, inter_pname, inter_psex, inter_page, inter_sendflag"
    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    For i = 1 To vasID.DataRowCnt
        If GetText(vasID, i, 12) = "0" Then
            SetText vasID, "결과", i, 12
        ElseIf GetText(vasID, i, 12) = "1" Then
            SetText vasID, "전송", i, 12
        End If
        
        ClearSpread vasRes
        SQL = "select '', a.INTER_SPECIMENID, b.equipcode, a.INTER_CODE, b.rscode, '', b.examname, a.inter_result, a.INTER_SENDFLAG, b.seqno, a.INTER_MEMO " & vbCrLf & _
              "from pat_res a, equipexam b " & vbCrLf & _
              "where a.inter_gubun = '" & gEquip & "' " & vbCrLf & _
              "and a.inter_code = b.examcode " & vbCrLf & _
              "and a.inter_date = '" & Format(txtToday, "yyyymmdd") & "' " & vbCrLf & _
              "and a.inter_gubun = b.equipno " & vbCrLf & _
              "and a.INTER_SPECIMENID = '" & GetText(vasID, i, ColBarcode) & "' " & vbCrLf & _
              "group by a.INTER_SPECIMENID, b.equipcode, a.INTER_CODE, b.rscode, b.examname, a.inter_result, a.INTER_SENDFLAG, b.seqno, a.INTER_MEMO " & vbCrLf & _
              "order by b.seqno"
        res = db_select_Vas(gLocal, SQL, vasRes)
        For j = 1 To vasRes.DataRowCnt
            For k = 1 To UBound(gArr_Exam)
                If Trim(gArr_Exam(k, 1)) = Trim(GetText(vasRes, j, 3)) Then
                    vasID.SetText colState + k, i, Trim(GetText(vasRes, j, 8))
                    Exit For
                End If
            Next k
        Next j
    Next i
End Sub

Private Sub cmdClear_Click()
    Dim lRow As Long
    
    txtMsg.Text = ""
    
    'ClearSpread vasID
        
'    If ChkAll.Value = 1 Then
        For lRow = 1 To vasID.DataRowCnt
            vasID.Row = lRow
            vasID.Col = 1

            If vasID.Value = 1 Then
                DeleteRow vasID, lRow, lRow
                lRow = lRow - 1
            End If
        Next lRow

        ChkAll.Value = 0
'    Else
'        vasID.Row = 1
'        vasID.Row2 = vasID.MaxRows
'        vasID.Col = 1
'        vasID.Col2 = vasID.MaxCols
'        vasID.BlockMode = True
'        vasID.BackColor = RGB(255, 255, 255)
'        vasID.Action = 3
'        vasID.BlockMode = False
'    End If

    vasActiveCell vasID, 1, colPID

    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 1
    vasRes.OperationMode = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
    End
End Sub

Private Sub cmdCloseRes_Click()
    sspRes.Visible = False
End Sub

Private Sub cmdConfig_Click()
    frmConfig.SSPanel_machine.Caption = "XS-1000i"
    frmConfig.Show 1
End Sub

Private Sub cmdListPrint_Click()

    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
        
On Error GoTo ErrGoto
 
    CommonDialog1.ShowPrinter
    
    If vasID.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    sCurDate = txtToday.Text
    
    vasID.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasID.PrintAbortMsg = "인쇄중 입니다 ..."
    vasID.PrintJobName = "VITROS Eci WorkList 출력"
    
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 울산인산병원 진단검사의학과"
    
    vasID.PrintHeader = sHead
    vasID.PrintFooter = sFoot

    vasID.PrintMarginTop = 680
    vasID.PrintMarginBottom = 680
'현재 SS가 비대칭으로 출력함
    vasID.PrintMarginLeft = 0
    vasID.PrintMarginRight = 0
    
    vasID.PrintColor = True
    vasID.PrintGrid = True
    
'Set printing range
    vasID.PrintType = 0  'SS_PRINT_ALL(default)

    vasID.PrintShadows = True

    vasID.Action = 13 'SS_ACTION_PRINT
    
ErrGoto:
    '사용자가 취소버튼을 눌렀습니다.
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
    Dim iRow As Integer
    Dim jRow As Integer
    Dim kRow As Integer
    
    '환자정보관련
    Dim sRack As String
    Dim sPos As String
    Dim sSampleNo As String
    Dim sPID As String
    Dim sPName As String
    Dim sPSex As String
    Dim sPAge As String
    
    Dim sExamName As String

    Dim sHead As String
    Dim sHead1 As String    '의뢰시간
    Dim sFoot As String
    Dim sSlip As String
    Dim sCurDate As String
    Dim sExamDate As String
    Dim sTitle As String
    Dim PageCnt As Integer

On Error GoTo ErrGoto
    
    CommonDialog1.ShowPrinter
    
    PageCnt = vasPrint.PrintPageCount
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
        
        If vasID.Value = 1 Then
            'sExamDate = Trim(txtToday.Text)
        
            sCurDate = Format(CDate(Date), "yyyy/mm/dd") & "   " & Format(CDate(Time), "hh:mm:dd")
            
            sRack = Trim(GetText(vasID, iRow, colRack))
            'sPos = Trim(GetText(vasID, iRow, ColPos))
            sSampleNo = Trim(GetText(vasID, iRow, colSampleNo))
            
            sPID = Trim(GetText(vasID, iRow, colPID))
            sPName = Trim(GetText(vasID, iRow, colPName))
            sPSex = Trim(GetText(vasID, iRow, colPSex))
            sPAge = Trim(GetText(vasID, iRow, colPAge))
    
            '보고일자
            SetText vasPrint, Format(CDate(Date), "yyyy-mm-dd"), 20, 7
            SetText vasPrint, Format(CDate(Time), "hh:mm:ss"), 20, 8

            
            sTitle = "혈액학검사"
        
            vasPrint.PrintOrientation = 2
    
            vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
            
'            sHead = "/fn""궁서체"" /fz""13"" /fb1 /fi0 /fu0 " & "/l" & "                            " & "▣ " & sTitle & " ▣" & "/n/n/n " & _
'                        "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l" & "Rack No : " & sRack & "           " & "Pos No : " & sPos & "           " & "SampleNo : " & sSampleNo & "/n/n" & _
'                        "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l" & "병록번호: " & sPID & "           " & " 환자성명: " & sPName & "           " & "성별/나이: " & sPSex & "/" & sPAge & "        " & "진료과:" & "" & "          " & "병동:" & "" & "/n/n"

            sHead = "/fn""궁서체"" /fz""13"" /fb1 /fi0 /fu0 " & "/l" & "                            " & "▣ " & sTitle & " ▣" & "/n/n/n " & _
                        "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l" & "Sample No : " & sSampleNo & "/n/n" & _
                        "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l" & "병록번호: " & sPID & "           " & " 환자성명: " & sPName & "           " & "성별/나이: " & sPSex & "/" & sPAge & "        " & "진료과:" & "" & "          " & "병동:" & "" & "/n/n"
                        
            sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 울산인산병원 진단검사의학과"
            
            vasPrint.PrintHeader = sHead
            vasPrint.PrintFooter = sFoot
            
            '검사결과
            vasID_Click 2, iRow

            vasID.Row = iRow
            vasID.Col = 1
            vasID.Value = 0
            
            vasPrint.PrintOrientation = PrintOrientationPortrait
            
            vasPrint.PrintMarginTop = 0
            vasPrint.PrintMarginBottom = 680
            
            '현재 SS가 비대칭으로 출력함
            vasPrint.PrintMarginLeft = 0
            vasPrint.PrintMarginRight = 0
            
            vasPrint.PrintColor = True
            vasPrint.PrintGrid = True
            
            'vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
            
            '원하는 셀까지만 출력함
            vasPrint.Row = 1
            vasPrint.Row2 = vasPrint.DataRowCnt + 1
            vasPrint.Col = 1
            vasPrint.Col2 = 9
            vasPrint.PrintType = PrintTypeCellRange

            vasPrint.PrintShadows = True
        
            vasPrint.Action = 13 'SS_ACTION_PRINT
        End If
    Next iRow

ErrGoto:
    '사용자가 취소버튼을 눌렀습니다.
    Exit Sub

End Sub

Private Sub cmdSetup_Click()
    frmEquipExam.SSPanel1.Caption = "  XS-1000i 장비 코드 설정"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub cmdTest_Click()
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send
    Dim sParam
    Dim sRet
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    'oSOAP.MSSoapInit gAddr
    
    
    sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13)
    sParam = sParam & "PID|||^C1^A130100001^DefaultDomain^PI" & Chr(13)
    sParam = sParam & "PV1||E|A1301" & Chr(13)
    sParam = sParam & "OBR|1||||||1" & Chr(13)
    sParam = Chr(11) & sParam & Chr(12) & Chr(13)
    'Debug.Print sParam
    
    Save_Raw_Data "Worklist Param : " & vbCrLf & sParam
    
    sParam = makeB64(sParam)
    
    'MsgBox oSOAP.detail
    
    send = oSOAP.MdbOrderList(sParam)
    
    send = makeUB64(send)
    
    Save_Raw_Data "Worklist Return : " & vbCrLf & send
    
    Text1 = send
    
    Set oSOAP = Nothing

    DoEvents
    
    Exit Sub

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Sub

Private Sub cmdWorkList_Click()
    frmPatSear.Left = 0
    frmPatSear.Top = 0
    frmPatSear.Show
    
    gWorkFlag = 0
End Sub

Private Sub Command1_Click()
    XS1000i Left(txtBuff, 1), Mid(txtBuff, 2)
    
    txtBuff.Text = ""
End Sub



Private Sub Command2_Click()
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send
    Dim sParam
    Dim sRet
    
    Dim i
    
    i = vasID.DataRowCnt + 1
    vasID.MaxRows = i
    
    vasID.SetText ColBarcode, i, Trim(txtBuff)

    Get_Sample_Info i

    Exit Sub
    
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    
    
'<SB> <HL7 message> <EB> <CR>
'<SB> = Start Block character (0x0B) 11
'<EB> = End Block character (0x1C) 12
'<CR> = Carriage Return Character (0x0D) 13
    
    sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13) '& Chr(10)
    'sParam = sParam & Chr(11) & "PID|||^CBC^B080100043^DefaultDomain^PI" & Chr(12) & Chr(13)
    'sParam = sParam & Chr(11) & "PID|||^^B080100043^DefaultDomain^PI" & Chr(12) & Chr(13)
    'sParam = sParam & Chr(11) & "PID|||^AIDS^B080100043^DefaultDomain^PI" & Chr(12) & Chr(13)
    sParam = sParam & "PID|||200902270043^C1^B080100043^DefaultDomain^PI" & Chr(13) '& Chr(10)
    sParam = sParam & "PV1||E|B0801" & Chr(13) '& Chr(10)
    sParam = sParam & "OBR|1||||||1" & Chr(13) '& Chr(10)
    sParam = Chr(11) & sParam & Chr(12) & Chr(13)
    'Debug.Print sParam
    
    sParam = makeB64(sParam)
    
    'MsgBox oSOAP.detail
    
    'send = oSOAP.MdbOrderList(sParam)
    send = oSOAP.New_SelectOrder(sParam)
    
    'send = oSOAP.New_SelectOrder(sParam)
    sParam = makeUB64(sParam)
    send = makeUB64(send)
    
    txtBuff = send
    Debug.Print sParam
    Debug.Print send
    
    Set oSOAP = Nothing

    DoEvents
    
    Exit Sub

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Sub

Private Sub Form_Activate()
    txtMsg.Text = ""
    
    vasRes.OperationMode = 0
    
    vasActiveCell vasID, 1, colPID
'    vasID.SetFocus
    
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 1
End Sub

Private Sub Form_Load()
    Dim sDate As String
    '1. 화면 및 변수 초기화
    '2. 데이타베이스에 Connect 하기 - Local - Server
    '3. Ini 내용 불러오기    GetSetup
    '4. Comport Open

    Me.Left = 0
    Me.Top = 0
    
    'Clear
    txtMsg.Text = ""

    ClearSpread vasID
    vasID.MaxRows = 1
    vasRes.OperationMode = 0
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 1
    
    GetSetup    'ini에서 DB정보 불러오기
        
'    If Not Connect_Server Then
'        MsgBox "서버에 연결되지 않았습니다."
'        Exit Sub
'    End If
    
    If Not Connect_Local Then
        MsgBox "로컬에 연결되지 않았습니다."
        Exit Sub
    End If

    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
    
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    raw_data = ""
    
    txtToday = Format(CDate(GetDateFull), "yyyy/mm/dd")
    
    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", CDate(txtToday.Text), -30), "yyyymmdd")
    
    SQL = "Delete from pat_res where INTER_DATE < '" & sDate & "' "
    SendQuery gLocal, SQL
    '===================================================================
    
    '검사코드 가져오기
    GetExamCode
    
    SQL = " Select exampart From EquipExam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table EquipExam Add Column exampart Text(50) "
        res = SendQuery(gLocal, SQL)
    End If
    
    '오더구분
    SQL = " Select OrdGubun From EquipExam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table EquipExam Add Column OrdGubun Text(1) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select INTER_RSCODE From pat_res "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table pat_res Add Column INTER_RSCODE Text(2) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Alter Table pat_res Alter Column INTER_SPECIMENID text(20) "
    res = SendQuery(gLocal, SQL)
    
    '검사종류
'    With cboGubun
'        .AddItem "1 검진"
'        .AddItem "2 진료"
'    End With
    
'    cboGubun.AddItem " ", 0
'    cboGubun.ListIndex = 1
    If Trim(GetSetting("MEDIMATE", "XS1000i", "SendMode", "0")) = "1" Then
        chkMode.BackColor = &HFF0000
        chkMode.ForeColor = &HFFFFFF
        chkMode.Caption = "Auto"
        
    Else
        chkMode.BackColor = &H8000&
        chkMode.ForeColor = &HFFFFFF
        chkMode.Caption = "Manual"
        
    End If

    'MultiSelect Mode
    vasRes.OperationMode = 1
    
    vasID.RowHeight(-1) = 12
        
    glRow = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    DisConnect_Server
    DisConnect_Local
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    Dim j As Integer
    
    gAllExam = ""
    gAllOcsExam = ""
    
    ClearSpread vasTemp
    
    '장비코드, 검사코드, 검사명
'    SQL = "Select EquipCode, ExamCode, ExamName, subcode, ocscode From EquipExam where equipno = '" & gEquip & "' " & vbCrLf & _
'          " Order by EquipCode"
    SQL = "Select equipcode, examcode, examname, ordgubun " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  seqno "
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If res > 0 Then
        ReDim gArr_ExamCode(1 To vasTemp.DataRowCnt, 1 To 3)
    Else
        SaveQuery SQL
    End If
    
    For i = 1 To vasTemp.DataRowCnt
'        'If IsNumeric(Trim(GetText(vasTemp, i, 1))) = True Then
'            gArr_ExamCode(i, 1) = i
'
'            For j = 1 To 2
'                gArr_ExamCode(i, j + 1) = Trim(GetText(vasTemp, i, j))
'            Next j
'
'            If gAllExam = "" Then
'                gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
'            Else
'                gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
'            End If
'
'            If Trim(GetText(vasTemp, i, 5)) <> "" Then
'                If gAllOcsExam = "" Then
'                    gAllOcsExam = Trim(GetText(vasTemp, i, 5))
'                Else
'                    gAllOcsExam = gAllOcsExam & "," & Trim(GetText(vasTemp, i, 5))
'                End If
'            End If
'        'End If

        vasID.MaxCols = colState + vasTemp.DataRowCnt + 1
        
        If IsNumeric(Trim(GetText(vasTemp, i, 1))) = True Then
            gArr_Exam(i, 1) = Trim(GetText(vasTemp, i, 1))
            gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))      '검사코드
            gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))      '검사명
            gArr_Exam(i, 4) = Trim(GetText(vasTemp, i, 4))      '처방구분
            
            vasID.MaxCols = colState + i + 1
            vasID.SetText colState + i, 0, Trim(gArr_Exam(i, 3))
            vasID.ColWidth(colState + i) = 6
            Select Case Trim(GetText(vasTemp, i, 4))
            Case "C"    'CBC
                If gAllExam = "" Then
                    gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
                Else
                    gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
                End If
            Case "D"    'Diff
                If gAllExam = "" Then
                    gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
                Else
                    gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
                End If
            End Select
        End If
    Next i
    
End Sub

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    Dim sGubun As String
    Dim LineData As String

    lsChar = MSComm1.Input
    
    Select Case lsChar
    Case chrSTX
        txtBuff.Text = ""
        
    Case chrETX
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtBuff

        sGubun = Left(txtBuff, 1)
        LineData = Mid(txtBuff, 2)
        
        XS1000i sGubun, LineData
        
        gPreData = chrACK
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        
        If sGubun = "R" Then    'Order
            SendOrder
        End If
    
    Case chrACK
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        
        gOrderMessage = chrSTX & "S21" & Format(txtToday.Text, "yyyymmdd") & Space(3) & _
                        SetChar(Trim(gsBarCode), 15, 1, " ") & Space(2) & Trim(gsRack) & Trim(gsTube) & _
                        "1" & gsPID & Space(100) & Space(97) & _
                        chrETX
                        
        If gOrderMessage <> "" Then
            SendOrder
        End If
    Case Else
        txtBuff.Text = txtBuff.Text & lsChar
    End Select
End Sub

Sub SendOrder()
    If gOrderMessage <> "" Then
        gPreData = gOrderMessage
        gOrderMessage = ""
        
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gPreData
        MSComm1.Output = gPreData
    End If
End Sub

Sub XS1000i(asGubun As String, asData As String)
    Dim MyVar As String
    Dim MyRet As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j, k As Integer
    
    Dim iRow As Integer
    Dim lRow As Integer
    Dim liRet As Integer
    
    Dim lsDistinctII As String
    Dim lsInqueryMode As String
    Dim lsDate As String
    Dim lsTime As String
    Dim lsRack As String
    Dim lsTube As String
    Dim lsID As String
    Dim lsIDInfo As String
    Dim lsSeqNo As String
    
    Dim lsPID As String
    Dim lsPName As String
    Dim lsPsex As String
    
    Dim lsData As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim lsFlag As String
    Dim lsExamCode As String
    Dim lsExamCode1 As String
    
    Dim lsRsCode As String
    Dim lsExamName As String

    Dim sDate As String
    Dim iExamCnt As Integer
    Dim sResult As String
    Dim m, n As Integer
    Dim lsSpcCode As String
    
    
    sDate = Format(txtToday, "yyyymmdd")
    
    Select Case asGubun
    Case "D"    'Analysis Data Format
        lsDistinctII = Mid(asData, 1, 1)
        
        If lsDistinctII = "1" Then
            lsID = Trim(Mid(asData, 32, 15))
            If Len(lsID) = 10 Then lsID = "20" & lsID
            lsSeqNo = Trim(Mid(asData, 19, 10))
            If IsNumeric(lsSeqNo) Then
                lsSeqNo = CStr(CDbl(lsSeqNo))
            End If
'            If Len(lsID) >= 10 Then                      '바코드번호
                lsRack = Trim(Mid(asData, 61, 6))
                lsTube = Trim(Mid(asData, 67, 2))
'            Else
'                If UCase(Left(lsID, 3)) = "ERR" Then    '바코드리딩 에러
'                    lsRack = Trim(Mid(asData, 61, 6))
'                    lsTube = Trim(Mid(asData, 67, 2))
'
'                    lsID = CInt(lsRack) & lsTube
'                Else                                    '메뉴얼
'                    lsRack = Trim(Mid(lsID, 1, 1))
'                    lsTube = Trim(Mid(lsID, 2, 2))
'                End If
'            End If
            
            If lsRack = "" Then
                lsRack = "0"
            End If
            
            gsRack = lsRack
            gsTube = lsTube
                            
            glRow = -1
            
            lRow = ScanCol(vasID, Trim(lsID), ColBarcode, 1)
            glRow = lRow
            If lRow = -1 Then
                lRow = vasID.DataRowCnt + 1
                If lRow > vasID.MaxRows Then
                    vasID.MaxRows = lRow
                End If
                
                glRow = lRow
                
                SetText vasID, Trim(lsID), glRow, ColBarcode
                SetText vasID, Trim(lsID), glRow, vasID.MaxCols
                SetText vasID, lsSeqNo, glRow, colRack
                'SetText vasID, lsTube, glRow, ColPos
                SetText vasID, "수신완료", glRow, colState
                
                vasActiveCell vasID, glRow, ColBarcode
                
                ClearSpread vasRes, 1, 1
            End If
            
            '샘플의 환자 정보 가져오기
            'If Trim(GetText(vasID, glRow, colPID)) = "" Then
                Get_Sample_Info glRow
            'End If
            
        End If
        
        If lsDistinctII = "2" Then
            lsID = Trim(Mid(asData, 32, 15))
            If Len(lsID) = 10 Then lsID = "20" & lsID
            lsSeqNo = Trim(Mid(asData, 19, 10))
            If IsNumeric(lsSeqNo) Then
                lsSeqNo = CStr(CDbl(lsSeqNo))
            End If
'            If UCase(Left(lsID, 3)) = "ERR" Then
'                lsID = CInt(gsRack) & gsTube
'            End If
            
            glRow = -1
            
            lRow = ScanCol(vasID, Trim(lsID), ColBarcode, 1)
            glRow = lRow
            
            If lRow = -1 Then
                lRow = vasID.DataRowCnt + 1
                If lRow > vasID.MaxRows Then
                    vasID.MaxRows = lRow
                End If
                
                glRow = lRow
            End If
            
            SetText vasID, Trim(lsID), glRow, ColBarcode
            SetText vasID, lsSeqNo, glRow, colRack
            'SetText vasID, gsTube, glRow, ColPos
            SetText vasID, "수신완료", glRow, colState
            SetText vasID, Trim(lsID), glRow, vasID.MaxCols
            
            vasActiveCell vasID, glRow, ColBarcode
            
            ClearSpread vasRes, 1, 1
            
            
            '샘플의 환자 정보 가져오기
            'If Trim(GetText(vasID, glRow, colPID)) = "" Then
                Get_Sample_Info glRow
            'End If
            
            lsData = Mid(asData, 47)
            
            '검사코드만큼 Row의 갯수를 설정
            SQL = "Select count(ExamCode) From EquipExam" & vbCrLf & _
                  " Where Equipno = '" & gEquip & "' "
            res = db_select_Col(gLocal, SQL)
            vasRes.MaxRows = gReadBuf(0)
        
            j = 0
            For i = 1 To 32
                ClearSpread vasTemp
                
                SQL = "Select ExamCode,ExamName, rscode From EquipExam" & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And EquipCode = '" & Format(i, "0#") & "'"
                
                res = db_select_Col(gLocal, SQL)
                
                If res = 1 And gReadBuf(0) <> "" Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsRsCode = Trim(gReadBuf(2))
                    
                    lsSpcCode = ""
                    
                    For ii = 1 To vasOrder.DataRowCnt
                        If Trim(lsExamCode) = Trim(GetText(vasOrder, ii, 1)) Then
                            lsSpcCode = Trim(GetText(vasOrder, ii, 2))
                            Exit For
                        End If
                    Next ii
                    
                    j = j + 1
                    
                    Select Case i
                    Case "1", "14", "15", "16", "17", "18"
                        sResult = Trim(Mid(lsData, 1, 6))
                        sResult = Left(sResult, 5)
                    Case Else
                        sResult = Trim(Mid(lsData, 1, 5))
                        sResult = Left(sResult, 4)
                    End Select
                    
                    'If sResult = "*000" Then
                    '    sResult = "."
                    'End If

                    If IsNumeric(sResult) Then
                        sResult = SetResult(sResult, i)                     '소수점 처리
                    'Else
                    '    sResult = "0"
                    
                    SetText vasRes, lsID, j, 2                          '검체번호
                    SetText vasRes, Format(i, "0#"), j, colEquipExam    '장비코드
                    SetText vasRes, lsExamCode, j, colExamCode          '검사코드
                    SetText vasRes, lsRsCode, j, colSubCode             '결과코드
                    SetText vasRes, lsExamName, j, colExamName          '검사명
                    SetText vasRes, sResult, j, colResult               '검사결과
                    SetText vasRes, sResult, j, colResult1              '검사결과
                         
                    For k = 1 To UBound(gArr_Exam)
                        If Trim(gArr_Exam(k, 1)) = Trim(GetText(vasRes, j, colEquipExam)) Then
                            vasID.SetText colState + k, glRow, sResult
                            Exit For
                        End If
                    Next k
                    
                    Save_Local_One glRow, j, "A"
                    
                    End If
                End If
                
                Select Case i
                Case 1, 14, 15, 16, 17, 18
                    lsData = Mid(lsData, 7)
                Case Else
                    lsData = Mid(lsData, 6)
                End Select
            Next i
            
            If chkMode.Caption = "Auto" Then
                If Trim(GetText(vasID, glRow, ColBarcode)) <> "" Then
                    liRet = Insert_Data(glRow)
                End If
                
                If liRet = 1 Then
                    SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 202, 255, 112
                    SetText vasID, "완료", glRow, colState
                    
                    vasID.Row = glRow
                    vasID.Col = 1
                    
                    vasID.Value = 0
                Else
                    SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 255, 0, 0
                    SetText vasID, "실패", glRow, colState
                End If
            
            End If
            
        End If
        
    Case "R"    'Inquiry Data Format
        lRow = vasID.DataRowCnt + 1
        If lRow > vasID.MaxRows Then
            vasID.MaxRows = lRow
        End If
        
        gsBarCode = ""
        gsRack = ""
        gsTube = ""
        
        lsInqueryMode = Mid(asData, 1, 1)
        lsID = Trim(Mid(asData, 5, 15))
        If Len(lsID) = 10 Then lsID = "20" & lsID
        gsBarCode = lsID
        
        lsRack = Trim(Mid(asData, 22, 6))
        lsTube = Trim(Mid(asData, 28, 2))
        
'        If UCase(Left(lsID, 3)) = "ERR" Then
'            lsID = CInt(lsRack) & lsTube
'            gsBarCode = lsID
'        End If
        
        gsRack = lsRack
        gsTube = lsTube
        
        SetText vasID, lsID, lRow, ColBarcode
        SetText vasID, CInt(lsRack) & "-" & lsTube, lRow, colRack
        'SetText vasID, lsTube, lRow, ColPos
        SetText vasID, "검사", lRow, colState
        SetText vasID, Trim(lsID), glRow, vasID.MaxCols
        
        '샘플의 환자 정보 가져오기
        'if Trim(GetText(vasID, lRow, colPID)) = "" Then
            Get_Sample_Info lRow
        'End If
        
        lsPID = ""
        lsPsex = ""
        gsPID = ""
        
        lsPID = Trim(GetText(vasID, lRow, colPID))
        If lsPID = "" Then
            lsPID = Space(16)
        Else
            lsPID = lsPID & Space(16 - Len(lsPID))
        End If
        gsPID = lsPID
        
        lsPsex = Trim(GetText(vasID, lRow, colPSex))
        If lsPsex = "" Then
            lsPsex = Space(1)
        Else
            Select Case lsPsex
            Case "M"
                lsPsex = "1"
            Case "F"
                lsPsex = "2"
            End Select
        End If

        lsData = ""
        
        lsData = Make_Order(Trim(lsID), lRow)

        gOrderMessage = chrSTX & _
                        "S11" & sDate & "000" & SetChar(lsID, 15, 1, " ") & "00" & _
                        lsRack & lsTube & "1" & lsPID & _
                        Space(40) & lsPsex & Space(8) & Space(20) & _
                        Space(20) & Space(40) & _
                        Space(18) & lsData & _
                        chrETX
    End Select
End Sub

'Sub XS1000i(asGubun As String, asData As String)
'ClassB 형식

'    Dim MyVar As String
'    Dim MyRet As String
'
'    Dim i As Integer
'    Dim j As Integer
'
'    Dim iRow As Integer
'    Dim lRow As Integer
'    Dim liRet As Integer
'
'    Dim lsDistinctII As String
'    Dim lsInqueryMode As String
'    Dim lsDate As String
'    Dim lsTime As String
'    Dim lsRack As String
'    Dim lsTube As String
'    Dim lsID As String
'    Dim lsIDInfo As String
'    Dim lsPName As String
'
'    Dim lsData As String
'
'    Dim lsCode As String
'    Dim lsRt As String
'    Dim lsFlag As String
'    Dim lsExamCode As String
'    Dim lsRsCode As String
'    Dim lsExamName As String
'
'    Dim sDate As String
'    Dim iExamCnt As Integer
'    Dim sResult As String
'    Dim m As Integer
'    Dim n As Integer
'
'    sDate = Format(txtToday, "yyyymmdd")
'
'    Select Case asGubun
'    Case "D"    'Analysis Data Format
'        lsDistinctII = Mid(asData, 1, 1)
'
'        If lsDistinctII = "1" Then
'            lsID = Trim(Mid(asData, 32, 15))
'
'            If Len(lsID) = 12 Then                      '바코드번호
'                lsRack = Trim(Mid(asData, 61, 6))
'                lsTube = Trim(Mid(asData, 67, 2))
'            Else
'                If UCase(Left(lsID, 3)) = "ERR" Then    '바코드리딩 에러
'                    lsRack = Trim(Mid(asData, 61, 6))
'                    lsTube = Trim(Mid(asData, 67, 2))
'
'                    lsID = CInt(lsRack) & lsTube
'                Else                                    '메뉴얼
'                    lsRack = Trim(Mid(lsID, 1, 1))
'                    lsTube = Trim(Mid(lsID, 2, 2))
'                End If
'            End If
'
'            If lsRack = "" Then
'                lsRack = "0"
'            End If
'
'            gRack = lsRack
'            gTube = lsTube
'
'            lRow = ScanCol(vasID, Trim(lsID), ColBarcode, 1)
'            If lRow = -1 Then
'                lRow = vasID.DataRowCnt + 1
'                If lRow > vasID.MaxRows Then
'                    vasID.MaxRows = lRow
'                End If
'
'                SetText vasID, Trim(lsID), lRow, ColBarcode
'                SetText vasID, CInt(lsRack), lRow, colRack
'                SetText vasID, lsTube, lRow, colTube
'                SetText vasID, "수신완료", lRow, colState
'
'                vasActiveCell vasID, lRow, ColBarcode
'
'                ClearSpread vasRes, 1, 1
'            End If
'
'            '샘플의 환자 정보 가져오기
'            Get_Sample_Info lRow
'
'        End If
'
'        If lsDistinctII = "2" Then
'            lsID = Trim(Mid(asData, 32, 15))
'
'            If UCase(Left(lsID, 3)) = "ERR" Then
'                lsID = CInt(gRack) & gTube
'            End If
'
'            lRow = ScanCol(vasID, Trim(lsID), ColBarcode, 1)
'            If lRow = -1 Then
'                lRow = vasID.DataRowCnt + 1
'                If lRow > vasID.MaxRows Then
'                    vasID.MaxRows = lRow
'                End If
'
'                SetText vasID, Trim(lsID), lRow, ColBarcode
'                SetText vasID, CInt(gRack), lRow, colRack
'                SetText vasID, gTube, lRow, colTube
'                SetText vasID, "수신완료", lRow, colState
'
'                vasActiveCell vasID, lRow, ColBarcode
'
'                ClearSpread vasRes, 1, 1
'            End If
'
'            '샘플의 환자 정보 가져오기
'            Get_Sample_Info lRow
'
'            lsData = Mid(asData, 47)
'
'            '검사코드만큼 Row의 갯수를 설정
'            SQL = "Select count(ExamCode) From EquipExam" & vbCrLf & _
'                  " Where Equipno = '" & gEquip & "' "
'            res = db_select_Col(gLocal, SQL)
'            vasRes.MaxRows = gReadBuf(0)
'
'            j = 0
'            For i = 1 To 30
'
'                gReadBuf(0) = "0"
'                SQL = "Select ExamCode, ExamName From EquipExam" & vbCrLf & _
'                      " Where Equip = '" & gEquip & "' " & vbCrLf & _
'                      "  And EquipCode = '" & Format(i, "0#") & "'"
'                res = db_select_Col(gServer, SQL)
'
'                'If (res = 1) And (gReadBuf(0) <> "") Then
'                If res = 1 And gReadBuf(2) <> "" Then
'                    lsExamCode = Trim(gReadBuf(0))
'                    lsRsCode = Trim(gReadBuf(1))
'                    lsExamName = Trim(gReadBuf(2))
'
'                    j = j + 1
'
'                    Select Case i
'                    Case "1", "14", "15", "16", "17", "18"
'                        sResult = Trim(Mid(lsData, 1, 6))
'                        sResult = Left(sResult, 5)
'                    Case Else
'                        sResult = Trim(Mid(lsData, 1, 5))
'                        sResult = Left(sResult, 4)
'                    End Select
'
'                    If IsNumeric(sResult) Then
'                        sResult = SetResult(sResult, i)
'                        SetText vasRes, lsID, j, ColBarcode                 '검체번호
'                        SetText vasRes, Format(i, "0#"), j, colEquipExam    '장비코드
'                        SetText vasRes, lsExamCode, j, colExamCode         '검사코드
'                        SetText vasRes, lsExamName, j, colExamName         '검사명
'                        SetText vasRes, sResult, j, colResult               '검사결과
'                        SetText vasRes, sResult, j, colResult1              '검사결과
'
''                        If sRType = "1" Then
''                            QC_Result sSpecID, sExamCode, sResult, iRow
''                        Else
''                            Check_Result sSpecID, Trim(GetText(vasID, llRow, colPID)), sExamCode, sResult, j, Trim(GetText(vasID, llRow, colPSex))
''                        End If
'
'                        Save_Local_One lRow, j, "A"  ''''내일 할 부분
'                    Else
'                        '2004/06/09 이상은
'                        'SetText vasRes, "", j, colResult
'                        '================================================================
'                        '결과값 없어도 항목 디스플레이 되도록
'                        SetText vasRes, lsID, j, ColBarcode  '검체번호
'                        sResult = SetResult(sResult, i)
'                        SetText vasRes, Format(i, "0#"), j, colEquipExam '장비코드
'                        SetText vasRes, lsExamCode, j, colExamCode         '검사코드
'                        SetText vasRes, lsExamName, j, colExamName         '검사명
'                        SetText vasRes, "", j, colResult    '검사결과
'                        SetText vasRes, "", j, colResult1   '검사결과
'
'                        Save_Local_One lRow, j, "A"
'                        '================================================================
'                    End If
'                End If
'
'                Select Case i
'                Case 1, 14, 15, 16, 17, 18
'                    lsData = Mid(lsData, 7)
'                Case Else
'                    lsData = Mid(lsData, 6)
'                End Select
'            Next i
'            gReadBuf(0) = ""
'            '수신중========================================================
'            SetText vasID, "수신완료", llRow, colState
'            SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
'            '==============================================================
'        End If
'
'    Case "R"    'Inquiry Data Format
'        lRow = vasID.DataRowCnt + 1
'        If lRow > vasID.MaxRows Then
'            vasID.MaxRows = lRow
'        End If
'
'        gBarCode = ""
'        gRack = ""
'        gTube = ""
'
'        lsInqueryMode = Mid(asData, 1, 1)
'        lsID = Trim(Mid(asData, 5, 15))
'        gBarCode = lsID
'
'
'        lsRack = Trim(Mid(asData, 22, 6))
'        lsTube = Trim(Mid(asData, 28, 2))
'
'        If UCase(Left(lsID, 3)) = "ERR" Then
'            lsID = CInt(lsRack) & lsTube
'            gBarCode = lsID
'        End If
'
'        gRack = lsRack
'        gTube = lsTube
'
'        SetText vasID, lsID, lRow, ColBarcode
'        SetText vasID, CInt(lsRack), lRow, colRack
'        SetText vasID, lsTube, lRow, colTube
'        SetText vasID, "Order", lRow, colState
'
'        '샘플의 환자 정보 가져오기
'        Get_Sample_Info lRow
'
'        lsData = Make_Order(Trim(lsID), lRow)
'
'        gOrderMessage = chrSTX & _
'                        "S11" & sDate & "000" & SetChar(lsID, 15, 1, " ") & "00" & _
'                        lsRack & lsTube & "1" & Space(16) & _
'                        Space(40) & " " & Space(8) & Space(20) & _
'                        Space(20) & Space(40) & _
'                        Space(18) & lsData & _
'                        chrETX
'
'    End Select
'End Sub

Sub SYSMEXK4500(asData As String)
    Dim i As Integer
    Dim j As Integer
    Dim llRow As Integer
    Dim iRow As Integer
    Dim jRow As Integer
    
    Dim MyVar As String
    Dim MyRet As String
    Dim lsData As String
    Dim sEquip As String
    Dim sRack As String
    Dim sPos As String
    Dim sSeqNo As String
    Dim sSpecID As String
    Dim sSpecName As String
    Dim sSpecDate As String
    Dim sSpecTime As String
    Dim sBarcode As String
    
    Dim sExamCode As String
    Dim sSubCode As String
    Dim sOcsCode As String
    Dim sExamName As String
    Dim sResult As String
    Dim liRet As Integer
    Dim vasRow As Integer
    
    If Trim(asData) = "" Then   '받은 신호가 없으면 나가기
        Exit Sub
    End If
    
    sBarcode = ""
    sBarcode = TxtBarcode
    
    MyVar = Trim(asData)
'    If Len(MyVar) < 75 Then '받은 신호 오류 검사
'        Select Case MyVar
'        Case chrNACK
'        Case Chr(127)   'DEL
'            txtMsg.Text = txtMsg.Text & "잘못된 신호가 있습니다"
'        Case Chr(24)    'CAN
'            txtMsg.Text = txtMsg.Text & "The Work List is Full"
'        Case Else
'        End Select
'        Exit Sub
'    End If
    
    sSeqNo = CStr(CCur(Trim(Mid(MyVar, 16, 5))))
    If IsNumeric(Trim(Mid(MyVar, 22, 13))) Then
        sSpecID = CStr(CCur(Trim(Mid(MyVar, 22, 13))))
    Else
        sSpecID = Trim(Mid(MyVar, 22, 13))
        If InStr(1, sSpecID, "ERR") > 0 Then
            sSpecID = Mid(sSpecID, 4)
            
            If IsNumeric(sSpecID) Then
                sSpecID = CStr(CCur(sSpecID))
            End If
        End If
    End If
    sSeqNo = Format(sSpecID, "00#")
    
    MyRet = Trim(Mid(MyVar, 54))   '결과부분만
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    glRow = -1
'    For iRow = 1 To vasID.DataRowCnt
'        If Trim(GetText(vasID, iRow, ColBarcode)) = sSeqNo Then
'            glRow = iRow
'            Exit For
'        End If
'    Next
    glRow = vasID.DataRowCnt + 1
    vasID.MaxRows = glRow
'    For iRow = 1 To vasID.DataRowCnt
'        If Trim(GetText(vasID, iRow, 12)) = "" Then
'            glRow = iRow
'
'            Exit For
'        End If
'    Next
    SetText vasID, sSeqNo, glRow, ColBarcode
'    If gWorkFlag > 0 Then
'        For iRow = 1 To vasID.DataRowCnt
'            If Trim(GetText(vasID, iRow, 3)) = "" Then
'                glRow = iRow
'                gWorkFlag = gWorkFlag - 1
'                Exit For
'            End If
'        Next iRow
'    Else
'        glRow = -1
'
'        For iRow = 1 To vasID.DataRowCnt
'            If Trim(GetText(vasID, iRow, 3)) = sBarcode Then
'                glRow = iRow
'                Exit For
'            End If
'        Next iRow
'
'        If glRow = -1 Then
'            For iRow = 1 To vasID.DataRowCnt
'                If Trim(GetText(vasID, iRow, 3)) = "" Then
'                    glRow = iRow
'                    Exit For
'                End If
'            Next iRow
'        End If
'    End If
'
'    If glRow = -1 Then
'        llRow = vasID.DataRowCnt + 1
'        glRow = llRow
'        If llRow > vasID.MaxRows Then
'            vasID.MaxRows = llRow
'        End If
'    End If

    vasActiveCell vasID, glRow, colRack
    
'    SetText vasID, sBarcode, glRow, ColBarcode
    
    vasRow = glRow
    
    If Trim(GetText(vasID, glRow, colPID)) = "" And Len(Trim(GetText(vasID, glRow, ColBarcode))) = 12 Then
        Get_Sample_Info glRow
    End If
    
   
    ClearSpread vasRes, 1, 1
            
    vasRes.MaxRows = 25
    
    '결과 잘라 넣기
    j = 0
    For i = 1 To vasRes.MaxRows
        If i = 1 Then
            sResult = Trim(Mid(MyRet, 1, 6))
        Else
            sResult = Trim(Mid(MyRet, 1, 5))
        End If
        'sResult = Trim(Left(sResult, 4))    '5번째는 flagcode임
        
        If IsNumeric(sResult) Then
            sResult = SetResult(sResult, i)
        End If
        
        sExamCode = ""
        sSubCode = ""
'        sOcsCode = ""
        sExamName = ""
        
        SQL = "Select examcode, rscode, '', examname From EquipExam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  And EquipCode = '" & Format(i, "0#") & "'"
        res = db_select_Col(gLocal, SQL)
        If IsNumeric(sResult) Then
            If Format(i, "0#") = "01" Then
                sResult = sResult * 100
            ElseIf Format(i, "0#") = "02" Then
                sResult = CInt(sResult * 10) & "만"
            ElseIf Format(i, "0#") = "08" Then
                sResult = Format(CCur(sResult), "#0") & "만"
            Else
                sResult = Left(sResult, Len(sResult) - 1)
            End If
        Else
            sResult = sResult
        End If
            
        If (res = 1) Then
            j = j + 1
            
            sExamCode = Trim(gReadBuf(0))
            sSubCode = Trim(gReadBuf(1))
'            sOcsCode = Trim(gReadBuf(2))
            sExamName = Trim(gReadBuf(3))
            
            SetText vasRes, CStr(i), j, colEquipExam
            SetText vasRes, sExamCode, j, colExamCode
            SetText vasRes, sSubCode, j, colSubCode
'            SetText vasRes, sOcsCode, j, colOcsCode
            SetText vasRes, sExamName, j, colExamName
            SetText vasRes, sResult, j, colResult
            SetText vasRes, sResult, j, colResult1
            
            'Local에 결과 저장하기
            Save_Local_One_이전 glRow, j, "A", sSeqNo
            
        End If
        If Format(i, "0#") = "01" Then
            MyRet = Mid(MyRet, 7)
        Else
            MyRet = Mid(MyRet, 6)
        End If
        
    Next i
    
'    vasID.SetText colState, glRow, "수신완료"
    
    SetText vasID, "Result", glRow, 12
'    If chkMode.Caption = "Auto" Then
''            vasID.Col = 1
''            vasID.Row = glRow
''            vasID.Value = 1
'            liRet = -1
'            If Trim(GetText(vasID, VasRow, colSampleNo)) <> "" Then
'                liRet = Insert_Data(VasRow)
'            End If
'
'            If liRet = 1 Then
'                SetBackColor vasID, VasRow, VasRow, colCheckBox, colCheckBox, 202, 255, 112
'                SetText vasID, "완료", VasRow, colState
'
'            Else
'                SetBackColor vasID, VasRow, VasRow, colCheckBox, colCheckBox, 255, 0, 0
'                SetText vasID, "실패", VasRow, colState
'            End If
'        End If
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim sGubun As String
    Dim sDate As String
    Dim lsID, lsPID, lsPName, lsAcpDate, lsOrdDate, lsSpcCode, lsExamCode As String
    Dim sBarcode As String
    Dim sRet, sSeg As String
    Dim i, j, k As Integer
    
    Get_Sample_Info = -1
   
    ClearSpread vasOrder
    vasOrder.MaxRows = 50
    'vasOrder.Visible = True
    
    sBarcode = Trim(GetText(vasID, asRow, ColBarcode))
    If Len(sBarcode) = 10 Then
        sBarcode = "20" & sBarcode
        vasID.SetText ColBarcode, asRow, sBarcode
    End If
    
    sRet = Get_OrderList(sBarcode)
    
    txtBuff = sRet
    
    sRet = Mid(sRet, InStr(1, sRet, Chr(11)) + 1)
    If InStr(1, sRet, Chr(12)) > 0 Then
        sRet = Left(sRet, InStr(1, sRet, Chr(12)) - 1)
    End If
    
    i = InStr(1, sRet, Chr(13))
    Do While i > 0
        sSeg = Left(sRet, i - 1)
        sRet = Mid(sRet, i + 1)
        
        Select Case Left(sSeg, InStr(1, sSeg, Chr(124)) - 1)
        Case "MSH"
        Case "PID"
            'PID|||200902240068^이경희^830601^2^20090224^20090224^DefaultDomain^PI
            k = 0
            j = InStr(1, sSeg, Chr(124))
            Do While j > 0
                k = k + 1
                
                If k = 4 Then
                    sSeg = Left(sSeg, j - 1)
                    Exit Do
                End If
                sSeg = Mid(sSeg, j + 1)
                j = InStr(1, sSeg, Chr(124))
            Loop
            k = 0
            j = InStr(1, sSeg, "^")
            Do While j > 0
                k = k + 1
                
                Select Case k
                Case 1
                    lsID = Left(sSeg, j - 1)
                Case 2
                    lsPName = Left(sSeg, j - 1)
                    vasID.SetText colPName, asRow, lsPName
                Case 3
                    lsPID = Left(sSeg, j - 1)
                    vasID.SetText colPID, asRow, lsPID
                Case 4
                Case 5
                    lsAcpDate = Left(sSeg, j - 1)
                Case 6
                    lsOrdDate = Left(sSeg, j - 1)
                    
                    Exit Do
                End Select
                sSeg = Mid(sSeg, j + 1)
                j = InStr(1, sSeg, "^")
            Loop
            
        Case "PV1"
        Case "OBR"
        Case "OBX"
            'OBX|1|ST|WB2570||||||||R
            
            k = 0
            j = InStr(1, sSeg, Chr(124))
            Do While j > 0
                k = k + 1
                
                If k = 3 Then
                    lsSpcCode = Left(sSeg, j - 1)
                ElseIf k = 4 Then
                    lsExamCode = Left(sSeg, j - 1)
                    k = vasOrder.DataRowCnt + 1
                    vasOrder.SetText 1, k, lsExamCode
                    vasOrder.SetText 2, k, lsSpcCode
                    Exit Do
                End If
                sSeg = Mid(sSeg, j + 1)
                j = InStr(1, sSeg, Chr(124))
            Loop
        End Select
        
        i = InStr(1, sRet, Chr(13))
    Loop
        
    Get_Sample_Info = 1


    gReadBuf(0) = ""
    gReadBuf(1) = ""
End Function

Function Make_Order(argNo As String, argRow As Integer) As String
'Order Text 만들기...
    
    Dim sRetOrder(2) As String     'Order Text넣을 변수
    Dim sOrder As String
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim sExecDate As String
    Dim sExecDate1 As String
    
    Dim sExamCode As String     '검사코드
    Dim sEquipCode As String
    Dim sOrdGubun As String
    
    Dim iCnt_Ord As Integer    'Order 갯수

    Dim llRow As Long

    
    If argNo = "" Then
        Exit Function
    End If
    
    Make_Order = -1
    
    sExecDate = Format(Trim(txtToday), "yyyymmdd")
    sExecDate1 = Format(DateAdd("d", "1", Trim(txtToday)), "yyyymmdd")
    
    '환자정보 조회
    'res = Get_Sample_Info(argRow)
    If vasOrder.DataRowCnt < 1 Then
        Make_Order = 0
        
        SetText vasID, "없음", argRow, colState
        SetForeColor vasID, argRow, argRow, 255, 0, 0
        
'        'CBC+Diff
        sOrder = "11111111" & "1111111111" & _
                 "11111" & "00" & "0000001000" & "000000000000000"

        'CBC
'        sOrder = "11111111" & "0000000000" & _
'                 "11111" & "00" & "0000001000" & "000000000000000"
                 
        Make_Order = sOrder
        
        Exit Function
    End If
        
'    '검사항목 불러오기
'    ClearSpread vasCode
'
'    SQL = " Select IN_CODE " & vbCrLf & _
'          " From EXAM_TOC " & vbCrLf & _
'          " Where RE_RCID = '" & Trim(GetText(vasID, argRow, colSampleNo)) & "'" & vbCrLf & _
'          " And IN_CODE in (" & gAllExam & ") "
'    res = db_select_Vas(gServer, SQL, vasCode)
'    If res = -1 Then
'        SaveQuery SQL
'
'        SetText vasID, "Order 없음", argRow, colState
'        SetForeColor vasID, argRow, argRow, 255, 0, 0
'
'        Exit Function
'    End If
    
    For i = 1 To 2
        sRetOrder(i) = "0"
    Next i
    
    'Order
    k = 1
    Do While k <= vasOrder.DataRowCnt
        sExamCode = Trim(GetText(vasOrder, k, 1))
        
        For j = 1 To UBound(gArr_Exam())
            If sExamCode = gArr_Exam(j, 2) Then
                Select Case gArr_Exam(j, 4)
                Case "C"
                    sRetOrder(1) = "1"
                Case "D"
                    sRetOrder(2) = "1"
                End Select
                
                Exit For
            End If
        Next j
        
        k = k + 1
    Loop
        
    sOrder = ""
        
    For i = 1 To 2
        sOrder = sOrder & sRetOrder(i)
    Next i
    
    If sOrder <> "" And sOrder = "10" Then          'CBC
        sOrder = "11111111" & "0000000000" & _
                 "11111" & "00" & "0000001000" & "000000000000000"
    ElseIf sOrder <> "" And sOrder = "11" Then      'CBC+Diff
        sOrder = "11111111" & "1111111111" & _
                 "11111" & "00" & "0000001000" & "000000000000000"
    End If

    Make_Order = sOrder
    
End Function

Function SetResult(asResult As String, aiItem As Integer) As String
    Dim iFloat As Integer

    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
'    Case 1, 14, 15, 16, 17, 18, 30
'        iFloat = 2
    Case 8
        iFloat = 0
    Case 1, 2, 14, 15, 16, 17, 18, 24, 30
        iFloat = 2
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
            SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
        Else
            SetResult = CStr(CCur(Left(asResult, 4 - iFloat)) & "." & Right(asResult, iFloat))
        End If
    End If
    
'    If aiItem = 1 Then
'        SetResult = Format(SetResult, "#0.0")
'    End If
End Function

Private Sub sspMode_Click()
    If sspMode.Caption = "수정모드" Then
        sspMode.Caption = "전송모드"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "전송모드" Then
        sspMode.Caption = "수정모드"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 0
        
        vasActiveCell vasRes, 1, colResult
        vasRes.SetFocus
    End If

End Sub

Private Sub subDel_Click()
    Dim i As Long
    
    Dim sDisk As String
    Dim sPos As String
    Dim sSeq As String
    
    Dim sSampleNo As String
    
    If MsgBox("결과를 삭제하시겠습니까?", vbYesNo, "알림") = vbNo Then
        Exit Sub
    End If
    
    i = vasID.ActiveRow
    
    sSeq = Trim(GetText(vasID, i, 0))
    sSampleNo = Trim(GetText(vasID, i, colSampleNo))
    
    SQL = " Delete From pat_res " & CR & _
          " Where inter_date = '" & Format(txtToday.Text, "yyyymmdd") & "' " & CR & _
          " And inter_gubun = '" & gEquip & "' " & CR & _
          " And INTER_SPECIMENID = '" & Trim(GetText(vasID, i, ColBarcode)) & "' " & CR & _
          " And inter_cham_id = '" & sSampleNo & "' "
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasID.DeleteRows i, 1
    If i > vasID.DataRowCnt Then
        i = vasID.DataRowCnt
    End If
    
    vasActiveCell vasID, i, colPID
    vasID.SetFocus
    
End Sub

Private Sub TxtBarcode_keydown(KeyCode As Integer, Shift As Integer)
    Dim BarLow As Integer
    Dim i As Integer
    Dim barFlag As Integer

    If KeyCode = vbKeyReturn Then
        
        barFlag = -1
        If vasID.DataRowCnt = 0 Then
            SetText vasID, TxtBarcode, 1, 3
            BarLow = 1
            Get_Sample_Info BarLow
            
            barFlag = 1
        End If
        
        If barFlag = -1 Then
            For i = 1 To vasID.DataRowCnt
                If Trim(GetText(vasID, i, 3)) = Trim(TxtBarcode) Then
                    SetText vasID, TxtBarcode, i, 3
                    Get_Sample_Info BarLow
                    
                    
                    barFlag = 1
                    
                End If
            Next
            
        End If
        If barFlag = -1 Then
            BarLow = vasID.DataRowCnt + 1
            vasID.MaxRows = BarLow
'            vasID.MaxRows = vasID.DataRowCnt + 1
'            DeleteRow vasID, BarLow, BarLow
            SetText vasID, TxtBarcode, BarLow, 3
            
            Get_Sample_Info BarLow
           
            
        End If
        TxtBarcode.Text = ""
    Else

    End If

End Sub

Private Sub txtToday_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim sGubun As String
'
'    sGubun = ""
'    sGubun = Left(cboGubun.Text, 1)
'
'    If KeyCode = vbKeyReturn Then
'        ClearSpread vasID
'
'        SQL = " Select seqno, '', barcode, pid, pname, '', psex, page, '', '', '', '', recedate " & CR & _
'              " From pat_res " & CR & _
'              " Where examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & CR & _
'              " And equipno = '" & gEquip & "' "
'
'        If sGubun <> " " Then
'            SQL = SQL & CR & _
'                " And gubun = '" & sGubun & "' "
'        End If
'
'        SQL = SQL & CR & _
'                " Group By seqno, barcode, pid, pname, psex, page, recedate" & CR & _
'                " Order By seqno "
'
'        res = db_select_Vas(gLocal, SQL, vasID, , 2)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'    End If
'
'    vasID.RowHeight(-1) = 12
End Sub

Private Sub txtUID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        gUID = txtUID
        WritePrivateProfileString "DATABASE", "UserID", gUID, App.Path & "\interface.ini"
    End If
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    Dim j As Integer
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
    
    ClearSpread vasPrint, 1, 1
    
    lsID = Trim(GetText(vasID, Row, colPID))
    
    ClearSpread vasRes
    ClearSpread vasPrint
    
    vasRes.Row = 1
    vasRes.Row2 = vasRes.MaxRows
    vasRes.CellType = CellTypeNumber
    
    '지정한 순번대로 디스플레이 할 것
'    SQL = "select '', a.barcode, a.equipcode, a.examcode, b.subcode, b.ocscode, a.examname, a.result, a.refflag, b.seqno " & vbCrLf & _
'          "FROM pat_res a, equipexam b " & vbCrLf & _
'          "WHERE a.examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND a.equipno = '" & gEquip & "' " & vbCrLf & _
'          "  AND a.Barcode = '" & Trim(GetText(vasID, Row, ColBarcode)) & "' " & vbCrLf & _
'          "  And a.equipno = b.equipno " & vbCrLf & _
'          "  And a.examcode = b.examcode " & vbCrLf & _
'          "  Group by a.barcode, a.equipcode,  a.examcode, b.subcode, b.ocscode, a.examname, a.result, a.refflag, b.seqno " & vbCrLf & _
'          "  order by b.seqno "
    SQL = "select '', a.INTER_SPECIMENID, b.equipcode, a.INTER_CODE, b.rscode, '', b.examname, a.inter_result, a.INTER_SENDFLAG, b.seqno, a.INTER_MEMO " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.inter_gubun = '" & gEquip & "' " & vbCrLf & _
          "and a.inter_code = b.examcode " & vbCrLf & _
          "and a.inter_date = '" & Format(txtToday, "yyyymmdd") & "' " & vbCrLf & _
          "and a.inter_gubun = b.equipno " & vbCrLf & _
          "and a.INTER_SPECIMENID = '" & GetText(vasID, Row, ColBarcode) & "' " & vbCrLf & _
          "group by a.INTER_SPECIMENID, b.equipcode, a.INTER_CODE, b.rscode, b.examname, a.inter_result, a.INTER_SENDFLAG, b.seqno, a.INTER_MEMO " & vbCrLf & _
          "order by b.seqno"
          
    res = db_select_Vas(gLocal, SQL, vasRes)
    res = db_select_Vas(gLocal, SQL, vasPrint)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    vasSort vasRes, 10
    vasSort vasPrint, 10
    
    For i = 1 To vasRes.DataRowCnt
        '참조치
        Select Case Trim(GetText(vasRes, i, colRCheck))
        Case "H"
            vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(205, 55, 0)
        Case "L"
            vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(65, 105, 225)
        Case ""
             vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(0, 0, 0)
        End Select
    Next i

    For j = 1 To vasPrint.DataRowCnt
        '참조치
        Select Case Trim(GetText(vasPrint, j, colRCheck))
        Case "H"
            vasPrint.Row = j
            vasPrint.Col = 7
            vasPrint.ForeColor = RGB(205, 55, 0)
        Case "L"
            vasPrint.Row = j
            vasPrint.Col = 7
            vasPrint.ForeColor = RGB(65, 105, 225)
        Case ""
            vasPrint.Row = j
            vasPrint.Col = 7
            vasPrint.ForeColor = RGB(0, 0, 0)
        End Select
    Next j
End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    sExamDate = GetDateFull
'
'    sCnt = ""
'    SQL = "delete FROM pat_res " & vbCrLf & _
'          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'          "  AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
'          "  AND barcode = '" & Trim(GetText(vasID, asRow1, ColBarcode)) & "' "
'    SaveQuery SQL
'    res = SendQuery(gLocal, SQL)
''    If res = -1 Then
''        SaveQuery SQL
''        Exit Function
''    End If
'
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
''    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
''        SetText vasExam, "1900-01-01", asRow, colExamDate
''    End If
'
'    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, receno, pid, " & _
'          "pname, pjumin, page, psex, resdate, " & _
'          "equipcode, examcode, examtype, result, sendflag, examname, " & _
'          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'          "VALUES ('" & Format(CDate(txtToday.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'          "'" & Trim(GetText(vasID, asRow1, ColBarcode)) & "', '', " & _
'          "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colJumin)) & "', " & _
'          "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
'          "'" & sExamDate & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', '', " & _
'          "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colPCheck)) & "', " & _
'          "'" & Trim(GetText(vasRes, asRow2, colDCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colUnit)) & "', " & _
'          "'" & Trim(GetText(vasRes, asRow2, colRef)) & "', '" & Trim(GetText(vasRes, asRow2, colPanic)) & "') "
'
'    SaveQuery SQL
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
    Dim Cnt As Integer
    Dim lsSeq, lsCnt As String
    Dim lsRes1, lsRes2 As String
    
    Dim lsBarcode As String     '바코드번호
    Dim lsRCID  As String       '처방전번호
    Dim lsPID As String         '챠트번호
    Dim lsPName As String
    Dim lsPsex As String
    Dim lsPAge As String
    Dim lsORDate As String
    Dim lsRsCode As String
    Dim lsRackPos As String
    Dim lsSpcCode As String
    
    Dim sDate As String
    
    Cnt = -1
    
    sDate = Format(txtToday.Text, "yyyymmdd")
    
    lsRes1 = Trim(GetText(vasRes, asRow2, colResult))
    lsRes2 = Trim(GetText(vasRes, asRow2, colResult))
    
    lsRsCode = Trim(GetText(vasRes, asRow2, colSubCode))
    
    lsSpcCode = Trim(GetText(vasRes, asRow2, colSpcCOde))
    
    lsBarcode = Trim(GetText(vasID, asRow1, ColBarcode))
    lsRackPos = Trim(GetText(vasID, asRow1, colRack))
    
    lsRCID = Trim(GetText(vasID, asRow1, colSampleNo))
    lsPID = Trim(GetText(vasID, asRow1, colPID))
    lsPName = Trim(GetText(vasID, asRow1, colPName))
    lsPsex = Trim(GetText(vasID, asRow1, colPSex))
    lsPAge = Trim(GetText(vasID, asRow1, colPAge))
    If lsPAge = "" Then
        lsPAge = 0
    End If

    lsSeq = Trim(GetText(vasID, asRow1, colSeq))
    
'    If GetText(vasID, asRow1, colSampleNo) = "" Then
'        SQL = " delete from pat_res " & vbCrLf & _
'              " WHERE inter_date = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'              "  AND inter_gubun = '" & gEquip & "' " & vbCrLf & _
'              "  AND inter_code = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf & _
'              "  AND inter_seq = '" & Trim(GetText(vasID, asRow1, ColBarcode)) & "' "
'    Else

        SQL = " delete from pat_res " & vbCrLf & _
              " WHERE inter_date = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
              "  AND inter_gubun = '" & gEquip & "' " & vbCrLf & _
              "  AND inter_code = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf & _
              "  AND inter_specimenid = '" & Trim(GetText(vasID, asRow1, ColBarcode)) & "' "
'    End If
    res = SendQuery(gLocal, SQL)
    
'    SQL = "INSERT INTO pat_res (INTER_DATE, INTER_SEQ, INTER_GUBUN, INTER_CODE, " & _
'          "INTER_CNT, INTER_SPECIMENID, INTER_CHAM_ID, INTER_PID, INTER_PNAME, INTER_PSEX, INTER_PAGE, " & vbCrLf & _
'          "INTER_ORDT, INTER_RESULT, INTER_RESULT1, INTER_TIME, INTER_SENDFLAG, INTER_RSCODE) " & vbCrLf & _
'          "VALUES ('" & Trim(sDate) & "', 0, '" & gEquip & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & _
'          "'" & lsCnt & "', '" & lsBarcode & "', '" & lsRCID & "', '" & lsPID & "', '" & lsPName & "', '" & lsPsex & "', '" & lsPAge & "', " & vbCrLf & _
'          " '', '" & lsRes1 & "', '" & lsRes2 & "', '" & Format(Time, "hhmmss") & "', '0', '" & lsRsCode & "')"

    SQL = "INSERT INTO pat_res (INTER_DATE, INTER_GUBUN, INTER_CODE, " & _
          "INTER_SPECIMENID, INTER_CHAM_ID, INTER_PID, INTER_PNAME, INTER_PSEX, INTER_PAGE, " & vbCrLf & _
          "INTER_ORDT, INTER_RESULT, INTER_RESULT1, INTER_TIME, INTER_SENDFLAG, INTER_RSCODE, INTER_RACKPOS, INTER_MEMO) " & vbCrLf & _
          "VALUES ('" & Trim(sDate) & "', '" & gEquip & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & _
          " '" & lsBarcode & "', '" & lsRCID & "', '" & lsPID & "', '" & lsPName & "', '" & lsPsex & "', " & lsPAge & ", " & vbCrLf & _
          " '', '" & lsRes1 & "', '" & lsRes2 & "', '" & Format(Time, "hhmmss") & "', '0', '" & lsRsCode & "', '" & lsRackPos & "', '" & lsSpcCode & "' )"
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Save_Local_One_이전(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, asSeq As String)
    Dim Cnt As Integer
    Dim lsSeq, lsCnt As String
    Dim lsRes1, lsRes2 As String
    
    Dim lsBarcode As String     '바코드번호
    Dim lsRCID  As String       '처방전번호
    Dim lsPID As String         '챠트번호
    Dim lsPName As String
    Dim lsPsex As String
    Dim lsPAge As String
    Dim lsORDate As String
    Dim lsRsCode As String
    
    Dim sDate As String
    
    Cnt = -1
    
    sDate = Format(txtToday.Text, "yyyymmdd")
    
    lsRes1 = Trim(GetText(vasRes, asRow2, colResult))
    lsRes2 = Trim(GetText(vasRes, asRow2, colResult))
    
    lsRsCode = Trim(GetText(vasRes, asRow2, colSubCode))
    lsBarcode = Trim(GetText(vasID, asRow1, ColBarcode))
    
    lsRCID = Trim(GetText(vasID, asRow1, colSampleNo))
    lsPID = Trim(GetText(vasID, asRow1, colPID))
    lsPName = Trim(GetText(vasID, asRow1, colPName))
    lsPsex = Trim(GetText(vasID, asRow1, colPSex))
    lsPAge = Trim(GetText(vasID, asRow1, colPAge))
    If lsPAge = "" Then
        lsPAge = 0
    End If
    
'    lsORDate = Trim(GetText(vasID, asRow1, colOrDate))
    
    lsSeq = Trim(GetText(vasID, asRow1, colSeq))
    
    If GetText(vasID, asRow1, colSampleNo) = "" Then
        SQL = " delete from pat_res " & vbCrLf & _
              " WHERE inter_date = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
              "  AND inter_gubun = '" & gEquip & "' " & vbCrLf & _
              "  AND inter_code = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf & _
              "  AND inter_seq = '" & Trim(GetText(vasID, asRow1, ColBarcode)) & "' "
    Else

        SQL = " delete from pat_res " & vbCrLf & _
              " WHERE inter_date = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
              "  AND inter_gubun = '" & gEquip & "' " & vbCrLf & _
              "  AND inter_code = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf & _
              "  AND inter_specimenid = '" & Trim(GetText(vasID, asRow1, ColBarcode)) & "' "
    End If
    res = SendQuery(gLocal, SQL)
    
    

'    SQL = "SELECT INTER_SEQ, INTER_CNT from k4500_res  " & vbCrLf & _
'          "where INTER_DATE = '" & Trim(sDate) & "' " & vbCrLf & _
'          "  and INTER_SEQ >= 0 " & vbCrLf & _
'          "  and INTER_GUBUN = '" & gEquip & "' " & vbCrLf & _
'          "  and INTER_CODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' " & vbCrLf & _
'          "  AND INTER_SPECIMENID = '" & Trim(lsBarcode) & "' " & vbCrLf & _
'          "  AND INTER_CHAM_ID = '" & Trim(lsRCID) & "' "
'    res = db_select_Col(gLocal, SQL)
'    Select Case res
'    Case -1
'        SaveQuery SQL
'        Exit Function
'    Case 0
'        Cnt = 0
'    End Select
    
'    lsCnt = Trim(GetText(vasRes, asRow2, colSeq))
'    lsCnt = vasRes.DataRowCnt
'    If Cnt = 0 Then
        SQL = "INSERT INTO pat_res (INTER_DATE, INTER_SEQ, INTER_GUBUN, INTER_CODE, " & _
              "INTER_CNT, INTER_SPECIMENID, INTER_CHAM_ID, INTER_PID, INTER_PNAME, INTER_PSEX, INTER_PAGE, " & vbCrLf & _
              "INTER_ORDT, INTER_RESULT, INTER_RESULT1, INTER_TIME, INTER_SENDFLAG, INTER_RSCODE) " & vbCrLf & _
              "VALUES ('" & Trim(sDate) & "', " & asSeq & ", '" & gEquip & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & _
              "'" & lsCnt & "', '" & lsBarcode & "', '" & lsRCID & "', '" & lsPID & "', '" & lsPName & "', '" & lsPsex & "', '" & lsPAge & "', " & vbCrLf & _
              " '', '" & lsRes1 & "', '" & lsRes2 & "', '" & Format(Time, "hhmmss") & "', '0', '" & lsRsCode & "')"
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
'    Else
'        SQL = "UPDATE k4500_res " & vbCrLf & _
'              "set INTER_RESULT = '" & lsRes1 & "', " & vbCrLf & _
'              "    INTER_RESULT1 = '" & lsRes2 & "', " & vbCrLf & _
'              "    INTER_SENDFLAG = '0' " & vbCrLf & _
'              "where INTER_DATE = '" & Trim(sDate) & "' " & vbCrLf & _
'              "  and INTER_GUBUN = '" & gEquip & "' " & vbCrLf & _
'              "  and INTER_CODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' " & vbCrLf & _
'              "  AND INTER_SPECIMENID = '" & lsBarcode & "' " & vbCrLf & _
'              "  AND INTER_CHAM_ID = '" & lsRCID & "' "
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'    End If
End Function

'Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim i As Integer
'    Dim iRow As Integer
'    Dim lRow As Long
'
'    iRow = vasID.ActiveRow
'    lRow = iRow
'
'    If KeyCode = vbKeyReturn Then
'        '환자정보 가져오기
'        Get_Sample_Info lRow
'
'        '로컬 데이터 다시 저장하기
'        For i = 1 To vasRes.DataRowCnt
'            Save_Local_One lRow, i, "A"
'        Next i
'
'        '기존 데이터는 지우기
'        SQL = " Delete From pat_res " & CR & _
'              " Where examdate = '" & Format(txtToday.Text, "yyyymmdd") & "' " & CR & _
'              " And equipno = '" & gEquip & "' " & CR & _
'              " And seqno = '" & Trim(GetText(vasID, lRow, colSampleNo)) & "' " & CR & _
'              " And barcode = '' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'        End If
'
'        vasID.SetFocus
'    End If
'End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    sspRes.Visible = True
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sSpecID As String
    Dim llRow As Long
    Dim iRow As Long
    Dim i As Integer
    llRow = vasID.ActiveRow
    If KeyCode = vbKeyReturn Then
        vasID.MaxRows = vasID.DataRowCnt + 1
        sSpecID = Trim(GetText(vasID, llRow, ColBarcode))

        SetText vasID, UCase(sSpecID), llRow, ColBarcode
        
        '샘플의 환자 정보 가져오기
        Get_Sample_Info llRow
        
        For i = 1 To vasRes.DataRowCnt
            Save_Local_One llRow, i, "0"
        Next

'        If vasID.MaxRows = llRow Then
'            InsertRow vasID, llRow + 1
'            vasActiveCell vasID, llRow + 1, colPID
'            vasID.SetFocus
'        Else
'            vasActiveCell vasID, vasID.MaxRows, colPID
'            vasID.SetFocus
'        End If
        
        vasID_Click ColBarcode, llRow + 1
        
    End If

End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop
End Sub


Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Response, Help
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    If KeyCode = vbKeyReturn Then
        vasIDRow = vasID.ActiveRow
        If vasResCol = colResult And _
           Trim(GetText(vasRes, vasResRow, colResult)) <> Trim(GetText(vasRes, vasResRow, colResult1)) Then
            
            Response = MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!", Help, 100)
            If Response = vbYes Then
                '판정, 델타, 패닉 수정
'                Check_Result Trim(GetText(vasID, vasIDRow, colPID)), _
'                             Trim(GetText(vasID, vasIDRow, colPID)), _
'                             Trim(GetText(vasRes, vasResRow, colExamCode)), _
'                             Trim(GetText(vasRes, vasResRow, colResult)), _
'                             vasResRow, Trim(GetText(vasID, vasIDRow, colPSex))
                If GetText(vasID, vasIDRow, colSampleNo) = "" Then
                SQL = " Update k4500_res " & vbCrLf & _
                      " Set inter_result = '" & Trim(GetText(vasRes, vasResRow, colResult)) & "' " & vbCrLf & _
                      " WHERE inter_date = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND inter_gubun = '" & gEquip & "' " & vbCrLf & _
                      "  AND inter_code = '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "'" & vbCrLf & _
                      "  AND inter_seq = '" & Trim(GetText(vasID, vasIDRow, ColBarcode)) & "' "
                Else

                SQL = " Update k4500_res " & vbCrLf & _
                      " Set inter_result = '" & Trim(GetText(vasRes, vasResRow, colResult)) & "' " & vbCrLf & _
                      " WHERE inter_date = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND inter_gubun = '" & gEquip & "' " & vbCrLf & _
                      "  AND inter_code = '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "'" & vbCrLf & _
                      "  AND inter_specimenid = '" & Trim(GetText(vasID, vasIDRow, ColBarcode)) & "' "
                End If
                res = SendQuery(gLocal, SQL)
'                SaveQuery SQL
                
                
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                
            End If
        End If
        
    End If
End Sub

Public Function Check_Result(argBarCode As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '결과종류
    Dim sLow        As String       '참조치
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " Select ResClassCode, Res_M_Low, Res_M_High, Res_F_Low, Res_F_High, " & CR & _
          "        PanicValueGubun, Panic_M_Low, Panic_M_High, Panic_F_Low, Panic_F_High, " & CR & _
          "        DeltaValueGubun, DeltaLow, DeltaHigh, Point " & CR & _
          "From ExamMaster " & CR & _
          " Where HID = '115' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_select_Col(gServer, SQL)
    
    sResClassCode = Trim(gReadBuf(0))
    
    If sResClassCode = "1" Then '숫자
'참조치 체크
        sLow = ""
        sHigh = ""
        
        '숫자인지 아닌지 확인
        If IsNumeric(sDiffRet) = False Then
           MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
           Check_Result = -1
           Exit Function
        End If
        
        If IsNumeric(gReadBuf(13)) Then
            If CCur(gReadBuf(13)) > 0 Then
                sTmpStr = "#0."
                For i = 1 To CInt(gReadBuf(13))
                    sTmpStr = sTmpStr & "0"
                Next i
            Else
                sTmpStr = "#0"
            End If
            sDiffRet = Format(sDiffRet, sTmpStr)
            SetText vasRes, sDiffRet, argRow, colResult
            SetText vasRes, sDiffRet, argRow, colResult1
        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(1))
            sHigh = Trim(gReadBuf(2))
        Case "F"
            sLow = Trim(gReadBuf(3))
            sHigh = Trim(gReadBuf(4))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   '이상
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   '이하
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


'Panic 체크
        sPanicLow = ""
        sPanicHigh = ""
        
        sPanicGubun = Trim(gReadBuf(5))
        
        Select Case asSex
        Case "M", ""
            sPanicLow = Trim(gReadBuf(6))
            sPanicHigh = Trim(gReadBuf(7))
        Case "F"
            sPanicLow = Trim(gReadBuf(8))
            sPanicHigh = Trim(gReadBuf(9))
        End Select
        
        If sPanicGubun = "0" Then '상한/하한
            If sPanicLow = "" Or sPanicHigh = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) > CCur(sDiffRet) Then
                    PanicRet = "L"
                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
                    PanicRet = "H"
                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
                    PanicRet = ""
                End If
            End If
        ElseIf sPanicGubun = "1" Then 'percent
            If sPanicLow = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "L"
                    Else
                        PanicRet = ""
                    End If
                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "H"
                    Else
                        PanicRet = ""
                    End If
                Else
                    PanicRet = ""
                End If
            End If
        End If
        

'Delta 체크
        sDeltaLow = ""
        sDeltaHigh = ""
                
        sTmpRece1 = ""
        sTmpRet1 = ""
        sTmpRece2 = ""
        sTmpRet2 = ""
        PreResult = ""
        
        sMax_ReceNo = ""
'        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
        sReceNo = argBarCode
       
'2004/06/09 이상은
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBarCode & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result"
              
'2004/12/30 이상은 - 정렬부분 추가
        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
              " Where HID = '115' " & CR & _
              " And PID = '" & Trim(argPID) & "' " & CR & _
              " And ReceNo < '" & argBarCode & "' " & CR & _
              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
              " Group By Result " & CR & _
              " Order by 2 desc "
        res = db_select_Col(gServer, SQL)
              
        If res > 0 And gReadBuf(0) <> "" Then
            PreResult = gReadBuf(0)
        Else
            PreResult = ""
        End If
      
        If PreResult <> "" Then
          'PreResult = Trim(gReadBuf(0))
          sDeltaGubun = Trim(gReadBuf(10))
          
          sDeltaLow = Trim(gReadBuf(11))
          sDeltaHigh = Trim(gReadBuf(12))
          
            '이전결과에서 현재결과 뺀값이 sDiffRet임 (2002년 3월 15일 수정)
'            sDiffRet = PreResult - sDiffRet
            sDiffRet1 = sDiffRet - PreResult
            If sDeltaGubun = "0" Then '상한/하한
                If sDeltaLow = "" Or sDeltaHigh = "" Then
                    DeltaRet = ""
                Else
                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
                        DeltaRet = "L"
                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
                        DeltaRet = "H"
                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
                        DeltaRet = ""
                    End If
                End If
              
            ElseIf sDeltaGubun = "1" Then 'percent
               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
                  DeltaRet = ""
               Else
                   If sDeltaLow = "" Then
                        DeltaRet = ""
                    Else
                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
                            DeltaRet = "D"
                        Else
                            DeltaRet = ""
                        End If
                    End If
               End If
            End If
        End If
        
    ElseIf sResClassCode = "2" Then '문자
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 이상은 수정
'        '검사 항목 결과 참조 코드 체크에서 1 이상일 경우만 판정되게
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002년 3월 12일 +-에서 +/-로 수정
'        '2002년 5월 13일 NON-REACTIVE 판정 안돼서 추가
'        '2003년 2월 4일 이상은 수정 - 0-1로 참조치는 1이나 판정됨
'        '=================================================================================
'        '2002년 5월 13일 1 : 40 미만 판정 안됨
'        '2002년 6월 11일 (결과참조가 1:로 시작하면 판정체크 안하게 수정)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "양성", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "양성", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "양성"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End Select
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
    End If
    
    SetText vasRes, RefRet, argRow, colRCheck
    SetText vasRes, PanicRet, argRow, colPCheck
    SetText vasRes, DeltaRet, argRow, colDCheck
    

    '2002년 2월 15일 수정 (판정시 H, L 일때 글자 색깔 변화)
    '2002년 3월 14일 수정 (판정시 L일때는 파란색 그 외는 빨간색)
    If RefRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If PanicRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If DeltaRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    ElseIf DeltaRet = "D" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    Check_Result = 1

End Function

