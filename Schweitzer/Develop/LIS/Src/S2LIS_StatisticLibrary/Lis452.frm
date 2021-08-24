VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm452TurnAroundTime 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "TurnAroundTime"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   Tag             =   "Turn Around Time"
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkER 
      BackColor       =   &H00DBE6E6&
      Caption         =   "응급실만..."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   11340
      TabIndex        =   44
      Top             =   465
      Width           =   1365
   End
   Begin VB.OptionButton optStatFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "비응급"
      ForeColor       =   &H00C56152&
      Height          =   330
      Index           =   2
      Left            =   10350
      TabIndex        =   43
      Top             =   495
      Width           =   900
   End
   Begin VB.OptionButton optStatFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "응급"
      ForeColor       =   &H005E5EE8&
      Height          =   330
      Index           =   1
      Left            =   9600
      TabIndex        =   42
      Top             =   495
      Width           =   825
   End
   Begin VB.OptionButton optStatFg 
      BackColor       =   &H00FFF2EE&
      Caption         =   "전체"
      Height          =   330
      Index           =   0
      Left            =   8715
      Style           =   1  '그래픽
      TabIndex        =   41
      Top             =   480
      Width           =   825
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&T)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   31
      Tag             =   "132"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.ComboBox cboWA 
      BackColor       =   &H00DFF3DC&
      Height          =   300
      Left            =   150
      Style           =   2  '드롭다운 목록
      TabIndex        =   30
      Top             =   1800
      Width           =   2430
   End
   Begin VB.ListBox lstTest 
      BackColor       =   &H00E2FCFC&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      ItemData        =   "Lis452.frx":0000
      Left            =   150
      List            =   "Lis452.frx":0064
      TabIndex        =   29
      Top             =   2490
      Width           =   2580
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   28
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   27
      Tag             =   "132"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13020
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "132"
      Top             =   390
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabData 
      Height          =   315
      Left            =   3045
      TabIndex        =   32
      Top             =   7995
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   556
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Turn Around Time"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Overtime Item List"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   360
      Left            =   1560
      TabIndex        =   33
      Top             =   450
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   635
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
      Format          =   85131264
      UpDown          =   -1  'True
      CurrentDate     =   36342.5951388889
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   360
      Left            =   4980
      TabIndex        =   34
      Top             =   465
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   635
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
      Format          =   85131264
      UpDown          =   -1  'True
      CurrentDate     =   36342.5951388889
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   285
      Left            =   60
      TabIndex        =   35
      Top             =   45
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "조회조건"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   285
      Left            =   2985
      TabIndex        =   36
      Top             =   1095
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "조회 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   285
      Left            =   60
      TabIndex        =   37
      Top             =   1095
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "조회 리스트"
      Appearance      =   0
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   735
      Top             =   -165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   255
      Top             =   -165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   345
      Index           =   1
      Left            =   165
      TabIndex        =   45
      Top             =   450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "접  수  일"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   345
      Index           =   0
      Left            =   150
      TabIndex        =   46
      Top             =   1425
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "WorkArea"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   345
      Index           =   2
      Left            =   150
      TabIndex        =   47
      Top             =   2115
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Test Name"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTest 
      Height          =   345
      Left            =   3060
      TabIndex        =   48
      Top             =   1515
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Routine Chemistry"
      Appearance      =   0
   End
   Begin VB.Frame fraData 
      BackColor       =   &H00DBE6E6&
      Height          =   6090
      Index           =   1
      Left            =   3060
      TabIndex        =   9
      Top             =   1890
      Width           =   11265
      Begin VB.TextBox txtmin 
         Height          =   345
         Left            =   3810
         TabIndex        =   12
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txthour 
         Height          =   345
         Left            =   2610
         TabIndex        =   11
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txtday 
         Height          =   345
         Left            =   1545
         TabIndex        =   10
         Top             =   300
         Width           =   765
      End
      Begin FPSpread.vaSpread tblOver 
         Height          =   5160
         Left            =   165
         TabIndex        =   13
         Tag             =   "45215"
         Top             =   750
         Width           =   10830
         _Version        =   196608
         _ExtentX        =   19103
         _ExtentY        =   9102
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16252927
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   20
         OperationMode   =   2
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15724527
         ShadowDark      =   16777215
         ShadowText      =   0
         SpreadDesigner  =   "Lis452.frx":01F2
         VisibleCols     =   6
         VisibleRows     =   2
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   3
         Left            =   165
         TabIndex        =   53
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "기준소요시간"
         Appearance      =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "분"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4365
         TabIndex        =   16
         Tag             =   "15104"
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "시간"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3375
         TabIndex        =   15
         Tag             =   "15104"
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "일"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2370
         TabIndex        =   14
         Tag             =   "15104"
         Top             =   345
         Width           =   180
      End
   End
   Begin VB.Frame fraData 
      BackColor       =   &H00DBE6E6&
      Height          =   6090
      Index           =   0
      Left            =   3060
      TabIndex        =   1
      Top             =   1890
      Width           =   11280
      Begin VB.CommandButton cmdGraph 
         BackColor       =   &H00F4F0F2&
         Caption         =   "그래프보기(&G)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6315
         Style           =   1  '그래픽
         TabIndex        =   7
         Tag             =   "132"
         Top             =   300
         Width           =   1470
      End
      Begin VB.TextBox txtTestcnt 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   375
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1875
         Width           =   3405
      End
      Begin VB.TextBox txtOTime 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   375
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2745
         Width           =   3405
      End
      Begin VB.TextBox txtRTime 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   375
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3615
         Width           =   3405
      End
      Begin VB.TextBox txtVTime 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   375
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4455
         Width           =   3405
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "요일별 검색"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   1275
      End
      Begin FPSpread.vaSpread tblTab1 
         Height          =   5205
         Left            =   240
         TabIndex        =   8
         Tag             =   "45213"
         Top             =   810
         Width           =   7545
         _Version        =   196608
         _ExtentX        =   13309
         _ExtentY        =   9181
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridSolid       =   0   'False
         MaxCols         =   5
         MaxRows         =   20
         Protect         =   0   'False
         RowHeaderDisplay=   0
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis452.frx":0900
         VisibleCols     =   4
         VisibleRows     =   2
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   4
         Left            =   7830
         TabIndex        =   49
         Top             =   1500
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "총 검사 건수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   5
         Left            =   7830
         TabIndex        =   50
         Top             =   2370
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "평균  From Collection to Receive"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   6
         Left            =   7830
         TabIndex        =   51
         Top             =   3255
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "평균 From Receive to Verify"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   7
         Left            =   7830
         TabIndex        =   52
         Top             =   4095
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "평균 From Collection to Verify"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraData 
      BackColor       =   &H00DBE6E6&
      Height          =   6105
      Index           =   2
      Left            =   3060
      TabIndex        =   17
      Top             =   1890
      Width           =   11295
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   240
         Left            =   225
         TabIndex        =   21
         Top             =   3225
         Width           =   2985
         Begin VB.OptionButton optBase 
            BackColor       =   &H00DBE6E6&
            Caption         =   "시간"
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   24
            Tag             =   "45304"
            Top             =   15
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.OptionButton optBase 
            BackColor       =   &H00DBE6E6&
            Caption         =   "일"
            Height          =   300
            Index           =   0
            Left            =   585
            TabIndex        =   23
            Tag             =   "45305"
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton optBase 
            BackColor       =   &H00DBE6E6&
            Caption         =   "분"
            Height          =   300
            Index           =   2
            Left            =   1950
            TabIndex        =   22
            Tag             =   "45304"
            Top             =   15
            Width           =   690
         End
         Begin VB.Label Label5 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "기 준 :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   25
            Tag             =   "15104"
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00F4F0F2&
         Caption         =   "되돌아가기(&B)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9615
         Style           =   1  '그래픽
         TabIndex        =   20
         Tag             =   "132"
         Top             =   150
         Width           =   1470
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일자별"
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   19
         Tag             =   "45305"
         Top             =   270
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "요일별"
         Height          =   300
         Index           =   1
         Left            =   1215
         TabIndex        =   18
         Tag             =   "45304"
         Top             =   270
         Width           =   1245
      End
      Begin ChartfxLibCtl.ChartFX chart1 
         Height          =   2520
         Left            =   210
         TabIndex        =   26
         Top             =   615
         Width           =   10860
         _cx             =   1710377684
         _cy             =   1710362973
         Build           =   7
         TypeMask        =   101187586
         CylSides        =   32
         Axis(0).MinorStep=   -20
         Axis(0).Decimals=   0
         Axis(2).MinorStep=   -1
         RGBBk           =   14737632
         nColors         =   1
         Pallete         =   "Lis452.frx":1192
         Colors          =   "Lis452.frx":1276
         TopFontMask     =   268435464
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   10
         nSer            =   1
         NumPoint        =   10
         NumSer          =   1
         Title(2)        =   "Turnaround Time(Average Count)"
         _Data_          =   "Lis452.frx":129E
      End
      Begin ChartfxLibCtl.ChartFX Chart2 
         Height          =   2490
         Left            =   210
         TabIndex        =   40
         Top             =   3525
         Width           =   10860
         _cx             =   1710377684
         _cy             =   1710362920
         Build           =   7
         TypeMask        =   101187586
         CylSides        =   32
         Axis(0).MinorStep=   -20
         Axis(0).Decimals=   0
         Axis(2).MinorStep=   -1
         RGBBk           =   14737632
         nColors         =   1
         Pallete         =   "Lis452.frx":1439
         Colors          =   "Lis452.frx":151D
         TopFontMask     =   268435464
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   10
         nSer            =   1
         NumPoint        =   10
         NumSer          =   1
         Title(2)        =   "Turnaround Time(Average Time)"
         _Data_          =   "Lis452.frx":1545
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   165
         X2              =   11130
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   165
         X2              =   11130
         Y1              =   3165
         Y2              =   3165
      End
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
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
      SpreadDesigner  =   "Lis452.frx":16E0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   6975
      Index           =   0
      Left            =   3000
      Top             =   1395
      Width           =   11445
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   6975
      Index           =   1
      Left            =   75
      Top             =   1395
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   630
      Index           =   2
      Left            =   75
      Top             =   345
      Width           =   14370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "부터"
      Height          =   225
      Left            =   4500
      TabIndex        =   39
      Tag             =   "15104"
      Top             =   525
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "까지"
      Height          =   225
      Left            =   7860
      TabIndex        =   38
      Tag             =   "15104"
      Top             =   525
      Width           =   360
   End
End
Attribute VB_Name = "frm452TurnAroundTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()
Private objSQL          As New clsLISSqlStatistic
'TAT용 DICTIONARY
Private objBase         As New clsDictionary
Private objOver         As New clsDictionary
Private objWeek         As New clsDictionary

Private MaxCnt          As Long
Private MaxTime         As Long
Private MaxTimeCnt      As Long

Private MaxWeekCnt      As Long
Private MaxWeekTime     As Long
Private MaxWeekTimeCnt  As Long



Private iFromRef As Integer


Private Sub GetWeekDictionary()
    objWeek.Clear
    objWeek.FieldInialize "week", "time1,time2,cnt"
    
    objWeek.Sort = False
    objBase.MoveFirst
    Do Until objBase.EOF
        If objWeek.Exists(objBase.Fields("weekday")) Then
            objWeek.KeyChange objBase.Fields("weekday")
            objWeek.Fields("time1") = Val(objWeek.Fields("time1")) + Val(objBase.Fields("time1"))
            objWeek.Fields("time2") = Val(objWeek.Fields("time2")) + Val(objBase.Fields("time2"))
            objWeek.Fields("cnt") = Val(objWeek.Fields("cnt")) + Val(objBase.Fields("cnt"))
        Else
            objWeek.AddNew objBase.Fields("weekday"), objBase.Fields("time1") & COL_DIV & _
                                                      objBase.Fields("time2") & COL_DIV & _
                                                      objBase.Fields("cnt")
        End If
        objBase.MoveNext
    Loop
    objWeek.Sort = True
End Sub

Private Sub WeekDisplay()
'--------------------
'요일별 조회시 사용함
'--------------------
    Dim TotO    As Long    '처방-접수시간
    Dim TotR    As Long    '접수-결과시간
    Dim TotC    As Long    '검사건수
    
    Dim strDay  As String
    Dim Hour    As String
    Dim Minu    As String
    Dim ii      As Integer: ii = 1
    
    Call medClearTable(tblTab1)
    
    If objWeek.RecordCount = 0 Then Exit Sub
    
    With tblTab1
        .MaxRows = objWeek.RecordCount + 20
        objWeek.MoveFirst
        .Row = 0: .Col = 0: .Value = "Weekend"
        Do Until objWeek.EOF
            .Row = ii
            .Col = 1
            Select Case objWeek.Fields("week")
                
                Case 1: .Value = "일요일": .ForeColor = vbRed
                Case 2: .Value = "월요일": .ForeColor = vbBlack
                Case 3: .Value = "화요일": .ForeColor = vbBlack
                Case 4: .Value = "수요일": .ForeColor = vbBlack
                Case 5: .Value = "목요일": .ForeColor = vbBlack
                Case 6: .Value = "금요일": .ForeColor = vbBlack
                Case 7: .Value = "토요일": .ForeColor = vbBlue
            End Select
            
            TotC = TotC + Val(objWeek.Fields("cnt"))
            TotO = TotO + Val(objWeek.Fields("time1"))
            TotR = TotR + Val(objWeek.Fields("time2"))
            
            .Col = 2: .Value = objWeek.Fields("cnt")
            
            Hour = Format((objWeek.Fields("time1") \ objWeek.Fields("cnt")) \ 60, "0#")
            Minu = Format((objWeek.Fields("time1") \ objWeek.Fields("cnt")) Mod 60, "0#")
            
            .Col = 3: .Value = Format(Hour & Minu, "0#:##"):
            
            Hour = Format((objWeek.Fields("time2") \ objWeek.Fields("cnt")) \ 60, "0#")
            Minu = Format((objWeek.Fields("time2") \ objWeek.Fields("cnt")) Mod 60, "0#")
            .Col = 4: .Value = Format(Hour & Minu, "0#:##"): .BackColor = &HFFFDEE
                
            Hour = Format(((Val(objWeek.Fields("time1")) + Val(objWeek.Fields("time2"))) \ Val(objWeek.Fields("cnt"))) \ 60, "0#")
            Minu = Format(((Val(objWeek.Fields("time1")) + Val(objWeek.Fields("time2"))) \ Val(objWeek.Fields("cnt"))) Mod 60, "0#")
            .Col = 5: .Value = Format(Hour & Minu, "0#:##"):
            ii = ii + 1
            objWeek.MoveNext
        Loop
        txtTestcnt = TotC & " 건"
        Hour = Format((TotO \ TotC) \ 60, "0#")
        Minu = Format((TotO \ TotC) Mod 60, "0#")
        txtOTime = Format(Hour & Minu, "0#:##")
            
        Hour = Format((TotR \ TotC) \ 60, "0#")
        Minu = Format((TotR \ TotC) Mod 60, "0#")
        txtRTime = Format(Hour & Minu, "0#:##")
        
        Hour = Format(((TotR + TotO) \ TotC) \ 60, "0#")
        Minu = Format(((TotR + TotO) \ TotC) Mod 60, "0#")
        txtVTime = Format(Hour & Minu, "0#:##")
        
        .Row = .DataRowCnt + 1
        .Col = 1: .Value = "합    계          ":
        .Col = 2: .Value = TotC & " 건":
        .Col = 3: .Value = txtOTime:
        .Col = 4: .Value = txtRTime:
        .Col = 5: .Value = txtVTime:
    End With
   
End Sub
Private Sub cboWA_Click()
    Dim RS        As Recordset
    Dim strTest   As String
    Dim strWA     As String
    Dim lngLen    As Long
    Dim ii        As Integer
    
    If cboWA.ListIndex < 0 Then Exit Sub
    Set RS = New Recordset
    If tabData.SelectedItem.Index <> 3 Then
        strWA = medGetP(cboWA.Text, 1, " ")
        RS.Open objSQL.GetWAvsTest(strWA), DBConn
        lstTest.Clear
        If Not RS.EOF Then
            
            Do Until RS.EOF
                
                strTest = RS.Fields("testcd").Value & ""
                lngLen = Len(strTest)
                If lngLen < 7 Then
                    For ii = 1 To 7 - lngLen
                        strTest = strTest & " "
                    Next
                End If
                
                lstTest.AddItem strTest & "   " & RS.Fields("abbrnm10").Value & ""
                RS.MoveNext
            Loop
        End If
        
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub chkDay_Click()
    
    If objBase.RecordCount = 0 Then Exit Sub
    
    If chkDay.Value = 1 Then
        Call WeekDisplay
    Else
        Call QueryDisplayer
    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set objSQL = Nothing
    Set objBase = Nothing
    Set objOver = Nothing
    Set objWeek = Nothing
   
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub LoadWorkAreaList()
'-----------------
'Workarea 가져오기
'-----------------
    Dim rsGetWA   As Recordset
    
    Dim ii        As Integer
    
    '검사코드 Clear
    
    Set rsGetWA = New Recordset
    rsGetWA.Open objSQL.GetWACd, DBConn
    
    cboWA.Clear
    With rsGetWA
        For ii = 1 To .RecordCount
            
            cboWA.AddItem "" & .Fields("WACd").Value & "   " & _
                            "" & .Fields("WANm").Value
            .MoveNext
        Next ii
        rsGetWA.Close
    End With
    cboWA.ListIndex = 0
    Set rsGetWA = Nothing
    dtpFromDate.Value = DateAdd("d", -7, GetSystemDate)
    dtpToDate.Value = GetSystemDate

End Sub
Private Sub Clear()
    txtRTime.Text = ""
    txtVTime.Text = ""
    txtOTime.Text = ""
    txtTestcnt.Text = ""
End Sub

Private Sub cmdQuery_Click()
        
    Me.MousePointer = 11
    Select Case tabData.SelectedItem.Index - 1
        Case 0:
            Call Clear
            
            Call Item_TurnAround
        Case 1:
            Call TimeOverList
    End Select
    Me.MousePointer = 0
End Sub

Private Sub TimeOverList()
    '-------------------------------------------
    'WorkArea별 TAT시간이 지난환자 목록 보여주기
    '-------------------------------------------
    Dim objPrgBar   As jProgressBar.clsProgress
    Dim RS          As Recordset
    Dim strWorkArea As String
    Dim strTestCd   As String
    
    Dim strFDT    As String
    Dim strTDT    As String
    
    Dim strRcvDt  As String
    Dim strRcvTm  As String
    Dim strExaDt  As String
    Dim strExaTm  As String
    
    Dim strCompare1 As String
    Dim strCompare2 As String
    Dim strCompare3 As String
    
    Dim TATtime     As String
    Dim Hour        As String
    Dim Minu        As String
    
    Dim strTmp      As String   'Dictionary키 중복체크위해서.~~~~
    Dim strTmp1     As String   '데이터 ADDNew
    
    Dim strSEXAGE   As String
    
    Dim strLocation As String
    Dim strStatFg   As String
    
    Dim ii          As Long
    Dim jj          As Long
    
    Dim lngTime     As Long
    
    
    lngTime = (Val(txtday) * 24 * 60) + (Val(txthour) * 60) + Val(txtmin)
    strStatFg = IIf(optStatFg(0).Value, "0", IIf(optStatFg(1).Value, "1", "2"))
    
    If lngTime = 0 Then
        MsgBox "기준 소요시간을 입력하신후 조회하세요", vbInformation + vbOKOnly, "기준소요시간 입력"
        Exit Sub
    End If
    
    objOver.Clear
    objOver.FieldInialize "workarea,accdt,accseq,testcd", "spccd,ptid,time,ptnm,testnm,sexage,location"
    
    strWorkArea = medGetP(cboWA.Text, 1, " ")
    strTestCd = medGetP(lstTest.List(lstTest.ListIndex), 1, " ")
    strFDT = Format(dtpFromDate.Value, "YYYYmmDD")
    strTDT = Format(dtpToDate.Value, "YYYYmmDD")
    
    Set RS = New Recordset
    
    RS.Open objSQL.GetOverTATTime(strFDT, strTDT, strTestCd, strStatFg, enStsCd.StsCd_LIS_MidRst, "L", chkER.Value), DBConn
    
    tblOver.MaxRows = 0
    If Not RS.EOF Then
        Set objPrgBar = New jProgressBar.clsProgress
        
        objPrgBar.Max = RS.RecordCount
        objPrgBar.Left = LisLabel1.Left
        objPrgBar.Top = LisLabel1.Top
        objPrgBar.Width = LisLabel1.Width
        objPrgBar.Container = Me
'        objPrgBar.SetMyForm Me
    
        objOver.Sort = False
        Do Until RS.EOF
            ii = ii + 1
            strRcvDt = RS.Fields("rcvdt").Value & ""
            strRcvTm = RS.Fields("rcvtm").Value & ""
            strExaDt = RS.Fields("vfydt").Value & ""
            strExaTm = RS.Fields("vfytm").Value & ""
            
            '접수시간
            strCompare1 = strRcvTm
            If Len(strCompare1) = 4 Then
                strCompare1 = strCompare2 & "00"
            Else
                strCompare1 = Mid(strCompare1, 1, 4) & "00"
            End If
            strCompare1 = Format(strRcvDt, "0###-##-##") & " " & Format(strCompare1, "0#:##:##")
            
            '결과시간
            strCompare2 = strExaTm
            If Len(strCompare2) = 4 Then
                strCompare2 = strCompare2 & "00"
            Else
                strCompare2 = Mid(strCompare2, 1, 4) & "00"
            End If
            
            strCompare2 = Format(strExaDt, "0###-##-##") & " " & Format(strCompare2, "0#:##:##")
            
            '접수-결과시간차
            TATtime = DateDiff("n", strCompare1, strCompare2)
            
            Select Case RS.Fields("bussdiv").Value & ""
                Case "1", "4"
                    strLocation = RS.Fields("deptcd").Value & ""
                Case "2", "3"
                    strLocation = RS.Fields("wardid").Value & ""
                    If strLocation <> "" Then
                        If RS.Fields("hosilid").Value & "" <> "" Then strLocation = strLocation & "/" & RS.Fields("hosilid").Value & ""
                    Else
                        If RS.Fields("hosilid").Value & "" <> "" Then strLocation = RS.Fields("hosilid").Value & ""
                    End If
            End Select
            strSEXAGE = GetSexAge(RS.Fields("ssn").Value & "")
            strTmp = RS.Fields("workarea").Value & "" & COL_DIV & RS.Fields("accdt").Value & "" & COL_DIV & RS.Fields("accseq").Value & "" & COL_DIV & RS.Fields("testcd").Value & ""
            strTmp1 = RS.Fields("spccd").Value & "" & "" & COL_DIV & _
                      RS.Fields("ptid").Value & "" & COL_DIV & _
                      TATtime & COL_DIV & _
                      RS.Fields("ptnm").Value & "" & COL_DIV & _
                      RS.Fields("abbrnm10").Value & "" & COL_DIV & _
                      strSEXAGE & COL_DIV & _
                      strLocation
            objOver.Sort = False
            If objOver.Exists(strTmp) Then
                objOver.KeyChange strTmp
                objOver.Fields("spccd") = RS.Fields("spccd").Value & ""
                objOver.Fields("ptid") = RS.Fields("ptid").Value & ""
                objOver.Fields("time") = TATtime
                objOver.Fields("ptnm") = RS.Fields("ptnm").Value & ""
                objOver.Fields("testnm") = RS.Fields("abbrnm10").Value & ""
                objOver.Fields("sexage") = strSEXAGE
                objOver.Fields("location") = strLocation
            Else
                objOver.AddNew strTmp, strTmp1
            End If
            objPrgBar.Value = ii
            '----------------------------------------
            '시간 체크를 해서 이상인것만 보여줄거야..
            '----------------------------------------
            If lngTime <= TATtime Then
                With tblOver
                    If .DataRowCnt >= .MaxRows Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    .Col = 1: .Value = RS.Fields("ptid").Value & ""
                    .Col = 2: .Value = RS.Fields("ptnm").Value & ""
                    .Col = 3: .Value = strSEXAGE
                    .Col = 4: .Value = strLocation
                    .Col = 5: .Value = Mid(strCompare1, 1, 10) & " " & Mid(strCompare1, 11, 6) 'Mid(strCompare1, 1, Len(strCompare1) - 3)
                    .Col = 6: .Value = Mid(strCompare2, 1, 10) & " " & Mid(strCompare2, 11, 6)
                    .Col = 7: .Value = GetEmpNm(RS.Fields("vfyid").Value & "")
                    Hour = Format((TATtime) \ 60, "0#")
                    Minu = Format((TATtime) Mod 60, "0#")
                    TATtime = Format(Hour & Minu, "0#:##")
                    .Col = 8: .Value = TATtime
                    
                End With
            End If
            RS.MoveNext
        Loop
        objOver.Sort = True
        cmdPrint.Enabled = True
        cmdExcel.Enabled = True
        cmdGraph.Enabled = True
        If tblOver.MaxRows = 0 Then
            MsgBox "해당조건의 조회리스트가 없습니다.", vbInformation + vbOKOnly, Me.Caption
            cmdPrint.Enabled = False
            cmdExcel.Enabled = False
            cmdGraph.Enabled = False
        End If
    Else
        MsgBox "해당조건의 조회리스트가 없습니다.", vbInformation + vbOKOnly, Me.Caption
        cmdPrint.Enabled = False
        cmdExcel.Enabled = False
        cmdGraph.Enabled = False
    End If
    Set RS = Nothing
    Set objPrgBar = Nothing
    
End Sub

Private Sub Item_TurnAround()
    '-------------------------
    'TAT 탭인 경우 사용하였슴.
    '-------------------------
    Dim objPrgBar   As jProgressBar.clsProgress
    Dim RS          As Recordset
    
    Dim strFromDt   As String
    Dim strToDt     As String
    Dim strWorkArea As String
    Dim strTestCd   As String
    Dim strOrdDt    As String
    Dim strOrdTm    As String
    Dim strRcvDt    As String
    Dim strRcvTm    As String
    Dim strExaDt    As String
    Dim strExaTm    As String
    Dim strWeek     As String
    Dim strStatFg   As String
    
    
    Dim strCompare1 As String
    Dim strCompare2 As String
    Dim strCompare3 As String
    
    Dim strTmp      As String
    Dim Cnt         As Long
    
    Dim Time1       As String         '처방-접수
    Dim Time2       As String         '접수-결과
    Dim ii          As Integer
    
    objBase.Clear
    objBase.FieldInialize "rcvdt", "time1,time2,Cnt,weekday"
    
    Call medClearTable(tblTab1)
    
    chkDay.Enabled = False
    strWorkArea = medGetP(cboWA.Text, 1, " ")
    strFromDt = Format(dtpFromDate.Value, "YYYYmmDD")
    strToDt = Format(dtpToDate.Value, "YYYYmmDD")
    
    strStatFg = IIf(optStatFg(0).Value, "0", IIf(optStatFg(1).Value, "1", "2"))
    
    strTestCd = medGetP(lstTest.List(lstTest.ListIndex), 1, " ")
    If strTestCd = "" Then
        MsgBox "검사항목을 선택하신후 조회하세요.", vbInformation + vbOKOnly, "검사항목선택"
        Exit Sub
    End If
    
    '## 5.0.12: 이상대(2004-12-29)
    '   - Workarea 조건을 추가하기 위해 GetTurnAroundDateRecord 대시 GetTurnAroundDateRecordX 사용
    Set RS = New Recordset
    Set RS = objSQL.GetTurnAroundDateRecordX(strFromDt, strToDt, strTestCd, enStsCd.StsCd_LIS_MidRst, "L", strStatFg, chkER.Value, strWorkArea)
    
    If Not RS.EOF Then
        Set objPrgBar = New jProgressBar.clsProgress
        
        objPrgBar.Max = RS.RecordCount
        objPrgBar.Left = LisLabel1.Left
        objPrgBar.Top = LisLabel1.Top
        objPrgBar.Width = LisLabel1.Width
        objPrgBar.Container = Me
        
        objBase.Sort = False
        Do Until RS.EOF
            ii = ii + 1
            strOrdDt = RS.Fields("coldt").Value & ""
            strOrdTm = RS.Fields("coltm").Value & ""
            strRcvDt = RS.Fields("rcvdt").Value & ""
            strRcvTm = RS.Fields("rcvtm").Value & ""
            strExaDt = RS.Fields("vfydt").Value & ""
            strExaTm = RS.Fields("vfytm").Value & ""
            strWeek = Weekday(Format(RS.Fields("rcvdt").Value & "", "####-##-##"))
            
            '처방시간
            strCompare1 = strOrdTm
            If Len(strCompare1) = 4 Then
                strCompare1 = strCompare1 & "00"
            Else
                strCompare1 = Mid(strCompare1, 1, 4) & "00"
            End If
            
            strCompare1 = Format(strOrdDt, "0###-##-##") & " " & Format(strCompare1, "0#:##:##")
            
            '접수시간
            strCompare2 = strRcvTm
            If Len(strCompare2) = 4 Then
                strCompare2 = strCompare2 & "00"
            Else
                strCompare2 = Mid(strCompare2, 1, 4) & "00"
            End If
            strCompare2 = Format(strRcvDt, "0###-##-##") & " " & Format(strCompare2, "0#:##:##")
            
            '결과시간
            strCompare3 = strExaTm
            If Len(strCompare3) = 4 Then
                strCompare3 = strCompare3 & "00"
            Else
                strCompare3 = Mid(strCompare3, 1, 4) & "00"
            End If
            
            strCompare3 = Format(strExaDt, "0###-##-##") & " " & Format(strCompare3, "0#:##:##")
            
            '처방-접수시간차
            Time1 = DateDiff("n", strCompare1, strCompare2)
            '접수-결과시간차
            Time2 = DateDiff("n", strCompare2, strCompare3)
            If objBase.Exists(strRcvDt) Then
                objBase.KeyChange strRcvDt
                objBase.Fields("time1") = Val(objBase.Fields("time1")) + Time1
                objBase.Fields("time2") = Val(objBase.Fields("time2")) + Time2
                objBase.Fields("cnt") = Val(objBase.Fields("cnt")) + 1
                Cnt = Cnt + 1
            Else
                objBase.AddNew strRcvDt, Time1 & COL_DIV & Time2 & COL_DIV & "1" & COL_DIV & strWeek
                strTmp = RS.Fields("rcvdt").Value & ""
                Cnt = 1
            End If
             
            objPrgBar.Message = Format(strTmp, "####-##-##") & " 일의 " & Cnt & " 번째를 처리하고 있습니다."
            objPrgBar.Value = ii
            RS.MoveNext
        Loop
        objBase.Sort = True
        chkDay.Enabled = True
        '일자별 조회결과 보여주기
        Call QueryDisplayer
        '요일별 Dictionary에 담기
        Call GetWeekDictionary
        
        cmdPrint.Enabled = True
        cmdExcel.Enabled = True
        cmdGraph.Enabled = True
    Else
        MsgBox "해당조건의 자료가 없습니다.", vbInformation + vbOKOnly, Me.Caption
        cmdPrint.Enabled = False
        cmdExcel.Enabled = False
        cmdGraph.Enabled = False
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub QueryDisplayer()
'------------------
'접수일자별 DISPLAY
'------------------
    Dim TotO As Long    '처방-접수시간
    Dim TotR As Long    '접수-결과시간
    Dim TotC As Long    '검사건수
    Dim Hour As String
    Dim Minu As String
    
    
    Dim ii   As Integer
    
    If objBase.RecordCount < 1 Then Exit Sub
    With tblTab1
        .MaxRows = 0
        .MaxRows = objBase.RecordCount + 20: ii = 1
        .Row = 0: .Col = 0: .Value = "Accssion Date"
        objBase.MoveFirst
        Do Until objBase.EOF
            .Row = ii
            .Col = 1: .Value = Format(objBase.Fields("rcvdt"), "####-##-##")
            
            TotC = TotC + Val(objBase.Fields("cnt"))
            TotO = TotO + Val(objBase.Fields("time1"))
            TotR = TotR + Val(objBase.Fields("time2"))
            
            .Col = 2: .Value = objBase.Fields("cnt")
            
            Hour = Format((objBase.Fields("time1") \ objBase.Fields("cnt")) \ 60, "0#")
            Minu = Format((objBase.Fields("time1") \ objBase.Fields("cnt")) Mod 60, "0#")
            .Col = 3: .Value = Format(Hour & Minu, "0#:##"):
            
            Hour = Format((objBase.Fields("time2") \ objBase.Fields("cnt")) \ 60, "0#")
            Minu = Format((objBase.Fields("time2") \ objBase.Fields("cnt")) Mod 60, "0#")
            .Col = 4: .Value = Format(Hour & Minu, "0#:##"): .BackColor = &HFFFDEE
                
            Hour = Format(((Val(objBase.Fields("time1")) + Val(objBase.Fields("time2"))) \ Val(objBase.Fields("cnt"))) \ 60, "0#")
            Minu = Format(((Val(objBase.Fields("time1")) + Val(objBase.Fields("time2"))) \ Val(objBase.Fields("cnt"))) Mod 60, "0#")
            .Col = 5: .Value = Format(Hour & Minu, "0#:##"):
            
            ii = ii + 1
            objBase.MoveNext
        Loop
    
    
        txtTestcnt = TotC & " 건"
        Hour = Format((TotO \ TotC) \ 60, "0#")
        Minu = Format((TotO \ TotC) Mod 60, "0#")
        txtOTime = Format(Hour & Minu, "0#:##")
            
        Hour = Format((TotR \ TotC) \ 60, "0#")
        Minu = Format((TotR \ TotC) Mod 60, "0#")
        txtRTime = Format(Hour & Minu, "0#:##")
        
        Hour = Format(((TotR + TotO) \ TotC) \ 60, "0#")
        Minu = Format(((TotR + TotO) \ TotC) Mod 60, "0#")
        txtVTime = Format(Hour & Minu, "0#:##")
        .Row = .DataRowCnt + 1
        .Col = 1: .Value = "합    계          ":
        .Col = 2: .Value = TotC & " 건":
        .Col = 3: .Value = txtOTime:
        .Col = 4: .Value = txtRTime:
        .Col = 5: .Value = txtVTime:
        
    End With
    
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Call LoadWorkAreaList
    Call Clear
    optStatFg(0).Value = True
    chkDay.Value = 0
    chkER.Value = 0
    fraData(0).ZOrder 0
    cmdPrint.Enabled = False
    cmdExcel.Enabled = False
    cmdGraph.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSQL = Nothing
    Set objBase = Nothing
    Set objOver = Nothing
    Set objWeek = Nothing
End Sub

Private Sub lstTest_Click()
    Select Case tabData.SelectedItem.Index
        Case 2
            txtday = "": txthour = "": txtmin = ""
    End Select
    lblTest.Caption = medGetP(cboWA.Text, 2, "   ") & "  :  " & Trim(Mid(lstTest.Text, 9)) & "(" & medGetP(lstTest.Text, 1, " ") & ")"
End Sub

Private Sub tabData_Click()
    fraData(tabData.SelectedItem.Index - 1).ZOrder 0
    Select Case tabData.SelectedItem.Index
        Case 1: lstTest.Enabled = True: lblTest.Caption = medGetP(cboWA.Text, 2, "   ") & "  :  " & Trim(Mid(lstTest.Text, 9)) & "(" & medGetP(lstTest.Text, 1, " ") & ")"
        Case 2: lstTest.Enabled = True: lblTest.Caption = medGetP(cboWA.Text, 2, "   ") & "  :  " & Trim(Mid(lstTest.Text, 9)) & "(" & medGetP(lstTest.Text, 1, " ") & ")"
                txtday = "": txthour = "": txtmin = ""
        Case 3: lstTest.Enabled = False: lblTest.Caption = medGetP(cboWA, 2, "   ") & "  : TAT until Verify FROM Receive"
    End Select
    cmdQuery.Enabled = True
    cmdExcel.Enabled = True
    cmdPrint.Enabled = True
End Sub

Private Sub cmdBack_Click()
    fraData(0).ZOrder 0
    cmdPrint.Enabled = True
    cmdExcel.Enabled = True
    cmdQuery.Enabled = True
End Sub
 Private Function GetSexAge(ByVal ssn As String) As String
    Dim strTmp As String
    Dim strSex As String
    Dim strAGE As String
    Dim strDOB As String
    
    Dim strYY  As String
    Dim strMM  As String
    Dim strDD  As String
    
    strYY = Mid(ssn, 1, 2)
    strMM = Mid(ssn, 3, 2)
    strDD = Mid(ssn, 5, 2)
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    If IsDate(strYY & "-" & strMM & "-" & strDD) = False Then
        strDD = "01"
    End If
    
    strSex = "기타": strAGE = "": strDOB = ""
    
    If ssn <> "" Then
        strTmp = Mid(ssn, 7, 1)
        Select Case strTmp
            Case "0": strSex = "F": strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1": strSex = "M": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2": strSex = "F": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3": strSex = "M": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4": strSex = "F": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strSex = "M": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        If Len(ssn) > 6 Then
            If strDOB <> "" Then
                strAGE = medFindAge(Replace(strDOB, "-", ""), "Y")
            End If
        Else
            strAGE = ""
        End If
        GetSexAge = strSex & "/" & strAGE
    Else
        GetSexAge = ""
    End If
End Function
Private Sub cmdGraph_Click()
'-----------------------------------------
'그래프보기(Default로 접수일자별(시간별임)
'-----------------------------------------
    If objBase.RecordCount = 0 Or objWeek.RecordCount = 0 Then
        MsgBox "조회를 하신후 그래프보기를 Click 하세요.", vbInformation + vbOKOnly, "그래프보기"
        Exit Sub
    End If
    fraData(2).ZOrder 0
    
    Dim tf      As Boolean
    
    With objBase
        .MoveFirst
        Do Until .EOF
            If tf = False Then
                MaxCnt = .Fields("cnt")
                MaxTime = .Fields("time2")
                MaxTimeCnt = .Fields("cnt")
            End If
            If MaxCnt < .Fields("cnt") Then MaxCnt = .Fields("cnt")
            
            If MaxTime < .Fields("time2") Then
                MaxTime = .Fields("time2")
                MaxTimeCnt = .Fields("cnt")
            End If
            
            tf = True
            .MoveNext
        Loop
    End With
    
    tf = False
    With objWeek
        .MoveFirst
        Do Until .EOF
            If tf = False Then
                MaxWeekCnt = .Fields("cnt")
                MaxWeekTime = .Fields("time2")
                MaxWeekTimeCnt = .Fields("cnt")
            End If
            If MaxWeekCnt < .Fields("cnt") Then MaxWeekCnt = .Fields("cnt")
            If MaxWeekTime < .Fields("time2") Then
                MaxWeekTime = .Fields("time2")
                MaxWeekTimeCnt = .Fields("cnt")
            End If
            
            tf = True
            .MoveNext
        Loop
    End With
    
    If chkDay.Value = 0 Then
        optDiv(0).Value = True
        optBase(2).Value = True
        Call ShowGraph_RcvDtCnt(0, MaxCnt)
        Call ShowGraph_RcvDTTime_Day(0, MaxTime, MaxTimeCnt)
    Else
        optDiv(1).Value = True
        optBase(2).Value = True
        Call ShowGraph_WeekCnt(0, MaxWeekCnt)
        Call ShowGraph_WeekTime_Day(0, MaxWeekTime, MaxWeekTimeCnt)
    End If
    cmdQuery.Enabled = False
    cmdPrint.Enabled = False
    cmdExcel.Enabled = False
End Sub
Private Sub ShowGraph_RcvDtCnt(ByVal iMinvalue As Long, ByVal iMaxValue As Long)
'--------------------
'접수일자별 Cnt Graph
'--------------------
    Dim iSeries As Integer
    Dim iPoints As Integer
    Dim avecnt  As Long
    
    Dim ii      As Integer
    
    iSeries = 1
    iPoints = objBase.RecordCount
    
    With chart1
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
        .OpenDataEx COD_VALUES, iSeries, iPoints
        .PointLabels = True
        .Scrollable = True
        .Axis(AXIS_Y).Decimals = 0
        .BottomGap = 10
        objBase.MoveFirst
        
        Do Until objBase.EOF
            .Axis(AXIS_X).Label(ii) = Format(Mid(objBase.Fields("rcvdt"), 5), "0#-##")
            .ValueEx(0, ii) = objBase.Fields("cnt")
            avecnt = avecnt + Val(objBase.Fields("cnt"))
            ii = ii + 1
            objBase.MoveNext
        Loop
        
        '-------------
        '검사건수 평균
        '-------------
        
        avecnt = avecnt / Val(objBase.RecordCount)
        
        .Axis(AXIS_Y).Min = iMinvalue - ((iMaxValue - iFromRef) / 10) '1
        .Axis(AXIS_Y).Max = iMaxValue + ((iMaxValue - iFromRef) / 10) '1
        
        .Axis(AXIS_Y).STEP = (iMaxValue - iMinvalue) / 3
        .OpenDataEx COD_CONSTANTS, 2, 0
        .ConstantLine(0).Value = avecnt
        .ConstantLine(0).LineColor = &H808080  '&H80&
        .ConstantLine(0).Axis = AXIS_Y
        .ConstantLine(0).Label = CStr(avecnt)
        .ConstantLine(0).LineWidth = 1
        .ConstantLine(0).LineStyle = CHART_DOT
        .CloseData COD_CONSTANTS
        
        .CloseData COD_VALUES + COD_SCROLLLEGEND
    End With
    
End Sub
Private Sub ShowGraph_WeekCnt(ByVal iMinvalue As Long, ByVal iMaxValue As Long)
'----------------
'요일별 Cnt Graph
'----------------
    Dim iSeries As Integer
    Dim iPoints As Integer
    Dim avecnt  As Long
    
    Dim ii      As Integer
    
    iSeries = 1
    iPoints = objWeek.RecordCount
    
    With chart1
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
        .OpenDataEx COD_VALUES, iSeries, iPoints
        .PointLabels = True
        .Scrollable = True
        .Axis(AXIS_Y).Decimals = 0
        .BottomGap = 10
        objWeek.MoveFirst
        Do Until objWeek.EOF
            Select Case objWeek.Fields("week")
                Case 1: .Axis(AXIS_X).Label(ii) = "일요일"
                Case 2: .Axis(AXIS_X).Label(ii) = "월요일"
                Case 3: .Axis(AXIS_X).Label(ii) = "화요일"
                Case 4: .Axis(AXIS_X).Label(ii) = "수요일"
                Case 5: .Axis(AXIS_X).Label(ii) = "목요일"
                Case 6: .Axis(AXIS_X).Label(ii) = "금요일"
                Case 7: .Axis(AXIS_X).Label(ii) = "토요일"
            End Select
            
            .ValueEx(0, ii) = objWeek.Fields("cnt")
            avecnt = avecnt + Val(objWeek.Fields("cnt"))
            ii = ii + 1
            objWeek.MoveNext
        Loop
        '------------
        '검사건수평균
        '------------
        avecnt = avecnt / Val(objWeek.RecordCount)
        .OpenDataEx COD_CONSTANTS, 2, 0
        
        .ConstantLine(0).Value = avecnt
        .ConstantLine(0).LineColor = &H808080  '&H80&
        .ConstantLine(0).Axis = AXIS_Y
        .ConstantLine(0).Label = CStr(avecnt)
        .ConstantLine(0).LineWidth = 1
        .ConstantLine(0).LineStyle = CHART_DOT
        .CloseData COD_CONSTANTS
        
        .Axis(AXIS_Y).Min = iMinvalue - ((iMaxValue - iFromRef) / 10) '1
        .Axis(AXIS_Y).Max = iMaxValue + ((iMaxValue - iFromRef) / 10) '1
        
        .Axis(AXIS_Y).STEP = (iMaxValue - iMinvalue) / 3
        
        .CloseData COD_VALUES + COD_SCROLLLEGEND
    End With
End Sub
Private Sub ShowGraph_RcvDTTime_Day(ByVal iMinvalue As Long, ByVal iMaxValue As Long, ByVal iMaxCnt As Long)
'--------------------
'접수일자별 Time Graph
'--------------------
    Dim iSeries As Integer
    Dim iPoints As Integer
    
    Dim Mode    As Integer
    
    Dim ii      As Integer
    
    Dim lngTime   As Long
    Dim avetime   As Long
    
    Dim YposMax   As Long
    
    Dim totCnt     As Long
    Dim totTime    As Double
    Dim totTimeCnt As Double
    
    
    Dim lngDTime As Long
    Dim lngHTime As Long
    
    
    
    lngDTime = 60 * 24
    lngHTime = 60
    
    
    iSeries = 1
    iPoints = objBase.RecordCount
    
    For ii = 0 To 2
        If optBase(ii).Value = True Then
            Mode = ii
        End If
    Next
    
    Select Case Mode
        Case 0:
            '검사시간이 1일보다 작으면 무조건 1일
            If iMaxValue <= lngDTime Then
                YposMax = 1
            Else
                YposMax = (iMaxValue / iMaxCnt) \ lngDTime
            End If
            If YposMax = 0 Then YposMax = 1
        Case 1:
            If iMaxValue <= lngHTime Then
                YposMax = 1
            Else
                YposMax = ((iMaxValue + (30 * iMaxCnt)) / iMaxCnt) \ lngHTime
            End If
        Case 2: YposMax = (iMaxValue) \ iMaxCnt
    End Select
    
    ii = 0
    With Chart2
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
        .OpenDataEx COD_VALUES, iSeries, iPoints
        .PointLabels = True
        .Scrollable = True
        
        .BottomGap = 10
        objBase.MoveFirst
        Do Until objBase.EOF
            Select Case Mode
                Case 0:
                    lngTime = (Val(objBase.Fields("time2")) / Val(objBase.Fields("cnt"))) \ lngDTime
                    If lngTime <= 0 Then lngTime = 1
                    totTime = totTime + (objBase.Fields("time2") \ lngDTime)
                Case 1:
                    lngTime = (Val(objBase.Fields("time2")) + (30 * Val(objBase.Fields("cnt")))) / Val(objBase.Fields("cnt")) \ lngHTime
                    If lngTime <= 0 Then lngTime = 1
                    totTime = totTime + (objBase.Fields("time2") \ lngHTime)
                Case 2:
                    lngTime = (Val(objBase.Fields("time2")) / Val(objBase.Fields("cnt"))):
                    totTime = totTime + objBase.Fields("time2")
            End Select
            totCnt = totCnt + objBase.Fields("cnt")
            .Axis(AXIS_X).Label(ii) = Format(Mid(objBase.Fields("rcvdt"), 5), "0#-##")
            .ValueEx(0, ii) = lngTime
            ii = ii + 1
            objBase.MoveNext
        Loop
        
        
        avetime = totTime \ totCnt
        If avetime = 0 Then avetime = 1
        
        .OpenDataEx COD_CONSTANTS, 2, 0
        .ConstantLine(0).Value = avetime
        .ConstantLine(0).LineColor = &H808080  '&H80&
        .ConstantLine(0).Axis = AXIS_Y
        .ConstantLine(0).Label = CStr(avetime)
        .ConstantLine(0).LineWidth = 1
        .ConstantLine(0).LineStyle = CHART_DOT
        .CloseData COD_CONSTANTS
        
        
        .Axis(AXIS_Y).Min = iMinvalue - ((YposMax - iFromRef) / 10) '1
        .Axis(AXIS_Y).Max = YposMax + ((YposMax - iFromRef) / 10) '1
        
        .Axis(AXIS_Y).STEP = (YposMax - iMinvalue) / 3
        
        .CloseData COD_VALUES + COD_SCROLLLEGEND
    End With


End Sub
Private Sub ShowGraph_WeekTime_Day(ByVal iMinvalue As Long, ByVal iMaxValue As Long, ByVal iMaxCnt As Long)
'--------------------
'요일별 Time Graph
'--------------------
    Dim iSeries As Integer
    Dim iPoints As Integer
    
    Dim Mode    As Integer
    
    Dim ii      As Integer
    
    Dim lngTime   As Long
    Dim avetime   As Long
    
    Dim YposMax   As Long
    
    Dim totCnt     As Long
    Dim totTime    As Double
    Dim totTimeCnt As Double
    
    
    Dim lngDTime As Long
    Dim lngHTime As Long
    
    
    
    lngDTime = 60 * 24
    lngHTime = 60
    
    
    iSeries = 1
    iPoints = objWeek.RecordCount
    
    For ii = 0 To 2
        If optBase(ii).Value = True Then
            Mode = ii
        End If
    Next
    
    Select Case Mode
        Case 0:
            '검사시간이 1일보다 작으면 무조건 1일
            If iMaxValue <= lngDTime Then
                YposMax = 1
            Else
                YposMax = (iMaxValue / iMaxCnt) \ lngDTime
            End If
            If YposMax = 0 Then YposMax = 1
        Case 1:
            If iMaxValue <= lngHTime Then
                YposMax = 1
            Else
                YposMax = ((iMaxValue + (30 * iMaxCnt)) / iMaxCnt) \ lngHTime
            End If
        Case 2: YposMax = (iMaxValue) \ iMaxCnt
    End Select
    
    ii = 0
    With Chart2
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
        .OpenDataEx COD_VALUES, iSeries, iPoints
        .PointLabels = True
        .Scrollable = True
        
        .BottomGap = 10
        objWeek.MoveFirst
        Do Until objWeek.EOF
            Select Case Mode
                Case 0:
                    lngTime = (Val(objWeek.Fields("time2")) / Val(objWeek.Fields("cnt"))) \ lngDTime
                    If lngTime <= 0 Then lngTime = 1
                    totTime = totTime + (objWeek.Fields("time2") \ lngDTime)
                Case 1:
                    lngTime = (Val(objWeek.Fields("time2")) + (30 * Val(objWeek.Fields("cnt")))) / Val(objWeek.Fields("cnt")) \ lngHTime
                    If lngTime <= 0 Then lngTime = 1
                    totTime = totTime + (objWeek.Fields("time2") \ lngHTime)
                Case 2:
                    lngTime = (Val(objWeek.Fields("time2")) / Val(objWeek.Fields("cnt"))):
                    totTime = totTime + objWeek.Fields("time2")
            End Select
            totCnt = totCnt + objWeek.Fields("cnt")
            
            Select Case objWeek.Fields("week")
                Case 1: .Axis(AXIS_X).Label(ii) = "일요일"
                Case 2: .Axis(AXIS_X).Label(ii) = "월요일"
                Case 3: .Axis(AXIS_X).Label(ii) = "화요일"
                Case 4: .Axis(AXIS_X).Label(ii) = "수요일"
                Case 5: .Axis(AXIS_X).Label(ii) = "목요일"
                Case 6: .Axis(AXIS_X).Label(ii) = "금요일"
                Case 7: .Axis(AXIS_X).Label(ii) = "토요일"
            End Select
            
            .ValueEx(0, ii) = lngTime
            ii = ii + 1
            objWeek.MoveNext
        Loop
        
        avetime = totTime \ totCnt
        If avetime = 0 Then avetime = 1
        
        
        .OpenDataEx COD_CONSTANTS, 2, 0
        .ConstantLine(0).Value = avetime
        .ConstantLine(0).LineColor = &H808080  '&H80&
        .ConstantLine(0).Axis = AXIS_Y
        .ConstantLine(0).Label = CStr(avetime)
        .ConstantLine(0).LineWidth = 1
        .ConstantLine(0).LineStyle = CHART_DOT
        .CloseData COD_CONSTANTS
        
        .Axis(AXIS_Y).Min = iMinvalue - ((YposMax - iFromRef) / 10) '1
        .Axis(AXIS_Y).Max = YposMax + ((YposMax - iFromRef) / 10) '1
        .Axis(AXIS_Y).STEP = (YposMax - iMinvalue) / 3
        
        .CloseData COD_VALUES + COD_SCROLLLEGEND
    End With
End Sub
Private Sub optDiv_Click(Index As Integer)
'요일별,접수일자별
    Select Case Index
        Case 0
            Call ShowGraph_RcvDtCnt(0, MaxCnt)
            Call ShowGraph_RcvDTTime_Day(0, MaxTime, MaxTimeCnt)

        Case 1
            Call ShowGraph_WeekCnt(0, MaxWeekCnt)
            Call ShowGraph_WeekTime_Day(0, MaxWeekTime, MaxWeekTimeCnt)
    End Select
    
End Sub
Private Sub optBase_Click(Index As Integer)
'시간기준 설정
    If optDiv(0).Value = True Then
        Call ShowGraph_RcvDTTime_Day(0, MaxTime, MaxTimeCnt)
    Else
        Call ShowGraph_WeekTime_Day(0, MaxWeekTime, MaxWeekTimeCnt)
    End If
End Sub
Private Sub cmdExcel_Click()

    Dim strTmp1  As String
    Dim strTmp2  As String
    
    Dim lngRows1 As Long
    Dim lngRows2 As Long
    
    Dim lngCols1 As Long
    Dim lngCols2 As Long
    
    If tblTab1.DataRowCnt = 0 And tblOver.DataRowCnt = 0 Then Exit Sub
    
    With tblTab1
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp1 = .Clip
        .BlockMode = False
        lngRows1 = .MaxRows: lngCols1 = .MaxCols
    End With
    
    With tblOver
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp2 = .Clip
        .BlockMode = False
        lngRows2 = .MaxRows: lngCols2 = .MaxCols
    End With
    
    With tblexcel
        .MaxRows = IIf(tabData.SelectedItem.Index = 1, lngRows1, lngRows2)
        .MaxCols = IIf(tabData.SelectedItem.Index = 1, lngCols1, lngCols2)
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Clip = ""
        Select Case tabData.SelectedItem.Index
            Case 1: .Clip = strTmp1
            Case 2: .Clip = strTmp2
        End Select
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    Select Case tabData.SelectedItem.Index
        Case 1:  DlgSave.FileName = "Turnaround Time"
        Case 2: DlgSave.FileName = "TAT OverTime"
    End Select
    DlgSave.ShowSave
    tblexcel.SaveTabFile (DlgSave.FileName)
    
End Sub


Private Sub TurnAroundListHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 10

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("TurnAround Time List", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    Call Print_Setting("업무구분 : Laboratory", PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    
    If chkDay.Value = 1 Then
        strTmp = "요일"
    Else
        strTmp = "접수일자"
    End If
    
    Call Print_Setting("출력구분 : " & strTmp & "별", PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    Call Print_Setting("업무영역 : " & Trim(medGetP(lblTest.Caption, 1, ":")), PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    Call Print_Setting("검사항목 : " & Trim(medGetP(lblTest.Caption, 2, ":")), PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
    Call TurnAroundBody("번호", strTmp, "건수", "채혈 ~ 접수", "접수 ~ 결과", "채혈 ~ 결과")
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub
Private Sub TurnAroundBody(ByVal sNo As String, ByVal sAccDt As String, ByVal sCnt As String, _
                           ByVal sTime1 As String, ByVal sTime2 As String, ByVal sTime3 As String)
                           
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call TurnAroundListHead
    End If
   
    Call Print_Setting(sNo, 5, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sAccDt, 20, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sCnt, 50, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sTime1, 70, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sTime2, 100, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sTime3, 130, LineSpace, 20, "L", "C")
End Sub

Private Sub PrintItem_TurnAround()
    Dim sNo     As String
    Dim sAccno  As String
    Dim sCnt    As String
    Dim sTime1  As String
    Dim sTime2  As String
    Dim sTime3  As String
    
    Dim ii          As Integer
    
    If tblTab1.DataRowCnt < 1 Then Exit Sub
    
    Call P_PrtSet
    Call TurnAroundListHead
    
    With tblTab1
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1:   sNo = .Row
            .Col = 1:   sAccno = .Value
            .Col = 2:   sCnt = .Value
            .Col = 3:   sTime1 = .Value
            .Col = 4:   sTime2 = .Value
            .Col = 5:   sTime3 = .Value
            Call TurnAroundBody(sNo, sAccno, sCnt, sTime1, sTime2, sTime3)
        Next
    End With
    
    Printer.EndDoc
End Sub


Private Sub TimeOverListHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 10

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("TAT OverTime List", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    Call Print_Setting("업무구분 : Laboratory", PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    Call Print_Setting("업무영역 : " & Trim(medGetP(lblTest.Caption, 1, ":")), PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    Call Print_Setting("검사항목 : " & Trim(medGetP(lblTest.Caption, 2, ":")), PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
    Call TimeOverListBody("번호", "환자ID", "환자명", "성별/나이", "Location", "접수일시", "결과입력일시", "입력자", "Time")
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub
Private Sub TimeOverListBody(ByVal sNo As String, ByVal sPtid As String, ByVal sPtnm As String, ByVal sSexAge As String, _
                             ByVal sLocation As String, ByVal sAccDt As String, ByVal sVfydt As String, _
                             ByVal sVfyNm As String, ByVal sTime As String)
                           
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call TimeOverListHead
    End If
   
    Call Print_Setting(sNo, 5, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sPtid, 15, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sPtnm, 35, LineSpace, 15, "L", "C", False)
    Call Print_Setting(sSexAge, 50, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sLocation, 70, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sAccDt, 90, LineSpace, 35, "L", "C", False)
    Call Print_Setting(sVfydt, 125, LineSpace, 35, "L", "C", False)
    Call Print_Setting(sVfyNm, 160, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sTime, 180, LineSpace, 20, "L", "C")
End Sub

Private Sub PrintTimeOverList()
    Dim sNo         As String
    Dim sPtid       As String
    Dim sPtnm       As String
    Dim sSexAge     As String
    Dim sLocation   As String
    Dim sAccDt      As String
    Dim sVfydt      As String
    Dim sVfyNm      As String
    Dim sTime       As String
    
    Dim ii          As Integer
    
    If tblOver.DataRowCnt < 1 Then Exit Sub
    
    Call P_PrtSet
    Call TimeOverListHead
    
    With tblOver
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 0:   sNo = .Row
            .Col = 1:   sPtid = .Value
            .Col = 2:   sPtnm = .Value
            .Col = 3:   sSexAge = .Value
            .Col = 4:   sLocation = .Value
            .Col = 5:   sAccDt = .Value
            .Col = 6:   sVfydt = .Value
            .Col = 7:   sVfyNm = .Value
            .Col = 8:   sTime = .Value
            Call TimeOverListBody(sNo, sPtid, sPtnm, sSexAge, sLocation, sAccDt, sVfydt, sVfyNm, sTime)
        Next
    End With
    
    Printer.EndDoc
End Sub

Private Sub cmdPrint_Click()
    
    Select Case tabData.SelectedItem.Index
        Case 1:
            If objBase.RecordCount = 0 Then Exit Sub
            If objWeek.RecordCount = 0 Then Exit Sub
            Me.MousePointer = 11
            Call PrintItem_TurnAround
            Me.MousePointer = 0
        Case 2:
            If objOver.RecordCount = 0 Then Exit Sub
            Me.MousePointer = 11
            Call PrintTimeOverList
            Me.MousePointer = 0
    End Select
    
End Sub
