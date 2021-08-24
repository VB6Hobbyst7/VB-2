VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm451AccCnt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "검사항목 통계"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   14490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   Tag             =   "접수건수 통계"
   WindowState     =   2  '최대화
   Begin VB.Frame fraInOut 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   9465
      TabIndex        =   66
      Top             =   -45
      Width           =   2355
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   225
         Index           =   2
         Left            =   1500
         TabIndex        =   69
         Top             =   255
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "외래"
         Height          =   225
         Index           =   1
         Left            =   765
         TabIndex        =   68
         Top             =   255
         Width           =   765
      End
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "입원"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   67
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "검사실"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Index           =   0
      Left            =   15
      TabIndex        =   35
      Top             =   8760
      Visible         =   0   'False
      Width           =   555
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   1590
         TabIndex        =   39
         Top             =   345
         Width           =   795
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "Lis451.frx":0000
         Left            =   2385
         List            =   "Lis451.frx":0016
         Style           =   2  '드롭다운 목록
         TabIndex        =   38
         Tag             =   "검사실"
         Top             =   315
         Width           =   780
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   37
         Top             =   345
         Width           =   1035
      End
      Begin VB.ComboBox cboBuilding 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   270
         Style           =   2  '드롭다운 목록
         TabIndex        =   36
         Top             =   660
         Width           =   2895
      End
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "Work Area"
      ForeColor       =   &H00864B24&
      Height          =   1020
      Index           =   1
      Left            =   75
      TabIndex        =   40
      Top             =   615
      Width           =   3510
      Begin VB.ComboBox cboWorkArea 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   150
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   645
         Width           =   3285
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   23
         Top             =   345
         Width           =   1035
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "Lis451.frx":002F
         Left            =   2640
         List            =   "Lis451.frx":0045
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Tag             =   "WorkArea"
         Top             =   300
         Width           =   780
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   1365
         TabIndex        =   34
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.TextBox txtTestCd 
      Height          =   345
      Left            =   10860
      TabIndex        =   30
      Top             =   1185
      Width           =   960
   End
   Begin VB.TextBox txtDeptCd 
      Height          =   315
      Left            =   7260
      TabIndex        =   29
      Top             =   1185
      Width           =   900
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "의뢰과"
      ForeColor       =   &H00864B24&
      Height          =   1020
      Index           =   3
      Left            =   7155
      TabIndex        =   24
      Top             =   615
      Width           =   3480
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   330
         Left            =   1350
         TabIndex        =   41
         Top             =   570
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         BackColor       =   15463405
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1050
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   31
         Tag             =   "DeptCd"
         Top             =   570
         Width           =   285
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   2
         ItemData        =   "Lis451.frx":005E
         Left            =   2595
         List            =   "Lis451.frx":0074
         Style           =   2  '드롭다운 목록
         TabIndex        =   27
         Tag             =   "의뢰과"
         Top             =   285
         Width           =   780
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   3
         Left            =   1320
         TabIndex        =   26
         Top             =   330
         Width           =   750
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   25
         Top             =   330
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Refresh"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "158"
      Top             =   30
      Width           =   1320
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   1770
      TabIndex        =   12
      Top             =   -45
      Width           =   6735
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   765
         TabIndex        =   0
         Top             =   195
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85131264
         CurrentDate     =   36238
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         Top             =   195
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85131264
         CurrentDate     =   36391
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "To"
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
         Left            =   3405
         TabIndex        =   15
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "From"
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
         Left            =   105
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "검사장비"
      ForeColor       =   &H00864B24&
      Height          =   1020
      Index           =   2
      Left            =   3630
      TabIndex        =   11
      Top             =   615
      Width           =   3480
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   1335
         TabIndex        =   21
         Top             =   360
         Width           =   750
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "Lis451.frx":008D
         Left            =   2610
         List            =   "Lis451.frx":00A3
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Tag             =   "검사장비"
         Top             =   315
         Width           =   795
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   1035
      End
      Begin VB.ComboBox cboEqpCd 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   150
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   645
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "검사항목"
      ForeColor       =   &H00864B24&
      Height          =   1020
      Left            =   10680
      TabIndex        =   10
      Top             =   615
      Width           =   3735
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   345
         Left            =   1455
         TabIndex        =   42
         Top             =   570
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   609
         BackColor       =   15463405
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1155
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   32
         Tag             =   "DeptCd"
         Top             =   555
         Width           =   285
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   28
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   4
         Left            =   1590
         TabIndex        =   22
         Top             =   285
         Width           =   750
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   3
         ItemData        =   "Lis451.frx":00BC
         Left            =   2865
         List            =   "Lis451.frx":00D2
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Tag             =   "검 사 항 목"
         Top             =   255
         Width           =   780
      End
   End
   Begin VB.Frame frmPrgBar 
      BackColor       =   &H00AFBCC5&
      BorderStyle     =   0  '없음
      Caption         =   "                                                                                    "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F5386&
      Height          =   1035
      Left            =   4560
      TabIndex        =   7
      Top             =   4260
      Visible         =   0   'False
      Width           =   6525
      Begin MSComctlLib.ProgressBar Prgbar 
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00A9B4BA&
         BackStyle       =   0  '투명
         Caption         =   "데이터를 로드중 입니다."
         Height          =   180
         Left            =   2355
         TabIndex        =   9
         Top             =   300
         Width           =   1980
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00DBE6E6&
         Height          =   1035
         Left            =   0
         Top             =   0
         Width           =   6525
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   1620
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "132"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&S)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "158"
      Top             =   30
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "127"
      Top             =   8535
      Width           =   1320
   End
   Begin TabDlg.SSTab tabView 
      Height          =   6795
      Left            =   75
      TabIndex        =   44
      Top             =   1650
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   11986
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "테이블"
      TabPicture(0)   =   "Lis451.frx":00EB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTotalCnt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ssDataBuf"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "spdStat"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "그래프"
      TabPicture(1)   =   "Lis451.frx":0107
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape4"
      Tab(1).Control(1)=   "cfxStat"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "cmdGrpShow"
      Tab(1).Control(4)=   "lstSort"
      Tab(1).ControlCount=   5
      Begin VB.ListBox lstSort 
         Height          =   240
         Left            =   -73305
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   60
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.CommandButton cmdGrpShow 
         BackColor       =   &H00D1DCD7&
         Caption         =   "Sho&w"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -62295
         Style           =   1  '그래픽
         TabIndex        =   56
         Tag             =   "158"
         Top             =   600
         Width           =   990
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   585
         Left            =   -74760
         TabIndex        =   46
         Top             =   360
         Width           =   12375
         Begin VB.ComboBox cboSeries 
            BackColor       =   &H00F7FFFF&
            Height          =   300
            ItemData        =   "Lis451.frx":0123
            Left            =   7875
            List            =   "Lis451.frx":0125
            Style           =   2  '드롭다운 목록
            TabIndex        =   52
            Tag             =   "검사실"
            Top             =   195
            Width           =   780
         End
         Begin VB.CheckBox chkTable 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Show Data Table"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10110
            TabIndex        =   51
            Top             =   240
            Width           =   2145
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Name"
            Height          =   240
            Index           =   1
            Left            =   8850
            TabIndex        =   50
            Top             =   240
            Width           =   840
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Count of"
            Height          =   240
            Index           =   0
            Left            =   6810
            TabIndex        =   49
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboXVal 
            Height          =   300
            ItemData        =   "Lis451.frx":0127
            Left            =   3405
            List            =   "Lis451.frx":0129
            Style           =   2  '드롭다운 목록
            TabIndex        =   48
            Top             =   195
            Width           =   1875
         End
         Begin VB.ComboBox cboYVal 
            Height          =   300
            ItemData        =   "Lis451.frx":012B
            Left            =   795
            List            =   "Lis451.frx":013B
            Style           =   2  '드롭다운 목록
            TabIndex        =   47
            Top             =   195
            Width           =   1875
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "Sort By : "
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
            Left            =   5790
            TabIndex        =   55
            Top             =   255
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "가로 : "
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
            Left            =   2805
            TabIndex        =   54
            Top             =   255
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "세로 : "
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
            Left            =   210
            TabIndex        =   53
            Top             =   255
            Width           =   660
         End
      End
      Begin FPSpread.vaSpread spdStat 
         Height          =   5745
         Left            =   570
         TabIndex        =   45
         Top             =   675
         Width           =   13380
         _Version        =   196608
         _ExtentX        =   23601
         _ExtentY        =   10134
         _StockProps     =   64
         BackColorStyle  =   1
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   6
         MaxRows         =   23
         OperationMode   =   1
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   13818331
         SpreadDesigner  =   "Lis451.frx":0161
         TextTip         =   4
      End
      Begin ChartfxLibCtl.ChartFX cfxStat 
         Height          =   5520
         Left            =   -74760
         TabIndex        =   58
         Top             =   975
         Width           =   13455
         _cx             =   2037210293
         _cy             =   2037196297
         Build           =   7
         TypeMask        =   101187586
         Style           =   -1179655
         RightGap        =   23
         TopGap          =   33
         AngleX          =   4
         AngleY          =   69
         View3DDepth     =   20
         MarkerShape     =   5
         MarkerSize      =   2
         Axis(0).MinorStep=   -10
         Axis(0).Max     =   90
         Axis(0).Decimals=   0
         Axis(0).TickMark=   -32767
         Axis(0).MinorTickMark=   -32766
         Axis(2).MinorStep=   -1
         Axis(2).Min     =   0
         Axis(2).Max     =   100
         RGBBk           =   14737632
         RGB2DBk         =   16777216
         nColors         =   10
         TopFontMask     =   268435464
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BottomFontMask  =   268435464
         BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendFontMask  =   268435464
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   10
         nSer            =   10
         NumPoint        =   10
         NumSer          =   10
      End
      Begin FPSpread.vaSpread ssDataBuf 
         Height          =   5745
         Left            =   570
         TabIndex        =   59
         Top             =   675
         Visible         =   0   'False
         Width           =   13335
         _Version        =   196608
         _ExtentX        =   23521
         _ExtentY        =   10134
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   22
         OperationMode   =   1
         ShadowColor     =   13818331
         ShadowDark      =   13818331
         SpreadDesigner  =   "Lis451.frx":08B6
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   1  '투명하지 않음
         Height          =   6480
         Left            =   -74985
         Top             =   315
         Width           =   14370
      End
      Begin VB.Label lblTotalCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12240
         TabIndex        =   61
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ToTal Count : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   10470
         TabIndex        =   60
         Top             =   390
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   1  '투명하지 않음
         Height          =   6480
         Left            =   15
         Top             =   315
         Width           =   14370
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   4005
      TabIndex        =   43
      Top             =   5220
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   1710
      TabIndex        =   63
      Top             =   2475
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
      SpreadDesigner  =   "Lis451.frx":0B48
   End
   Begin MedControls1.LisLabel lblCondition 
      Height          =   510
      Left            =   75
      TabIndex        =   64
      Top             =   45
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   900
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
      Caption         =   "조회기간"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   510
      Left            =   8520
      TabIndex        =   65
      Top             =   45
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   900
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
      Caption         =   "조회유형"
      Appearance      =   0
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   1  '투명하지 않음
      Height          =   1035
      Left            =   3945
      Top             =   4500
      Visible         =   0   'False
      Width           =   6525
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00A9B4BA&
      BackStyle       =   0  '투명
      Caption         =   "데이터를 로드중 입니다."
      Height          =   180
      Left            =   6300
      TabIndex        =   62
      Top             =   4800
      Visible         =   0   'False
      Width           =   1980
   End
End
Attribute VB_Name = "frm451AccCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_BUILDING = 5
Private Const COL_WORKAREA = 1
Private Const COL_EQUIPMENT = 2
Private Const COL_DEPTNM = 3
Private Const COL_TESTNM = 4
Private Const COL_COUNT = 6
Private Const COL_SERIES = 7
Private Const COL_POINTS = 8

Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private objSQL                  As New clsLISSqlStatistic
Dim rsDeptStat          As Recordset
Dim rsTestStat          As Recordset
Dim rsDeptTestStat      As Recordset
Dim QueryFlag           As Boolean
Dim MsgFg               As Boolean

Dim ColWid(6)           As Double
Dim SortKeys(6)         As Integer
Dim totCnt              As Long
Dim GrpColor(100)       As Long
Dim iPrgbarCount        As Long
Dim SubTot(6)           As Long

'Workarea별 검사코드 담아주는 Dictionary
Private objDic As clsDictionary
Public Event LastFormUnload()

Private Sub cboBuilding_Click()
    Call LoadEqpList
End Sub

Private Sub cboWorkArea_Click()
    Dim RS        As Recordset
    Dim sWorkarea As String
    Dim SSQL      As String
    
    Set objDic = New clsDictionary
    objDic.Clear
    objDic.FieldInialize "testcd", "testnm"
    
    If cboWorkArea.ListIndex < 0 Then Exit Sub
    
    sWorkarea = medGetP(cboWorkArea.Text, 1, " ")
    
    Set RS = New Recordset
    SSQL = objSQL.GetWorkareaTestItem(sWorkarea)
    
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            If objDic.Exists(RS.Fields("testcd").Value & "") = False Then
                objDic.AddNew RS.Fields("testcd").Value & "", RS.Fields("testnm").Value & ""
            End If
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub

Private Sub cboXVal_Click()
    cboSeries.Clear
End Sub

Private Sub cboYVal_Click()

    With cboXVal
        .Clear
        Select Case cboYVal.ListIndex
            Case 0: '전체
'                .AddItem "검사실":    .ItemData(0) = COL_BUILDING
                .AddItem "Work Area": .ItemData(0) = COL_WORKAREA
                .AddItem "검사장비":  .ItemData(1) = COL_EQUIPMENT
                .AddItem "의뢰과":    .ItemData(2) = COL_DEPTNM
                .AddItem "검사항목":  .ItemData(3) = COL_TESTNM
            Case 1: 'worarea
'                .AddItem "Work Area": .ItemData(0) = COL_WORKAREA
                .AddItem "검사장비":  .ItemData(0) = COL_EQUIPMENT
                .AddItem "의뢰과":    .ItemData(1) = COL_DEPTNM
                .AddItem "검사항목":  .ItemData(2) = COL_TESTNM
            Case 2: '검사장비
                .AddItem "검사항목":  .ItemData(0) = COL_TESTNM
            Case 3: '의뢰과
                .AddItem "검사항목":  .ItemData(0) = COL_TESTNM
        End Select
    End With
    cboSeries.Clear

End Sub

Private Sub chkAll_Click(Index As Integer)
    Dim ChkValue As Boolean

    ChkValue = IIf(chkAll(Index).Value = 0, True, False)
    Select Case Index
    Case 0:
        cboBuilding.Enabled = ChkValue
    Case 1:
        cboWorkArea.Enabled = ChkValue
    Case 2:
        cboEqpCd.Enabled = ChkValue
    Case 3:
        txtDeptCd.Text = ""
        txtDeptCd.Enabled = ChkValue
        cmdHelpList(0).Enabled = ChkValue
        lblDeptNm.Caption = ""
    Case 4:
        txtTestCd.Text = ""
        txtTestCd.Enabled = ChkValue
        cmdHelpList(1).Enabled = ChkValue
        lblTestNm.Caption = ""
    End Select
End Sub

Private Sub chkTable_Click()
    If chkTable.Value = 1 Then
        cfxStat.DataEditor = True
        cfxStat.DataEditorObj.BkColor = &HE0E0E0
'        cfxStat.DataEditorObj.Docked = TGFP_BOTTOM 2001/04/18
        cfxStat.DataEditorObj.AutoSize = True
        cfxStat.DataEditorObj.Font = "돋움"
        cfxStat.DataEditorObj.SizeToFit
    Else
        cfxStat.DataEditor = False
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    dtpStart.Value = Now
    dtpEnd.Value = Now
    dtpStart.SetFocus
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp      As String
    
    With spdStat
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "AccCount"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)

End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdGrpShow_Click()
    If cboXVal.ListIndex < 0 Or cboYVal.ListIndex < 0 Then Exit Sub
    Me.MousePointer = 13
    Call clearcfx(cfxStat)
    DoEvents
    Call ShowGraph
    Me.MousePointer = 0
End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Dim objDept As clsBasisData
    
'    Set objDept = New clsBasisData
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        Select Case Index
            Case 0:
                .FormCaption = "진료과 조회"
                .ColumnHeaderText = "진료과코드;진료과명"
                Call .LoadPopUp(GetSQLDeptList) ', 3400, 6500) ', ObjLISComCode.DeptCd)
                txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
                lblDeptNm.Caption = medGetP(.SelectedString, 2, ";")
            Case 1:
                .FormCaption = "검사항목 조회"
                .ColumnHeaderText = "검사항목코드;검사명"
                Call .LoadPopUp(objSQL.GetAccTest) ', 3400, 9800)
                txtTestCd.Text = medGetP(.SelectedString, 1, ";")
                lblTestNm.Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
    
'    Set objDept = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdRefresh_Click()
    Call clearcfx(cfxStat)
    Call ShowData
End Sub

Private Sub cmdStart_Click()

    Dim sStartDate As String, sEndDate As String

    If dtpStart.Value > dtpEnd.Value Then
        MsgBox "Duration input Error"
        Exit Sub
    End If

    sStartDate = Format(dtpStart.Value, CS_DateDbFormat)
    sEndDate = Format(dtpEnd.Value, CS_DateDbFormat)
    Me.MousePointer = 11
'    MouseRunning   2001/04/18
    QueryFlag = ReadData  ' True 이면 조회가 이루어 졌음을 의미
'    MouseDefault   2001/04/18

    If QueryFlag Then
        chkAll(1).Enabled = False
        chkAll(2).Enabled = False
        chkAll(3).Enabled = False
        chkAll(4).Enabled = False
        dtpStart.Enabled = False
        dtpEnd.Enabled = False
        cmdStart.Enabled = False

        cmdRefresh.Enabled = True
        cmdGrpShow.Enabled = True
        cmdPrint.Enabled = True
        cmdExcel.Enabled = True
        Call cmdRefresh_Click
    Else
        MsgBox "해당 자료가 없습니다...", vbInformation
    End If
    Me.MousePointer = 0

End Sub


Private Sub dtpEnd_Validate(Cancel As Boolean)
    Call clearspdStat
    Call clearcfx(cfxStat)
End Sub

Private Sub dtpStart_Validate(Cancel As Boolean)
    Call clearspdStat
    Call clearcfx(cfxStat)
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim i As Integer

'    Me.Show
    Call ClearRtn

    DoEvents

    Call LoadBuildingList
    Call LoadWorkAreaList
    Call LoadEqpList

    dtpStart.Value = Format(Now, "yyyy-mm-dd")
    dtpEnd.Value = Format(Now, "yyyy-mm-dd")

    chkAll(0).Value = 1
    chkAll(1).Value = 1
    chkAll(2).Value = 1
    chkAll(3).Value = 1
    chkAll(4).Value = 1

    ColWid(1) = 16
    ColWid(2) = 18
    ColWid(3) = 18
    ColWid(4) = 38.5
    ColWid(5) = 0

    GrpColor(0) = &HCC99FF
    GrpColor(1) = &HFF99CC
    GrpColor(2) = &H8080FF
    GrpColor(3) = &HFFCC00
    GrpColor(4) = &HDF6A3E     '&H864B24
    GrpColor(5) = &HFFFF
    GrpColor(6) = &H808066
    GrpColor(7) = &HFF9999
    GrpColor(8) = &H663399
    GrpColor(9) = &H0


    MsgFg = True
    For i = 1 To cboSort.Count
        cboSort(i - 1).ListIndex = i
    Next
    MsgFg = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    QueryFlag = False
    
    Set objSQL = Nothing
    Set objDic = Nothing
    Set rsDeptStat = Nothing
    Set rsTestStat = Nothing
    Set rsDeptTestStat = Nothing

End Sub

Private Sub clearspdStat()
    With spdStat
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxRows = 0
    End With
End Sub

Private Sub clearcfx(Ccfx As ChartFX)
    With Ccfx
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
    End With
End Sub

Public Sub LoadBuildingList()
    Dim tmpRs   As Recordset
    Dim i       As Integer
    Dim SqlStmt As String

    Set tmpRs = New Recordset
    SqlStmt = objSQL.GetBuildCd
    
    tmpRs.Open SqlStmt, DBConn
    
    cboBuilding.Clear
    For i = 1 To tmpRs.RecordCount
        cboBuilding.AddItem Trim("" & tmpRs.Fields("BuildCd").Value) & "   " & _
                            Trim("" & tmpRs.Fields("BuildNm").Value)
        tmpRs.MoveNext
    Next

    Set tmpRs = Nothing
    Set objSQL = Nothing
    
    If cboBuilding.ListCount > 0 Then cboBuilding.ListIndex = 0 'medComboFind(cboBuilding, objSysInfo.BuildingCd)
    
End Sub

Private Sub LoadWorkAreaList()
    Dim rsGetWA     As Recordset
    Dim sSqlGetWA   As String
    Dim i           As Integer
    
    Set rsGetWA = New Recordset
    rsGetWA.Open objSQL.GetWACd, DBConn

    cboWorkArea.Clear
    For i = 1 To rsGetWA.RecordCount
        cboWorkArea.AddItem "" & rsGetWA.Fields("WACd").Value & " " & _
                            "" & rsGetWA.Fields("WANm").Value
        rsGetWA.MoveNext
    Next i

    Set rsGetWA = Nothing

End Sub

Public Sub LoadEqpList()
    Dim rsEQCode        As Recordset
    Dim sSqlGetEQCode   As String
    Dim strBldCd        As String
    Dim i               As Integer
    
    sSqlGetEQCode = objSQL.GetEqpCd
    If cboBuilding.ListIndex > 0 Then
        strBldCd = medGetP(cboBuilding.Text, 1, " ")
        sSqlGetEQCode = objSQL.GetEqpCd(False, strBldCd)
    End If
 
    Set rsEQCode = New Recordset
    rsEQCode.Open sSqlGetEQCode, DBConn

    With cboEqpCd
        .Clear
        .AddItem "수작업"
        For i = 0 To rsEQCode.RecordCount - 1
            .AddItem "" & rsEQCode.Fields("EqpCd").Value & "   " & _
                     "" & rsEQCode.Fields("EqpNm").Value
            rsEQCode.MoveNext
        Next i
    End With

    Set rsEQCode = Nothing
End Sub

Private Function ReadData() As Boolean
    Dim objProBar   As jProgressBar.clsProgress
    Dim RS          As Recordset
    Dim sInOut      As String               '입원 / 외래 구분조회변수
    Dim SqlStmt     As String
    Dim strTmp      As String               '진료과 초기 변수
    Dim sWorkarea   As String               'workarea
    Dim sWorkareaNm As String               'workarea 명
    Dim BlnFG       As Boolean              '처음의 진료과를 건너뛰기 위한 선언
    
    Dim i           As Integer
    Dim j           As Long
    Dim kk          As Long                 '마지막 진료과를 담기위한 변수
    
    ReadData = False
    lblMsg = "입력된 기간동안의 검사건수를 집계하고 있습니다..."
   
    Set objProBar = New jProgressBar.clsProgress
    With objProBar
        .Container = Me
        .Left = tabView.Left + 1700
        .Top = tabView.Top + 30 '
        .Width = (tabView.Width - 1700)
        .Height = 280
        .Message = "자료를 읽기 위해 준비중입니다..."
'        .SetMyForm Me
'        .Choice = True
'        .XPos = tabView.Left + 1700
'        .YPos = tabView.Top + 30 '
'        .XWidth = (tabView.Width - 1700)
'        .ForeColor = &H864B24
'        .Appearance = aPlate
'        .YHeight = 280
'        .Msg = "자료를 읽기 위해 준비중입니다..."
'        .Value = 1
        DoEvents
    End With
    
    sInOut = ""
    
    '입원/외래 구분 조회
    If optInOut(0).Value = True Then sInOut = "2"
    If optInOut(1).Value = True Then sInOut = "1"
    
    'workarea 별로 검색할시
    objSQL.WorkArea = ""
    If chkAll(1).Value = 0 And cboWorkArea.ListIndex > -1 Then objSQL.WorkArea = medGetP(cboWorkArea.Text, 1, " ")
    
    SqlStmt = objSQL.GetAccCnt_Bussdiv(dtpStart.Value, dtpEnd.Value, sInOut)
    
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If RS.EOF Then
        Set RS = Nothing
        Set objProBar = Nothing
        Exit Function
    End If
    
'========================================================================================
'  WorkArea별로 선택해서 조회하였을경우의 Flow
'  Workarea 별로 모든 검사항목을 진료과별로 보여준다.
'========================================================================================
    If chkAll(1).Value = 0 And cboWorkArea.ListIndex > -1 Then
        With objProBar
            .Max = RS.RecordCount * 2
            .DisplayMessage = False
        End With
            
        Dim objShow As clsDictionary
        
        Set objShow = New clsDictionary
        

        objShow.Clear
        objShow.FieldInialize "workarea,deptnm,testcd", "cnt,eqpcd,testnm,buildcd"
        
        objShow.Sort = False

        
        sWorkarea = medGetP(cboWorkArea, 1, " ")
        sWorkareaNm = medGetP(cboWorkArea, 2, " ")
        
        
        strTmp = ""
        BlnFG = False
        kk = 0
        With RS
            .MoveFirst
            Do Until RS.EOF
                kk = kk + 1
                
                If strTmp <> .Fields("deptnm").Value & "" Then
                    If BlnFG = True Then
                        objDic.MoveFirst
                        Do Until objDic.EOF
                            If objShow.Exists(sWorkarea & COL_DIV & strTmp & COL_DIV & objDic.Fields("testcd")) = False Then
                                objShow.AddNew sWorkarea & COL_DIV & strTmp & COL_DIV & objDic.Fields("testcd"), _
                                               "" & COL_DIV & "" & COL_DIV & objDic.Fields("testnm") & COL_DIV & "" & COL_DIV & ""
                            End If
                            objDic.MoveNext
                        Loop
                    End If
                    BlnFG = True
                    strTmp = RS.Fields("deptnm").Value & ""
                    
                Else
                    If kk = .RecordCount Then
                        objDic.MoveFirst
                        '검사항목 담는게 빠졌당.
                        
                        Do Until objDic.EOF
                            If objShow.Exists(sWorkarea & COL_DIV & strTmp & COL_DIV & objDic.Fields("testcd")) = False Then
                                objShow.AddNew sWorkarea & COL_DIV & strTmp & COL_DIV & objDic.Fields("testcd"), _
                                               "" & COL_DIV & "" & COL_DIV & objDic.Fields("testnm") & COL_DIV & "" & COL_DIV & ""
                            End If
                            objDic.MoveNext
                        Loop
                    End If
                End If
                objProBar.Value = kk
                RS.MoveNext
            Loop
            RS.MoveFirst
            Do Until .EOF
                If objShow.Exists(sWorkarea & COL_DIV & .Fields("deptnm").Value & "" & COL_DIV & .Fields("testcd").Value & "") Then
                    objShow.KeyChange sWorkarea & COL_DIV & .Fields("deptnm").Value & COL_DIV & .Fields("testcd").Value & ""
                    objShow.Fields("cnt") = Val(objShow.Fields("cnt")) + Val(.Fields("cnt").Value & "")
                    objShow.Fields("eqpcd") = cboEqpCd.List(medComboFind(cboEqpCd, "" & .Fields("eqpcd").Value))
                    objShow.Fields("testnm") = .Fields("testnm").Value & ""
                    objShow.Fields("buildcd") = .Fields("buildcd").Value & ""
                End If
                kk = kk + 1
                objProBar.Value = kk
                .MoveNext
            Loop
        End With
        
        With ssDataBuf
            .MaxRows = 0
            objProBar.Value = 1
            objProBar.Max = objShow.RecordCount
            
            objShow.MoveFirst
            Do Until objShow.EOF
                
                j = j + 1
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1: .Value = cboWorkArea.Text
                .Col = 2: .Value = objShow.Fields("eqpcd")

                .Col = 3: .Value = "" & objShow.Fields("deptnm")

                .Col = 4: .Value = "" & objShow.Fields("testcd")
                          .Value = .Value & Space(10 - Len(.Value)) & "" & objShow.Fields("testnm")
                .Col = 5: .Value = objShow.Fields("buildcd")
                
                .Col = 6: .Value = "" & objShow.Fields("cnt")

                Me.MousePointer = 11
                DoEvents
                objProBar.Value = j
                objShow.MoveNext
            Loop
            Me.MousePointer = 0
        End With
    Else
        '==================
        '일반 조건의 검색
        '==================
        With objProBar
            .Max = RS.RecordCount
            .DisplayMessage = False
        End With
        
        With ssDataBuf
            .MaxRows = 0
            Do Until RS.EOF
                j = j + 1
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 5
                    i = medComboFind(cboBuilding, "" & RS.Fields("buildcd").Value)
                    .Value = cboBuilding.List(i)
                .Col = 1
                    i = medComboFind(cboWorkArea, "" & RS.Fields("workarea").Value)
                    .Value = cboWorkArea.List(i)
                .Col = 2
                If Trim("" & RS.Fields("eqpcd").Value) = "" Then
                    .Value = "수작업"
                Else
                    i = medComboFind(cboEqpCd, "" & RS.Fields("eqpcd").Value)
                    .Value = cboEqpCd.List(i)
                End If
    
                .Col = 3: .Value = "" & RS.Fields("deptnm").Value
    
                .Col = 4: .Value = "" & RS.Fields("testcd").Value
                          .Value = .Value & Space(10 - Len(.Value)) & "" & RS.Fields("testnm").Value
    
                .Col = 6: .Value = "" & RS.Fields("cnt").Value
    
    
                objProBar.Value = j
                DoEvents
                RS.MoveNext
            Loop
        End With
    End If

    Set RS = Nothing
    Set objShow = Nothing
    Set objProBar = Nothing
    ReadData = True
End Function

Private Sub cboSort_Click(Index As Integer)

    Dim i As Integer
    Dim j As Integer

    j = Val(cboSort(Index).Tag)
    If cboSort(Index).ListIndex = 0 Then
        chkSubTot(Index).Value = 0
        Exit Sub
    End If

    cboSort(Index).Tag = cboSort(Index).ListIndex
    SortKeys(cboSort(Index).ListIndex) = Index + 1

    If MsgFg Then Exit Sub
    MsgFg = True
    For i = 0 To cboSort.Count - 1
        If i <> Index Then
            If Val(cboSort(i).Tag) = cboSort(Index).ListIndex Then
                If cboSort(i).ListIndex > 0 Then
                    cboSort(i).ListIndex = j
                Else
                    cboSort(i).Tag = j
                    SortKeys(j) = i + 1
                End If
            End If
        End If
    Next
    MsgFg = False
End Sub

Private Sub ShowData()
    Dim K(6)    As String
    Dim FirstFg As Boolean
    Dim i       As Integer

    FirstFg = True
    K(1) = "": K(2) = "": K(3) = ""
    K(4) = "": K(5) = "": K(6) = ""
    SubTot(1) = 0: SubTot(2) = 0
    SubTot(3) = 0: SubTot(4) = 0
    totCnt = 0

    With ssDataBuf
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .SortKey(1) = SortKeys(1)
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKey(2) = SortKeys(2)
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKey(3) = SortKeys(3)
        .SortKeyOrder(3) = SortKeyOrderAscending
        .SortKey(4) = SortKeys(4)
        .SortKeyOrder(4) = SortKeyOrderAscending
        .SortKey(5) = SortKeys(5)
        .SortKeyOrder(5) = SortKeyOrderAscending
        .SortKey(6) = SortKeys(6)
        .SortKeyOrder(6) = SortKeyOrderAscending
        .SortBy = SortByRow
        .Action = ActionSort
        .BlockMode = False

        spdStat.MaxRows = 0
        spdStat.Row = 0

        For i = 0 To cboSort.Count - 1
            spdStat.Col = Val(cboSort(i).Tag)
            If cboSort(i).ListIndex = 0 Then
                spdStat.ColHidden = True
            Else
                spdStat.ColHidden = False
            End If
            spdStat.ColWidth(spdStat.Col) = ColWid(i + 1)
        Next

        .Row = 0
        Call SetValue(1, K(1))
        Call SetValue(2, K(2))
        Call SetValue(3, K(3))
        Call SetValue(4, K(4))
        Call SetValue(5, K(5))

        For i = 1 To .MaxRows
            .Row = i

            If i > 0 Then
                If chkAll(0).Value = 0 Then
                    .Col = 5
                    If .Value <> cboBuilding.Text Then GoTo Skip
                End If
                If chkAll(1).Value = 0 Then
                    .Col = 1
                    If .Value <> cboWorkArea.Text Then GoTo Skip
                
                End If
                If chkAll(2).Value = 0 Then
                    .Col = 2
                    If .Value <> cboEqpCd.Text Then GoTo Skip
                End If
                
                If chkAll(3).Value = 0 Then
                    .Col = 3
                    If .Value <> Trim(lblDeptNm.Caption) Then GoTo Skip
                End If
                
                If chkAll(4).Value = 0 Then
                    .Col = 4
                    If medGetP(.Value, 1, Space(8 - Len(Trim(txtTestCd.Text)))) <> txtTestCd.Text Then GoTo Skip
                End If
                

            End If

            .Col = SortKeys(1)
            If K(1) <> .Value Then
                If Not FirstFg Then
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
                    If chkSubTot(SortKeys(1)).Value = 1 Then Call SetSubTot(1)
                End If
                If cboSort(SortKeys(1) - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(1, K(1))
                    Call SetValue(2, K(2))
                    Call SetValue(3, K(3))
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If
            .Col = SortKeys(2)
            If K(2) <> .Value Then
                If Not FirstFg Then
                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
                End If
                If cboSort(SortKeys(2) - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(2, K(2))
                    Call SetValue(3, K(3))
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If
            .Col = SortKeys(3)
            If K(3) <> .Value Then
                If Not FirstFg Then
                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
                End If
                If cboSort(SortKeys(3) - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(3, K(3))
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If
            .Col = SortKeys(4)
            If K(4) <> .Value Then
                If Not FirstFg Then
                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                End If
                If cboSort(.Col - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If

            .Col = SortKeys(5)
            If K(5) <> .Value Then
                'If chkSubTot(.Col - 1).Value = 1 Then Call SetSubTot(5)
                If cboSort(.Col - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(5, K(5))
                End If
            End If

            .Col = 6: spdStat.Col = 6
            spdStat.Value = Val(spdStat.Value) + Val(.Value)
            
            
            SubTot(1) = SubTot(1) + Val(.Value)
            SubTot(2) = SubTot(2) + Val(.Value)
            SubTot(3) = SubTot(3) + Val(.Value)
            SubTot(4) = SubTot(4) + Val(.Value)
            SubTot(5) = SubTot(5) + Val(.Value)
            FirstFg = False

            totCnt = totCnt + Val(.Value)
        
Skip:
        Next

    End With
    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
    If chkSubTot(SortKeys(1)).Value = 1 Then Call SetSubTot(1)

    lblTotalCnt.Caption = Format(totCnt, "###,###,###,###")
    tabView.Tab = 0
    spdStat.SetFocus

End Sub

Private Sub SetValue(ByVal Col As Integer, ByRef SvVal As String)
    With ssDataBuf
        .Col = SortKeys(Col)
        spdStat.Col = Col
        spdStat.Value = .Value
        SvVal = .Value
    End With
End Sub

Private Sub SetSubTot(ByVal Col As Integer)
    Dim lngColor As Long
    
    With spdStat
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = Col
        .Value = "소  계"
        lngColor = .BackColor
        .Col = 6
        .Value = SubTot(Col)
        .Col = Col: .COL2 = .MaxCols
        .Row = .Row: .Row2 = .Row
        .BlockMode = True
        .BackColor = &HEEEEEE        'lngColor
        .ForeColor = &HB9602F
        .CellBorderStyle = CellBorderStyleDot
        .CellBorderType = 8  '16
        .Action = ActionSetCellBorder
        '.FontBold = True
        .BlockMode = False
        SubTot(Col) = 0
    End With
End Sub

Private Sub optSort_Click(Index As Integer)
    If Index = 0 Then
        cboSeries.Enabled = True
    Else
        cboSeries.Enabled = False
    End If
End Sub

Private Sub spdStat_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row = 0 Then Exit Sub
    If Col = Val(cboSort(4).Tag) Then
        spdStat.Row = Row
        spdStat.Col = Col
        If spdStat.Value = "소  계" Or Trim(spdStat.Value) = "" Then
            ShowTip = False
            Exit Sub
        End If
        MultiLine = 1
        TipText = "  " & spdStat.Value
        TipWidth = 3000
        spdStat.TextTipDelay = 200
        'Call spdStat.SetTextTipAppearance("굴림", 9, False, False, &HEEFDF2, vbBlue)    '&H996666)
        Call spdStat.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
        ShowTip = True
    Else
        ShowTip = False
    End If
End Sub

Private Sub ShowGraph()
    Dim K(2)    As String
    Dim tmpStr  As String
    
    Dim FirstFg As Boolean
    
    Dim iSeries As Integer
    Dim iPoints As Integer
    Dim iSS     As Integer
    Dim iPT     As Integer
    Dim i       As Integer
    Dim j       As Integer
    
    Dim iCnt    As Long
    Dim iVal    As Long
    

    FirstFg = True
    K(1) = "": K(2) = ""
    iSeries = 0: iPoints = 0    ': iMaxValue = 100

    iSeries = GetCount(cboYVal.ItemData(cboYVal.ListIndex), 7)
    
    iPoints = GetCount(cboXVal.ItemData(cboXVal.ListIndex), 8)

    Call InitDraw(iSeries, iPoints)
    
    Call AddToSortList(iPoints)

    With ssDataBuf
        cfxStat.Title(CHART_TOPTIT) = cboYVal.Text & "  :  " & cboXVal.Text
        cfxStat.ClearData CD_VALUES
        cfxStat.ClearLegend CHART_LEGEND
        cfxStat.OpenDataEx COD_VALUES, iSeries, iPoints

        cfxStat.BottomGap = 20
        If iSeries = 1 Then
            cfxStat.FixedGap = 28 * iSeries
        Else
            cfxStat.FixedGap = 15 * iSeries
        End If
        'cfxStat.Axis(AXIS_Y).Max = iMaxValue

        cfxStat.Scrollable = True
        cfxStat.PointLabels = True
        cfxStat.Axis(AXIS_X).STEP = 1
        cfxStat.Axis(AXIS_X).Decimals = 0
        'cfxStat.PointLabelsFont.Bold = False

        Call SetSerLeg
        Call SetLegend
        Call chkTable_Click

        For i = 0 To iSeries - 1
            cfxStat.Series(i).Color = GrpColor(i)
        Next

        For i = lstSort.ListCount To 1 Step -1
            tmpStr = lstSort.List(i - 1)
            iPT = Val(medGetP(tmpStr, 3, ":"))

            If optSort(0).Value Then
                cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i) = medGetP(tmpStr, 2, ":")
            Else
                cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i) = medGetP(tmpStr, 1, ":")
            End If
            cfxStat.Legend(lstSort.ListCount - i) = cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i)

            For j = 1 To .MaxRows
                .Row = j
                .Col = COL_POINTS
                If Val(.Value) - 1 = iPT Then
                    .Col = COL_SERIES:  iSS = Val(.Value)
                    .Col = COL_COUNT:   iCnt = Val(.Value)
                    iVal = cfxStat.ValueEx(iSS - 1, lstSort.ListCount - i)
                    cfxStat.ValueEx(iSS - 1, lstSort.ListCount - i) = iVal + iCnt
                    cfxStat.KeyLeg(iPT) = cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i)
                    .Col = cboYVal.ItemData(cboYVal.ListIndex)
                    cfxStat.SerLeg(iSS - 1) = .Value
                End If
            Next
        Next

        cboSeries.Clear
        For i = 1 To iSeries
            cboSeries.AddItem cfxStat.SerLeg(i - 1)
        Next

        'cfxStat.Axis(AXIS_Y).Max = iMaxValue + 1
        cfxStat.CloseData COD_VALUES + COD_SCROLLLEGEND

    End With

End Sub

Private Sub InitDraw(ByVal nSeries As Integer, ByVal nPoints As Integer)

    Dim iMaxValue   As Long
    Dim iCnt        As Long
    Dim iVal        As Long
    Dim iSS         As Integer
    Dim iPT         As Integer
    Dim i           As Integer

    With ssDataBuf

        cfxStat.ClearData CD_VALUES
        cfxStat.OpenDataEx COD_VALUES, nSeries, nPoints

        For i = 0 To .MaxRows - 1
            .Row = i + 1
            .Col = COL_SERIES:  iSS = Val(.Value)
            .Col = COL_POINTS:  iPT = Val(.Value)
            .Col = COL_COUNT:   iCnt = Val(.Value)

            iVal = cfxStat.ValueEx(iSS - 1, iPT - 1)
            
            cfxStat.ValueEx(iSS - 1, iPT - 1) = iVal + iCnt

            iVal = cfxStat.ValueEx(iSS - 1, iPT - 1)
            
            If iMaxValue < iVal Then iMaxValue = iVal

            .Col = cboXVal.ItemData(cboXVal.ListIndex)
            'cfxStat.Axis(AXIS_X).Label(iPT - 1) = .Value
            cfxStat.Legend(iPT - 1) = .Value
        Next i

        cfxStat.Axis(AXIS_Y).Max = iMaxValue + 1

    End With

End Sub

Private Sub AddToSortList(ByVal nPoints As Integer)

    Dim tmpStr  As String
    Dim i       As Integer
    Dim nSeries As Integer

    lstSort.Clear
    If cboSeries.ListCount = 0 Or cboSeries.ListIndex < 0 Then
        nSeries = 0
    Else
        nSeries = cboSeries.ListIndex
    End If
    With cfxStat
        For i = 0 To nPoints - 1
            If optSort(0).Value Then
                tmpStr = Format(.ValueEx(nSeries, i), "0#####")
                tmpStr = tmpStr & ":" & .Legend(i)
                'tmpStr = tmpStr & ":" & .Axis(AXIS_X).Label(i)
            Else
                'tmpStr = .Axis(AXIS_X).Label(i)
                tmpStr = .Legend(i)
                tmpStr = tmpStr & ":" & Format(.ValueEx(nSeries, i), "0#####")
            End If
            tmpStr = tmpStr & ":" & Format(i, "0####")
            lstSort.AddItem tmpStr
        Next
    End With

End Sub

Private Sub SetSerLeg()

    With cfxStat
        .SerLegBox = True
'        .SerLegBoxObj.Docked = TGFP_TOP    2001/04/18
        .SerLegBoxObj.Height = 100
'        .SerLegBoxObj.Sizeable = BAS_ALWAYS    2001/04/18
        .SerLegBoxObj.BkColor = &HE0E0E0   '&HD2D9DB   '&HD1DCD7
        .SerLegBoxObj.SizeToFit
    End With

End Sub

Private Sub SetLegend()

    With cfxStat
        .LegendBox = True
        .LegendBoxObj.AutoSize = True
        .LegendBoxObj.Moveable = True
'        .LegendBoxObj.Docked = TGFP_RIGHT  2001/04/18
        .LegendBoxObj.Width = 100
'        .LegendBoxObj.Sizeable = BAS_ALWAYS 2001/04/18
        '.LegendBoxObj.FontMask = CF_SMALLFONTS
        .LegendBoxObj.BkColor = &HE0E0E0   '&HD2D9DB  '&HD1DCD7
        .LegendBoxObj.Font = "smallfonts"
        '.Axis(AXIS_X).PixPerUnit = 30
        'cfxStat.LegendBoxObj.SizeToFit
    End With

End Sub

Private Function GetCount(ByVal iCol As Integer, ByVal iCol1 As Integer) As Integer

    Dim i       As Integer
    Dim iCount  As Integer
    Dim K       As String

    If iCol = 0 Then
        GetCount = 1
        ssDataBuf.Row = 1: ssDataBuf.Row2 = ssDataBuf.MaxRows
        ssDataBuf.Col = iCol1: ssDataBuf.COL2 = iCol1
        ssDataBuf.BlockMode = True
        ssDataBuf.Value = 1
        ssDataBuf.BlockMode = False
        Exit Function
    End If

    iCount = 0: K = ""
    With ssDataBuf
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .SortKey(1) = iCol
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortBy = SortByRow
        .Action = ActionSort
        .BlockMode = False

        For i = 1 To .MaxRows
            .Row = i
            .Col = iCol
            If K <> .Value Then
                iCount = iCount + 1
                K = .Value
            End If
            .Col = iCol1
            .Value = iCount
        Next
    End With

    GetCount = iCount

End Function

Private Sub ClearRtn()

    Dim i As Integer

    ssDataBuf.MaxRows = 0
    spdStat.MaxRows = 0
    lblTotalCnt.Caption = ""

    For i = 0 To chkAll.Count - 1
        chkAll(i).Value = 1
        chkSubTot(i).Value = 0
        cboSort(i).ListIndex = i + 1
    Next

    Call clearcfx(cfxStat)

    tabView.Tab = 0
    optSort(0).Value = True

    dtpStart.Enabled = True
    dtpEnd.Enabled = True
    cmdStart.Enabled = True
    cmdRefresh.Enabled = False
    cmdGrpShow.Enabled = False
    cmdPrint.Enabled = False
    cmdExcel.Enabled = False
    
    chkAll(1).Enabled = True
    chkAll(2).Enabled = True
    chkAll(3).Enabled = True
    chkAll(4).Enabled = True
End Sub

Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtDeptCd_LostFocus()
    Dim strDeptCd   As String
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    strDeptCd = Trim(txtDeptCd.Text)
    
    If strDeptCd <> "" Then
'        Set objDept = New clsBasisData
        strDept = GetDeptNm(strDeptCd)
'        Set objDept = Nothing
        If strDept = "" Then
            MsgBox "등록되지 않은 진료과입니다. 진료과코드를 확인하십시요!", vbCritical, "입력오류"
            txtDeptCd.Text = ""
            txtDeptCd.SetFocus
        Else
            lblDeptNm.Caption = strDept ' ObjLISComCode.DeptCd.Fields("deptnm")
        End If
        
'        If ObjLISComCode.DeptCd.Exists(strDeptCd) = False Then
'            MsgBox "등록되지 않은 진료과입니다. 진료과코드를 확인하십시요!", vbCritical, "입력오류"
'            txtDeptCd.Text = ""
'            txtDeptCd.SetFocus
'        Else
'            ObjLISComCode.DeptCd.KeyChange strDeptCd
'            lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'        End If
    End If
    
End Sub

Private Sub txtTestCd_LostFocus()

    Dim RS  As Recordset
    
    If Trim(txtTestCd.Text) = "" Then Exit Sub
    
    Set RS = New Recordset
    
    RS.Open objSQL.GetAccTest(Trim(txtTestCd.Text)), DBConn
    
    If RS.RecordCount > 0 Then
        lblTestNm.Caption = RS.Fields("abbrnm10").Value & ""
    Else
        MsgBox "등록되지 않은 검사코드입니다. 검사코드를 확인하십시요!", vbCritical, "입력오류"
        txtTestCd.Text = ""
        txtTestCd.SetFocus
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub cmdPrint_Click()
    Call GeneralCount
    Exit Sub
End Sub


Private Sub GeneralCountHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 10

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("검사건수(접수일) 통계", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    strTmp = Format(dtpStart.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpEnd.Value, "YYYY년 MM월 DD일")
    
    Call Print_Setting("조회기간 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    If optInOut(0).Value Then strTmp = "[ 입 원 ]"
    If optInOut(1).Value Then strTmp = "[ 외 래 ]"
    If optInOut(2).Value Then strTmp = "[ 전 체 ]"
    Call Print_Setting("조회유형 : " & strTmp, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    
    strTmp = "[ 전 체 ]"
    If chkAll(1).Value = 0 Then strTmp = cboWorkArea.Text
    Call Print_Setting("업무영역 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    strTmp = "[ 전 체 ]"
    If chkAll(2).Value = 0 Then strTmp = cboEqpCd.Text
    Call Print_Setting("검사장비 : " & strTmp, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    strTmp = "[ 전 체 ]"
    If chkAll(3).Value = 0 Then strTmp = "[ " & txtDeptCd.Text & " ] " & lblDeptNm.Caption
    Call Print_Setting("의 뢰 과 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    strTmp = "[ 전 체 ]"
    If chkAll(4).Value = 0 Then strTmp = "[ " & txtTestCd.Text & " ] " & lblTestNm.Caption
    Call Print_Setting("검사항목 : " & strTmp, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    strTmp = Format(GetSystemDate, "YYYY년 MM월 DD일")
    Call Print_Setting("출 력 일 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("조회건수 : " & lblTotalCnt.Caption, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
    Call GeneralCountBody("업무영역", "검사장비", "의뢰과", "검사항목", "건수")
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub

Private Sub GeneralCountBody(ByVal sWork As String, ByVal sEqpCd As String, ByVal sDept As String, _
                             ByVal sTest As String, ByVal sCnt As String)
                           
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call GeneralCountHead
    End If
   
    Call Print_Setting(sWork, 5, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sEqpCd, 40, LineSpace, 50, "L", "C", False)
    Call Print_Setting(sDept, 70, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sTest, 110, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sCnt, 180, LineSpace, 30, "L", "C")
End Sub

Private Sub GeneralCount()
    Dim sWork   As String
    Dim sEqpCd  As String
    Dim sDept   As String
    Dim sTest   As String
    Dim sCnt    As String
    
    
    Dim ii          As Integer
    
    If spdStat.DataRowCnt < 1 Then Exit Sub
    
    Call P_PrtSet
    Call GeneralCountHead
    
    With spdStat
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1:   sWork = .Value
            .Col = 2:   sEqpCd = .Value
            .Col = 3:   sDept = .Value
            .Col = 4:   sTest = .Value
            .Col = 6:   sCnt = .Value
            Call GeneralCountBody(sWork, sEqpCd, sDept, sTest, sCnt)
        Next
    End With
    
    Printer.EndDoc
End Sub
