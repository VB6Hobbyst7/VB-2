VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm311QCResultEntry_N 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdUrine 
      Caption         =   "Urine 소견"
      Height          =   495
      Left            =   60
      TabIndex        =   47
      Top             =   8520
      Width           =   1215
   End
   Begin FPSpread.vaSpread tblWorkList 
      Height          =   735
      Left            =   3840
      TabIndex        =   46
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
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
      MaxRows         =   3
      ScrollBars      =   0
      SpreadDesigner  =   "frm311QCResultEntry_N.frx":0000
   End
   Begin VB.CommandButton NExT 
      Caption         =   "다음"
      Height          =   495
      Left            =   8280
      TabIndex        =   45
      Top             =   8535
      Width           =   1215
   End
   Begin VB.CommandButton BeFoRe 
      Caption         =   "이전"
      Height          =   495
      Left            =   6840
      TabIndex        =   44
      Top             =   8535
      Width           =   1215
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   315
      Left            =   75
      TabIndex        =   20
      Top             =   1815
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   556
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
      Caption         =   "◈ Reject 소견"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   660
      Left            =   6810
      TabIndex        =   30
      Top             =   -45
      Width           =   7665
      Begin VB.CheckBox chkWard 
         BackColor       =   &H00DBE6E6&
         Caption         =   "병동"
         Height          =   345
         Left            =   5160
         TabIndex        =   48
         Top             =   210
         Value           =   1  '확인
         Width           =   675
      End
      Begin VB.ComboBox cboWorkarea 
         Height          =   300
         Left            =   45
         Style           =   2  '드롭다운 목록
         TabIndex        =   35
         Top             =   225
         Width           =   2085
      End
      Begin VB.CommandButton cmdPopUnverify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미입력 조회"
         Height          =   480
         Left            =   6150
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   150
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpFromDt 
         Height          =   300
         Left            =   2205
         TabIndex        =   32
         Top             =   225
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   84148225
         CurrentDate     =   37925
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   300
         Left            =   3630
         TabIndex        =   33
         Top             =   225
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   84148225
         CurrentDate     =   37925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "~"
         Height          =   180
         Left            =   3465
         TabIndex        =   34
         Top             =   285
         Width           =   135
      End
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   90
      TabIndex        =   29
      Top             =   8190
      Visible         =   0   'False
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MedControls1.LisLabel LisLabel7 
      Height          =   315
      Left            =   75
      TabIndex        =   19
      Top             =   615
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   556
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
      Caption         =   "◈ 컨트롤 정보"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   18
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblRst 
      Height          =   4230
      Left            =   75
      TabIndex        =   15
      Top             =   3960
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   7461
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   13
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "frm311QCResultEntry_N.frx":0287
      TextTip         =   2
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1920
      Left            =   75
      TabIndex        =   12
      Top             =   2025
      Width           =   4890
      Begin VB.CommandButton cmdPopComment 
         BackColor       =   &H00F4F0F2&
         Height          =   300
         Left            =   4515
         Picture         =   "frm311QCResultEntry_N.frx":09CD
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   1560
         Width           =   300
      End
      Begin VB.TextBox txtComment 
         Height          =   1695
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   165
         Width           =   4425
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   960
      Left            =   75
      TabIndex        =   3
      Top             =   855
      Width           =   14400
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   9735
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검사장비"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   6750
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Lot No."
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   4260
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Level"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   4260
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   525
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "컨트롤 비고"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblControl 
         Height          =   360
         Left            =   1365
         TabIndex        =   5
         Top             =   150
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "6_CAR  28-Carbamazepine Nomal"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblLevel 
         Height          =   360
         Left            =   5580
         TabIndex        =   6
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "High"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblLotNo 
         Height          =   360
         Left            =   8070
         TabIndex        =   7
         Top             =   150
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "123456789012345"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEqp 
         Height          =   360
         Left            =   11055
         TabIndex        =   8
         Top             =   150
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C007  Abbott Axsym"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblMakeCd 
         Height          =   360
         Left            =   1365
         TabIndex        =   9
         Top             =   525
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "동양화학"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRemark 
         Height          =   360
         Left            =   5580
         TabIndex        =   10
         Top             =   525
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "이거 정상적으로 사용할수 있는것이에욤"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   45
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "컨트롤 정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   45
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   525
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "컨트롤 제조사"
         Appearance      =   0
      End
   End
   Begin ChartfxLibCtl.ChartFX cfxResult 
      Height          =   2115
      Left            =   4980
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1815
      Width           =   9495
      _cx             =   2037203308
      _cy             =   2037190291
      Build           =   7
      TypeMask        =   -1794588671
      Style           =   -106979329
      LeftGap         =   38
      RightGap        =   4
      TopGap          =   3
      BottomGap       =   20
      WallWidth       =   8
      View3DDepth     =   60
      TypeEx          =   0
      StyleEx         =   1
      CylSides        =   2
      MarkerShape     =   2
      MarkerSize      =   2
      Volume          =   10
      BorderColor     =   16777263
      Axis(0).MinorStep=   -2
      Axis(0).Min     =   -1
      Axis(0).Max     =   11
      Axis(0).Style   =   10280
      Axis(0).GridColor=   0
      Axis(1).Min     =   0
      Axis(1).Max     =   100
      Axis(1).Decimals=   0
      Axis(1).Style   =   10344
      Axis(1).TickMark=   -32767
      Axis(1).GridColor=   0
      Axis(2).Step    =   1
      Axis(2).MinorStep=   1
      Axis(2).Style   =   14368
      Axis(2).PixPerUnit=   19
      Axis(2).GridColor=   0
      RGBBk           =   14936810
      RGB2DBk         =   16777215
      RGB3DBk         =   16777215
      nColors         =   3
      Pallete         =   "frm311QCResultEntry_N.frx":0A7F
      Colors          =   "frm311QCResultEntry_N.frx":0B63
      LeftFontMask    =   805306368
      BeginProperty LeftFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RightFontMask   =   1879048192
      BeginProperty RightFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TopFontMask     =   268435456
      BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomFontMask  =   268435456
      BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Axis(2).FontMask=   268435456
      BeginProperty Axis(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Axis(0).FontMask=   268435456
      BeginProperty Axis(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FixedFontMask   =   268435456
      BeginProperty FixedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LegendFontMask  =   268435456
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PointLabelsFontMask=   268435456
      BeginProperty PointLabelsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PointFontMask   =   268435456
      BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      nPts            =   20
      nSer            =   3
      NumPoint        =   20
      NumSer          =   3
      Stripes         =   "frm311QCResultEntry_N.frx":0B9B
      Fixed           =   "frm311QCResultEntry_N.frx":0C1B
      _Data_          =   "frm311QCResultEntry_N.frx":0C9B
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   660
      Left            =   75
      TabIndex        =   0
      Top             =   -45
      Width           =   6735
      Begin VB.CheckBox chkBarcode 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드"
         Height          =   180
         Left            =   5790
         TabIndex        =   2
         Top             =   285
         Width           =   870
      End
      Begin VB.TextBox txtAccNo 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1245
         TabIndex        =   1
         Text            =   "123456789012"
         Top             =   180
         Width           =   1590
      End
      Begin MedControls1.LisLabel lblBarNo 
         Height          =   360
         Left            =   4065
         TabIndex        =   4
         Top             =   195
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01-031028-1234"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAccNo 
         Height          =   360
         Left            =   45
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "접수 번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBarcode 
         Height          =   360
         Left            =   2865
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   195
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "바코드 번호"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraLabel 
      BackColor       =   &H00E3EAEA&
      BorderStyle     =   0  '없음
      Height          =   1950
      Left            =   90
      TabIndex        =   21
      Top             =   165
      Width           =   480
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   135
         Left            =   90
         TabIndex        =   22
         Top             =   75
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   238
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "+3SD"
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   135
         Left            =   90
         TabIndex        =   23
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   238
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "+2SD"
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   135
         Left            =   90
         TabIndex        =   24
         Top             =   630
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   238
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "+1SD"
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   120
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   900
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   212
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "Mean"
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   135
         Left            =   120
         TabIndex        =   26
         Top             =   1170
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   238
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "-1SD"
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   135
         Left            =   120
         TabIndex        =   27
         Top             =   1455
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   238
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "-2SD"
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   135
         Left            =   120
         TabIndex        =   28
         Top             =   1740
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   238
         BackColor       =   14936810
         ForeColor       =   14508337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "-3SD"
         LeftGab         =   0
      End
   End
End
Attribute VB_Name = "frm311QCResultEntry_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()

Private objRst As clsQcResult
Private MsgFg As Boolean
Private ChangeFg As Boolean

Private mvarParentHwnd As Long

'** 보류버튼 추가로 인한 변수 By M.G.Choi 2006.09.01
Private Const F_OneRow = 1
Private blnExpect       As Boolean
'-------------------------------
' 2009.02.19 양성현 추가
Private currWorkList As Integer
'-------------------------------

Public Property Let ParentHwnd(ByVal vData As Long)
    mvarParentHwnd = vData
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mvarParentHwnd
End Property

Public Sub CallByExternal(ByVal pAccNo As String)
'외부에서 이화면을 콜할때 사용한다.
    
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    
    If cboWorkarea.ListCount = 0 Then
        Call LoadWorkArea
    End If
    
    dtpFromDt.Value = DateAdd("d", -7, GetSystemDate)
    dtpToDt.Value = GetSystemDate
    
    Call InitForm
    Call ClearGraph
    
    txtAccNo.Text = pAccNo
    
    Call LoadData
End Sub

Private Sub chkBarcode_Click()
    If chkBarcode.Value = 1 Then
        lblAccNo.Caption = "바코드 번호"
        lblBarcode.Caption = "접수 번호"
    Else
        lblAccNo.Caption = "접수 번호" '
        lblBarcode.Caption = "바코드 번호"
    End If
    
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    
    Call InitForm
    Call ClearGraph
End Sub

Private Sub chkWard_Click()
    If chkWard.Value = 0 Then
        cboWorkarea.Enabled = True
        cboWorkarea.ListIndex = 0
    Else
        cboWorkarea.Enabled = False
        cboWorkarea.ListIndex = 15
    End If
End Sub

Private Sub cmdClear_Click()
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    
    Call InitForm
    Call ClearGraph
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
    If IsLastForm Then Call UnloadForm(Me)
'    If IsLastForm Then
'        If mvarParentHwnd <> 0 Then
'            Call SendMessage(mvarParentHwnd, WM_CLOSE, 0&, 0&)
'        End If
'    End If
End Sub

Private Sub cmdPopComment_Click()
    Dim objPop As clsPopUpList
    Dim Rs As Recordset
    Dim strSql As String
    
    If Trim(txtAccNo.Text) = "" Then Exit Sub
    If tblRst.ActiveRow > tblRst.DataRowCnt Then Exit Sub
    
    strSql = " select text1 from " & T_LAB034 & _
             " where " & DBW("cdindex=", LC4_QCRejReason)
             
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF Then
        MsgBox "코멘트 템플릿이 등록되지 않았습니다.", vbExclamation
    Else
        Set objPop = New clsPopUpList
        
        With objPop
            .Recordset = Rs
            .FormCaption = "코멘트 템플릿 조회"
            .FormHeight = 2775
            .FormWidth = 4635
            .ColumnHeaderText = "코멘트"
            .ColumnHeaderWidth = "4155.24"
            .HideSearchTool = True
            
            .LoadPopUp
            
            txtComment.Text = txtComment.Text & medGetP(.SelectedString, 1, .Delimiter)
            
            objRst.Item(tblRst.ActiveRow).TxtFg = "1"
            objRst.Item(tblRst.ActiveRow).RstText = txtComment.Text
            Call tblRst.SetText(7 + F_OneRow, tblRst.ActiveRow, objRst.Item(tblRst.ActiveRow).RstText)
            tblRst.SetFocus
        End With
    End If
    
    Set Rs = Nothing
    Set objPop = Nothing
End Sub

Private Sub cmdPopUnverify_Click()
    Dim objPop As clsPopUpList
    Dim Rs As Recordset
    Dim strSql As String
    Dim strSpcYY As String
    Dim strSpcNo As String
    Dim i   As Integer

    strSql = " select  /*+ ROW(a) */  /*+ INDEX_DESC(a S2LAB201_IDX3) use S2LAB201_IDX3 */ distinct  a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "( a.accdt,3)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "a.accseq") & " as accno, " & _
            " b.ctrlcd" & FUNC_CONCAT & "'   '" & FUNC_CONCAT & "c.ctrlnm as control,a.coldt,a.coltm,a.spcyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "a.spcno") & " as barno,a.workarea,a.accdt,a.accseq,b.ctrlcd,b.levelcd,b.lotno " & _
            " from " & T_LAB201 & " a, " & T_LAB026 & " b, " & T_LAB021 & " c " & _
            " where " & DBW("a.coldt>=", Format(dtpFromDt.Value, CS_DateDbFormat)) & _
            " and " & DBW("a.coldt<=", Format(dtpToDt.Value, CS_DateDbFormat)) & _
            " and a.qcfg='1' " & _
            " and a.stscd<'5' " & _
            IIf(cboWorkarea.ListIndex > 0, " and " & DBW("a.workarea=", Trim(medGetP(cboWorkarea.Text, 2, COL_DIV))), "") & _
            " and a.workArea = b.workArea " & _
            " and a.accdt=b.accdt " & _
            " and a.accseq=b.accseq " & _
            " and b.ctrlcd=c.ctrlcd " & _
            " and b.levelcd=c.levelcd "
            If chkWard.Value = 1 Then
                strSql = strSql & " and c.eqpcd NOT IN ('E067','E068','E084','E070','E071','E072','E073','E074','E083','E069') "
            End If
            strSql = strSql & " order by b.ctrlcd,b.levelcd,b.lotno, a.coldt, a.coltm "
            
' 미입력리스트 조회 소팅 변경 2016.04.14 온승호
'            " order by a.workarea,a.accdt,a.accseq,b.ctrlcd,b.levelcd,b.lotno, a.coltm "
                
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF Then
        MsgBox "미입력된 검사가 없습니다.", vbInformation
    Else
        Set objPop = New clsPopUpList
        objPop.Recordset = Rs
'        objPop.FormWidth = 4635
        objPop.FormCaption = "미입력 리스트 조회"
        objPop.FormWidth = 7965
        objPop.FontName = "굴림체"
        objPop.ColumnHeaderText = "접수번호;컨트롤;채취일자;채취시간;검체번호"
        objPop.ColumnHeaderWidth = "1305.071;2700.284;1170.142;1170.142;1170.142"
        objPop.ColumnHeaderAlign = "0;0;2;2;2"
        '0 왼쪽, 1 오른쪽, 2 가운데
        objPop.LoadPopUp
        
        lblBarNo.Caption = ""
        Call InitForm
        Call ClearGraph
        
        If chkBarcode.Value = 1 Then
            strSpcYY = medGetP(medGetP(objPop.SelectedString, 5, objPop.Delimiter), 1, "-")
            strSpcNo = medGetP(medGetP(objPop.SelectedString, 5, objPop.Delimiter), 2, "-")
            
            txtAccNo.Text = strSpcYY & Format(strSpcNo, LIS_BarFormat)
        Else
            txtAccNo.Text = medGetP(objPop.SelectedString, 1, objPop.Delimiter)
        End If

        Rs.MoveFirst
        With tblWorkList
            .MaxRows = Rs.RecordCount
            For i = 1 To .MaxRows
                .Row = i
            
                If txtAccNo.Text = Rs.Fields("accno").Value & "" Then currWorkList = i
            
                .Col = 1: .Value = Rs.Fields("accno").Value & ""
                .Col = 2: .Value = Rs.Fields("control").Value & ""
                .Col = 3: .Value = Rs.Fields("coldt").Value & ""
                .Col = 4: .Value = Rs.Fields("coltm").Value & ""
                .Col = 5: .Value = Rs.Fields("barno").Value & ""
                .Col = 6: .Value = Rs.Fields("workarea").Value & ""
                .Col = 7: .Value = Rs.Fields("accdt").Value & ""
                .Col = 8: .Value = Rs.Fields("accseq").Value & ""
                .Col = 9: .Value = Rs.Fields("ctrlcd").Value & ""
                .Col = 10: .Value = Rs.Fields("levelcd").Value & ""
                .Col = 11: .Value = Rs.Fields("lotno").Value & ""
                Rs.MoveNext
            Next i
        End With

        Call LoadData
    End If
        
    Set Rs = Nothing
    Set objPop = Nothing
End Sub

Private Sub cmdSave_Click()
    If Trim(txtAccNo.Text) = "" Then Exit Sub
    
    MousePointer = vbHourglass
    
    objRst.l201_vfyid = ObjSysInfo.EmpId
    If objRst.SaveResult(tblRst, ObjSysInfo.EmpId) Then
        txtAccNo.Text = ""
        lblBarNo.Caption = ""
        
        Call InitForm
        Call ClearGraph
'---------------------------
' 2009.02.19 양성현 추가
        Call AfterSave
'---------------------------
    End If
    
    MousePointer = vbDefault
End Sub

Private Sub cmdUrine_Click()
    Dim iCnt As Integer
    Dim varTmp
    
    For iCnt = 1 To tblRst.MaxRows
        tblRst.GetText 1, iCnt, varTmp
        
        If varTmp <> "" Then
            objRst.Item(iCnt).TxtFg = "1"
            objRst.Item(iCnt).RstText = "육안결과일치"
            tblRst.SetText 8, iCnt, "육안결과일치"
        End If
    Next
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    txtAccNo.SetFocus
End Sub

Private Sub Form_Load()
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    
    Call LoadWorkArea
    
    dtpFromDt.Value = DateAdd("d", -7, GetSystemDate)
    dtpToDt.Value = GetSystemDate
    
    Call InitForm
    Call ClearGraph
End Sub

Private Sub LoadWorkArea()
    Dim objWA As clsLISSqlQc
    Dim Rs As Recordset
    
    Set objWA = New clsLISSqlQc
       
    Set Rs = New Recordset
    Rs.Open objWA.GetWorkArea, DBConn
    
    cboWorkarea.Clear
    cboWorkarea.addItem "전 체"
    Do Until Rs.EOF
        cboWorkarea.addItem Format(Rs.Fields("field1").Value & "", "!" & String(50, "@")) & COL_DIV & _
                           Rs.Fields("cdval1").Value & ""
        Rs.MoveNext
    Loop
    
    If cboWorkarea.ListCount > 0 And chkWard.Value = 0 Then
        cboWorkarea.ListIndex = 0
        cboWorkarea.Enabled = True
    Else
        cboWorkarea.ListIndex = 15
        cboWorkarea.Enabled = False
    End If
    Set Rs = Nothing
    Set objWA = Nothing
End Sub

Private Sub InitForm()
    lblControl.Caption = ""
    lblLevel.Caption = ""
    lblLotNo.Caption = ""
    lblEqp.Caption = ""
    lblMakeCd.Caption = ""
    lblRemark.Caption = ""
    txtComment.Text = ""
    blnExpect = False
    
    Call ClearTable
End Sub

Private Sub ClearTable()
    Dim i As Long
    
    Call medClearTable(tblRst)
    tblRst.MaxRows = 12
    
    tblRst.Row = 1: tblRst.Row2 = tblRst.MaxRows
    tblRst.Col = 8 + F_OneRow: tblRst.Col2 = 8 + F_OneRow
    tblRst.BlockMode = True
    tblRst.CellType = CellTypeStaticText
    tblRst.BlockMode = False
    
    
'    With tblRst
'        For i = 1 To tblRst.MaxRows
'            .Row = i
'            .Col = 1
'            .CellType = CellTypeStaticText
'            .TypeHAlign = TypeHAlignLeft
'
'            .Col = 2
'            .CellType = CellTypeEdit
'            .TypeHAlign = TypeHAlignRight
'
'            .Col = 3: .Col2 = 6
'            .CellType = CellTypeStaticText
'            .TypeHAlign = TypeHAlignCenter
'
'            .Col = 7
'            .CellType = CellTypeEdit
'            .TypeHAlign = TypeHAlignLeft
'
'            .Col = 8
'            .CellType = CellTypeStaticText
'            .TypeHAlign = TypeHAlignCenter
'        Next
'    End With
End Sub

Private Sub ClearGraph()
    With cfxResult
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
    End With
    fraLabel.Visible = False
'    If OldRow > 0 Then
'        tblRst.Row = OldRow
'        tblRst.Col = 7
'        tblRst.TypeButtonText = ""
'        OldRow = -1
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRst = Nothing
End Sub

'--------------------------------------------------------
' 2009.02.19 양성현 추가
Private Sub BeFoRe_Click()
    Call BeforeSave
End Sub

Private Sub NExT_Click()
    Call AfterSave
End Sub

Private Sub AfterSave()
    currWorkList = currWorkList + 1
    With tblWorkList
        If .MaxRows < currWorkList Then
            currWorkList = .MaxRows
            MsgBox "마지막 자료 입니다."
        End If
        .Row = currWorkList
        .Col = 1
        txtAccNo.Text = .Value
    End With
    Call LoadData
End Sub

Private Sub BeforeSave()
    currWorkList = currWorkList - 1
    If currWorkList < 1 Then
        currWorkList = 1
        MsgBox "처음 자료 입니다."
    End If
    With tblWorkList
        .Row = currWorkList
        .Col = 1
        txtAccNo.Text = .Value
    End With
    Call LoadData
End Sub

'--------------------------------------------------------

Private Sub tblRst_Advance(ByVal AdvanceNext As Boolean)
'    Dim lngRow As Long
'
'    lngRow = IIf(AdvanceNext, tblRst.DataRowCnt, 1)
'    Call tblRst_LeaveCell(tblRst.ActiveCol, lngRow, tblRst.ActiveCol, lngRow, False)
End Sub

Private Sub tblRst_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Row = 0 Then
        Exit Sub
    End If
    
    If Col = 9 Then
        With tblRst
            .Row = Row: .Col = 2
            If Trim(.Value) = "" Then
                MsgBox "결과가 입력되지 않았습니다.", vbExclamation
                Exit Sub
            End If
    
            Call ShowGraph(Row)
        
        End With
    End If
    
End Sub

Private Sub ShowGraph(ByVal pRow As Long)

    Dim i As Long, j As Long
    Dim lngSeries As Long, lngPoints As Long
    Dim dblMaxValue As Double, dblMinValue As Double
    Dim dblFromRef As Double, dblToRef As Double
    Dim sPnt As Long, ePnt As Long
    Dim sXVal As Long, eXVal As Long
    Dim tmpStr As String
    Dim strFormat As String

    lngSeries = 1
    lngPoints = 0

    Call ClearGraph

    With cfxResult
        .ClearData CD_VALUES

        .RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
        .OpenDataEx COD_VALUES, lngSeries, objRst.Item(pRow).HistoryCnt + 1

'        .TopGap = 10
'        .BottomGap = 20
'        .LeftGap = 80
'        .FixedGap = 33
        .TopGap = 3
        .BottomGap = 20
        .LeftGap = 50
        .Grid = CHART_NOGRID
        .Scrollable = True
        
        strFormat = "##0" & IIf(objRst.Item(pRow).AvalVal = 9 Or objRst.Item(pRow).AvalVal = 0, "", "." & String(objRst.Item(pRow).AvalVal, "0"))
        
        dblFromRef = Format(objRst.Item(pRow).MeanVal - (objRst.Item(pRow).SdVal * 3), strFormat) 'Val(Format(objRst.Item(pRow).MeanVal, strFormat)) - (Val(Format(objRst.Item(pRow).SdVal, strFormat)) * 3)
        dblToRef = Format(objRst.Item(pRow).MeanVal + (objRst.Item(pRow).SdVal * 3), strFormat) 'Val(Format(objRst.Item(pRow).MeanVal, strFormat)) + (Val(Format(objRst.Item(pRow).SdVal, strFormat)) * 3)
        dblMinValue = dblFromRef
        dblMaxValue = dblToRef

        .Scrollable = True
        .PointLabels = True
        .RGBFont(CHART_POINTFT) = vbBlue
        .Axis(AXIS_X).STEP = 1

        .ThisSerie = 0
        For i = objRst.Item(pRow).HistoryCnt To 1 Step -1

            On Error GoTo Err_Trap

            .Value(lngPoints) = objRst.Item(pRow).RstHistory(i)
            If .Value(lngPoints) > dblMaxValue Then
                .KeyLeg(lngPoints) = CStr(i + 1) & "(+)"
            ElseIf .Value(lngPoints) < dblMinValue Then
                .KeyLeg(lngPoints) = CStr(i + 1) & "(-)"
            Else
                .KeyLeg(lngPoints) = CStr(i + 1)
            End If
            lngPoints = lngPoints + 1
        Next
        
        .Axis(AXIS_Y).Decimals = objRst.Item(pRow).AvalVal
        .Axis(AXIS_Y).Min = dblMinValue ' - objRst.Item(pRow).SdVal
        .Axis(AXIS_Y).Max = dblMaxValue '+ objRst.Item(pRow).SdVal

        .Axis(AXIS_Y).STEP = objRst.Item(pRow).SdVal 'dblFromRef - dblToRef
        
        Dim varRstCd As Variant
        Call tblRst.GetText(2, pRow, varRstCd)
        
        .Value(lngPoints) = Format(Val(objRst.Item(pRow).RstCd), strFormat) 'objRst.Item(pRow).RstCd)
        
        If .Value(lngPoints) > dblMaxValue Then
            .KeyLeg(lngPoints) = CStr(i + 1) & "(+)"
        ElseIf .Value(lngPoints) < dblMinValue Then
            .KeyLeg(lngPoints) = CStr(i + 1) & "(-)"
        Else
            .KeyLeg(lngPoints) = CStr(i + 1)
        End If

        .CloseData COD_VALUES
'------------------------------------------
        .OpenDataEx COD_STRIPES, 6, 0

        Call DrawStripe(0, Format(objRst.Item(pRow).N_3SdVal, strFormat), Format(objRst.Item(pRow).N_2SdVal, strFormat))
        Call DrawStripe(1, Format(objRst.Item(pRow).N_2SdVal, strFormat), Format(objRst.Item(pRow).N_1SdVal, strFormat))
        Call DrawStripe(2, Format(objRst.Item(pRow).N_1SdVal, strFormat), Format(objRst.Item(pRow).MeanVal, strFormat))
        Call DrawStripe(3, Format(objRst.Item(pRow).MeanVal, strFormat), Format(objRst.Item(pRow).P_1SdVal, strFormat))
        Call DrawStripe(4, Format(objRst.Item(pRow).P_1SdVal, strFormat), Format(objRst.Item(pRow).P_2SdVal, strFormat))
        Call DrawStripe(5, Format(objRst.Item(pRow).P_2SdVal, strFormat), Format(objRst.Item(pRow).P_3SdVal, strFormat))

        .CloseData COD_STRIPES
'-------------------------------------------
        .OpenDataEx COD_CONSTANTS, 7, 0

        Call DrawLine(0, Format(objRst.Item(pRow).P_3SdVal, strFormat))
        Call DrawLine(1, Format(objRst.Item(pRow).P_2SdVal, strFormat))
        Call DrawLine(2, Format(objRst.Item(pRow).P_1SdVal, strFormat))
        Call DrawLine(3, Format(objRst.Item(pRow).MeanVal, strFormat), 0)
        Call DrawLine(4, Format(objRst.Item(pRow).N_1SdVal, strFormat))
        Call DrawLine(5, Format(objRst.Item(pRow).N_2SdVal, strFormat))
        Call DrawLine(6, Format(objRst.Item(pRow).N_3SdVal, strFormat))

        .CloseData COD_CONSTANTS
'----------------------------------------------
'        .OpenDataEx COD_VALUES, lngSeries, objRst.Item(pRow).HistoryCnt + 1
'
'        .Axis(AXIS_Y).Min = dblMinValue '- objRst.Item(pRow).SdVal
'        .Axis(AXIS_Y).Max = dblMaxValue '+ objRst.Item(pRow).SdVal
'
'        .Axis(AXIS_Y).STEP = objRst.Item(pRow).SdVal 'dblFromRef - dblToRef
'
'        .CloseData COD_VALUES

    End With
    
'    fraLabel.Visible = True
    Exit Sub

Err_Trap:
    Resume Next
End Sub

Private Sub DrawStripe(ByVal lngCnt As Long, ByVal dblFromVal As Double, ByVal dblToVal As Double)

    With cfxResult.Stripe(lngCnt)
        .Axis = AXIS_Y
        .Color = Choose(lngCnt + 1, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF)
        .From = dblFromVal
        .To = dblToVal
    End With

End Sub

Private Sub DrawLine(ByVal lngCnt As Long, ByVal dblValue As Double, Optional ByVal lngStyle As Long = CHART_DOT)

    With cfxResult.ConstantLine(lngCnt)
        .Value = dblValue
        .LineColor = &H808080  '&H80&
        .Axis = AXIS_Y
        .Label = Choose(lngCnt + 1, "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)")
        .LineWidth = 1
        .LineStyle = lngStyle
    End With

End Sub

Private Sub tblRst_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i       As Integer
    
    '## 보류표시 Clear
    If chkWard.Value = 1 Then Exit Sub
    
    If Row = 0 And Col = 4 Then
        With tblRst
            .Col = 4
            blnExpect = IIf(blnExpect, False, True)
            For i = 1 To .MaxRows
                .Row = i
                If .CellType = CellTypeCheckBox Then
                    .Value = IIf(blnExpect, 0, 1)
                End If
            Next
        End With
    End If
    
End Sub

Private Sub tblRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    If Row > tblRst.DataRowCnt Then Exit Sub
    tblRst.Col = 1: tblRst.Row = Row
    If tblRst.Value = "" Then Exit Sub

    If Col = 2 Then
        objRst.Item(Row).TxtFg = ""
        objRst.Item(Row).RstText = ""
        objRst.Item(Row).RstCd = ""
        objRst.Item(Row).RaDiv = ""
        objRst.Item(Row).RstType = ""
        
        tblRst.Row = Row
        tblRst.Col = 7 + F_OneRow
        tblRst.Value = ""
        tblRst.Col = 5 + F_OneRow
        tblRst.Value = ""
        txtComment.Text = ""
        ChangeFg = True
    End If
    
    If Col = 7 + F_OneRow Then
        objRst.Item(Row).TxtFg = ""
        objRst.Item(Row).RstText = ""
        
        txtComment.Text = ""
        ChangeFg = True
    End If
End Sub

Private Sub tblRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim pValue As Variant
    Dim strReason As String
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub

    If tblRst.DataRowCnt = 0 Then Exit Sub
    If Row = 0 Then Exit Sub
    If Row > tblRst.DataRowCnt Then Exit Sub
    tblRst.Col = 1: tblRst.Row = Row
    If tblRst.Value = "" Then Exit Sub

    Call ClearGraph
    Call ShowResultText(NewRow)        '결과소견 Display

    If Col = 2 Then
        If MsgFg Then Exit Sub
        If Not ChangeFg Then Exit Sub
    
        With objRst
            Call tblRst.GetText(Col, Row, pValue)
            Debug.Print pValue
            If pValue = "Weak reactive" Or pValue = "Weak Reactive" Then
               Call tblRst.SetText(Col, Row, "Weak Reactive")
            Else
               'Call tblRst.SetText(Col, Row, UCase(Mid(pValue, 1, 1)) & LCase(Mid(pValue, 2)))
            End If
            
            If pValue = "" Then Exit Sub
            If CheckResult(Row, pValue) = False Then
                If .Item(Row).RstText = "" Then
                    Cancel = True
                    tblRst.SetFocus
                End If
            End If
    
            Call tblRst.SetText(7 + F_OneRow, Row, .Item(Row).RstText)
            tblRst.EditEnterAction = EditEnterActionDown
        End With
    End If
        
    If Col = 7 + F_OneRow Then
        If MsgFg Then Exit Sub
        If Not ChangeFg Then Exit Sub
        
        tblRst.Col = 7 + F_OneRow
        tblRst.Row = Row
        If tblRst.Value = "" And objRst.Item(Row).RaDiv = "R" Then
            MsgFg = True
            MsgBox "Reject 사유를 반드시 입력해야 합니다.", vbExclamation
            MsgFg = False
            
            Cancel = True
            tblRst.SetFocus
        ElseIf tblRst.Value <> "" Then
            objRst.Item(Row).TxtFg = "1"
            objRst.Item(Row).RstText = tblRst.Value
            txtComment.Text = tblRst.Value
        End If
        
    End If
'    If NewRow > tblRst.DataRowCnt Then
'        tblRst.Row = 1
'        tblRst.Col = Col
'        tblRst.Action = ActionActiveCell
'    End If
End Sub

Private Function CheckResult(ByVal pRow As Long, ByVal pValue As String) As Boolean
    
'수치결과면
 '1, 유효숫자 체크 2, 멀티룰 체크
'문자결과면
 ' 결과코드와의 매칭만 시도
    With objRst.Item(pRow)
        .RstCd = ""
        .RstVal = 0
        .RaDiv = ""
        .RaName = ""
        .RaColor = 0
        .VfyDt = ""
        .VfyTm = ""
        .VfyId = ""
'        .RstType = IIf(IsNumeric(pValue), "N", "F")
        
        CheckResult = True
        
        If IsNumeric(pValue) Then
            If .RefCd <> "" Then
                If CheckString(pRow, pValue) Then
                    .RaDiv = "A"
                    .RaName = "Accept"
                    .RaColor = DCM_LightBlue
                Else
                    .RaDiv = "R"
                    .RaName = "Reject"
                    .RaColor = DCM_LightRed
                    CheckResult = False
                End If
                
                .RstType = "F"
            Else
                If CheckNumeric(pRow, pValue) Then  '정상
                    If MultiRule(pRow, pValue) Then '정상
                        .RaDiv = "A"
                        .RaName = "Accept"
                        .RaColor = DCM_LightBlue
                    Else
                        .RaDiv = "R"
                        .RaName = "Reject"
                        .RaColor = DCM_LightRed
                        CheckResult = False
                    End If
                Else    'AR 판단을 할수 없다.. 유효숫자 입력 오류기때문에...
                    CheckResult = False
                    Exit Function
                End If
                
                If .RaDiv = "A" Then
                    If Val(pValue) > Val(objRst.Item(pRow).P_2SdVal) Or Val(pValue) < Val(objRst.Item(pRow).N_2SdVal) Then
                        .RaDiv = "A"
                        .RaName = "Warning"
                        .RaColor = DCM_LightRed
'                        Stop
                        'strTmp = "Step1 - Once 3SD"
                    End If
                End If
                
                .RstType = "N"
            End If
            
            objRst.Item(pRow).RstCd = pValue
            objRst.Item(pRow).RstVal = IIf(objRst.Item(pRow).RefCd = "", pValue, 0)
        Else
            If objRst.Item(pRow).RefCd <> "" Then
                If CheckString(pRow, pValue) Then
                    .RaDiv = "A"
                    .RaName = "Accept"
                    .RaColor = DCM_LightBlue
                Else
                    .RaDiv = "R"
                    .RaName = "Reject"
                    .RaColor = DCM_LightRed
                    CheckResult = False
                End If
                
                .RstType = "F"
            Else
                CheckResult = False
                MsgFg = True
                MsgBox "수치결과로만 입력하십시오.", vbExclamation
                MsgFg = False
                
                Exit Function
            End If
            
            objRst.Item(pRow).RstCd = pValue
            objRst.Item(pRow).RstVal = 0
        End If
    
        tblRst.Row = pRow
        tblRst.Col = 5 + F_OneRow
        tblRst.Value = objRst.Item(pRow).RaName
        tblRst.ForeColor = objRst.Item(pRow).RaColor
    End With
End Function

Private Function CheckNumeric(ByVal pRow As Long, ByVal pValue As String) As Boolean
'수치결과인 경우
'먼저 유효숫자 체크를 하고 멀티룰 체크를 수행한다.
'유효숫자가 잘못입력 된경우 오류 메시지만 띄워주고 True리턴
    Dim strTmp As String
    
    CheckNumeric = False
    
    strTmp = medGetP(pValue, 2, ".")
    
    With objRst.Item(pRow)
        '유효숫자 체크
        If .AvalVal = 9 Or .AvalVal = 0 Then
            If InStr(pValue, ".") > 0 Then
                MsgFg = True
                MsgBox "유효숫자 입력오류 입니다. 정수형만 입력하십시오.", vbCritical
                MsgFg = False
                Exit Function
            End If
        Else
            If Len(strTmp) <> .AvalVal Then
                MsgFg = True
                MsgBox "유효숫자 입력오류입니다. 소숫점이하 " & .AvalVal & "자리까지 입력하십시오.", vbCritical
                MsgFg = False
                Exit Function
            End If
        End If
    End With
    
    CheckNumeric = True
End Function

Private Function CheckString(ByVal pRow As Long, ByVal pValue As String) As Boolean
'문자 결과인 경우
'알파코드 참고치하고 비고한다.
'정상이면 True를 리턴한다.
    Dim strReason As String
    
    CheckString = True
    
    With objRst.Item(pRow)
        If .RefCd <> pValue Then
            CheckString = False
            
'            If objRst.Item(pRow).RstText <> "" Then
'                If InStr(objRst.Item(pRow).RstText, "참고치와 결과치가 서로 다름") > 0 Then Exit Function
'            End If
            
            If objRst.Item(pRow).RstText <> "" Then Exit Function
            
            MsgFg = True
            MsgBox "참고치와 결과치가 서로 다릅니다.", vbExclamation
            MsgFg = False
            
            strReason = objRst.Item(pRow).RstText
            If strReason = "" Then
                While (strReason = "")
                    MsgFg = True
                    strReason = InputBox("Reject 소견을 입력하십시오.", "소견 입력")
                    MsgFg = False
                Wend
                objRst.Item(pRow).TxtFg = "1"
                
                '-- 원본 ======================================================================
                'objRst.Item(pRow).RstText = "참고치와 결과치가 서로 다름" & vbCrLf & strReason & vbCrLf
                '==============================================================================
                
                '-- 수정 By M.G.Choi 2005.11.09
                objRst.Item(pRow).RstText = "" & vbCrLf & strReason & vbCrLf
                
            End If
        End If
    End With
End Function

Private Function MultiRule(ByVal pRow As Long, ByVal pValue As String) As Boolean
    Dim pRstVal As Double
    Dim tmpVal As Double
    Dim lngCnt As Long
    Dim i As Long
    
    Dim strTmp As String
    
    MultiRule = True
    pRstVal = Val(pValue)
    
    'SD값이 설정되어 있지 않는 경우
    If objRst.Item(pRow).SdVal = 0 Then Exit Function
    
'Step1 : Once 3SD
    If Mid$(objRst.Item(pRow).WsSet, 1, 1) = "1" Then
        If pRstVal > objRst.Item(pRow).P_3SdVal Or pRstVal < objRst.Item(pRow).N_3SdVal Then
            MultiRule = False
            strTmp = "Step1 - Once 3SD"
            GoTo Reason
        End If
    End If

'Step2 : Once 4SD
    If Mid$(objRst.Item(pRow).WsSet, 2, 1) = "1" Then
        If pRstVal > objRst.Item(pRow).P_2SdVal And objRst.Item(pRow).RstHistory(1) < objRst.Item(pRow).N_2SdVal Then
            MultiRule = False
            strTmp = "Step2 - Once 4SD"
            GoTo Reason
        ElseIf objRst.Item(pRow).RstHistory(1) > objRst.Item(pRow).P_2SdVal And pRstVal < objRst.Item(pRow).N_2SdVal Then
            MultiRule = False
            strTmp = "Step2 - Once 4SD"
            GoTo Reason
        End If
    End If
    
'Step3 : Twice 2SD
    If Mid$(objRst.Item(pRow).WsSet, 3, 1) = "1" Then
        If pRstVal > objRst.Item(pRow).P_2SdVal And objRst.Item(pRow).RstHistory(1) > objRst.Item(pRow).P_2SdVal Then
            MultiRule = False
            strTmp = "Step3 - Twice 2SD"
            GoTo Reason
        ElseIf pRstVal < objRst.Item(pRow).N_2SdVal And objRst.Item(pRow).RstHistory(1) < objRst.Item(pRow).N_2SdVal Then
            MultiRule = False
            strTmp = "Step3 - Twice 2SD"
            GoTo Reason
        End If
    End If

'Step4 : 4 Times 1SD
    If Mid$(objRst.Item(pRow).WsSet, 4, 1) = "1" Then
        lngCnt = 0
        If pRstVal > objRst.Item(pRow).P_1SdVal And objRst.Item(pRow).HistoryCnt >= 3 Then    '(+)1SD
            lngCnt = 1
            For i = 1 To 3
                If objRst.Item(pRow).RstHistory(i) <= objRst.Item(pRow).P_1SdVal Then Exit For
                lngCnt = lngCnt + 1
            Next
        ElseIf pRstVal < objRst.Item(pRow).N_1SdVal And objRst.Item(pRow).HistoryCnt >= 3 Then  '(-)1SD
            lngCnt = 1
            For i = 1 To 3
                If objRst.Item(pRow).RstHistory(i) >= objRst.Item(pRow).N_1SdVal Then Exit For
                lngCnt = lngCnt + 1
            Next
        End If
    
        'CS_TotCnt개의 결과중 CS_ChkCnt번 (+,-)1SD를 연속해서 벗어남.
        If lngCnt = 4 Then
            MultiRule = False
            strTmp = "Step4 - 4 Times 1SD"
            GoTo Reason
        End If
    End If

'Step5 : 10 Times Trend
    If Mid$(objRst.Item(pRow).WsSet, 5, 1) = "1" Then
        lngCnt = 0
        tmpVal = pRstVal
        If tmpVal > objRst.Item(pRow).MeanVal And objRst.Item(pRow).HistoryCnt >= 9 Then
            lngCnt = 1
            For i = 1 To 9
                If objRst.Item(pRow).RstHistory(i) <= objRst.Item(pRow).MeanVal Then Exit For
                lngCnt = lngCnt + 1
            Next i
        ElseIf tmpVal < objRst.Item(pRow).MeanVal And objRst.Item(pRow).HistoryCnt >= 9 Then
            lngCnt = 1
            For i = 1 To 9
                If objRst.Item(pRow).RstHistory(i) >= objRst.Item(pRow).MeanVal Then Exit For
                lngCnt = lngCnt + 1
            Next i
        End If
        
        If lngCnt = 10 Then
            MultiRule = False
            strTmp = "Step5 - 10 Times Trend"
            GoTo Reason
        End If
    End If
    
    MultiRule = True
    Exit Function

Reason:
    Dim strReason As String
    
    If strTmp <> "" Then
        MultiRule = False

        If objRst.Item(pRow).RstText <> "" Then Exit Function
        
        MsgFg = True
        MsgBox strTmp, vbExclamation
        MsgFg = False
        
        strReason = objRst.Item(pRow).RstText
        If strReason = "" Then
            While (strReason = "")
                MsgFg = True
                strReason = InputBox("Reject 소견을 입력하십시오.", "소견 입력")
                MsgFg = False
            Wend
            objRst.Item(pRow).TxtFg = "1"
            objRst.Item(pRow).RstText = strTmp & vbCrLf & strReason & vbCrLf
        End If
    End If
End Function

Private Sub tblRst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim objPop As clsPopUpList
    Dim Rs As Recordset
    Dim strSql As String
    Dim RLAB001     As New ADODB.Recordset
    
    If Col <> 2 Then Exit Sub
    If Row > tblRst.DataRowCnt Then Exit Sub
    tblRst.Col = 1: tblRst.Row = Row
    If tblRst.Value = "" Then Exit Sub
    
    With tblRst
        strSql = " select cdval2,field1 from " & T_LAB031 & _
                 " where " & DBW("cdindex=", LC2_ItemResult) & _
                 " and " & DBW("cdval1=", objRst.Item(Row).TestCd)
        Debug.Print strSql
        Set Rs = New Recordset
        Rs.Open strSql, DBConn
        
        If Rs.EOF = False Then
            Set objPop = New clsPopUpList
            objPop.Recordset = Rs
            objPop.HideSearchTool = True
            objPop.SelectByClick = True
            objPop.FormWidth = 4635
            objPop.FormHeight = 2880
            objPop.FormCaption = "결과코드 찾기"
            objPop.ColumnHeaderText = "결과코드;결과코드명"
            objPop.ColumnHeaderWidth = "1110.047;3075.024"
            objPop.HideColumnHeaders = True
            objPop.LoadPopUp
            
            .Col = 2
            .Row = Row
            
            '2014-03-27 PSK WORKAREA = '03' BLOOD BANK 혈액은행이 아닐경우 결과코드값이 아닌 결과값을 가져오자
            RLAB001.Open "SELECT * FROM S2LAB001 WHERE TESTCD = '" & objRst.Item(Row).TestCd & "' AND WORKAREA = '03'", DBConn, adOpenForwardOnly, adLockReadOnly
            If Not RLAB001.EOF() And Not RLAB001.BOF() Then
                .Value = medGetP(objPop.SelectedString, 2, objPop.Delimiter)
            Else
                .Value = medGetP(objPop.SelectedString, 1, objPop.Delimiter)
            End If
            RLAB001.Close: Set RLAB001 = Nothing

            '2014-03-28 PSK 모두 결과값으로 처리한다함.
            '.Value = medGetP(objPop.SelectedString, 2, objPop.Delimiter)
            
            Call tblRst_EditChange(Col, Row)
            
            Call tblRst_LeaveCell(Col, Row, Col + 1, Row, False)
        End If
            
    End With
    
    Set Rs = Nothing
    Set objPop = Nothing
End Sub

Private Sub tblRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strAccNo As String
    Dim strVfyDt As String
    Dim strVfyTm As String
    
    If Row = 0 Or tblRst.DataRowCnt = 0 Then ShowTip = False: Exit Sub
    If Col <> 3 Then ShowTip = False: Exit Sub
    
    tblRst.Row = Row
    tblRst.Col = 3
    If tblRst.Value = "" Then
        ShowTip = False
        Exit Sub
    End If
    
    tblRst.Row = Row
    tblRst.Col = 9 + F_OneRow
    strAccNo = tblRst.Value
    tblRst.Col = 10 + F_OneRow
    strVfyDt = Format(tblRst.Value, "0###-0#-0#")
    tblRst.Col = 11 + F_OneRow
    strVfyTm = Format(tblRst.Value, "0#:0#:0#")
    
    MultiLine = 1
    TipText = vbNewLine & " 최근 결과접수번호 : " & strAccNo & vbNewLine
    TipText = TipText & " 최근 결과일시 : " & strVfyDt & " " & _
                                               strVfyTm & vbNewLine
                                               
    TipWidth = 4000
    Call tblRst.SetTextTipAppearance("굴림체", 9, True, False, &HEEFDF2, &H996666)
    ShowTip = True
End Sub



Private Sub txtAccNo_Change()
    Dim lngLen As Long
    Static lngAccDt As Long
    
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtAccNo.Name Then Exit Sub
    
    If chkBarcode.Value = 0 Then    '접수번호로 입력할때만 유효
        With txtAccNo
            lngLen = Len(Trim(.Text))
            
            If lngLen < 2 Then
                lngAccDt = 0
            End If
            
            If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
                                
                lngAccDt = GetLenOfAccDt
            End If
            
            If lngLen > 2 And lngLen = lngAccDt + 3 And lngAccDt <> 0 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    End If
    
    If lblControl.Caption <> "" Then
        lblBarNo.Caption = ""
        Call InitForm
        Call ClearGraph
    End If
End Sub

Private Function GetLenOfAccDt() As Long
    Dim objSQL As clsQcOrder
    Dim Rs As Recordset
    Dim strSql As String
    
    Set objSQL = New clsQcOrder
    strSql = objSQL.SqlCommonCode(T_LAB032, lc3_workarea, Mid(txtAccNo.Text, 1, 2))
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF Then
        GetLenOfAccDt = 6
    Else
        Select Case Rs.Fields("field2").Value & ""
            Case enLabDiv.LabDiv_ByDay:       '일단위
                GetLenOfAccDt = 6
            Case enLabDiv.LabDiv_ByMonth:       '월단위
                GetLenOfAccDt = 4
            Case enLabDiv.LabDiv_ByYear:       '년단위
                GetLenOfAccDt = 2
            Case enLabDiv.LabDiv_BySpc:       '검체단위
                GetLenOfAccDt = 4
            Case Else:
                GetLenOfAccDt = 6
        End Select
    End If
    
    Set Rs = Nothing
    Set objSQL = Nothing
End Function

Private Sub txtAccNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtAccNo.Text) = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    If chkBarcode.Value = 0 Then '접수번호를 입력할때만 유효
        If KeyAscii = vbKeyBack Then
            With txtAccNo
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    If Len(.Text) > 2 Then
                        .Text = Mid(.Text, 1, Len(.Text) - 2)
                        .SelStart = Len(.Text)
                        KeyAscii = 0
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub txtAccNo_LostFocus()
    If Trim(txtAccNo.Text) = "" Then Exit Sub
    If lblControl.Caption <> "" Then Exit Sub
    
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    
    Call LoadData
End Sub

Private Sub LoadData()
    Dim strAccNo As String
    
    If Trim(txtAccNo.Text) = "" Then Exit Sub
    
    If chkBarcode.Value = 1 Then
        strAccNo = GetAccNo
        lblBarNo.Caption = strAccNo
    Else
        strAccNo = Trim(txtAccNo.Text)
        lblBarNo.Caption = GetSpcNo
    End If

    If strAccNo = "" Then
        MsgBox "유효한 바코드 정보가 아닙니다.", vbExclamation
        txtAccNo.SetFocus
        Exit Sub
    End If
    
    If lblBarNo.Caption = "" Then
        MsgBox "유효한 접수번호가 아닙니다.", vbExclamation
        txtAccNo.SetFocus
        Exit Sub
    End If

    Set objRst = Nothing
    Set objRst = New clsQcResult
    
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim strAccSeq As String
    
    strWorkArea = medGetP(strAccNo, 1, "-")
    strAccDt = medGetP(strAccNo, 2, "-")
    strAccDt = IIf(strAccDt Like "99*", "19" & strAccDt, "20" & strAccDt)
    strAccSeq = medGetP(strAccNo, 3, "-")
    
    '접수내역 불러오기
    DoEvents
    If objRst.getSlipTable(strWorkArea, strAccDt, strAccSeq) = False Then
        MsgBox "접수내역이 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    '결과내역 불러오기
    DoEvents
    If objRst.getRstTable(strWorkArea, strAccDt, strAccSeq, prgProgress) = False Then
        MsgBox "결과내역이 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    If Val(objRst.l201_stscd) >= enStsCd.StsCd_LIS_FinRst Then '최종결과난 경우
    
    ElseIf objRst.l201_qcfg < "1" Then  '일반검사 접수번호를 입력한 경우
        MsgBox "QC 처방이 아닙니다. 일반결과 등록 메뉴를 사용하십시오.", vbExclamation
        Exit Sub
    ElseIf objRst.l201_qcfg = "2" Then    '외부정도관리
        MsgBox "외부정도관리 접수번호입니다. 외부정도관리 화면에서 등록하십시오.", vbExclamation
        Exit Sub
    End If

    With objRst
        If .TestCount > 0 Then
            lblControl.Caption = Format(.l201_ctrlcd, "!" & String(10, "@")) & .l201_ctrlnm
            lblLevel.Caption = IIf(.l201_levelcd = "L", "Low", IIf(.l201_levelcd = "N", "Normal", "High"))
            lblLotNo.Caption = .l201_lotno
            lblEqp.Caption = Format(.l201_eqpcd, "!" & String(10, "@")) & .l201_eqpnm
            lblMakeCd.Caption = .l201_makecd
            lblRemark.Caption = .l201_remark
            
'스푸레드에 담아주는..
            Call LoadResultData '.SetSpread(tblRst, 1)
            Call ShowResultText(1)
            ChangeFg = False
        Else
            MsgBox "결과등록할 검사항목이 없습니다.", vbExclamation
            Exit Sub
        End If
    End With
End Sub

Private Function GetAccNo() As String
    Dim Rs As Recordset
    Dim strSql As String
    Dim strSpcYY As String
    Dim strSpcNo As Long
    
    strSpcYY = Mid(Trim(txtAccNo.Text), 1, P_SpcYyLength)
    strSpcNo = Format(Mid(Trim(txtAccNo.Text), P_SpcYyLength + 1, P_SpcNoLength), "#0")
        
    strSql = " select workarea,accdt,accseq from " & T_LAB201 & _
             " where " & DBW("spcyy=", strSpcYY) & _
             " and " & DBW("spcno=", strSpcNo)
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF Then
        GetAccNo = ""
    Else
        GetAccNo = Rs.Fields("workarea").Value & "" & "-" & _
                   Mid(Rs.Fields("accdt").Value, 3) & "" & "-" & _
                   Rs.Fields("accseq").Value & ""
    End If
    
    Set Rs = Nothing
End Function

'----------------------------------------------------------------
' 2009.02.18 양성현 수정
Private Function GetSpcNo() As String
    With tblWorkList
        .Row = currWorkList
        .Col = 5
        GetSpcNo = Mid(.Value, 1, 2) & "" & Format(medGetP(.Value, 2, "-"), LIS_BarFormat)
    
    End With
End Function

'Private Function GetSpcNo() As String
'    Dim Rs As Recordset
'    Dim strSQL As String
'    Dim strWorkArea As String
'    Dim strAccDt As String
'    Dim strAccSeq As String
'
'    strWorkArea = medGetP(Trim(txtAccNo.Text), 1, "-")
'    strAccDt = IIf(medGetP(Trim(txtAccNo.Text), 2, "-") Like "99*", _
'                  "19" & medGetP(Trim(txtAccNo.Text), 2, "-"), _
'                  "20" & medGetP(Trim(txtAccNo.Text), 2, "-"))
'    strAccSeq = medGetP(Trim(txtAccNo.Text), 3, "-")
'
'    strSQL = " select spcyy,spcno from " & T_LAB201 & _
'             " where " & DBW("workarea=", strWorkArea) & _
'             " and " & DBW("accdt=", strAccDt) & _
'             " and " & DBW("accseq=", strAccSeq)
'
'    Set Rs = New Recordset
'    Rs.Open strSQL, DBConn
'
'    If Rs.EOF Then
'        GetSpcNo = ""
'    Else
'        GetSpcNo = Rs.Fields("spcyy").Value & "" & Format(Rs.Fields("spcno").Value & "", LIS_BarFormat)
'    End If
'
'    Set Rs = Nothing
'End Function
'----------------------------------------------------------------
Private Sub LoadResultData()
    Dim i As Long
    
    Call ClearTable
        
    tblRst.ReDraw = False
    
    For i = 1 To objRst.TestCount
        With objRst.Item(i)
            If tblRst.DataRowCnt >= tblRst.MaxRows Then
                tblRst.MaxRows = tblRst.MaxRows + 1
                tblRst.RowHeight(-1) = 12
            End If
            
            tblRst.Row = tblRst.DataRowCnt + 1
            
            tblRst.Col = 1: tblRst.Value = .TestNm
            tblRst.Col = 2: tblRst.Value = Trim(.RstCd)
            Debug.Print .RstCd
            
            '** 보류체크 추가 By M.G.Choi 2006.09.01
            If .DetailFg <> "" Then
                tblRst.Col = 4: tblRst.Value = ""
                tblRst.CellType = CellTypeStaticText ' 5
            Else
                If .VfyDt = "" Then
                    tblRst.Col = 4: tblRst.Value = "1"
                End If
            End If
            
            tblRst.Col = 4 + F_OneRow: tblRst.Value = .RstUnit
            tblRst.Col = 5 + F_OneRow: tblRst.Value = IIf(.RstCd = "", "", IIf(.RaDiv = "A", "Accept", "Reject"))
            tblRst.Col = 7 + F_OneRow: tblRst.Value = .RstText
            
            If .RaDiv = "A" Then
                If Val(Trim(.RstCd)) > Val(objRst.Item(tblRst.Row).P_2SdVal) Or Val(Trim(.RstCd)) < Val(objRst.Item(tblRst.Row).N_2SdVal) Then
                    Debug.Print .RstCd
                    Debug.Print Val(objRst.Item(tblRst.Row).P_2SdVal)
                    Debug.Print Val(objRst.Item(tblRst.Row).N_2SdVal)
                                        
                    '2014-03-03 "0" 값을 비교하는경우 발생 SKIP 처리함
                    If Val(objRst.Item(tblRst.Row).P_2SdVal) <> 0 And Val(objRst.Item(tblRst.Row).P_2SdVal) <> 0 Then
                        tblRst.Col = 5 + F_OneRow:
                        tblRst.ForeColor = vbRed
                        tblRst.Value = "Warning"
                    Else
                        tblRst.Col = 5 + F_OneRow:
                        tblRst.ForeColor = vbBlue
                        tblRst.Value = IIf(.RstCd = "", "", IIf(.RaDiv = "A", "Accept", "Reject"))
                    End If
                Else
                    tblRst.Col = 5 + F_OneRow:
                    tblRst.ForeColor = vbBlue
                    tblRst.Value = IIf(.RstCd = "", "", IIf(.RaDiv = "A", "Accept", "Reject"))
                End If
            End If
            
            
            Call GetRecentRst(tblRst.Row)    'col=3
            Call GetRefValue(tblRst.Row) 'col=6
        End With
    Next i
    
'    With tblRst
'        .Col = 1: .Col2 = .MaxCols
'        .Row = .DataRowCnt + 1: .Row2 = .MaxRows
'        .BlockMode = True
'        .CellType = CellTypeStaticText
'        .BlockMode = False
'    End With
    
    Call SetForeColor
    
    tblRst.ReDraw = True
    
    On Error Resume Next
    tblRst.Row = 1: tblRst.Col = 2
    tblRst.SetFocus
    tblRst.Action = ActionActiveCell
End Sub

Private Function GetRecentRst(ByVal pRow As Long)
    Dim Rs As Recordset
    Dim strSql As String
    Dim strFormat As String
    
    strFormat = "##0" & IIf(objRst.Item(pRow).AvalVal = 9 Or objRst.Item(pRow).AvalVal = 0, "", "." & String(objRst.Item(pRow).AvalVal, "0"))
    
    strSql = " select a.testcd, a.rstcd, a.vfydt, a.vfytm,a.radiv,a.workarea,a.accdt,a.accseq" & _
            " from " & T_LAB026 & " a, " & T_LAB026 & " b " & _
            " where " & DBW("b.workarea=", objRst.Item(pRow).WorkArea) & _
            " and " & DBW("b.accdt=", objRst.Item(pRow).AccDt) & _
            " and " & DBW("b.accseq=", objRst.Item(pRow).AccSeq) & _
            " and " & DBW("b.testcd=", objRst.Item(pRow).TestCd) & _
            " and a.ctrlcd = b.ctrlcd " & _
            " and a.levelcd = b.levelcd " & _
            " and a.lotno = b.lotno " & _
            " and a.testcd = b.testcd " & _
            " and (a.vfydt<>' ' or a.vfydt  is not null) " & _
            " and (a.workarea" & FUNC_CONCAT & "a.accdt" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "a.accseq") & ") <> (b.workarea" & FUNC_CONCAT & "b.accdt" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "b.accseq") & ") " & _
            " order by a.testcd, a.vfydt desc, a.vfytm desc "
    
    Debug.Print strSql
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF = False Then
        With tblRst
            .Row = pRow
            .Col = 3
            
            Select Case Rs.Fields("testcd").Value
             Case "B2047", "B2000A", "B2000B", "B2021D", "B2061", "B2023", "B2000AC", "B2000BC"
                  .Value = Trim(Rs.Fields("rstcd").Value & "")
             Case Else
                  .Value = Format(Rs.Fields("rstcd").Value & "", strFormat)
            End Select
            If Rs.Fields("radiv").Value & "" = "A" Then
                .ForeColor = DCM_LightBlue
            ElseIf Rs.Fields("radiv").Value & "" = "R" Then
                .ForeColor = DCM_LightRed
            End If
            
            .Col = 9 + F_OneRow
            .Value = Rs.Fields("workarea").Value & "" & "-" & Mid(Rs.Fields("accdt").Value & "", 3) & "-" & Rs.Fields("accseq").Value & ""
            .Col = 10 + F_OneRow
            .Value = Rs.Fields("vfydt").Value & ""
            .Col = 11 + F_OneRow
            .Value = Rs.Fields("vfytm").Value & ""
        End With
    End If
    
    Set Rs = Nothing
'    If objRst.Item(pRow).HistoryCnt <> 0 Then
'        With tblRst
'            .Col = 3
'            .Row = pRow
'            .Value = Format(objRst.Item(pRow).RstHistory(1), strFormat)
'
'            .ForeColor = IIf(objRst.Item(pRow).RstHistoryRaDiv(1) = "R", DCM_LightRed, DCM_LightBlue)
'        End With
'    End If
    
    
    
End Function

Private Sub GetRefValue(ByVal pRow As Long)
    Dim strTmp As String
    
    With objRst.Item(pRow)
        If .RefCd = "" Then
            
            strTmp = Format(.MeanVal, "##0" & IIf(.AvalVal = 9 Or .AvalVal = 0, "", "." & String(.AvalVal, "0")))
            strTmp = strTmp & "(" & Format(.MinVal, "##0" & IIf(.AvalVal = 9 Or .AvalVal = 0, "", "." & String(.AvalVal, "0")))
            strTmp = strTmp & "~" & Format(.MaxVal, "##0" & IIf(.AvalVal = 9 Or .AvalVal = 0, "", "." & String(.AvalVal, "0"))) & ")"
            
            tblRst.Row = pRow
            tblRst.Col = 6 + F_OneRow
            tblRst.Value = strTmp
            
            tblRst.Row = pRow: tblRst.Row2 = pRow
            tblRst.Col = 8 + F_OneRow: tblRst.Col2 = 8 + F_OneRow
            tblRst.BlockMode = True
            tblRst.CellType = CellTypeButton
            tblRst.BlockMode = False
        ElseIf .RefCd <> "" Then
            strTmp = .RefCd
            
            tblRst.Row = pRow
            tblRst.Col = 6 + F_OneRow
            tblRst.Value = strTmp
            
            tblRst.Row = pRow: tblRst.Row2 = pRow
            tblRst.Col = 8 + F_OneRow: tblRst.Col2 = 8 + F_OneRow
            tblRst.BlockMode = True
            tblRst.CellType = CellTypeStaticText
            tblRst.BlockMode = False
        End If
    End With
End Sub

Private Function GetMean(ByVal strTestcd As String) As String
    Dim Rs As Recordset
    Dim sSql As String

    sSql = " select meanval from " & T_LAB024 & _
           " where " & DBW("ctrlcd=", objRst.l201_ctrlcd) & _
           " and " & DBW("levelcd=", objRst.l201_levelcd) & _
           " and " & DBW("lotno=", objRst.l201_lotno) & _
           " and " & DBW("testcd=", strTestcd)
         
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetMean = Rs.Fields("meanval").Value & ""
    End If
    
    Set Rs = Nothing

End Function

Private Sub SetForeColor()
    Dim i As Long
    
    With tblRst
        .Row = -1
        .Col = 1
        .ForeColor = DCM_MidBlue
        
        .Col = 4 + F_OneRow
        .ForeColor = DCM_Brown
        
        .Col = 6 + F_OneRow
        .ForeColor = DCM_Green
        
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 5 + F_OneRow
            If objRst.Item(i).RaDiv = "R" Then
                .ForeColor = DCM_LightRed
            Else
                .ForeColor = DCM_LightBlue
            End If
        Next
    End With
End Sub

Private Sub ShowResultText(ByVal pRow As Long)

    Dim varTmp As Variant

    If pRow <= 0 Then Exit Sub

    If objRst.TestCount <= 0 Then Exit Sub
    If pRow > objRst.TestCount Then Exit Sub
    
    If objRst.Item(pRow).TxtFg <> "0" Then
        Call tblRst.GetText(1, pRow, varTmp)
        LisLabel8.Caption = "◈ Reject 소견 - " & varTmp
        txtComment.Text = objRst.Item(pRow).RstText
        txtComment.BackColor = vbWhite
        txtComment.Enabled = True
        cmdPopComment.Enabled = True
    End If

End Sub

Private Sub txtComment_LostFocus()
    Dim lngRow As Long
    
    lngRow = tblRst.ActiveRow
    
    If tblRst.DataRowCnt = 0 Then Exit Sub
    If tblRst.ActiveRow > tblRst.DataRowCnt Then Exit Sub
    If lngRow > tblRst.DataRowCnt Then Exit Sub
    
    objRst.Item(lngRow).TxtFg = "1"
    objRst.Item(lngRow).RstText = txtComment.Text
    Call tblRst.SetText(7 + F_OneRow, lngRow, objRst.Item(lngRow).RstText)
    
    On Error Resume Next
    tblRst.SetFocus
End Sub
