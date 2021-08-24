VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCResSch 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Elecsys QC 조회"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   14970
   StartUpPosition =   2  '화면 가운데
   Begin MSComCtl2.MonthView monvCal 
      Height          =   2220
      Left            =   8070
      TabIndex        =   0
      Top             =   930
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   95485953
      CurrentDate     =   36878
   End
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   14895
      Begin VB.ComboBox cboPart 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1140
         TabIndex        =   38
         Top             =   1050
         Width           =   3045
      End
      Begin VB.CheckBox chkTwin 
         Height          =   285
         Index           =   3
         Left            =   7740
         TabIndex        =   37
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkTwin 
         Height          =   285
         Index           =   2
         Left            =   7740
         TabIndex        =   36
         Top             =   795
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkTwin 
         Height          =   285
         Index           =   1
         Left            =   7740
         TabIndex        =   35
         Top             =   525
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkTwin 
         Height          =   285
         Index           =   0
         Left            =   7740
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdSummary 
         Caption         =   "Summary"
         Height          =   615
         Left            =   12570
         TabIndex        =   32
         Top             =   750
         Width           =   1095
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1395
         Left            =   12870
         TabIndex        =   31
         Top             =   180
         Visible         =   0   'False
         Width           =   1545
         _Version        =   393216
         _ExtentX        =   2725
         _ExtentY        =   2461
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   4
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   0
         ShadowText      =   32768
         SpreadDesigner  =   "QCResSch.frx":0000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "출  력"
         Height          =   615
         Left            =   11460
         TabIndex        =   29
         Top             =   750
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9930
         TabIndex        =   23
         Top             =   660
         Width           =   285
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9930
         TabIndex        =   22
         Top             =   1110
         Width           =   285
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종 료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13680
         TabIndex        =   21
         Top             =   750
         Width           =   1095
      End
      Begin VB.TextBox txtLotNo 
         Appearance      =   0  '평면
         BackColor       =   &H00DEFDFE&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   6150
         TabIndex        =   20
         Top             =   270
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtLotNo 
         Appearance      =   0  '평면
         BackColor       =   &H00EDFEEE&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   1
         Left            =   6150
         TabIndex        =   19
         Top             =   540
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtLotNo 
         Appearance      =   0  '평면
         BackColor       =   &H00E7E8FE&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   2
         Left            =   6150
         TabIndex        =   18
         Top             =   810
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtLotNo 
         Appearance      =   0  '평면
         BackColor       =   &H00FDEADB&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   3
         Left            =   6150
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조 회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10350
         TabIndex        =   16
         Top             =   750
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "전체출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12360
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox lstLevel 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   1
         Left            =   4320
         Style           =   1  '확인란
         TabIndex        =   14
         Top             =   270
         Width           =   1845
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "기  간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8340
         TabIndex        =   12
         Top             =   270
         Width           =   1575
      End
      Begin VB.CommandButton cmdPreLot 
         BackColor       =   &H00C0C0C0&
         Caption         =   "이전 Lot No."
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
         Left            =   10350
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   180
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95485953
         CurrentDate     =   37207
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   24
         Top             =   240
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   7276
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "정도관리 조회"
         ForeColor       =   4210752
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin MSMask.MaskEdBox mskQCDate 
         Height          =   315
         Index           =   0
         Left            =   8550
         TabIndex        =   25
         Top             =   630
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskQCDate 
         Height          =   315
         Index           =   1
         Left            =   8550
         TabIndex        =   26
         Top             =   1065
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   39
         Top             =   1080
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Height          =   405
         Left            =   90
         Top             =   180
         Visible         =   0   'False
         Width           =   2670
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3915
      Left            =   0
      TabIndex        =   1
      Top             =   1620
      Width           =   14925
      _Version        =   65536
      _ExtentX        =   26326
      _ExtentY        =   6906
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread vasSummary 
         Height          =   2565
         Left            =   7290
         TabIndex        =   33
         Top             =   540
         Visible         =   0   'False
         Width           =   6705
         _Version        =   393216
         _ExtentX        =   11827
         _ExtentY        =   4524
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
         MaxCols         =   13
         MaxRows         =   50
         SpreadDesigner  =   "QCResSch.frx":04E9
      End
      Begin FPSpread.vaSpread vasDisplay 
         Height          =   1935
         Left            =   2400
         TabIndex        =   30
         Top             =   1740
         Visible         =   0   'False
         Width           =   4245
         _Version        =   393216
         _ExtentX        =   7488
         _ExtentY        =   3413
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
         SpreadDesigner  =   "QCResSch.frx":125F
      End
      Begin FPSpread.vaSpread vasResult 
         Height          =   1575
         Left            =   6810
         TabIndex        =   28
         Top             =   2100
         Visible         =   0   'False
         Width           =   7635
         _Version        =   393216
         _ExtentX        =   13467
         _ExtentY        =   2778
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   100
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "QCResSch.frx":587D
      End
      Begin FPSpread.vaSpread vasLevel 
         Height          =   1875
         Left            =   6180
         TabIndex        =   27
         Top             =   1140
         Visible         =   0   'False
         Width           =   7635
         _Version        =   393216
         _ExtentX        =   13467
         _ExtentY        =   3307
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "QCResSch.frx":6A50
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   5010
         Visible         =   0   'False
         Width           =   1575
      End
      Begin FPSpread.vaSpread vasNuRes 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   6588
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         OperationMode   =   2
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "QCResSch.frx":AAE3
      End
      Begin FPSpread.vaSpread vasGenList 
         Height          =   2865
         Left            =   720
         TabIndex        =   4
         Top             =   450
         Width           =   9345
         _Version        =   393216
         _ExtentX        =   16484
         _ExtentY        =   5054
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "QCResSch.frx":EC63
      End
      Begin Threed.SSPanel sspTitle 
         Height          =   675
         Left            =   540
         TabIndex        =   5
         Top             =   450
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   1191
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.26
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         Alignment       =   1
         Begin VB.Label lblTitle 
            BackStyle       =   0  '투명
            Caption         =   "내부정도관리"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   300
            TabIndex        =   6
            Top             =   90
            Width           =   1575
         End
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   4185
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Width           =   14925
      _Version        =   65536
      _ExtentX        =   26326
      _ExtentY        =   7382
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   14250
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ChartfxLibCtl.ChartFX ChartFX1 
         Height          =   4095
         Left            =   120
         TabIndex        =   8
         Top             =   90
         Width           =   9945
         _cx             =   17542
         _cy             =   7223
         Build           =   19
         TypeMask        =   1183318017
         MarkerShape     =   2
         BorderWidth     =   2
         Axis(2).Min     =   0
         Axis(2).Max     =   100
         Axis(2).TickMark=   -32767
         RGBBk           =   16777216
         nColors         =   16
         Colors          =   "QCResSch.frx":155D0
         _Data_          =   "QCResSch.frx":15670
      End
      Begin ChartfxLibCtl.ChartFX ChartFX2 
         Height          =   4095
         Left            =   10050
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   4815
         _cx             =   8493
         _cy             =   7223
         Build           =   19
         TypeMask        =   109576196
         RightGap        =   38
         MarkerShape     =   1
         Axis(0).Max     =   80
         Axis(1).Min     =   0
         Axis(1).Max     =   80
         Axis(2).Min     =   0
         Axis(2).Max     =   100
         Axis(2).Decimals=   2
         nColors         =   16
         Colors          =   "QCResSch.frx":15831
         _Data_          =   "QCResSch.frx":158D1
      End
   End
End
Attribute VB_Name = "frmQCResSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim iCalIndex   As Integer
    
    Dim iEquip_Cnt  As Integer   '선택한 장비 수
    Dim iLevel_Cnt  As Integer   '선택한 Level 수
    Dim iExam_Cnt   As Integer    '선택한 검사항목 수
    Dim iCnt        As Integer
    Dim iSum        As Currency
    Dim iMax        As Currency     'Mean + 2SD
    Dim iMin        As Currency     'Mean - 2SD
    Dim gIndex      As Integer

Private Sub cboPart_Click()
    Dim sPart As String
    Dim sLevel As String
    
    Dim i As Integer
    
    If cboPart.ListIndex < 0 Or cboPart.ListCount <= cboPart.ListIndex Then
        Exit Sub
    End If
    
    IsolateCode cboPart.List(cboPart.ListIndex)
    sPart = gCode
        
    lstLevel(1).Clear
    SQL = "Select levelno, max(levelname) from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and equipcode = '" & sPart & "' " & vbCrLf & _
          "group by levelno " & vbCrLf & _
          "order by levelno "
    res = db_select_List(gLocal, SQL, lstLevel(1), 1)

    sLevel = ""
    For i = 0 To lstLevel(1).ListCount - 1
        If lstLevel(1).Selected(i) = True Then
            'IsolateCode lstLevel(1).List(i)
            If sLevel = "" Then
                sLevel = "'" & Trim(lstLevel(1).List(i)) & "'"
            Else
                sLevel = sLevel & ", '" & Trim(lstLevel(1).List(i)) & "'"
            End If
        End If
    Next i
   
    For i = 0 To 3
        txtLotNo(i).Text = ""
        txtLotNo(i).Visible = False
        chkTwin(i).Value = 0
        chkTwin(i).Visible = False
    Next i
   
End Sub

Private Sub chkDate_Click()
Dim sLot    As String
Dim sEquip  As String
Dim i       As Integer
Dim sPart   As String

If chkDate.Value = 1 Then
    IsolateCode cboPart.Text
    sPart = Trim(gCode)

    mskQCDate(0).Enabled = True
    mskQCDate(1).Enabled = True
    cmdCalendar(0).Enabled = True
    cmdCalendar(1).Enabled = True

    '날짜Setting
    sLot = ""
    For i = 0 To lstLevel(1).ListCount - 1
        If lstLevel(1).Selected(i) = True Then
            If sLot = "" Then
                sLot = "'" & Trim(txtLotNo(i).Text) & "'"
                Exit For
            End If
        End If
    Next i
    
    If sLot = "" Then
        MsgBox "선택된 LotNo가 존재하지 않습니다.", vbInformation, "알림"
        chkDate.Value = 0
        Exit Sub
    Else
        SQL = "Select validstart,validend From qcexam " & vbCrLf & _
              " Where equipcode = '" & sPart & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And lotno = " & sLot & " "
        
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
            mskQCDate(0).Text = Left(Trim(gReadBuf(0)), 4) & "-" & Mid(Trim(gReadBuf(0)), 5, 2) & "-" & Mid(Trim(gReadBuf(0)), 7, 2)
            mskQCDate(1).Text = Left(Trim(gReadBuf(1)), 4) & "-" & Mid(Trim(gReadBuf(1)), 5, 2) & "-" & Mid(Trim(gReadBuf(1)), 7, 2)
        Else
            MsgBox "선택된 LotNo가 존재하지 않습니다.", vbInformation, "알림"
            chkDate.Value = 0
            Exit Sub
        End If
    End If

Else
    mskQCDate(0).Enabled = False
    mskQCDate(1).Enabled = False
    cmdCalendar(0).Enabled = False
    cmdCalendar(1).Enabled = False
End If
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    If Index = 0 Then
       monvCal.Top = 930
       iCalIndex = 0
    Else
       monvCal.Top = 1380
       iCalIndex = 1
    End If
    
    If monvCal.Visible = True Then
        monvCal.Visible = False
    ElseIf monvCal.Visible = False Then
        If IsDate(mskQCDate(Index).Text) Then
         monvCal.Value = mskQCDate(Index).Text
        Else
         monvCal.Value = Format(Date, "yyyy-mm-dd")
        End If
        monvCal.Visible = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreLot_Click()
    frmDisplayLot.Show
End Sub

Private Sub cmdPrint_Click()
Dim l, r, t, b As Integer
Dim px, py As Integer
Dim w, h, gap As Integer
Dim PageCount As Integer
Dim i         As Integer
Dim sHead       As String


sHead = "조회기간 : " & mskQCDate(0).Text & " ~ " & mskQCDate(1).Text
sHead = sHead & "   검사파트 : " & cboPart.Text & ""
sHead = sHead & "   Lot No. : "
For i = 0 To lstLevel(1).ListCount - 1
    If lstLevel(1).Selected(i) = True Then
        sHead = sHead & lstLevel(1).List(i) & "(" & Trim(txtLotNo(i)) & ")  "
    End If
Next i


Printer.Orientation = 2
vasNuRes.PrintUseDataMax = True

Call vasNuRes.OwnerPrintPageCount(Printer.hDC, 550, 800, (Printer.Width - 300), (Printer.Height / 2), PageCount)

For i = 1 To PageCount
    'Printer.Orientation = 2
   ' vasNuRes.PrintUseDataMax = True
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); sHead; Tab(150); "Page : " & i & "/" & PageCount
    Printer.FontSize = 9
    Call vasNuRes.OwnerPrintDraw(Printer.hDC, 550, 800, (Printer.Width - 300), (Printer.Height / 2), i)
    If i = 1 Then ' Chart 그리기
        Printer.Print ""
        px = Printer.TwipsPerPixelX
        py = Printer.TwipsPerPixelY
        w = Printer.Width
        h = Printer.Height
        gap = 100 / px
        t = 10 * gap
        b = ((h / 2) / px) - (20 * gap)
        r = (w / px) - (gap * 5)
        l = 3 * gap

        t = b + (20 * gap)
        b = (h / px) - (10 * gap)

        If ChartFX2.Visible = True Then
            r = (w / px) - (gap * 60)
                ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0
            l = (w / px) - (gap * 55)
            r = (w / px) - (gap * 2)
                ChartFX2.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0
        Else
             r = (w / px) - (gap * 5)
                ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0
        End If
    End If
    Printer.NewPage
Next i

Printer.EndDoc

End Sub

Private Sub cmdSearch_Click()
Dim sPart       As String
Dim sLot        As String
Dim sLevel      As String
Dim sExam       As String
Dim sExamCode   As String
Dim sExamDate   As String

Dim i           As Integer
Dim iRow        As Integer
Dim iCol        As Integer
Dim iLevel      As Integer

Dim iMean       As Currency
Dim sSD         As Currency
Dim sCV         As Currency

Dim sRet        As String

Dim iTmpRow As Long


ChartFX2.Visible = False

ClearSpread vasNuRes
ClearSpread vasSummary

IsolateCode cboPart.Text
sPart = Trim(gCode)

sLot = ""
sLevel = ""
iCnt = 0
For i = 0 To lstLevel(1).ListCount - 1
    If lstLevel(1).Selected(i) = True Then
        iLevel = iLevel + 1
        If sLot = "" Then
            sLot = "'" & Trim(txtLotNo(i).Text) & "'"
        Else
            sLot = sLot & ", '" & Trim(txtLotNo(i).Text) & "'"
        End If
        If sLevel = "" Then
            sLevel = "'" & Trim(lstLevel(1).List(i)) & "'"
        Else
            sLevel = sLevel & ", '" & Trim(lstLevel(1).List(i)) & "'"
        End If
    End If
Next i

If iLevel = 2 Then
    For i = 0 To lstLevel(1).ListCount - 1
        If lstLevel(1).Selected(i) = True Then
            chkTwin(i).Value = 1
        End If
    Next i
End If

    'Spread에 검사 리스트 뿌리기
    SQL = "Select equipcode, examname, levelno, levelname, t_mean, t_sd, t_mean + '/' + t_sd " & vbCrLf & _
          " from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and equipcode = '" & sPart & "' " & vbCrLf & _
          "  and levelname in (" & sLevel & ") " & vbCrLf & _
          "  and lotno in (" & sLot & ")  " & vbCrLf & _
          "order by levelno "
          
    res = db_select_Vas(gLocal, SQL, vasNuRes)
    res = db_select_Vas(gLocal, SQL, vasSummary)

    '검사일시만 먼저 column header에 세번째 칼럼부터 넣는다
    SQL = "Select distinct examdate + ' ' + mid(examtime,1,3) + '0'  " & vbCrLf & _
          "From qc_res " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and equipcode = '" & sPart & "' " & vbCrLf & _
          "  and levelname In (" & sLevel & ") "
    If chkDate.Value = 1 Then
        SQL = SQL & "And examdate Between '" & SeperatorCls(Trim(mskQCDate(0).Text)) & "' And '" & SeperatorCls(Trim(mskQCDate(1).Text)) & "' "
    End If
    SQL = SQL & " Order By 1 Desc"
    res = db_select_HVas(gLocal, SQL, vasNuRes, 0, 13)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For i = 13 To vasNuRes.MaxCols
        With vasNuRes
            .ColWidth(i) = 8
            .Row = -1
            .Col = i
            .TypeHAlign = TypeHAlignRight
        End With
    Next i

    ClearSpread Form_Main.vasTemp
    
    '검사결과를 해당 검사항목과 검사일에 맞게 넣기
    SQL = "Select equipcode, levelname, examdate + ' ' + mid(examtime,1,3) + '0' , result " & CR & _
          "From qc_res " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and equipcode = '" & sPart & "' " & vbCrLf & _
          "  and levelname In (" & sLevel & ") "
    If chkDate.Value = 1 Then
        SQL = SQL & "And examdate Between '" & SeperatorCls(Trim(mskQCDate(0).Text)) & "' And '" & SeperatorCls(Trim(mskQCDate(1).Text)) & "' "
    End If
             
    SQL = SQL & " Order BY 1, 2, 3 desc"
    res = db_select_Vas(gLocal, SQL, Form_Main.vasTemp)
   
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    iTmpRow = 1
    For iRow = 1 To vasNuRes.DataRowCnt
        iCnt = 0
        iSum = 0
        iMax = 0
        iMin = 0
        sExamCode = Trim(GetText(vasNuRes, iRow, 1))
        sLevel = Trim(GetText(vasNuRes, iRow, 4))
        '최대,최소 판정하여 Text Color 바꾸기
        '최대 iMax = Mean + 2SD (Target)
        '최소 iMin = Mean - 2SD (Target)
        If GetText(vasNuRes, iRow, 5) <> "" And GetText(vasNuRes, iRow, 5) <> "" Then
            iMax = CCur(Trim(GetText(vasNuRes, iRow, 5))) + CCur(Trim(GetText(vasNuRes, iRow, 6)))
            iMin = CCur(Trim(GetText(vasNuRes, iRow, 5))) - CCur(Trim(GetText(vasNuRes, iRow, 6)))
            SetText vasNuRes, CStr(iMax), iRow, 8
            SetText vasNuRes, CStr(iMin), iRow, 9
        End If
        
        For iCol = 13 To vasNuRes.MaxCols
            sExamDate = Trim(GetText(vasNuRes, 0, iCol))
            i = SchResult(sExamCode, sLevel, sExamDate, iTmpRow, iRow, iCol)
            If i > 0 Then
                iTmpRow = iTmpRow + 1
            End If
        Next iCol
        '각 검사항목별 MEAN
        If iCnt > 0 Then
            SetText vasNuRes, Format(iSum / iCnt, "####0.00"), iRow, 10
        End If
    Next iRow
    
    '각 검사항목별 SD, CV 계산해서 Display
    For iRow = 1 To vasNuRes.DataRowCnt
        iMean = 0
        sSD = 0
        sCV = 0
        iCnt = 0
        If Trim(GetText(vasNuRes, iRow, 8)) <> "" And Trim(GetText(vasNuRes, iRow, 10)) <> "" Then
            iMean = CCur(Trim(GetText(vasNuRes, iRow, 10)))
            For iCol = 13 To vasNuRes.MaxCols
                sRet = Trim(GetText(vasNuRes, iRow, iCol))
                If IsNumeric(sRet) Then
                    sSD = sSD + ((CCur(sRet) - iMean) * (CCur(sRet) - iMean))
                    iCnt = iCnt + 1
                End If
            Next iCol
        End If
        '각 검사항목별SD, CV 계산해서 Display
        iCnt = iCnt - 1
        If iCnt > 0 Then
            sSD = Sqr(sSD / iCnt)
            SetText vasNuRes, Format(sSD, "####0.00"), iRow, 11
            sCV = sSD / iMean * 100
            SetText vasNuRes, Format(sCV, "####0.00"), iRow, 12
        End If
    Next iRow
    
If vasNuRes.DataRowCnt < 12 Then
    vasNuRes.MaxRows = 11
Else
    vasNuRes.MaxRows = vasNuRes.DataRowCnt
End If
    
'Summary에 Mean,SD,CV넣기
For iRow = 1 To vasNuRes.DataRowCnt
    SetText vasSummary, GetText(vasNuRes, iRow, 10), iRow, 8
    SetText vasSummary, GetText(vasNuRes, iRow, 11), iRow, 9
    SetText vasSummary, GetText(vasNuRes, iRow, 12), iRow, 10
   
'Min,Max,Count 구하기
    SQL = "Select equipcode,levelname,Count(*) " & CR & _
          "From qcexam " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And equipcode = '" & GetText(vasNuRes, iRow, 1) & "' " & CR & _
          " And levelname = '" & GetText(vasNuRes, iRow, 4) & "' " & CR & _
          " And lostno In (" & sLot & ") " & CR & _
          " And result <> '' "
    
    If chkDate.Value = 1 Then
        SQL = SQL & "And examdate Between '" & Trim(mskQCDate(0).Text) & "' And '" & Trim(mskQCDate(1).Text) & "' "
    End If
             
    SQL = SQL & " Group BY equipcode,levelname"
    
    res = db_select_Col(gLocal, SQL)

    If res > 0 Then
        SetText vasSummary, gReadBuf(2), iRow, 13
    End If
    
'Max
    SQL = "Select Max(result) " & CR & _
          "From qcexam " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And equipcode = '" & GetText(vasNuRes, iRow, 1) & "' " & CR & _
          " And levelname = '" & GetText(vasNuRes, iRow, 4) & "' " & CR & _
          " And lostno In (" & sLot & ") " & CR & _
          " And result <> '' "
    
    If chkDate.Value = 1 Then
        SQL = SQL & "And examdate Between '" & Trim(mskQCDate(0).Text) & "' And '" & Trim(mskQCDate(1).Text) & "' "
    End If
             
    SQL = SQL & " Group BY equipcode,levelname"
    
    res = db_select_Col(gLocal, SQL)

    If res > 0 Then
        SetText vasSummary, gReadBuf(0), iRow, 12
    End If
    
'Min
    SQL = "Select Min(result) " & CR & _
          "From qcexam " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And equipcode = '" & GetText(vasNuRes, iRow, 1) & "' " & CR & _
          " And levelname = '" & GetText(vasNuRes, iRow, 4) & "' " & CR & _
          " And lostno In (" & sLot & ") " & CR & _
          " And result <> '' "
    
    If chkDate.Value = 1 Then
        SQL = SQL & "And examdate Between '" & Trim(mskQCDate(0).Text) & "' And '" & Trim(mskQCDate(1).Text) & "' "
    End If
             
    SQL = SQL & " Group BY equipcode,levelname"
    
    res = db_select_Col(gLocal, SQL)

    If res > 0 Then
        SetText vasSummary, gReadBuf(0), iRow, 11
    End If
    
Next iRow
    
vasNuRes_Click 1, 1

    
End Sub

 Function SchResult(argCode As String, argLevel As String, argDate As String, ByVal argStart As Integer, ByVal argSetRow As Integer, ByVal argSetCol As Integer)
    Dim compCode    As String
    Dim compLevel   As String
    Dim compDate    As String
    Dim sRet        As String
    Dim i           As Integer
    Dim iSchRow     As Integer
    
    Dim j           As Integer
    Dim sTmp        As String
    
    SchResult = -1
    For i = argStart To Form_Main.vasTemp.DataRowCnt
        compCode = Trim(GetText(Form_Main.vasTemp, i, 1))
        compLevel = Trim(GetText(Form_Main.vasTemp, i, 2))
        compDate = Trim(GetText(Form_Main.vasTemp, i, 3))
        If (compCode = argCode) And (compLevel = argLevel) And (Left(Trim(compDate), 11) = Left(Trim(argDate), 11)) Then
            SetText vasNuRes, Trim(GetText(Form_Main.vasTemp, i, 4)), argSetRow, argSetCol
            sRet = Trim(GetText(Form_Main.vasTemp, i, 4))
            SchResult = i
            If sRet <> "" Then
                If Not IsNumeric(sRet) Then
                    sTmp = ""
                    For j = 1 To Len(sRet)
                        If Mid(sRet, j, 1) <> "<" And Mid(sRet, j, 1) <> ">" And Mid(sRet, j, 1) <> "=" Then
                            sTmp = sTmp & Mid(sRet, j, 1)
                        End If
                    Next j
                    sRet = sTmp
                End If
                iSum = iSum + CCur(sRet)
                iCnt = iCnt + 1
                'Mean + 2SD (최대) Mean - 2SD (최소) ForeColor 바꾸기
                If CCur(sRet) > iMax Then
                    vasNuRes.Row = argSetRow
                    vasNuRes.Col = argSetCol
                    vasNuRes.ForeColor = RGB(205, 55, 0)
                ElseIf CCur(sRet) < iMin Then
                    vasNuRes.Row = argSetRow
                    vasNuRes.Col = argSetCol
                    vasNuRes.ForeColor = RGB(0, 0, 205)
                End If
            End If
            Exit Function
        End If
    Next i
End Function

Private Sub cmdSummary_Click()
    Dim sHead   As String
    Dim sHead1  As String
    Dim sHead2  As String
    Dim sFoot   As String
    Dim sCurDate As String
    Dim i       As Integer
    
    If vasSummary.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If

    sCurDate = GetDateFull
    
    sHead1 = mskQCDate(0).Text
    If (IsDate(mskQCDate(0).Text) = True And mskQCDate(0).Text <> mskQCDate(1).Text) Then
        sHead1 = sHead1 & " ~ " & mskQCDate(1).Text
    End If

sHead2 = "검사코드: " & cboPart.Text & ""
sHead2 = sHead2 & "   검사장비: Elecsys 2010"

sHead2 = sHead2 & "   Lot No.: "
For i = 0 To lstLevel(1).ListCount - 1
    If lstLevel(1).Selected(i) = True Then
        sHead2 = sHead2 & lstLevel(1).List(i) & "(" & Trim(txtLotNo(i)) & ") "
    End If
Next i

    sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ Summary Result ▣" & "/n/n " & _
                "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l" & "조회 일자 : " & sHead1 & "/n" & _
                "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l" & "" & sHead2 & "/n"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "SCL 부산"
    vasSummary.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasSummary.PrintAbortMsg = "인쇄중 입니다 ..."
    vasSummary.PrintJobName = "Auto LIS - 환자리스트"
    vasSummary.PrintHeader = sHead
    vasSummary.PrintFooter = sFoot
    vasSummary.PrintMarginTop = 900
    vasSummary.PrintMarginBottom = 720
'현재 SS가 비대칭으로 출력함
'    vasSummary.PrintMarginLeft = 720
    vasSummary.PrintMarginLeft = 300
    vasSummary.PrintMarginRight = 300
    
    vasSummary.PrintColor = True
    vasSummary.PrintGrid = True
'Set printing range
    vasSummary.PrintType = 0  'SS_PRINT_ALL(default)

    vasSummary.PrintShadows = True

    vasSummary.Action = 13 'SS_ACTION_PRINT
End Sub

Private Sub Command1_Click()
Dim l, r, t, b As Integer
Dim px, py As Integer
Dim w, h, gap As Integer
Dim PageCount As Integer
Dim i         As Integer
Dim sHead       As String


sHead = "  조회기간 : " & mskQCDate(0).Text & " ~ " & mskQCDate(1).Text
sHead = sHead & "   검사코드 : " & cboPart.Text & ""
sHead = sHead & "   검사장비 : Elecsys 2010"
'For i = 0 To lstLevel(1).ListCount - 1
'    If lstLevel(1).Selected(i) = True Then
'        sHead = sHead & lstLevel(1).List(i) & "(" & Trim(txtLotNo(i)) & ")  "
'    End If
'Next i

Printer.Orientation = 2
vasLevel.PrintUseDataMax = True
vasResult.PrintUseDataMax = True

Call vasResult.OwnerPrintPageCount(Printer.hDC, 550, (Printer.Height / 4) - 300, (Printer.Width - 300), (Printer.Height / 2) + 300, PageCount)

For i = 1 To PageCount
   'Printer.Orientation = 2
   'vasNuRes.PrintUseDataMax = True
    Printer.FontSize = 13
    Printer.Print ""
    Printer.Print Tab(5); "▣  Quality Control Graph  ▣ "
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print Tab(5); sHead; Tab(150); "Page : " & i & "/" & PageCount
    Printer.FontSize = 9
    Call vasLevel.OwnerPrintDraw(Printer.hDC, 550, 1000, (Printer.Width - 300), (Printer.Height / 4), 1)
    Call vasResult.OwnerPrintDraw(Printer.hDC, 550, (Printer.Height / 4) - 300, (Printer.Width - 300), (Printer.Height / 2) + 300, i)
    If i = 1 Then ' Chart 그리기
        Printer.Print ""
        px = Printer.TwipsPerPixelX
        py = Printer.TwipsPerPixelY
        w = Printer.Width
        h = Printer.Height
        gap = 100 / px
        t = 10 * gap
        b = ((h / 2) / px)
        r = (w / px) - (gap * 5)
        l = 3 * gap

        t = b + (10 * gap)
        b = (h / px) - (5 * gap)

        If ChartFX2.Visible = True Then
            r = (w / px) - (gap * 60)
                ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0
            l = (w / px) - (gap * 55)
            r = (w / px) - (gap * 2)
                ChartFX2.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0
        Else
             r = (w / px) - (gap * 5)
                ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0
        End If
    End If
    Printer.NewPage
Next i

Printer.EndDoc
End Sub

Private Sub Command2_Click()
'Dim iSeries As Integer
'Dim iCol    As Integer
'Dim i       As Integer
'Dim j       As Integer
'Dim iMax    As Integer
'Dim iMin    As Integer
'Dim jMax    As Integer
'Dim jMin    As Integer
'Dim iPoint  As Integer
'Dim xMean   As Currency
'Dim xSD     As Currency
'Dim yMean   As Currency
'Dim ySD     As Currency
'
'Dim xLow    As Currency
'Dim xMid    As Currency
'Dim xHigh   As Currency
'Dim xVal    As Currency
'
'Dim yLow    As Currency
'Dim yMid    As Currency
'Dim yHigh   As Currency
'Dim yVal    As Currency
'
'Const CHART_HIDDEN = 1E+308
'
''*************FXChart 그리기**********************************
''Series 구하기
'    iSeries = 0
'    For iCol = 2 To vasGenList.MaxCols
'        vasGenList.Row = 1
'        vasGenList.Col = iCol
'        If vasGenList.CellType = CellTypeCheckBox Then
'            If vasGenList.Value = 1 Then iSeries = iSeries + 1
'        End If
'    Next iCol
'
''Point 구하기
'    iPoint = 0
'    iPoint = vasGenList.DataRowCnt - 4
'
' i = 0
' For iCol = 2 To vasGenList.MaxCols
'    vasGenList.Row = 1
'    vasGenList.Col = iCol
'    If vasGenList.CellType = CellTypeCheckBox Then
'        i = i + 1
'        If i = 1 Then
'            xMean = CCur(Trim(GetText(vasGenList, vasGenList.DataRowCnt - 2, iCol)))
'            xSD = CCur(Trim(GetText(vasGenList, vasGenList.DataRowCnt - 1, iCol)))
'            xLow = xMean - 2 * (xSD)
'            xMid = xMean
'            xHigh = xMean + 2 * (xSD)
'        ElseIf i = 2 Then
'            yMean = CCur(Trim(GetText(vasGenList, vasGenList.DataRowCnt - 2, iCol)))
'            ySD = CCur(Trim(GetText(vasGenList, vasGenList.DataRowCnt - 1, iCol)))
'            yLow = yMean - 2 * (ySD)
'            yMid = yMean
'            yHigh = yMean + 2 * (ySD)
'        End If
'    End If
' Next iCol
'
'iMax = yHigh + 10
'iMin = yLow - 10
'jMax = xHigh + 10
'jMin = xLow - 10
'
'ChartFX2.OpenDataEx COD_VALUES Or COD_REMOVE, 1, iPoint
'ChartFX2.OpenDataEx COD_XVALUES, 1, iPoint
'
'For j = 0 To iPoint - 1
'    If GetText(vasGenList, j + 2, 3) = "" Then
'        xVal = 0
'    Else
'        xVal = CCur(GetText(vasGenList, j + 2, 3))
'    End If
'
'    If xVal > jMax Then
'        jMax = xVal
'    End If
'
'    If jMin > xVal Then
'        jMin = xVal
'    End If
'
'
'    If GetText(vasGenList, j + 2, 5) = "" Then
'           ChartFX2.ValueEx(0, j) = CHART_HIDDEN
'    Else
'        yVal = CCur(GetText(vasGenList, j + 2, 5))
'
'        If GetText(vasGenList, j + 2, 3) <> "" Then
'            If GetText(vasGenList, j + 2, 3) = "" Then
'                ChartFX2.ValueEx(0, j) = CHART_HIDDEN
'            Else
'                ChartFX2.Series(0).Xvalue(j) = xVal
'                ChartFX2.Series(0).Yvalue(j) = yVal
'            End If
'        End If
'    End If
'
'    If yVal > iMax Then
'        iMax = yVal
'    End If
'
'    If iMin > yVal Then
'        iMin = yVal
'    End If
'
'    ChartFX2.Axis(AXIS_Y).Max = iMax
'    ChartFX2.Axis(AXIS_Y).Min = iMin
'
'    ChartFX2.Axis(AXIS_X).Max = jMax
'    ChartFX2.Axis(AXIS_X).Min = jMin
'Next j
'
'ChartFX2.CloseData COD_XVALUES
'ChartFX2.CloseData COD_VALUES
'
'
''*****************Stripe Colors*****************
'    ChartFX2.OpenDataEx COD_STRIPES, 3, 0
'
'    With ChartFX2.Stripe(2)
'        .Axis = AXIS_Y
'        .From = yHigh
'        .To = iMax
'        '.Color = CHART_PALETTECOLOR Or 4
'        .Color = RGB(255, 255, 255)
'    End With
'
'If yLow < 0 Then
'    With ChartFX2.Stripe(1)
'        .Axis = AXIS_Y
'        '.From = 0
'        .To = yLow
'        .Color = RGB(255, 255, 255)
'    End With
'
'Else
'    With ChartFX2.Stripe(1)
'        .Axis = AXIS_Y
'        .From = 0
'        .To = yLow
'        .Color = RGB(255, 255, 255)
'    End With
'End If
'
'    With ChartFX2.Stripe(0)
'        .Axis = AXIS_X
'        .From = xLow
'        .To = xHigh
'        .Color = RGB(222, 243, 255)
'    End With
'    ChartFX2.CloseData COD_STRIPES
'
'ChartFX2.RGB2DBk = RGB(255, 255, 255)



End Sub

Private Sub Form_Load()
    gEquip = "Elecsys2010"
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    CommonDialog1.Flags = cdlPDReturnDC Or cdlPDPrintSetup

mskQCDate(1).Text = Data2Pict(Date, "9999-99-99")
DTPicker1.Value = Data2Pict(Date, "9999-99-99")
mskQCDate(0).Text = Data2Pict(DateAdd("m", -1, mskQCDate(1).Text), "9999-99-99")

ClearGraph ChartFX1
ClearGraph ChartFX2

    SQL = "Select equipcode + '  ' + examname from equipexam order by equipcode "
    db_select_Combo gLocal, SQL, cboPart
End Sub

Private Sub lstLevel_Click(Index As Integer)
    Dim sPart As String
    Dim sLevel As String
    Dim sDate   As String
    Dim sCnt    As String
    
    
    Dim i As Integer
    Dim iIndex As Integer
    
    ClearSpread vaSpread1
    vaSpread1.Visible = False
        
    If cboPart.ListIndex < 0 Then
        cboPart.SetFocus
        Exit Sub
    End If
    
    sDate = SeperatorCls(DTPicker1.Value)

    IsolateCode cboPart.Text
    sPart = gCode
    
    sLevel = ""

   'Lot No. 불러오기
    iIndex = lstLevel(1).ListIndex
    gIndex = iIndex
    If lstLevel(1).Selected(iIndex) = True Then
       sLevel = lstLevel(1).Text
       'sLevel = "'" & Trim(gCode) & "'"
       
       txtLotNo(iIndex).Visible = True
       chkTwin(iIndex).Visible = True
       
        SQL = "Select lotno From qcexam " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & sPart & "' " & vbCrLf & _
              "  and levelname = '" & sLevel & "' " & vbCrLf & _
              "  And ValidStart <= '" & sDate & "' " & vbCrLf & _
              "  And ValidEnd >= '" & sDate & "' "
        res = db_select_Vas(gLocal, SQL, vaSpread1)
        If res < 1 Then
            SaveQuery SQL
        End If
        sCnt = vaSpread1.DataRowCnt
        
        Select Case sCnt
        Case 0
            lstLevel(1).Selected(iIndex) = False
            txtLotNo(iIndex).Visible = False
            chkTwin(iIndex).Visible = False
            'sspMsg.Caption = " 알림 ☞ 해당조건으로 등록된 Lot No.가 존재하지 않습니다."
        Case 1
            txtLotNo(iIndex).Text = Trim(GetText(vaSpread1, 1, 1))
        
        '조회기간
        
            SQL = "Select validstart,validend From qcexam " & CR & _
                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and equipcode = '" & sPart & "' " & vbCrLf & _
                  "  and levelname = " & sLevel & " " & vbCrLf & _
                  "  And lotno = '" & txtLotNo(iIndex).Text & "' "
            
            res = db_select_Col(gLocal, SQL)
            
            If res > 0 And mskQCDate(0).Enabled = False And gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                mskQCDate(0).Text = Trim(gReadBuf(0))
                mskQCDate(1).Text = Data2Pict(Date, "9999-99-99")
            End If
            
        Case Else
            'sspMsg.Caption = " 알림 ☞ 해당 Lot No.를 선택하십시요."
            vaSpread1.Visible = True
        End Select
    Else
        txtLotNo(iIndex).Text = ""
        txtLotNo(iIndex).Visible = False
        chkTwin(iIndex).Value = 0
        chkTwin(iIndex).Visible = False
    End If
    
    
End Sub

Private Sub monvCal_DateClick(ByVal DateClicked As Date)
mskQCDate(iCalIndex).Text = Format(DateClicked, "yyyy-mm-dd")
monvCal.Visible = False
End Sub

Private Sub monvCal_LostFocus()
monvCal.Visible = False
End Sub

Private Sub mskQCDate_GotFocus(Index As Integer)
    SelectFocus mskQCDate(Index)
End Sub

Private Sub mskQCDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            mskQCDate(1).SetFocus
        End If
    End If
End Sub


Private Sub vasNuRes_Click(ByVal Col As Long, ByVal Row As Long)
Dim sExamCode       As String
Dim sExamName       As String
Dim sName           As String
Dim sRet            As String
Dim sLevel1         As String
Dim sLevel2         As String
Dim iLevel          As Integer
Dim iStr            As Long
Dim iRow            As Long
Dim jRow            As Long
Dim iCol            As Long
Dim jCol            As Long
Dim iSeries         As Integer
Dim iPoint          As Integer
Dim xRow            As Integer
Dim i               As Integer
Dim j               As Integer
Dim k               As Integer
Dim iMod            As Integer
Dim sMax(0 To 4)            As Currency
Dim sMin(0 To 4)            As Currency
Dim iGab            As Currency
Dim xMax            As Currency
Dim xMin            As Currency
Dim xSD             As Currency
Dim yMax            As Currency
Dim yMin            As Currency
Dim ySD             As Currency
Dim xVal            As Currency
Dim yVal            As Currency
Dim iMax            As Currency
Dim iMin            As Currency
Dim iMean           As Currency
Dim iSD1            As Currency
Dim iSD2            As Currency
Dim jMax            As Currency
Dim jMin            As Currency
Dim jMean           As Currency
Dim jSD1            As Currency
Dim jSD2            As Currency
Dim sDate           As String
Dim sTime           As String
Dim sResult         As String
Dim sAxis

Dim sToday          As String
Dim sYesterday      As String

Dim sTmp            As String

sToday = Data2Pict(Date, "9999-99-99")
sYesterday = DateAdd("d", -1, sToday)

Const CHART_HIDDEN = 1E+308

ClearSpread vasLevel
ClearSpread vasResult
ClearSpread vasDisplay

'*************FXChart 그리기**********************************
xRow = 0
iPoint = 0

'Point 구하기
iPoint = vasNuRes.DataColCnt - 12

'Series 구하기
iSeries = 0
iRow = vasNuRes.ActiveRow
sExamCode = Trim(GetText(vasNuRes, iRow, 1))

If sExamCode = "" Then
   Exit Sub
End If

For iRow = 1 To vasNuRes.DataRowCnt
    If sExamCode = Trim(GetText(vasNuRes, iRow, 1)) Then
        iSeries = iSeries + 1
        If iSeries = 1 Then
            xRow = iRow
        End If
        
        If Trim(GetText(vasNuRes, iRow, 2)) <> "" Then
           sExamName = Trim(GetText(vasNuRes, iRow, 2))
        End If
    End If
Next iRow

If xRow = 0 Then
    Exit Sub
End If

'X축의 간격
If iPoint > 20 Then
   ChartFX1.Scrollable = True
   ChartFX1.Axis(AXIS_X).PixPerUnit = 30
   ChartFX1.Axis(AXIS_X).LabelValue = 2
   'ChartFX1.Axis(AXIS_X).Style = AS_2LEVELS
End If

'Chart Title
ChartFX1.Title(CHART_TOPTIT) = sExamName
ChartFX1.Fonts(CHART_TOPTIT) = CF_BOLD Or CF_ITALIC Or 12
ChartFX1.RGBFont(CHART_TOPTIT) = RGB(0, 0, 0)

ChartFX1.OpenDataEx COD_VALUES Or COD_REMOVE, iSeries, iPoint '/Graph 그리기
ChartFX1.OpenDataEx COD_CONSTANTS, iSeries * 5, 0             '/Constant Line 그리기

i = 0
iMod = 0
If GetText(vasNuRes, xRow, 8) <> "" And GetText(vasNuRes, xRow, 6) <> "" And GetText(vasNuRes, xRow, 9) <> "" Then
    sMax(0) = CCur(Trim(GetText(vasNuRes, xRow, 8))) + (CCur(Trim(GetText(vasNuRes, xRow, 6))) / 2)
    sMin(0) = CCur(Trim(GetText(vasNuRes, xRow, 9))) - (CCur(Trim(GetText(vasNuRes, xRow, 6))) / 2)
Else
'    sMax(0) = 1
'    sMin(0) = 0
    MsgBox "Target Mean/SD가 설정되지 않았습니다.", vbInformation, "확인"
    ClearGraph ChartFX1
    Exit Sub
End If

For iRow = xRow To xRow + iSeries - 1
'************************검사코드,레벨,Mean,SD,CV*******************
    For jCol = 1 To 12
        SetText vasLevel, GetText(vasNuRes, iRow, jCol), i + 1, jCol
    Next jCol
'*******************************************************************
    SetText vasDisplay, GetText(vasNuRes, iRow, 4), 0, i + 3
jRow = 1

'************************검사일자,레벨별 결과***********************
    For jCol = 13 To vasNuRes.MaxCols
        If i = 0 Then
            IsolateCode Trim(GetText(vasNuRes, 0, jCol))
            sDate = Trim(gCode)
            sTime = Trim(gName)
            SetText vasDisplay, sDate, jRow, 1
            SetText vasDisplay, sTime, jRow, 2
        End If
        
        sResult = Trim(GetText(vasNuRes, iRow, jCol))
        SetText vasDisplay, sResult, jRow, i + 3
        
        jRow = jRow + 1
    
    Next jCol
'********************************************************************
    Select Case i
    Case 0
        ChartFX1.Series(i).Color = RGB(105, 89, 205)
        ChartFX1.Axis(AXIS_Y).TextColor = RGB(105, 89, 205)
        iGab = 0
    Case 1
        ChartFX1.Series(i).Color = RGB(238, 64, 0)
        ChartFX1.Axis(AXIS_Y2).TextColor = RGB(238, 64, 0)
        iGab = 0
    Case 2
        ChartFX1.Series(i).Color = RGB(0, 139, 69)
        iGab = CCur(Trim(GetText(vasNuRes, iRow, 5))) - CCur(Trim(GetText(vasNuRes, iRow - 2, 5)))
    Case 3
        ChartFX1.Series(i).Color = RGB(255, 193, 37)
        iGab = CCur(Trim(GetText(vasNuRes, iRow, 5))) - CCur(Trim(GetText(vasNuRes, iRow - 2, 5)))
    End Select
    ChartFX1.Series(i).MarkerShape = i + 1
  
    iMod = i Mod 2
    If iMod = 0 Then
        sAxis = AXIS_Y
        ChartFX1.Series(i).YAxis = AXIS_Y
    Else
        sAxis = AXIS_Y2
        ChartFX1.Series(i).YAxis = AXIS_Y2
    End If
    
    
    If iMod <> 0 Then
        If GetText(vasNuRes, iRow, 8) <> "" And GetText(vasNuRes, iRow, 6) <> "" And GetText(vasNuRes, iRow, 9) <> "" Then
          sMax(iMod) = CCur(Trim(GetText(vasNuRes, iRow, 8))) + (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2)
          sMin(iMod) = CCur(Trim(GetText(vasNuRes, iRow, 9))) - (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2)
        Else
          sMax(iMod) = CCur(Trim(GetText(vasNuRes, iRow, 10))) + 1
          sMin(iMod) = CCur(Trim(GetText(vasNuRes, iRow, 10))) - 1
        End If
    End If
    
'    ChartFX1.Axis(sAxis).Max = sMax(iMod)
'    ChartFX1.Axis(sAxis).Min = sMin(iMod)
    If i = 0 Or i = 1 Then
        ChartFX1.Axis(sAxis).Max = sMax(iMod)
        ChartFX1.Axis(sAxis).Min = sMin(iMod)
        ChartFX1.Axis(sAxis).STEP = (sMax(iMod) - sMin(iMod)) / 6
    End If
    
    j = 0
    ChartFX1.ThisSerie = i
    sName = Trim(GetText(vasNuRes, iRow, 4))
    ChartFX1.Series(i).Legend = sName
    
    For iCol = 13 To vasNuRes.MaxCols
        sRet = Trim(GetText(vasNuRes, iRow, iCol))
        If Not IsNumeric(sRet) Then
            sTmp = ""
            For k = 1 To Len(sRet)
                If Mid(sRet, k, 1) <> "<" And Mid(sRet, k, 1) <> ">" And Mid(sRet, k, 1) <> "=" Then
                    sTmp = sTmp & Mid(sRet, k, 1)
                End If
            Next k
            sRet = sTmp
        End If
        If sRet = "" Then '값이 존재하지 않으면
            ChartFX1.Value(j) = CHART_HIDDEN
'            If Trim(GetText(vasNuRes, iRow, 10)) <> "" Then
'                ChartFX1.Value(j) = CCur(Trim(GetText(vasNuRes, iRow, 10))) - iGab
'            End If
        Else
'            If sRet > sMax(iMod) Then
'                ChartFX1.Axis(sAxis).Max = sRet
'                sMax(iMod) = sRet
'            ElseIf sRet < sMin(iMod) Then
'                ChartFX1.Axis(sAxis).Min = sRet
'                sMin(iMod) = sRet
'            End If
            ChartFX1.Value(j) = CCur(sRet) - iGab
        End If
        'IsolateCode Trim(GetText(vasNuRes, 0, iCol))
        ChartFX1.Legend(j) = Trim(GetText(vasNuRes, 0, iCol))
        j = j + 1
    Next iCol

If i = 2 Or i = 3 Then
    j = i Mod 2
    
    If Trim(GetText(vasNuRes, iRow, 5)) <> "" Then
        'Mean
        With ChartFX1.ConstantLine(3 * (i - 2))
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 5))) - iGab
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            '.Label = Trim(GetText(vasNuRes, iRow, 5))
            .Label = "MEAN"
            'set the label of the constant line to be aligned to the right
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
    End If
        
    sRet = Trim(GetText(vasNuRes, iRow, 8))
    If sRet <> "" Then
        'Mean + 2SD
        With ChartFX1.ConstantLine(3 * (i - 2) + 1)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 8))) - iGab
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            '.Label = Trim(GetText(vasNuRes, iRow, 8))
            .Label = "Mean + 2SD"
            'set the label of the constant line to be aligned to the right
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
        'Mean + 1SD
        With ChartFX1.ConstantLine(3 * (i - 2) + 2)
'            .Value = (CCur(Trim(GetText(vasNuRes, iRow, 8))) / 2) - iGab
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 5))) + (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2) - iGab
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            '.Label = CCur(Trim(GetText(vasNuRes, iRow, 5))) + (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2)
            .Label = "Mean + 1SD"
            'set the label of the constant line to be aligned to the right
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
    End If
        
    sRet = Trim(GetText(vasNuRes, iRow, 9))
    If sRet <> "" Then
        'Mean - 2SD
        With ChartFX1.ConstantLine(3 * (i - 2) + 3)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 9))) - iGab
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            '.Label = Trim(GetText(vasNuRes, iRow, 9))
            .Label = "Mean - 2SD"
            'set the label of the constant line to be aligned to the right
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
        'Mean - 1SD
        With ChartFX1.ConstantLine(3 * (i - 2) + 4)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 5))) - (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2) - iGab
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            '.Label = CCur(Trim(GetText(vasNuRes, iRow, 5))) - (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2)
            .Label = "Mean - 1SD"
            'set the label of the constant line to be aligned to the right
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
    End If
    ChartFX1.RGBFont(CHART_FIXEDFT) = RGB(0, 139, 69)
    
ElseIf i = iSeries - 1 And i <> 2 And i <> 3 Then
    j = i Mod 2
    
    If Trim(GetText(vasNuRes, iRow, 5)) <> "" Then
        'Mean
        With ChartFX1.ConstantLine(0)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 5)))
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            .Label = "MEAN"
            '.LineColor = ChartFX1.Series(i).Color
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
    End If
        
    sRet = Trim(GetText(vasNuRes, iRow, 8))
    If sRet <> "" Then
        'Mean + 2SD
        With ChartFX1.ConstantLine(1)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 8)))
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .Label = "Mean + 2SD"
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
        'Mean + 1SD
        With ChartFX1.ConstantLine(2)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 5))) + (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2)
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .Label = "Mean + 1SD"
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
    End If
        
    sRet = Trim(GetText(vasNuRes, iRow, 9))
    If sRet <> "" Then
        'Mean - 2SD
        With ChartFX1.ConstantLine(3)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 9)))
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .Label = "Mean - 2SD"
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
        'Mean - 1SD
        With ChartFX1.ConstantLine(4)
            .Value = CCur(Trim(GetText(vasNuRes, iRow, 5))) - (CCur(Trim(GetText(vasNuRes, iRow, 6))) / 2)
            If j = 0 Then
            .Axis = AXIS_Y 'horizontal
            Else
            .Axis = AXIS_Y2
            End If
            If j = 1 Then
            .Style = .Style Or CC_RIGHTALIGNED
            End If
            '.LineColor = ChartFX1.Series(i).Color
            .Label = "Mean - 1SD"
            .LineStyle = 2  'CHART_DOT
            .LineWidth = 2    '3
        End With
    End If
   
End If
'    'the interval between 2 consecutive tickmarks is equal to 10
'    ChartFX1.Axis(AXIS_Y).LabelValue = 10
'
'    'assign each label
'    For i = 0 To 9
'        ChartFX1.Axis(AXIS_Y).Label(i) = "Label " & Str(i)
'    Next i
'
'    With ChartFX1.Axis(AXIS_Y)
'        'set the text color of the y axis labels
'        .TextColor = RGB(255, 255, 255)
'        'set the font properties for the y axis labels
'        With .Font
'            .Bold = False
'            .Italic = True
'            .size = 10
'            .Name = "Times New Roman"
'        End With
'    End With
   
    i = i + 1
Next iRow

ChartFX1.CloseData COD_VALUES
ChartFX1.CloseData COD_CONSTANTS

'******************************************************************************
'************************* Twin Ploat 그리기 **********************************
iLevel = 0
For i = 0 To 3
 If chkTwin(i).Value = 1 Then
    iLevel = iLevel + 1
    If iLevel = 1 Then
        lstLevel(1).ListIndex = i
        sLevel1 = lstLevel(1).Text
    Else
        lstLevel(1).ListIndex = i
        sLevel2 = lstLevel(1).Text
    End If
 End If
Next i

ChartFX2.MultipleColors = True

ChartFX2.OpenDataEx COD_VALUES Or COD_REMOVE, 1, iPoint
ChartFX2.OpenDataEx COD_XVALUES, 1, iPoint
ChartFX2.OpenDataEx COD_COLORS, iPoint, 0
If iLevel = 2 Then
    For i = xRow To xRow + iSeries - 1
        If sLevel1 = GetText(vasNuRes, i, 4) Then
            iRow = i
        ElseIf sLevel2 = GetText(vasNuRes, i, 4) Then
            jRow = i
        End If
    Next i

    If Trim(GetText(vasNuRes, iRow, 8)) = "" Or Trim(GetText(vasNuRes, iRow, 9)) = "" Or Trim(GetText(vasNuRes, iRow, 6)) = "" Or Trim(GetText(vasNuRes, jRow, 8)) = "" Or Trim(GetText(vasNuRes, jRow, 9)) = "" Or Trim(GetText(vasNuRes, jRow, 6)) = "" Then
        MsgBox "Target Mean Or SD가 설정되지 않았습니다.", vbInformation, "확인"
        ChartFX2.Visible = False
        Exit Sub
    End If

    ChartFX2.Visible = True

    xMax = Trim(GetText(vasNuRes, iRow, 8))
    xMin = Trim(GetText(vasNuRes, iRow, 9))
    xSD = Trim(GetText(vasNuRes, iRow, 6))
    yMax = Trim(GetText(vasNuRes, jRow, 8))
    yMin = Trim(GetText(vasNuRes, jRow, 9))
    ySD = Trim(GetText(vasNuRes, jRow, 6))
    
    iMax = xMax
    iMin = xMin
    iMean = Trim(GetText(vasNuRes, iRow, 5))
    iSD1 = iMean - (xSD / 2)
    iSD2 = iMean + (xSD / 2)
    jMax = yMax
    jMin = yMin
    jMean = Trim(GetText(vasNuRes, jRow, 5))
    jSD1 = jMean - (ySD / 2)
    jSD2 = jMean + (ySD / 2)
    
    j = 0
    For iCol = 13 To vasNuRes.DataColCnt
        '특정 포인트 칼라 바꾸기
        IsolateCode Trim(GetText(vasNuRes, 0, iCol))
        If SeperatorCls(sToday) = gCode Then
           ' ChartFX2.Series(0).Color = RGB(205, 55, 0) 'Red
            ChartFX2.Color(j) = RGB(205, 55, 0) 'Red
        ElseIf SeperatorCls(sYesterday) = gCode Then
            'ChartFX2.Series(0).Color = RGB(255, 174, 185) 'Pink
            ChartFX2.Color(j) = RGB(255, 215, 0) 'Yellow
        Else
            'ChartFX2.Series(0).Color = RGB(16, 78, 139) 'Blue
            ChartFX2.Color(j) = RGB(16, 78, 139) 'Blue
        End If
    
    
        'xVal = Trim(GetText(vasNuRes, xRow, iCol))
        If IsNumeric(Trim(GetText(vasNuRes, iRow, iCol))) = False Then
            ChartFX2.ValueEx(0, j) = CHART_HIDDEN
        Else
            'yVal = Trim(GetText(vasNuRes, xRow + 1, iCol))
            If IsNumeric(Trim(GetText(vasNuRes, jRow, iCol))) = False Then
                ChartFX2.ValueEx(0, j) = CHART_HIDDEN
            Else
                xVal = Trim(GetText(vasNuRes, iRow, iCol))
                yVal = Trim(GetText(vasNuRes, jRow, iCol))
                
                ChartFX2.Series(0).Xvalue(j) = xVal
                ChartFX2.Series(0).Yvalue(j) = yVal

                If xVal > xMax Then
                    xMax = xVal
                ElseIf xVal < xMin Then
                    xMin = xVal
                End If

                If yVal > yMax Then
                    yMax = yVal
                ElseIf yVal < yMin Then
                    yMin = yVal
                End If
            End If
        End If

        j = j + 1
    Next iCol

    yMax = yMax + ySD
    yMin = yMin - ySD
    xMax = xMax + xSD
    xMin = xMin - xSD

    ChartFX2.Axis(AXIS_Y).Max = yMax
    ChartFX2.Axis(AXIS_Y).Min = yMin

    ChartFX2.Axis(AXIS_X).Max = xMax
    ChartFX2.Axis(AXIS_X).Min = xMin

    ChartFX2.CloseData COD_XVALUES
    ChartFX2.CloseData COD_VALUES
    ChartFX2.CloseData COD_COLORS
    
    '*****************Stripe Colors*****************
    ChartFX2.OpenDataEx COD_CONSTANTS, iSeries * 4, 0             '/Constant Line 그리기
    ChartFX2.OpenDataEx COD_STRIPES, 9, 0

    With ChartFX2.Stripe(0)
        .Axis = AXIS_X
        .From = iSD1
        .To = iSD2
        .Color = RGB(0, 191, 255)
    End With

    With ChartFX2.Stripe(1)
        .Axis = AXIS_Y
        .From = jSD2
        .To = yMax
        '.Color = CHART_PALETTECOLOR Or 4
        .Color = RGB(222, 243, 255)
    End With
    
    With ChartFX2.Stripe(2)
        .Axis = AXIS_Y
        .From = jSD1
        .To = yMin
        '.Color = CHART_PALETTECOLOR Or 4
        .Color = RGB(222, 243, 255)
    End With

    With ChartFX2.Stripe(3)
        .Axis = AXIS_X
        .From = iMin
        .To = iSD1
        .Color = RGB(222, 243, 255)
    End With

    With ChartFX2.Stripe(4)
        .Axis = AXIS_X
        .From = iSD2
        .To = iMax
        .Color = RGB(222, 243, 255)
    End With

    With ChartFX2.Stripe(5)
        .Axis = AXIS_Y
        .From = jMax
        .To = yMax
        '.Color = CHART_PALETTECOLOR Or 4
        .Color = RGB(255, 255, 255)
    End With

    If yMin < jMin Then
        With ChartFX2.Stripe(6)
            .Axis = AXIS_Y
            .From = yMin
            .To = jMin
            .Color = RGB(255, 255, 255)
        End With
    Else
        With ChartFX2.Stripe(6)
            .Axis = AXIS_Y
            .From = 0
            .To = jMin
            .Color = RGB(255, 255, 255)
        End With
    End If

    With ChartFX2.Stripe(7)
        .Axis = AXIS_X
        .From = iMin
        .To = xMin
        .Color = RGB(255, 255, 255)
    End With

    With ChartFX2.Stripe(8)
        .Axis = AXIS_X
        .From = iMax
        .To = xMax
        .Color = RGB(255, 255, 255)
    End With




'    With ChartFX2.Stripe(2)
'        .Axis = AXIS_Y
'        .From = jMax
'        .To = yMax
'        '.Color = CHART_PALETTECOLOR Or 4
'        .Color = RGB(255, 255, 255)
'    End With
'
'    If yMin < jMin Then
'        With ChartFX2.Stripe(1)
'            .Axis = AXIS_Y
'            .From = yMin
'            .To = jMin
'            .Color = RGB(255, 255, 255)
'        End With
'    Else
'        With ChartFX2.Stripe(1)
'            .Axis = AXIS_Y
'            .From = 0
'            .To = jMin
'            .Color = RGB(255, 255, 255)
'        End With
'    End If
'
'    With ChartFX2.Stripe(0)
'        .Axis = AXIS_X
'        .From = iMin
'        .To = iMax
'        .Color = RGB(222, 243, 255)
'    End With

    'Mean
    With ChartFX2.ConstantLine(0)
        .Value = jMean
        .Axis = AXIS_Y 'horizontal
        .Label = "Mean"
        'set the label of the constant line to be aligned to the right
        '.LineColor = RGB(34, 139, 34)   'Green
        .LineStyle = CHART_SOLID
        .LineWidth = 1
    End With

'    'Mean + 2SD
'    With ChartFX2.ConstantLine(1)
'        .Value = jMax
'        .Axis = AXIS_Y 'horizontal
'        .Label = "Mean + 2SD"
'        'set the label of the constant line to be aligned to the right
'        .LineColor = RGB(34, 139, 34)   'Green
'        .LineStyle = CHART_DOT
'        .LineWidth = 2
'    End With
'
'    'Mean - 2SD
'    With ChartFX2.ConstantLine(2)
'        .Value = jMin
'        .Axis = AXIS_Y 'horizontal
'        .Label = "Mean - 2SD"
'        'set the label of the constant line to be aligned to the right
'        .LineColor = RGB(34, 139, 34)   'Green
'        .LineStyle = CHART_DOT
'        .LineWidth = 2
'    End With

    With ChartFX2.ConstantLine(1)
        .Value = iMean
        .Axis = AXIS_X 'vertical
        .Label = "Mean"
        'set the label of the constant line to be aligned to the right
        '.LineColor = RGB(255, 69, 0)  'Red
        .LineStyle = CHART_SOLID
        .LineWidth = 1
    End With

'    'Mean + 2SD
'    With ChartFX2.ConstantLine(4)
'        .Value = iMax
'        .Axis = AXIS_X 'vertical
'        .Label = "Mean + 2SD"
'        'set the label of the constant line to be aligned to the right
'        .LineColor = RGB(255, 69, 0)  'Red
'        .LineStyle = CHART_DOT
'        .LineWidth = 2
'    End With
'
'    'Mean - 2SD
'    With ChartFX2.ConstantLine(5)
'        .Value = iMin
'        .Axis = AXIS_X 'vertical
'        .Label = "Mean - 2SD"
'        'set the label of the constant line to be aligned to the right
'        .LineColor = RGB(255, 69, 0)  'Red
'        .LineStyle = CHART_DOT
'        .LineWidth = 2
'    End With

'    'Mean + 1SD
'    With ChartFX2.ConstantLine(8)
'        .Value = iSD2
'        .Axis = AXIS_X 'vertical
'        '.Label = "Mean + 2SD"
'        'set the label of the constant line to be aligned to the right
'        .LineColor = RGB(255, 69, 0)  'Red
'        .LineStyle = CHART_DOT
'        .LineWidth = 2
'    End With
'
'    'Mean - 1SD
'    With ChartFX2.ConstantLine(9)
'        .Value = iSD1
'        .Axis = AXIS_X 'vertical
'        '.Label = "Mean - 2SD"
'        'set the label of the constant line to be aligned to the right
'        .LineColor = RGB(255, 69, 0)  'Red
'        .LineStyle = CHART_DOT
'        .LineWidth = 2
'    End With

    ChartFX2.CloseData COD_STRIPES
    ChartFX2.CloseData COD_CONSTANTS

    ChartFX2.RGB2DBk = RGB(255, 255, 255)

Else
    ChartFX2.Visible = False
End If

'************vasDisplay에 나열된것 다시 정열하기*********************
For iRow = 1 To vasDisplay.DataRowCnt
    jRow = iRow Mod 12
    iStr = iRow \ 12

    If jRow = 0 Then
       jRow = 12
       iStr = iStr - 1
    End If
    iStr = iStr * (iSeries + 2)
    For iCol = 1 To (iSeries + 2)
        SetText vasResult, GetText(vasDisplay, iRow, iCol), jRow, iStr + iCol
        If jRow = 1 Then
            SetText vasResult, GetText(vasDisplay, 0, iCol), 0, iStr + iCol
        End If
    Next iCol

Next iRow


'정열하기
For iCol = 1 To vasResult.MaxCols
    jCol = iCol Mod (iSeries + 2)
    'If jCol = 0 Then jCol = iSeries + 2
    
            Select Case jCol
            Case 1
                With vasResult
                    .ColWidth(iCol) = 10
                    .Row = -1
                    .Col = iCol
                    .TypeHAlign = TypeHAlignCenter
                End With
            Case 2
                With vasResult
                    .ColWidth(iCol) = 7
                    .Row = -1
                    .Col = iCol
                    .TypeHAlign = TypeHAlignCenter
                End With
            Case Else
                With vasResult
                    .ColWidth(iCol) = 8.5
                    .Row = -1
                    .Col = iCol
                    .TypeHAlign = TypeHAlignRight
                End With
            End Select
Next iCol

iRow = 1
If vasLevel.DataRowCnt > 0 Then
    For i = 0 To lstLevel(1).ListCount - 1
        If lstLevel(1).Selected(i) = True Then
            SetText vasLevel, Trim(txtLotNo(i)), iRow, 5
            iRow = iRow + 1
            'sHead = sHead & lstLevel(1).List(i) & "(" & Trim(txtLotNo(i)) & ")  "
        End If
    Next i
End If
'**********************************************************************



End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sLotNo  As String

sLotNo = Trim(GetText(vaSpread1, Row, 1))

If sLotNo <> "" Then
    txtLotNo(gIndex).Text = sLotNo
    vaSpread1.Visible = False
End If
End Sub

