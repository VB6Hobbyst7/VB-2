VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmCalData 
   Caption         =   "Calendar 통계"
   ClientHeight    =   6705
   ClientLeft      =   840
   ClientTop       =   765
   ClientWidth     =   9885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9885
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   270
      TabIndex        =   52
      Top             =   90
      Width           =   4965
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   979
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.ComboBox cmbYear 
         Height          =   300
         Left            =   945
         Style           =   2  '드롭다운 목록
         TabIndex        =   55
         Top             =   135
         Width           =   1230
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   300
         Left            =   2160
         Style           =   2  '드롭다운 목록
         TabIndex        =   54
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox txtSysDate 
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "조회년월"
         Height          =   240
         Left            =   90
         TabIndex        =   56
         Top             =   180
         Width           =   780
      End
   End
   Begin FPSpreadADO.fpSpread sprItem 
      Height          =   7125
      Left            =   6660
      TabIndex        =   51
      Top             =   675
      Width           =   5055
      _Version        =   196608
      _ExtentX        =   8916
      _ExtentY        =   12568
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   120
      ScrollBars      =   2
      SpreadDesigner  =   "frmCalData.frx":0000
      Appearance      =   2
   End
   Begin FPSpreadADO.fpSpread sprSLip 
      Height          =   4200
      Left            =   270
      TabIndex        =   50
      Top             =   3465
      Width           =   6135
      _Version        =   196608
      _ExtentX        =   10821
      _ExtentY        =   7408
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   40
      ScrollBars      =   2
      SpreadDesigner  =   "frmCalData.frx":1035
      Appearance      =   2
   End
   Begin Threed.SSPanel panelCalendar 
      Height          =   2670
      Left            =   270
      TabIndex        =   0
      Top             =   675
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   4710
      _StockProps     =   15
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   0
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "일"
         ForeColor       =   255
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   1
         Left            =   900
         TabIndex        =   2
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "월"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   2
         Left            =   1770
         TabIndex        =   3
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "화"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "수"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   4
         Left            =   3510
         TabIndex        =   5
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "목"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   5
         Left            =   4380
         TabIndex        =   6
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "금"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdWeek 
         Height          =   375
         Index           =   6
         Left            =   5250
         TabIndex        =   7
         Top             =   0
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "토"
         ForeColor       =   16711680
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   1
         Left            =   900
         TabIndex        =   9
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   2
         Left            =   1770
         TabIndex        =   10
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   11
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   4
         Left            =   3510
         TabIndex        =   12
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   5
         Left            =   4380
         TabIndex        =   13
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   6
         Left            =   5250
         TabIndex        =   14
         Top             =   405
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   7
         Left            =   0
         TabIndex        =   15
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   8
         Left            =   900
         TabIndex        =   16
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   9
         Left            =   1770
         TabIndex        =   17
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   10
         Left            =   2640
         TabIndex        =   18
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   11
         Left            =   3510
         TabIndex        =   19
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   12
         Left            =   4380
         TabIndex        =   20
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   13
         Left            =   5250
         TabIndex        =   21
         Top             =   780
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   14
         Left            =   0
         TabIndex        =   22
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   15
         Left            =   900
         TabIndex        =   23
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   16
         Left            =   1770
         TabIndex        =   24
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   17
         Left            =   2640
         TabIndex        =   25
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   18
         Left            =   3510
         TabIndex        =   26
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   19
         Left            =   4380
         TabIndex        =   27
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   20
         Left            =   5250
         TabIndex        =   28
         Top             =   1155
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   21
         Left            =   0
         TabIndex        =   29
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   22
         Left            =   900
         TabIndex        =   30
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   23
         Left            =   1770
         TabIndex        =   31
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   24
         Left            =   2640
         TabIndex        =   32
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   25
         Left            =   3510
         TabIndex        =   33
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   26
         Left            =   4380
         TabIndex        =   34
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   27
         Left            =   5250
         TabIndex        =   35
         Top             =   1530
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   28
         Left            =   0
         TabIndex        =   36
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   29
         Left            =   900
         TabIndex        =   37
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   30
         Left            =   1770
         TabIndex        =   38
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   31
         Left            =   2640
         TabIndex        =   39
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   32
         Left            =   3510
         TabIndex        =   40
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   33
         Left            =   4380
         TabIndex        =   41
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   34
         Left            =   5250
         TabIndex        =   42
         Top             =   1905
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   35
         Left            =   0
         TabIndex        =   43
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   36
         Left            =   900
         TabIndex        =   44
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   37
         Left            =   1770
         TabIndex        =   45
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   38
         Left            =   2640
         TabIndex        =   46
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   39
         Left            =   3510
         TabIndex        =   47
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   40
         Left            =   4380
         TabIndex        =   48
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDay 
         Height          =   375
         Index           =   41
         Left            =   5250
         TabIndex        =   49
         Top             =   2280
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmCalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nLastDate        As Integer
Public nFirstWeek       As Integer

Public Function Select_LastDate(ByVal sYear As String, ByVal sMonth As String) As Integer

    Dim sDate   As String
    Dim i       As Integer
    
    For i = 31 To 28 Step -1
        sDate = sYear & "-" & sMonth & "-" & Trim(Str(i))
        If IsDate(sDate) Then
            Select_LastDate = i
            Exit Function
        End If
    Next
    

End Function
Public Function Select_FirstWeek(ByVal sYear As String, ByVal sMonth As String) As Integer
    Dim sDate   As String
    
    sDate = sYear & "-" & sMonth & "-" & "01"
    
    Select Case Format(sDate, "aaa")
        Case "일": Select_FirstWeek = 0
        Case "월": Select_FirstWeek = 1
        Case "화": Select_FirstWeek = 2
        Case "수": Select_FirstWeek = 3
        Case "목": Select_FirstWeek = 4
        Case "금": Select_FirstWeek = 5
        Case "토": Select_FirstWeek = 6
    End Select
    
End Function
Public Sub CalDisplay()
    Dim i           As Integer
    Dim nDay        As Integer
    Dim nOpCnt      As Integer
    Dim nLaCnt      As Integer
    Dim nDispCnt    As Integer
    Dim adoCount    As ADODB.Recordset
    
    For i = 0 To 41
        frmCalData.cmdDay(i).Caption = ""
        frmCalData.cmdDay(i).Tag = ""
    Next
    
    
    nDay = 1
    For i = nFirstWeek To (nLastDate + nFirstWeek) - 1
        frmCalData.cmdDay(i).Caption = Format(Trim(Str(nDay)), "@@")
        frmCalData.cmdDay(i).Tag = frmCalData.cmbYear.Text & "-" & Left(frmCalData.cmbMonth, 2) & "-" & Format(Trim$(Str(nDay)), "00")
        If frmCalData.cmdDay(i).Tag <> "" Then
            GoSub Get_GeneralOrder
            If nDispCnt = 0 Then
                frmCalData.cmdDay(i).Caption = frmCalData.cmdDay(i).Caption & "  ( )"
            Else
                frmCalData.cmdDay(i).Caption = frmCalData.cmdDay(i).Caption & " (" & Format(Trim$(Str(nDispCnt)), "@@") & ")"
            End If
        End If
        nDay = nDay + 1
    Next
    Exit Sub

Get_GeneralOrder:
    StrSql = ""
    StrSql = StrSql & " SELECT COUNT(*) Count"
    StrSql = StrSql & " FROM   TWEXAM_ORDER"
    StrSql = StrSql & " WHERE  COLLDate = TO_DATE('" & frmCalData.cmdDay(i).Tag & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    JEOBSUYN = '*'"
    StrSql = StrSql & " AND    SLIPNO1 > 0 "
    StrSql = StrSql & " AND    SLIPNO1 < 52"
    If False = adoSetOpen(StrSql, adoCount) Then Return
    nDispCnt = adoCount.Fields("Count").Value & ""
    Call adoSetClose(adoCount)
    Return

End Sub




Private Sub cmbMonth_Click()
    nLastDate = Select_LastDate(cmbYear, Left(cmbMonth, 2))
    nFirstWeek = Select_FirstWeek(cmbYear, Left(cmbMonth, 2))
    Call CalDisplay

End Sub

Private Sub cmbYear_Click()
    
    nLastDate = Select_LastDate(cmbYear, Left(cmbMonth, 2))
    nFirstWeek = Select_FirstWeek(cmbYear, Left(cmbMonth, 2))
    DoEvents: Call CalDisplay

End Sub

Public Sub cmdDay_Click(Index As Integer)
    Dim iSLno       As Integer
    Dim sCount      As Integer
    
    txtSysDate.Text = Format(cmdDay(Index).Tag, "yyyy-MM-dd")
    
    Call SpreadSetClear(sprSLip)
    GoSub Get_SLIPcount
    
    GoSub Set_Color
    
    
    Exit Sub
    
Set_Color:
    For i = 0 To 41
        If frmCalData.cmdDay(i).Tag = txtSysDate.Text Then
            frmCalData.cmdDay(i).ForeColor = RGB(0, 0, 255)
        Else
            frmCalData.cmdDay(i).ForeColor = RGB(0, 0, 0)
        End If
    Next
    
    Return
    
    
Get_SLIPcount:
    StrSql = ""
    StrSql = StrSql & "  SELECT a.SLIPNO1, b.Codenm,  COUNT(*) Count"
    StrSql = StrSql & "  FROM   TWEXAM_ORDER   a,"
    StrSql = StrSql & "         TWEXAM_SPECODE b"
    StrSql = StrSql & "  WHERE  a.COLLDate = TO_DATE('" & txtSysDate.Text & "','YYYY-MM-DD') "
    StrSql = StrSql & "  AND    a.JeobsuYn = '*'"
    StrSql = StrSql & "  AND    a.SLipno1  > 0 "
    StrSql = StrSql & "  AND    a.SLipno1  < 52"
    StrSql = StrSql & "  AND    a.SLipno1  = b.Codeky"
    StrSql = StrSql & "  AND    b.Codegu   = '12'"
    StrSql = StrSql & "  GROUP BY SLIPNO1, b.Codenm"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprSLip.Row = sprSLip.DataRowCnt + 1
        sprSLip.Col = 2: sprSLip.Text = adoSet.Fields("SLipno1").Value & ""
                         iSLno = Val(adoSet.Fields("SLipno1").Value & "")
        sprSLip.Col = 3: sprSLip.Text = adoSet.Fields("Codenm").Value & ""
        sprSLip.Col = 4: GoSub Get_PtCount
                         sprSLip.Text = sCount
        sprSLip.Col = 5: sprSLip.Text = adoSet.Fields("Count").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
Get_PtCount:
    Dim adoPtCnt        As ADODB.Recordset
    
    sCount = 0
    StrSql = ""
    StrSql = StrSql & " SELECT ptno"
    StrSql = StrSql & " FROM   TWEXAM_Order"
    StrSql = StrSql & " WHERE  COLLDate = TO_DATE('" & txtSysDate.Text & "','YYYY-MM-DD') "
    StrSql = StrSql & " AND    JeobsuYN = '*'"
    StrSql = StrSql & " AND    SLipno1  = " & iSLno
    StrSql = StrSql & " GROUP  BY Ptno"
    If False = adoSetOpen(StrSql, adoPtCnt) Then Return
    sCount = adoPtCnt.RecordCount
    Call adoSetClose(adoPtCnt)
    
    Return
    
    
End Sub

Private Sub Form_Activate()
    
    For i = 0 To 41
        If frmCalData.cmdDay(i).Tag = txtSysDate.Text Then
            frmCalData.cmdDay(i).SetFocus
            Call cmdDay_Click(i)
            Exit Sub
        End If
    Next

End Sub

Private Sub Form_Load()
    
    txtSysDate.Text = Dual_Date_Get("yyyy-MM-dd")
    
    GoSub Set_Year_Data
    GoSub Set_Month_Data
    GoSub ToYYYYMM_Set_Combo
    nLastDate = Select_LastDate(cmbYear, Left(cmbMonth, 2))
    nFirstWeek = Select_FirstWeek(cmbYear, Left(cmbMonth, 2))
    
    DoEvents: Call CalDisplay
    

    Exit Sub
    
'/----------------------------------------------------------------

Set_Year_Data:
    For i = 1901 To 2100
        cmbYear.AddItem Trim$(Str(i))
    Next
    
    Return


Set_Month_Data:
    For i = 1 To 12
        cmbMonth.AddItem Format(Trim$(Str(i)), "00") & " 월"
    Next
    Return
    
ToYYYYMM_Set_Combo:
    cmbYear.Text = Dual_Date_Get("yyyy")
    cmbMonth.Text = Dual_Date_Get("MM") & " 월"
    Return

    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub sprSLip_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim iSLipno1        As Integer
    
    If Row = 0 Then Exit Sub
    If Row > sprSLip.DataRowCnt Then Exit Sub
    
    
    sprSLip.Row = Row
    sprSLip.Col = 2: iSLipno1 = Val(sprSLip.Text)
    
    GoSub Set_Color
    Call SpreadSetClear(sprItem)
    GoSub Get_ITemData
    Exit Sub
    
Set_Color:
    sprSLip.Row = -1
    sprSLip.Col = -1
    sprSLip.ForeColor = RGB(0, 0, 0)
    
    sprSLip.Row = Row
    sprSLip.Row2 = Row
    sprSLip.Col = 2
    sprSLip.Col2 = sprSLip.MaxCols
    sprSLip.BlockMode = True
    sprSLip.ForeColor = RGB(0, 0, 255)
    sprSLip.BlockMode = False
    
    Return

Get_ITemData:
    StrSql = ""
    StrSql = StrSql & " SELECT  itemCd, ItemName,"
    StrSql = StrSql & "         SUM(DECODE(itemcd, '', '', 1)) Count"
    StrSql = StrSql & " FROM(   SELECT  DISTINCT a.itemcd, b.Routinnm ItemName, "
    StrSql = StrSql & "                 a.JeobsuDt, a.JeobsuT1, a.Jeobsut2"
    StrSql = StrSql & "         FROM    TWEXAM_Order   a,"
    StrSql = StrSql & "                 TWEXAM_Routine b"
    StrSql = StrSql & "         WHERE   a.COLLDate = TO_DATE('" & txtSysDate.Text & "','YYYY-MM-DD') "
    StrSql = StrSql & "         AND     a.JeobsuYN = '*'"
    StrSql = StrSql & "         AND     a.SLipno1  = " & iSLipno1
    StrSql = StrSql & "         AND     a.ItemCd   = b.RoutinCD"
    StrSql = StrSql & "         GROUP BY a.ItemCd, b.Routinnm, a.Jeobsudt, a.Jeobsut1, a.Jeobsut2)"
    StrSql = StrSql & " GROUP BY  itemcd, itemname"
    StrSql = StrSql & " UNION ALL"
    StrSql = StrSql & " SELECT a.itemcd , b.Itemnm ItemName, Count(*) Count"
    StrSql = StrSql & " FROM   TWEXAM_Order  a,"
    StrSql = StrSql & "        TWEXAM_ITEMML b"
    StrSql = StrSql & " WHERE  a.COLLDate = TO_DATE('" & txtSysDate.Text & "','YYYY-MM-DD') "
    StrSql = StrSql & " AND    a.JeobsuYN = '*'"
    StrSql = StrSql & " AND    a.SLipno1  = " & iSLipno1
    StrSql = StrSql & " AND    a.ItemCd   = b.Codeky"
    StrSql = StrSql & " GROUP BY  a.itemcd, b.Itemnm"
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprItem.Row = sprItem.DataRowCnt + 1
        sprItem.Col = 1: sprItem.Text = adoSet.Fields("itemCd").Value & ""
        sprItem.Col = 2: sprItem.Text = adoSet.Fields("itemName").Value & ""
        sprItem.Col = 3: sprItem.Text = adoSet.Fields("Count").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
    
End Sub

Private Sub sprSLip_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then Exit Sub
    If Row > sprSLip.DataRowCnt Then Exit Sub
    
    Call sprSLip_ButtonClicked(1, Row, 1)
    
End Sub
