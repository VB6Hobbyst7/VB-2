VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form INTface41 
   BorderStyle     =   0  '없음
   Caption         =   "해당일의 검사결과 받아 보기"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   1200
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7500
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin Threed.SSFrame fraWork 
      Height          =   6705
      Left            =   30
      TabIndex        =   10
      Top             =   60
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   11827
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdList 
         Height          =   5700
         Left            =   90
         TabIndex        =   7
         Top             =   870
         Width           =   11595
         _Version        =   196608
         _ExtentX        =   20452
         _ExtentY        =   10054
         _StockProps     =   64
         BackColorStyle  =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   7
         MaxRows         =   1
         ScrollBars      =   2
         SelectBlockOptions=   6
         SpreadDesigner  =   "INFACE41.frx":0000
         UserResize      =   0
         VisibleCols     =   3
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   5700
         Left            =   8400
         TabIndex        =   8
         Top             =   870
         Width           =   3285
         _Version        =   196608
         _ExtentX        =   5794
         _ExtentY        =   10054
         _StockProps     =   64
         BackColorStyle  =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   6
         MaxRows         =   21
         ScrollBars      =   2
         SpreadDesigner  =   "INFACE41.frx":1158
         UserResize      =   0
         VisibleCols     =   4
         VisibleRows     =   21
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   585
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   6855
         _Version        =   65536
         _ExtentX        =   12091
         _ExtentY        =   1032
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboSelect 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5220
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   165
            Width           =   1545
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   315
            Index           =   0
            Left            =   1170
            TabIndex        =   0
            Top             =   135
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   390
            Left            =   120
            TabIndex        =   13
            Top             =   90
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   688
            _StockProps     =   15
            Caption         =   "접수일자"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   390
            Left            =   4170
            TabIndex        =   14
            Top             =   120
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   688
            _StockProps     =   15
            Caption         =   "조회조건"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            RoundedCorners  =   0   'False
            MouseIcon       =   "INFACE41.frx":1609
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   315
            Index           =   1
            Left            =   2670
            TabIndex        =   1
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
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
         Begin VB.Label Label4 
            Caption         =   "-"
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
            Left            =   2520
            TabIndex        =   15
            Top             =   180
            Width           =   105
         End
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   450
         Left            =   8070
         TabIndex        =   4
         Top             =   270
         Visible         =   0   'False
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "선택"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "INFACE41.frx":1EE3
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   450
         Left            =   7140
         TabIndex        =   3
         Top             =   270
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "조회"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "INFACE41.frx":1EFF
      End
      Begin Threed.SSCommand cmdclose 
         Height          =   450
         Left            =   10830
         TabIndex        =   6
         Top             =   270
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "종 료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "INFACE41.frx":1F1B
      End
      Begin Threed.SSCommand cmdInitial 
         Height          =   450
         Left            =   9930
         TabIndex        =   5
         Top             =   270
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "취소"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "INFACE41.frx":1F37
      End
      Begin Threed.SSCommand cmdOrder 
         Height          =   960
         Left            =   3660
         TabIndex        =   18
         Top             =   -540
         Visible         =   0   'False
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1693
         _StockProps     =   78
         Caption         =   "전송"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
         Picture         =   "INFACE41.frx":1F53
      End
      Begin Threed.SSCommand cmdMake 
         Height          =   450
         Left            =   9000
         TabIndex        =   19
         Top             =   270
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "생성"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "INFACE41.frx":23A5
      End
      Begin VB.FileListBox FileBep2000 
         Height          =   990
         Left            =   6360
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   2085
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   450
         Left            =   9450
         TabIndex        =   17
         Top             =   60
         Visible         =   0   'False
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "삭 제"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "INFACE41.frx":23C1
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   450
         Left            =   5460
         TabIndex        =   16
         Top             =   30
         Visible         =   0   'False
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "열기"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "INFACE41.frx":23DD
      End
   End
   Begin FPSpread.vaSpread spdface 
      Height          =   5625
      Left            =   30
      TabIndex        =   9
      Top             =   1140
      Visible         =   0   'False
      Width           =   11775
      _Version        =   196608
      _ExtentX        =   20770
      _ExtentY        =   9922
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      BackColorStyle  =   1
      ColsFrozen      =   1
      EditEnterAction =   2
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   14
      SelectBlockOptions=   2
      SpreadDesigner  =   "INFACE41.frx":23F9
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   3390
      TabIndex        =   20
      Top             =   6810
      Width           =   8445
      _Version        =   65536
      _ExtentX        =   14896
      _ExtentY        =   1032
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtdeChoice 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7050
         TabIndex        =   26
         Top             =   150
         Width           =   795
      End
      Begin VB.TextBox txtChoice 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4110
         TabIndex        =   25
         Top             =   150
         Width           =   795
      End
      Begin VB.TextBox txtAll 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   24
         Top             =   150
         Width           =   795
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   2910
         TabIndex        =   21
         Top             =   120
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "선택검체수"
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   390
         Left            =   5490
         TabIndex        =   22
         Top             =   120
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "미선택검체수"
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   390
         Left            =   300
         TabIndex        =   23
         Top             =   120
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "전체검체수"
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "INTface41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private transcfg        As commset
Dim errfound            As Integer
Dim Porttag             As Integer
Dim RcvBuffer           As String
Dim TestNameTable(30)   As TestNameTbl
Dim nowcount            As Integer

Dim EmerCnt             As Integer
Dim SampleCnt           As Integer
Dim EditCnt             As Integer  'Express처럼 한 슬립이 두 프레임 이상 거쳐 나올 경우 사용
Dim PrevCnt             As Integer
Dim CurSampCnt          As Integer
Dim Startslip           As String

Dim StartBCol           As Integer
Dim EndBCol             As Integer
Dim StartBRow           As Integer
Dim EndBRow             As Integer
Dim identbOpenKey       As Integer
Dim IdList              As String
Dim BlockKey            As Integer
Dim Errkey              As Integer

'public에 slipno As String 선언
Dim phase               As Integer
Dim bufcnt              As Integer
Dim wkbuf               As String
Dim ix1                 As Integer
Dim PrevReq             As Integer
Dim OldRow              As Long
Dim long_slip1          As Long
Dim long_slip2          As Long

Dim Test_OpenFlag       As Integer
Dim F_iComm_Cnt         As Integer  '99.11.23 YEJ

Dim f_iWork_Row         As Integer
Dim f_adoCn             As ADODB.Connection

Private Type TYPE_TESTID
    sPatid      As String
    sPatnm      As String
    sOrderId    As String
    sRackno     As String
    sPosition   As String
    sTestID(1 To 100)   As String
    iTestCnt            As Integer
End Type
Dim f_tpTestList()  As TYPE_TESTID
Dim f_iTestCnt     As Integer

Dim AxOrder(5)     As String
Dim AxResult(13)   As String
Dim tmpOrder       As String
Dim fBEP2000(100) As String
Dim sBEP2000(10) As String
Dim mBEP2000(10) As String
 
Sub f_subGet_검사항목(ByVal sKeyno As String, _
                      ByRef sTestID() As String, ByRef iCnt As Integer)

    Dim adoRs   As New ADODB.Recordset
    Dim sqldoc  As String
    
    Dim labdate As String, numgbn As String, labsqno As String
    Dim iIdx    As Integer
    
    If InStr(sKeyno, "-") > 0 Then
        labdate = Mid$(sKeyno, 1, 8)
        numgbn = Mid$(sKeyno, 10, 1)
        labsqno = Mid$(sKeyno, 12, 5)
    Else
        labdate = Mid$(sKeyno, 1, 8)
        numgbn = Mid$(sKeyno, 9, 1)
        labsqno = Mid$(sKeyno, 10, 5)
    End If
    
    sqldoc = "select SLIPCD+ORDCD+SPCCD from LAB_DB..LAB030M" _
           & " where LABDATE = '" & labdate & "'" _
           & "   and NUMGBN  = '" & numgbn & "'" _
           & "   and LABSQNO = '" & labsqno & "'" _
           & "   and SUBCD   = ''" _
           & "   and SLIPCD + ORDCD + SPCCD in ("
           
    sqldoc = sqldoc + "''"
    
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")
    
    tbcode.MoveLast
    If tbcode.RecordCount > 0 Then tbcode.MoveFirst
    
    Do While Not tbcode.EOF
        sqldoc = sqldoc + ",'" + tbcode!code & "" & "'"
        tbcode.MoveNext
    Loop
    tbcode.Close:   dbcode.Close

    sqldoc = sqldoc + ")"
    
    adoRs.CursorLocation = adUseClient
    adoRs.Open sqldoc, f_adoCn, adOpenStatic, adLockReadOnly
    
    If adoRs.RecordCount > 0 Then adoRs.MoveFirst
    
    iCnt = 0
    Do While Not adoRs.EOF
    
        For iIdx = 1 To 30
            If InStr(TestNameTable(iIdx).code, Trim$(adoRs(0) & "")) > 0 Then
                iCnt = iCnt + 1
                sTestID(iCnt) = Trim(TestNameTable(iIdx).eqno)
            End If
        Next
        
        adoRs.MoveNext
    Loop
    adoRs.Close:    Set adoRs = Nothing
    
End Sub

'
'   참고치 판정
'
Private Function Chk_Ref(sOrdCd As String, sSubNo As String, sRes As String, _
                        sex As String) As String

    Dim sStr    As String
    Dim sData() As String
    Dim iRet_Cd As Integer
    
    Dim LowVal  As Single
    Dim HighVal As Single
    Dim RefVal  As Single
    Dim RefChar As String

    Chk_Ref = ""

    If sex = "M" Then
        sStr = " Select REFLOM, REFHIM, REFCHAR, REFCHK "
    ElseIf sex = "F" Then
        sStr = " Select REFLOF, REFHIF, REFCHAR, REFCHK "
    End If
    
    sStr = sStr & "  From LAB01_DB..DJA060M " _
            & " where ORDCD = '" & sOrdCd & "'" _
            & "   and SUBCD = '" & sSubNo & "'"
    
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 4, sStr, sData()
            
            If Trim(sData(4)) = "C" Then    '참고치 문자
                RefChar = Trim(sData(3))
                If RefChar <> Trim(sRes) Then
                    Chk_Ref = "*"
                End If
            ElseIf Trim(sData(4)) = "N" Then        '숫자
                RefVal = CSng(Val(Trim(sRes)))
                LowVal = CSng(sData(1)): HighVal = CSng(sData(2))
            
                If RefVal > HighVal Then
                    Chk_Ref = "H"
                ElseIf RefVal < LowVal Then
                    Chk_Ref = "L"
                End If
            End If
            
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
    
End Function


Sub f_subClear_Form()

    f_iWork_Row = 0
    
'    mskOrderno.Text = 1
'    mskRackno.Text = "A"
'    mskPosition.Text = 1
    
'    txtStatus.Text = ""
'    txtStatus.Visible = False
    fraWork.Visible = True
    
    spdList.MaxRows = 0
    
    With spdWorkList
        .MaxRows = 21
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = 21
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
    With spdface
        .MaxRows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = 14
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
End Sub

Private Sub Update_DJC020M(OrdSqNo As String)

    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
    Dim sqldoc  As String
    
    sStr = " Select count(*) from LAB03_DB..DJC050M " _
            & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
            & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
            & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
            & "   and RSTGBN = '' "
    
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 1, sStr, tData()
            
            If Val(tData(1)) = 0 Then
                sqldoc = " Update LAB03_DB..DJC020M " _
                        & "   set ORDSTAT = '1' " _
                        & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
                        & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
                        & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
                        & "   AND ORDSTAT IN ('0','6') "
                iRet_Cd = QSqlDBExec(sqldoc, QsqlCode)
            End If
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
                
End Sub


Private Function Append_To_Server(P_Key As String, iCnt As Integer, sOrdNo As String, RtnCd As String) As Integer
                
    Dim iRet    As Integer
    Dim sStr    As String
    Dim rData() As String
    Dim sLabNo    As String
    Dim II      As Integer
    Dim sqldoc  As String
    
    'sLabNo = Left(P_Key, 8) & Mid(P_Key, 10, 1) & Right(P_Key, 5)
    sLabNo = P_Key
    
    Append_To_Server = True
    
    '----- Server결과등록
    With Insert_Server(iCnt)
        '--- 검사 Order 내역 Table Update
        sStr = " Update LAB03_DB..DJC050M " _
                & "   set RSTGBN = 'Y' " _
                & " where ORDDATE = '" & Left(sOrdNo, 8) & "'" _
                & "   and DEPTCD = '" & Mid(sOrdNo, 9, 2) & "'" _
                & "   and SEQNO = '" & Right(sOrdNo, 5) & "'"
        If Mid(RtnCd, 4, 1) = "0" Then
            sStr = sStr & "   and ORDCD = '" & RtnCd & "'"
        Else
            sStr = sStr & "   and ORDCD = '" & .ordcd & "'"
        End If
        
        If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
            '--- Update
            sStr = " Update LAB04_DB..DJD010M " _
                    & "   set RSTVAL = '" & .Result & "', " _
                    & "       REFVAL = '" & .Ref & "', " _
                    & "       RSTID  = '" & D0COM_USERID & "', " _
                    & "       RSTDATE = '" & Format(Now, "YYYYMMDD") & "'" _
                    & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                    & "   and NUMGBN  = '" & Mid(sLabNo, 9, 1) & "'" _
                    & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                    & "   and ORDCD = '" & .ordcd & "'" _
                    & "   and SUBCD = '" & .SubNo & "'" _
                    & "   and ORDERNO = '" & sOrdNo & "'"
            
            If QSqlDBExec(sStr, QsqlConn) <> QSQL_SUCCESS Then
                '--- Insert(Sub 검사항목인 경우-조회 후 입력처리)
                ReDim INSDATA(1 To 5) As String
                
                '--- Insert할 항목 조회
                sqldoc = " Select DISTINCT REQGBN, SPCGBN, RETGBN, RTNCD, IDNO " _
                        & "  from LAB04_DB..DJD010M " _
                        & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                        & "   and NUMGBN = '" & Mid(sLabNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                        & "   and ORDCD = '" & .ordcd & "'" _
                        & "   and ORDERNO = '" & sOrdNo & "'"
                        
                If QSqlDBExec(sqldoc, QsqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                        QSqlGetField 5, sStr, rData()
                        
                        For II = 1 To 5
                            INSDATA(II) = Trim(rData(II))
                        Next II
                    Else
                        iRet = QSqlSelectFree(QsqlConn)
    '                    Append_To_Server = False
                        Exit Function
                    End If
                Else
                    iRet = QSqlSelectFree(QsqlConn)
    '                Append_To_Server = False
                    Exit Function
                End If
                iRet = QSqlSelectFree(QsqlConn)
                '--- 조회된 자료로 Insert처리
                sStr = " Insert into LAB04_DB..DJD010M ( " _
                        & " LABDATE, NUMGBN,  LABSQNO, ORDCD,  SUBCD, " _
                        & " RSTVAL,  REFVAL,  PANVAL,  DELVAL, REQGBN, " _
                        & " SPCGBN,  RETGBN,  RTNCD,   IDNO,   RSTID, " _
                        & " RSTDATE, ORDERNO, SYSDATE, SYSTIME ) values ( " _
                        & "'" & Left(sLabNo, 8) & "', " _
                        & "'" & Mid(sLabNo, 9, 1) & "', " _
                        & "'" & Right(sLabNo, 5) & "', " _
                        & "'" & .ordcd & "', " _
                        & "'" & .SubNo & "', " _
                        & "'" & .Result & "', " _
                        & "'" & .Ref & "', " _
                        & "'', '',"
                For II = 1 To 5
                    sStr = sStr & "'" & INSDATA(II) & "', "
                Next II
                sStr = sStr & "'" & D0COM_USERID & "', " _
                        & "'" & Format(Now, "YYYYMMDD") & "', " _
                        & "'" & sOrdNo & "', " _
                        & "'" & Format(Now, "YYYYMMDD") & "', " _
                        & "'" & Format(Now, "HHMMSS") & "') "
                        
                If QSqlDBExec(sStr, QsqlConn) <> QSQL_SUCCESS Then
                    Append_To_Server = False
                    Exit Function
                End If
                '---------------
            End If
        End If
        
        '--- 초기화
        .ordcd = ""
        .SubNo = ""
        .Result = ""
        .Ref = ""
        '-----------
    End With
    
End Function


Sub add_db_identb(sample As String, slip As String)
   
   If RecordExist(identb, "PrimaryKey", sample) Then
      identb.Edit
   Else
      identb.AddNew
      identb!seq_no = sample
   End If
   
   identb!slip_no = slip
   identb.Update
   identb.MoveLast

End Sub

Sub add_db_resulttb(sample2 As String, tcd As String, trt As String)
   
   tcd = Right$(tcd, 2)
   If RecordExists(resulttb, "PrimaryKey", sample2, tcd) Then
      resulttb.Edit
   Else
      resulttb.AddNew
      resulttb!seq_no = sample2
      resulttb!TestCode = tcd
   End If
   resulttb!TestResult = trt
   resulttb.Update
   resulttb.MoveLast

End Sub

Function FindIdList(position As Integer) As Integer
     
     FindIdList = InStr(position + 1, RcvBuffer, Chr(13))
     IdList = Mid$(RcvBuffer, FindIdList + 1, 1)

End Function

Sub PhaseCfg_Protocol()
    
    Dim wkdat          As String
    Dim ix1            As Integer
    Dim ir             As Integer
    
    Erase AxResult
    Erase AxOrder
    
    'wkbuf = txtStatus.Text
    tmpOrder = ""
'    SampleCnt = 0
    
    For ix1 = 1 To Len(wkbuf)
        wkdat = Mid$(wkbuf, ix1, 1)
        If Trim(wkdat) = "" Then Exit For
        Select Case Asc(wkdat)
        Case 5 'ENQ
        
        Case 21 'NAK
        
        Case 2  'STX
            phase = 2
            
            Select Case Mid$(wkbuf, ix1 + 2, 1)
'                Case "H" '-- Header
'                Case "P" '-- Patient
'                Case "C" '-- Comment
'                Case "Q" '-- Request
'                Case "L" '-- Terminator
                Case "O" '-- Order
                    'If InStr(wkbuf, AxOrder(1)) <> 0 Then
                    
                    For ir = 1 To 4
                        AxOrder(ir) = GetByOne(Mid(wkbuf, ix1, 100), wkbuf)
                    Next

                    For ir = 1 To 3
                        AxOrder(ir) = GetByOne1(AxOrder(4), AxOrder(4))
                    Next
                    
                    If tmpOrder = "" Then
                        tmpOrder = AxOrder(1)
                        SampleCnt = SampleCnt + 1
                    Else
                        If tmpOrder <> AxOrder(1) Then
                            tmpOrder = AxOrder(1)
                            SampleCnt = SampleCnt + 1
                        End If
                    End If
                    'Order    번호 : AxOrder(1) ex) 200109250002
                    'Lec      번호 : AxOrder(2) ex) E
                    'Lec별일련번호 : AxOrder(4) ex)
                Case "R" '-- Result
                    For ir = 1 To 12
                        AxResult(ir) = GetByOne(Mid(wkbuf, ix1, 100), wkbuf)
                    Next
                    '결과순서 : AxResult(2)
                    '검 사 명 : AxResult(3)
                    '결    과 : AxResult(4)
                    '단    위 : AxResult(5)
                    
                    '-- 2001.11.10 판정치가 있는경우 판정치를 넣어준다.
                    '-- 1|...........^^F|결과값 ==> 결과치
                    '-- 2|...........^^P|참고값 ==> 참고치
                    '-- 3|...........^^I|판  정 ==> 판정치(NONREACTIVE & NEGATIVE & POSITIVE)
                    If AxResult(2) = 1 And Right(Trim(AxResult(3)), 1) = "F" Then
                        AxResult(3) = Mid(AxResult(3), 4, Len(AxResult(3)) - 4)
                        For ir = 1 To 3
                            AxResult(ir) = GetByOne1(AxResult(3), AxResult(3))
                        Next
                        '검사명 : AxResult(2)
                        Call edit_data
                    
                    ElseIf AxResult(2) = 3 And Right(Trim(AxResult(3)), 1) = "I" Then
                            AxResult(3) = Mid(AxResult(3), 4, Len(AxResult(3)) - 4)
                            For ir = 1 To 3
                                AxResult(ir) = GetByOne1(AxResult(3), AxResult(3))
                            Next
                            '검사명 : AxResult(2)
                            
                            Call edit_data
                        
                        'End If
                    End If
                Case Else
                    'Exit Sub
            End Select
        Case Else
        
        End Select
                   
    Next
             
    'txtguide.Text = "Data 전송 완료!!"

End Sub

Private Function RecordExist(Tb As Recordset, IndexName As String, Samp As String) As Integer
         
         Dim CurrRecord As Variant

         If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
            RecordExist = False
            Exit Function
         End If

         '''CurrRecord = Tb.Bookmark
         Tb.MoveFirst
         Tb.Index = IndexName
         Tb.Seek "=", Samp

         If Tb.NoMatch Then
            '''Tb.Bookmark = CurrRecord
            RecordExist = False
         Else
            RecordExist = True
         End If
         
End Function
Private Function RecordExists(Tb As Recordset, IndexName As String, samp2 As String, tcd2 As String) As Integer
         Dim CurrRecord As Variant

         If Tb.RecordCount < 1 Then
            RecordExists = False
            Exit Function
         End If

         '''CurrRecord = Tb.Bookmark
         Tb.MoveFirst
         Tb.Index = IndexName
         Tb.Seek "=", samp2, tcd2

         If Tb.NoMatch Then
            '''Tb.Bookmark = CurrRecord
            RecordExists = False
         Else
             RecordExists = True
         End If

End Function
Sub edit_data()
        
    Dim seqno       As String
    Dim tresult(1 To 30) As String
    Dim tcode       As String
    Dim i           As Integer
    Dim a           As Integer
    Dim ix1         As Integer
    Dim pos         As Integer
    Dim tmpbuff     As String
    Dim NextPos     As Integer
    Dim tmpbuffer   As String
    Dim StartPos    As Integer
    Dim temp
    Dim no_tmp, no_tmp1, no_tmp2
    Dim iC          As Integer
    Dim chk_Id      As Boolean
    Dim tmp_iC      As Integer
    
'    SampleCnt = SampleCnt + 1
    
    '-- sampleNo.얻기->slipno에 해당
    Call spdface.GetText(1, SampleCnt, no_tmp)
    Call spdWorkList.GetText(3, SampleCnt, no_tmp1)
    Call spdWorkList.GetText(4, SampleCnt, no_tmp2)

    If Trim(no_tmp) = "" Then
'        MsgBox "Work List를 먼저 등록 하십시요", vbInformation, "Work List 등록"
        Exit Sub
    End If

'---검사결과값 얻기 ------------------------------------------------------------
    Erase tresult
    chk_Id = False
    
    For iC = 1 To spdWorkList.MaxRows
        Call spdWorkList.GetText(3, iC, no_tmp1)
        Call spdWorkList.GetText(4, iC, no_tmp2)
        
        If Trim(no_tmp1) = "" And Trim(no_tmp2) = "" Then Exit For
        
        If AxOrder(2) = no_tmp1 And AxOrder(4) = Format(no_tmp2, "00") Then
            tmp_iC = SampleCnt
            SampleCnt = iC
            chk_Id = True
            Exit For
        Else
'            If AxOrder(2) = no_tmp1 And AxOrder(4) = no_tmp2 Then
'                SampleCnt = iC
'                chk_Id = False
'            End If
        End If
    Next
    
    If chk_Id = False Then Exit Sub
    
    For ix1 = 1 To 30
        If Trim(TestNameTable(ix1).Name) = "" Then Exit For
        If InStr(AxResult(2), Trim(TestNameTable(ix1).Name)) <> 0 Then
            If Not IsNumeric(AxResult(4)) Then
                If UCase(Left(Trim(AxResult(4)), 1)) = "N" Then
                    tresult(ix1) = "음성"
                Else
                    tresult(ix1) = "양성"
                End If
            Else
                tresult(ix1) = AxResult(4)
            End If
            
            Exit For
        End If
    Next ix1

'---검사명, 검사결과값 spread에 뿌리기------------------------------------------------------------
'    txslipno.Text = slipno
'    txsapno.Text = Format(PrevCnt + SampleCnt, "0000")
'    txsapno.Text = Format(SampleCnt, "0000")
'    txtguide.Text = "TX Data!!"

    If SampleCnt = 1 Then
        spdface.Row = 1
    Else
'        Call Row_Plus(spdface)
        If SampleCnt >= spdface.MaxRows Then
            spdface.MaxRows = spdface.MaxRows + 1
            spdface.Row = SampleCnt
        Else
            spdface.Row = SampleCnt
        End If
    End If
'---검사명, 검사결과값 db에 등록 ------------------------------------------------------------
    tcode = ""
    For i = 1 To 30
        If Trim(TestNameTable(i).code) <> "" Then
            tcode = Format$(i, "00")
            If tcode <> "" Then
                add_db_identb Format(SampleCnt + PrevCnt, "0000"), CStr(no_tmp)
                If tresult(i) <> "" Then
                    Call spdface.GetText(1, SampleCnt, temp)
                
                    If temp <> "" Then
                        add_db_resulttb Format(PrevCnt + SampleCnt, "0000"), tcode, tresult(i)
                    End If
                    Call spdsettext(spdface, TestNameTable(i).col_cnt, SampleCnt, tresult(i))
                End If
            End If
        End If
    Next
    
    identbOpenKey = True   'DB에 등록이 되었으므로 결과 등록이 가능한 조건이 되었음을 나타내는 키
    SampleCnt = tmp_iC
    
End Sub


Sub edit_data1()
        
    Dim tresult     As String
    Dim tcode       As String
    Dim i           As Integer
    Dim sTmpbuff    As String
    Dim sTestID     As String, sTestVal As String
    
    Dim slip_no     As String   '접수일자+구분+작업번호
    Dim seq_no      As String   'Sample 순서
    Dim iPos1       As Integer, iPos2   As Integer, iPos3   As Integer
    Dim iRow        As Integer
    Dim vTmp        As Variant
    Dim sDate       As String, iEtc As Integer
    Dim j           As Integer
    Dim tmpAssay(7) As String
    
'    sDate = Mid$(mskOrdDate, 1, 4) + Mid$(mskOrdDate, 5, 2) + Mid$(mskOrdDate, 7, 2)
    
'임시Remark    If Not sDate = Mid$(AxResult(13), 10) Then Exit Sub
    
    seq_no = AxResult(2)
    
    iRow = D0SUB_SPREADGETROW(spdface, spdface.MaxCols, seq_no)
    If iRow < 1 Then Exit Sub
    
    spdface.GetText 1, iRow, vTmp:  slip_no = Trim$(vTmp)
    seq_no = Format$(nowcount + iRow, "0000")

'---검사명, 검사결과값 spread에 뿌리기------------------------------------------------------------
'    txsapno.Text = Format(seq_no, "0000")
'    txtguide.Text = "TX Data!!"
    
    Call add_db_identb(seq_no, slip_no)
    
    For j = 1 To 7
        tmpAssay(j) = GetByOne1(AxResult(3), AxResult(3))
    Next
    
    sTmpbuff = tmpAssay(5)

    Do While seq_no > 0
        sTestID = tmpAssay(4)
        sTestVal = AxResult(4) '& AxResult(5)
        
'        iEtc = 0
'        For i = 1 To spdface.MaxCols - 2
'            If Trim$(TestNameTable(i).eqno) = sTestID Then
'                iEtc = Val(TestNameTable(i).etc)
'                Exit For
'            End If
'        Next
'
'        sTestVal = Round(Val(sTestVal), iEtc)
        
        If Not sTestID = "" Then
            For i = 1 To spdface.MaxCols - 2
                If TestNameTable(i).eqno = sTestID Then
                    
                    If sTestVal <> "" Then
                        add_db_resulttb seq_no, sTestID, sTestVal
                        spdface.SetText TestNameTable(i).col_cnt, iRow, sTestVal
                    End If
                End If
            Next
        End If
        
        sTmpbuff = Mid$(sTmpbuff, iPos3 + 55)
        seq_no = Format$(nowcount + iRow, "0000")
    Loop
    
'    spdface.SetText spdface.MaxCols, iRow, seq_no
        
    F_iComm_Cnt = F_iComm_Cnt + 1
    
    identbOpenKey = True   'DB에 등록이 되었으므로 결과 등록이 가능한 조건이 되었음을 나타내는 키

End Sub
Sub Test()
    
    Dim rv%
    
    Test_OpenFlag = 1

    Open App.Path & "\dump_axsym.dat" For Input As #3
'    Open App.Path & "\axsym.log" For Input As #3
    Test_OpenFlag = 2
    wkbuf = ""
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

    Close #3

    Call PhaseCfg_Protocol
    
End Sub

Private Sub cboSelect_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub


Private Sub cmdAdd_Click()

    Dim iRow1   As Integer, iRow2   As Integer
    Dim vTmp    As Variant
    Dim sKeyno  As String, sTestID  As String, sPatnm   As String
    Dim iRack   As String, iPos    As Integer
    Dim sSendBuf    As String
    'Dim adoRs As ADODB.Recordset
    'Dim sqlDoc As String
    Dim icol As Integer
    
    For iRow1 = 1 To spdList.MaxRows
        spdList.GetText 3, iRow1, vTmp:  sPatnm = Trim$(vTmp)
        spdList.GetText 2, iRow1, vTmp:  sKeyno = Trim$(vTmp)
        spdList.GetText 1, iRow1, vTmp
        
'        spdList.Tag
        If Trim(vTmp) = "1" And Not sKeyno = "" Then
            With spdWorkList
                f_iWork_Row = f_iWork_Row + 1
                
                If f_iWork_Row > .MaxRows Then .MaxRows = .MaxRows + 1
                
                .SetText 1, f_iWork_Row, sKeyno
                .SetText 2, f_iWork_Row, sPatnm
                For icol = 3 To .MaxCols
                    If sBEP2000(icol - 2) <> "" Then .SetText icol, f_iWork_Row, "(    )"
                Next
                '-- 검사항목
                'Set adoRs = New ADODB.Recordset
                'sqlDoc = "select a.SLIPCD + a.ORDCD + a.SPCCD " _
                       & "  from LAB_DB..LAB030M a " _
                       & " where a.SLIPCD + a.ORDCD + a.SPCCD = '" & sBEP2000(f_iWork_Row) & "' " _
                       & "   and a.LABDATE = '" & Mid(sKeyno, 1, 8) & "'" _
                       & "   and a.NUMGBN = '" & Mid(sKeyno, 10, 1) & "'" _
                       & "   and a.LABSQNO = '" & Mid(sKeyno, 12, 5) & "'" _
                       & "   and a.SUBCD   = ''" _
                       & "   and (a.RSTVAL = '' or a.RSTVAL is null)"
                
                'adoRs.CursorLocation = adUseClient
                'adoRs.Open sqlDoc, f_adoCn, adOpenStatic, adLockReadOnly
            
                'If adoRs.RecordCount > 0 Then
                    'adoRs.MoveFirst
                'End If
                
                'Call add_db_identb(Format(f_iWork_Row, "0000"), Trim(sKeyno))
            End With
                    
        End If
    Next

    With spdList
        For iRow1 = 1 To .MaxRows
            If iRow1 > .MaxRows Then Exit For
            .GetText 1, iRow1, vTmp
            
            If Trim(vTmp) = "1" Then
                .Row = iRow1
                .Action = SS_ACTION_DELETE_ROW
                
                .MaxRows = .MaxRows - 1
                iRow1 = iRow1 - 1
            End If
        Next
    End With

End Sub

Private Sub cmdclose_Click()
    
    Unload Me
    FrmFlag = 0

End Sub

Private Sub cmdDelete_Click()

    Dim rv As Integer
    Dim i  As Integer
    Dim CurrentTbRows As Integer
    Dim ExistTxtKey As Integer
    Dim tmpSlip
    Dim seq_no      As String
    
    If StartBRow = -1 And EndBRow = -1 Then
        StartBRow = 1
        EndBRow = nowcount - PrevCnt
    End If
    
    For i = StartBRow To EndBRow
        rv = spdface.GetText(1, i, tmpSlip)
        If tmpSlip = "" Then
            ExistTxtKey = False
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If identbOpenKey = True And ExistTxtKey = True Then
        If StartBCol = -1 And EndBCol = -1 And BlockKey = True Then
            
            rv = MsgBox("블록으로 지정된 Slip을 삭제하시겠습니까?", 4, Title & "  " & "Slip No. 삭제 확인!!")
            If rv = 7 Then
                BlockKey = False
                spdface.EditMode = True
                spdface.EditMode = False
                cmdclose.SetFocus
                Exit Sub
            End If
            
            identb.Index = "primarykey"
            resulttb.Index = "Seq_No"
            
'            identb.MoveLast
'            CurrentTbRows = nowcount - PrevCnt
            
            For i = Val(StartBRow) To Val(EndBRow)
                With spdface
                    .Row = i
                    .Col = .MaxCols:    seq_no = .Text
                End With
                
                identb.Seek "=", seq_no
                If Not identb.NoMatch Then identb.Delete
                
                SampleCnt = SampleCnt - 1
                
                resulttb.Seek "=", seq_no
                If resulttb.NoMatch = False Then
                   Do Until resulttb.EOF
                       If resulttb!seq_no <> seq_no Then Exit Do
                       
                       resulttb.Delete
                            
                       resulttb.MoveNext
                   Loop
                End If
            Next
            
        '삭제하는 Spread 라인의 텍스트를 지움.
            spdface.BlockMode = True
            spdface.Col = -1
            spdface.Col2 = -1
            spdface.Row = StartBRow
            spdface.Row2 = EndBRow
            spdface.Action = SS_ACTION_DELETE_ROW
            spdface.BlockMode = False

        '1st Column(SlipNo)의 색깔을 노란색
            spdface.BlockMode = True
            spdface.Col = 1
            spdface.Col2 = 1
            spdface.Row = -1
            spdface.Row2 = -1
            spdface.BackColor = &HC0FFFF
            spdface.BlockMode = False
            
'            txsapno = ""
'            txtguide = "Data 삭제!!"
            
        Else
        
            MsgBox "잘못된 삭제 방법입니다." & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "삭제를 하십시요!!"
        
        End If
   Else
   
        MsgBox "데이터가 없거나 검사 결과 전송을 받지 않으셨습니다!!"
        
   End If

   BlockKey = False
   spdface.EditMode = True
   spdface.EditMode = False
   
'현재의 Row를 점검
   spdface.Row = nowcount - PrevCnt
   
End Sub

Private Sub cmdInitial_Click()
    
    Call f_subClear_Form
    
'    fraWork.Visible = True
'    FileBep2000.Path = "\\medcom\bep2000\Exportfiles"
    FileBep2000.Visible = False

End Sub

Private Sub cmdMake_Click()
    Dim iRow  As Integer
    Dim icol  As Integer
    Dim sTxt  As String
    Dim sName As String
    
    On Error GoTo ErrLst
    
    If spdList.MaxRows < 1 Then MsgBox "대상자가 없습니다", vbInformation, Me.Caption: Exit Sub
    
    Open ImportPath & "\" & Format(Now, "yymmdd") & "ipt.txt" For Output As #2
    
    sTxt = "Patient ID,"
    For icol = 4 To spdList.MaxCols
        If icol = spdList.MaxCols Then
            sTxt = sTxt + "Test name"
        Else
            sTxt = sTxt + "Test name,"
        End If
    Next
    
    Print #2, sTxt + Chr(13) & Chr(10);
    With spdList
        For iRow = 1 To .MaxRows
            sTxt = ""
            .Row = iRow: .Col = 1
            If .Value <> 1 Then GoTo rst2
            For icol = 2 To .MaxCols
                .Col = icol
                If icol = 3 Then GoTo rst1
                If icol < 3 Then
                    sTxt = sTxt + Trim(.Text) + ","
                Else
                    If Trim(.Text) <> "" Then
                        sTxt = sTxt + Trim(mBEP2000(icol - 3)) + ","
                    Else
                        sTxt = sTxt + ","
                    End If
                End If
rst1:
            Next icol
            If Right(sTxt, 1) = "," Then
                sTxt = Mid(sTxt, 1, Len(sTxt) - 1)
            End If
            
            Print #2, sTxt + Chr(13) & Chr(10);
rst2:
        Next iRow
    End With

    Close #2
    MsgBox "업무나열서가 생성되었습니다", vbInformation, Me.Caption
    
    Call cmdInitial_Click
    Exit Sub

ErrLst:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, Me.Caption
    
End Sub

Private Sub cmdOrder_Click()

'    Dim iRow1   As Integer, iRow2   As Integer
'    Dim iCnt    As Integer
'    Dim vTmp    As Variant
'    Dim sKeyno  As String, sTestID  As String, sPatnm   As String
'    Dim sRack   As String, sPos     As String
'    Dim sSendBuf    As String
'
'    iRow2 = 0
'    For iRow1 = 1 To spdWorkList.MaxRows
'        spdWorkList.GetText 1, iRow1, vTmp:  sKeyno = Trim$(vTmp)
'        spdWorkList.GetText 2, iRow1, vTmp:  sPatnm = Trim$(vTmp)
'        spdWorkList.GetText 3, iRow1, vTmp:  sRack = Trim$(vTmp)
'        spdWorkList.GetText 4, iRow1, vTmp:  sPos = Trim$(vTmp)
'
'        If sKeyno = "" Then Exit For
'
'        iRow2 = iRow2 + 1
'        ReDim Preserve f_tpTestList(1 To iRow2) As TYPE_TESTID
'
'        With f_tpTestList(iRow2)
'            Call f_subGet_검사항목(sKeyno, .sTestID, .iTestCnt)
'
'            .sPatnm = sPatnm
'            .sPatid = Mid$(sKeyno, 1, 8) + Mid$(sKeyno, 10, 1) + Mid$(sKeyno, 12, 5)
''            .sOrderId = "1" + Format(iRow2 + Val(mskOrderno.Text) - 1, "000") + Space(11) & " 00/00/0000"
'            .sOrderId = iRow2
'            .sRackno = "  " + CStr(sRack)
'            .sPosition = IIf(Val(sPos) > 9, "", " ") + CStr(sPos)
'
'        End With
'
'    Next
'
'    For iRow1 = 1 To iRow2
'
'        With f_tpTestList(iRow1)
'            For iCnt = 1 To .iTestCnt
''                sSendBuf = sSendBuf & "55 " & .sTestID(iCnt) & Chr(10)
'            Next
'
'            '-- 결과 보기
'            If iRow1 > spdface.MaxRows Then spdface.MaxRows = spdface.MaxRows + 1
'
'            spdface.SetText 1, iRow1, Mid(.sPatid, 1, 8) + "-" + Mid$(.sPatid, 9, 1) + "-" + Mid$(.sPatid, 10, 5)
'            spdface.SetText 2, iRow1, .sPatnm
'            spdface.SetText spdface.MaxCols, iRow1, Mid$(.sOrderId, 1, 4)
'
'        End With
'    Next
'
'    'Call Test
'    Comm1.Output = Chr$(5) 'ENQ
'    txtguide.Text = "Data 수신대기"
'
'    CurSampCnt = iRow2
'
'    If iCnt = 0 Then
'        MsgBox "Work List에 등록할 자료가 없습니다. 조회를 실행해 주십시오.", vbCritical, Me.Caption
'    Else
'        fraWork.Visible = False
'
'        txtStatus.Text = ""
'        txtStatus.Visible = True
'        Timer1.Enabled = True
'    End If
'
End Sub

Private Sub cmdQuery_Click()

    Dim adoRs   As New ADODB.Recordset
    Dim adoRs1  As New ADODB.Recordset
    Dim sqldoc  As String
    
    Dim iRow    As Integer
    Dim vTmp    As Variant
    
    Dim iRet    As Integer
    Dim sStr    As String
    Dim tData() As String
    Dim icol    As Integer
    
    '--- 조회조건 체크
    If Not IsDate(Format(mskDate(0), "####-##-##")) Then
        MsgBox "조회를 원하는 접수일자를 입력해 주십시오.", vbExclamation
        mskDate(0).SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Format(mskDate(1), "####-##-##")) Then
        MsgBox "조회를 원하는 접수일자를 입력해 주십시오.", vbExclamation
        mskDate(1).SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    spdList.MaxRows = 0
    
    sqldoc = "select distinct" _
           & "       a.LABDATE,  a.NUMGBN, a.LABSQNO, b.PATNM" _
           & "  from LAB_DB..LAB030M a, PAT_DB..PAT010M b" _
           & " where a.LABDATE between '" & mskDate(0).Text & "' and '" & mskDate(1).Text & "'" _
           & "   and a.SUBCD   = ''" _
           & "   and a.SLIPCD + a.ORDCD + a.SPCCD in ("
    
    sqldoc = sqldoc + "''"
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")

    tbcode.MoveLast
    If tbcode.RecordCount > 0 Then tbcode.MoveFirst
    
    Do While Not tbcode.EOF
        sqldoc = sqldoc & ",'" + tbcode!code & "" & "'"
        tbcode.MoveNext
    Loop
    tbcode.Close:   dbcode.Close
    sqldoc = sqldoc & ")"
    
    If cboSelect.ListIndex = 0 Then
        sqldoc = sqldoc & "   and (a.RSTVAL = '' or a.RSTVAL is null) "
    Else
        sqldoc = sqldoc & "   and not (a.RSTVAL = '' or a.RSTVAL is null)"
    End If
    sqldoc = sqldoc _
           & "   and a.IDNO = b.IDNO " _
           & " order by a.LABDATE, a.NUMGBN, a.LABSQNO "
                
    adoRs.CursorLocation = adUseClient
    adoRs.Open sqldoc, f_adoCn, adOpenStatic, adLockReadOnly
    
    If adoRs.RecordCount > 0 Then
        adoRs.MoveFirst
        txtAll.Text = adoRs.RecordCount
        cmdMake.Enabled = True
    Else
        cmdMake.Enabled = False
    End If
    
    iRow = 0
    Do While Not adoRs.EOF
        iRow = iRow + 1
        
        With spdList
            If iRow > .MaxRows Then .MaxRows = .MaxRows + 1
            .SetText 1, iRow, "1"
            .SetText 2, iRow, adoRs(0) & "" + "-" + adoRs(1) & "" + "-" & adoRs(2) & ""
            .SetText 3, iRow, Trim$(adoRs(3) & "")
            sqldoc = "select a.SLIPCD + a.ORDCD + a.SPCCD as wCode" _
                   & "  from LAB_DB..LAB030M a " _
                   & " where a.SUBCD   = ''" _
                   & "   and a.LABDATE = '" & adoRs(0) & "'" _
                   & "   and a.NUMGBN = '" & adoRs(1) & "'" _
                   & "   and a.LABSQNO = '" & adoRs(2) & "'" _
                   & "   and a.SLIPCD + a.ORDCD + a.SPCCD in ("
            
            sqldoc = sqldoc + "''"
            
            For icol = 1 To .MaxCols - 3
                sqldoc = sqldoc & ",'" + sBEP2000(icol) & "" & "'"
            Next
            
            sqldoc = sqldoc & ")"
            
            If cboSelect.ListIndex = 0 Then
                sqldoc = sqldoc & "   and (a.RSTVAL = '' or a.RSTVAL is null) "
            Else
                sqldoc = sqldoc & "   and not (a.RSTVAL = '' or a.RSTVAL is null)"
            End If
                        
            adoRs1.CursorLocation = adUseClient
            adoRs1.Open sqldoc, f_adoCn, adOpenStatic, adLockReadOnly
            
            If adoRs1.RecordCount > 0 Then adoRs1.MoveFirst
            
            For icol = 1 To .MaxCols - 3
                If adoRs1.EOF Then Exit For
                If Trim(sBEP2000(icol)) = "" & Trim(adoRs1(0)) Then
                    .SetText icol + 3, iRow, "      (          )"
                    adoRs1.MoveNext
                End If
            Next
        
        End With
        adoRs1.Close
        adoRs.MoveNext
    Loop
    adoRs.Close:    Set adoRs = Nothing
    
    If spdList.MaxRows = 0 Then _
        MsgBox "해당자료가 존재하지 않습니다.", vbInformation, Me.Caption
    
    Me.MousePointer = 0
    
End Sub


'Private Sub Comm1_OnComm()
'    Dim wkdat   As String
'    Dim pnlDump As String
'
'    Screen.MousePointer = 11
'
'    Select Case Comm1.CommEvent
'       ' Events
'        Case MSCOMM_EV_SEND     ' There are SThreshold number of
'                               ' character in the transmit buffer.
'        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
'
'            'txtguide.Text = "Data 수신준비"
'            wkbuf = Comm1.Input
'            Select Case Asc(wkbuf)
'                Case 3  '-- ETX
''                    Timer1.Enabled = False
''                    Timer2.Enabled = False
'                Case 21 '-- NAK
'
'                    Timer1.Enabled = True
'                    Timer1.Interval = 12000
'                    Timer2.Enabled = False
'                Case 4  '-- EOT
'                    Timer1.Enabled = False
'                    Timer2.Enabled = False
''                    Call Test
'                    Call PhaseCfg_Protocol
'                    txtguide.Text = "Data 전송 완료!!"
'                    txtStatus.Text = ""
'                    wkbuf = ""
'                Case 5  '-- ENQ
'                    txtguide.Text = "Data 수신 중!!"
'                    Timer2.Enabled = True
'                    Timer2.Interval = 1000
'                    Timer1.Enabled = False
''                    txtguide.Text = "Data 수신 중!!"
'                    DoEvents
'                Case Else
'
'            End Select
'
'            txtStatus.Text = txtStatus.Text + wkbuf
'
''            Print #1, wkbuf;    'Test
''            wkbuf = wkbuf + wkbuf
'
'
'        Case MSCOMM_EV_CTS      'j
'        Case MSCOMM_EV_DSR      ' Change in the DSR line.
'        Case MSCOMM_EV_CD       ' Change in the CD line.
'        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
'        ' Errors
'        Case MSCOMM_ER_BREAK    ' A Break was received.
'        ' Code to handle a BREAK goes here, and so on.
'        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
'        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
'        Case MSCOMM_ER_FRAME    ' Framing Error.
'        Case MSCOMM_ER_OVERRUN  ' Data Lost.
'        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
'        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
'        Case MSCOMM_ER_RXPARITY ' Parity Error.
'        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
'    End Select
'
'    Screen.MousePointer = 0
'
'End Sub
'Private Sub Command1_Click()
'
'    Call Test
'
'End Sub

Private Sub FileBep2000_DblClick()
    Dim wkbuf
    
    Test_OpenFlag = 1

    Open "\\medcom\bep2000\exportfiles\" & FileBep2000.filename For Input As #3
    
    Test_OpenFlag = 2
    wkbuf = ""
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

'    Debug.Print wkbuf
    Close #3
    
    Call psDataDefine(wkbuf)

    
End Sub

Private Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = Left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Private Function Text_Change(FSend_Str As String, FCheck_Char As String, FChange_Char As String) As String
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

Private Sub psDataDefine(ByVal sRstText As String)

Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
Dim Channel_No  As String       ' 문자형 변수
Dim Patiant_No  As String       ' 환자번호
Dim pGrid_Point As Integer      ' 해당 검사자 Point
Dim Max_Arary_Cnt As Integer    ' 검사 항목수
'-------------------------------' 임시 변수들.....
Dim sDeCnt      As Integer
Dim pDoCount    As Integer
Dim Loop_Count  As Integer
Dim FunStr1 As String, FunStr2 As String, FunStr3 As String, FunStr4 As String
Dim sRtn As Integer, sChannel As String, sRstValue As Single, sUnit As String
Dim sPoint1 As Integer
Dim sPoint2 As Integer
Dim sLname As String
Dim fmatVal

'    sRstText = brbarcd
    '------------------------------<<< fUrinscan300() 배열 Clear 한다.         >>>----------
    For Loop_Count = 1 To 100: fBEP2000(Loop_Count) = "": Next Loop_Count
    '------------------------------<<< fUrinscan300() 배열에 구분하여 넣는다.  >>>----------

    pDoCount = 0
    Do While InStr(sRstText, Chr$(13)) > 0
        'pDoCount = pDoCount + 1
        If pDoCount = 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "[") > 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "]") > 0 Then
            '-- pDoCount = 0 : 검사명
            fBEP2000(pDoCount) = Text_Redefine(sRstText, Chr$(13))
            fBEP2000(pDoCount) = Mid(fBEP2000(pDoCount), InStr(Text_Redefine(sRstText, Chr$(13)), "[") + 1, InStr(Text_Redefine(sRstText, Chr$(13)), "]") - 2)
                Debug.Print fBEP2000(pDoCount)
            pDoCount = pDoCount + 1
        ElseIf pDoCount = 1 And InStr(Text_Redefine(sRstText, Chr$(13)), "[") > 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "]") > 0 Then
            fBEP2000(pDoCount) = Text_Redefine(sRstText, Chr$(13))
                Debug.Print fBEP2000(pDoCount)
            pDoCount = pDoCount + 1
        ElseIf pDoCount = 2 Then
            '-- pDoCount = 2 : Header
            fBEP2000(pDoCount) = Text_Redefine(sRstText, Chr$(13))
                Debug.Print fBEP2000(pDoCount)
            pDoCount = pDoCount + 1
        ElseIf pDoCount > 2 Then
            '-- pDoCount > 2 : Result
            If Mid(Trim(Text_Redefine(sRstText, Chr$(13))), 2, 1) <> """" And InStr(Text_Redefine(sRstText, Chr$(13)), "[") = 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "]") = 0 Then
                fBEP2000(pDoCount) = Text_Redefine(sRstText, Chr$(13))
                Debug.Print fBEP2000(pDoCount)
                pDoCount = pDoCount + 1
            End If
        End If
        
        sRstText = Mid$(sRstText, InStr(sRstText, Chr$(13)) + 2)
        If pDoCount > 99 Then
            sRstText = ""
            Exit Do
        End If
    Loop
             
    'fmat(0) = "Patient ID;Assay;Raw value;Qualitative data;Well Location;Flag"
    'fmat(1) = "Patient ID;Assay;Raw value;Evaluation;Well Location;Flag"
    'fmat(2) = "Patient ID;Assay;Raw value;Corrected OD;Evaluation;IU/l;Well Location;Flag"
    'fmat(3) = "Patient ID;Assay;Reader value;Qual. value;Well Location"

    Select Case fmatVal
    Case "Raw value"
    
    Case "Reader value"
    
    Case "Qual. value"
    
    Case "Evaluation"
    
    Case "Qualitative data"
    
    Case "Corrected OD"
    
    Case "IU/ml"
    
    Case "Titer"
    
    Case "mIU/ml"
    
    Case "IU/l"
    
    Case Else
    
    End Select
    
             
'    Max_Arary_Cnt = brSpread.MaxCols - 6   ' 앞에서부터 5까지는 환자 정보 이기때문에.... -5를 한다.
'                                            '해당 배열은  brItem(),brChannel() 이다.
'    pGrid_Point = 0
'
'    With brSpread
'        For pDoCount = 1 To .MaxRows
'            .Row = pDoCount: .Col = 16
'            '--------------------------------------------------------------------------
'            '-- Modify Date : 2003/04/30
'            '-- Modify Dev  : sw,oh
'            '-- Modify 내용 : 바코드번호를 가지고 있을 경우 그 번호의 대상자를 찾는다.
'            '--------------------------------------------------------------------------
'            If Len(fUrinscan300(2)) > 20 And optBar.Value = True Then
'                .Col = 2
'                If Trim(brSpread.Text) = Trim(Mid$(fUrinscan300(2), 12, 10)) Then             ' fMidtron(0) = 환자 ID가 Key
'                    pGrid_Point = pDoCount
'                    Exit For
'                End If
'            ElseIf optSeq.Value = True Then
'                .Col = 0
'                If Val((brSpread.StartingRowNumber + pDoCount) - 1) = Val(Trim(Mid$(fUrinscan300(2), 7, 4))) Then             ' fMidtron(0) = 환자 ID가 Key
'                    pGrid_Point = pDoCount
'                    Exit For
'                End If
'            End If
'
'        Next pDoCount
'        If pGrid_Point > 0 Then                                   ' 해당 대상자를 찿으면 ....
'        '----------------------------------------------<<<<<<<<<,  세부검사항목을 찿는다.  >>>>>>>----------
'            For sDeCnt = 1 To Max_Arary_Cnt
'                For pDoCount = 1 To 24
'
'                    .Row = pGrid_Point
'                    .Col = sDeCnt + 6
'                    Channel_No = Trim(Mid$(brChannel(sDeCnt), 1, 4))         'Val(brChannel(sDeCnt))               '  Channel이 숫자이기 때문에 숫자로 치환한다.
'                    If Channel_No = Trim(Mid$(fUrinscan300(pDoCount), 1, 3)) Then
'                        If Len(Trim(Mid$(fUrinscan300(pDoCount), 4, 5))) > 0 Then
'                            If Trim(Mid$(fUrinscan300(pDoCount), 4, 7)) = "+-" Then
'                                .Text = "   +/-"
'                            Else
'                                '-- 2003/03/04 수정(by osw)
'                                If Trim(InStr(fUrinscan300(pDoCount), "+++++")) Then
'                                    .Text = Trim(Mid$(fUrinscan300(pDoCount), 4, 9))
'                                ElseIf Trim(InStr(fUrinscan300(pDoCount), "++++")) Then
'                                    .Text = Trim(Mid$(fUrinscan300(pDoCount), 4, 8))
'                                ElseIf Trim(InStr(fUrinscan300(pDoCount), "+++")) Then
'                                    .Text = Trim(Mid$(fUrinscan300(pDoCount), 4, 7))
'                                Else
'                                    .Text = Trim(Mid$(fUrinscan300(pDoCount), 4, 7))
'                                End If
'                            End If
'                        Else
'                            .Text = Trim(Mid$(fUrinscan300(pDoCount), 11, 8))
'                        End If
'                        Exit For
'                    End If
'                Next pDoCount
'            Next sDeCnt
'        End If
'    End With
    
End Sub


Private Sub Form_Load()
    Dim icol    As Integer
    
    'form을 가운데에 위치
    Me.Top = 0
    Me.Left = 0
    Me.Height = INTmain00.Height - INTmain00.pnlMain.Height - 500
    Me.Width = INTmain00.Width - 200
    
    fraWork.ZOrder 0
    
    Dim tablerows As Integer
    Dim iRow As Integer
    Dim i As Integer
    Dim TestItemNo As Integer
    
    Set f_adoCn = New ADODB.Connection
    f_adoCn.Open p_adoCnStr_1
    
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")
    
    tbcode.MoveLast
    tablerows = tbcode.RecordCount
        
    tbcode.MoveFirst
   
    iRow = 0
    Do While Not tbcode.EOF
        
        iRow = iRow + 1
        
        TestNameTable(iRow).eqno = tbcode!EQIPNO & ""
        TestNameTable(iRow).code = tbcode!code & ""
        TestNameTable(iRow).Name = tbcode!Name & ""
        TestNameTable(iRow).Mname = tbcode!Mname & ""
        
        If TestNameTable(iRow).code <> "" Then
            
            TestItemNo = TestItemNo + 1
            TestNameTable(iRow).col_cnt = TestItemNo + 2
            spdface.MaxCols = TestNameTable(iRow).col_cnt
            
            '-- 2001.11.12추가
            spdface.Font = "굴림체"
            spdface.FontSize = 9
            
            For icol = 3 To spdface.MaxCols
                spdface.ColWidth(icol) = 7
            Next
            
            Call spdsettext(spdface, TestNameTable(iRow).col_cnt, 0, TestNameTable(iRow).Name)
            sBEP2000(iRow) = TestNameTable(iRow).code
            mBEP2000(iRow) = Trim(TestNameTable(iRow).Mname)
            
        End If
        
        tbcode.MoveNext
    Loop

    With spdface
        .MaxCols = .MaxCols + 1
        .Col = .MaxCols
        .ColHidden = True
    End With

    SampleCnt = 0
    F_iComm_Cnt = 1 '99.11.24 YEJ 추가
    
    mskDate(0).Text = Format(Now, "yyyymmdd")
    mskDate(1).Text = Format(Now, "yyyymmdd")
    
    With cboSelect
        .AddItem ("미등록 자료")
        .AddItem ("등록된 자료")
        .ListIndex = 0
    End With
    
    tbcode.Close
    dbcode.Close
    
'--------------츄가------------
    'Call spdsettext(spdWorkList, 3, 0, "Sex")
    'Call spdsettext(spdWorkList, 4, 0, "OrdNo")
    'Call spdsettext(spdWorkList, 5, 0, "RtnCd")
    'Call spdsettext(spdWorkList, 6, 0, "SampleNo")
    
    'txtmmdd.Text = Format(month(Now), "00") & Format(day(Now), "00")
           
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")

    tbcode.MoveFirst

    iRow = 0: TestItemNo = 0
    Do While Not tbcode.EOF
        iRow = iRow + 1
        TestNameTable(iRow).eqno = tbcode!EQIPNO & ""
        TestNameTable(iRow).code = tbcode!code & ""
        TestNameTable(iRow).Name = tbcode!Name & ""

        If TestNameTable(iRow).code <> "" Then
            TestItemNo = TestItemNo + 1
            TestNameTable(iRow).col_cnt = TestItemNo + 3
            spdList.MaxCols = TestNameTable(iRow).col_cnt

            Call spdsettext(spdList, TestNameTable(iRow).col_cnt, 0, TestNameTable(iRow).Name)
            spdList.ColWidth(TestNameTable(iRow).col_cnt) = 15
        End If

        tbcode.MoveNext
    Loop

    tbcode.Close:   dbcode.Close

'------------------------------
    
    
    
'1st Column(SlipNo)의 색깔을 노란색
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.Col2 = 1
    spdface.Row = -1
    spdface.Row2 = -1
    spdface.BackColor = &HC0FFFF
    spdface.BlockMode = False

'Interface Result를 일단 Lock
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.Col2 = spdface.MaxCols
    spdface.Row = 1
    spdface.Row2 = spdface.MaxRows
    spdface.Lock = True
    spdface.BlockMode = False

'Spread.Row Initialization
    spdface.Row = 0
    
'    tbcode.Close
'    dbcode.Close
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Porttag = 3 Then
        'Comm1.PortOpen = False
        identb.Close
        resulttb.Close
        Db.Close
        identbOpenKey = False
        Close #1
        Close #2
    End If
    
    If Not f_adoCn.State = adStateClosed Then
        f_adoCn.Close:  Set f_adoCn = Nothing
    End If
    
End Sub


Private Sub mskDate_GotFocus(Index As Integer)
    
    With mskDate(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With

End Sub

Private Sub mskDate_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    ElseIf Not KeyAscii = vbKeyBack Then
        mskDate(Index).SelLength = 1
    End If
    
End Sub

'Private Sub mskOrderno_GotFocus()
'
'    mskOrderno.SelStart = 0
'    mskOrderno.SelLength = Len(mskOrderno.Text)
'
'End Sub

'Private Sub mskPosition_GotFocus()
'
'    mskPosition.SelStart = 0
'    mskPosition.SelLength = Len(mskPosition.Text)
'
'End Sub
'
'Private Sub mskPosition_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'        KeyAscii = 0
'    End If
'
'End Sub


'Private Sub mskRackno_GotFocus()
'
'    mskRackno.SelStart = 0
'    mskRackno.SelLength = Len(mskRackno.Text)
'
'End Sub
'
'Private Sub mskRackno_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'        KeyAscii = 0
'    End If
'
'End Sub


Private Sub spdface_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    StartBCol = CInt(BlockCol)
    StartBRow = CInt(BlockRow)
    EndBCol = CInt(BlockCol2)
    EndBRow = CInt(BlockRow2)
    BlockKey = True
End Sub

Private Sub spdList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim sNo As Long
    Dim eNo As Long
    Dim iRow    As Integer
    Dim Tmp As Variant
    
    If BlockRow = 0 Or BlockRow2 = 0 Or BlockRow = BlockRow2 Then Exit Sub
    
    If BlockRow < BlockRow2 Then
        sNo = BlockRow: eNo = BlockRow2
    Else
        sNo = BlockRow2: eNo = BlockRow
    End If

    For iRow = sNo To eNo
        With spdList
            Call .GetText(1, iRow, Tmp)
            If Tmp = True Then
                Call .SetText(1, iRow, "0")
            Else
                Call .SetText(1, iRow, "1")
            End If
        End With
    Next iRow
    
    spdList.SelModeSelected = False
    txtChoice.Text = "0"
    txtdeChoice.Text = "0"
    spdList.Col = 1
    For iRow = 1 To spdList.MaxRows
        spdList.Row = iRow
        If spdList.Value = "1" Then
            txtChoice.Text = Val(txtChoice.Text) + 1
        End If
    Next
    txtdeChoice.Text = Val(txtAll.Text) - Val(txtChoice.Text)
cmdQuery.SetFocus
End Sub


Private Sub spdList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim iRow As Integer
    
    txtChoice.Text = "0"
    txtdeChoice.Text = "0"
    spdList.Col = 1
    For iRow = 1 To spdList.MaxRows
        spdList.Row = iRow
        If spdList.Value = "1" Then
            txtChoice.Text = Val(txtChoice.Text) + 1
        End If
    Next
    txtdeChoice.Text = Val(txtAll.Text) - Val(txtChoice.Text)

End Sub


Private Sub spdList_GotFocus()
    With spdList
        If OldRow <> 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = &H80000005
        End If
        If .ActiveRow = 1 Then
            .Row = .ActiveRow
            .Col = -1
    
            If .Lock = False Then
                .BackColor = &HEEEFFF
                OldRow = .ActiveRow
            End If
        End If
    End With
End Sub

Private Sub spdList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
      If Row <> NewRow Then
        If OldRow = 0 Then OldRow = Row
        
        With spdList
            If NewRow <> -1 Then
                .Row = OldRow
                .Col = -1
                .BackColor = &H80000005
                
                .Row = NewRow
                .Col = -1

                .BackColor = &HEEEFFF
                OldRow = NewRow
                'Call Disp_Data(OldRow)     '해당 Sample의 Order표시
            End If
        End With
    End If
End Sub

Private Sub cmdDown_Click()

    FileBep2000.Path = "\\medcom\bep2000\Exportfiles"
    FileBep2000.Visible = True
    
End Sub

'Private Sub Timer1_Timer()
'
'    Comm1.Output = Chr$(5) 'ENQ
'
'End Sub
'
'
'Private Sub Timer2_Timer()
'
'    Comm1.Output = Chr$(6) 'ACK
'
'End Sub

