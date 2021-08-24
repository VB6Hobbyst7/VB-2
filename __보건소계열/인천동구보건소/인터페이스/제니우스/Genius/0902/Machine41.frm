VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form INTface41 
   BorderStyle     =   0  '없음
   Caption         =   "해당일의 검사결과 받아 보기"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   1200
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7500
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdface 
      Height          =   6030
      Left            =   4020
      OleObjectBlob   =   "Machine41.frx":0000
      TabIndex        =   2
      Top             =   1080
      Width           =   7740
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3825
      Top             =   7200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2970
      Top             =   7065
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Command1"
      Height          =   330
      Left            =   11415
      TabIndex        =   32
      Top             =   855
      Visible         =   0   'False
      Width           =   420
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1035
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1826
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "검사 인터페이스 작업을 수행합니다!!"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   675
         Width           =   3675
      End
      Begin VB.Label LblMMDD 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   7
         Top             =   225
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   870
      Left            =   10080
      TabIndex        =   21
      Top             =   135
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "삭   제"
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
      Picture         =   "Machine41.frx":03D6
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   45
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   512
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin Threed.SSCommand cmdclose 
      Height          =   870
      Left            =   10905
      TabIndex        =   1
      Top             =   135
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "종   료"
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
      Picture         =   "Machine41.frx":1D78
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   645
      Left            =   7770
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5850
      Width           =   2160
      _Version        =   65536
      _ExtentX        =   3810
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "서버에       등록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "Machine41.frx":371A
   End
   Begin Threed.SSFrame fmeGuide 
      Height          =   585
      Left            =   6870
      TabIndex        =   19
      Top             =   225
      Width           =   2325
      _Version        =   65536
      _ExtentX        =   4101
      _ExtentY        =   1032
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
      Begin VB.TextBox txtguide 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   170
         Width           =   2085
      End
   End
   Begin Threed.SSFrame fmeSlipSeq 
      Height          =   585
      Left            =   4305
      TabIndex        =   3
      Top             =   225
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   1032
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
      Begin VB.TextBox txsapno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1170
         TabIndex        =   4
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "일련번호"
         Height          =   210
         Left            =   210
         TabIndex        =   5
         Top             =   240
         Width           =   870
      End
   End
   Begin Threed.SSFrame fmeOrder 
      Height          =   5460
      Left            =   5175
      TabIndex        =   10
      Top             =   1350
      Width           =   6360
      _Version        =   65536
      _ExtentX        =   11218
      _ExtentY        =   9631
      _StockProps     =   14
      Caption         =   "WORKLIST  작성"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdWorklist1 
         Height          =   4590
         Left            =   2760
         OleObjectBlob   =   "Machine41.frx":3FF4
         TabIndex        =   11
         Top             =   420
         Width           =   3195
      End
      Begin FPSpread.vaSpread spdWorklist2 
         Height          =   4590
         Left            =   6630
         OleObjectBlob   =   "Machine41.frx":119AE
         TabIndex        =   18
         Top             =   420
         Width           =   3165
      End
      Begin VB.TextBox txtSId 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   14
         Top             =   1035
         Width           =   1110
      End
      Begin VB.TextBox txtPos 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1290
         MaxLength       =   2
         TabIndex        =   13
         Top             =   765
         Width           =   825
      End
      Begin VB.TextBox txtRackNo 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   12
         Top             =   495
         Width           =   825
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   17
         Top             =   795
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rack No."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   16
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "SampleID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   15
         Top             =   1065
         Width           =   1005
      End
   End
   Begin Threed.SSFrame sFrame1 
      Height          =   6060
      Left            =   0
      TabIndex        =   22
      Top             =   1050
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   10689
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
         Height          =   4140
         Left            =   30
         OleObjectBlob   =   "Machine41.frx":1F368
         TabIndex        =   23
         Top             =   1260
         Width           =   3945
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1035
         Left            =   30
         TabIndex        =   24
         Top             =   165
         Width           =   3945
         _Version        =   65536
         _ExtentX        =   6959
         _ExtentY        =   1826
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
         BorderWidth     =   2
         BevelInner      =   1
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
            Left            =   1320
            Style           =   2  '드롭다운 목록
            TabIndex        =   25
            Top             =   555
            Width           =   1605
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   345
            Left            =   1320
            TabIndex        =   26
            Top             =   135
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
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
            Left            =   60
            TabIndex        =   27
            Top             =   120
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
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
            Left            =   60
            TabIndex        =   28
            Top             =   510
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
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
            MouseIcon       =   "Machine41.frx":20440
         End
         Begin Threed.SSCommand cmdQuery 
            Height          =   870
            Left            =   2970
            TabIndex        =   29
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   1535
            _StockProps     =   78
            Caption         =   "접수조회"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   2
            RoundedCorners  =   0   'False
            Picture         =   "Machine41.frx":20D1A
         End
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   495
         Left            =   30
         TabIndex        =   30
         Top             =   5475
         Width           =   3930
         _Version        =   65536
         _ExtentX        =   6932
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Work List 등록"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
         RoundedCorners  =   0   'False
         Picture         =   "Machine41.frx":215F4
      End
   End
   Begin Threed.SSCommand cmdResultReg 
      Height          =   645
      Left            =   690
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   2160
      _Version        =   65536
      _ExtentX        =   3810
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "결과  등록  준비"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "Machine41.frx":21610
   End
   Begin Threed.SSCommand cmdOrder 
      Height          =   645
      Left            =   3165
      TabIndex        =   9
      Top             =   6210
      Width           =   2070
      _Version        =   65536
      _ExtentX        =   3651
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "Order       내리기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "Machine41.frx":220DA
   End
   Begin Threed.SSCommand cmdLoad 
      Height          =   870
      Left            =   9270
      TabIndex        =   33
      Top             =   135
      Visible         =   0   'False
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Load"
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
      Picture         =   "Machine41.frx":22BA4
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
Dim identbOpenKey      As Integer
Dim IdList               As String
Dim BlockKey            As Integer
Dim Errkey              As Integer
Dim next_flag               As Integer

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

Dim TotLen As Integer

Private Function AscLen(para As String) As Integer
    '해당 문자열의 한글이 포함된 Ascii Length를 가진다.
    Dim i As Integer
    
    AscLen = 0
    For i = 1 To Len(para)
        AscLen = AscLen + IIf(Asc(Mid(para, i, 1)) > 0, 1, 2)
    Next i
    
End Function

'
'   참고치 판정
'
Private Function Chk_Ref(sOrdCd As String, sSubNo As String, sRes As String, _
                        sex As String) As String

    Dim sStr    As String
    Dim SData() As String
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
            QSqlGetField 4, sStr, SData()
            
            If Trim(SData(4)) = "C" Then    '참고치 문자
                RefChar = Trim(SData(3))
                If RefChar <> Trim(sRes) Then
                    Chk_Ref = "*"
                End If
            ElseIf Trim(SData(4)) = "N" Then        '숫자
                RefVal = CSng(Val(Trim(sRes)))
                LowVal = CSng(SData(1)): HighVal = CSng(SData(2))
            
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


Private Sub Update_DJC020M(OrdSqNo As String)

    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
    Dim SqlDoc  As String
    
    sStr = " Select count(*) from LAB03_DB..DJC050M " _
            & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
            & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
            & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
            & "   and RSTGBN = '' "
    
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 1, sStr, tData()
            
            If Val(tData(1)) = 0 Then
                SqlDoc = " Update LAB03_DB..DJC020M " _
                        & "   set ORDSTAT = '1' " _
                        & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
                        & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
                        & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
                        & "   AND ORDSTAT IN ('0','6') "
                iRet_Cd = QSqlDBExec(SqlDoc, QsqlCode)
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
    Dim SqlDoc  As String
    
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
            sStr = sStr & "   and ORDCD = '" & .OrdCd & "'"
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
                    & "   and ORDCD = '" & .OrdCd & "'" _
                    & "   and SUBCD = '" & .SubNo & "'" _
                    & "   and ORDERNO = '" & sOrdNo & "'"
            
            If QSqlDBExec(sStr, QsqlConn) <> QSQL_SUCCESS Then
                '--- Insert(Sub 검사항목인 경우-조회 후 입력처리)
                ReDim INSDATA(1 To 5) As String
                
                '--- Insert할 항목 조회
                SqlDoc = " Select DISTINCT REQGBN, SPCGBN, RETGBN, RTNCD, IDNO " _
                        & "  from LAB04_DB..DJD010M " _
                        & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                        & "   and NUMGBN = '" & Mid(sLabNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                        & "   and ORDCD = '" & .OrdCd & "'" _
                        & "   and ORDERNO = '" & sOrdNo & "'"
                        
                If QSqlDBExec(SqlDoc, QsqlConn) = QSQL_SUCCESS Then
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
                        & "'" & .OrdCd & "', " _
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
        .OrdCd = ""
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
      identb!Seq_No = sample
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
      resulttb!Seq_No = sample2
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
    
    Dim wkdat As String
    Dim ix1 As Integer
        
    TotLen = AscLen(wkbuf)
    
    'For ix1 = 1 To Len(wkbuf)
'    If TotLen < 120 Then Exit Sub
    For ix1 = 1 To TotLen
        wkdat = Mid$(wkbuf, ix1, 1)
        
        If wkdat = "" Then Exit For
            
            Select Case phase

                Case 1
                    Select Case Asc(wkdat)
                        Case 1
                            phase = 2
                    End Select

                Case 2
                    Select Case Asc(wkdat)
                        Case 2 ', 3
                            If Len(RcvBuffer) = 127 Then
                                Call edit_data
                                RcvBuffer = ""
                                phase = 1
                                ix1 = ix1 + 69
'                            Else
'                                RcvBuffer = RcvBuffer & wkdat
                            End If
                        Case Else
                            If Trim(wkdat) = "?" Then wkdat = "0"
                            If Trim(wkdat) = "같" Then wkdat = Space$(2)
                            RcvBuffer = RcvBuffer & wkdat
                            phase = 2
                            If Len(RcvBuffer) = 127 Then
                                Call edit_data
                                RcvBuffer = ""
                                phase = 1
                                ix1 = ix1 + 69
                            End If
                    End Select

            End Select
    Next
    
End Sub
Sub PhaseCfg_Protocol1()
    
    Dim wkdat As String
    Dim ix1 As Integer

    
    For ix1 = 1 To Len(wkbuf) * 1.5

        wkdat = Mid$(wkbuf, ix1, 1)
        If wkdat = "" Then Exit For
            Select Case phase

                Case 1
                    Select Case Asc(wkdat)
                        Case 1
                            phase = 2
                    End Select

                Case 2
                    Select Case Asc(wkdat)
                        Case 2, 3
                            If Len(RcvBuffer) = 129 Then
                               Call edit_data
                                RcvBuffer = ""
                                phase = 1
                                ix1 = ix1 + 69
                            Else
                                RcvBuffer = RcvBuffer & wkdat
                            End If
                        Case Else
                            If Trim(wkdat) = "?" Then wkdat = "0"
                            If Trim(wkdat) = "같" Then wkdat = Space$(2)
                            RcvBuffer = RcvBuffer & wkdat
                            phase = 2
                            If Len(RcvBuffer) = 129 Then
                                Call edit_data
                                RcvBuffer = ""
                                phase = 1
                                ix1 = ix1 + 69
                            End If
                    End Select

            End Select
    Next
    
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

         If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
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
        
    Dim SEQNO       As String
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
    
    Dim j           As Integer
    Dim StNm        As Integer
    
'    If Left$(RcvBuffer, 1) = "<" Then
'        next_flag = True
'        RcvBuffer = ""
'        Exit Sub
'    ElseIf Left$(RcvBuffer, 1) = ":" Then
'        next_flag = False
'        Exit Sub
'    End If
'
'    a = InStr(RcvBuffer, "E")
'    If a = 0 Then
'        next_flag = False
'        Exit Sub
'    End If
    
'sampleNo.얻기->slipno에 해당
    Call spdface.GetText(1, SampleCnt + 1, temp)
'
'    If temp = "" Then
'       Exit Sub
'    End If
    
'    long_slip1 = Val(Mid$(RcvBuffer, 15, 6))
'    slipno = Right$("0000" + Trim$(Mid$(RcvBuffer, 15, 6)), 4)

'    'Date 얻기
'    tmpbuffer = ""
'    tmpbuffer = Mid$(RcvBuffer, 21, 8)
'
'    'Time 얻기
'    tmpbuffer = ""
'    tmpbuffer = Mid$(RcvBuffer, 30, 5)

    '샘플번호=> seqno
    
    If long_slip2 = 0 Then
        SampleCnt = SampleCnt + 1
        long_slip2 = long_slip1
    ElseIf long_slip1 <= long_slip2 Then
        SampleCnt = SampleCnt + 1
        long_slip2 = long_slip1
    Else
        SampleCnt = SampleCnt + (long_slip1 - long_slip2)
        long_slip2 = long_slip1
    End If

'---검사결과값 얻기 ------------------------------------------------------------
    For ix1 = 1 To 30
        tresult(ix1) = ""
    Next ix1
    
    If Len(RcvBuffer) = 127 Then
        StNm = 81
    Else
        StNm = 80
    End If
    
'    If SampleCnt = 1 Then
'        StNm = 81
'    Else
'        StNm = 80
'    End If
    
    
    For j = 1 To 12
'        If j <= 18 And j <> 8 And j <> 12 Then
        Select Case j
'            Case 12
'                tresult(j) = Trim$(Mid$(RcvBuffer, StNm, 3))
            Case 1, 2, 3, 4, 6, 7, 9, 10, 11, 12
                tresult(j) = Trim$(Mid$(RcvBuffer, StNm, 4))
                StNm = StNm + 4
            Case 5, 8
                tresult(j) = Trim$(Mid$(RcvBuffer, StNm, 3))
                StNm = StNm + 3
        End Select
        
'            tresult(j) = Trim$(Mid$(RcvBuffer, StNm, 4))
'        Else
'            tresult(j) = Trim$(Mid$(RcvBuffer, StNm, 3))
'        End If
'        StNm = StNm + Len(tresult(j))
        If Len(RcvBuffer) <= StNm Then Exit For
    Next
'    a = InStr(tresult(3), "/ul")
'    If a <> 0 Then
''        tresult(3) = Left(tresult(3), a - 1)
'        tresult(3) = Trim$(Mid(tresult(3), a + Len("/ul")))
'    End If
'
'    tresult(4) = Trim$(Mid$(RcvBuffer, 83, 5)) 'NIT
'
'    tresult(5) = Trim$(Mid$(RcvBuffer, 95, 17)) 'PRO
'    a = InStr(tresult(5), "mg/dl")
'    If a <> 0 Then
''        tresult(5) = Left(tresult(5), a - 1)
'        tresult(5) = Trim$(Mid(tresult(5), a + Len("mg/dl")))
'    End If
'
'    tresult(6) = Trim$(Mid$(RcvBuffer, 115, 17)) 'GLU
'    a = InStr(tresult(6), "mg/dl")
'    If a <> 0 Then
''        tresult(6) = Left(tresult(6), a - 1)
'        tresult(6) = Trim$(Mid(tresult(6), a + Len("mg/dl")))
'    End If
'
'    tresult(7) = Trim$(Mid$(RcvBuffer, 135, 17)) 'KET
'    a = InStr(tresult(7), "mg/dl")
'    If a <> 0 Then
''        tresult(7) = Left(tresult(7), a - 1)
'        tresult(7) = Trim$(Mid(tresult(7), a + Len("mg/dl")))
'    End If
'
'    tresult(8) = Trim$(Mid$(RcvBuffer, 155, 17)) 'UBG
'    a = InStr(tresult(8), "mg/dl")
'    If a <> 0 Then
''        tresult(8) = Left(tresult(8), a - 1)
'        tresult(8) = Trim$(Mid(tresult(8), a + Len("mg/dl")))
'    End If
'
'    tresult(9) = Trim$(Mid$(RcvBuffer, 175, 17)) 'BIL
'    a = InStr(tresult(9), "mg/dl")
'    If a <> 0 Then
''        tresult(9) = Left(tresult(9), a - 1)
'        tresult(9) = Trim$(Mid(tresult(9), a + Len("mg/dl")))
'    End If
'
'    tresult(10) = Trim$(Mid$(RcvBuffer, 195, 17)) 'ERY
'    a = InStr(tresult(10), "/ul")
'    If a <> 0 Then
''        tresult(10) = Left(tresult(10), a - 1)
'         tresult(10) = Trim$(Mid(tresult(10), a + Len("/ul")))
'   End If
'
''Result 약간의 변신
'    For i = 1 To 10
'        If UCase(tresult(i)) = "NEG" Or UCase(tresult(i)) = "NORM" Then
'            tresult(i) = "음성"
'        ElseIf UCase(tresult(i)) = "POS" Then
'            tresult(i) = "양성"
'        End If
'    Next

'---검사명, 검사결과값 spread에 뿌리기------------------------------------------------------------
    'txslipno.Text = slipno
    txsapno.Text = Format(PrevCnt + SampleCnt, "0000")
    txtguide.Text = "TX Data!!"

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
        If TestNameTable(i).code <> "" Then
            tcode = Format$(i, "00")
            If tcode <> "" Then
                'add_db_identb SeqNo, slipno
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
    
    
End Sub
Sub Test()
'
'    RcvBuffer = ";E                4 30.06.98 13:47 SG1.020      PH  6      LEU        neg      NITpos  pos PRO   25 mg/dl   1+ GLU       norm      KET        neg      UBG       norm      BIL        neg      ERY        neg      NAG                 21"
'
'
'    Call edit_data
    

    Open App.Path & "\dump.dat" For Input As #3
    Test_OpenFlag = 2
    wkbuf = ""
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

    Close #3

    Call PhaseCfg_Protocol

End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim i           As Integer
    Dim K           As Integer
    
    Dim Tmp         As Variant
    Dim tmpName     As Variant
    
    Dim temp        As Variant
    Dim ChkExist    As Integer
    
    'Modify
    Dim iListExist%
    Dim iRetVal%
    
    With spdList
        If .MaxRows = 0 Then
            MsgBox "Work List에 등록할 자료가 없습니다. 조회를 실행해 주십시오.", vbCritical
            Exit Sub
        End If
        
        ChkExist = False
        
        For i = 1 To .MaxRows
            Call .GetText(1, i, Tmp)
            If Tmp = True Then
                Call .GetText(2, i, Tmp)
                
                '--- WorkList Spread에 존재하는제 체크
                If spdface.MaxRows <> 0 Then
                    For K = 1 To spdface.MaxRows
                        Call spdface.GetText(1, K, temp)
                        If Trim(Tmp) = Left(Trim(temp), 16) Then
                            MsgBox "접수번호 " & Trim(Tmp) & " 가 WorkList에 존재합니다. " _
                                    & "확인해 주십시오.", vbCritical
                            
                            iListExist = True
                            Exit For
                        Else
                            iListExist = False
                        End If
                    Next K
                    
                    '---WorkList Spread에 존재하지 않으면 넣는다.
                    If iListExist = False Then
                        iRetVal = Row_Plus(spdface)
                        'Call spdWork.SetText(1, spdWork.MaxRows, Trim(tmp))     '접수번호
                        
                        Call .GetText(3, i, tmpName)
                        Call spdsettext(spdface, 1, iRetVal + 1, Trim(Tmp))
                        Call spdsettext(spdface, 2, iRetVal + 1, Trim(tmpName))   '성명
                        
                        nowcount = nowcount + 1
                        Call add_db_identb(Format(nowcount, "0000"), Trim(Tmp))
                        
                        .Col = -1: .Row = i
                        .Action = SS_ACTION_DELETE_ROW
                        .MaxRows = .MaxRows - 1
                        i = i - 1
                    End If
                    '---
                        
                Else
                    Exit Sub
                End If
                '-------------------------------------
                
                ChkExist = True
            End If
        Next i
    End With
    
    iRetVal = Row_Plus(spdface)
    
    CurSampCnt = iRetVal
    
'    MsgBox CurSampCnt
    
    If ChkExist <> True Then
        MsgBox "Work List에 등록할 자료가 없습니다. 조회를 실행해 주십시오.", vbCritical
    End If
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
            
            identb.MoveLast
            CurrentTbRows = nowcount - PrevCnt
            
            For i = Val(StartBRow) To Val(EndBRow)
                identb.Seek "=", Format$(i + PrevCnt, "0000")
                identb.Delete
                SampleCnt = SampleCnt - 1
                resulttb.Seek "=", Format$(i + PrevCnt, "0000")
                
                If resulttb.NoMatch = False Then
                   Do Until resulttb.EOF
                       If resulttb!Seq_No <> Format$(i + PrevCnt, "0000") Then Exit Do
                       
                       resulttb.Delete
                            
                       resulttb.MoveNext
                   Loop
                End If
            Next
            
            If EndBRow <> CurrentTbRows Then
                For i = EndBRow + 1 To CurrentTbRows
                    
                    identb.Seek "=", Format$(i + PrevCnt, "0000")
                    identb.Edit
                    identb!Seq_No = Format$(i + PrevCnt - (EndBRow - StartBRow + 1), "0000")
                    identb.Update
                    
                    resulttb.Seek "=", Format(i + PrevCnt, "0000")
                    
                    If resulttb.NoMatch = False Then
                       Do Until resulttb.EOF
                           If resulttb!Seq_No <> Format(i + PrevCnt, "0000") Then Exit Do
                           
                           resulttb.Edit
                           resulttb!Seq_No = Format$(i + PrevCnt - (EndBRow - StartBRow + 1), "0000")
                           resulttb.Update
                           
                           resulttb.MoveNext
                       Loop
                    End If
                    
                Next
            End If
            
        '삭제하는 Spread 라인의 텍스트를 지움.
            spdface.BlockMode = True
            spdface.Col = -1
            spdface.col2 = -1
            spdface.Row = StartBRow
            spdface.row2 = EndBRow
            spdface.Action = SS_ACTION_DELETE_ROW
            spdface.BlockMode = False

        '1st Column(SlipNo)의 색깔을 노란색
            spdface.BlockMode = True
            spdface.Col = 1
            spdface.col2 = 1
            spdface.Row = -1
            spdface.row2 = -1
            spdface.BackColor = &HC0FFFF
            spdface.BlockMode = False
            
            nowcount = nowcount - (EndBRow - StartBRow + 1)   '삭제한 후에 다시 삭제할 경우를 대비
            'SampleCnt = SampleCnt - (EndBRow - StartBRow + 1)   '삭제한 후에 다시 삭제할 경우를 대비
            txsapno = Format(nowcount, "0000")
            txtguide = "Data 삭제!!"
            
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

Private Sub cmdLoad_Click()
' Call Test
    Comm1.Output = Chr$(21) '-- NAK
    
End Sub

Private Sub cmdOrder_Click()
    
    Call Send_Order
    
    fmeOrder.Visible = False
    cmdOrder.Visible = False
    
    cmdReg.Visible = True
    cmdDelete.Visible = True
    spdface.Visible = True
    fmeSlipSeq.Visible = True
    
    fmeGuide.Left = 9720
    fmeGuide.Top = 0
    fmeGuide.Width = 2115
    fmeGuide.Height = 585

    txtguide.Left = 150
    txtguide.Top = 180
    txtguide.Width = 1815
    txtguide.Height = 315
    
    txtguide.Text = "RX Results!!"
    
End Sub

Private Sub cmdQuery_Click()
    Dim iRet    As Integer
    Dim sStr    As String
    Dim tData() As String
    
    Dim db As Database
    Dim Rs As Recordset
    
    '--- 조회조건 체크
    If Not IsDate(mskDate) Then
        MsgBox "조회를 원하는 접수일자를 입력해 주십시오.", vbExclamation
        mskDate.SetFocus
        Exit Sub
    End If
    
    '--- Index Open
    'If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlConn) <> QSQL_SUCCESS Then Exit Sub
    
    
    spdList.MaxRows = 0
    
    sStr = "Select Distinct D1.LABDATE,  D1.PARTGBN, D1.LABSEQ, B2.NAME " _
            & "  from H_RESULT D1, PERSON B2  " _
            & " where D1.LABDATE = '" & Trim(mskDate.ClipText) & "'" _
            & "   and D1.PARTGBN + D1.SPECIMENCD + D1.TESTITEMSEQ in ("
    
    Set dbcode = OpenDatabase(FileName & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable", dbOpenTable)

    
    sStr = sStr & "''"
    If tbcode.RecordCount > 0 Then
        
        tbcode.MoveFirst
        Do Until tbcode.EOF
            If (Trim(tbcode!code) & "") <> "" Then
            sStr = sStr & ",'" & Trim(Mid(tbcode!code, 2, 9)) & "'"
            End If
            tbcode.MoveNext
        Loop
        tbcode.Close
        dbcode.Close
    End If
    sStr = sStr & ") "
    
'            & "   and SUBSTRING(D1.ORDCD,1,2) = 'HB' "
    If cboSelect.ListIndex = 0 Then
        sStr = sStr & "  and (D1.RESULT = '' or D1.RESULT is null) "
    Else
        sStr = sStr & "  and Not(D1.RESULT = '' or D1.RESULT is null)"
    End If
    
'    sStr = sStr & "  and  D1.IDLEFT = B2.IDLEFT " _
                & "  and  D1.IDRIGHT = B2.IDRIGHT " _
                & "order by D1.PARTGBN, D1.LABSEQ "
                
    sStr = sStr & "  and  D1.REGNO = B2.REGNO " _
                & "order by D1.PARTGBN, D1.LABSEQ "

'    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
    Set db = OpenDatabase(SemiDb)
'    Set db = OpenDatabase("E:\인터페이스\인천동구보건소\SemiLIS\Source\Database\Semilis.mdb")

    Set Rs = db.OpenRecordset(sStr)
        
    If Rs.RecordCount > 0 Then
        Rs.MoveLast
        Rs.MoveFirst
        
        Do Until Rs.EOF
            'iRet = QSqlGetRow(sStr, QsqlConn)
            'If iRet <> QSQL_SUCCESS Then
            '    Exit Do
            'End If
            
            'QSqlGetField 4, sStr, tData()
                    
            With spdList
                .MaxRows = .MaxRows + 1
                Call .SetText(2, .MaxRows, Rs.Fields(0) & "-" & Rs.Fields(1) & "-" & Rs.Fields(2))
                Call .SetText(3, .MaxRows, Trim(Rs.Fields(3)))
            End With
            Rs.MoveNext
        Loop 'Until (iRet <> QSQL_SUCCESS)
    End If
    'iRet = QSqlSelectFree(QsqlConn)
    
    If spdList.MaxRows = 0 Then
        MsgBox "해당자료가 존재하지 않습니다.", vbExclamation
    End If
    
    '--- Index Close
    'iRet = Qsqlclose(QsqlConn, ONECLOSE)
    
    Set Rs = Nothing

End Sub

Private Sub cmdReg_Click()
    
    Dim i%, rt%, seqnoVar, slipnoVar, tcode$, tresult$
    Dim tmpSlip    'general에 선언
    Dim tmpResult
    Dim ExistTxtKey As Integer
    
    Dim FileName    As String
    Dim iRet    As Integer
    

    Dim iR  As Integer
    Dim iC  As Integer
    Dim Tmp As Variant
    Dim ix1 As Integer
        
    Dim LabNo   As String
    Dim SampNo  As String
    Dim sSex    As String
    Dim sOrderNo    As String
    Dim sRtnCd  As String
    
    Dim ChkTrans    As Integer
    Dim ChkExist    As Integer
    
    If StartBRow = -1 And EndBRow = -1 Then
        StartBRow = 1
        EndBRow = nowcount - PrevCnt
    End If

    For i = StartBRow To EndBRow
        rt = spdface.GetText(1, i, tmpSlip)
        rt = spdface.GetText(2, i, tmpResult)
        If tmpSlip = "" Or tmpResult = "" Then
            ExistTxtKey = False
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If spdface.MaxCols = 12 Then
    Else
        MsgBox "결과 등록 준비를 아직 하지 않았습니다. 결과 등록 준비를 클릭한 후에만 서버에 등록을 할 수 있습니다!!"
        Exit Sub
    End If

    
    If identbOpenKey = True And ExistTxtKey = True Then
        If StartBCol = -1 And EndBCol = -1 And BlockKey = True Then
            identb.Index = "primarykey"
            identb.MoveFirst

            Screen.MousePointer = 11

            For i = Val(StartBRow) To Val(EndBRow)
            'MDB(Access)에 일단 등록
                DoEvents

                   rt = spdface.GetText(1, i, seqnoVar)
                   rt = spdface.GetText(1, i, slipnoVar)

                   identb.Index = "primarykey"
                   identb.Seek "=", Format(i + PrevCnt, "0000")

                   If identb.NoMatch = False Then
                       identb.Edit
                       identb!ChkResult = "*"
                       identb.Update
                       spdface.Col = 1
                       spdface.col2 = 1
                       spdface.Row = i
                       spdface.row2 = i
                       spdface.BlockMode = True
                       spdface.BackColor = &HFFFFC0
                       spdface.BlockMode = False
                   End If


               Screen.MousePointer = 0
           Next

           
        '--- Index Open
        If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlConn) <> QSQL_SUCCESS Then Exit Sub
        If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlCode) <> QSQL_SUCCESS Then Exit Sub
        
        MousePointer = 11
        
        ChkExist = False
        '----- Server로 결과등록
        For iR = 1 To CurSampCnt
            '--- Check Box
            Call spdface.GetText(1, iR, Tmp)
'            If tmp = True Then
            If Trim(Tmp) <> "" Then     'yk
                With spdface
                    Call .GetText(1, iR, Tmp)
                    LabNo = Left(Tmp, 8) & Mid(Tmp, 10, 1) & Mid(Tmp, 12, 5)      '접수번호
                    'Call .GetText(4, iR, tmp)
                    'SampNo = Trim(tmp)          'Sample No
                    Call .GetText(10, iR, Tmp): sSex = Trim(Tmp)
                    Call .GetText(11, iR, Tmp): sOrderNo = Trim(Tmp)
                    Call .GetText(12, iR, Tmp): sRtnCd = Trim(Tmp)
                End With
                            
                ChkExist = True
                
                For iC = 1 To 8
                    If Trim(TestNameTable(iC).code) <> "" Then
                        Call spdface.GetText(iC + 1, iR, Tmp)
                        With Insert_Server(iC)
                            .OrdCd = TestNameTable(iC).code
                            .SubNo = ""
                            .Result = Trim(Tmp)
                            .Ref = Chk_Ref(.OrdCd, .SubNo, .Result, sSex)
                            ''--- Hi_Result 내용 Update
                            ''Call Update_DB_Result(SampNo, TestNameTable(iC).EqCd, .Result, iC)
                        End With
                    End If
                Next iC
                                    
                '----- 구조체에 저장된 결과 Server에 등록
                ret = QSqlBeginTrans()
                DBEngine.Workspaces(0).BeginTrans
                ChkTrans = True
                
                For ix1 = 1 To 8
                    '----- 검사항목별 결과입력(Batch)
                    If Append_To_Server(LabNo, ix1, sOrderNo, sRtnCd) <> True Then
                        ChkTrans = False
                        Exit For
                    End If
                Next ix1
                
                If ChkTrans = False Then
                    DBEngine.Workspaces(0).Rollback
                    ret = QSqlRollBack()      'TRANSACTION 에러종료
                Else
                    DBEngine.Workspaces(0).CommitTrans
                    ret = QSqlCommitTrans()    'TRANSACTION 정상종료
                    '--- 진료과별 처방내역 Update
                    ''Call Update_DJC020M(LabNo)
                    Call Update_DJC020M(sOrderNo)
                    
                    '--- 등록체크 Update(MDB)
                    'Call Update_RegChk(SampNo)

                End If
            End If
        Next iR
        
        '--- Index Close
        ret = Qsqlclose(QsqlConn, ONECLOSE)
        ret = Qsqlclose(QsqlCode, ONECLOSE)
    
        If ChkExist <> True Then
            MsgBox "선택된 자료가 없습니다. 등록을 원하는 환자를 선택해 주십시오.", vbInformation
        End If
        
        MousePointer = 0
        
        
        Else

            MsgBox "잘못된 서버에 결과 등록 방법입니다." & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "결과 등록을 하십시요!!"

        End If
   Else

        MsgBox "데이터가 없거나 검사 결과를 전송받지 않으셨습니다!!"

   End If

   BlockKey = False
   spdface.EditMode = True
   spdface.EditMode = False

'현재의 마지막 Row를 점검
   spdface.Row = nowcount - PrevCnt
    
    
    
    
    
   
End Sub

Private Sub cmdResultReg_Click()
    
    Dim tData() As String
    Dim sStr    As String
    Dim sLabNo  As String
    Dim tmpSlip
    Dim ExistTxtKey As Integer
    Dim CurRow  As Long
    Dim i       As Integer
    Dim rt      As Integer
    Dim tmpResult
    Dim SqlDoc  As String
    Dim SqlConn As Long
    
    If CallLabKey = True Then
        Call Test
    End If
    If CallLabKey = True Then
        Call Test
    End If
    If CallLabKey = True Then
        Call Test
    End If
    
   If S0SUB_Open(D0COM_SERVER01, Me.hWnd, SqlConn) <> QSQL_SUCCESS Then Exit Sub
   
    spdface.MaxCols = 12
    
    Call spdsettext(spdface, 10, 0, "Sex")
    Call spdsettext(spdface, 11, 0, "OrdNo")
    Call spdsettext(spdface, 12, 0, "RtnCd")
    
    For i = 1 To CurSampCnt
        rt = spdface.GetText(1, i, tmpSlip)
        rt = spdface.GetText(2, i, tmpResult)
        If tmpSlip = "" Or tmpResult = "" Then
            ExistTxtKey = False
            spdface.MaxCols = 9
            MsgBox "Worklist 등록 또는 검사기기로부터 결과가 전송되지 않았습니다!!"
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If ExistTxtKey = True And identbOpenKey = True Then
        
        For i = 1 To CurSampCnt
        
            With spdface
                '----- 환자정보 조회/표시
                SqlDoc = " Select Distinct C2.PATNM, C2.SEX, D1.ORDERNO, D1.RTNCD " _
                        & "  from LAB04_DB..DJD010M D1, LAB03_DB..DJC020M C2 " _
                        & " where D1.LABDATE = '" & Left$(tmpSlip, 8) & "'" _
                        & "   and D1.NUMGBN  = '" & Mid$(tmpSlip, 10, 1) & "'" _
                        & "   and D1.LABSQNO = '" & Mid$(tmpSlip, 12, 5) & "'" _
                        & "   and SUBSTRING(D1.ORDERNO,1,8) = C2.ORDDATE " _
                        & "   and SUBSTRING(D1.ORDERNO,9,2) = C2.DEPTCD " _
                        & "   and SUBSTRING(D1.ORDERNO,11,5) = C2.SEQNO " _
                        & "   and substring(D1.ORDCD,1,2) = 'HB' " _
                        & "   and D1.IDNO = C2.IDNO "
                        
                If QSqlDBExec(SqlDoc, SqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, SqlConn) = QSQL_SUCCESS Then
                        
                        QSqlGetField 4, sStr, tData()
                        
                        ''.MaxRows = .MaxRows + 1
                        
        '''                sLabNo = Left$(TableSam!Lab_ID, 8) & "-" & Mid$(TableSam!Lab_ID, 9, 1) & "-" _
        '''                        & Right$(TableSam!Lab_ID, 5)
        '''                Call .SetText(2, .MaxRows, Trim(sLabNo))        '접수번호
        '''                Call .SetText(3, .MaxRows, Trim(tData(1)))      '이름
        '''                Call .SetText(4, .MaxRows, Trim$(TableSam!SampNo))  'Sample No
                        Call .SetText(10, i, Trim(tData(2)))     'Sex
                        Call .SetText(11, i, Trim(tData(3)))     'OrderNo
                        Call .SetText(12, i, Trim(tData(4)))
                        '--- 결과 표시
                        'Call DIsp_Result(TableSam!SampNo, .MaxRows)
                    End If
                End If
                Return_cd = QSqlSelectFree(SqlConn)
        
            End With
        
        Next
    
    End If
    
    Call Qsqlclose(SqlConn, ONECLOSE)
    
End Sub

Private Sub cmdTest_Click()
    Call Test
End Sub

Private Sub Comm1_OnComm()
    
    Screen.MousePointer = 11
    'Dim wkbuf As String
    
    
    Select Case Comm1.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                               ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
                          
            Timer1.Enabled = False
            Timer2.Enabled = True
            wkbuf = Comm1.Input
            
            Print #1, wkbuf;    'Test
           
            Call PhaseCfg_Protocol
            
'            Comm1.Output = Chr$(6) '-- ACK
            
            txtguide.Text = "Data 전송 완료!!"
                
        Case MSCOMM_EV_CTS      'j
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
           
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    'form을 가운데에 위치
    Me.Top = 0
    Me.Left = 0
    Me.Height = INTmain00.Height - INTmain00.pnlMain.Height - 500
    Me.Width = INTmain00.Width - 200
    
    Dim tablerows As Integer
    Dim sRow As Integer
    Dim i As Integer
    Dim TestItemNo As Integer
    
        
    Set dbcode = OpenDatabase(FileName & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")
    
    tbcode.MoveLast
    tablerows = tbcode.RecordCount
        
    tbcode.MoveFirst
   
    For sRow = 1 To 30
        TestNameTable(sRow).name = tbcode!name & ""
        TestNameTable(sRow).code = tbcode!code & ""
        If TestNameTable(sRow).code <> "" Then
            TestItemNo = TestItemNo + 1
            TestNameTable(sRow).col_cnt = TestItemNo + 2
            spdface.MaxCols = TestNameTable(sRow).col_cnt
            Call spdsettext(spdface, TestNameTable(sRow).col_cnt, 0, TestNameTable(sRow).name)
        End If
        tbcode.MoveNext
    Next
    
    SampleCnt = 0
    
    Timer1.Enabled = True
    Timer2.Enabled = False
'    If TestItemNo >= 17 Then
'        spdface.MaxCols = TestItemNo + 1
'        For i = 1 To TestItemNo
'            Call spdsettext(spdface, i + 1, 0, TestNameTable(i).name)
'        Next
'    Else
'        spdface.MaxCols = 9
'        For i = 1 To TestItemNo
'            Call spdsettext(spdface, i + 1, 0, TestNameTable(i).name)
'        Next
'
'        For i = TestItemNo + 1 To spdface.MaxCols - 1
'            Call spdsettext(spdface, i + 1, 0, "-")
'        Next
'    End If
    
    mskDate = Format(Now, "yyyy-mm-dd")
    With cboSelect
        .AddItem ("미등록 자료")
        .AddItem ("등록된 자료")
        .ListIndex = 0
    End With
    
'1st Column(SlipNo)의 색깔을 노란색
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.col2 = 1
    spdface.Row = -1
    spdface.row2 = -1
    spdface.BackColor = &HC0FFFF
    spdface.BlockMode = False

'Interface Result를 일단 Lock
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.col2 = spdface.MaxCols
    spdface.Row = 1
    spdface.row2 = spdface.MaxRows
    spdface.Lock = True
    spdface.BlockMode = False

'Spread.Row Initialization
    spdface.Row = 0
    
    tbcode.Close
    dbcode.Close
    
    LblMMDD.Caption = Val(Left$(textmmdd, 2)) & "월" & " " & Val(Right$(textmmdd, 2)) & "일"
    
        
    If OrderKey = True Then
        fmeOrder.Visible = True
        
        spdface.Visible = False
        cmdReg.Visible = False
        cmdDelete.Visible = False
        fmeSlipSeq.Visible = False
        
        For i = 1 To 30
                       
            If i < 16 Then
                fmeOrder.Width = 6360
                Call spdChksettext(spdWorklist1, 1, i, TestNameTable(i).code)
                Call spdsettext(spdWorklist1, 2, i, TestNameTable(i).name)
            Else
                If TestNameTable(i).code <> "" Or TestNameTable(i).name <> "" Then
                    fmeOrder.Width = 10260
                    Call spdChksettext(spdWorklist2, 1, i - 15, TestNameTable(i).code)
                    Call spdsettext(spdWorklist2, 2, i - 15, TestNameTable(i).name)
                End If
            End If
        
        Next i
        
        fmeGuide.Left = 5610
        fmeGuide.Top = 0
        fmeGuide.Width = 6240
        fmeGuide.Height = 585
        
        txtguide.Width = 5000
        txtguide.Text = "Worklist & Order Process"
        
    Else
        fmeOrder.Visible = False
        cmdOrder.Visible = False
        
        txtguide.Text = "RX Results!!"
    End If

    
    On Error GoTo PortOpenErr:
    errfound = False
    
    Set dbcomm = OpenDatabase(FileName & commstr)
    Set tbcomm = dbcomm.OpenRecordset("cfgcomm")

    tbcomm.MoveFirst
        
    With transcfg
        .Port = tbcomm!Port
        .data_bit = tbcomm!data_bit
        .stop_bit = tbcomm!stop_bit
        .baud_rate = tbcomm!baud_rate
        .parity = tbcomm!parity
        .blocksize = tbcomm!blocksize
    End With
    
    tbcomm.Close
    dbcomm.Close
    
    With Comm1
        .CommPort = transcfg.Port
        .Settings = transcfg.baud_rate & "," & transcfg.parity & "," & transcfg.data_bit & "," & transcfg.stop_bit
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With
          
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
    
    Porttag = 1     'PortOpen에 성공을 나타냄
    identbOpenKey = False
    
    Set db = OpenDatabase(FileName & "comm\" & strmmdd + ".mdb")
    Set identb = db.OpenRecordset("sp_identify")
    Set resulttb = db.OpenRecordset("sp_result")
    
    identbOpenKey = True
    Porttag = 2     'OpenDB에 성공을 나타냄

'해당일의 이전까지의 샘플의 갯수를 db.identb에서 읽어 옴
    identb.Index = "Primarykey"
    
    For i = 1 To 9999
        identb.Seek "=", Format(i, "0000")
        
        If identb.NoMatch = False Then
        Else
            If i = 1 Then
                nowcount = 0
            Else
                nowcount = i - 1
            End If
            
            Exit For
        End If
    Next
    
    'nowcount = identb.RecordCount
    PrevCnt = nowcount
        
    phase = 1
    EditCnt = 0
    SampleCnt = 0
    Test_OpenFlag = 0
    long_slip1 = 0
    long_slip2 = 0
    
    Open FileName & machstr & ".log" For Output As #1
    
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
    
    Open FileName & machstr & "Buff.log" For Output As #2
          
    Porttag = 3     'OpenSamFile에 성공을 나타냄
    FrmFlag = 41
          
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
          
    If TestKey = 1 Then
        cmdTest.Visible = True
    End If
              
    Timer1.Enabled = True
    Timer1.Interval = 2000

Exit Sub

PortOpenErr:
    
    errfound = True
    
    MsgBox "통신 구성 에러!!  통신구성을 다시 설정해 주십시요!!"
    
    If Porttag = 2 Then     'OpenDB까지 성공, OpenSamFile에서 실패
        identb.Close
        resulttb.Close
        db.Close
        Close #1
    End If
    
    Porttag = 0
    
    If Test_OpenFlag = 1 Then    'Sub Test에서 Open시 에러발생하면 Test_OpenFlag = 1가 되고,
        Close #3                   '완전히 Open되면 Test_OpenFlag = 2
        Porttag = 3
    End If
    
    Resume Next
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Porttag = 3 Then
        Comm1.PortOpen = False
        identb.Close
        resulttb.Close
        db.Close
        identbOpenKey = False
        Close #1
        Close #2
    End If
End Sub

Private Sub mskDate_GotFocus()
    With mskDate
        .SelStart = 8
        .SelLength = .MaxLength
    End With
End Sub

Private Sub mskDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    
        KeyAscii = 0
    
    End If
    
    
End Sub

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
    Dim irow    As Integer
    Dim Tmp As Variant
    
    If BlockRow = 0 Or BlockRow2 = 0 Or BlockRow = BlockRow2 Then Exit Sub
    
    If BlockRow < BlockRow2 Then
        sNo = BlockRow: eNo = BlockRow2
    Else
        sNo = BlockRow2: eNo = BlockRow
    End If

    For irow = sNo To eNo
        With spdList
            Call .GetText(1, irow, Tmp)
            If Tmp = True Then
                Call .SetText(1, irow, "0")
            Else
                Call .SetText(1, irow, "1")
            End If
        End With
    Next irow
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

Private Sub Timer1_Timer()
    
    Comm1.Output = Chr$(21) '-- NAK
    
End Sub


Private Sub Timer2_Timer()
    
    Comm1.Output = Chr$(6) '-- ACK

'    Timer2.Enabled = False
End Sub


