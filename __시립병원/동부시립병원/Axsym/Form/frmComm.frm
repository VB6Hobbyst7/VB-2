VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmComm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "인터페이스"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   FillStyle       =   0  '단색
   Icon            =   "frmComm.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15210
   WindowState     =   2  '최대화
   Begin VB.CheckBox ChkError 
      Caption         =   "에러내역"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Left            =   6810
      TabIndex        =   27
      Top             =   150
      Width           =   1245
   End
   Begin VB.CheckBox chkMdbSave 
      Caption         =   "Auto로컬등록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   8310
      TabIndex        =   25
      Top             =   150
      Value           =   1  '확인
      Width           =   1545
   End
   Begin VB.CheckBox chkSvrSave 
      Caption         =   "Auto서버등록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   9900
      TabIndex        =   24
      Top             =   150
      Width           =   1545
   End
   Begin FPSpreadADO.fpSpread spdResult1 
      Height          =   4875
      Left            =   90
      TabIndex        =   22
      Top             =   480
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   8599
      _StockProps     =   64
      ColsFrozen      =   8
      EditEnterAction =   5
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   22
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmComm.frx":08CA
   End
   Begin VB.CommandButton cmdACK 
      Caption         =   "Dump Test"
      Height          =   345
      Left            =   3360
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame fraCmdBar 
      Height          =   705
      Left            =   60
      TabIndex        =   8
      Top             =   9240
      Width           =   15075
      Begin VB.CommandButton cmdAction 
         Caption         =   "Exit"
         Height          =   375
         Index           =   4
         Left            =   13500
         TabIndex        =   26
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Run"
         Height          =   375
         Index           =   0
         Left            =   8220
         TabIndex        =   12
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Stop"
         Height          =   375
         Index           =   1
         Left            =   9540
         TabIndex        =   11
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Clear"
         Height          =   375
         Index           =   2
         Left            =   10860
         TabIndex        =   10
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Save"
         Height          =   375
         Index           =   3
         Left            =   12180
         TabIndex        =   9
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "작업대기 중.."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   900
         TabIndex        =   14
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " 상태 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   13
         Top             =   315
         Width           =   615
      End
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14070
      Top             =   9750
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13650
      Top             =   9750
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   12510
      Top             =   8940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":146A
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":14C8
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1526
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1584
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":15E2
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1640
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   14445
      Top             =   8850
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   13080
      Top             =   9750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":169E
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1C38
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":21D2
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":276C
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2FFE
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3158
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":32B2
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame FrameResult 
      Height          =   3525
      Left            =   90
      TabIndex        =   4
      Top             =   5460
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   6218
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpreadADO.fpSpread spdRstDetail 
         Height          =   3195
         Left            =   150
         TabIndex        =   23
         Top             =   210
         Width           =   7695
         _Version        =   393216
         _ExtentX        =   13573
         _ExtentY        =   5636
         _StockProps     =   64
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm.frx":340C
         ScrollBarTrack  =   1
      End
   End
   Begin Threed.SSFrame FrameInterface 
      Height          =   3525
      Left            =   8220
      TabIndex        =   2
      Top             =   5460
      Width           =   6885
      _Version        =   65536
      _ExtentX        =   12144
      _ExtentY        =   6218
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtCom 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmComm.frx":398A
         Top             =   480
         Width           =   6645
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Interface Dump Data :"
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
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   4785
      End
   End
   Begin MSComctlLib.ListView lvwCuData 
      Height          =   2940
      Left            =   3540
      TabIndex        =   5
      Top             =   6270
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Threed.SSFrame FrameError 
      Height          =   3525
      Left            =   8220
      TabIndex        =   0
      Top             =   5460
      Visible         =   0   'False
      Width           =   6885
      _Version        =   65536
      _ExtentX        =   12144
      _ExtentY        =   6218
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton CmdErrClear 
         Caption         =   "내역지움"
         Height          =   285
         Left            =   5520
         TabIndex        =   28
         Top             =   150
         Width           =   1275
      End
      Begin VB.ListBox lstErr 
         Height          =   2940
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6675
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Error List :"
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
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Receive :"
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
      Left            =   13740
      TabIndex        =   21
      Top             =   150
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Send : "
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
      Left            =   12705
      TabIndex        =   20
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Port : "
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
      Left            =   11670
      TabIndex        =   19
      Top             =   150
      Width           =   615
   End
   Begin VB.Image imgReceive 
      Height          =   240
      Left            =   14730
      Picture         =   "frmComm.frx":3990
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgSend 
      Height          =   240
      Left            =   13410
      Picture         =   "frmComm.frx":3F1A
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgPort 
      Height          =   240
      Left            =   12270
      Picture         =   "frmComm.frx":44A4
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11490
      TabIndex        =   18
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label lblSubMenu 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Axsym  Interface"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   135
      TabIndex        =   16
      Top             =   90
      Width           =   2460
   End
   Begin VB.Label picSubMenu 
      Appearance      =   0  '평면
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   90
      TabIndex        =   17
      Top             =   30
      Width           =   2565
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "순서"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "등록번호"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "성  명"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "검체번호"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "검체번호"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "상 태"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "검사항목"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult
Attribute Result1.VB_VarHelpID = -1

Private mAdoRs      As ADODB.Recordset
Private mAdoRs1     As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String
Private f_strJOB_FLAG   As String
Private f_strOrdList    As String
Private f_intSampleNo   As Integer

Private f_blnWorkList   As Boolean
Private f_lngWork_Row   As Long
Dim ReceiveData      As String
Dim SendFlg          As Boolean
Dim Patiant_Recevid As Integer

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Private Type typeNOVA
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(100)   As String
    Result(100)   As String
    Status(100)   As String
    Rerun(100)    As String
End Type

Dim NOVA As typeNOVA
Dim flgETB As Boolean
Dim fAxsym(100) As String
Dim fAxsymCfg(100) As Integer
Dim fAxsymSize(100, 1) As Integer

Dim fChannel() As String

Dim SeqNo As String

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
End Type

Private f_typCode() As TYPE_CD

Dim OrderCnt As Integer
Dim SendCount As Integer

Dim CountTest As Integer, sErrorFlag As Boolean
Dim cntCheckSum      As Integer
Dim flgETX           As Boolean

Private Type typeAXSYM
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(100)   As String
    Result(100)   As String
    Status(100)   As String
    Rerun(100)    As String
End Type

Dim AXSYM As typeAXSYM

Const Field_      As String = "|"
Const Repeat_     As String = "\"
Const Component_  As String = "^"
Const Escape_     As String = "&"
Const Slash_      As String = "/"
Dim cntField_     As Integer '|
Dim cntRepeat_    As Integer '\
Dim cntComponent_ As Integer '^
Dim cntEscape_    As Integer '&
Dim cntSlash_     As Integer '/

Dim fAXSYM_1(100) As String
'Dim SendData(10)     As String
Dim SendData     As String
Dim HostOutput       As String

Dim phase  As Integer
Dim bufcnt As Integer
Dim state  As String
Dim SndPhase As Integer
Dim FrameNo As Integer

Private strRcvbufR As String

'------------------
Dim cInterface As New clsIInterface ' Interface Class

Private Function f_funGet_ConvertResult(ByVal strRstval As String) As String

    Dim intPos  As Integer
    Dim strTmp1 As String, strTmp2  As String
    
    intPos = InStr(strRstval, "E")
    If intPos > 0 Then
        strTmp1 = Mid$(strRstval, 1, intPos - 1)
        strTmp2 = Mid$(strRstval, intPos + 1)
        
        If Mid$(strTmp2, 1, 1) = "-" Then
            f_funGet_ConvertResult = Round(Val(strTmp1) * (0.1 ^ Val(Mid$(strTmp2, 2))), 2)
        Else
            f_funGet_ConvertResult = Round(Val(strTmp1) * (10 ^ Val(Mid$(strTmp2, 2))), 2)
        End If
    Else
        f_funGet_ConvertResult = strRstval
    End If
    
End Function

Private Function MakeCS(Source As String) As String
    Dim x      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    
    For x = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, x, 1))
    Next x
    AddCS = AddCS + Asc(Chr(13)) + Asc(ETX)
    AddCS = AddCS Mod &H100
    SumCS = Hex(AddCS)
    If Len(SumCS) = 1 Then
        ChkCS = "0" & SumCS
    Else
        ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
        ChkCS = ChkCS & Right(SumCS, 1)
    End If
    MakeCS = ChkCS
End Function

Private Function f_funGet_SpreadRow(ByVal objSpd As fpSpread, ByVal intCol As Integer, _
                                    ByVal strPara As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For intRow = 1 To .MaxRows
            .GetText intCol, intRow, varTmp
            If Trim$(varTmp) = strPara Then
                f_funGet_SpreadRow = intRow
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub f_subSet_ComCharacter()

    MSG_STX = Chr(COM_STX)
    MSG_ETX = Chr(COM_ETX)
    MSG_ENQ = Chr(COM_ENQ)
    MSG_EOT = Chr(COM_EOT)
    MSG_ACK = Chr(COM_ACK)
    MSG_NAK = Chr(COM_NACK)
    MSG_CR = Chr(COM_CR)
    MSG_LF = Chr(COM_LF)
    MSG_CRLF = Chr(COM_CR) & Chr(COM_LF)
    
End Sub


Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intRow  As Integer, intCnt  As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear
    
    intRow = 1
    intCol = 9
    intCol2 = 1
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdRstDetail
        .MaxRows = 10
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFLM, REFHM, REFLF, REFHF, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") " & _
             "   and ((TESTCD <> '') and (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst:        ReDim fChannel(adoRS.RecordCount + intCol)
    
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFLM") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFHM") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("REFLF") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REFHF") & "")
            itemX.SubItems(12) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(13) = Trim(adoRS.Fields("REMARK") & "")
            itemX.Tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdRstDetail
            If intRow > .MaxRows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(adoRS("TESTNM") & "")
            intRow = intRow + 1
            
        End With
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"
        f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
        f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = strTmp
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    
    Set adoRS = Nothing

Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub ChkError_Click()
    If ChkError.Value = vbChecked Then
        FrameInterface.Visible = False
        FrameError.Visible = True
    Else
        FrameInterface.Visible = True
        FrameError.Visible = False
    End If
End Sub

Private Sub cmdACK_Click()
    Dim RecData As String
    Dim sRs As Object
    
'    'Call f_subSet_WorkList
'    'Call f_subSet_TestList("10124")
'
'    Set sRs = f_subSet_TestList("10124")
'
'    If Not sRs.EOF Then
'        Debug.Print Trim(sRs("검체번호")) & ""
'        Debug.Print Trim(sRs("품목코드")) & ""
'    End If
'
'    Exit Sub
'    'Call COM_OUTPUT(ACK)
'          RecData = "|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|2004031714235384"
'RecData = RecData & "1H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|2004031714255083" & vbCr
'RecData = RecData & "2Q|1|^115596||^^^ALL||||||||O00" & vbCr
'RecData = RecData & "3L|13C" & vbCr
'RecData = RecData & "" & vbCr

    RecData = RecData & "1H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101123" & vbCr
    RecData = RecData & "7A" & vbCr
    RecData = RecData & "2P|1||||" & vbCr
    RecData = RecData & "2F" & vbCr
    RecData = RecData & "3O|1||1^E^01|^^^817^HIV_Ag_Ab^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "1E" & vbCr
    RecData = RecData & "4R|1|^^^817^HIV_Ag_Ab^UNDILUTED^^F|0.55|S/CO||||R||FSE||20031110180612" & vbCr
    RecData = RecData & "73" & vbCr
    RecData = RecData & "5R|2|^^^817^HIV_Ag_Ab^UNDILUTED^^P|25.53|Rate||||R||FSE||20031110180612" & vbCr
    RecData = RecData & "2C" & vbCr
    RecData = RecData & "6R|3|^^^817^HIV_Ag_Ab^UNDILUTED^^I|NEGATIVE|||||R||FSE||20031110180612" & vbCr
    RecData = RecData & "F1" & vbCr
    RecData = RecData & "7L|1" & vbCr
    RecData = RecData & "40" & vbCr
    RecData = RecData & "0H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101125" & vbCr
    RecData = RecData & "7B" & vbCr
    RecData = RecData & "1P|1||||" & vbCr
    RecData = RecData & "2E" & vbCr
    RecData = RecData & "2O|1||3^E^03|^^^106^HBsAg^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "CD" & vbCr
    RecData = RecData & "3R|1|^^^106^HBsAg^UNDILUTED^^F|0.98|S/N||||F||FSE||20031110180414" & vbCr
    RecData = RecData & "D5" & vbCr
    RecData = RecData & "4R|2|^^^106^HBsAg^UNDILUTED^^P|5.37|Rate||||F||FSE||20031110180414" & vbCr
    RecData = RecData & "9B" & vbCr
    RecData = RecData & "5R|3|^^^106^HBsAg^UNDILUTED^^I|NEGATIVE|||||F||FSE||20031110180414" & vbCr
    RecData = RecData & "90" & vbCr
    RecData = RecData & "6L|1" & vbCr
    RecData = RecData & "3F" & vbCr
    RecData = RecData & "7H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101125" & vbCr
    RecData = RecData & "82" & vbCr
    RecData = RecData & "0P|1||||" & vbCr
    RecData = RecData & "2D" & vbCr
    RecData = RecData & "1O|1||3^E^03|^^^118^AUSAB^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "96" & vbCr
    RecData = RecData & "2R|1|^^^118^AUSAB^UNDILUTED^^F|117.0|mIU/mL||||F||FSE||20031110180930" & vbCr
    RecData = RecData & "EC" & vbCr
    RecData = RecData & "3R|2|^^^118^AUSAB^UNDILUTED^^P|1032.63|Rate||||F||FSE||20031110180930" & vbCr
    RecData = RecData & "F7" & vbCr
    RecData = RecData & "4R|3|^^^118^AUSAB^UNDILUTED^^I|REACTIVE|||||F||FSE||20031110180930" & vbCr
    RecData = RecData & "5C" & vbCr
    RecData = RecData & "5L|1" & vbCr
    RecData = RecData & "3E" & vbCr
    RecData = RecData & "6H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101126" & vbCr
    RecData = RecData & "82" & vbCr
    RecData = RecData & "7P|1||||" & vbCr
    RecData = RecData & "34" & vbCr
    RecData = RecData & "0O|1||2^E^02|^^^208^hTSH_II^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "6F" & vbCr
    RecData = RecData & "1R|1|^^^208^hTSH_II^UNDILUTED^^F|69.250|uIU/mL||||R||FSE||20031110175558" & vbCr
    RecData = RecData & "22" & vbCr
    RecData = RecData & "2R|2|^^^208^hTSH_II^UNDILUTED^^P|1186.83|Rate||||R||FSE||20031110175558" & vbCr
    RecData = RecData & "F4" & vbCr
    RecData = RecData & "3L|1" & vbCr
    RecData = RecData & "3C" & vbCr
    RecData = RecData & "4H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101126" & vbCr
    RecData = RecData & "80" & vbCr
    RecData = RecData & "5P|1||||" & vbCr
    RecData = RecData & "32" & vbCr
    RecData = RecData & "6O|1||5^E^05|^^^208^hTSH_II^UNDILUTED|S||||||||||||||||||||F" & vbCr
    RecData = RecData & "7C" & vbCr
    RecData = RecData & "7R|1|^^^208^hTSH_II^UNDILUTED^^F|1.385|uIU/mL||||F||FSE||20031110171945" & vbCr
    RecData = RecData & "E3" & vbCr
    RecData = RecData & "0R|2|^^^208^hTSH_II^UNDILUTED^^P|54.92|Rate||||F||FSE||20031110171945" & vbCr
    RecData = RecData & "7B" & vbCr
    RecData = RecData & "1L|1" & vbCr
    RecData = RecData & "3A" & vbCr
    RecData = RecData & "2H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101127" & vbCr
    RecData = RecData & "7F" & vbCr
    RecData = RecData & "3P|1||||" & vbCr
    RecData = RecData & "30" & vbCr
    RecData = RecData & "4O|1||4^E^04|^^^106^HBsAg^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "D1" & vbCr
    RecData = RecData & "5R|1|^^^106^HBsAg^UNDILUTED^^F|0.96|S/N||||F||FSE||20031110180641" & vbCr
    RecData = RecData & "D7" & vbCr
    RecData = RecData & "6R|2|^^^106^HBsAg^UNDILUTED^^P|5.29|Rate||||F||FSE||20031110180641" & vbCr
    RecData = RecData & "A0" & vbCr
    RecData = RecData & "7R|3|^^^106^HBsAg^UNDILUTED^^I|NEGATIVE|||||F||FSE||20031110180641" & vbCr
    RecData = RecData & "94" & vbCr
    RecData = RecData & "0L|1" & vbCr
    RecData = RecData & "39" & vbCr
    RecData = RecData & "1H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101127" & vbCr
    RecData = RecData & "7E" & vbCr
    RecData = RecData & "2P|1||||" & vbCr
    RecData = RecData & "2F" & vbCr
    RecData = RecData & "3O|1||4^E^04|^^^118^AUSAB^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "9A" & vbCr
    RecData = RecData & "4R|1|^^^118^AUSAB^UNDILUTED^^F|0.2|mIU/mL||||F||FSE||20031110181234" & vbCr
    RecData = RecData & "85" & vbCr
    RecData = RecData & "5R|2|^^^118^AUSAB^UNDILUTED^^P|23.78|Rate||||F||FSE||20031110181234" & vbCr
    RecData = RecData & "9C" & vbCr
    RecData = RecData & "6R|3|^^^118^AUSAB^UNDILUTED^^I|NONREACTIVE|||||F||FSE||20031110181234" & vbCr
    RecData = RecData & "47" & vbCr
    RecData = RecData & "7L|1" & vbCr
    RecData = RecData & "40" & vbCr
    RecData = RecData & "0H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101128" & vbCr
    RecData = RecData & "7E" & vbCr
    RecData = RecData & "1P|1||||" & vbCr
    RecData = RecData & "2E" & vbCr
    RecData = RecData & "2O|1||3^E^03|^^^419^CEA^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "F8" & vbCr
    RecData = RecData & "3R|1|^^^419^CEA^UNDILUTED^^F|14.9|ng/mL||||F||FSE||20031110180747" & vbCr
    RecData = RecData & "F3" & vbCr
    RecData = RecData & "4R|2|^^^419^CEA^UNDILUTED^^P|123.87|Rate||||F||FSE||20031110180747" & vbCr
    RecData = RecData & "35" & vbCr
    RecData = RecData & "5L|1" & vbCr
    RecData = RecData & "3E" & vbCr
    RecData = RecData & "6H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101128" & vbCr
    RecData = RecData & "84" & vbCr
    RecData = RecData & "7P|1||||" & vbCr
    RecData = RecData & "34" & vbCr
    RecData = RecData & "0O|1||6^E^06|^^^106^HBsAg^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "D1" & vbCr
    RecData = RecData & "1R|1|^^^106^HBsAg^UNDILUTED^^F|0.90|S/N||||F||FSE||20031110180858" & vbCr
    RecData = RecData & "D7" & vbCr
    RecData = RecData & "2R|2|^^^106^HBsAg^UNDILUTED^^P|4.94|Rate||||F||FSE||20031110180858" & vbCr
    RecData = RecData & "A7" & vbCr
    RecData = RecData & "3R|3|^^^106^HBsAg^UNDILUTED^^I|NEGATIVE|||||F||FSE||20031110180858" & vbCr
    RecData = RecData & "9A" & vbCr
    RecData = RecData & "4L|1" & vbCr
    RecData = RecData & "3D" & vbCr
    RecData = RecData & "5H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101129" & vbCr
    RecData = RecData & "84" & vbCr
    RecData = RecData & "6P|1||||" & vbCr
    RecData = RecData & "33" & vbCr
    RecData = RecData & "7O|1||6^E^06|^^^419^CEA^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "03" & vbCr
    RecData = RecData & "0R|1|^^^419^CEA^UNDILUTED^^F|4.4|ng/mL||||F||FSE||20031110181017" & vbCr
    RecData = RecData & "B1" & vbCr
    RecData = RecData & "1R|2|^^^419^CEA^UNDILUTED^^P|39.76|Rate||||F||FSE||20031110181017" & vbCr
    RecData = RecData & "FD" & vbCr
    RecData = RecData & "2L|1" & vbCr
    RecData = RecData & "3B" & vbCr
    RecData = RecData & "3H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|20031128101129" & vbCr
    RecData = RecData & "82" & vbCr
    RecData = RecData & "4P|1||||" & vbCr
    RecData = RecData & "31" & vbCr
    RecData = RecData & "5O|1||6^E^06|^^^118^AUSAB^UNDILUTED|R||||||||||||||||||||F" & vbCr
    RecData = RecData & "A0" & vbCr
    RecData = RecData & "6R|1|^^^118^AUSAB^UNDILUTED^^F|>1000.0|mIU/mL||>||F||FSE||20031110181329" & vbCr
    RecData = RecData & "97" & vbCr
    RecData = RecData & "7R|2|^^^118^AUSAB^UNDILUTED^^P|3151.87|Rate||||F||FSE||20031110181329" & vbCr
    RecData = RecData & "08" & vbCr
    RecData = RecData & "0R|3|^^^118^AUSAB^UNDILUTED^^I|REACTIVE|||||F||FSE||20031110181329" & vbCr
    RecData = RecData & "5B" & vbCr
    RecData = RecData & "1L|1" & vbCr
    RecData = RecData & "3A" & vbCr


    Call ComReceive(RecData)
    
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdSave
        Case 4:     Call Form_QueryUnload(0, 0)
        Case Else
    End Select

End Sub

Private Sub cmdSave()
    Dim intRow, intCol As Integer
    Dim varTmp
    Dim itemX As ListItem
    Dim Channel_No  As String
    Dim strRstval As String
    Dim strTestcd As String
    Dim varBarno, varSPnm, varSPid, varZation, varRegNo, varRegDt
    Dim varSex As String, varAge As String, varRef As String
    Dim sqlDoc As String
    Dim sqlRet As Integer
    
    With spdResult1
        For intRow = 1 To .MaxRows
            For intCol = 9 To .MaxCols
                .GetText intCol, 0, varTmp
                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                If Not itemX Is Nothing Then
                    Channel_No = itemX.Tag
                    strTestcd = itemX.ListSubItems(1)
                    
                    .GetText 1, intRow, varBarno
                    .GetText 4, intRow, varSPnm
                    .GetText 5, intRow, varSPid '--성별/나이
                    .GetText 6, intRow, varZation
                    If varZation = "입원" Then
                        varZation = "1"
                    ElseIf varZation = "외래" Then
                        varZation = "2"
                    Else
                        varZation = "3"
                    End If
                    .GetText 7, intRow, varRegNo
                    .GetText 8, intRow, varRegDt
                    varSex = Mid(varSPid, 1, 1)
                    varAge = Mid(varSPid, 3)
                    varRef = ""
                    .GetText intCol, intRow, varTmp
                    strRstval = varTmp
                    
                    '-- Channel_No = "106", "118", "841", "817" : 문자형결과
                    If Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817" Then
                        If varSex = "남" Then
                            If Val(strRstval) < Val(itemX.ListSubItems(8)) Then
                                varRef = "L"
                            ElseIf Val(strRstval) > Val(itemX.ListSubItems(9)) Then
                                varRef = "H"
                            End If
                        Else
                            If Val(strRstval) < Val(itemX.ListSubItems(10)) Then
                                varRef = "L"
                            ElseIf Val(strRstval) > Val(itemX.ListSubItems(11)) Then
                                varRef = "H"
                            End If
                        End If
                    End If
                        
                    '-- Local등록
                    If chkMdbSave.Value = 1 And Trim(strRstval) <> "" Then
                        sqlDoc = "Update INTERFACE003" & _
                                 "   set RESULT1  = '" & strRstval & "', REFERENCE = '" & varRef & "'" & _
                                 " where SPCNO   = '" & varBarno & "'" & _
                                 "   and TESTCD  = '" & itemX.Text & "'" & _
                                 "   and REGDATE = '" & varRegDt & "'"

                        AdoCn_Jet.Execute sqlDoc, sqlRet
    
                        If sqlRet = 0 Then
                            sqlDoc = "insert into INTERFACE003(" & _
                                     "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
                                     "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
                                     "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
                                     "            '" & strRstval & "','','', '" & varRef & "','',''," & _
                                     "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
                            AdoCn_Jet.Execute sqlDoc
                        End If
                    End If
                    
                    '-- Server등록
                    If chkSvrSave.Value = 1 Then
                    
                    End If
                End If
                                        
                Set itemX = Nothing
            Next intCol
        Next intRow
    End With

End Sub

Private Sub cmdClear()
    
    txtCom.Text = ""
    lstErr.Clear
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

End Sub

Private Sub cmdExit()
    
    If frmComm.comEQP.PortOpen = True Then
        If MsgBox("인터페이스중입니다." & Chr(10) & _
               "작업을 종료하면 받고있거나 검사중인 데이터를 잃게 됩니다" & Chr(10) & _
               "종료하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            
            Unload Me
            
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "장비와 통신연결 성공.."
        cmdAction(2).Enabled = False
        cmdAction(3).Enabled = False
        cmdAction(4).Enabled = False
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
        cmdAction(2).Enabled = True
        cmdAction(3).Enabled = True
        cmdAction(4).Enabled = True
    End If
        
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdStop()
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If comEQP.PortOpen Then comEQP.PortOpen = False
    If comEQP.PortOpen Then
        Call ShowMessage("중지 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
        cmdAction(2).Enabled = False
        cmdAction(3).Enabled = False
        cmdAction(4).Enabled = False
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
        cmdAction(2).Enabled = True
        cmdAction(3).Enabled = True
        cmdAction(4).Enabled = True
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub CmdErrClear_Click()
    lstErr.Clear
End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim fRcvString As String
    
    Dim Buffer As String
    Dim iBufLen As Integer
    Dim BufChar As String
    Dim I%

    On Error GoTo ErrorTrap
    
    CallForm = "frmInterface - Private Sub comEQP_OnComm()"
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            
            Buffer = comEQP.Input
            iBufLen = Len(Buffer)
            For I = 1 To iBufLen
                BufChar = Mid(Buffer, I, 1)
                Select Case cInterface.phase
                    Case 1                  ' ENQ 대기
                        Select Case Asc(BufChar)
                            Case 5          ' ENQ
                                Call COM_OUTPUT(Chr(6))
                                cInterface.phase = 2
                                cInterface.bufcnt = 1
                       End Select
                    Case 2                  ' LF 대기
                        Select Case Asc(BufChar)
                            Case 2          ' STX
                                Call cInterface.clearRcvbuf
                            Case 4          ' EOT
                                If cInterface.state = "Q" Then
                                    Call COM_OUTPUT(Chr(5))
                                    cInterface.Snd_Phase = 1
                                    cInterface.FrameN = 1
                                End If
                                cInterface.phase = 3
                            Case 10        ' LF
                                Call psDataDefine(Buffer, fChannel(), spdResult1)
                                Call COM_OUTPUT(Chr(6))
                                cInterface.phase = 2
                            Case Else
                                Call cInterface.addRcvbuf(BufChar)
                        End Select
                    Case 3                  ' ACK 대기
                        Select Case Asc(BufChar)
                            Case 6
                                If cInterface.state = "Q" Then
                                    Call SendOrdData
                                End If
                            Case 5
                                Call COM_OUTPUT(Chr(6))
                                cInterface.phase = 2
                            Case 21
                                Call COM_OUTPUT(Chr(5))
                            Case 4
                                cInterface.phase = 1
                        End Select
                End Select
            Next
        
        Case comEvSend
        
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) 변경 감지"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) 변경 감지"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) 변경 감지"
        Case comEvRing
            strEVMsg = " 전화 벨이 울리는 중"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) 감지"

        ' 오류 메시지
        Case comBreak
            strERMsg = " 중단 신호 수신"
        Case comCDTO
            strERMsg = " 반송파 검출 시간 초과"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) 시간 초과"
        Case comDCB
            strERMsg = " 포트에 대한 장치 제어 블록(DCB) 검색 중 예기치 못한 오류"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) 시간 초과"
        Case comFrame
            strERMsg = " 프레이밍 오류"
        Case comOverrun
            strERMsg = " 패리티 오류"
        Case comRxOver
            strERMsg = " 수신 버퍼 초과"
        Case comRxParity
            strERMsg = " 패리티 오류"
        Case comTxFull
            strERMsg = " 전송 버퍼에 여유가 없음"
        Case Else
            strERMsg = " 알 수 없는 오류 또는 이벤트"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
    
    Exit Sub
ErrorTrap:
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub SendOrdData()
    Dim sSndBuf As String
    Dim ChkS    As String
    Dim LabDate As String
    Dim TestDat As String
    Dim I       As Integer
    Dim sSampleNo As String
    Dim sRs As Object
    Dim intRow  As Integer
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim strOrder As String
    Dim sndOrder As String
    
    On Error GoTo ErrorTrap
    
    CallForm = "frmInterface - Private Sub SendOrdData()"
    
    strOrder = ""
    
    Select Case cInterface.Snd_Phase
        Case 1      ' Header Record
            Debug.Print "----> H"
            sSndBuf = cInterface.FrameN & "H|\^&||||||||||P|1|" & vbCr & Chr(3)
            cInterface.Snd_Phase = 2
        Case 2      ' Patient Record
            Debug.Print "----> P"
            sSndBuf = cInterface.FrameN & "P|1||" & Trim$(AXSYM.SampleNo) & "||" & vbCr & Chr(3)
            cInterface.Snd_Phase = 3
        Case 3      ' Order Record
            Debug.Print "----> O"
            sSampleNo = AXSYM.SampleNo
        
            With spdResult1
                '-- 오더 리스트에 있는 검체를 다시 스캔했을 경우..
                If SeqSearch(spdResult1, Trim(sSampleNo), 2) > 0 Then
                    lstErr.AddItem "검체번호 " & sSampleNo & "은(는) 이미 오더리스트에 있습니다."
                    Exit Sub
                End If
                
                Set sRs = f_subSet_TestList(Trim(sSampleNo))
                
                If Not sRs.EOF Then
                    .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                    intRow = SeqNullSearch(spdResult1, sRs("검체번호"), 1)
                    
                    .SetText 1, intRow, Trim(sRs("검체번호")) & ""
                    .SetText 4, intRow, Trim(sRs("성명")) & ""
                    .SetText 5, intRow, IIf(Trim(sRs("성별코드")) = "M", "남", "여") & "/" & Format(Now, "yyyy") - Format(Trim(sRs("생년월일")), "yyyy") - 1
                    If sRs("입외구분") = "1" Then
                        .SetText 6, intRow, "입원"
                    ElseIf sRs("입외구분") = "2" Then
                        .SetText 6, intRow, "외래"
                    Else
                        .SetText 6, intRow, "퇴원"
                    End If
                    .SetText 7, intRow, Trim(sRs("병록번호")) & ""
                    .SetText 8, intRow, Trim(sRs("접수일자")) & ""
                    '-- 검사항목조회
                    Do Until sRs.EOF
                        strEqpCd = f_funGet_CODE(Trim(sRs("품목코드")) & "")
'                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                        
                        If Not itemX Is Nothing Then
                            spdResult1.Row = intRow
                            spdResult1.Col = itemX.Index + 8
                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
                            DoEvents
                        
                            If strOrder = "" Then
                                strOrder = "^^^" & Trim(itemX.Tag)
                            Else
                                strOrder = strOrder & "\^^^" & Trim(itemX.Tag)
                            End If
                        
                        End If
                        
                        sRs.MoveNext
                    Loop
                    sSndBuf = cInterface.FrameN & "O|1|" & Trim$(AXSYM.SampleNo) & "||" & strOrder & "|R||||||N||||||||||||||Q" & vbCr & Chr(3)
                    'Debug.Print tmp
                End If
            End With
            cInterface.Snd_Phase = 4
        Case 4      'Terminator Record
            Debug.Print "----> L"
            sSndBuf = cInterface.FrameN & "L|1" & vbCr & Chr(3)
            cInterface.Snd_Phase = 5
            Debug.Print sSndBuf
        Case 5      ' EOT
            Debug.Print "----> EOT"
            'comEQP.Output = Chr(4)   'EOT
            Call COM_OUTPUT(Chr(4))
            cInterface.FrameN = 1
            cInterface.phase = 1
            cInterface.Snd_Phase = 1
            cInterface.state = ""
            Exit Sub
    End Select
    
    ChkS = getChkSum(sSndBuf)
    sndOrder = Chr(2) & sSndBuf & ChkS & vbCr & vbLf
    Call COM_OUTPUT(sndOrder)
    cInterface.addFrameN
        
    Exit Sub
ErrorTrap:
    Call ErrMsgProc(CallForm)

End Sub

Private Sub ComReceive(ByRef RecData As String)

    Dim sStxCheck As Integer, sEnqCheck As Integer, sEtxCheck As Integer
    Dim sLfCheck As Integer, sCrcheck As Integer, ii As Integer
    Dim MHead As String, Pinfo As String, OutputData As String, com_sTemp As String
    Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    
    Static OrgMsg As String

    strRec = RecData
    Print #1, strRec;
    Call COM_INPUT(strRec)
'    Debug.Print strRec
    
    For ii = 1 To Len(strRec)
        Select Case Mid(strRec, ii, 1)
            Case STX:
                    ii = ii + 1 'Frame Number
                    If sErrorFlag Then
                        sErrorFlag = False
                        cntCheckSum = 2
                    Else
                        cntCheckSum = 0
                    End If
            Case ETX:
                    cntCheckSum = cntCheckSum + 1
                    Call COM_OUTPUT(ACK)
'                        Debug.Print "[HOST] " & ACK
                    flgETX = True
            Case ETB:
                    If Mid(ReceiveData, ii, 2) = vbCr & vbLf Then
                        ReceiveData = left(ReceiveData, Len(ReceiveData) - 2) 'Remove CR & LF
                    End If
                    cntCheckSum = cntCheckSum + 1
                    Call COM_OUTPUT(ACK)
'                        Debug.Print "[HOST] " & ACK
                    flgETB = True
                    sErrorFlag = True
            Case vbCr:
                    If flgETB = True Then
                        flgETB = False
                    Else
'                            ReceiveTheDataAXSYM
                        '---------------------------------------------
'                        Dim sTxfile As String
'
'                        sTxfile = App.Path & "\" & Format(Now, "yyyyMMdd") & ".LOG"
'                        If Len(Dir(sTxfile)) = 0 Then
'                            Open sTxfile For Output As #1
'                            Close #1
'                        End If
'                        Open sTxfile For Append As #1
'                            Print #1, "RCV=> "; ReceiveData
'                        Close #1
                        '---------------------------------------------
                        Call psDataDefine(ReceiveData, fChannel(), spdResult1)
                        GoSub ClearReceiveData
                    End If
            Case vbLf:
                '
            Case ENQ:
                    Call COM_OUTPUT(ACK)
            Case ACK:
                    If SendFlg = True Then
'                            SendTest
                        Call SendTest(ReceiveData, fChannel(), spdResult1)
                    Else
                        Call COM_OUTPUT(EOT)
'                           Debug.Print "[HOST] " & EOT
                        AXSYM.SID = ""
                    End If
            Case NAK:
                    If AXSYM.SID <> "" Then
                        SendCount = SendCount - 1
                        Call SendTest(ReceiveData, fChannel(), spdResult1)
                    Else
                        Call COM_OUTPUT(EOT)
                    End If
            Case EOT:
                    Call COM_OUTPUT(ENQ)
                    cntCheckSum = 0
                    GoSub ClearReceiveData
            Case Else:
                Select Case cntCheckSum
                    Case 1:
                        cntCheckSum = cntCheckSum + 1
                    Case 2:
                        cntCheckSum = 0
                    Case Else:
                        ReceiveData = ReceiveData & Mid(strRec, ii, 1)
                End Select
                
        End Select
    Next ii
    
    Exit Sub
    
ClearReceiveData:
    ReceiveData = ""
    cntField_ = 0
    cntRepeat_ = 0
    cntComponent_ = 0
    cntEscape_ = 0
    cntSlash_ = 0
    Return
    
errOnComm:
    Exit Sub
            
End Sub


Private Sub SendTest(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)
    If SendCount <= 0 Then
'        SendTheSample
        Call SendTheSample(ReceiveData, fChannel(), spdResult1)
    Else
        Call COM_OUTPUT(SendData)
        SendCount = SendCount - 1
        If SendCount = 0 Then
           SendFlg = False
        Else
           SendFlg = True
        End If
    End If
End Sub

Private Sub SendTheSample(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)
    On Error GoTo SendTheSampleSub_
    Dim sRs As Object
    Dim Loop_Count As Integer
    Dim FunStr1 As String
    Dim PatientID As String
    Dim PatientNo As String
    Dim ii As Integer, sDeCnt    As Integer
    Dim Testcd As String
    Dim OutputData As String, sOrderLst As String
    Dim EndStr, strEqpCd, sChannel, strOrdLst As String
    Dim intRow  As Integer
    Dim itemX As ListItem
    Dim sHead As String, sPInfo As String, sOrder As String, sLast As String
    Dim sTmp  As String
    Dim sSampleNo  As String
    Dim strOrder As String
    Dim ChkS    As String
    
    On Error GoTo errSend
    
    strOrder = ""
    
    Select Case SendCount
        Case 0      ' Header Record
'            tmp = objInt.FrameNo & "H|\^&|||AxSYM^3.60^1180^H1P1O1R1C1Q1L1M1|||||||P|1|" & Format(DBConn.getSysDate, "yyyyMMddhhMMss") & Chr(13) & Chr(3)
            sTmp = "1H|\^&||||||||||P|1|" & vbCr & Chr(3)
            SendCount = 1
        Case 1      ' Patient Record
'            tmp = objInt.FrameNo & "P|1||" & Trim$(objAxsym.SpcYY & objAxsym.SpcNo) & "||" & vbCr & Chr(3)
            sTmp = "1P|1||" & Trim$(AXSYM.SampleNo) & "||" & Chr(13) & Chr(3)
            SendCount = 2
        Case 2      ' Order Record
            
            sSampleNo = AXSYM.SampleNo
            
            With spdResult1
                Set sRs = f_subSet_TestList(Trim(sSampleNo))
                If Not sRs.EOF Then
                    .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                    intRow = SeqNullSearch(spdResult1, sRs("검체번호"), 1)
                    
                    .SetText 1, intRow, Trim(sRs("검체번호")) & ""
                    .SetText 4, intRow, Trim(sRs("성명")) & ""
                    .SetText 5, intRow, IIf(Trim(sRs("성별코드")) = "M", "남", "여") & "/" & Format(Now, "yyyy") - Format(Trim(sRs("생년월일")), "yyyy") - 1
                    If sRs("입외구분") = "1" Then
                        .SetText 6, intRow, "입원"
                    ElseIf sRs("입외구분") = "2" Then
                        .SetText 6, intRow, "외래"
                    Else
                        .SetText 6, intRow, "퇴원"
                    End If
                    .SetText 7, intRow, Trim(sRs("병록번호")) & ""
                    .SetText 8, intRow, Trim(sRs("접수일자")) & ""
                    '-- 검사항목조회
                    Do Until sRs.EOF
                        strEqpCd = f_funGet_CODE(Trim(sRs("품목코드")))
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
        
                        If Not itemX Is Nothing Then
                            spdResult1.Row = intRow
                            spdResult1.Col = itemX.Index + 8
                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
                            DoEvents
                        End If
                        
                        If strOrder = "" Then
                            strOrder = "^^^" & Trim(itemX.Tag)
                        Else
                            strOrder = strOrder & "\^^^" & Trim(itemX.Tag)
                        End If
                        sRs.MoveNext
                    Loop
                End If
            End With
    
            'sOrder = "1O|1|" & mvarSpcYY & mvarSpcNo & "||" & strOrder & "|R||||||N||||||||||||||Q" & vbCr & Chr(3)
            sTmp = "1O|1|" & AXSYM.SampleNo & "||" & strOrder & "|R||||||N||||||||||||||Q" & Chr(13) & Chr(3)
            SendCount = 3
        Case 3      'Terminator Record
            sTmp = "1L|1" & vbCr & Chr(3)
            SendCount = 4
        Case 4      ' EOT
            comEQP.Output = Chr(4)   'EOT
            SendCount = 0
            Exit Sub
    End Select
    
    
    lblStatus.Caption = "Order 전송 중.."
    ChkS = getChkSum(sTmp)
    'comEQP.Output = Chr(2) & tmp & ChkS & Chr(13) & Chr(10)
    
    SendData = Chr(2) & sTmp & ChkS & Chr(13) & Chr(10)
    
    'SendCount = Int((Len(SendData) / 230)) + 1
    
    Call COM_OUTPUT(SendData)
    Debug.Print SendData
    
    'SendCount = SendCount - 1
    If SendCount = 0 Then
       SendFlg = False
    Else
       SendFlg = True
    End If
    Exit Sub
    
SendTheSampleSub_:
    Call COM_OUTPUT(EOT)
    Exit Sub
    
errSend:

End Sub

Public Function getChkSum(sMsg As String) As String
    Dim I%
    Dim iChkSum As Integer
    
    iChkSum = 0
    For I = 1 To Len(sMsg)
        iChkSum = (iChkSum + Asc(Mid(sMsg, I, 1)))
    Next
    iChkSum = iChkSum Mod 256
    getChkSum = Right("0" & Hex(iChkSum), 2)
    
End Function

Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long

Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Val(Trim(.Text)) = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With
    
End Function

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    Dim ii As Long
    Dim jj As Long
    Dim KK As Long
    Dim Found As Boolean
    Dim SID As String
    Dim SEX As String
    Dim AGE As Long
    Dim sql As String
    Dim OutputData As String, sOrderLst As String
    Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
    Dim Channel_No  As String       ' 문자형 변수
    Dim Patiant_No  As String       ' 환자번호
    Dim pGrid_Point As Integer      ' 해당 검사자 Point
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    '-------------------------------' 임시 변수들.....
    Dim sDeCnt      As Integer
    Dim pDoCount    As Integer
    Dim Loop_Count  As Integer
    Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
    Dim sPatiant_No As Long
    
    Dim sSeq As String
    Dim sCol As Integer
    Dim iCnt As Integer
    Dim intRow As Integer, intCol As Integer, intIdx As Integer
    Dim varTmp
    Dim strRstval As String, strRefVal As String
    Dim tmpRstval As String
    Dim itemX  As ListItem
    Dim sqlDoc As String
    Dim sqlRet As Integer
    Dim strTestcd As String
    Dim varBarno, varSPnm, varSPid, varZation, varRegNo, varRegDt
    Dim varSex As String, varAge As String, varRef As String
    Dim blnFlag As Boolean
    Dim strBarNo As String
    Dim sRs As Object
    Dim CntsTxt As Integer
    Dim sTxt As String, tTxt As String

    On Error GoTo ErrorTrap
    
    CallForm = "frmInterface - Private Sub psDataDefine()"
    
    ReceiveData = strdata
    
    ' 결과버퍼 통합
    cInterface.addBufs
    
    Dim sIDCode$, rcvbufs$, sendbufs$, tmp$
    
    ReceiveData = cInterface.getrcvbufs
    
    'sIDCode = Mid(ReceiveData, 2, 1)
    
'    Debug.Print "받은데이타:" & ReceiveData
    Call COM_INPUT(ReceiveData)
    On Error GoTo errReceive
    For iCnt = 1 To Len(ReceiveData)
        Select Case Mid(ReceiveData, iCnt, 1)
            
            Case "H" 'Message Header
                GoSub Clear_AXSYM_
                Exit For
            Case "P" 'Patient Informatioin
                GoSub Clear_AXSYM_
                Exit For
            Case "O" 'Test Order
                sSeq = Trim(AXSYM.SID)
                sCol = 2
                
                'Patiant_Recevid = SeqSearch(spdResult1, sSeq, sCol)
                Erase fAxsym
            
                Do While InStr(ReceiveData, Chr$(124)) > 0 '--Chr(29)
                    pDoCount = pDoCount + 1
                    fAxsym(pDoCount) = Text_Redefine(ReceiveData, Chr$(124))
                    ReceiveData = Mid$(ReceiveData, InStr(ReceiveData, Chr$(124)) + 1)
                    If pDoCount > 99 Then
                        ReceiveData = ""
                        Exit Do
                    End If
                    If pDoCount < 10 Then fAxsym(pDoCount) = Text_Change(fAxsym(pDoCount), Chr$(29), "")
'                    Debug.Print fAxsym(pDoCount)
                Loop
                
                SeqNo = Mid(fAxsym(4), 1, InStr(fAxsym(4), "^") - 1)
                Exit For
    
            Case "R" 'Test Result
                pDoCount = 0
            
                Erase fAxsym
            
                Do While InStr(ReceiveData, Chr$(124)) > 0 '--Chr(29)
                    pDoCount = pDoCount + 1
                    fAxsym(pDoCount) = Text_Redefine(ReceiveData, Chr$(124))
                    ReceiveData = Mid$(ReceiveData, InStr(ReceiveData, Chr$(124)) + 1)
                    If pDoCount > 99 Then
                        ReceiveData = ""
                        Exit Do
                    End If
                    If pDoCount < 10 Then fAxsym(pDoCount) = Text_Change(fAxsym(pDoCount), Chr$(29), "")
                Loop
                
                Channel_No = Mid(fAxsym(3), 4, 3)
                '-- 숫자형 결과
                If Len(SeqNo) > 0 And Right(fAxsym(3), 1) = "F" Then 'And (Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817") Then
                    intRow = 0
                    With spdResult1
                        intRow = SeqSearch(brspread, SeqNo, 1)
                        '-- 해당번호 찾음
                        If intRow > 0 Then
                            For intCol = 9 To .MaxCols
                                .GetText intCol, 0, varTmp
                                Channel_No = Mid(fAxsym(3), 4, 3)
                                If Right(Channel_No, 1) = "^" Then Channel_No = Mid(Channel_No, 1, 2)
                                
                                If Right(fAxsym(3), 1) = "P" Then Exit Sub
                                
                                strRstval = Trim(fAxsym(4))
                                tmpRstval = Replace(strRstval, ">", "")
                                tmpRstval = Replace(tmpRstval, "<", "")
                                Select Case Channel_No
                                    Case "106" '-- Hbs Ag
                                        If Val(tmpRstval) >= 2 And Val(tmpRstval) <= 10 Then '-- Weekly Positive
                                            strRstval = "weakly Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) > 10 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) < 2 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "118" '-- Hbs Ab
                                        If Val(tmpRstval) >= 10 And Val(tmpRstval) <= 20 Then '-- WeeklyPositive
                                            strRstval = "weakly Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) > 20 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) < 10 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "841" '-- HCV
                                        If Val(tmpRstval) > 1 Then '-- Positive
                                            strRstval = "" '"Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "817" '-- Hiv
                                        If Val(tmpRstval) > 1 Then '-- Positive
                                            strRstval = ""
                                        ElseIf Val(tmpRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case Else
                                
                                End Select
                                
'                                strRstval = Mid(strRstval, 1, 20)
                                                                
                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                If Not itemX Is Nothing Then
                                    If Channel_No = itemX.Tag Then
                                        .SetText intCol, intRow, strRstval
                                    
                                        strTestcd = itemX.ListSubItems(1)
                                        
                                        .GetText 1, intRow, varBarno
                                        .GetText 4, intRow, varSPnm
                                        .GetText 5, intRow, varSPid '--성별/나이
                                        .GetText 6, intRow, varZation
                                        If varZation = "입원" Then
                                            varZation = "1"
                                        ElseIf varZation = "외래" Then
                                            varZation = "2"
                                        Else
                                            varZation = "3"
                                        End If
                                        .GetText 7, intRow, varRegNo
                                        .GetText 8, intRow, varRegDt
                                        varSex = Mid(varSPid, 1, 1)
                                        varAge = Mid(varSPid, 3)
                                        varRef = ""
                                        
                                        If Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817" Then
                                            If varSex = "남" Then
                                                If fAxsym(4) < Val(itemX.ListSubItems(8)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(9)) Then
                                                    varRef = "H"
                                                End If
                                            Else
                                                If fAxsym(4) < Val(itemX.ListSubItems(10)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(11)) Then
                                                    varRef = "H"
                                                End If
                                            End If
                                        End If
                                        
                                        spdResult1.Col = intCol
                                        spdResult1.Row = intRow
                                        spdResult1.ForeColor = IIf(varRef <> "", vbRed, vbBlack)
                                        
                                        '-- Local등록
                                        If chkMdbSave.Value = 1 Then
                                            sqlDoc = "Update INTERFACE003" & _
                                                     "   set RESULT1  = '" & strRstval & "', REFERENCE = '" & varRef & "'" & _
                                                     " where SPCNO   = '" & varBarno & "'" & _
                                                     "   and TESTCD  = '" & itemX.Text & "'" & _
                                                     "   and REGDATE = '" & varRegDt & "'"
                
                                            AdoCn_Jet.Execute sqlDoc, sqlRet
                        
                                            If sqlRet = 0 Then
                                                sqlDoc = "insert into INTERFACE003(" & _
                                                         "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN, AUTOMSG, MANUALMSG)" & _
                                                         "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
                                                         "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
                                                         "            '" & strRstval & "','','', '" & varRef & "','',''," & _
                                                         "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','','','')"
                                                AdoCn_Jet.Execute sqlDoc
                                            End If
                                        End If
                                        
                                        '-- Server등록
                                        If chkSvrSave.Value = 1 Then
                                        
                                        End If
                                    End If
                                End If
                                                        
                                Set itemX = Nothing
                            Next
                        '-- 해당번호 못찾음
                        Else
                            intRow = SeqNullSearch(spdResult1, "", 1)
                            If intRow = 0 Then
                                .MaxRows = .MaxRows + 1
                                intRow = .MaxRows
                                lstErr.AddItem "검체번호 " & SeqNo & "은(는) 오더리스트에 없는 결과입니다."
                                .SetText 1, .MaxRows, SeqNo
                                .SetText 8, .MaxRows, Format(Now, "yyyy-mm-dd")
                            Else
                                .SetText 1, intRow, SeqNo
                                .SetText 8, intRow, Format(Now, "yyyy-mm-dd")
                                lstErr.AddItem "검체번호 " & SeqNo & "은(는) 오더리스트에 없는 결과입니다."
                            End If
                            For intCol = 9 To .MaxCols
                                .GetText intCol, 0, varTmp
                                Channel_No = Mid(fAxsym(3), 4, 3)
                                If Right(Channel_No, 1) = "^" Then Channel_No = Mid(Channel_No, 1, 2)

                                If Right(fAxsym(3), 1) = "P" Then Exit Sub

                                strRstval = Trim(fAxsym(4))
                                tmpRstval = Replace(strRstval, ">", "")
                                tmpRstval = Replace(tmpRstval, "<", "")
                                Select Case Channel_No
                                    Case "106" '-- Hbs Ag
                                        If Val(tmpRstval) >= 2 And Val(tmpRstval) <= 10 Then '-- WeeklyPositive
                                            strRstval = "weakly Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) > 10 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) < 2 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "118" '-- Hbs Ab
                                        If Val(tmpRstval) >= 10 And Val(tmpRstval) <= 20 Then '-- WeeklyPositive
                                            strRstval = "weakly Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) > 20 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) < 10 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "841" '-- HCV
                                        If Val(tmpRstval) > 1 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(tmpRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "817" '-- Hiv
                                        If Val(tmpRstval) > 1 Then '-- Positive
                                            strRstval = ""
                                        ElseIf Val(tmpRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case Else
                                
                                End Select
                                
'                                strRstval = Mid(strRstval, 1, 20)
                                
                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                If Not itemX Is Nothing Then
                                    If Channel_No = itemX.Tag Then
                                        .SetText intCol, intRow, strRstval
                                    
                                        strTestcd = itemX.ListSubItems(1)
                                        
                                        .GetText 1, intRow, varBarno
                                        .GetText 4, intRow, varSPnm
                                        .GetText 5, intRow, varSPid '--성별/나이
                                        .GetText 6, intRow, varZation
                                        If varZation = "입원" Then
                                            varZation = "1"
                                        ElseIf varZation = "외래" Then
                                            varZation = "2"
                                        Else
                                            varZation = "3"
                                        End If
                                        .GetText 7, intRow, varRegNo
                                        .GetText 8, intRow, varRegDt
                                        varSex = Mid(varSPid, 1, 1)
                                        varAge = Mid(varSPid, 3)
                                        varRef = ""
                                        
                                        If Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817" Then
                                            If varSex = "남" Then
                                                If fAxsym(4) < Val(itemX.ListSubItems(8)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(9)) Then
                                                    varRef = "H"
                                                End If
                                            Else
                                                If fAxsym(4) < Val(itemX.ListSubItems(10)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(11)) Then
                                                    varRef = "H"
                                                End If
                                            End If
                                        End If
                                        
                                        spdResult1.Col = intCol
                                        spdResult1.Row = intRow
                                        spdResult1.ForeColor = IIf(varRef <> "", vbRed, vbBlack)

                                        '-- Local등록
                                        If chkMdbSave.Value = 1 Then
                                            sqlDoc = "Update INTERFACE003" & _
                                                     "   set RESULT1  = '" & strRstval & "', REFERENCE = '" & varRef & "'" & _
                                                     " where SPCNO   = '" & varBarno & "'" & _
                                                     "   and TESTCD  = '" & itemX.Text & "'" & _
                                                     "   and REGDATE = '" & varRegDt & "'"
                
                                            AdoCn_Jet.Execute sqlDoc, sqlRet
                        
                                            If sqlRet = 0 Then
'                                                sqlDoc = "insert into INTERFACE003(" & _
'                                                         "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
'                                                         "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
'                                                         "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
'                                                         "            '" & strRstval & "','','', '" & varRef & "','',''," & _
'                                                         "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
'                                                AdoCn_Jet.Execute sqlDoc
                                                sqlDoc = "insert into INTERFACE003(" & _
                                                         "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN, AUTOMSG, MANUALMSG)" & _
                                                         "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
                                                         "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
                                                         "            '" & strRstval & "','','', '" & varRef & "','',''," & _
                                                         "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','','','')"
                                                AdoCn_Jet.Execute sqlDoc
                                            End If
                                        End If
                                    End If
                                End If
                                                        
                                Set itemX = Nothing
                            Next
                            
                        End If
                    End With

                End If
            Case "C" 'Comment
                Exit For
                
            Case "L" 'Message Termination
                If SendFlg = True Then Exit Sub
                AXSYM.SID = Right(Trim(AXSYM.SID), 10)
                If AXSYM.SID = "" Then Exit Sub
                If AXSYM.SampleNo = "" Then AXSYM.SampleNo = "0"
                If AXSYM.SampleTy = "" Then AXSYM.SampleTy = "1"
                Exit For
               
            Case "Q" 'Request Information
                SendFlg = True
                SendCount = 0
                cInterface.state = "Q"
                If Len(ReceiveData) > 0 Then
                    AXSYM.SampleNo = Mid(ReceiveData, InStr(ReceiveData, "^") + 1)
                    AXSYM.SampleNo = Mid(AXSYM.SampleNo, 1, InStr(AXSYM.SampleNo, "|") - 1)
                End If
                
                If AXSYM.SampleNo = "" Then AXSYM.SampleNo = "0"
                If AXSYM.SampleTy = "" Then AXSYM.SampleTy = "1"
                Exit For
            Case Else
                
        End Select
    Next iCnt
    Exit Sub
    
Clear_AXSYM_:
    AXSYM.TestDate = ""
    AXSYM.TestTime = ""
    AXSYM.SampleNo = ""
    AXSYM.SID = ""
    AXSYM.SampleTy = ""
    AXSYM.RackNo = ""
    AXSYM.Position = ""
    AXSYM.Priority = ""
    For ii = 0 To 100
        AXSYM.TestId(ii) = ""
        AXSYM.Result(ii) = ""
        AXSYM.Status(ii) = ""
        AXSYM.Rerun(ii) = ""
    Next ii
    flgETB = False
    SendCount = 0
    SendFlg = False
    Return
    
    Exit Sub
ErrorTrap:
    Call ErrMsgProc(CallForm)

errReceive:


End Sub


Public Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Public Function Text_Change(FSend_Str As String, FCheck_Char As String, FChange_Char As String) As String
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

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function



Private Sub Form_Activate()

    If IS_SET = False Then Unload Me
    
    Call cmdRun           ' 실행

End Sub

Private Sub Form_Load()
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    'Call cmdRun           ' 실행
    
    Open App.Path + "\" + "Axsym.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    cInterface.phase = 1
    bufcnt = 0
    
    DoEvents

End Sub

Private Sub f_subSet_ItemHeader()
    
    '검사코드 테이블
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "검사결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFLM", "참고치남(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFHM", "참고치남(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFLF", "참고치여(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFHF", "참고치여(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "재검", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With

End Sub

Private Function f_subSet_WorkList()
    Dim sqlRet      As Integer
    Dim gSql        As String

On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"

    Set AdoRs_ORACLE = New ADODB.Recordset

    gSql = "select 처방전코드, 처방전명, 검체번호, 검체명, 품목코드, 품목명, 접수일자, 입외구분, 병록번호, 성명, 생년월일, 나이, 성별코드, 검체코드, 과코드, 특기사항, 처리구분코드 " & _
           "  from cli.검사검체1v " & _
           "   where (처리구분코드 <> 'N' or 처리구분코드 <> 'R') " & _
           "   and 접수일자 = '2004-03-17' " & _
           "   and 검체번호 = '9135' " & _
           " order by 검체번호 "

'           " where 처방전코드 = '250' " & _

'    gSql = "select * from cli.검사기품목 "
    
    With DataRs(gSql)
        Dim ii As Integer
        
        Do Until .EOF
        
            'Debug.Print .Fields("처방전코드")
            If .Fields("처방전코드") = "249" Then
'                Stop
'                Debug.Print .Fields("처방전코드")
'                Debug.Print .Fields("처방전명")
                Debug.Print .Fields("품목코드")
                Debug.Print .Fields("품목명")
            End If
            'Debug.Print .Fields("처방전명")
            
            .MoveNext
        Loop
        .Close
    End With

Exit Function

ErrorTrap:
    'Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)


End Function

Private Function f_funGet_CODE(ByVal strOrdcd As String) As String

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE = ""
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

Private Function f_subSet_TestList(ByVal strRecei As String)
    Dim sqlRet      As Integer
    Dim gSql        As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    gSql = "select * " & _
           "  from cli.검사검체1v " & _
           " where 검체번호 = '" & strRecei & "'" & _
           "   and (처리구분코드 <> 'N' or 처리구분코드 <> 'R' or 처리구분코드 <> 'E') " & _
           " order by 품목코드 "
    
'           " where 처방전코드 = '250' " & _

    
    Set f_subSet_TestList = DataRs(gSql)
    
'    Set AdoRs_ORACLE = New ADODB.Recordset
'    AdoRs_ORACLE.CursorLocation = adUseClient
'    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE
'
'    Set f_subSet_TestList = AdoRs_ORACLE

Exit Function

ErrorTrap:
'    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function


Private Sub f_subGet_Setting()
    
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subGet_Setting()"
    Set objComSetting = New clsCommon
    
    With objComSetting
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpProperty(INS_CODE)
    End With
    Set objComSetting = Nothing
    
    If mAdoRs Is Nothing Then
        IS_SET = False
        MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
            Set mAdoRs = Nothing
            Exit Sub
        Else
            IS_SET = True
            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")
            
            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
'                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
'                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
'                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
'                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
'                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
'                .RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
'                .SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
            End With
            Call Del_OldData
        End If
    End If
    
    Set mAdoRs = Nothing
Exit Sub

ErrRoutine:
    Set objComSetting = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If frmComm.comEQP.PortOpen <> False Then
        If MsgBox("인터페이스중입니다." & Chr(10) & _
               "작업을 종료하면 받고있거나 검사중인 데이터를 잃게 됩니다" & Chr(10) & _
               "종료하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
               
            Call cmdStop
            Set Result = Nothing
            
            Close #1
            Unload Me
        Else
'            Call cmdStop
'            Set Result = Nothing
'
            Close #1
            Exit Sub
        End If
    Else
        Call cmdStop
        Set Result = Nothing
        
        Close #1
        Unload Me
    End If
    
End Sub


Private Sub imgReceive_DblClick()

'    If FrameResult.Visible = False Then
'        FrameResult.Visible = True
'    Else
'        FrameInterface.Visible = True
'        FrameInterface.ZOrder 0
'    End If

End Sub

Private Sub imgSend_DblClick()
    
'    If FrameResult.Visible = True Then
'        FrameResult.Visible = False
'    Else
'        FrameInterface.Visible = True
'        FrameInterface.ZOrder 0
'    End If

End Sub


Private Sub Order_Ready(ByVal ACK As String)

    Static msgIndex As Long
    
    Select Case ACK
        Case Chr(COM_ENQ)
            msgIndex = 1
        Case Chr(COM_ACK)
            msgIndex = msgIndex + 1
        Case Chr(COM_NACK)
            msgIndex = msgIndex
        Case Chr(COM_EOT)
            msgIndex = 7
            Set Order = Nothing
        Case Else
        
    End Select
    
    Select Case msgIndex
        Case 1
            Call COM_OUTPUT(Order.MSG_ENQ)
        Case 2
            Call COM_OUTPUT(Order.MSG_HEADER)
        Case 3
            Call COM_OUTPUT(Order.MSG_PATIENT)
        Case 4
            Call COM_OUTPUT(Order.MSG_ORDER)
        Case 5
            Call COM_OUTPUT(Order.MSG_TERMINATION)
        Case 6
            Call COM_OUTPUT(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then Exit Sub
    
    intCol1 = 9
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstDetail.Row = intRow1
            spdRstDetail.Col = intCol2
            spdRstDetail.Text = .Text
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstDetail.MaxRows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If
        
        Next
    End With
    
End Sub

'Private Sub spdResult1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim intCol1 As Integer
'    Dim intCol2 As Integer
'    Dim intRow1 As Integer
'    Dim intRow2 As Integer
'    Dim iCnt    As Integer
'
'    intCol1 = 9
'    intCol2 = 2
'    intRow1 = 1
'
'    With spdResult1
'        For iCnt = intCol1 To .MaxCols
'            .Row = NewRow
'            .Col = intCol1
'
'            spdRstDetail.Row = intRow1
'            spdRstDetail.Col = intCol2
'            spdRstDetail.Text = .Text
'
'            intRow1 = intRow1 + 1
'            intCol1 = intCol1 + 1
'
'            If intRow1 > spdRstDetail.MaxRows Then
'                intRow1 = 1
'                intCol2 = intCol2 + 2
'            End If
'
'        Next
'    End With
'
'End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim I As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For I = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(I).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - I)) + (70 * (cmdAction.UBound - I)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub



