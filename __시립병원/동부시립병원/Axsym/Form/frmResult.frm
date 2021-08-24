VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmResult 
   Caption         =   "결과조회"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15210
   WindowState     =   2  '최대화
   Begin Threed.SSCommand cmdSel 
      Height          =   360
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   450
      Width           =   285
      _Version        =   65536
      _ExtentX        =   503
      _ExtentY        =   644
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "frmResult.frx":0000
   End
   Begin FPSpreadADO.fpSpread spdResult1 
      Height          =   4005
      Left            =   90
      TabIndex        =   29
      Top             =   450
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   7064
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColsFrozen      =   7
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
      MaxCols         =   21
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      ShadowText      =   0
      SpreadDesigner  =   "frmResult.frx":046E
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1125
      Left            =   90
      TabIndex        =   16
      Top             =   4470
      Width           =   5385
      _Version        =   65536
      _ExtentX        =   9499
      _ExtentY        =   1984
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
      Begin VB.TextBox txtAge 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   4830
         TabIndex        =   26
         Top             =   270
         Width           =   450
      End
      Begin VB.TextBox txtHospNo 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   3480
         TabIndex        =   23
         Top             =   765
         Width           =   1380
      End
      Begin VB.TextBox txtInGbn 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         TabIndex        =   22
         Top             =   765
         Width           =   1440
      End
      Begin VB.TextBox txtSex 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   4260
         TabIndex        =   20
         Top             =   270
         Width           =   450
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   3150
         TabIndex        =   18
         Top             =   270
         Width           =   1080
      End
      Begin VB.TextBox txtBarno 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         TabIndex        =   17
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   5
         Left            =   4740
         TabIndex        =   27
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "검체번호"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   25
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "이름"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   3
         Left            =   2700
         TabIndex        =   24
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "입원구분"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   21
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "병록번호"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   0
         Left            =   2700
         TabIndex        =   19
         Top             =   765
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdSerch 
      Caption         =   "조  회"
      Height          =   315
      Left            =   5640
      TabIndex        =   14
      Top             =   90
      Width           =   855
   End
   Begin VB.CheckBox chkRSH 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      Caption         =   "RSH"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   7740
      TabIndex        =   6
      Top             =   9060
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CheckBox chkRCP 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      Caption         =   "RCP"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   7020
      TabIndex        =   5
      Top             =   9060
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.ComboBox cboReg 
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   -210
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ComboBox cboRegGbn 
      Height          =   300
      Left            =   3990
      TabIndex        =   1
      Top             =   90
      Width           =   1455
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   90
      TabIndex        =   0
      Top             =   9330
      Width           =   15030
      Begin VB.CommandButton cmdAction 
         Caption         =   "Close"
         Height          =   375
         Index           =   4
         Left            =   13530
         TabIndex        =   37
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "결과지"
         Height          =   375
         Index           =   3
         Left            =   12210
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Clear"
         Height          =   375
         Index           =   2
         Left            =   10890
         TabIndex        =   12
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   9570
         TabIndex        =   11
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "서버등록"
         Height          =   375
         Index           =   0
         Left            =   8250
         TabIndex        =   10
         Top             =   120
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
         Left            =   360
         TabIndex        =   36
         Top             =   210
         Width           =   1200
      End
   End
   Begin MSComCtl2.DTPicker dtpRegDate 
      Height          =   300
      Left            =   1065
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
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
      Format          =   25231361
      CurrentDate     =   37112
   End
   Begin Threed.SSCommand cmdSel 
      Height          =   360
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   450
      Width           =   285
      _Version        =   65536
      _ExtentX        =   503
      _ExtentY        =   644
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "frmResult.frx":0C5D
   End
   Begin Threed.SSFrame FrameResult 
      Height          =   4785
      Left            =   5520
      TabIndex        =   9
      Top             =   4470
      Width           =   9585
      _Version        =   65536
      _ExtentX        =   16907
      _ExtentY        =   8440
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
         Height          =   4335
         Left            =   270
         TabIndex        =   30
         Top             =   270
         Width           =   9045
         _Version        =   393216
         _ExtentX        =   15954
         _ExtentY        =   7646
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   8
         MaxRows         =   14
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmResult.frx":10DF
         ScrollBarTrack  =   1
      End
   End
   Begin MSComctlLib.ListView lvwCuData 
      Height          =   3420
      Left            =   7050
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   6033
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
   Begin MSComctlLib.ImageList imlList 
      Left            =   13980
      Top             =   60
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
            Picture         =   "frmResult.frx":16BD
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":171B
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":1779
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":17D7
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":1835
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":1893
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame FrameError 
      Height          =   3645
      Left            =   90
      TabIndex        =   31
      Top             =   5610
      Width           =   5385
      _Version        =   65536
      _ExtentX        =   9499
      _ExtentY        =   6429
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
      Begin VB.ListBox lstMsg 
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
         Height          =   3120
         Left            =   150
         TabIndex        =   32
         Top             =   390
         Width           =   5115
      End
      Begin VB.ListBox lstErr 
         Height          =   780
         Left            =   150
         TabIndex        =   38
         Top             =   2580
         Width           =   5115
      End
      Begin VB.TextBox txtMsg 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   975
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   35
         Text            =   "frmResult.frx":18F1
         Top             =   1980
         Width           =   5115
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "첨부소견"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   180
         TabIndex        =   34
         Top             =   1770
         Width           =   4785
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "자동소견"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   180
         TabIndex        =   33
         Top             =   180
         Width           =   4785
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "접수일자 :"
      Height          =   180
      Left            =   150
      TabIndex        =   28
      Top             =   150
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "등록구분 :"
      Height          =   180
      Left            =   2970
      TabIndex        =   3
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmResult"
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

Private Sub f_subGet_JobList(ByVal strKeyno As String, ByRef strOrder As String, _
                             ByRef intOrdCnt As Integer, ByRef strSpec As String, _
                             ByRef strPcFlag As String)

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    strOrder = "":  strPcFlag = "  ":   strSpec = "SE": intOrdCnt = 0
    sqlDoc = "select ORD_CODE, CHART_NO From L3A01" & _
             " where SAMPLE_DATE = '" & Mid$(strKeyno, 1, 8) & "'" & _
             "   and SAMPLE_SEQ  = " & Format(Mid$(strKeyno, 9, 3), "##0") & "" & _
             "   and PART        = '" & Mid$(strKeyno, 12, 2) & "'"
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD, REMARK, AUTOVERIFY from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "ORD_CODE = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            Select Case Trim(adoRS2(2) & "")
                Case "128": strSpec = "PL"
                Case Else:  strSpec = "SE"
            End Select
            
            If Trim(adoRS2("TESTCD_EQP") & "") = "XXX" Then
                strOrder = strOrder + "06A ," + Trim$(adoRS2("AUTOVERIFY") & "") + ",": strPcFlag = "PC"
            Else
                strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + " ," + Trim$(adoRS2("AUTOVERIFY") & "") + ","
            End If
            intOrdCnt = intOrdCnt + 1
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
    
End Sub



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
    
    intRow = 1
    intCol = 8
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
        '.MaxRows = 7
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFLM, REFHM, REFLF, REFHF, DELTA, DELTAGBN, PANICL, PANICH " & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") " & _
             "   and ((TESTCD <> '') and (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst:        ReDim fChannel(adoRS.RecordCount + intCol)
    
    Do While Not adoRS.EOF
        'Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TESTCD") & ""), , "LST")
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
            itemX.SubItems(14) = " " 'Trim(adoRS.Fields("MSGLOW") & "")
            itemX.SubItems(15) = " " 'Trim(adoRS.Fields("MSGHIGH") & "")
'            itemX.SubItems(16) = Trim(adoRS.Fields("MANUALMSG") & "")
            itemX.Tag = Trim(adoRS.Fields("TESTCD") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdRstDetail
            If intRow > .MaxRows Then .MaxRows = .MaxRows + 1
            
            .SetText intCol2, intRow, Trim$(adoRS("TESTNM") & "")
            
            Set itemX = lvwCuData.FindItem(Trim$(adoRS("TESTNM") & ""), lvwSubItem, , lvwWhole)
            If Not itemX Is Nothing Then
                .SetText 5, intRow, itemX.ListSubItems(8)
                .SetText 6, intRow, itemX.ListSubItems(9)
                .SetText 7, intRow, itemX.ListSubItems(10)
                .SetText 8, intRow, itemX.ListSubItems(11)
            End If
                                    
            Set itemX = Nothing
'                    End If
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

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdSave
        Case 1:     Call cmdPrint
        Case 2:     Call cmdClear
        Case 3:     Call cmdResultPrint
        Case 4:     Call cmdExit
        Case Else
    End Select

End Sub

Private Sub cmdResultPrint()
   ' Dim intRow As Integer
    Dim varTmp
    
    If spdResult1.ActiveRow < 1 Then
        MsgBox "인쇄할 환자를 선택하세요", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Const iXX = 10, iYY = 0
    
    Dim sgCol   As Single, sgRow    As Single
    
    ReDim sPatDat(0 To 8) As String
    ReDim sRstdat(1 To spdRstDetail.MaxCols) As String
     
    Dim sTemp   As String, iTemp    As Integer
    Dim iCol    As Integer, iRow    As Integer
    
    sPatDat(0) = dtpRegDate.Value   '-- 접수일자
    sPatDat(1) = txtBarno.Text      '-- 검체번호
    sPatDat(2) = txtName.Text       '-- 이름
    sPatDat(3) = txtSex.Text        '-- 성별
    sPatDat(4) = txtAge.Text        '-- 나이
    sPatDat(5) = txtInGbn.Text      '-- 입원구분
    sPatDat(6) = txtHospNo.Text     '-- 병록번호
    sPatDat(7) = lstMsg.Text        '-- 자동소견
    sPatDat(8) = txtMsg.Text        '-- 소견
    
'    With spdRstDetail
'        .Row = .ActiveRow
'        .Col = 1:   sPatDat(2) = .Text  '접수일자
'        If txtDeptCd.Text = "60" Or txtDeptCd.Text = "51" Or txtDeptCd.Text = "55" Or _
'           txtDeptCd.Text = "57" Then
'            .Col = 9:  sPatDat(3) = .Text  '보험종류
'        Else
'            .Col = 10:  sPatDat(3) = .Text  '보험종류
'        End If
'        .Col = 11:  sPatDat(4) = .Text  '조합기호
'        .Col = 12:  sPatDat(5) = .Text  '의료보험증번호
'        .Col = 2:   sPatDat(6) = .Text  '이름
'        .Col = 3:   sSexAge = .Text     '성별(나이)
'        .Col = 4:   sPatDat(7) = .Text  '주민등록번호
'    End With
    
    ' 99-01-15 OJM
'    If sPatDat(2) = "" Then
'        MsgBox "인쇄할 환자를 선택하세요", vbInformation, Me.Caption
'        Exit Sub
'    End If
    ' 99-01-15 OJM
    
'    sPatDat(6) = sPatDat(6) + Space(10) + sSexAge
'    sPatDat(7) = Mid$(sPatDat(7), 1, 6) + "-" + Mid$(sPatDat(7), 7) ' 99/10/6 ojm 수정
      
      
    With Printer
        .EndDoc
        .PaperSize = vbPRPSA4
        .Orientation = 1
        .FontName = "굴림체"
        Printer.FontSize = 11
        Printer.ScaleMode = 6
        
        sgCol = Printer.TextWidth(" "): sgRow = Printer.TextHeight(" ")
        
        'Call Print_Title(iXX, iYY, sPatDat())
        For iRow = 1 To spdRstDetail.MaxRows
            'With spdResult1
                'For iCol = 1 To .MaxCols
'                    .Row = iRow
'                    .Col = iCol:   sRstdat(iCol) = .Text
'                Next
            'End With
            For iCol = 1 To spdRstDetail.MaxCols
                spdRstDetail.GetText 2, iRow, varTmp
                sRstdat(iCol) = varTmp
                
                If sRstdat(1) = "" Then Exit For
                
                If .CurrentY > .ScaleHeight - iYY - (sgRow * 5) Then
                    .NewPage:   iTemp = 0
                    'Call Print_Title(iXX, iYY, sPatDat())
                End If
                
                If Not Mid$(sRstdat(7), 1, 2) = sTemp Then
                    If Not iTemp = 0 Then Printer.Print
                
                    .CurrentX = iXX + 2: Printer.Print "[" + sRstdat(6) + "]"
                    sTemp = Mid$(sRstdat(7), 1, 2)
                End If
                
                If Len(sRstdat(7)) > 7 Then
                    .CurrentX = iXX + 13:   Printer.Print sRstdat(1);
                Else
                    .CurrentX = iXX + 7:    Printer.Print sRstdat(1);
                End If
                .CurrentX = iXX + 65:   Printer.Print sRstdat(2);
                .CurrentX = iXX + 110:  Printer.Print sRstdat(3);
                .CurrentX = iXX + 120:  Printer.Print sRstdat(4);
                .CurrentX = iXX + 155:  Printer.Print sRstdat(5)
                
                iTemp = iTemp + 1
            Next
        Next
        
        .EndDoc
    End With

End Sub

Sub Print_Title(ByVal iXX As Integer, ByVal iYY As Integer _
              , ByRef sPat() As String)

    Dim sgCol   As Single, sgRow    As Single

    sgCol = Printer.TextWidth(" "): sgRow = Printer.TextHeight(" ")
    
    With Printer
    
    .DrawWidth = 2
    
    Printer.Line (iXX, iYY + 15)-(.ScaleWidth - iXX, iYY + 60), , B
    
    .CurrentX = iXX + 5:   .CurrentY = iYY + 20: Printer.Print "진 료 과 : " + sPat(1)
    .CurrentX = iXX + 5:   .CurrentY = iYY + 30: Printer.Print "접수일자 : " + sPat(2)
    .CurrentX = iXX + 5:   .CurrentY = iYY + 40: Printer.Print "주민번호 : " + sPat(7)
    .CurrentX = iXX + 5:   .CurrentY = iYY + 50: Printer.Print "이    름 : " + sPat(6)
    
'    If Mid$(DEPT_담당과, 3, 1) = "Y" Or txtDeptCd.Text = p_과_건강검진실 Then
'        .CurrentX = iXX + 100: .CurrentY = iYY + 30: Printer.Print "보험종류   : " + sPat(3)
'        .CurrentX = iXX + 100: .CurrentY = iYY + 40: Printer.Print "조합기호   : " + sPat(4)
'        .CurrentX = iXX + 100: .CurrentY = iYY + 50: Printer.Print "보험증번호 : " + sPat(5)
'    ElseIf txtDeptCd.Text = "60" Or txtDeptCd.Text = "51" Or txtDeptCd.Text = "55" Or _
'           txtDeptCd.Text = "57" Then
'        .CurrentX = iXX + 100: .CurrentY = iYY + 30: Printer.Print "등록번호   : " + sPat(3)
'        .CurrentX = iXX + 100: .CurrentY = iYY + 40: Printer.Print "전화번호   : " + sPat(4)
'        If txtDeptCd.Text = p_과_영유아실 Then
'            .CurrentX = iXX + 100: .CurrentY = iYY + 50: Printer.Print "엄마이름   : " + sPat(5)
'        End If
'    End If
    .CurrentX = iXX + 2:  .CurrentY = iYY + 75:    Printer.Print "검사명"
    .CurrentX = iXX + 65:  .CurrentY = iYY + 75:   Printer.Print "결과"
    .CurrentX = iXX + 120: .CurrentY = iYY + 75:   Printer.Print "참고치"
    
    Printer.Line (iXX, iYY + 80)-(.ScaleWidth - iXX, iYY + 80)

    Printer.Line (iXX, .ScaleHeight - iYY - sgRow * 3)-(.ScaleWidth - iXX, .ScaleHeight - iYY - sgRow * 3)
    
    .CurrentY = .ScaleHeight - iYY - sgRow * 2
'    .CurrentX = .ScaleWidth - iXX - .TextWidth(p_기타_보건소)
'    Printer.Print p_기타_보건소
    
    .CurrentY = iYY + 85
    
    End With
    
End Sub

Private Sub cmdClear()
    
    dtpRegDate = Now
    cboReg.Clear
    cboRegGbn.Clear
    
    lstMsg.Clear
    
    cboReg.AddItem "접수일"
    cboReg.AddItem "등록일"
    cboReg.ListIndex = 0
    
    cboRegGbn.AddItem "전체자료"
    cboRegGbn.AddItem "등록자료"
    cboRegGbn.AddItem "미등록자료"
    cboRegGbn.ListIndex = 0
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        '.BackColor = vbWhite
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    With spdRstDetail
        .Col = 2:   .Col2 = 4
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    txtBarno.Text = ""
    txtName.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    txtInGbn.Text = ""
    txtHospNo.Text = ""
    txtMsg.Text = ""


End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Private Function f_subSet_TestList(ByVal strRecei As String) As OraDynaset
    Dim sqlRet      As Integer
    Dim gSql        As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
      
    gSql = "select 품목코드 " & _
           "  from cli.검사검체1v " & _
           " where 검체번호 = '" & strRecei & "'" & _
           "   and (처리구분코드 <> 'N' or 처리구분코드 <> 'R' or 처리구분코드 <> 'E') " & _
           " order by 품목코드 "
    
    Set f_subSet_TestList = DataRs(gSql)
    
Exit Function

ErrorTrap:
    Set f_subSet_TestList = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub cmdSave()
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim sRs As Object
    
    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String, strBarNo     As String, strTime      As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As fpSpread
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    Dim gSql    As String
    Dim sTxt    As String
    Dim tTxt    As String
    Dim CntsTxt As Integer
    
    Dim varBarno As Variant
    Dim varSPid  As Variant
    Dim varSPnm  As Variant
    Dim varSex   As Variant
    Dim varAge   As Variant
    Dim varZation  As Variant
    Dim varRef   As Variant
    Dim varRefH  As Variant
    Dim varRefL  As Variant
    Dim varRegDt As Variant
    Dim varRegNo As Variant
    Dim lstCnt   As Integer
    Dim lstMessage   As String
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"
    
On Error GoTo ErrorRoutine
    
    Me.MousePointer = 11
    
    Set objSpd = spdResult1
    
    CntsTxt = 0:    sTxt = "":  tTxt = ""
    
    With objSpd
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp:         strBarNo = Trim$(varTmp)
            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
            
            .GetText 1, intRow, varTmp
            
            If strBarNo = "" Then Exit For
            
            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 8 To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp
                            strTestcd = itemX.ListSubItems(1)
                            
                            .GetText 2, intRow, varBarno
                            .GetText 3, intRow, varSPnm
                            .GetText 4, intRow, varSPid
                            .GetText 5, intRow, varZation
                            If varZation = "입원" Then
                                varZation = "1"
                            ElseIf varZation = "외래" Then
                                varZation = "2"
                            Else
                                varZation = "3"
                            End If
                            .GetText 6, intRow, varRegNo
                            .GetText 7, intRow, varRegDt
                            varSex = Mid(varSPid, 1, 1)
                            varAge = Mid(varSPid, 3)
                            varRef = ""
                            If varSex = "남" Then
                                If varTmp < Val(itemX.ListSubItems(8)) Then
                                    varRef = "L"
                                ElseIf varTmp > Val(itemX.ListSubItems(9)) Then
                                    varRef = "H"
                                End If
                            Else
                                If varTmp < Val(itemX.ListSubItems(10)) Then
                                    varRef = "L"
                                ElseIf varTmp > Val(itemX.ListSubItems(11)) Then
                                    varRef = "H"
                                End If
                            End If
                            
                            
                            For lstCnt = 0 To lstMsg.ListCount - 1
                                lstMessage = lstMessage & lstMsg.List(lstCnt)
                            Next
                            
                            sqlDoc = "Update INTERFACE003" & _
                                     "   set RESULT1  = '" & varTmp & "', REFERENCE = '" & varRef & "' " & _
                                     " where SPCNO   = '" & varBarno & "'" & _
                                     "   and TESTCD  = '" & itemX.Tag & "'" & _
                                     "   and REGDATE = '" & varRegDt & "'"
                                     
                                     '"       AUTOMSG  = '" & lstMessage & "', MANUALMSG = '" & Trim(txtMsg.Text) & "' " & _

                            AdoCn_Jet.Execute sqlDoc, sqlRet
                            
                            If sqlRet = 0 Then
                                sqlDoc = "insert into INTERFACE003(" & _
                                         "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
                                         "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
                                         "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
                                         "            '" & varTmp & "','','', '" & varRef & "','',''," & _
                                         "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
                                         
                                AdoCn_Jet.Execute sqlDoc
                            End If
                            
'                            blnFlag = False
'                            Set sRs = f_subSet_TestList(strBarNo)
'
'                            Dim tmpTestcd As Variant
'                            Dim tmpCnt    As Integer
'
'                            tmpTestcd = Split(strTestcd, ",")
'                            Do Until sRs.EOF
'                                For tmpCnt = 0 To UBound(tmpTestcd)
'                                    If sRs("품목코드") = tmpTestcd(tmpCnt) Then
'                                        blnFlag = True
'                                        Exit Do
'                                    End If
'                                    sRs.MoveNext
'                                Next tmpCnt
'                            Loop
'                            Set sRs = Nothing
                            
'                            If blnFlag Then
                                CntsTxt = CntsTxt + 1
                                sTxt = sTxt + itemX.SubItems(13) + "," + varTmp + ","
'                            End If
                        End If
                                                
                        Set itemX = Nothing
                    End If
                Next
                
                
                If CntsTxt > 0 Then
                    spdResult1.SetText 1, intRow, "0"
                    spdResult1.Row = intRow
                    spdResult1.Col = -1: spdResult1.BackColor = &HFFF8F0
                    
                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                             " where SPCNO   = '" & strBarNo & "'" & _
                             "   and REGDATE = '" & varRegDt & "'"
                             
                    AdoCn_Jet.Execute sqlDoc
                    
                    sTxt = "ax2004," + strBarNo + "," + CStr(CntsTxt) + "," + sTxt + vbLf
                    tTxt = tTxt + sTxt
                    Debug.Print tTxt
                    CntsTxt = 0:    sTxt = ""
                    
                    'lstErr.ForeColor = vbBlue
                    lstErr.AddItem "검체번호 " & strBarNo & "번 서버등록 성공"
                Else
                    'lstErr.ForeColor = vbRed
                    lstErr.AddItem "검체번호 " & strBarNo & "번 서버등록 실패"
                    'MsgBox strErrMsg, vbInformation, Me.Caption
                End If
                
            End If
            
        Next
    End With
    
    'If CntsTxt > 0 And sTxt <> "" And tTxt <> "" Then
        Open App.Path & "\ax2004.dat" For Output As #2
    
        Print #2, tTxt
        
        tTxt = ""
        
        Close #2
        
        Dim Ret As Long
        
        '-- Remote Copy
        If chkRCP.Value = 1 Then
            Ret = WinExec("rcp.exe -a ax2004.dat med.lab:/usr/tmp/ax2004.dat", 2)
        End If
        
        '-- Remote Sheel
        If chkRCP.Value = 1 And chkRSH.Value = 1 Then
            Ret = WinExec("rsh.exe med -l lab ./exp_interface.sh clbdvt02m3 ax2004", 2)
        End If
        
        MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
    'End If
    
    Me.MousePointer = 0
    
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)

End Sub


Private Sub cmdPrint()
    
    Dim strPage As String
    Dim strArea As String
    Dim strPDate As String
    
    strPage = "Page : " & Space(7) & "/p" & " of " & spdRstDetail.PrintPageCount
    strArea = ""
    strPDate = "출력일자:" & Format(Now, "yyyy년mm월dd일")
    
    With SpPrint
        .strTitle = "/fn""굴림체""/fz""20""/fb1/fi0/fu1/fk0/fs1" _
                  & "/f1/c검사결과(" & Format(dtpRegDate.Value, "yyyy-mm-dd") & ")/n"
        .strBaseDate = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs1" _
                     & "/f1/c" & "" & "/n/n"
        .strPageCount = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & strPage & "/n"
        .strAreaName = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/l" & strArea
        .strPrintDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & strPDate & "/n/n"
                      
        .strSpcNo = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r검체번호 : " & Trim(txtBarno.Text) & "/n"
        .strName = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r이름 : " & Trim(txtName.Text) & "  " & Trim(txtSex.Text) & "/" & Trim(txtAge.Text) & "/n"
        '.strSex = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & Trim(txtSex.Text) & "/" & Trim(txtAge.Text) & "/n"
        '.strAge = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & Trim(txtAge.Text) & ""
        .strGbn = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r입원구분 : " & Trim(txtInGbn.Text) & "/n"
        .strChart = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r병록번호 : " & Trim(txtHospNo.Text) & "/n"
        
    End With

    Call Load_From(frmSpPreview)
    
End Sub


Private Sub Load_From(ByVal frm As Form)
    
    With frm
        .Show
        .SetFocus
        
    End With
    
End Sub


Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    With spdResult1
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp
            If Trim$(varTmp) <> "" Then
                .SetText 1, intRow, IIf(Index = 0, "1", "")
                If Index = 0 Then cmdSel(Index).Visible = False: cmdSel(Index + 1).Visible = False
            End If
        Next
    End With
    
    If Index = 0 Then
        cmdSel(0).Visible = False: cmdSel(1).Visible = True
    Else
        cmdSel(0).Visible = True: cmdSel(1).Visible = False
    End If

End Sub

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
            If Trim(.Text) = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With
    
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

Private Sub cmdSerch_Click()
    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intRow  As Integer, intCnt  As Integer
    
On Error GoTo ErrRoutine

    CallForm = "frmInterface - Private Sub cmdSerch_Click()"
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
'        .BackColor = vbWhite
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
        
    intRow = 1
    intCol = 1
    lblStatus.Caption = ""
    
    sqlDoc = "select * from INTERFACE003"
'    If cboReg.ListIndex = 0 Then '-- 접수일
'        spdResult1.SetText 7, 0, "접수일자"
'        sqlDoc = sqlDoc & " where RegDate = '" & Format(dtpRegDate.Value, "yyyy-mm-dd") & "'"
'    Else                     '-- 등록일
        spdResult1.SetText 7, 0, "등록일자"
        sqlDoc = sqlDoc & " where TransDt = '" & Format(dtpRegDate.Value, "yyyy-mm-dd") & "'"
'    End If
    
    If cboRegGbn.ListIndex = 1 Then     '-- 등록
        sqlDoc = sqlDoc & "   and ServerGbn = 'Y' "
    ElseIf cboRegGbn.ListIndex = 2 Then '-- 미등록
        sqlDoc = sqlDoc & "   and ServerGbn = '' "
    End If
    sqlDoc = sqlDoc & " order by SpcNo, TestCd "
                          
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
    Else
        lblStatus.Caption = "조회결과가 없습니다."
        spdResult1.MaxRows = 1
        Set adoRS = Nothing
        Exit Sub
    End If
    
    Do While Not adoRS.EOF
        With spdResult1
            intRow = SeqSearch(spdResult1, Trim$(adoRS("SPCNO") & ""), 2)
            If intRow = 0 Then intRow = SeqNullSearch(spdResult1, Trim$(adoRS("SPCNO") & ""), 2)
            If intRow = 0 Then .MaxRows = .MaxRows + 1: intRow = .MaxRows
            .SetText 2, intRow, Trim$(adoRS("SPCNO") & "")
            .SetText 3, intRow, Trim$(adoRS("NAME") & "")
            .SetText 4, intRow, Trim$(adoRS("SEX") & "") & "/" & Trim$(adoRS("AGE") & "")
            If Trim$(adoRS("HOSPZATION") & "") = "1" Then     '-- 입원
                .SetText 5, intRow, "입원"
            ElseIf Trim$(adoRS("HOSPZATION") & "") = "2" Then '-- 외래
                .SetText 5, intRow, "외래"
            ElseIf Trim$(adoRS("HOSPZATION") & "") = "3" Then '-- 퇴원
                .SetText 5, intRow, "퇴원"
            End If
            
            .SetText 6, intRow, Trim$(adoRS("HOSPNO") & "")
            
            If cboReg.ListIndex = 0 Then '-- 접수일
                .SetText 7, intRow, Trim$(adoRS("REGDATE") & "")
            Else                         '-- 등록일
                .SetText 7, intRow, Trim$(adoRS("REGDATE") & "")
            End If
            
            Set itemX = lvwCuData.FindItem(Trim$(adoRS("TestCd") & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 7
                .SetText intCol, intRow, Trim$(adoRS("Result1") & "")
                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS("Reference") & "") <> "", vbRed, vbBlack)
                .Col = 1: .Row = intRow: .Value = 1
            End If
        
        End With
        
        adoRS.MoveNext
        'intRow = intRow + 1
    Loop
    
    spdResult1.MaxRows = intRow
    
    Set adoRS = Nothing

Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub Form_Load()
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    DoEvents
    SendCount = 0
    
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
            Call .Add(, "MSGLOW", "자동판정(L)", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "MSGHIGH", "자동판정(H)", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "MESSAGE", "소견", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With
    
   
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


Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim itemX   As ListItem
    
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    Dim varTmp
    
    If Row = 0 Then Exit Sub
    With spdRstDetail
        .Col = 2:   .Col2 = 4
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    intCol1 = 8
    intCol2 = 2
    intRow1 = 1
    
    lstMsg.Clear
    txtMsg.Text = ""
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstDetail.Row = intRow1
            spdRstDetail.Col = intCol2
            spdRstDetail.Text = .Text
            
            varTmp = ""
            spdResult1.GetText 2, Row, varTmp: txtBarno.Text = varTmp
            If Trim(varTmp) = "" Then cmdAction(1).Enabled = True: Exit Sub
            cmdAction(1).Enabled = True
            spdResult1.GetText 3, Row, varTmp: txtName.Text = varTmp
            spdResult1.GetText 4, Row, varTmp: txtSex.Text = Mid(varTmp, 1, 1)
            spdResult1.GetText 4, Row, varTmp: txtAge.Text = Mid(varTmp, 3)
            spdResult1.GetText 5, Row, varTmp: txtInGbn.Text = varTmp
            spdResult1.GetText 6, Row, varTmp: txtHospNo.Text = varTmp
             
            spdResult1.GetText intCol1, 0, varTmp
            Set itemX = lvwCuData.FindItem(varTmp, lvwSubItem, , lvwWhole)
            If Not itemX Is Nothing Then
                spdResult1.Row = Row
                spdResult1.Col = intCol1
                If txtSex.Text = "남" Then
                    If spdResult1.Text <> "" And itemX.SubItems(8) <> "" And Val(spdResult1.Text) < Val(itemX.SubItems(8)) Then
                        spdRstDetail.Col = 3
                        spdRstDetail.Row = intRow1
                        spdRstDetail.ForeColor = vbRed
                        spdRstDetail.SetText 3, intRow1, "Low"
                        lstMsg.AddItem itemX.SubItems(14)
                    ElseIf spdResult1.Text <> "" And itemX.SubItems(9) <> "" And Val(spdResult1.Text) > Val(itemX.SubItems(9)) Then
                        spdRstDetail.Col = 3
                        spdRstDetail.Row = intRow1
                        spdRstDetail.ForeColor = vbRed
                        spdRstDetail.SetText 3, intRow1, "High"
                        lstMsg.AddItem itemX.SubItems(15)
                    End If
                ElseIf txtSex.Text = "여" Then
                    If spdResult1.Text <> "" And itemX.SubItems(10) <> "" And spdResult1.Text < itemX.SubItems(10) Then
                        spdRstDetail.Col = 3
                        spdRstDetail.Row = intRow1
                        spdRstDetail.ForeColor = vbRed
                        spdRstDetail.SetText 3, intRow1, "Low"
                        lstMsg.AddItem itemX.SubItems(14)
                    ElseIf spdResult1.Text <> "" And itemX.SubItems(11) <> "" And spdResult1.Text > itemX.SubItems(11) Then
                        spdRstDetail.Col = 3
                        spdRstDetail.Row = intRow1
                        spdRstDetail.ForeColor = vbRed
                        spdRstDetail.SetText 3, intRow1, "High"
                        lstMsg.AddItem itemX.SubItems(15)
                    End If
                End If
            End If
             
            varTmp = ""
            
            If cboReg.ListIndex = 0 Then '-- 접수일
                sqlDoc = "select RegDate, Result1 from INTERFACE003"
                sqlDoc = sqlDoc & " where RegDate < '" & Format(dtpRegDate.Value, "yyyy-mm-dd") & "'"
            Else                     '-- 등록일
                sqlDoc = "select TransDt, Result1, ManualMsg from INTERFACE003"
                sqlDoc = sqlDoc & " where TransDt < '" & Format(dtpRegDate.Value, "yyyy-mm-dd") & "'"
            End If
            
            If cboRegGbn.ListIndex = 1 Then     '-- 등록
                sqlDoc = sqlDoc & "   and ServerGbn = 'Y' "
            ElseIf cboRegGbn.ListIndex = 2 Then '-- 미등록
                sqlDoc = sqlDoc & "   and ServerGbn Is Null "
            End If
            
            spdResult1.GetText 6, Row, varTmp
            sqlDoc = sqlDoc & "   and HospNo = '" & varTmp & "' "
            
            spdResult1.GetText intCol1, 0, varTmp
            Set itemX = lvwCuData.FindItem(varTmp, lvwSubItem, , lvwWhole)
            If Not itemX Is Nothing Then
                sqlDoc = sqlDoc & "   and TestCd = '" & itemX & "'"
            End If
            
            sqlDoc = sqlDoc & " order by 1 desc "
                                  
            adoRS.CursorLocation = adUseClient
            adoRS.Open sqlDoc, AdoCn_Jet
            
            If adoRS.RecordCount > 0 Then
                adoRS.MoveFirst
                spdRstDetail.SetText 4, intRow1, Trim(adoRS.Fields("Result1") & "")
                'If Not IsNull(adoRS.Fields("ManualMsg")) Then txtMsg.Text = Trim(adoRS.Fields("ManualMsg") & "")
            Else
                spdRstDetail.SetText 4, intRow1, ""
            End If
            adoRS.Close
            
            intCol1 = intCol1 + 1
            intRow1 = intRow1 + 1
        Next
    End With

End Sub

Private Sub spdResult1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    
    spdResult1.Col = 1
    spdResult1.Col2 = 8
    spdResult1.Row = 1
    spdResult1.Row2 = spdResult1.MaxRows
    spdResult1.BlockMode = True
    spdResult1.BackColor = 14286847
    spdResult1.BlockMode = False
    
    spdResult1.Col = 9
    spdResult1.Col2 = spdResult1.MaxCols
    spdResult1.Row = 1
    spdResult1.Row2 = spdResult1.MaxRows
    spdResult1.BlockMode = True
    spdResult1.BackColor = vbWhite
    spdResult1.BlockMode = False
    
    spdResult1.Col = 1
    spdResult1.Col2 = spdResult1.MaxCols
    spdResult1.Row = y / 295
    spdResult1.Row2 = y / 295
    spdResult1.BlockMode = True
    spdResult1.BackColor = &HFFC0FF
    spdResult1.BlockMode = False

End Sub

Private Sub spdRstDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    
    spdRstDetail.Col = 1
    spdRstDetail.Col2 = spdRstDetail.MaxCols
    spdRstDetail.Row = 1
    spdRstDetail.Row2 = spdRstDetail.MaxRows
    spdRstDetail.BlockMode = True
    spdRstDetail.BackColor = vbWhite
    spdRstDetail.BlockMode = False
    
    spdRstDetail.Col = 1
    spdRstDetail.Col2 = spdRstDetail.MaxCols
    spdRstDetail.Row = y / 310
    spdRstDetail.Row2 = y / 310
    spdRstDetail.BlockMode = True
    spdRstDetail.BackColor = &H80C0FF
    spdRstDetail.BlockMode = False
End Sub
