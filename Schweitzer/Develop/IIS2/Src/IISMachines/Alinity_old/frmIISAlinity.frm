VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{D74ED2A2-3650-4720-93BC-FDDD8DCBC769}#1.0#0"; "Han2EngOCX.ocx"
Begin VB.Form frmIISAlinity 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Alinity"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows ?⺻??
   Begin HAN2ENGOCXLib.Han2EngOCX Han2EngOCX1 
      Height          =   255
      Left            =   8610
      TabIndex        =   27
      Top             =   8670
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin MSWinsockLib.Winsock wSck 
      Left            =   8070
      Top             =   8580
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   4710
      TabIndex        =   26
      Top             =   150
      Visible         =   0   'False
      Width           =   615
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   107
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "?? ȯ??????"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   495
      Left            =   11475
      Style           =   1  '?׷???
      TabIndex        =   0
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ??????(&C)"
      Height          =   495
      Left            =   12694
      Style           =   1  '?׷???
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "?? ??(&X)"
      Height          =   495
      Left            =   13913
      Style           =   1  '?׷???
      TabIndex        =   2
      Top             =   8567
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1290
      Left            =   6548
      TabIndex        =   4
      Top             =   405
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   5
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "00000001"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   6
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "?̻???"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStatFg 
         Height          =   315
         Left            =   6795
         TabIndex        =   7
         Top             =   165
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "????"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   315
         Left            =   1245
         TabIndex        =   8
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "?̻??? ?Ʊ?"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   9
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "??????"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   6795
         TabIndex        =   10
         Top             =   525
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "Blood"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   1245
         TabIndex        =   11
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "???? / 29"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   12
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "65????"
         Appearance      =   0
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "?? ü ?? :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   5760
         TabIndex        =   20
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "???޿??? :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "??  ?? : "
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   3105
         TabIndex        =   18
         Top             =   975
         Width           =   810
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "?????? :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   3105
         TabIndex        =   17
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ó???? :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   3105
         TabIndex        =   16
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "????/???? :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   15
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "??     ?? :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   14
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ȯ  ?? ID :"
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6870
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7530
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3555
      Left            =   105
      TabIndex        =   21
      Top             =   525
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   6271
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   11
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISAlinity.frx":0000
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   105
      TabIndex        =   22
      Top             =   4663
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   7779
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISAlinity.frx":045D
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   98
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   107
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "?? ?˻????? ????Ʈ"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Left            =   98
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4247
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "?? ?˻??Ϸ? ????Ʈ"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6690
      Left            =   6555
      TabIndex        =   25
      Top             =   1710
      Width           =   8580
      _Version        =   393216
      _ExtentX        =   15134
      _ExtentY        =   11800
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   22
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISAlinity.frx":0BFE
      TextTip         =   2
   End
End
Attribute VB_Name = "frmIISAlinity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ???ϸ?  : frmIISAlinity.frm
'   ?ۼ???  : ??????
'   ??  ??  : Alinity ??????
'   ?ۼ???  : 2019-12-19
'   ??  ??  :
'   ??  ??  :
'       1. ????????
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady?? Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
End Enum

'## tblComplete?? Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccSexAge = 6
    ccDoctNm = 7:       ccDeptNm = 8
    ccWardNm = 9:       ccStatFg = 10
    ccSpcNm = 11:       ccQcFg = 12
    ccSendCnt = 13:     ccResult = 14
End Enum

'## tblResult?? Column Enum
Private Enum TResultEnum
    ccTestNm = 1
    ccEqpResult = 2
    ccLISResult = 3
    ccUnit = 4
    ccHLDiv = 5
    ccDPDiv = 6
    ccRef = 7
    ccInfo = 8
End Enum

'## Clear Enum
Private Enum ClearEnum
    ccAll = 1
    ccLabel = 2
End Enum

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '???????̽? Ŭ????
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '?˾??޴?
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors          '???????̽? ???? ?÷???
Private mOrder      As clsIISIntOrder           '???????? Ŭ????

Private mEqpCd  As String   '?????ڵ?
Private mEqpKey As String   '????Ű

Dim strDump As String
Dim strRackPos  As String

Private gSckPort    As String
Private gEqp1       As String
Private gEqp2       As String


Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Private Function GetAlinityConfig(ByVal strConfigNm As String) As String
    Dim strFileName As String
    Dim strReturnedString As String
    
    GetAlinityConfig = ""
    
    strFileName = App.Path & "\Alinity.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "ALINITY", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetAlinityConfig = strReturnedString
    
End Function


Private Sub Command1_Click()

    strDump = ""
    strDump = strDump & "1H|\^&|||Alinity^2.20^52500030827^H1P1O1R1C1Q1L1|||||||P|1|20070801102234" & vbNewLine
    strDump = strDump & "07" & vbNewLine
    strDump = strDump & "2P|1||01155107|01155107" & vbNewLine
    strDump = strDump & "DB" & vbNewLine
    strDump = strDump & "3O|1|17000416672|17000416672^J401^1|^^^161^Anti-HCV^UNDILUTED^P|R||||||||||||||||||||F" & vbNewLine
    strDump = strDump & "0B" & vbNewLine
    strDump = strDump & "4R|1|^^^161^Anti-HCV^UNDILUTED^P^48229HN00^04462^^F|0.06|S/CO||||F||FSE^FSE||20070801102234|iSR03698" & vbNewLine
    strDump = strDump & "42" & vbNewLine
    strDump = strDump & "5R|2|^^^161^Anti-HCV^UNDILUTED^P^48229HN00^04462^^I|Nonreactive|||||F||FSE^FSE||20070801102234|iSR03698" & vbNewLine
    strDump = strDump & "ED" & vbNewLine
    strDump = strDump & "6R|3|^^^161^Anti-HCV^UNDILUTED^P^48229HN00^04462^^P|1448|RLU||||F||FSE^FSE||20070801102234|iSR03698" & vbNewLine
    strDump = strDump & "3C" & vbNewLine
    strDump = strDump & "7L|1" & vbNewLine
    strDump = strDump & "40" & vbNewLine
    strDump = strDump & ""



    Call MSComm_OnComm

End Sub

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
    Me.MDIActiveX.WindowState = ccMaximize
End Sub

Private Sub Form_Load()
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    Set mOrder = New clsIISIntOrder
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    'TCP-IP ???Ż?????
    'Call GetEqpComm
    
    gSckPort = GetAlinityConfig("TCPPORT")
    gEqp1 = GetAlinityConfig("EQP1")
    gEqp2 = GetAlinityConfig("EQP2")


    If gSckPort <> "" And IsNumeric(gSckPort) Then
        wSck.LocalPort = CInt(gSckPort)
        wSck.Listen
    End If
    
    DoEvents
    
    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mOrder = Nothing
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISAlinity = Nothing
End Sub

Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '?????? ǥ?? Ŭ????

    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing

    '## ?????? ?????? ??ư???? ????????, ?????? ???? ??????
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tblReady_Click(ByVal Col As Long, ByVal Row As Long)
    Dim objAccInfo  As clsIISAccInfo    '???????? Ŭ????
    Dim vBarNo      As Variant          'Spread?? ???ڵ???ȣ
    Dim strSpcYy    As String           '??ü????
    Dim lngSpcNo    As Long             '??ü??ȣ

    If Row = 0 Then Exit Sub

    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
    If vBarNo = "" Then Exit Sub
    
    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)

    '## tblResult, Label?? ????ǥ??
    Call SetLabel(objAccInfo)
    Call SetResult(objAccInfo)

    Set objAccInfo = Nothing
End Sub

Private Sub tblReady_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Set mPopup = New clsIISPopup
    With mPopup
        .AddMenu DELETE, "Delete"
        .AddMenu DELETEALL, "Delete All"
        .PopupMenus Me.hWnd
    End With
    Set mPopup = Nothing
End Sub

Private Sub tblComplete_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strQcFg As String   'QC????
    Dim strInfo As String   '?????? ?߰?????
    Dim strTemp As String
    Dim i       As Long

    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    With tblComplete
        .Row = Row

        .Col = TCompleteEnum.ccQcFg:    strQcFg = .Text
        If strQcFg = "0" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
        ElseIf strQcFg = "1" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
        End If

        For i = TCompleteEnum.ccResult To .DataColCnt
            .Col = i:   strTemp = .Text

            '## 1.0.1: ??????(2005-06-24)
            '   - ????ǥ?? ???׼???
            If tblResult.MaxRows <= tblResult.DataRowCnt Then
                tblResult.MaxRows = tblResult.MaxRows + 1
                tblResult.Row = tblResult.MaxRows
            Else
                tblResult.Row = tblResult.DataRowCnt + 1
            End If

            tblResult.Col = TResultEnum.ccTestNm:       tblResult.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
            tblResult.Col = TResultEnum.ccUnit:         tblResult.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
            tblResult.Col = TResultEnum.ccHLDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
            tblResult.Col = TResultEnum.ccDPDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
            tblResult.Col = TResultEnum.ccRef:          tblResult.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
            tblResult.Col = TResultEnum.ccLISResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
            tblResult.Col = TResultEnum.ccEqpResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
        Next i
    End With
End Sub

Private Sub MSComm_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long

    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long

            Buffer = MSComm.Input
            'Buffer = strDump
            
            Call mIntLib.WriteLog(Buffer, ccEqp)

            Debug.Print Buffer
            
            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case mIntLib.Phase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                Call mIntLib.ClearBuffer
                                mIntLib.Phase = 2
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                            Case ACK
                                If mIntLib.State = "Q" Then
                                    Call SendOrder
                                Else    '-- Edit by Sewon,Oh(2006.08.02)
                                    MSComm.Output = ACK
                                End If
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Call mIntLib.ClearBuffer
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                            Case STX
                                mIntLib.BufCnt = 1
                                Call mIntLib.ClearBuffer
                            Case ETB
                                mIntLib.IsETB = True
                                mIntLib.Phase = 3
                            Case ETX
                                mIntLib.BufCnt = mIntLib.BufCnt + 1
                                mIntLib.Phase = 3
                            Case EOT
                                mIntLib.Phase = 1
                            Case vbCr, vbLf
                            Case Else
                                If mIntLib.IsETB = False Then
                                    Call mIntLib.AddBuffer(BufChar)
                                Else
                                    mIntLib.IsETB = False
                                End If
                        End Select
                    Case 3      '## Transfer Phase
                        Select Case BufChar
                            Case vbCr
                            Case vbLf
                                mIntLib.Phase = 4
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                mIntLib.Phase = 2
                            Case EOT
                                Call EditRcvData
                                If mIntLib.State = "Q" Then
                                    mIntLib.SndPhase = 1
                                    mIntLib.FrameNo = 0
                                    MSComm.Output = ENQ
                                    Call mIntLib.WriteLog(ENQ, ccPCLog)
                                End If
                                mIntLib.Phase = 1
'                                MSComm.Output = ACK
                        End Select
                End Select
            Next i

        Case comEvSend
        Case comEvCTS
            EVMsg$ = "CTS ???? ????"
        Case comEvDSR
            EVMsg$ = "DSR ???? ????"
        Case comEvCD
            EVMsg$ = "CD ???? ????"
        Case comEvRing
            EVMsg$ = "??ȭ ???? ?︮?? ??"
        Case comEvEOF
            EVMsg$ = "EOF ????"

        '???? ?޽???
        Case comBreak
            ERMsg$ = "?ߴ? ??ȣ ????"
        Case comCDTO
            ERMsg$ = "?ݼ??? ???? ?ð? ?ʰ?"
        Case comCTSTO
            ERMsg$ = "CTS ?ð? ?ʰ?"
        Case comDCB
            ERMsg$ = "DCB ?˻? ????"
        Case comDSRTO
            ERMsg$ = "DSR ?ð? ?ʰ?"
        Case comFrame
            ERMsg$ = "?????̹? ????"
        Case comOverrun
            ERMsg$ = "?и?Ƽ ????"
        Case comRxOver
            ERMsg$ = "???? ???? ?ʰ?"
        Case comRxParity
            ERMsg$ = "?и?Ƽ ????"
        Case comTxFull
            ERMsg$ = "???? ???ۿ? ?????? ????"
        Case Else
            ERMsg$ = "?? ?? ???? ???? ?Ǵ? ?̺?Ʈ"
    End Select

    If Len(EVMsg$) Then
        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
        StatusBar.Panels(2).Text = ERMsg$
    End If
End Sub

Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '???????̽? ??ü???? Ŭ????
    Dim objIntNms    As clsIISIntNms     '?????? ?˻??׸? ?÷??? Ŭ????
    Dim objBuffer    As clsIISBuffer     '????Ŭ????

    Dim strRcvBuf    As String   '?????? Data
    Dim strType      As String   '?????? Record Type
    Dim strBarNo     As String   '?????? ???ڵ???ȣ
    Dim strSeg       As String   '?????? Segment
    Dim strPos       As String   '?????? Position
    Dim strIntBase   As String   '?????? ???????? ?˻???
    Dim strIntResult As String   '?????? ????
    Dim strFlag      As String   '?????? ????,???? ????(F:????,I:????)
    Dim strResult    As String   'LIS????
    Dim strTemp      As String
    Dim strEqpCd     As String

    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strType = Mid$(strRcvBuf, 2, 1)

        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "Q"    '## Request Information
                '## ???ڵ???ȣ ??ȸ
                strBarNo = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))

                With mOrder
                    .ClsClear
                    .BarNo = strBarNo
                End With
                Call GetOrder(strBarNo)
                mIntLib.State = "Q"

            Case "O"    '## Order
                strTemp = mGetP(strRcvBuf, 4, "|")
                strBarNo = mGetP(strTemp, 1, "^")
                strSeg = mGetP(strTemp, 2, "^")
                strPos = mGetP(strTemp, 3, "^")
                
                '3O|1|18001305818|18001305818^T592^3|^^^639^_HIV Ag/Ab^UNDILUTED^P|
                '3O|1|18001304191|18001304191^T597^5|^^^561^Syphilis^UNDILUTED^P|
                
                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = strPos & "/" & strSeg
                    strRackPos = .SpcPos
                End With
                
            '-- 2019.12.19 ??????
            '-- ????(R) ä?ο? ?ִ? ?ڸ?Ʈ?? Architect ??????.
            'R ???? IGbn:= TokenStr(sData, '|', 14);  ?????? ???Ϳ?
            '1?????? : Ai04026
            '2?????? : Ai04025
            '4R|1|^^^598^CA19-9XR^UNDILUTED^F|6.62|U/mL||||F||Admin^Admin||20191220093826|Ai04026

        
            Case "R"    '## Result
                strTemp = mGetP(strRcvBuf, 3, "|")
                strIntBase = mGetP(strTemp, 4, "^")
                strFlag = mGetP(strTemp, 7, "^")
                strIntResult = mGetP(strRcvBuf, 4, "|")
                
                '2020.01.08 : ???񺰷? ?????Ͽ? ???????? ???????? ????
'                '???񺰷? ?????Ͽ? ????
'                strEqpCd = mGetP(strRcvBuf, 14, "|")
'
'                If strEqpCd = gEqp1 Then
'                    mEqpCd = "E122" '1?????? Ai04026
'                    EqpCd = "E122"
'                Else
'                    mEqpCd = "E123" '2?????? Ai04025
'                    EqpCd = "E123"
'                End If
                
                Select Case strFlag
                    Case "F"    '## ????
                        strIntBase = strIntBase & "N"
                        strResult = strIntResult
                    Case "I"    '## ????
                        strIntBase = strIntBase & "C"
                        '-- Edit by Sewon,Oh(2007.08.02)
                        'Anti-HCV(???ֿ??????? ?䱸????)
                        '???????????? ???????? ?ʰ? ??ü?????? ?????Ѵ?.
                        If strIntBase = "385C" Then
                            Select Case CDbl(strResult)
                                '2018.09.17 ????
                                Case Is < 1: strResult = "N"
                                Case 1 To 5: strResult = "WeaklyPositive"
                                Case Is > 5: strResult = "P"
                            End Select
                            strIntResult = strResult
                        '2018.11.15 ?߰? : HAVIgM
                        ElseIf strIntBase = "800C" Then
                            Select Case CDbl(strResult)
                                Case Is < 0.8: strResult = "N"
                                Case 0.8 To 1.21: strResult = "WeaklyPositive"
                                Case Is > 1.21: strResult = "P"
                            End Select
                            strIntResult = strResult
                        Else
                            Select Case Mid$(strIntResult, 1, 1)
                                '2018.09.17 ????
                                Case "N":   strResult = "N"
                                Case "G":   strResult = "Grayzone"
                                Case "R":   strResult = "P"
                                Case "P":   strResult = "P"
                                
                            End Select
                        End If
                    

                    '-- Edit by Sewon,Oh(2016.10.26)
                    'HbsAb2(Anti-Hbe)
                    '???????????? ??? ?????????? ?????Ѵ?.
                    '>=10 : Positive
                    '< 10 : Negative
                    Case "P"
                        'HbsAb2(Anti-Hbe)
                        If strIntBase = "694" Then
                            strIntBase = strIntBase & "C"
                            If IsNumeric(strResult) Then
                                Select Case CCur(strResult)
                                    Case Is < 10:  strResult = "Negative"
                                    Case Is >= 10: strResult = "Positive"
                                    '2018.09.17 ????
                                    Case Is < 10:  strResult = "N"
                                    Case Is >= 10: strResult = "P"
                                End Select
                                strIntResult = strResult
                            Else
                                'strResult = "Positive"
                                '2018.09.17 ????
                                strResult = "P"
                                strIntResult = strResult
                            End If
                        End If
                        
                        
                End Select
                
                '?ŵ??˻?
                If strIntBase = "565C" Then
                    If UCase(strResult) = "NEGATIVE" Then
                        '2018.09.17 ????
                        strResult = "N"
                    ElseIf UCase(strResult) = "POSITIVE" Then
                        'strResult = "Reactive"
                        '2018.09.17 ????
                        strResult = "P"
                    End If
                Else
                    If UCase(strResult) = "NONREACTIVE" Then
                        'strResult = "Negative"
                        '2018.09.17 ????
                        strResult = "N"
                    ElseIf UCase(strResult) = "REACTIVE" Then
                        'strResult = "Positive"
                        '2018.09.17 ????
                        strResult = "P"
                    End If
                
                    If UCase(strIntResult) = "NONREACTIVE" Then
                        'strIntResult = "Negative"
                        '2018.09.17 ????
                        strIntResult = "N"
                    ElseIf UCase(strIntResult) = "REACTIVE" Then
                        'strIntResult = "Positive"
                        '2018.09.17 ????
                        strIntResult = "P"
                    End If
                End If
                
                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                         strIntResult, strResult, strRackPos)
                    mIntLib.State = "R"
                End If

            Case "L"    '## Terminator
                '## DB?? ????????
                If mIntLib.State = "R" Then
                    Call SaveServer(objIntInfo)
                    Set objIntInfo = Nothing
                    mIntLib.State = ""
                End If
        End Select
    Next
    Set objIntNms = Nothing
    Set objBuffer = Nothing
End Sub

Private Sub RcvSocketData(ByVal pRcvData As String)
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = pRcvData
    
    Call mIntLib.WriteLog(Buffer, ccEqp)

    'Debug.Print Buffer
    
    lngBufLen = Len(Buffer)
    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)

        Select Case mIntLib.Phase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        mIntLib.Phase = 2
                        wSck.SendData (ACK)
                        
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                    Case ACK
                        If mIntLib.State = "Q" Then
                            Call SendOrder
                        Else    '-- Edit by Sewon,Oh(2006.08.02)
                            wSck.SendData (ACK)
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        wSck.SendData (ACK)
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                    Case STX
                        mIntLib.BufCnt = 1
                        Call mIntLib.ClearBuffer
                    Case ETB
                        mIntLib.IsETB = True
                        mIntLib.Phase = 3
                    Case ETX
                        mIntLib.BufCnt = mIntLib.BufCnt + 1
                        mIntLib.Phase = 3
                    Case EOT
                        mIntLib.Phase = 1
                    Case vbCr, vbLf
                    Case Else
                        If mIntLib.IsETB = False Then
                            Call mIntLib.AddBuffer(BufChar)
                        Else
                            mIntLib.IsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        mIntLib.Phase = 4
                        wSck.SendData (ACK)
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        mIntLib.Phase = 2
                    Case EOT
                        Call EditRcvData
                        If mIntLib.State = "Q" Then
                            mIntLib.SndPhase = 1
                            mIntLib.FrameNo = 0
                            wSck.SendData (ENQ)
                            Call mIntLib.WriteLog(ENQ, ccPCLog)
                        End If
                        mIntLib.Phase = 1
                End Select
        End Select
    Next i
End Sub
'-----------------------------------------------------------------------------'
'   ???? : ????????, ????????, ȭ??ǥ??
'   ?μ? :
'       - pIntInfo : ???????̽? ??ü???? Ŭ????
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '???????? Ŭ????
    Dim vBarNo      As Variant 'Spread?? ???ڵ???ȣ
    Dim strBarNo    As String  '???ڵ???ȣ
    Dim strSpcYy    As String  '??ü????
    Dim lngSpcNo    As Long    '??ü??ȣ
    Dim i           As Long

    Me.MousePointer = vbHourglass

    strBarNo = pIntInfo.BarNo

    '## ????????
    If mIntLib.CheckResult(pIntInfo) = -1 Then
        '## ?????????? ?????? ????ǥ??
        Call SetComplete1(pIntInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        '## ?????????? ?????? ????ǥ??
        strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)

        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing

        '## ClientDb, Server?? ????????
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        '190022942188
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        
        StatusBar.Panels(2).Text = "??ü??ȣ:" & strBarNo & " ?? ?????????? ???????? ?߽??ϴ?."
    End If

    '## tblReady???? ???۵? ??ü????
    With tblReady
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
            If Mid(CStr(vBarNo), 1, SPCYYLEN + SPCNOLEN) = Mid(strBarNo, 1, SPCYYLEN + SPCNOLEN) Then
                Call .DeleteRows(i, 1)
                Exit For
            End If
        Next i
    End With

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   ???? : ?ش? ???ڵ???ȣ?? ???? ???????? ??ȸ, tblReady, tblResult?? ǥ??
'   ?μ? :
'       - pBarNo : ???ڵ???ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '???????? Ŭ????
    Dim strOutput  As String            '?۽??? ??????

    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label?? ????ǥ??
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)

        mOrder.AccInfo = objAccInfo
        mOrder.NoOrder = False
        Set objAccInfo = Nothing
    Else
        mOrder.NoOrder = True
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ???? : ???????? ????
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '?۽??? ??????
    
    Select Case mIntLib.SndPhase
        Case 1  '## Header
            strOutput = mIntLib.GetFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            'slSender.Add('1H|\^&||||||||||P|1|'+CR+ETX);
            
            '## ???????? ?????? ?Ǵ??Ͽ? SndPhase????
            If mOrder.NoOrder = True Then
                '## ?????????? ???°???
                mIntLib.SndPhase = 3
            Else
                mIntLib.SndPhase = 2
            End If

        Case 2  '## Patient
            'strOutput = mIntLib.GetFrameNo & "P|1||" & mOrder.AccInfo.PtId & "||" & vbCr & ETX
            strOutput = mIntLib.GetFrameNo & "P|1||" & mOrder.AccInfo.PtId & "|" & Han2EngOCX1.HanToEng(mOrder.AccInfo.Name) & "|" & vbCr & ETX
            'slSender.Add('2P|1|'+TMaster.FPID+'|||'+CR+ ETX);
            mIntLib.SndPhase = 4

        Case 3  '## No Order
            strOutput = mIntLib.GetFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
            mIntLib.SndPhase = 5

        Case 4  '## Order
            With mOrder
                If .IsSending = False Then  '## ???? ??????
                    strOutput = "O|1|" & .BarNo & "||" & .GetOrder & "|R||||||N||||||||||||||Q"
                    If Len(strOutput) > 230 Then
                        .IsSending = True
                        .Order = Mid$(strOutput, 231)
                        strOutput = mIntLib.GetFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        mIntLib.SndPhase = 4
                    Else
                        strOutput = mIntLib.GetFrameNo & strOutput & vbCr & ETX
                        mIntLib.SndPhase = 5
                    End If
                Else                        '## ???? ???ڿ??? ??????
                    strOutput = .Order
                    If Len(strOutput) > 230 Then
                        .Order = Mid$(strOutput, 231)
                        strOutput = mIntLib.GetFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        mIntLib.SndPhase = 4
                    Else
                        .IsSending = False
                        strOutput = mIntLib.GetFrameNo & strOutput & vbCr & ETX
                        mIntLib.SndPhase = 5
                    End If
                End If
            End With

        Case 5  '## Termianator
            strOutput = mIntLib.GetFrameNo & "L|1" & vbCr & ETX
            mIntLib.SndPhase = 6

        Case 6  '## EOT
            mIntLib.State = ""
            wSck.SendData (EOT)
            Call mIntLib.WriteLog(EOT, ccPCLog)
            Exit Sub
    End Select

    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
    Debug.Print strOutput
    wSck.SendData (strOutput)
    Call mIntLib.WriteLog(strOutput, ccPCLog)
End Sub

'-----------------------------------------------------------------------------'
'   ???? : tblReady?? ????ǥ??
'   ?μ? :
'       - pAccInfo : ???????? Ŭ????
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

        If pAccInfo.QcFg = "0" Then         '## ?Ϲݰ?ü
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        ElseIf pAccInfo.QcFg = "1" Then     '## QC??ü
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
        End If
        Call .SetActiveCell(1, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ???? : tblComplete?? ????ǥ?? (?????????? ??????)
'   ?μ? :
'       - pIntInfo : ???????̽? ??ü???? Ŭ????
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pIntInfo As clsIISIntInfo)
    Dim objIntResult As clsIISIntResult     '???????̽? ???? Ŭ????
    Dim i            As Long

    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = pIntInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pIntInfo.BarNo
        .Col = TCompleteEnum.ccSendCnt: .Text = pIntInfo.IntResults.Count

        For Each objIntResult In pIntInfo.IntResults
            If .MaxCols <= .DataColCnt Then
                .MaxCols = .MaxCols + 1
            End If
            .Col = TCompleteEnum.ccResult + i
            .ColHidden = True
            .Text = objIntResult.IntNm & DIV & objIntResult.IntResult
            i = i + 1
        Next
        Set objIntResult = Nothing

        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ???? : tblComplete?? ????ǥ?? (?????????? ??????)
'   ?μ? :
'       - pAccInfo : ???????? Ŭ????
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '???????? Ŭ????
    Dim objQCResult As clsIISQCResult   'QC???????? Ŭ????
    Dim i           As Long

    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        
        .Col = TCompleteEnum.ccNo:      .Text = strRackPos 'pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

        i = 0
        If pAccInfo.QcFg = "0" Then         '## ?Ϲݰ?ü
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.PtId
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.Name
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.Sex & " / " & pAccInfo.Age
            .Col = TCompleteEnum.ccDoctNm:  .Text = pAccInfo.OrdDoctNm
            .Col = TCompleteEnum.ccDeptNm:  .Text = pAccInfo.DeptNm
            .Col = TCompleteEnum.ccWardNm:  .Text = pAccInfo.WardNm
            .Col = TCompleteEnum.ccStatFg:  .Text = IIf(pAccInfo.StatFg = "1", "Y", "N")
            .Col = TCompleteEnum.ccSpcNm:   .Text = pAccInfo.SpcNm
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt

            For Each objResult In pAccInfo.Results
                If objResult.RstCd <> "" Then
                    If .MaxCols <= .DataColCnt Then
                        .MaxCols = .MaxCols + 1
                    End If
                    .Col = TCompleteEnum.ccResult + i
                    .ColHidden = True
                    .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                            DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & DIV & _
                            IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "")
                    i = i + 1
                End If
            Next
            Set objResult = Nothing
        ElseIf pAccInfo.QcFg = "1" Then     '## QC??ü
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.CtrlCd
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.LevelCd
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.LotNo
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt

            For Each objQCResult In pAccInfo.QCResults
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                .Text = objQCResult.IntNm.IntNm & DIV & objQCResult.IntResult & DIV & _
                        objQCResult.RstCd & DIV & objQCResult.Unit & DIV & objQCResult.RADiv
                i = i + 1
            Next
            Set objQCResult = Nothing
        End If
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ???? : tblResult ????ǥ??
'   ?μ? :
'       - pAccInfo : ???????? Ŭ????
'-----------------------------------------------------------------------------'
Private Sub SetResult(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '???????? Ŭ????
    Dim objQCResult As clsIISQCResult   'QC???????? Ŭ????

    Call mTblClear(tblResult)
    If pAccInfo.QcFg = "0" Then         '## ?Ϲݰ?ü
        For Each objResult In pAccInfo.Results
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If

                .Col = TResultEnum.ccTestNm:    .Text = objResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objResult.RstCd
                .Col = TResultEnum.ccUnit:      .Text = objResult.Unit
                .Col = TResultEnum.ccHLDiv:     .Text = objResult.HLDiv
                .Col = TResultEnum.ccDPDiv:     .Text = objResult.DPDiv
                .Col = TResultEnum.ccRef
                .Text = IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "")
            End With
        Next
    ElseIf pAccInfo.QcFg = "1" Then     '## QC??ü
        For Each objQCResult In pAccInfo.QCResults
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If

                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
            End With
        Next
    End If

    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ???? : Label?? ȯ??????, ???????? ǥ??
'   ?μ? :
'       - pAccInfo : ???????? Ŭ????
'-----------------------------------------------------------------------------'
Private Sub SetLabel(ByVal pAccInfo As clsIISAccInfo)
    Call CtlClear(ccLabel)

    If pAccInfo.QcFg = "0" Then         '## ?Ϲݰ?ü
        Call LabelShow("0")
        lblPtId.Caption = pAccInfo.PtId
        lblName.Caption = pAccInfo.Name
        lblSexAge.Caption = pAccInfo.Sex & " / " & pAccInfo.Age
        lblDoctNm.Caption = pAccInfo.OrdDoctNm
        lblDeptNm.Caption = pAccInfo.DeptNm
        lblWardNm.Caption = pAccInfo.WardNm
        lblStatFg.Caption = IIf(pAccInfo.StatFg = "1", "Y", "N")
        lblSpcNm.Caption = pAccInfo.SpcNm
    ElseIf pAccInfo.QcFg = "1" Then     '## QC??ü
        Call LabelShow("1")
        lblPtId.Caption = pAccInfo.CtrlCd
        lblName.Caption = pAccInfo.LevelCd
        lblSexAge.Caption = pAccInfo.LotNo
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ???? : ??ü?????? ???? Label?? ?ٸ??? ǥ??
'   ?μ? :
'       - pQcFg : 0(?Ϲݰ?ü), 1(QC??ü)
'-----------------------------------------------------------------------------'
Private Sub LabelShow(ByVal pQcFg As String)
    Dim i As Long

    If pQcFg = "0" Then         '## ?Ϲݰ?ü
        lblControl.Caption = "ȯ  ?? ID :"
        lblLevel.Caption = "??     ?? :"
        lblLotNo.Caption = "????/???? :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = True
        Next i

        lblDoctNm.Visible = True:   lblDeptNm.Visible = True
        lblWardNm.Visible = True:   lblStatFg.Visible = True
        lblSpcNm.Visible = True
    ElseIf pQcFg = "1" Then     '## QC??ü
        lblControl.Caption = "Control :"
        lblLevel.Caption = "Level   :"
        lblLotNo.Caption = "Lot No  :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = False
        Next i

        lblDoctNm.Visible = False:   lblDeptNm.Visible = False
        lblWardNm.Visible = False:   lblStatFg.Visible = False
        lblSpcNm.Visible = False
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ???? : ???????? ??????ȸ, ??Ʈ Open
'-----------------------------------------------------------------------------'
Private Sub GetEqpComm()
    Dim objComm     As clsIISEqpComm    '???ż??? Ŭ????
    Dim strErrMsg   As String           '?????޽???

    '## ???ż??? ??????ȸ
    Set objComm = mIntLib.GetEqpComm
    If objComm Is Nothing Then Exit Sub

    With objComm
        MSComm.CommPort = .Port
        MSComm.Settings = .GetSettings
    End With
    Set objComm = Nothing

On Error GoTo Errors
    '## ??Ʈ Open
    With MSComm
        '## ?̹? ??Ʈ?? ????????
        If .PortOpen Then
            strErrMsg = mEqpCd & " ?????? ??????Ʈ?? ?̹? ?????ֽ??ϴ?."
            Error.SetLog App.EXEName, "frmIISAlinity", "GetEqpComm", strErrMsg, Now
            Call mIntLib_EqpError("E004")
            Exit Sub
        End If

        .RThreshold = 1
        .SThreshold = 1
        .RTSEnable = True
        .PortOpen = True
    End With

    '## ???????? ?????????? ????
    Call mIntLib.DelHistoryData
    Exit Sub

Errors:
    '## ?ٸ? ??ġ???? ??Ʈ?? ?????ϴ? ????
    If Err.Number = 8005 Then
        strErrMsg = mEqpCd & " ?????? ?????? ??Ʈ?? ?̹? ???????Դϴ?."
        Error.SetLog App.EXEName, "frmIISAlinity", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ???? : ??Ʈ?? ?ʱ?ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear(Optional ByVal pFlag As ClearEnum = ccAll)
    lblPtId.Caption = "":       lblName.Caption = ""
    lblSexAge.Caption = "":     lblDoctNm.Caption = ""
    lblDeptNm.Caption = "":     lblWardNm.Caption = ""
    lblStatFg.Caption = "":     lblSpcNm.Caption = ""

    If pFlag = ccAll Then
        Call mTblClear(tblReady):   Call mTblClear(tblComplete)
        Call mTblClear(tblResult)
    End If
End Sub

'------------------------------------------------------------------'
'   ???? : ???????? ???? ????ó??
'------------------------------------------------------------------'
Private Sub mIntLib_EqpError(ByVal pCode As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey)
End Sub

'------------------------------------------------------------------'
'   ???? : ??ü???? ????ó??1
'------------------------------------------------------------------'
Private Sub mIntLib_SpcError(ByVal pCode As String, ByVal pBarNo As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo)
End Sub

'------------------------------------------------------------------'
'   ???? : ??ü???? ????ó??2
'------------------------------------------------------------------'
Private Sub mIntLib_SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo, pPtId, pName)
End Sub

'------------------------------------------------------------------'
'   ???? : Popup ?޴? Click ?̺?Ʈ
'------------------------------------------------------------------'
Private Sub mPopup_Click(ByVal vMenuID As Long)
    Dim vBarNo      As Variant  'Spread?? ???ڵ???ȣ
    Dim strSpcYy    As String   '??ü????
    Dim lngSpcNo    As Long     '??ü??ȣ

    Select Case vMenuID
        Case DELETE     '## Delete
            With tblReady
                Call .GetText(TReadyEnum.ccBarNo, .ActiveRow, vBarNo)
                If vBarNo <> "" Then
                    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
                    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
                    Call mIntLib.AccInfos.Remove(strSpcYy, lngSpcNo)
                    Call .DeleteRows(.ActiveRow, 1)
                    Call mTblClear(tblResult)
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

Private Sub wSck_Close()
    
    If wSck.State <> sckClosed Then
        wSck.Close
    End If
    wSck.LocalPort = CInt(gSckPort)
    wSck.Listen

    StatusBar.Panels(2).Text = ""

End Sub

Private Sub wSck_ConnectionRequest(ByVal requestID As Long)
    
    If wSck.State <> sckClosed Then
        wSck.Close
    End If
    
    wSck.Accept requestID
    
    StatusBar.Panels(2).Text = mEqpKey & " : " & gSckPort & "?? ??Ʈ?? ?????Ǿ????ϴ?"

End Sub

Private Sub wSck_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvBuffer As String
   
    wSck.GetData strRcvBuffer

    Call RcvSocketData(strRcvBuffer)

End Sub


