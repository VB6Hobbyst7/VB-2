VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{D74ED2A2-3650-4720-93BC-FDDD8DCBC769}#1.0#0"; "Han2EngOCX.ocx"
Begin VB.Form frmIISCobas8000 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cobas8000"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows �⺻��
   Begin HAN2ENGOCXLib.Han2EngOCX Han2Eng 
      Height          =   495
      Left            =   8190
      TabIndex        =   31
      Top             =   8130
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   9060
      TabIndex        =   30
      Top             =   8010
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6540
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� ȯ������"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8E4D8&
      Height          =   1290
      Left            =   6540
      TabIndex        =   5
      Top             =   390
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   6
         Top             =   165
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   9450
         TabIndex        =   7
         Top             =   165
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�̻��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStatFg 
         Height          =   315
         Left            =   5595
         TabIndex        =   8
         Top             =   165
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   315
         Left            =   1245
         TabIndex        =   9
         Top             =   525
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�̻�� �Ʊ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   9450
         TabIndex        =   10
         Top             =   525
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   5595
         TabIndex        =   11
         Top             =   525
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TabIndex        =   12
         Top             =   885
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "���� / 29"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   9450
         TabIndex        =   13
         Top             =   885
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "65����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAreaNm 
         Height          =   315
         Left            =   5595
         TabIndex        =   14
         Top             =   885
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�������"
         Appearance      =   0
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "ȯ  �� ID :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   23
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "��     �� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   22
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "����/���� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   21
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "ó���� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   8625
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "����� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   8625
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "��  �� : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   8625
         TabIndex        =   18
         Top             =   975
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "���޿��� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "�� ü �� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   4560
         TabIndex        =   16
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "�Ƿ����� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   4560
         TabIndex        =   15
         Top             =   975
         Width           =   900
      End
   End
   Begin IISCobas8000.sckStringData sck 
      Height          =   300
      Left            =   11610
      TabIndex        =   3
      Top             =   8310
      Visible         =   0   'False
      Width           =   660
      _extentx        =   1164
      _extenty        =   529
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7710
      Top             =   7860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   495
      Left            =   11475
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      Height          =   495
      Left            =   12694
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   13913
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   8567
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6585
      Top             =   7770
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7185
      Top             =   7845
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3555
      Left            =   90
      TabIndex        =   24
      Top             =   510
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   6271
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      ShadowColor     =   16773087
      SpreadDesigner  =   "frmIISCobas8000.frx":0000
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   90
      TabIndex        =   25
      Top             =   4650
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   7779
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      ShadowColor     =   16773087
      SpreadDesigner  =   "frmIISCobas8000.frx":04CB
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   90
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻��� ����Ʈ"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6690
      Left            =   6540
      TabIndex        =   27
      Top             =   1695
      Width           =   8580
      _Version        =   393216
      _ExtentX        =   15134
      _ExtentY        =   11800
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      ShadowColor     =   16773087
      SpreadDesigner  =   "frmIISCobas8000.frx":0D37
      TextTip         =   2
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4230
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻�Ϸ� ����Ʈ"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblStatus 
      Height          =   315
      Left            =   6540
      TabIndex        =   29
      Top             =   8700
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   ""
      Appearance      =   0
   End
End
Attribute VB_Name = "frmIISCobas8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISCobasE602.frm
'   �ۼ���  : ������
'   ��  ��  : CobasE602 �����
'   �ۼ���  :
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady�� Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
End Enum

'## tblComplete�� Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccSexAge = 6
    ccDoctNm = 7:       ccDeptNm = 8
    ccWardNm = 9:       ccStatFg = 10
    ccSpcNm = 11:       ccQcFg = 12
    ccSendCnt = 13:     ccResult = 14
End Enum

'## tblResult�� Column Enum
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

'## Datalog Length ���
Private Const RACKNOLEN     As Long = 4
Private Const CUPPOSLEN     As Long = 2
Private Const SAMPLENOLEN   As Long = 4
Private Const TESTCDLEN     As Long = 2
Private Const RESULTLEN     As Long = 6
Private Const FLAGSLEN      As Long = 2

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '�������̽� Ŭ����
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '�˾��޴�
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors          '�������̽� ���� �÷���
Private mOrder      As clsIISIntOrder           '�������� Ŭ����

Private mEqpCd  As String   '����ڵ�
Private mEqpKey As String   '���Ű

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Command1_Click()
    Dim strData As String
    

    
              strData = "MSH|^~\&|cobas 8000||host||20151105094703||OUL^R22|26634||2.5||||NE||UNICODE UTF-8|" & vbCr
    strData = strData & "PID|1||||^||||" & vbCr
    strData = strData & "SPM||501||S1||not|||||P|||^^^^|||20151103112702|||||||||||" & vbCr
    strData = strData & "SAC||||||||||50022|1|" & vbCr
    strData = strData & "OBR|1|||53|" & vbCr
    strData = strData & "TQ1|1||||||||R|" & vbCr
    strData = strData & "OBX|1||53||14.67|U/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151103110339|bmserv^SYSTEM||0|e602^1^MU1#e602#3#1^9|20151103112206|" & vbCr
    strData = strData & "TCD|53|1|" & vbCr
    strData = strData & "SID|53^^53^ASY^1^0|185357|123910|" & vbCr
    strData = strData & "NTE|1|I|0|I|" & vbCr
    strData = strData & "OBR|2|||67|" & vbCr
    strData = strData & "TQ1|1||||||||R|" & vbCr
    strData = strData & "OBX|2||67||227.7|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151103110834|bmserv^SYSTEM||0|e602^1^MU1#e602#4#1^12|20151103112702|" & vbCr
    strData = strData & "TCD|67|1|" & vbCr
    strData = strData & "SID|67^^67^ASY^4^0|187143|40268|" & vbCr
    strData = strData & "NTE|1|I|0|I|" & vbCr
    strData = strData & "OBR|3|||127|" & vbCr
    strData = strData & "TQ1|1||||||||R|" & vbCr
    strData = strData & "OBX|3||127||0.212|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151103110813|bmserv^SYSTEM||0|e602^2^MU1#e602#4#2^13|20151103112642|" & vbCr
    strData = strData & "TCD|127|1|" & vbCr
    strData = strData & "SID|127^^127^ASY^8^0|185926|43971|" & vbCr
    strData = strData & "NTE|1|I|0|I|" & vbCr
    
strData = ""
strData = strData & "MSH|^~\&|cobas 8000||host||20151105105156||OUL^R22|26980||2.5||||NE||UNICODE UTF-8|" & vbCr
strData = strData & "PID|1||||^||||" & vbCr
strData = strData & "SPM||321||S1||not|||||P|||^^^^|||20151102153853|||||||||||" & vbCr
strData = strData & "SAC||||||||||50009|1|" & vbCr
strData = strData & "OBR|1|||52|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|1||52||14.89|U/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102150829|bmserv^SYSTEM||0|e602^1^MU1#e602#4#1^12|20151102152717|" & vbCr
strData = strData & "TCD|52|1|" & vbCr
strData = strData & "SID|52^^52^ASY^6^0|187902|36423|" & vbCr
strData = strData & "NTE|1|I|0|I|" & vbCr
strData = strData & "OBR|2|||63|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|2||63||1.05|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102150911|bmserv^SYSTEM||0|e602^2^MU1#e602#4#2^13|20151102152737|" & vbCr
strData = strData & "TCD|63|1|" & vbCr
strData = strData & "SID|63^^63^ASY^4^0|186661|19308|" & vbCr
strData = strData & "NTE|1|I|0|I|" & vbCr
strData = strData & "OBR|3|||68|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|3||68||0.155|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102150808|bmserv^SYSTEM||0|e602^1^MU1#e602#4#1^12|20151102152637|" & vbCr
strData = strData & "TCD|68|1|" & vbCr
strData = strData & "SID|68^^68^ASY^14^0|183535|84549|" & vbCr
strData = strData & "NTE|1|I|0|I|" & vbCr
strData = strData & "OBR|4|||127|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|4||127||0.255|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102150953|bmserv^SYSTEM||0|e602^2^MU1#e602#4#2^13|20151102152822|" & vbCr
strData = strData & "TCD|127|1|" & vbCr
strData = strData & "SID|127^^127^ASY^3^0|185926|42335|" & vbCr
strData = strData & "NTE|1|I|0|I|" & vbCr
strData = strData & "OBR|5|||155|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|5||155||4.90|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102151035|bmserv^SYSTEM||0|e602^2^MU1#e602#4#2^13|20151102152902|" & vbCr
strData = strData & "TCD|155|1|" & vbCr
strData = strData & "SID|155^^155^ASY^2^0|188301|16306|" & vbCr
strData = strData & "NTE|1|I|0|I|" & vbCr
strData = strData & "OBR|6|||178|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|6||178||0.005|ng/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102151220|bmserv^SYSTEM||0|e602^1^MU1#e602#4#1^12|20151102153047|" & vbCr
strData = strData & "TCD|178|1|" & vbCr
strData = strData & "SID|178^^178^ASY^11^0|187549|8176|" & vbCr
strData = strData & "NTE|1|I|0|I|" & vbCr
strData = strData & "OBR|7|||266|" & vbCr
strData = strData & "TQ1|1||||||||R|" & vbCr
strData = strData & "OBX|7||266||0.300|IU/L|^TECH~^NORM~^CRIT~^USER||||F|||20151102151117|bmserv^SYSTEM||0|e602^2^MU1#e602#4#2^13|20151102153852|" & vbCr
strData = strData & "TCD|266|1|" & vbCr
strData = strData & "SID|266^^266^ASY^8^0|186770|49386|" & vbCr
strData = strData & "NTE|1|I|27^Technical limit over (lower)|I|" & vbCr
strData = strData & ""


          strData = "MSH|^~\&|cobas 8000||host||20150310041213||TSREQ|307564||2.5||||ER||UNICODE UTF-8|" & vbCr
strData = strData & "QPD|TSREQ|307564|" & "Z1511050051" & "||50750|4||||S1|SC|R1|R|" & vbCr
strData = strData & "RCP|I|1|R|" & vbCr
strData = strData & ""

          strData = "MSH|^~\&|cobas 8000||host||20151106091636||TSREQ|29096||2.5||||ER||UNICODE UTF-8|" & vbCr
strData = strData & "QPD|TSREQ|29096|**********************||60345|1||||S1|SC|R1|R|" & vbCr
strData = strData & "RCP|I|1|R|" & vbCr
strData = strData & ""

Call RcvSocketDataHL7(strData)
    
    
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
    'Call GetEqpComm
    'Call GetWinSockComm
    
    DoEvents
    
'    TxtIP = Winsock1.LocalIP
    Winsock1.LocalPort = CLng(50003)
    Winsock1.Listen
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mOrder = Nothing
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISCobas8000 = Nothing
End Sub

Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '������ ǥ�� Ŭ����

    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing

    '## ������ ������ ��ư���� �������, ������ ��� ������
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
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim strSpcYy    As String           '��ü����
    Dim lngSpcNo    As Long             '��ü��ȣ

    If Row = 0 Then Exit Sub

    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
    If vBarNo = "" Then Exit Sub
    
    strSpcYy = vBarNo 'Mid$(vBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)

    '## tblResult, Label�� ����ǥ��
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
    Dim strQcFg As String   'QC����
    Dim strInfo As String   '������ �߰�����
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

            '## 1.0.2: �̻��(2005-06-20)
            '   - ȭ��ǥ�� ���׼���
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
            tblResult.Col = TResultEnum.ccInfo:
                strInfo = mGetP(strTemp, TResultEnum.ccInfo, DIV)
                If Trim$(strInfo) <> "" Then
                    tblResult.Text = vbCrLf & Space(2) & strInfo & vbCrLf
                End If
            tblResult.Col = TResultEnum.ccEqpResult
                tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
                If Trim$(strInfo) = "" Then
                    tblResult.ForeColor = vbBlack
                Else
                    tblResult.ForeColor = vbRed
                End If
        Next i
    End With
End Sub

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strInfo As String       '������ �߰�����
    
    If Row = 0 Then Exit Sub
    With tblResult
        .Row = Row: .Col = TResultEnum.ccInfo
        strInfo = .Text
        If Trim(strInfo) = "" Then Exit Sub
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strInfo
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
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
            Call mIntLib.WriteLog(Buffer, ccEqp)

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
                                If mIntLib.state = "Q" Then Call SendOrder
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
                            Case vbCr, vbLf
                            Case EOT
                                mIntLib.Phase = 1
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
                                If mIntLib.state = "Q" Then
                                    mIntLib.SndPhase = 1
                                    mIntLib.FrameNo = 0
                                    MSComm.Output = ENQ
                                    Call mIntLib.WriteLog(ENQ, ccPCLog)
                                End If
                                mIntLib.Phase = 1
                        End Select
                End Select
            Next i

        Case comEvSend
        Case comEvCTS
            EVMsg$ = "CTS ���� ����"
        Case comEvDSR
            EVMsg$ = "DSR ���� ����"
        Case comEvCD
            EVMsg$ = "CD ���� ����"
        Case comEvRing
            EVMsg$ = "��ȭ ���� �︮�� ��"
        Case comEvEOF
            EVMsg$ = "EOF ����"

        '���� �޽���
        Case comBreak
            ERMsg$ = "�ߴ� ��ȣ ����"
        Case comCDTO
            ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            ERMsg$ = "CTS �ð� �ʰ�"
        Case comDCB
            ERMsg$ = "DCB �˻� ����"
        Case comDSRTO
            ERMsg$ = "DSR �ð� �ʰ�"
        Case comFrame
            ERMsg$ = "�����̹� ����"
        Case comOverrun
            ERMsg$ = "�и�Ƽ ����"
        Case comRxOver
            ERMsg$ = "���� ���� �ʰ�"
        Case comRxParity
            ERMsg$ = "�и�Ƽ ����"
        Case comTxFull
            ERMsg$ = "���� ���ۿ� ������ ����"
        Case Else
            ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select

    If Len(EVMsg$) Then
        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
        StatusBar.Panels(2).Text = ERMsg$
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '�������̽� ��ü���� Ŭ����
    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
    Dim objBuffer    As clsIISBuffer     '����Ŭ����

    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���
    Dim strFlag      As String   '������ Abnormal Flag
    Dim strComm      As String   '������ Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strType = Mid$(strRcvBuf, 2, 1)
        
        Select Case strType
            Case "H"    '## Header
            Case "Q"    '## Request Information
            Case "P"    '## Patient
                strBarNo = Format$(mGetP(strRcvBuf, 3, "|"), String$(SPCLEN, "#"))
                '2P|1|150005776650|||||||||||||||||||||||49

                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = strTubePos & "/" & strRackNo
                End With
            Case "O"    '## Order
                '3O|1|||Flu A+B|||||||||||P4C
'                strBarNo = Format$(mGetP(strRcvBuf, 3, "|"), String$(SPCLEN, "#"))
'                Set objIntInfo = New clsIISIntInfo
'                With objIntInfo
'                    .BarNo = strBarNo
'                    .SpcPos = strTubePos & "/" & strRackNo
'                End With
            Case "R"    '## Result
                '## ������ �˻��, ���, Abnormal Flag
                strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                strTemp2 = mGetP(strRcvBuf, 4, "|")
                
                strIntBase = strTemp1
                strResult = strTemp2
                
                If strResult <> "" Then
                    If objIntNms.ExistIntBase(strIntBase) Then
                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                             strResult, strResult)
                    End If
                End If
                mIntLib.state = "R"
                
            Case "L"    '## Terminator
                '## DB�� �������
                If mIntLib.state = "R" Then
                    Call SaveServer(objIntInfo)
                    Set objIntInfo = Nothing
                    mIntLib.state = ""
                End If
        End Select
    Next
    Set objIntNms = Nothing
    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : euc-kr ��utf-8 �� ��ȯ
'-----------------------------------------------------------------------------'
Function utf8Encode(ByVal strEucKr As String) As String
     Dim i As Long
     Dim ansi() As Byte
     Dim ascii As Integer
     Dim encText As String
    
     ansi = StrConv(strEucKr, vbFromUnicode)
     encText = ""
    
     For i = 0 To UBound(ansi)
          ascii = ansi(i)
         
          Select Case ascii
               Case 48 To 57, 65 To 90, 97 To 122
                    encText = encText & Chr(ascii)
               Case 32
                    encText = encText & "+"
               Case Else
                    If ascii < 16 Then
                         encText = encText & "%0" & Hex(ascii)
                    Else
                         encText = encText & "%" & Hex(ascii)
                    End If
          End Select
     Next i
    
     utf8Encode = encText

End Function





'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataHL7()
    Dim objIntInfo   As clsIISIntInfo    '�������̽� ��ü���� Ŭ����
    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
    Dim objBuffer    As clsIISBuffer     '����Ŭ����

    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���
    Dim strFlag      As String   '������ Abnormal Flag
    Dim strComm      As String   '������ Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim strOrgBarNo  As String
    Dim strRackType  As String
    Dim strSmpCntType As String
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strType = Mid$(strRcvBuf, 1, 3)
        
        Select Case strType
            Case "MSH"    '## Header
            Case "QPD"    '## Request Information
                '-- Test selection inquiry for routine rack
                'QPD|TSREQ|15161|321070||50094|2||||S1|SC|R1|R|
                
                '-- Test selection inquiry for STAT rack
                'QPD|TSREQ|15164|321040||40002|3||||S1|SC|R1|S|
                
                '-- Routine rack (AL) with acknowledgment
                'QPD|TSREQ|15161|321070||50094|2||||S1|SC|R1|R|

                '------- Query ----------
                'MSH|^~\&|cobas 8000||host||20150310041213||TSREQ|307564||2.5||||ER||UNICODE UTF-8|
                'QPD|TSREQ|307564|506831016010||50750|4||||S1|SC|R1|R|
                'RCP|I|1|R|

                '## ���ڵ��ȣ ��ȸ
                strOrgBarNo = Trim$(mGetP(strRcvBuf, 4, "|"))
                strBarNo = Mid(strOrgBarNo, 1, 10)
                strRackNo = Trim$(mGetP(strRcvBuf, 6, "|"))
                strTubePos = Trim$(mGetP(strRcvBuf, 7, "|"))
                strRackType = Trim$(mGetP(strRcvBuf, 11, "|"))
                strSmpCntType = Trim$(mGetP(strRcvBuf, 12, "|"))
                
                If InStr(strOrgBarNo, "*") > 0 Then
                    strOrgBarNo = mIntLib.GetBarcodeNo("RFM", strRackNo, strTubePos)
                    strBarNo = strOrgBarNo
                End If
                
                With mOrder
                    .ClsClear
                    .BarNo = strBarNo
                    .OrgBarNo = strOrgBarNo
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RackType = strRackType
                    .SmpCntType = strSmpCntType
                End With
                
                Call GetOrder(strBarNo)
                mIntLib.state = "Q"
            
            
            Case "PID"    '## Patient
                'PID|1||||^||||
            
            Case "SPM"    '## Barcode
                'SPM||321||S1||not|||||P|||^^^^|||20151102153853|||||||||||

                strOrgBarNo = mGetP(strRcvBuf, 3, "|")
                strBarNo = Mid(strOrgBarNo, 1, 10)
                
            Case "SAC"    '## Rack/Pos
                'SAC||||||||||50009|1|

                strRackNo = mGetP(strRcvBuf, 11, "|")
                strTubePos = mGetP(strRcvBuf, 12, "|")
                
                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = strTubePos & "/" & strRackNo
                    .OrgBarNo = strOrgBarNo
                    .RackNo = strRackNo
                    .PosNo = strTubePos
                End With
            
            Case "OBX"    '## Result
                'OBX|1||52||14.89|U/mL|^TECH~^NORM~^CRIT~^USER||||F|||20151102150829|bmserv^SYSTEM||0|e602^1^MU1#e602#4#1^12|20151102152717|
                
                strTemp1 = mGetP(strRcvBuf, 4, "|")
                strTemp2 = mGetP(strRcvBuf, 6, "|")
                
                strIntBase = strTemp1
                strResult = strTemp2
                
                If strResult <> "" Then
                    If objIntNms.ExistIntBase(strIntBase) Then
                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                             strResult, strResult)
                    End If
                End If
                mIntLib.state = "R"
                
        End Select
    Next
    
    '## DB�� �������
    If mIntLib.state = "R" Then
        Call SaveServer(objIntInfo)
        Set objIntInfo = Nothing
        mIntLib.state = ""
    End If
    
    If mIntLib.state = "Q" Then
        'Call GetOrder(strBarNo)
        Call SendOrder
        Set objIntInfo = Nothing
        mIntLib.state = ""
    End If
    
    Set objIntNms = Nothing
    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������, �������, ȭ��ǥ��
'   �μ� :
'       - pIntInfo : �������̽� ��ü���� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim vBarNo      As Variant 'Spread�� ���ڵ��ȣ
    Dim strBarNo    As String  '���ڵ��ȣ
    Dim strSpcYy    As String  '��ü����
    Dim lngSpcNo    As Long    '��ü��ȣ
    Dim i           As Long

    Me.MousePointer = vbHourglass

    strBarNo = pIntInfo.BarNo

    '## �������
    If mIntLib.CheckResult(pIntInfo) = -1 Then
        '## ���������� ������ ���ǥ��
        Call SetComplete1(pIntInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        '## ���������� ������ ���ǥ��
        strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)

        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)

        '## ClientDb, Server�� �������
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        Set objAccInfo = Nothing
        
        StatusBar.Panels(2).Text = "��ü��ȣ:" & strBarNo & " �� ���������� ������� �߽��ϴ�."
    End If

    '## tblReady���� ���۵� ��ü����
    With tblReady
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
            'If CStr(vBarNo) = strBarNo Then
            If Mid(CStr(vBarNo), 1, 11) = Mid(strBarNo, 1, 11) Then
                Call .DeleteRows(i, 1)
                Exit For
            End If
        Next i
    End With

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, tblReady, tblResult�� ǥ��
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����
    Dim strOutput  As String            '�۽��� ������

    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label�� ����ǥ��
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '�۽��� ������
    Dim strOrder  As String
    Dim varOrder  As Variant
    Dim i         As Integer
    Dim j         As Integer

        'lsTmp = Chr(11) & "MSH|^~\&|SCC|SCC|LIS|cobasIT1000|" & Format(Date, "yyyymmdd") & Format(Time, "hhnnss") & "||ACK|ROCHE20101101|P|2.3" & Chr(13)
        'lsTmp = lsTmp & "MSA|AA|" & lsSendID & "|Message Ok" & Chr(28) & Chr(13)
        
    j = 0
                strOutput = Chr(11) & "MSH|^~\&|Host||cobas8k^3.0.5.21^cobas8k||" & Format(Date, "yyyymmdd") & Format(Time, "hhnnss") & "||OML^O33|3166032||2.5||||NE||UNICODE UTF-8" & Chr(13)
    strOutput = strOutput & "PID|1|" & mOrder.GetPtId & "|||" & Han2Eng.HanToEng(mOrder.GetPtNm) & "||" & mOrder.GetPtSsn & "|" & mOrder.GetPtSex & Chr(13)
    'strOutput = strOutput & "PID|1|" & mOrder.GetPtId & "|||" & "aaa" & "||" & mOrder.GetPtSsn & "|" & mOrder.GetPtSex & Chr(13)
    strOutput = strOutput & "SPM||" & mOrder.OrgBarNo & "||" & mOrder.RackType & "||not|||||P|||^^^^|||||||||||||" & mOrder.SmpCntType & Chr(13)
    strOutput = strOutput & "SAC||||||||||" & mOrder.RackNo & "|" & mOrder.TubePos & Chr(13)
    
    If mOrder.NoOrder = False Then
        strOrder = mOrder.GetOrder
        varOrder = Split(strOrder, "|")
        For i = 0 To UBound(varOrder)
            If varOrder(i) <> "" Then
                j = j + 1
                strOutput = strOutput & "TQ1|1||||||||R" & Chr(13)
                strOutput = strOutput & "OBR|" & j & "|||" & varOrder(i) & "^|||||||A" & Chr(13)
            Else
                Exit For
            End If
        Next
    End If
    
    strOutput = strOutput & Chr(28)
    j = 0
        
    'Call mIntLib.WriteLog(strOutput, ccPCLog)
'    strOutput = utf8Encode(strOutput)
'    strOutput = URLEncodeUTF8(strOutput)
    
    sck.ProcSendMessage strOutput
    
    'Call Winsock1.SendData(strOutput)
    Call mIntLib.WriteLog(strOutput, ccPCLog)

End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblReady�� ����ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    Dim lngWorkNo As Long   'WorkNo

    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        '.Col = TReadyEnum.ccNo:     .Text = mOrder.TubePos & "/" & Mid$(mOrder.RackNo, 2)
        .Col = TReadyEnum.ccNo:     .Text = mOrder.TubePos & "/" & mOrder.RackNo
        .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

        If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.DeptNm ' pAccInfo.PtId
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
        End If
        Call .SetActiveCell(1, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblComplete�� ����ǥ�� (���������� ������)
'   �μ� :
'       - pIntInfo : �������̽� ��ü���� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pIntInfo As clsIISIntInfo)
    Dim objIntResult As clsIISIntResult     '�������̽� ��� Ŭ����
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
            
            .Text = objIntResult.IntNm & DIV & objIntResult.Result & DIV & DIV & DIV & DIV & _
                    DIV & DIV & objIntResult.Info
            i = i + 1
        Next
        Set objIntResult = Nothing

        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblComplete�� ����ǥ�� (���������� ������)
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim vTubePos    As Variant          'Spread�� Tube Position
    Dim i           As Long

    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        
        .Col = TCompleteEnum.ccNo:      .Text = pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

        i = 0
        If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.PtId
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.Name
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
            .Col = TCompleteEnum.ccDoctNm:  .Text = pAccInfo.OrdDoctNm
            .Col = TCompleteEnum.ccDeptNm:  .Text = pAccInfo.DeptNm
            .Col = TCompleteEnum.ccWardNm:  .Text = pAccInfo.WardNm
            .Col = TCompleteEnum.ccStatFg:  .Text = IIf(pAccInfo.StatFg = "1", "Y", "N")
            .Col = TCompleteEnum.ccSpcNm:   .Text = pAccInfo.SpcNm
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt

            For Each objResult In pAccInfo.Results
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                        DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & _
                        DIV & mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal) & DIV & _
                        objResult.IntInfo
                i = i + 1
            Next
            Set objResult = Nothing
        ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
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
'   ��� : tblResult ����ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetResult(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����

    Call mTblClear(tblResult)
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
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
                .Col = TResultEnum.ccRef:       .Text = mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal)
            End With
        Next
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
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
'   ��� : Label�� ȯ������, �������� ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetLabel(ByVal pAccInfo As clsIISAccInfo)
    Call CtlClear(ccLabel)

    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        Call LabelShow("0")
        lblPtId.Caption = pAccInfo.PtId
        lblName.Caption = pAccInfo.Name
        lblSexAge.Caption = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
        lblDoctNm.Caption = pAccInfo.OrdDoctNm
        lblDeptNm.Caption = pAccInfo.DeptNm
        lblWardNm.Caption = pAccInfo.WardNm
        lblStatFg.Caption = IIf(pAccInfo.StatFg = "1", "Y", "N")
        lblSpcNm.Caption = pAccInfo.SpcNm
        lblAreaNm.Caption = pAccInfo.DeptNm
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        Call LabelShow("1")
        lblPtId.Caption = pAccInfo.CtrlCd
        lblName.Caption = pAccInfo.LevelCd
        lblSexAge.Caption = pAccInfo.LotNo
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��ü������ ���� Label�� �ٸ��� ǥ��
'   �μ� :
'       - pQcFg : 0(�Ϲݰ�ü), 1(QC��ü)
'-----------------------------------------------------------------------------'
Private Sub LabelShow(ByVal pQcFg As String)
    Dim i As Long

    If pQcFg = "0" Then         '## �Ϲݰ�ü
        lblControl.Caption = "ȯ  �� ID :"
        lblLevel.Caption = "��     �� :"
        lblLotNo.Caption = "����/���� :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = True
        Next i

        lblDoctNm.Visible = True:   lblDeptNm.Visible = True
        lblWardNm.Visible = True:   lblStatFg.Visible = True
        lblSpcNm.Visible = True
    ElseIf pQcFg = "1" Then     '## QC��ü
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
'   ��� : ������Ʈ Open
'-----------------------------------------------------------------------------'
Private Sub GetWinSockComm()

    If Winsock1.state <> sckClosed Then
        Winsock1.Close
    End If
      
    Winsock1.LocalPort = 10002
    Winsock1.Listen

'    lblWinSock.Caption = "Socket : Connecting.."

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ����� ������ȸ, ��Ʈ Open
'-----------------------------------------------------------------------------'
Private Sub GetEqpComm()
    Dim objComm     As clsIISEqpComm    '��ż��� Ŭ����
    Dim strErrMsg   As String           '�����޽���

    '## ��ż��� ������ȸ
    Set objComm = mIntLib.GetEqpComm
    If objComm Is Nothing Then Exit Sub

    With objComm
        MSComm.CommPort = .Port
        MSComm.Settings = .GetSettings
    End With
    Set objComm = Nothing

On Error GoTo Errors
    '## ��Ʈ Open
    With MSComm
        '## �̹� ��Ʈ�� �������
        If .PortOpen Then
            strErrMsg = mEqpCd & " ����� �����Ʈ�� �̹� �����ֽ��ϴ�."
            Error.SetLog App.EXEName, "frmIISElecsys2010", "GetEqpComm", strErrMsg, Now
            Call mIntLib_EqpError("E004")
            Exit Sub
        End If

        .RThreshold = 1
        .SThreshold = 1
        .RTSEnable = True
        .PortOpen = True
    End With

    '## �������� ���������� ����
    Call mIntLib.DelHistoryData
    Exit Sub

Errors:
    '## �ٸ� ��ġ���� ��Ʈ�� ����ϴ� ���
    If err.Number = 8005 Then
        strErrMsg = mEqpCd & " ��� ������ ��Ʈ�� �̹� ������Դϴ�."
        Error.SetLog App.EXEName, "frmIISElecsys2010", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ������ Abnormal Flag�� ���� ������ȸ
'-----------------------------------------------------------------------------'
Private Function GetContent(ByVal pFlags As String) As String
    Dim strFlag     As String   'Abnormal Flag
    Dim strContent  As String   'Abnormal Content
    Dim strTemp     As String
    Dim i           As Long
    
    If pFlags = "" Then Exit Function
    
    For i = 1 To Len(Trim(pFlags))
        strFlag = Mid$(pFlags, i, 1)
        If i = 1 Then
            strTemp = Space(2) & "[Abnormality flag]: " & strFlag & vbCrLf
        Else
            strTemp = strTemp & vbCrLf & Space(2) & "[Abnormality flag]: " & strFlag & vbCrLf
        End If
        
        strContent = ""
        Select Case strFlag
            Case "/": strContent = "Test no performed: test has been requisitioned but not performed due to any reason."
            Case "S": strContent = "Result extracted for repeat run"
            Case "?": strContent = "Calculation unable due to abnormal photometric data. UNIT in STOP mode (Incl. Lamp OFF), etc."
            Case "n": strContent = "8087 error"
            Case "R": strContent = "Reagent level detection error"
            Case "#": strContent = "Sample level detection error"
            Case "!": strContent = "A/D error of photometry"
            Case ">": strContent = "The absolute OD value is over 2.665."
            Case "<": strContent = "The absolute OD value is under 0.99."
            Case "-": strContent = "The final result is negative."
            Case "U": strContent = "Reagent absorbance value at P0 of Reagent Blank run, is smaller than the lower limit of the Parameter."
            Case "u": strContent = "Reagent absorbance value at P0 or p8 is lower than the lower limit specified in the Parameters in routine run."
            Case "Y": strContent = "Reagent absorbance value at P16 of Reagent Blank run, is greater than the upper limit of the Parameter."
            Case "y": strContent = "Reagent absorbance value at P0 or p8 is higher than the upper limit specified in the Parameters in routine run."
            Case "@": strContent = "Abnormally high result: absorbance of every wavelength is more than 2.5."
            Case "$": strContent = "No linearity validation conducted because less than 3 data obtained in the kinetics."
            Case "D": strContent = "Too quick reaction slope in increasing kinetics, absorbance at P-START is higher than MAX. OD in increasing FIXED assay, or too slow reaction slope in decreasing kinetics (=no reaction observed)"
            Case "B": strContent = "Too quick reaction slope in increasing kinetics, or absorbance at P-END is lower than MIN. OD in increasing FIXED assay."
            Case "*": strContent = "Linearity error in kinetics"
            Case "P": strContent = "Result higher than DECIDE RANGE designated in parameters."
            Case "N": strContent = "Result lower than DECIDE RANGE designated in parameters."
            Case "&": strContent = "Data check 2 error"
            Case "Z": strContent = "Data check 1 error"
            Case "F": strContent = "Result higher than the dynamic range specified in the Parameters"
            Case "G": strContent = "Result lower than the dynamic range specified in the Parameters"
            Case "p": strContent = "Result beyond the panic value specified in the Parameters"
            Case "T": strContent = "Abnormality found in the Inter-Item Check"
            Case "H": strContent = "Result higher than the normal value range specified in the Parameters"
            Case "L": strContent = "Result lower than the normal value range specified in the Parameters"
            Case "W": strContent = "Abnormality in WB data. Photocal has not been performed."
            Case "J": strContent = "Result higher than the repeat run range specified in the Parameters"
            Case "K": strContent = "Result higher than the repeat run range specified in the Parameters"
        End Select
        
        If strContent <> "" Then
            strTemp = strTemp & Space(2) & "[Content]: " & strContent & vbCrLf
        End If
    Next i
    
    GetContent = vbCrLf & strTemp
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear(Optional ByVal pFlag As ClearEnum = ccAll)
    lblPtId.Caption = "":       lblName.Caption = ""
    lblSexAge.Caption = "":     lblDoctNm.Caption = ""
    lblDeptNm.Caption = "":     lblWardNm.Caption = ""
    lblStatFg.Caption = "":     lblSpcNm.Caption = ""
    lblAreaNm.Caption = ""

    If pFlag = ccAll Then
        Call mTblClear(tblReady):   Call mTblClear(tblComplete)
        Call mTblClear(tblResult)
    End If
End Sub

'------------------------------------------------------------------'
'   ��� : ����� ���� ����ó��
'------------------------------------------------------------------'
Private Sub mIntLib_EqpError(ByVal pCode As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey)
End Sub

'------------------------------------------------------------------'
'   ��� : ��ü���� ����ó��1
'------------------------------------------------------------------'
Private Sub mIntLib_SpcError(ByVal pCode As String, ByVal pBarNo As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo)
End Sub

'------------------------------------------------------------------'
'   ��� : ��ü���� ����ó��2
'------------------------------------------------------------------'
Private Sub mIntLib_SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo, pPtId, pName)
End Sub

'------------------------------------------------------------------'
'   ��� : Popup �޴� Click �̺�Ʈ
'------------------------------------------------------------------'
Private Sub mPopup_Click(ByVal vMenuID As Long)
    Dim vBarNo      As Variant  'Spread�� ���ڵ��ȣ
    Dim strSpcYy    As String   '��ü����
    Dim lngSpcNo    As Long     '��ü��ȣ

    Select Case vMenuID
        Case DELETE     '## Delete
            With tblReady
                Call .GetText(TReadyEnum.ccBarNo, .ActiveRow, vBarNo)
                If vBarNo <> "" Then
                    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
                    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
                    Call mIntLib.AccInfos.Remove(strSpcYy, lngSpcNo)
                    Call .DeleteRows(.ActiveRow, 1)
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

'Private Sub Winsock1_Close()
'
'    lblWinSock.Caption = "Socket : Close!!!"
'    'Call Sleep(1000)
'    Call GetWinSockComm
'
'End Sub
'
'Private Sub Winsock1_Connect()
'    lblWinSock.Caption = "Socket : Connection OK!"
'End Sub
'
'Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'    Winsock1.Close
'    Winsock1.Accept requestID
'End Sub
'
Public Sub RcvSocketDataHL7(ByVal lsData As String)
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = lsData
    Call mIntLib.WriteLog(Buffer, ccEqp)
    
    lngBufLen = Len(Buffer)
    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)

        Select Case mIntLib.Phase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ""
                        Call mIntLib.ClearBuffer
                        mIntLib.BufCnt = 1
                        mIntLib.Phase = 2
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case vbLf
                    Case vbCr
                        mIntLib.BufCnt = mIntLib.BufCnt + 1
                    Case Else
                        If Asc(BufChar) = 28 Then
                        '    Exit For
                        Else
                            Call mIntLib.AddBuffer(BufChar)
                        End If
                End Select
        End Select
    Next i

    Call EditRcvDataHL7
    mIntLib.Phase = 1
    
    Exit Sub

ErrHandle:
    Resume Next
End Sub


Public Sub RcvSocketData(ByVal lsData As String)
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = lsData
    Call mIntLib.WriteLog(Buffer, ccEqp)

    lngBufLen = Len(Buffer)
    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)

        Select Case mIntLib.Phase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        mIntLib.Phase = 2
                        'Call Winsock1.SendData(ACK)
                        sck.ProcSendMessage ACK
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                    Case ACK
                        'If mIntLib.State = "Q" Then Call SendOrder
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        sck.ProcSendMessage ACK
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
                    Case vbCr, vbLf
                    Case EOT
                        mIntLib.Phase = 1
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
                        sck.ProcSendMessage ACK
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        mIntLib.Phase = 2
                    Case EOT
                        Call EditRcvData
                        'If mIntLib.State = "Q" Then
                        '    mIntLib.SndPhase = 1
                        '    mIntLib.FrameNo = 0
                        '    MSComm.Output = ENQ
                        '    Call mIntLib.WriteLog(ENQ, ccPCLog)
                        'End If
                        mIntLib.Phase = 1
                End Select
        End Select
    Next i

    'Winsock1.SendData lsTmp

    Exit Sub

ErrHandle:
    Resume Next
End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    sck.Accept requestID

    Winsock1.Close
    Winsock1.Listen
    
    
'    If sck.state <> sckClosed Then
'        sck.Accept requestID
'
'        Winsock1.Close
'        Winsock1.Listen
        lblStatus.Caption = "  ���� C8000�� ���ӵǾ����ϴ�"
'    End If
    
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvBuffer As String
    Dim strSndBuffer As String
   
    Winsock1.GetData strRcvBuffer
    Debug.Print strRcvBuffer


    strSndBuffer = "ORDER"
    Winsock1.SendData (strSndBuffer)

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Number & " >> " & Description
    Winsock1.Close
End Sub

