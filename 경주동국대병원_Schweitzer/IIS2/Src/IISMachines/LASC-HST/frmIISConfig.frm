VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIISConfig 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "LASC-HST 설정"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4185
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저 장(&S)"
      Height          =   495
      Left            =   1650
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   4331
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   2865
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   4331
      Width           =   1215
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   187
      MaxLength       =   3
      TabIndex        =   6
      Top             =   3911
      Width           =   2370
   End
   Begin VB.ComboBox cboBaud 
      BackColor       =   &H00F7FFF7&
      Height          =   300
      Left            =   1702
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   1189
      Width           =   2370
   End
   Begin VB.ComboBox cboStopbit 
      BackColor       =   &H00F7FFF7&
      Height          =   300
      Left            =   1702
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   2329
      Width           =   2370
   End
   Begin VB.ComboBox cboPort 
      BackColor       =   &H00F7FFF7&
      Height          =   300
      ItemData        =   "frmIISConfig.frx":0000
      Left            =   1702
      List            =   "frmIISConfig.frx":0002
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   619
      Width           =   2370
   End
   Begin VB.ComboBox cboDatabit 
      BackColor       =   &H00F7FFF7&
      Height          =   300
      Left            =   1702
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   1759
      Width           =   2370
   End
   Begin VB.ComboBox cboParity 
      BackColor       =   &H00F7FFF7&
      Height          =   300
      ItemData        =   "frmIISConfig.frx":0004
      Left            =   1702
      List            =   "frmIISConfig.frx":0006
      Style           =   2  '드롭다운 목록
      TabIndex        =   5
      Top             =   2899
      Width           =   2370
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   112
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   109
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "■ LASC-HST 통신설정"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3431
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "■ 오더전송 시간간격"
      Appearance      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "초"
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
      Left            =   2640
      TabIndex        =   15
      Top             =   3990
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "▶ Baud Rate"
      Height          =   180
      Left            =   187
      TabIndex        =   14
      Top             =   1249
      Width           =   1110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "▶ Stop bit"
      Height          =   180
      Left            =   187
      TabIndex        =   13
      Top             =   2404
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "▶ Parity bit"
      Height          =   180
      Left            =   187
      TabIndex        =   12
      Top             =   2974
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "▶ Data bit"
      Height          =   180
      Left            =   187
      TabIndex        =   11
      Top             =   1819
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "▶ Port"
      Height          =   180
      Left            =   187
      TabIndex        =   10
      Top             =   679
      Width           =   570
   End
End
Attribute VB_Name = "frmIISConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISConfig.frm
'   작성자  : 이상대
'   내  용  : LASC-HST 설정폼
'   작성일  : 2005-09-15
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqpKey As String   '장비키

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Form_Load()
    Call InitCombo
    
On Error Resume Next
    cboPort.Text = IIf(mPort = 0, "", CStr(mPort))
    cboBaud.Text = mBaudRate
    cboDatabit.Text = mDataBit
    cboStopbit.Text = mStopBit
    cboParity.Text = mParityBit
    txtInterval.Text = CStr(mInterval)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISConfig = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim strFileNm   As String   'INI파일 경로+파일명

    strFileNm = IniPath & "IIS.ini"

    mPort = CInt(cboPort.Text)
    mBaudRate = cboBaud.Text
    mDataBit = cboDatabit.Text
    mStopBit = cboStopbit.Text
    mParityBit = cboParity.Text
    mInterval = CLng(IIf(txtInterval.Text = "", 20, txtInterval.Text))
    
    Call mWriteINI(strFileNm, mEqpKey, "Port", mPort)
    Call mWriteINI(strFileNm, mEqpKey, "BaudRate", mBaudRate)
    Call mWriteINI(strFileNm, mEqpKey, "DataBit", mDataBit)
    Call mWriteINI(strFileNm, mEqpKey, "StopBit", mStopBit)
    Call mWriteINI(strFileNm, mEqpKey, "ParityBit", mParityBit)
    Call mWriteINI(strFileNm, mEqpKey, "Interval", mInterval)
    
    MsgBox "정상적으로 저장되었습니다.", vbInformation, "정보"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Combobox 초기화
'-----------------------------------------------------------------------------'
Private Sub InitCombo()
    '## Port
    With cboPort
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
    End With
    
    '## Baud Rate
    With cboBaud
        .Clear
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "28800"
    End With
    
    '## Data bit
    With cboDatabit
        .Clear
        .AddItem "7"
        .AddItem "8"
    End With
    
    '## Stop bit
    With cboStopbit
        .Clear
        .AddItem "1"
        .AddItem "2"
    End With
    
    '## Parity bit
    With cboParity
        .Clear
        .AddItem "None"
        .AddItem "Even"
        .AddItem "Odd"
    End With
End Sub

