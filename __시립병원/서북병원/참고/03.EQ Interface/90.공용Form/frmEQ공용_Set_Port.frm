VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEQ공용_Set_Port 
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   3795
   ClientLeft      =   4050
   ClientTop       =   1665
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_Set_Port.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   3795
   Begin VB.ComboBox cboSERIALPARITY 
      Height          =   300
      ItemData        =   "frmEQ공용_Set_Port.frx":263A
      Left            =   1320
      List            =   "frmEQ공용_Set_Port.frx":263C
      Style           =   2  '드롭다운 목록
      TabIndex        =   5
      Top             =   2700
      Width           =   2415
   End
   Begin VB.ComboBox cboSERIALSTOPBIT 
      Height          =   300
      Left            =   1320
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.ComboBox cboSERIALSTARTBIT 
      Height          =   300
      Left            =   1320
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   1890
      Width           =   2415
   End
   Begin VB.ComboBox cboSERIALDATABIT 
      Height          =   300
      ItemData        =   "frmEQ공용_Set_Port.frx":263E
      Left            =   1320
      List            =   "frmEQ공용_Set_Port.frx":2640
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   1500
      Width           =   2415
   End
   Begin VB.ComboBox cboSERIALBAUD 
      Height          =   300
      ItemData        =   "frmEQ공용_Set_Port.frx":2642
      Left            =   1320
      List            =   "frmEQ공용_Set_Port.frx":2644
      TabIndex        =   1
      Text            =   "cboSERIALBAUD"
      Top             =   1110
      Width           =   2415
   End
   Begin VB.ComboBox cboSERIALPORT 
      Height          =   300
      ItemData        =   "frmEQ공용_Set_Port.frx":2646
      Left            =   1320
      List            =   "frmEQ공용_Set_Port.frx":2648
      TabIndex        =   0
      Text            =   "cboSERIALPORT"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CheckBox chkSERIALRTS 
      Caption         =   "RTS Enabled"
      Height          =   225
      Left            =   1320
      TabIndex        =   6
      Top             =   3165
      Width           =   1665
   End
   Begin VB.CheckBox chkSERIALDTR 
      Caption         =   "DTR Enabled"
      Height          =   225
      Left            =   1320
      TabIndex        =   7
      Top             =   3510
      Width           =   1665
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2820
      TabIndex        =   9
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1860
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   11
      Top             =   600
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "패리티"
      Height          =   195
      Index           =   14
      Left            =   60
      TabIndex        =   17
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "정지 비트"
      Height          =   195
      Index           =   13
      Left            =   60
      TabIndex        =   16
      Top             =   2340
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "시작 비트"
      Height          =   195
      Index           =   12
      Left            =   60
      TabIndex        =   15
      Top             =   1950
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "데이터 비트"
      Height          =   195
      Index           =   11
      Left            =   60
      TabIndex        =   14
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "전송속도"
      Height          =   195
      Index           =   10
      Left            =   75
      TabIndex        =   13
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "COM PORT"
      Height          =   195
      Index           =   8
      Left            =   75
      TabIndex        =   12
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "MSComm Port Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   9
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   1
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   1755
   End
End
Attribute VB_Name = "frmEQ공용_Set_Port"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SUB_MM_CANCEL() As Boolean
    Call SUB_MM_KEY_CLEAR
End Function

Private Sub SUB_MM_INITIAL()
    Me.Height = 4275
    Me.Width = 3915
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    
    '/COM PORT ADDITEM
    cboSERIALPORT.AddItem "1"
    cboSERIALPORT.AddItem "2"
    cboSERIALPORT.AddItem "3"
    cboSERIALPORT.AddItem "4"
    cboSERIALPORT.AddItem "5"
    cboSERIALPORT.AddItem "6"
    cboSERIALPORT.AddItem "7"
    cboSERIALPORT.AddItem "8"
    cboSERIALPORT.AddItem "9"
    cboSERIALPORT.AddItem "10"
    cboSERIALPORT.AddItem "11"
    cboSERIALPORT.AddItem "12"
    cboSERIALPORT.AddItem "13"
    
    '/전송속도 ADDITEM
    cboSERIALBAUD.AddItem "100"
    cboSERIALBAUD.AddItem "150"
    cboSERIALBAUD.AddItem "300"
    cboSERIALBAUD.AddItem "600"
    cboSERIALBAUD.AddItem "1200"
    cboSERIALBAUD.AddItem "2400"
    cboSERIALBAUD.AddItem "4800"
    cboSERIALBAUD.AddItem "9600"
    cboSERIALBAUD.AddItem "14400"
    cboSERIALBAUD.AddItem "19200"
    cboSERIALBAUD.AddItem "28800"
    cboSERIALBAUD.AddItem "38400"
    cboSERIALBAUD.AddItem "56000"
    cboSERIALBAUD.AddItem "57600"
    cboSERIALBAUD.AddItem "128000"
    cboSERIALBAUD.AddItem "256000"
    
    '/데이터 비트 ADDITEM
    cboSERIALDATABIT.AddItem "7"
    cboSERIALDATABIT.AddItem "8"
    
    '/시작 비트 ADDITEM
    cboSERIALSTARTBIT.AddItem "1"
    cboSERIALSTARTBIT.AddItem "2"
    
    '/정지 비트 ADDITEM
    cboSERIALSTOPBIT.AddItem "1"
    cboSERIALSTOPBIT.AddItem "1.5"
    cboSERIALSTOPBIT.AddItem "2"
    
    '/패리티 ADDITEM
    cboSERIALPARITY.AddItem "N"
    cboSERIALPARITY.AddItem "E"
    cboSERIALPARITY.AddItem "O"
    
    Call SUB_MM_CANCEL
End Sub

Private Sub SUB_MM_KEY_CLEAR()
    cboSERIALPORT.ListIndex = -1
    cboSERIALBAUD.ListIndex = -1
    cboSERIALDATABIT.ListIndex = -1
    cboSERIALSTARTBIT.ListIndex = -1
    cboSERIALSTOPBIT.ListIndex = -1
    cboSERIALPARITY.ListIndex = -1
    chkSERIALRTS.Value = 0
    chkSERIALDTR.Value = 0
End Sub

Public Function FUNC_MM_SAVE() As Boolean
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_CONF "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        ADC_LOC.BeginTrans

        gstrQuy = "UPDATE EQ_CONF SET "
        gstrQuy = gstrQuy & vbCrLf & "       SERIALPORT     = '" & cboSERIALPORT & "', "            '/RS232 SERIAL PORT
        gstrQuy = gstrQuy & vbCrLf & "       SERIALBAUD     = '" & cboSERIALBAUD & "', "            '/RS232 통신속도
        gstrQuy = gstrQuy & vbCrLf & "       SERIALDATABIT  = '" & cboSERIALDATABIT & "', "         '/RS232 DATABIT(7,8)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALSTARTBIT = '" & cboSERIALSTARTBIT & "', "        '/RS232 STARTBIT(1,2)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALSTOPBIT  = '" & cboSERIALSTOPBIT & "', "         '/RS232 STOPBIT(1,2)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALPARITY   = '" & cboSERIALPARITY & "', "          '/RS232 PARITY(E,N,O)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALRTS      = '" & Val(chkSERIALRTS.Value) & "', "  '/RS232 RTS(0,1)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALDTR      = '" & Val(chkSERIALDTR.Value) & "' "   '/RS232 DTR(0,1)
        If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
        
        ADC_LOC.CommitTrans
    End If
    
    Call CloseDB_LOC
    
    FUNC_MM_SAVE = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_MM_VIEW() As Boolean
    FUNC_MM_VIEW = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_CONF "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        cboSERIALPORT = Trim(ADR_LOC!SERIALPORT & "")
        cboSERIALBAUD = Trim(ADR_LOC!SERIALBAUD & "")
        Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALDATABIT & ""), cboSERIALDATABIT)
        Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALSTARTBIT & ""), cboSERIALSTARTBIT)
        Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALSTOPBIT & ""), cboSERIALSTOPBIT)
        Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALPARITY & ""), cboSERIALPARITY)
        
        chkSERIALRTS.Value = Val(ADR_LOC!SERIALRTS & "")
        chkSERIALDTR.Value = Val(ADR_LOC!SERIALDTR & "")
        
        ADR_LOC.Close: Set ADR_LOC = Nothing
    End If
    Call CloseDB_LOC
    
    FUNC_MM_VIEW = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Private Sub cboSERIALBAUD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSERIALDATABIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSERIALPARITY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSERIALPORT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSERIALSTARTBIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSERIALSTOPBIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chkSERIALDTR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chkSERIALRTS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If MsgBox("저장하겠습니까?", vbQuestion + vbOKCancel, "저장질의") = vbCancel Then Exit Sub
        
    If FUNC_MM_SAVE = True Then
        MsgBox "저장되었습니다!", vbInformation, "확인"
        Call SUB_MM_CANCEL
        Call FUNC_MM_VIEW
        
        MsgBox "환경설정이 변경되었습니다." & vbCrLf & _
               "프로그램을 재 실행하십시오!", vbInformation, "프로그램 종료"
        End
    Else
        MsgBox "저장오류!", vbCritical, "확인"
    End If
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdQuit_Click
    End Select
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
    Call FUNC_MM_VIEW
End Sub
