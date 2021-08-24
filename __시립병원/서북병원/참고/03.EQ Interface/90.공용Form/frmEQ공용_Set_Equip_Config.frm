VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEQ공용_Set_Equip_Config 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Client info Setting"
   ClientHeight    =   7050
   ClientLeft      =   6915
   ClientTop       =   1905
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_Set_Equip_Config.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "이거 어떻게 할까요?"
      Height          =   5235
      Left            =   3720
      TabIndex        =   27
      Top             =   1560
      Width           =   5655
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   3600
         ScaleHeight     =   300
         ScaleWidth      =   1965
         TabIndex        =   39
         Top             =   780
         Width           =   1995
         Begin VB.OptionButton optRECEIVETYPE 
            Caption         =   "직접"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   41
            Top             =   0
            Width           =   765
         End
         Begin VB.OptionButton optRECEIVETYPE 
            Caption         =   "간접"
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   40
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1260
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   36
         Top             =   1500
         Width           =   1995
         Begin VB.OptionButton optORDYN 
            Caption         =   "No"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   38
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optORDYN 
            Caption         =   "Yes"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   37
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   60
         ScaleHeight     =   3225
         ScaleWidth      =   3165
         TabIndex        =   33
         Top             =   1860
         Width           =   3195
         Begin FPSpread.vaSpread sprEQORD 
            Height          =   2895
            Left            =   60
            TabIndex        =   34
            Top             =   300
            Width           =   3075
            _Version        =   393216
            _ExtentX        =   5424
            _ExtentY        =   5106
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   1
            MaxRows         =   10
            SpreadDesigner  =   "frmEQ공용_Set_Equip_Config.frx":000C
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "처방코드"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.TextBox txtDEPTCODE 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   3600
         TabIndex        =   32
         Text            =   "txtDEPTCODE"
         Top             =   300
         Width           =   1995
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   29
         Top             =   720
         Width           =   1995
         Begin VB.OptionButton optZIPYN 
            Caption         =   "미사용"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   31
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton optZIPYN 
            Caption         =   "사용"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   30
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.TextBox txtZIPNM 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   60
         TabIndex        =   28
         Text            =   "txtZIPNM"
         Top             =   1140
         Width           =   5715
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방유무"
         Height          =   180
         Left            =   60
         TabIndex        =   45
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Image획득방식"
         Height          =   180
         Index           =   7
         Left            =   2400
         TabIndex        =   44
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "설치장소"
         Height          =   180
         Index           =   1
         Left            =   2460
         TabIndex        =   43
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "가상프린터"
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.Frame fra상세내역 
      Caption         =   "[상세내역]"
      Height          =   4335
      Left            =   60
      TabIndex        =   15
      Top             =   1500
      Width           =   3555
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1320
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   23
         Top             =   240
         Width           =   1995
         Begin VB.OptionButton optSERIALYN 
            Caption         =   "미사용"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   1
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton optSERIALYN 
            Caption         =   "사용"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   0
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.PictureBox picSERIALPORT 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3225
         ScaleWidth      =   3165
         TabIndex        =   16
         Top             =   600
         Width           =   3195
         Begin VB.CheckBox chkSERIALDTR 
            Caption         =   "DTR Enabled"
            Height          =   225
            Left            =   1200
            TabIndex        =   9
            Top             =   2850
            Width           =   1665
         End
         Begin VB.CheckBox chkSERIALRTS 
            Caption         =   "RTS Enabled"
            Height          =   225
            Left            =   1200
            TabIndex        =   8
            Top             =   2505
            Width           =   1665
         End
         Begin VB.ComboBox cboSERIALPORT 
            Height          =   300
            ItemData        =   "frmEQ공용_Set_Equip_Config.frx":0314
            Left            =   1200
            List            =   "frmEQ공용_Set_Equip_Config.frx":0316
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALBAUD 
            Height          =   300
            ItemData        =   "frmEQ공용_Set_Equip_Config.frx":0318
            Left            =   1200
            List            =   "frmEQ공용_Set_Equip_Config.frx":031A
            TabIndex        =   3
            Text            =   "cboSERIALBAUD"
            Top             =   450
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALDATABIT 
            Height          =   300
            ItemData        =   "frmEQ공용_Set_Equip_Config.frx":031C
            Left            =   1200
            List            =   "frmEQ공용_Set_Equip_Config.frx":031E
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   840
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALSTARTBIT 
            Height          =   300
            Left            =   1200
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   1230
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALSTOPBIT 
            Height          =   300
            Left            =   1200
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   1620
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALPARITY 
            Height          =   300
            ItemData        =   "frmEQ공용_Set_Equip_Config.frx":0320
            Left            =   1200
            List            =   "frmEQ공용_Set_Equip_Config.frx":0322
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   2040
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "COM PORT"
            Height          =   195
            Index           =   8
            Left            =   -45
            TabIndex        =   22
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전송속도"
            Height          =   195
            Index           =   10
            Left            =   -45
            TabIndex        =   21
            Top             =   510
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "데이터 비트"
            Height          =   195
            Index           =   11
            Left            =   -60
            TabIndex        =   20
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "시작 비트"
            Height          =   195
            Index           =   12
            Left            =   -60
            TabIndex        =   19
            Top             =   1290
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "정지 비트"
            Height          =   195
            Index           =   13
            Left            =   -60
            TabIndex        =   18
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "패리티"
            Height          =   195
            Index           =   14
            Left            =   -60
            TabIndex        =   17
            Top             =   2100
            Width           =   1155
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Serial Port"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   300
         Width           =   990
      End
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
      Height          =   615
      Left            =   6180
      TabIndex        =   11
      Top             =   660
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
      Height          =   615
      Left            =   5220
      TabIndex        =   10
      Top             =   660
      Width           =   915
   End
   Begin VB.Label lbl장비SEQ 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "lbl장비SEQ"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   26
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label lbl장비코드 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비코드"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   25
      Top             =   720
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   7080
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "EMR Interface Client Information"
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
      Height          =   240
      Index           =   9
      Left            =   180
      TabIndex        =   14
      Top             =   60
      Width           =   3300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비SEQ"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비코드"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   720
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
      Width           =   7035
   End
End
Attribute VB_Name = "frmEQ공용_Set_Equip_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function MM_CANCEL() As Boolean
    lbl장비코드 = ""
    lbl장비SEQ = ""

    Call MM_KEY_CLEAR
End Function

Private Sub MM_INITIAL()
    Me.Height = 7470
    Me.Width = 7275
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    
    '/COM PORT ADDITEM
    cboSERIALPORT.AddItem "1"
    cboSERIALPORT.AddItem "2"
    cboSERIALPORT.AddItem "3"
    cboSERIALPORT.AddItem "4"
    cboSERIALPORT.AddItem "5"
    cboSERIALPORT.AddItem "6"
    
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
    
    Call MM_CANCEL
End Sub

Private Sub MM_KEY_CLEAR()

    optORDYN(0).Value = True '/처방유무
    optRECEIVETYPE(0).Value = True '/Image 획득방식
    
    optSERIALYN(0).Value = True '/Serial Port
    picSERIALPORT.Enabled = True
    cboSERIALPORT.ListIndex = -1
    cboSERIALBAUD = ""
    cboSERIALDATABIT.ListIndex = -1
    cboSERIALSTARTBIT.ListIndex = -1
    cboSERIALSTOPBIT.ListIndex = -1
    cboSERIALPARITY.ListIndex = -1
    
    chkSERIALRTS.Value = 0
    chkSERIALDTR.Value = 0
    
    If sprEQORD.MaxRows > 0 Then sprEQORD.MaxRows = 0: sprEQORD.MaxRows = 1
End Sub

Public Function MM_VIEW() As Boolean
    
    If Trim(lbl장비코드) = "" Then Exit Function
    If Trim(lbl장비SEQ) = "" Then Exit Function
    
    
    
    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_CONF "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        
        If Trim(ADR_LOC!SERIALYN & "") = "Y" Then
            optSERIALYN(0).Value = True
            Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALPORT & ""), cboSERIALPORT)
            cboSERIALBAUD = Trim(ADR_LOC!SERIALBAUD & "")
            '''Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALBAUD & ""), cboSERIALBAUD)
            Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALDATABIT & ""), cboSERIALDATABIT)
            Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALSTARTBIT & ""), cboSERIALSTARTBIT)
            Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALSTOPBIT & ""), cboSERIALSTOPBIT)
            Call SET_CBO_DT_ALL(Trim(ADR_LOC!SERIALPARITY & ""), cboSERIALPARITY)
            
            chkSERIALRTS.Value = Val(ADR_LOC!SERIALRTS & "")
            chkSERIALDTR.Value = Val(ADR_LOC!SERIALDTR & "")
            
        Else
            optSERIALYN(1).Value = True
        End If
        ADR_LOC.Close: Set ADR_LOC = Nothing
    Else
    
    End If

    Call CloseDB_LOC
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
    If ConnDB_LOC = False Then End
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_CONF "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        ADC_LOC.BeginTrans

        gstrQuy = "UPDATE EQ_CONF SET "
        gstrQuy = gstrQuy & vbCrLf & "       SERIALPORT     = '" & cboSERIALPORT & "', " '/RS232 SERIAL PORT
        gstrQuy = gstrQuy & vbCrLf & "       SERIALBAUD     = '" & cboSERIALBAUD & "', " '/RS232 통신속도
        gstrQuy = gstrQuy & vbCrLf & "       SERIALDATABIT  = '" & cboSERIALDATABIT & "', " '/RS232 DATABIT(7,8)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALSTARTBIT = '" & cboSERIALSTARTBIT & "', " '/RS232 STARTBIT(1,2)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALSTOPBIT  = '" & cboSERIALSTOPBIT & "', " '/RS232 STOPBIT(1,2)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALPARITY   = '" & cboSERIALPARITY & "', " '/RS232 PARITY(E,N,O)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALRTS      = '" & Val(chkSERIALRTS.Value) & "', "  '/RS232 RTS(0,1)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALDTR      = '" & Val(chkSERIALDTR.Value) & "' " '/RS232 DTR(0,1)
        
        If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: End
        
        ADC_LOC.CommitTrans
        
        MsgBox "저장되었습니다!", vbInformation, "확인"
    End If
    
    Call CloseDB_LOC
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
    Call MM_INITIAL
    
    lbl장비코드 = gtypEQ_INFO.EQUIPCODE & "." & gtypEQ_INFO.EQUIPNM
    lbl장비SEQ = gtypEQ_INFO.EQUIPSEQ
    
    Call MM_VIEW
End Sub

Private Sub Label13_Click()

End Sub

Private Sub optORDYN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub optRECEIVETYPE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub optSERIALYN_Click(Index As Integer)
    Select Case Index
        Case 0 '/사용
            picSERIALPORT.Enabled = True
            Call MM_VIEW
        Case 1 '/미사용
            picSERIALPORT.Enabled = False
            cboSERIALPORT.ListIndex = -1
            cboSERIALBAUD.ListIndex = -1
            cboSERIALDATABIT.ListIndex = -1
            cboSERIALSTARTBIT.ListIndex = -1
            cboSERIALSTOPBIT.ListIndex = -1
            cboSERIALPARITY.ListIndex = -1

            chkSERIALRTS.Value = 0
            chkSERIALDTR.Value = 0
    End Select
End Sub

Private Sub optSERIALYN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQUIPSEQ_Change()
    Call MM_KEY_CLEAR
End Sub

Private Sub txtEQUIPSEQ_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQUIPSEQ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MM_VIEW
End Sub

Private Sub optZIPYN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDEPTCODE_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtDEPTCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtZIPNM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtZIPNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
