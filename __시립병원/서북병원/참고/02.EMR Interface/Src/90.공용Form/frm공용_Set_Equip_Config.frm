VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frm공용_Set_Equip_Config 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Client info Setting"
   ClientHeight    =   8145
   ClientLeft      =   7290
   ClientTop       =   4260
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm공용_Set_Equip_Config.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra상세내역 
      Caption         =   "[상세내역]"
      Height          =   6615
      Left            =   60
      TabIndex        =   25
      Top             =   1500
      Width           =   7035
      Begin VB.TextBox txtZIPNM 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1260
         TabIndex        =   5
         Text            =   "txtZIPNM"
         Top             =   1080
         Width           =   5715
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1260
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   45
         Top             =   660
         Width           =   1995
         Begin VB.OptionButton optZIPYN 
            Caption         =   "사용"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optZIPYN 
            Caption         =   "미사용"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   4
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.TextBox txtREMARK 
         Height          =   735
         IMEMode         =   10  '한글 
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "frm공용_Set_Equip_Config.frx":000C
         Top             =   5760
         Width           =   6795
      End
      Begin VB.TextBox txtDEPTCODE 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1260
         TabIndex        =   0
         Text            =   "txtDEPTCODE"
         Top             =   240
         Width           =   1995
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   3780
         ScaleHeight     =   3225
         ScaleWidth      =   3165
         TabIndex        =   41
         Top             =   2100
         Width           =   3195
         Begin FPSpread.vaSpread sprEQORD 
            Height          =   2895
            Left            =   60
            TabIndex        =   18
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
            SpreadDesigner  =   "frm공용_Set_Equip_Config.frx":0018
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "처방코드"
            Height          =   180
            Left            =   120
            TabIndex        =   42
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   4980
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   35
         Top             =   1740
         Width           =   1995
         Begin VB.OptionButton optORDYN 
            Caption         =   "Yes"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optORDYN 
            Caption         =   "No"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   17
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   4980
         ScaleHeight     =   300
         ScaleWidth      =   1965
         TabIndex        =   34
         Top             =   240
         Width           =   1995
         Begin VB.OptionButton optRECEIVETYPE 
            Caption         =   "간접"
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   2
            Top             =   0
            Width           =   765
         End
         Begin VB.OptionButton optRECEIVETYPE 
            Caption         =   "직접"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   1
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1260
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   33
         Top             =   1740
         Width           =   1995
         Begin VB.OptionButton optSERIALYN 
            Caption         =   "미사용"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   7
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton optSERIALYN 
            Caption         =   "사용"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.PictureBox picSERIALPORT 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   60
         ScaleHeight     =   3225
         ScaleWidth      =   3165
         TabIndex        =   26
         Top             =   2100
         Width           =   3195
         Begin VB.CheckBox chkSERIALDTR 
            Caption         =   "DTR Enabled"
            Height          =   225
            Left            =   1200
            TabIndex        =   15
            Top             =   2850
            Width           =   1665
         End
         Begin VB.CheckBox chkSERIALRTS 
            Caption         =   "RTS Enabled"
            Height          =   225
            Left            =   1200
            TabIndex        =   14
            Top             =   2505
            Width           =   1665
         End
         Begin VB.ComboBox cboSERIALPORT 
            Height          =   300
            ItemData        =   "frm공용_Set_Equip_Config.frx":034A
            Left            =   1200
            List            =   "frm공용_Set_Equip_Config.frx":034C
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   60
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALBAUD 
            Height          =   300
            ItemData        =   "frm공용_Set_Equip_Config.frx":034E
            Left            =   1200
            List            =   "frm공용_Set_Equip_Config.frx":0350
            TabIndex        =   9
            Text            =   "cboSERIALBAUD"
            Top             =   450
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALDATABIT 
            Height          =   300
            ItemData        =   "frm공용_Set_Equip_Config.frx":0352
            Left            =   1200
            List            =   "frm공용_Set_Equip_Config.frx":0354
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   840
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALSTARTBIT 
            Height          =   300
            Left            =   1200
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   1230
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALSTOPBIT 
            Height          =   300
            Left            =   1200
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   1620
            Width           =   1875
         End
         Begin VB.ComboBox cboSERIALPARITY 
            Height          =   300
            ItemData        =   "frm공용_Set_Equip_Config.frx":0356
            Left            =   1200
            List            =   "frm공용_Set_Equip_Config.frx":0358
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   2100
            Width           =   1155
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "가상프린터"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "비고"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   5520
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "설치장소"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Image획득방식"
         Height          =   180
         Index           =   7
         Left            =   3780
         TabIndex        =   38
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방유무"
         Height          =   180
         Left            =   3780
         TabIndex        =   37
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Serial Port"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   1800
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   40
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
      TabIndex        =   39
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
Attribute VB_Name = "frm공용_Set_Equip_Config"
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
    fra상세내역.Enabled = False

    optORDYN(0).Value = True '/처방유무
    optRECEIVETYPE(0).Value = True '/Image 획득방식
    
    optSERIALYN(1).Value = True '/Serial Port
    picSERIALPORT.Enabled = False
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
    
    Call MM_KEY_CLEAR
    
    If OpenDB(gstrREG_DB_CONSTR) = False Then Exit Function
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & Trim(Left(lbl장비코드, 5)) & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  = '" & Trim(lbl장비SEQ) & "' "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: Exit Function
    
    If Not ADR Is Nothing Then
        fra상세내역.Enabled = True
        
        txtDEPTCODE = Trim(ADR!DEPTCODE & "") '/설치장소
        
        If Trim(ADR!ZIPYN & "") = "Y" Then '/가상프린터 사용여부(Y.사용,N.미사용)
            optZIPYN(0).Value = True
        Else
            optZIPYN(1).Value = True
        End If
        txtZIPNM = Trim(ADR!ZIPNM & "") '/가상프린터 Device Name
        
        If Trim(ADR!ORDYN & "") = "Y" Then '/처방유무
            optORDYN(0).Value = True
        Else
            optORDYN(1).Value = True
        End If
        
        If Trim(ADR!RECEIVETYPE & "") = "1" Then '/Image 획득방식
            optRECEIVETYPE(0).Value = True
        Else
            optRECEIVETYPE(1).Value = True
        End If
        
        If Trim(ADR!SERIALYN & "") = "Y" Then '/Serial Port
            optSERIALYN(0).Value = True
        
            picSERIALPORT.Enabled = True
            
            Call SET_CBO_DT_ALL(Trim(ADR!SERIALPORT & ""), cboSERIALPORT)
            cboSERIALBAUD = Trim(ADR!SERIALBAUD & "")
            '''Call SET_CBO_DT_ALL(Trim(ADR!SERIALBAUD & ""), cboSERIALBAUD)
            Call SET_CBO_DT_ALL(Trim(ADR!SERIALDATABIT & ""), cboSERIALDATABIT)
            Call SET_CBO_DT_ALL(Trim(ADR!SERIALSTARTBIT & ""), cboSERIALSTARTBIT)
            Call SET_CBO_DT_ALL(Trim(ADR!SERIALSTOPBIT & ""), cboSERIALSTOPBIT)
            Call SET_CBO_DT_ALL(Trim(ADR!SERIALPARITY & ""), cboSERIALPARITY)
            
            chkSERIALRTS.Value = Val(ADR!SERIALRTS & "")
            chkSERIALDTR.Value = Val(ADR!SERIALDTR & "")
        Else
            optSERIALYN(1).Value = True
        End If
        
        txtREMARK = Trim(ADR!REMARK & "") '/비고
        
        ADR.Close: Set ADR = Nothing
    Else
        fra상세내역.Enabled = True
    End If

    '/처방코드
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & Trim(Left(lbl장비코드, 5)) & "' "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY ORDCD "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: Exit Function
    
    If Not ADR Is Nothing Then
        With sprEQORD
            If .MaxRows > 0 Then .MaxRows = 0
        
            Do Until ADR.EOF
                .MaxRows = .MaxRows + 1: .Row = .MaxRows
                
                .Col = 1: .Text = Trim(ADR!ORDCD & "")
                
                ADR.MoveNext
            Loop
            .MaxRows = .MaxRows + 1
        End With
        
        ADR.Close: Set ADR = Nothing
    End If

    Call CloseDB
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
    If OpenDB(gstrREG_DB_CONSTR) = False Then End
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & Trim(Left(lbl장비코드, 5)) & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  = '" & Trim(lbl장비SEQ) & "' "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
    
    If Not ADR Is Nothing Then
        ADR.Close: Set ADR = Nothing
        
        ADC.BeginTrans

        gstrQuy = "UPDATE MM_EMR_CONF SET "
        gstrQuy = gstrQuy & vbCrLf & "       SERIALYN       = '" & IIf(optSERIALYN(0).Value = True, "Y", "N") & "', " '/RS232 SERIAL 사용여부(Y.사용, N.미사용)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALPORT     = '" & cboSERIALPORT & "', " '/RS232 SERIAL PORT
        gstrQuy = gstrQuy & vbCrLf & "       SERIALBAUD     = '" & cboSERIALBAUD & "', " '/RS232 통신속도
        gstrQuy = gstrQuy & vbCrLf & "       SERIALDATABIT  = '" & cboSERIALDATABIT & "', " '/RS232 DATABIT(7,8)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALSTARTBIT = '" & cboSERIALSTARTBIT & "', " '/RS232 STARTBIT(1,2)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALSTOPBIT  = '" & cboSERIALSTOPBIT & "', " '/RS232 STOPBIT(1,2)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALPARITY   = '" & cboSERIALPARITY & "', " '/RS232 PARITY(E,N,O)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALRTS      = '" & Val(chkSERIALRTS.Value) & "', "  '/RS232 RTS(0,1)
        gstrQuy = gstrQuy & vbCrLf & "       SERIALDTR      = '" & Val(chkSERIALDTR.Value) & "', " '/RS232 DTR(0,1)
        gstrQuy = gstrQuy & vbCrLf & "       RECEIVETYPE    = '" & IIf(optRECEIVETYPE(0).Value = True, "1", "2") & "', " '/Image획득방식(1.직접, 2.간접)
'''        gstrQuy = gstrQuy & vbCrLf & "       EQUIPPORT      = '', " '/장비PC접속포트(RECEIVETYPE 간접일경우)
        gstrQuy = gstrQuy & vbCrLf & "       ORDYN          = '" & IIf(optORDYN(0).Value = True, "Y", "N") & "', " '/처방유무(Y.처방, N.미처방)
        gstrQuy = gstrQuy & vbCrLf & "       ZIPYN          = '" & IIf(optZIPYN(0).Value = True, "Y", "N") & "', " '/ZanImagePrinter사용여부(Y.사용,N.미사용)
        gstrQuy = gstrQuy & vbCrLf & "       ZIPNM          = '" & Trim(txtZIPNM) & "', " '/ZanImagePrinter Device Name
        gstrQuy = gstrQuy & vbCrLf & "       DEPTCODE       = '" & Trim(txtDEPTCODE) & "', " '/진료과코드
        gstrQuy = gstrQuy & vbCrLf & "       REMARK         = '" & Trim(TEXT_LSET(Trim(txtREMARK), 200)) & "' " '/비고
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE      = '" & Trim(Left(lbl장비코드, 5)) & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ       = '" & Trim(lbl장비SEQ) & "' "
        If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
        
        '/처방코드
        gstrQuy = "DELETE MM_EMR_EQORD "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE      = '" & Trim(Left(lbl장비코드, 5)) & "' "
        If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    
        For intX = 1 To sprEQORD.MaxRows
            If Trim(GET_CELL(sprEQORD, 1, intX)) <> "" Then
                gstrQuy = "INSERT INTO MM_EMR_EQORD (EQUIPCODE, ORDCD)"
                gstrQuy = gstrQuy & vbCrLf & " VALUES ('" & Trim(Left(lbl장비코드, 5)) & "', '" & Trim(GET_CELL(sprEQORD, 1, intX)) & "') "
                If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            End If
        Next intX
        
        ADC.CommitTrans
    End If
    
    Call CloseDB
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
            
            Call SET_CBO_DT_ALL("1", cboSERIALPORT)
            Call SET_CBO_DT_ALL("9600", cboSERIALBAUD)
            Call SET_CBO_DT_ALL("8", cboSERIALDATABIT)
            Call SET_CBO_DT_ALL("1", cboSERIALSTARTBIT)
            Call SET_CBO_DT_ALL("1", cboSERIALSTOPBIT)
            Call SET_CBO_DT_ALL("N", cboSERIALPARITY)
            
            chkSERIALRTS.Value = 1
            chkSERIALDTR.Value = 1
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

Private Sub optZIPYN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub sprEQORD_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If sprEQORD.MaxRows = Row Then sprEQORD.MaxRows = sprEQORD.MaxRows + 1
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
