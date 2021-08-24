VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmPassword 
   Caption         =   "작업자 사번 및 비밀번호 확인"
   ClientHeight    =   2190
   ClientLeft      =   2880
   ClientTop       =   3435
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5145
   Begin Threed.SSPanel SSPanel3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1740
      Width           =   4920
      _Version        =   65536
      _ExtentX        =   8678
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "최선을 다 하겠습니다 !!!"
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1590
      Left            =   1740
      TabIndex        =   5
      Top             =   120
      Width           =   3300
      _Version        =   65536
      _ExtentX        =   5821
      _ExtentY        =   2805
      _StockProps     =   15
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TxtPassword 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   1485
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   585
         Width           =   1680
      End
      Begin VB.TextBox TxtSabun 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1485
         TabIndex        =   0
         Top             =   90
         Width           =   1680
      End
      Begin Threed.SSCommand CmdCancel 
         Height          =   420
         Left            =   1755
         TabIndex        =   3
         Top             =   1080
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "취 소[&C]"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand CmdOk 
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "확 인[&O]"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelIdnumber 
         Appearance      =   0  '평면
         BackColor       =   &H00FFC0C0&
         Caption         =   "ID_Number"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label LabelPassward 
         Appearance      =   0  '평면
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pass_Word"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   630
         Width           =   1230
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1590
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1590
      _Version        =   65536
      _ExtentX        =   2805
      _ExtentY        =   2805
      _StockProps     =   15
      Caption         =   "&H00FFFFFF&"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelInner      =   1
      Begin VB.Image Image1 
         Height          =   1320
         Left            =   135
         Picture         =   "BasPass1.frx":0000
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1320
      End
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FnPassCount     As Integer '비밀번호 오류횟수
Dim FnSabunCnt      As Integer '사번 오류 횟수

Dim FstrGrade       As String
Dim FstrPassWord    As String
Dim FstrPassName    As String
Dim FstrPassGrade   As String
Dim FstrPassPart    As String * 1
Dim FstrPassCharge  As String


Private Sub cmdCancel_click()
    
    Call DbAdoDisConnect
    
    End
    
End Sub

Private Sub cmdOk_Click()
    Dim strPassWord     As String
    
    FnPassCount = FnPassCount + 1
    
    If FnPassCount > 2 Then
        GstrMsgPrompt = "ID와 비밀번호를 확인후" & vbCrLf
        GstrMsgPrompt = GstrMsgPromot & "다시 작업을 하십시오."
        MsgBox GstrMsgPrompt, vbCritical + vbExclamation, "작업중지"
        
        Call DbAdoDisConnect
        
        End
    End If
    
    '사번을 Check
    If TxtSabun.Text = "" Then
        MsgBox "사번이 공란입니다.", , "확인"
        TxtSabun.SetFocus
        Exit Sub
    End If
    
    '입력한 미밀번호를 Check
    strPassWord = UCase(Trim(txtPassword.Text))
    If strPassWord <> FstrPassWord Then
        MsgBox "정확한 비밀번호를 입력하세요", , "확인"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    FstrPassGrade = ""
    Call INSA_READ '퇴사자 여부를 Check
    Call PassWordCheck_Grade  '프로그램 사용 가능여부 Check
    
    '작업자 내역을 Global 변수에 저장
    GnJobSabun = Val(TxtSabun.Text)
    GstrJobName = FstrPassName
    GstrJobPart = FstrPassPart
    GstrPassPart = FstrPassPart
    GstrJobGrade = FstrPassGrade
    
    Unload Me

End Sub


Private Sub Form_Load()
    
    TxtSabun.Text = ""
    txtPassword.Text = ""
    FnPassCount = 0:       FnSabunCnt = 0
    
    GnJobSabun = 0:        GstrJobName = ""
    GstrJobPart = "":      GstrJobGrade = ""

    Call FORM_CENTER(Me)
    
End Sub

Private Sub INSA_READ()
    Dim strTDate            As String
    
    '전산실ID는 인사마스타 CHECK 않함
'B    If TxtSabun.Text = "4349" Then Exit Sub
    
'B    SQL = "SELECT TO_CHAR(TOIDAY,'YYYY-MM-DD') TDate "
'B    SQL = SQL & " FROM KOSMOS_ADM.INSA_MST "
'B    SQL = SQL & "WHERE SABUN = '" & Format(TxtSabun.Text, "00000") & "' "
'B    Result = AdoOpenSet(AdoRes, SQL)
    
    strsql = "          SELECT  Name                                      "
    strsql = strsql & "   FROM  TWBAS_PASS                                "
    strsql = strsql & "  WHERE  ProgramID  = '" & GstrPassProgramID & "'  "
    strsql = strsql & "    AND  IDnumber   = " & Val(txtIdnumber.Text)
    Result = adoSQL(strsql)
    
    If rowindicator > 0 Then
        strTDate = Trim(AdoGetString(AdoRes, "TDate", 0))
        If strTDate < GstrSysDate And strTDate <> "" Then
            GstrMsgPromot = "퇴사자는 작업이 불가능함" & vbCrLf
            GstrMsgPrompt = GstrMsgPrompt & "퇴사일자: " & AdoGetString(AdoRes, "TDate") & "일" & vbCrLf
            MsgBox GstrMsgPrompt, vbCritical + vbExclamation, "종 료"
            
            Call AdoCloseSet(AdoRes)
            
            Call DbAdoDisConnect
            
            End
        End If
    End If
    
    Call AdoCloseSet(AdoRes)
    
End Sub

Private Sub TxtPassWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtSabun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub TxtSabun_LostFocus()
    TxtSabun.Text = Trim(TxtSabun.Text)
    
    If TxtSabun.Text = "" Then Exit Sub
    FstrPassName = ""
    
    '사번을 오류로 3번이상이면 강제로 종료함
    FnSabunCnt = FnSabunCnt + 1
    If FnSabunCnt > 3 Then
        GstrMsgPrompt = "정확한 사번을 확인후" & vbCrLf
        GstrMsgPrompt = GstrMsgPromot & "다시 작업을 하십시오."
        MsgBox GstrMsgPrompt, vbCritical + vbExclamation, "작업중지"
        
        Call DbAdoDisConnect
        
        End
    End If
    
    'BAS_PASS에서 해당 사번이 등록되었는지 점검
    SQL = "SELECT Name, PassWard, Grade, Part, ProgramID,Charge "
    SQL = SQL & " FROM  KOSMOS_PMPA.BAS_PASS "
    SQL = SQL & "WHERE  ProgramID = ' ' "
    SQL = SQL & "  AND  IDnumber = " & Val(TxtSabun.Text) & " "
    Result = AdoOpenSet(AdoRes, SQL)
    If rowindicator > 0 Then
        FstrPassWord = Trim(AdoGetString(AdoRes, "PassWard", 0))
        FstrPassName = Trim(AdoGetString(AdoRes, "Name", 0))
        FstrPassGrade = Trim(AdoGetString(AdoRes, "Grade", 0))
        FstrPassPart = AdoGetString(AdoRes, "Part", 0)
        FstrPassCharge = Trim(AdoGetString(AdoRes, "Charge", 0))
    End If
        
    Call AdoCloseSet(AdoRes)
    
    If FstrPassName = "" Then  '이름이 공란이면
        GstrMsgPrompt = "사번이 등록되지 않았습니다." & vbCrLf
        GstrMsgPrompt = GstrMsgPrompt & "사번을 확인후 맞으시면" & vbCrLf
        GstrMsgPrompt = GstrMsgPrompt & "전산실에 연락 주십시오."
        MsgBox GstrMsgPrompt, vbCritical + vbExclamation, "확인"
        TxtSabun.SetFocus
        Exit Sub
    End If
    
End Sub
Private Sub PassWordCheck_Grade()

    Dim strProgPass          As String
    
    If GstrPassProgramID = "" Then Exit Sub   '공용P/G
    If TxtSabun.Text = "4349" Then Exit Sub  '전산실
    
    ' PROG_EXE에서 해당 ProgPass를 찾기
    SQL = "SELECT ProgPass FROM KOSMOS_PMPA.PROG_EXE "
    SQL = SQL & " WHERE ExeCode = '" & Trim(GstrPassProgramID) & "' "
    Result = AdoOpenSet(AdoRes, SQL)
    strProgPass = ""
    If rowindicator > 0 Then strProgPass = Trim(AdoGetString(AdoRes, "ProgPass", 0))
    Call AdoCloseSet(AdoRes)
    
    If strProgPass = "" Then strProgPass = GstrPassProgramID

    ' 해당 ProgPass의 사용권한을 Check
    SQL = "SELECT Name FROM KOSMOS_PMPA.BAS_PASS "
    SQL = SQL & " WHERE ProgramID = '" & strProgPass & "' "
    SQL = SQL & "   AND IDnumber  = " & Val(TxtSabun.Text) & " "
    Result = AdoOpenSet(AdoRes, SQL)
    If rowindicator = 0 Then
        GstrMsgPrompt = "이 프로그램의 사용 권한이 없습니다." & vbCrLf
        GstrMsgPrompt = GstrMsgPrompt & "전산실에 확인을 하십시오." & vbCrLf
        MsgBox GstrMsgPrompt, vbCritical + vbExclamation, "작업중지"
        
        Call DbAdoDisConnect
        
        End
    Else
        Call AdoCloseSet(AdoRes)
    End If
    
End Sub

