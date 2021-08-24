VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FrmIdPass 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "사용자 번호 & 비밀번호 입력"
   ClientHeight    =   2370
   ClientLeft      =   1815
   ClientTop       =   1665
   ClientWidth     =   7905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '없음
      Height          =   1335
      Left            =   6075
      ScaleHeight     =   1335
      ScaleWidth      =   1575
      TabIndex        =   7
      Top             =   585
      Width           =   1575
      Begin Threed.SSCommand cmdOk 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "확 인 [&O]"
         ForeColor       =   16711680
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "취 소 [&C]"
         ForeColor       =   128
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   4080
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtIdnumber 
      Height          =   375
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '단일 고정
      Height          =   1725
      Left            =   360
      Picture         =   "Frmpass2.frx":0000
      Stretch         =   -1  'True
      Top             =   495
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "비 밀 번 호"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1650
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "사용자 번호"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   930
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "병원 정보 시스템에 접속하기 위해서 ID와 비밀번호를 입력 하십시요"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "FrmIdPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strGrade                As String
Dim nPassCount              As Integer
Dim sdate                   As Date
Dim stime                   As Date

Private Sub PassWordCheck_Grade()
    
    strSql = "          SELECT  Name                                      "
    strSql = strSql & "   FROM  TW_MIS_PMPA.TWBAS_PASS                                "
    strSql = strSql & "  WHERE  ProgramID  = '" & GstrPassProgramID & "'  "
    strSql = strSql & "    AND  IDnumber   = " & Val(txtIdnumber.Text)
    Result = AdoOpenSet(rs, strSql)
    
    If Result = -1 Or rowindicator = 0 Then
        strGrade = "NO"
    Else
        strGrade = "OK"
    End If


End Sub

Private Sub CmdCancel_Click()
    
    Call DbAdoDisConnect
    End

End Sub

Private Sub cmdOk_Click()
Dim strConnect      As String

     nPassCount = nPassCount + 1
    
    GoSub PassWordCheck
    
    If strConnect = "NO" Then
        If nPassCount > 3 Then
            MsgBox "ID 와 Password를 확인후에 다시 시작하십시요", 48, "경고"
            Call DbAdoDisConnect
            End
        Else
            MsgBox "Password가 틀림니다 !", 48, "주의"
            txtIdnumber.SetFocus
        End If
    Else
        Unload Me
    End If
    Exit Sub

'/------------------------------------------------------
PassWordCheck:

    If Not IsNumeric(txtIdnumber.Text) Then txtIdnumber.Text = "0"
    
    strSql = "         SELECT  Name, PassWord, Grade, Part, ProgramID,    "
    strSql = strSql & "        TO_CHAR(sysdate,'YYYY-MM-DD') SDate,     "
    strSql = strSql & "        TO_CHAR(sysdate,'HH24:MI:SS') STime      "
    strSql = strSql & "  FROM  TW_MIS_PMPA.TWBAS_PASS                                 "
    strSql = strSql & " Where  ProgramID = ' '                            "
    strSql = strSql & "   AND  IDnumber  = " & txtIdnumber.Text
    Result = AdoOpenSet(rs, strSql)
    
    If Result <> -1 And rowindicator <> 0 Then
        sdate = AdoGetString(rs, "SDATE", 0)
        stime = AdoGetString(rs, "STIME", 0)
        GstrSysDate = AdoGetString(rs, "SDATE", 0)

        GstrPassWord = AdoGetString(rs, "PassWord", 0)
        GstrPassName = AdoGetString(rs, "Name", 0)
        GstrPassGrade = AdoGetString(rs, "Grade", 0)
        GstrPassPart = AdoGetString(rs, "Part", 0)
        GstrPassProgramID = AdoGetString(rs, "ProgramID", 0)
        GstrPassId = txtIdnumber.Text
        If txtPassword.Text <> GstrPassWord Then
            Result = -1
        End If
    End If
                
    If Result = -1 Or rowindicator = 0 Then
        strConnect = "NO"
    Else
        strConnect = "OK"
    End If
    Return
    
End Sub

Private Sub Form_Load()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub TxtIdnumber_GotFocus()
    txtIdnumber.SelStart = 0
    txtIdnumber.SelLength = Len(txtIdnumber.Text)
End Sub

Private Sub TxtIdnumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtPassWord_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub TxtPassWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub TxtPassWord_LostFocus()
    txtPassword.Text = UCase(txtPassword.Text)
End Sub
