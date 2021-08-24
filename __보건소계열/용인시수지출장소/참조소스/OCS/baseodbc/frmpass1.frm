VERSION 4.00
Begin VB.Form FrmIdPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "사용자 번호 & 비밀번호 입력"
   ClientHeight    =   2040
   ClientLeft      =   2160
   ClientTop       =   3360
   ClientWidth     =   7920
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "굴림체"
      charset         =   1
      weight          =   400
      size            =   12
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2445
   Left            =   2100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Top             =   3015
   Width           =   8040
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      Begin Threed.SSCommand cmdOk 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1335
         _version        =   65536
         _extentx        =   2355
         _extenty        =   661
         _stockprops     =   78
         caption         =   "확 인 [&O]"
         forecolor       =   12582912
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _version        =   65536
         _extentx        =   2355
         _extenty        =   661
         _stockprops     =   78
         caption         =   "취 소 [&C]"
         forecolor       =   128
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   4080
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtIdnumber 
      Height          =   375
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "☎（02）1366－5954"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   360
      Picture         =   "FRMPASS1.frx":0000
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "비 밀 번 호"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "사용자 번호"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
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
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim strSqlDef       As String
Dim strConnect      As String
Dim strGrade        As String
Dim nPassCount      As Integer


Private Sub PassWordCheck_Grade()
    
    strSqlDef = "FOR 1 SELECT Name FROM TWBAS_PASS "
    strSqlDef = strSqlDef & " WHERE ProgramID = '" & GstrPassProgramID & "' "
    strSqlDef = strSqlDef & " AND IDnumber = " & Val(TxtIdnumber.Text)
    Result = dosql(strSqlDef)
    
    If Result = -1 Or rowindicator = 0 Then
        strGrade = "NO"
    Else
        strGrade = "OK"
    End If

End Sub



Private Sub PassWordCheck()

    If Not IsNumeric(TxtIdnumber.Text) Then TxtIdnumber.Text = "0"
    
    strSqlDef = "FOR 1 SELECT Name, PassWord, Grade, Part, ProgramID FROM TW_MIS_PMPA.TWBAS_PASS WHERE "
    strSqlDef = strSqlDef & " ProgramID = ' ' AND IDnumber = " & TxtIdnumber.Text
    
    Result = dosql(strSqlDef)
    
    If Result <> -1 And rowindicator <> 0 Then
        GstrPassWord = GlueGetString("PassWord", 0)
        GstrPassName = GlueGetString("Name", 0)
        GstrPassGrade = GlueGetString("Grade", 0)
        GstrPassPart = GlueGetString("Part", 0)
        GstrPassProgramID = GlueGetString("ProgramID", 0)
        GstrPassID = Trim(TxtIdnumber.Text)
        'GstrPassProgramID = " "
        If txtPassword.Text <> GstrPassWord Then
            Result = -1
        End If
    End If
               
    GstrPassID = Format(TxtIdnumber.Text, "000000")
    
    If Result = -1 Or rowindicator = 0 Then
        strConnect = "NO"
    Else
        strConnect = "OK"
    End If

End Sub


Private Sub cmdCancel_Click()
    
    Call DbDisConnect
    End

End Sub

Private Sub cmdOk_Click()

    nPassCount = nPassCount + 1
    
    Call PassWordCheck
    
    If strConnect = "NO" Then
        If nPassCount > 3 Then
            MsgBox "ID 와 Password를 확인후에 다시 시작하십시요", 48, "경고"
            Call DbDisConnect
            End
        Else
            MsgBox "Password가 틀림니다 !", 48, "주의"
            TxtIdnumber.SetFocus
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
End Sub


Private Sub TxtIdnumber_GotFocus()
    TxtIdnumber.SelStart = 0
    TxtIdnumber.SelLength = Len(TxtIdnumber.Text)

End Sub


Private Sub TxtIdnumber_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub TxtPassWord_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub


Private Sub TxtPassWord_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub TxtPassWord_LostFocus()
    txtPassword.Text = UCase(txtPassword.Text)

End Sub


