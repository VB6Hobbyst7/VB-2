VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmIdPassOCS 
   BorderStyle     =   3  '고정 대화 상자
   Caption         =   "사용자 번호 & 비밀번호 입력"
   ClientHeight    =   990
   ClientLeft      =   2055
   ClientTop       =   2670
   ClientWidth     =   7125
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   990
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   1920
      ScaleHeight     =   915
      ScaleWidth      =   5115
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtIdnumber 
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Top             =   45
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   1680
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "취 소 [&C]"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   60
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "확 인 [&O]"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "사용자 번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "비 밀 번 호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1455
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '단일 고정
      Height          =   990
      Left            =   15
      Picture         =   "Frmpass3.frx":0000
      Top             =   0
      Width           =   1860
   End
End
Attribute VB_Name = "FrmIdPassOCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSqlDef       As String
Dim strConnect      As String
Dim strGrade        As String
Dim nPassCount      As Integer

Dim rs              As rdoResultset



Private Sub PassWordCheck_Grade()
    
    strSqlDef = " SELECT Name FROM TWBAS_PASS "
    strSqlDef = strSqlDef & " WHERE ProgramID = '" & GstrPassProgramID & "' "
    strSqlDef = strSqlDef & " AND IDnumber = " & Val(txtIdnumber.Text)
    
    Result = RdoOpenSet(rs, strSqlDef)
    
    If Result = -1 Or Rowindicator = 0 Then
        strGrade = "NO"
    Else
        strGrade = "OK"
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub



Private Sub PassWordCheck()

    If Not IsNumeric(txtIdnumber.Text) Then txtIdnumber.Text = "0"
    
    strSqlDef = " SELECT Name, PassWord, Grade, Part, ProgramID FROM TW_MIS_PMPA.TWBAS_PASS WHERE "
    strSqlDef = strSqlDef & "  ProgramID = ' ' AND IDnumber = " & txtIdnumber.Text
    
    Result = RdoOpenSet(rs, strSqlDef)
    
    If Result <> -1 And Rowindicator <> 0 Then
        GstrPassWord = rs.rdoColumns("PassWord") & ""
        GstrPassName = rs.rdoColumns("Name") & ""
        GstrPassGrade = rs.rdoColumns("Grade") & ""
        GstrPassPart = rs.rdoColumns("Part") & ""
        GstrPassProgramID = rs.rdoColumns("ProgramID") & ""
        'GstrPassProgramID = " "
        If txtPassword.Text <> GstrPassWord Then
            Result = -1
        End If
    End If
               
    rs.Close
    Set rs = Nothing
               
    GstrIdnumber = Format(txtIdnumber.Text, "000000")
    
    If Result = -1 Or Rowindicator = 0 Then
        strConnect = "NO"
    Else
        strConnect = "OK"
    End If

End Sub


Private Sub CmdCancel_Click()
    
    GstrIdnumber = "0"
    
    Set FrmIdPassOCS = Nothing
    
    Unload Me

End Sub

Private Sub CmdOK_Click()

    nPassCount = nPassCount + 1
    
    Call PassWordCheck
    
    If strConnect = "NO" Then
        'If nPassCount > 3 Then
        '    MsgBox "ID 와 Password를 확인후에 다시 시작하십시요", 48, "경고"
        '    Call DbDisConnect
        '    End
        'Else
            MsgBox "Password가 틀림니다 !", 48, "주의"
            txtIdnumber.SetFocus
        'End If
    Else
        Set FrmIdPassOCS = Nothing
        Unload Me
    End If

End Sub

Private Sub TxtIdnumber_GotFocus()
    
    txtIdnumber.SelStart = 0
    txtIdnumber.SelLength = Len(txtIdnumber.Text)

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


