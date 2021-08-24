VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmIdPass 
   BorderStyle     =   3  '고정 대화 상자
   Caption         =   "사용자 번호 & 비밀번호 입력"
   ClientHeight    =   1980
   ClientLeft      =   1110
   ClientTop       =   1485
   ClientWidth     =   6480
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
   ScaleHeight     =   1980
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '사용 못함
      Left            =   3180
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1350
      Width           =   1500
   End
   Begin VB.TextBox txtIdnumber 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3180
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   1500
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   405
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "취 소 [&C]"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   405
      Left            =   4920
      TabIndex        =   2
      Top             =   825
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "확 인 [&O]"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '단일 고정
      Height          =   1605
      Left            =   180
      Picture         =   "Frmid2.frx":0000
      Top             =   150
      Width           =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "비 밀 번 호"
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
      Left            =   1935
      TabIndex        =   6
      Top             =   1410
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "사용자 번호"
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
      Left            =   1935
      TabIndex        =   5
      Top             =   930
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "병원 정보 시스템에 접속하기 위해서 ID와 비밀번호를 입력 하십시요....."
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1935
      TabIndex        =   4
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "FrmIdPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSqlDef       As String
Dim strConnect      As String
Dim strGrade        As String
Dim nPassCount      As Integer


Private Function PassWordCheck_Grade(argProgramID As String, argIDno As Long) As String
    
    strSqlDef = "SELECT Name FROM TWBAS_PASS " & _
                " WHERE ProgramID = '" & argProgramID & "' " & _
                "   AND IDnumber  =  " & argIDno
    
    
    If OpenRDO(strSqlDef, 0) Then
        PassWordCheck_Grade = "OK"
        RdoSet(0).Close
    Else
        PassWordCheck_Grade = "NO"
    End If
    
End Function


Private Function PassWordCheck(argIDnum As Long) As String

    strSqlDef = "SELECT Name, PassWord, Class, Grade, Part, ProgramID, DeptCode " & _
                "  FROM TWBAS_PASS " & _
                " WHERE ProgramID = ' ' " & _
                "   AND IDnumber  =  " & argIDnum
    
    If OpenRDO(strSqlDef, 0) Then
        GstrPassWord = RdoSet(0).rdoColumns("Password")
        GstrPassName = RdoSet(0).rdoColumns("Name")
        GstrPassClass = RdoSet(0).rdoColumns("Class")
        GstrPassGrade = RdoSet(0).rdoColumns("Grade")
        GstrPassPart = RdoSet(0).rdoColumns("Part")
        GstrPassDept = RdoSet(0).rdoColumns("DeptCode")
        RdoSet(0).Close
    
        If txtPassword.Text <> GstrPassWord Then
            PassWordCheck = "PS"
        Else
            PassWordCheck = "OK"
        End If
    Else
            PassWordCheck = "NO"
    End If
    
    GstrPassIDnumber = Format$(argIDnum)
End Function


Private Sub CmdCancel_Click()
    
    RdoDB.Close
    End

End Sub


Private Sub CmdOK_Click()
    
    nPassCount = nPassCount + 1
    
    Select Case PassWordCheck(Val(txtIdnumber.Text))
        Case "PS":  MsgBox "비밀번호가 틀림니다 !", vbCritical, "재확인요망"
        Case "NO":  MsgBox "해당 ID 가 없습니다 !", vbCritical, "재확인요망"
        
        Case "OK":
        Select Case PassWordCheck_Grade(GstrPassProgramID, Val(txtIdnumber.Text))
               Case "NO":   MsgBox "이프로그램을 사용하실 권한이 없습니다 !", vbInformation, "알림"
               
               Case "OK":   Unload Me
                            Call Read_Announce_Ment
                            Exit Sub
        End Select
    End Select
    
    If nPassCount > 3 Then
        MsgBox "ID 와 Password를 확인후에 다시 시작하십시요", vbCritical, "경고"
        RdoDB.Close
        End
    End If
    
End Sub


Private Sub Form_Load()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 200
    
End Sub


Private Sub TxtIdnumber_GotFocus()
    
    txtIdnumber.SelStart = 0
    txtIdnumber.SelLength = Len(txtIdnumber.Text)
    
End Sub


Private Sub txtIdnumber_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then SendKeys "{TAB}"
    
End Sub

Private Sub txtPassword_GotFocus()
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub


Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then SendKeys "{TAB}"
    
End Sub

Private Sub txtPassword_LostFocus()
    
    txtPassword.Text = UCase(txtPassword.Text)

End Sub

