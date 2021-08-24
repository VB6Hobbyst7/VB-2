VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmPass 
   BorderStyle     =   3  '고정 대화 상자
   Caption         =   "사용자 확인"
   ClientHeight    =   1965
   ClientLeft      =   2445
   ClientTop       =   3075
   ClientWidth     =   6720
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
   ScaleHeight     =   1965
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPasswd 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '사용 못함
      Left            =   3540
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   1500
   End
   Begin VB.TextBox TxtIDnum 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3540
      MaxLength       =   6
      TabIndex        =   0
      Top             =   660
      Width           =   1500
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   420
      Left            =   5220
      TabIndex        =   3
      Top             =   1110
      Width           =   1300
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "취소 [&C]"
      ForeColor       =   128
      Font            =   "FrmPass1.frx":0000
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   420
      Left            =   5220
      TabIndex        =   2
      Top             =   620
      Width           =   1300
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "확인 [&O]"
      ForeColor       =   12582912
      Font            =   "FrmPass1.frx":0025
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "☎ (02) 3664-5954"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '단일 고정
      Height          =   990
      Left            =   240
      Picture         =   "FrmPass1.frx":004A
      Top             =   540
      Width           =   1860
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
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
      Left            =   2265
      TabIndex        =   6
      Top             =   1200
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
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
      Left            =   2265
      TabIndex        =   5
      Top             =   720
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "병원 정보 시스템에 접속하기 위해서 ID와 비밀번호를 입력 하십시요"
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
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   6330
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSqlDef       As String
Dim strConnect      As String
Dim strGrade        As String
Dim nPassCount      As Integer

Dim LsUserId         As String * 6
Dim LsUserName       As String * 20
Dim LsPassWord       As String * 4
Dim LsDeptNo         As String * 4
Dim LsRank           As Integer


Private Sub PassWordCheck()

    If Not IsNumeric(TxtIDnum.Text) Then TxtIDnum.Text = "0"
    
    strSqlDef = "FOR 1 SELECT * FROM TWEXAM_PASSWORD WHERE UserId = '" & TxtIDnum.Text & "'"
    Result = dosql(strSqlDef)
    
    If Result <> -1 And rowindicator <> 0 Then
        LsUserId = GlueGetString("UserId", 0)
        LsUserName = GlueGetString("UserName", 0)
        LsPassWord = UCase$(Trim$(GlueGetString("PassWord", 0)))
        LsDeptNo = GlueGetString("DeptCode", 0)
        LsRank = GlueGetString("Rank", 0)
        
        If UCase$(Trim$(TxtPasswd.Text)) <> UCase$(Trim$(LsPassWord)) Then
            Result = -1
        End If
    End If
               
    LsUserId = Format(TxtIDnum.Text, "000000")
    
    If Result = -1 Or rowindicator = 0 Then
        GsUserid = ""
        GsUserName = ""
        UserRank = 0
        strConnect = "NO"
    Else
        strConnect = "OK"
        GsUserid = LsUserId
        GsUserName = LsUserName
        UserRank = LsRank
    End If

End Sub


Private Sub CmdCancel_Click()
    
    GbStart = False
    Call DbDisConnect
    End

End Sub

Private Sub CmdOK_Click()
    
    nPassCount = nPassCount + 1
    
    Call PassWordCheck
    
    If strConnect = "NO" Then
        If nPassCount > 3 Then
            MsgBox "ID 와 Password를 확인후에 다시 시작하십시요", 48, "경고"
            Call DbDisConnect
            End
        Else
            MsgBox "Password가 틀림니다 !", 48, "주의"
            TxtIDnum.SetFocus
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
End Sub



Private Sub TxtIDnum_GotFocus()
    
    TxtIDnum.SelStart = 0
    TxtIDnum.SelLength = Len(TxtIDnum.Text)
    
End Sub


Private Sub TxtIDnum_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then KeyCode = 0: SendKeys "{TAB}"
    
End Sub

Private Sub TxtPasswd_GotFocus()
    
    TxtPasswd.SelStart = 0
    TxtPasswd.SelLength = Len(TxtPasswd.Text)
    
End Sub


Private Sub TxtPasswd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then KeyCode = 0: SendKeys "{TAB}"
    
End Sub

