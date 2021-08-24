VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Login"
   ClientHeight    =   3570
   ClientLeft      =   3240
   ClientTop       =   2925
   ClientWidth     =   5265
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5265
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2700
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtPW 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '사용 못함
      Left            =   3030
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2130
      Width           =   1575
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3030
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   945
      Left            =   510
      Top             =   390
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "순천향대학교 병원 응급실"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   780
      TabIndex        =   9
      Top             =   900
      Width           =   3885
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '투명
      Caption         =   "* 사용자 ID나 Password 가 잘못되었습니다."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   3360
      Width           =   4515
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3900
      TabIndex        =   6
      Top             =   2820
      Width           =   645
   End
   Begin VB.Label lblCommit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   2820
      Width           =   645
   End
   Begin VB.Label lblPW 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "비밀번호 :"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1770
      TabIndex        =   2
      Top             =   2100
      Width           =   1155
   End
   Begin VB.Label lblID 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "아이디 :"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      Top             =   1710
      Width           =   1155
   End
   Begin VB.Label lblEquipName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Cobas Taqman Interface"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   555
      Left            =   780
      TabIndex        =   0
      Top             =   330
      Width           =   4005
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    lblErr = ""
'''    GetSetup
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    Unload frmInterface
    
End Sub

Private Sub lblCancel_Click()

    Unload Me
    
End Sub

Private Sub lblCommit_Click()

Dim lsWK As Integer

    If Trim(txtID.Text) = "" Then
        lblErr = "* 사용자 아이디를 입력하세요."
        txtID.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPW.Text) = "" Then
        lblErr = "* 비밀번호를 입력하세요."
        txtPW.SetFocus
        Exit Sub
    End If
    GetSetup
    gUserName = ""
    
    Connect_Server
    SQL = "SELECT USER_NAME FROM CPL0135 " & vbCrLf & _
          "WHERE LOGIN_ID = '" & Trim(txtID.Text) & "' AND PASSWORD = '" & Trim(txtPW.Text) & "'"
    res = db_select_Col(gServer, SQL)
    
    DisConnect_Server
    
    If res > 0 Then
        lblErr = ""
        gUserName = Trim(gReadBuf(0))
        gUserID = Trim(txtID.Text)
        frmInterface.Caption = " COBAS TaqMan Interface " & "[" & gUserName & "]"
        frmInterface.Show 0
        Unload Me
    Else
        lblErr = "* 아이디/패스워드가 일치하지 않습니다."
        txtPW.Text = ""
        txtID.Text = ""
        txtID.SetFocus
        
    End If
    
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lsWK As Integer

    If KeyCode = 13 Then
        If Trim(txtID.Text) = "" Then
            lblErr = "* 사용자 아이디를 입력하세요."
            txtID.SetFocus
            Exit Sub
        Else
            txtPW.SetFocus
        End If
    End If
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtPW.Text) = "" Then
            lblErr = "* 비밀번호를 입력하세요."
            txtPW.SetFocus
            Exit Sub
        Else
            lblErr = ""
            lblCommit_Click
            
        End If
        
    End If
End Sub
