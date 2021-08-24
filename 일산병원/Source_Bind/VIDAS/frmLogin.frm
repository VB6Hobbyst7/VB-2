VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 로그인"
   ClientHeight    =   3375
   ClientLeft      =   3240
   ClientTop       =   2925
   ClientWidth     =   5730
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5730
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer1 
      Left            =   1170
      Top             =   2280
   End
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   -1170
      TabIndex        =   9
      Top             =   3000
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
      Left            =   3870
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2550
      Visible         =   0   'False
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
      Left            =   3870
      TabIndex        =   3
      Top             =   2100
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "* 아이디만 입력한 후 로그인하세요"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   900
      TabIndex        =   12
      Top             =   1710
      Width           =   4515
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   210
      Picture         =   "frmLogin.frx":058A
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   210
      Picture         =   "frmLogin.frx":06D4
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   210
      Picture         =   "frmLogin.frx":081E
      Top             =   2940
      Width           =   240
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID를 입력 하십시오."
      Height          =   180
      Left            =   480
      TabIndex        =   11
      Top             =   2970
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   120
      Picture         =   "frmLogin.frx":0968
      Stretch         =   -1  'True
      Top             =   90
      Width           =   555
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   465
      Left            =   360
      Top             =   900
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "일산병원"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   525
      Left            =   690
      TabIndex        =   10
      Top             =   180
      Width           =   1905
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H008080FF&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   90
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   30
      Top             =   2130
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "진단검사의학과 "
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2670
      TabIndex        =   8
      Top             =   330
      Width           =   3915
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '투명
      Caption         =   "* 사용자 ID나 Password 가 잘못되었습니다."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   570
      TabIndex        =   7
      Top             =   1410
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
      Left            =   4830
      TabIndex        =   6
      Top             =   2970
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
      Left            =   4050
      TabIndex        =   5
      Top             =   2970
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
      Left            =   2610
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
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
      Left            =   2610
      TabIndex        =   1
      Top             =   2130
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   -30
      Top             =   2130
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblEquipName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "ABL 800 Basic Interface"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   885
      Left            =   600
      TabIndex        =   0
      Top             =   900
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   0
      Picture         =   "frmLogin.frx":10B6
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   5745
   End
   Begin VB.Image Image3 
      Height          =   2010
      Left            =   0
      Picture         =   "frmLogin.frx":2040
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Dim i As Integer
    lblErr = ""
    
    GetSetup

    lblEquipName.Caption = App.EXEName
    
    imgNet1.ZOrder 0
    Timer1.Interval = 500
    Timer1.Enabled = True

'    lblEquipName.Caption = App.Title

'''    For i = 1 To 3
'''        If Not Connect_PRServer Then
'''            cn_cnt = cn_cnt + 1
'''            If cn_cnt = 3 Then
''''                If Not Connect_DRServer Then
'''                    MsgBox "연결되지 않았습니다."
'''                    cn_Server_Flag = False
'''                    Exit Sub
''''                Else
''''                    cn_Server_Flag = True
''''                End If
'''            End If
'''        Else
'''            cn_Server_Flag = True
'''            Exit For
'''        End If
'''    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Unload frmInterface
    
End Sub

Private Sub lblCancel_Click()

    Unload Me
    
End Sub

Private Sub lblCommit_Click()
'Dim lsWK As Integer
Dim blnUser As Boolean

    blnUser = False

    If Trim(txtID.Text) = "" Then
        lblErr = "* 사용자 아이디를 입력하세요."
        txtID.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtPW.Text) = "" Then
'        lblErr = "* 비밀번호를 입력하세요."
'        txtPW.SetFocus
'        Exit Sub
'    End If
    
    'blnUser = GetUser(Trim(txtID.Text), Trim(txtPW.Text))
    gUserID = Trim(txtID.Text)
'    gDB_Parm.User = gUserID
     
'    If blnUser = False Then
'        lblErr = "* 아이디나 패스워드가 일치하지 않습니다."
'        txtPW.Text = ""
'        txtID.Text = ""
'        txtID.SetFocus
'    Else
        lblErr = ""
        'frmInterface.lblUser.Caption = gUserID
        gIFUser = Trim(txtID.Text)
        frmInterface.Show 0
        Unload Me
'    End If
    
    
    
'    If Trim(gWorker_Info.WK_PW) = Trim(txtPW.Text) And Trim(gWorker_Info.WK_ID) = Trim(txtID.Text) Then
'        lblErr = ""
'        frmInterface.lblUser.Caption = "사용자 : " & gWorker_Info.WK_NM
'        frmInterface.Show 0
'        Me.Hide
'
'    Else
'        lblErr = "* 비밀번호를 확인하세요."
'        txtPW.Text = ""
'        txtPW.SetFocus
'    End If
End Sub

Private Sub Timer1_Timer()
    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call lblCancel_Click
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
            lblCommit_Click
'            txtPW.Text = "1"
'            txtPW.SetFocus
        End If
'            lsWK = Get_WKID(Trim(txtID.Text))
'            If lsWK > 0 Then
'                lblErr = ""
'                txtPW.SetFocus
'
'            Else
'                lblErr = "* 존재하지 않는 아이디입니다."
'                txtID.Text = ""
'                txtID.SetFocus
'                Exit Sub
'            End If
'        End If
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
