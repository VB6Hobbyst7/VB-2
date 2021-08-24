VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   " 로그인"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   7125
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":000C
   ScaleHeight     =   4770
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtPW 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3900
      Width           =   1545
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   60
   End
   Begin VB.CheckBox chkSave 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID저장"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4140
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2940
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   3240
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2940
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   3900
      Width           =   705
   End
   Begin VB.OptionButton optServer 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "ERP"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   5670
      TabIndex        =   5
      Top             =   3660
      Width           =   1005
   End
   Begin VB.OptionButton optServer 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "LOCAL"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   5670
      TabIndex        =   6
      Top             =   3390
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "국도 바코드 발행 시스템"
      BeginProperty Font 
         Name            =   "문체부 바탕체"
         Size            =   21.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   720
      TabIndex        =   10
      Top             =   1470
      Width           =   5685
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   4695
      Left            =   60
      Top             =   60
      Width           =   7035
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   4140
      Picture         =   "frmLogin.frx":7163
      Top             =   3990
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   4140
      Picture         =   "frmLogin.frx":72AD
      Top             =   3990
      Width           =   240
   End
   Begin VB.Label lblUserNm 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "홍길동"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      TabIndex        =   9
      Top             =   3570
      Width           =   1395
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID를 입력 하십시오."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4500
      TabIndex        =   8
      Top             =   4020
      Width           =   2205
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   4140
      Picture         =   "frmLogin.frx":73F7
      Top             =   3990
      Width           =   240
   End
   Begin VB.Label lblHospNm 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   $"frmLogin.frx":7541
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   8190
      TabIndex        =   7
      Top             =   1500
      Width           =   5565
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4140
      Picture         =   "frmLogin.frx":7562
      Top             =   210
      Width           =   2490
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Call txtPW_KeyPress(vbKeyReturn)
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        End
    End If
    
End Sub

Private Sub Form_Load()

    imgNet1.ZOrder 0
    Timer1.Interval = 500
    Timer1.Enabled = True
    
    Call CtlInitializing
    
End Sub

Private Sub optServer_Click(Index As Integer)

    'ERP
    If Index = 0 Then
        Call WritePrivateProfileString("DB", "DBCONN", "2", App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("DB", "DBTYPE", "2", App.PATH & "\KDBAR.ini")
    'LOCAL
    Else
        Call WritePrivateProfileString("DB", "DBCONN", "1", App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("DB", "DBTYPE", "4", App.PATH & "\KDBAR.ini")
    End If
    
End Sub

Private Sub Picture1_DblClick()
    txtID.Text = "KDBAR"
    txtPW.Text = "0810"
    txtPW.SetFocus
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


Private Sub txtID_GotFocus()
    
    txtID.SelStart = 0
    txtID.SelLength = Len(txtID.Text)

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    Dim strUserName As String
    
    If KeyAscii = vbKeyReturn Then
        If txtID.Text <> "" Then
            strUserName = ""
            lblUserNm.Caption = ""
            strUserName = Get_UserName(txtID.Text)
            
            If strUserName = "" Then
                txtID.SetFocus
            Else
                lblUserNm.Caption = strUserName
                txtPW.SetFocus
            End If
        Else
            MsgBox "아이디 또는 비밀번호를 확인해주세요", vbOKOnly + vbCritical, Me.Caption
        End If
    End If

End Sub

Private Sub txtID_LostFocus()
    
    If txtID.Text <> "" Then
        lblUserNm.Caption = Get_UserName(txtID.Text)
    End If

End Sub

Private Sub txtPW_GotFocus()
    
    txtPW.SelStart = 0
    txtPW.SelLength = Len(txtPW.Text)

End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    Dim i       As Integer
    Dim strPW   As String
    Dim strUserName As String
    
    If KeyAscii = vbKeyReturn Then
        strUserName = ""
        If Trim(txtID.Text) <> "" And Trim(txtPW.Text) <> "" Then
            strUserName = Get_UserName(txtID.Text, txtPW.Text)
              
            If strUserName <> "" Then
                If chkSave.Value = "1" Then
                    
                    Call WritePrivateProfileString("USER", "USERID", txtID.Text, App.PATH & "\KDBAR.ini")
                    Call WritePrivateProfileString("USER", "USERNM", lblUserNm.Caption, App.PATH & "\KDBAR.ini")
                    If chkSave.Value = "1" Then
                        Call WritePrivateProfileString("USER", "SAVEPW", "Y", App.PATH & "\KDBAR.ini")
                    Else
                        Call WritePrivateProfileString("USER", "SAVEPW", "", App.PATH & "\KDBAR.ini")
                    End If
                End If
                        
                'frmMain.Show
                
                frmMDI.Show
                Unload Me
            Else
                MsgBox "아이디 또는 비밀번호를 확인해주세요", vbOKOnly + vbCritical, Me.Caption
                
                txtPW.SelStart = 0
                txtPW.SelLength = Len(txtPW.Text)
                
            End If
        Else
            MsgBox "아이디 또는 비밀번호를 확인해주세요", vbOKOnly + vbCritical, Me.Caption
        
            txtPW.SelStart = 0
            txtPW.SelLength = Len(txtPW.Text)
        End If
    End If
    
End Sub


Public Sub CtlInitializing()
    Dim i           As Integer
    Dim strPW       As String
    Dim strOrgPW    As String
    
    txtID.Text = ""
    txtPW.Text = ""
    lblUserNm.Caption = ""
    
    If gDBCONN = "1" Then
        optServer(1).Value = True
    Else
        optServer(0).Value = True
    End If
    
    If gKUKDO.SAVEPW = "Y" Then
        chkSave.Value = "1"
        txtID.Text = gKUKDO.USERID
        'txtPW.Text = gKUKDO.USERPW
        lblUserNm.Caption = gKUKDO.USERNM
    End If
    
End Sub

