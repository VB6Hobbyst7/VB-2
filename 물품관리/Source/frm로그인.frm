VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm로그인 
   BorderStyle     =   0  '없음
   Caption         =   "사용자 확인"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSFrame SSFrame1 
      Height          =   1335
      Left            =   3870
      TabIndex        =   0
      Top             =   2670
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2355
      _Version        =   262144
      Begin VB.TextBox txtPswd 
         Alignment       =   2  '가운데 맞춤
         Height          =   300
         IMEMode         =   3  '사용 못함
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtUserNm 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   510
         Width           =   1575
      End
      Begin VB.TextBox txtUserCd 
         Alignment       =   2  '가운데 맞춤
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   1575
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   300
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   262144
         Caption         =   "사용자번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   300
         Left            =   210
         TabIndex        =   4
         Top             =   510
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   262144
         Caption         =   "사용자명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   300
         Left            =   210
         TabIndex        =   6
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   262144
         Caption         =   "비밀번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdConfirm 
      Height          =   420
      Left            =   4680
      TabIndex        =   7
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "확 인"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   420
      Left            =   5790
      TabIndex        =   8
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "취 소"
      ButtonStyle     =   2
   End
End
Attribute VB_Name = "frm로그인"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me
    End

End Sub

Private Sub cmdConfirm_Click()

    If Len(txtUserCd.Text) > 0 Then
        If txtPswd.Tag = Trim(txtPswd.Text) Then
            
            gUserId = Trim(txtUserCd.Text)
            frm메인화면.stsBar.Panels(3).Text = txtUserNm.Text
            frm메인화면.Enabled = True
            
            Unload Me
        Else
            MsgBox "비밀번호가 잘못되었습니다.!", vbCritical
            txtPswd.SetFocus
        End If
    Else
        MsgBox "사용자번호를 입력하세요.!", vbCritical
        txtUserCd.SetFocus
    End If
    
End Sub

Private Sub Form_Load()

    txtUserCd.Text = ""
    txtUserNm.Text = ""
    txtPswd.Text = ""
    
End Sub

Private Sub txtPswd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cmdConfirm.SetFocus

End Sub

Private Sub txtUserCd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then txtPswd.SetFocus
    
End Sub

Private Sub txtUserCd_LostFocus()

    If Len(txtUserCd.Text) > 0 And Me.ActiveControl.Name <> "cmdCancel" Then
        txtUserCd.Text = UCase(txtUserCd.Text)
        gSql = "select usernm, pswd from mstUSER where usercd = '" & Trim(txtUserCd.Text) & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    txtUserNm.Text = "" & .Fields("usernm").Value
                    txtPswd.Tag = "" & .Fields("pswd").Value
                Else
                    MsgBox "등록되지 않은 사용자 입니다.!", vbCritical
                    txtUserCd.Text = ""
                    txtUserNm.Text = ""
                    txtPswd.Tag = ""
                    txtUserCd.SetFocus
                End If
                .Close
            End If
        End With
    End If

End Sub
