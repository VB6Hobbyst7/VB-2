VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Begin VB.Form medNewPasswd 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "비밀번호 변경"
   ClientHeight    =   4005
   ClientLeft      =   6615
   ClientTop       =   2220
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FAF3FA&
      Height          =   4005
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   4215
      TabIndex        =   6
      Top             =   0
      Width           =   4275
      Begin VB.TextBox txtLoginId 
         Appearance      =   0  '평면
         Height          =   345
         IMEMode         =   3  '사용 못함
         Left            =   2310
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1230
         Width           =   1605
      End
      Begin MedControls1.LisLabel lblUserId 
         Height          =   345
         Left            =   885
         TabIndex        =   10
         Top             =   555
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   609
         BackColor       =   16446458
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.CommandButton cmdCANCEL 
         BackColor       =   &H00E0CFE0&
         Caption         =   "취소(&X)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1725
         MaskColor       =   &H000000FF&
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0CFE0&
         Caption         =   "확인(&O)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2895
         MaskColor       =   &H000000FF&
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtOLD 
         Appearance      =   0  '평면
         Height          =   345
         IMEMode         =   3  '사용 못함
         Left            =   2310
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1665
         Width           =   1605
      End
      Begin VB.TextBox txtNEW1 
         Appearance      =   0  '평면
         Height          =   345
         IMEMode         =   3  '사용 못함
         Left            =   2310
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2100
         Width           =   1620
      End
      Begin VB.TextBox txtNEW2 
         Appearance      =   0  '평면
         Height          =   345
         IMEMode         =   3  '사용 못함
         Left            =   2310
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2550
         Width           =   1620
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "사용자ID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   180
         Left            =   1410
         TabIndex        =   11
         Top             =   1305
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   330
         Picture         =   "medNewPasswd.frx":0000
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblOld 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "비밀번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   315
         Left            =   645
         TabIndex        =   9
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label lblNew 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "새 비밀번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Top             =   2205
         Width           =   1335
      End
      Begin VB.Label lblMore 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "한번 더.."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   315
         Left            =   210
         TabIndex        =   7
         Top             =   2655
         Width           =   2070
      End
   End
End
Attribute VB_Name = "medNewPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCANCEL_Click()
   
   Dim Resp As VbMsgBoxResult
   
   If txtNEW1.Text <> "" Then
      Resp = MsgBox("저장없이 종료하시겠습니까?", vbYesNo)
      If Resp = vbNo Then Exit Sub
   End If
   
   Unload Me
   Set medNewPasswd = Nothing
   
End Sub

Private Sub cmdOK_Click()
   
   If txtLoginId.Text = "" Then
      MsgBox "사용자 ID를 입력하세요."
      LoginId.SetFocus
      Exit Sub
   End If
   
   If txtNEW2.Text <> txtNEW1.Text Then
      MsgBox "새 비밀번호를 한번 더 정확히 입력하세요.."
      txtNEW2.SetFocus
      Exit Sub
   End If
   
   Dim SqlStmt As String
   
   SqlStmt = "Update " & TB_LAB015 & " set logonid  " & DBStr(txtLoginId.Text, 3) & " password " & DBStr(txtNEW1.Text, 2) & _
                 " where empid = " & MyUser.EmpId
   
   DbConn.BeginTrans
   DbConn.Execute (SqlStmt)
   DbConn.CommitTrans
   
   Unload Me
   Set medNewPasswd = Nothing

End Sub

Private Sub Form_Load()
   'Me.Show
   lblUserId.Caption = MyUser.EmpId & "  " & MyUser.EmpNm
   txtLoginId.Text = MyUser.LogonId
   cmdOK.Enabled = False
End Sub

Private Sub txtLoginId_GotFocus()
   With txtLoginId
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtNEW1_GotFocus()
   With txtNEW1
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub txtNEW2_GotFocus()
   With txtNEW2
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub txtNEW2_LostFocus()
   If ActiveControl.Name = cmdCANCEL.Name Then Exit Sub
   If txtNEW2.Text <> txtNEW1.Text Then
      MsgBox "새 비밀번호를 한번 더 입력하세요.."
      txtNEW2.SetFocus
      Exit Sub
   End If
   cmdOK.Enabled = True
   cmdOK.SetFocus
End Sub

Private Sub txtOLD_LostFocus()
   If ActiveControl.Name = cmdCANCEL.Name Then Exit Sub
   If txtOLD.Text <> MyUser.Password Then
      MsgBox "기존 비밀번호를 정확히 입력하세요.."
      txtOLD.SetFocus
      Exit Sub
   End If
   txtNEW1.SetFocus
   
End Sub
