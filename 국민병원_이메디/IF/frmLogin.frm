VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  '단일 고정
   Caption         =   " 로그인"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5670
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture2 
      Align           =   2  '아래 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   5670
      TabIndex        =   4
      Top             =   2235
      Width           =   5670
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "저장"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4260
         TabIndex        =   2
         Top             =   690
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Left            =   150
         Top             =   600
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
         Left            =   2550
         TabIndex        =   0
         Top             =   240
         Width           =   1575
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
         Left            =   2550
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblUserNm 
         BackStyle       =   0  '투명
         Caption         =   "홍길동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblID 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "아이디 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1290
         TabIndex        =   6
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblPW 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "비밀번호 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1290
         TabIndex        =   5
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   5670
      TabIndex        =   3
      Top             =   0
      Width           =   5670
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmLogin.frx":000C
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmLogin.frx":0156
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label labMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "사용자 ID를 입력 하십시오."
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   2010
         Width           =   2205
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmLogin.frx":02A0
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label lblPartNm 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "생화학검사실"
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
         Left            =   3780
         TabIndex        =   11
         Top             =   780
         Width           =   1755
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00400000&
         BorderWidth     =   2
         Height          =   375
         Left            =   390
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label lblMachNm 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "ABL 800 Basic "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   1260
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabNm 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "진단검사의학과"
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
         Left            =   3450
         TabIndex        =   8
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label lblHospNm 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "전남대학교 병원"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
    
'    txtID.SetFocus
    
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


Private Sub txtID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtPW.SetFocus
    End If

End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
Dim i       As Integer
Dim strPW   As String

    If KeyAscii = vbKeyReturn Then
        If UCase(txtID.Text) = "OKSOFT" And UCase(txtPW.Text) = "0810" Then
            
            If chkSave.Value = "1" Then
                For i = 1 To Len(txtPW.Text)
                    strPW = strPW & Asc(Mid(txtPW.Text, i, 1))
                Next
                strPW = "%" & strPW & "@#"
                
                Call WritePrivateProfileString("HOSP", "USERID", txtID.Text, App.PATH & "\INI\" & gMACH & ".ini")
                Call WritePrivateProfileString("HOSP", "USERNM", lblUserNm.Caption, App.PATH & "\INI\" & gMACH & ".ini")
                Call WritePrivateProfileString("HOSP", "USERPW", strPW, App.PATH & "\INI\" & gMACH & ".ini")
                
            End If
            frmMain.Show
            Unload Me
        Else
            MsgBox "아이디 또는 비밀번호를 확인해주세요"
        End If
    End If
    
End Sub


Public Sub CtlInitializing()
    Dim i           As Integer
    Dim strPW       As String
    Dim strOrgPW    As String
    
    lblHospNm.Caption = gHOSP.HOSPNM
    lblLabNm.Caption = gHOSP.LABNM
    lblPartNm.Caption = gHOSP.PARTNM
    lblMachNm.Caption = gHOSP.MACHNM
    If gHOSP.SAVEPW = "Y" Then
        If gHOSP.USERPW <> "" Then
            strPW = Mid(gHOSP.USERPW, 2)
            strPW = Mid(strPW, 1, Len(strPW) - 2)
            For i = 1 To Len(strPW) Step 2
                strOrgPW = strOrgPW & Chr(Mid(strPW, i, 2))
            Next
            
            chkSave.Value = "1"
            txtID.Text = gHOSP.USERID
            txtPW.Text = strOrgPW
            lblUserNm.Caption = gHOSP.USERNM
        End If
    End If
    
End Sub

