VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   " 로그인"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   4050
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":000C
   ScaleHeight     =   5175
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   450
   End
   Begin VB.CheckBox chkSave 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   4710
      Width           =   825
   End
   Begin VB.TextBox txtPW 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      IMEMode         =   3  '사용 못함
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3930
      Width           =   1545
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   5
      Top             =   3300
      Width           =   1545
   End
   Begin HSCotrol.CButton cmdConfirm 
      Height          =   585
      Left            =   2970
      TabIndex        =   9
      Top             =   3270
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1032
      BackColor       =   16777215
      Caption         =   "확인"
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
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   16711680
   End
   Begin HSCotrol.CButton cmdExit 
      Height          =   255
      Left            =   2970
      TabIndex        =   10
      Top             =   3930
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   450
      BackColor       =   16777215
      Caption         =   "종료"
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
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Interface Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   900
      TabIndex        =   11
      Top             =   2220
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHospNm 
      Alignment       =   1  '오른쪽 맞춤
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
      Left            =   360
      TabIndex        =   0
      Top             =   510
      Width           =   3345
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   150
      Picture         =   "frmLogin.frx":7163
      Top             =   270
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   5115
      Left            =   30
      Top             =   30
      Width           =   3975
   End
   Begin VB.Label lblUserNm 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "홍길동"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   1410
      TabIndex        =   7
      Top             =   3630
      Width           =   1395
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   180
      Picture         =   "frmLogin.frx":89D5
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   180
      Picture         =   "frmLogin.frx":8B1F
      Top             =   4680
      Width           =   240
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID를 입력 하십시오."
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   4710
      Width           =   2205
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   180
      Picture         =   "frmLogin.frx":8C69
      Top             =   4680
      Width           =   240
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
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   210
      TabIndex        =   3
      Top             =   1860
      Width           =   3615
      WordWrap        =   -1  'True
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
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1260
      Width           =   1755
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1590
      TabIndex        =   1
      Top             =   990
      Width           =   2085
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirm_Click()

    Call LoginConfirm

End Sub

Private Sub cmdExit_Click()
    
    End

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

Private Sub Picture1_DblClick()
    txtID.Text = "OKSOFT"
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


Private Sub txtID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtPW.SetFocus
    End If

End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        Call LoginConfirm
    
    End If
    
End Sub

Private Sub LoginConfirm()
    Dim i       As Integer
    Dim strPW   As String

    If UCase(txtID.Text) = "OKSOFT" And UCase(txtPW.Text) = "0810" Then
        
        If chkSave.Value = "1" Then
            For i = 1 To Len(txtPW.Text)
                strPW = strPW & Asc(Mid(txtPW.Text, i, 1))
            Next
            strPW = "%" & strPW & "@#"
            
            Call WritePrivateProfileString("HOSP", "USERID", txtID.Text, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("HOSP", "USERNM", lblUserNm.Caption, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("HOSP", "USERPW", strPW, App.PATH & "\INI\" & gMACH & ".ini")
            If chkSave.Value = "1" Then
                Call WritePrivateProfileString("HOSP", "SAVEPW", "Y", App.PATH & "\INI\" & gMACH & ".ini")
            Else
                Call WritePrivateProfileString("HOSP", "SAVEPW", "", App.PATH & "\INI\" & gMACH & ".ini")
            End If
        End If
        
        
        frmInterface.Show
        Unload Me
    Else
        If UCase(txtID.Text) <> "" Then
            gHOSP.USERID = txtID.Text
            Call WritePrivateProfileString("HOSP", "USERID", txtID.Text, App.PATH & "\INI\" & gMACH & ".ini")
            
            If chkSave.Value = "1" Then
                Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")
            Else
                Call WritePrivateProfileString("HOSP", "SAVEAUTO", "", App.PATH & "\INI\" & gMACH & ".ini")
            End If
            
            frmInterface.Show
            Unload Me
        Else
            MsgBox "아이디 또는 비밀번호를 확인해주세요"
        End If
    End If

End Sub

Private Sub CtlInitializing()
    Dim i           As Integer
    Dim strPW       As String
    Dim strOrgPW    As String
    
    lblHospNm.Caption = gHOSP.HOSPNM
    lblLabNm.Caption = gHOSP.LABNM
    lblPartNm.Caption = gHOSP.PARTNM
    lblMachNm.Caption = gHOSP.MACHNM
    lblUserNm.Caption = gHOSP.USERNM
    chkSave.Value = IIf(gHOSP.SAVEAUTO = "Y", "1", "0")
    
    If gHOSP.SAVEAUTO = "Y" Then
        txtID.Text = gHOSP.USERID
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

