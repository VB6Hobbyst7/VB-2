VERSION 5.00
Begin VB.Form frmDatabase 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Setup Database"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Reg등록"
      Height          =   400
      Left            =   510
      TabIndex        =   10
      Top             =   2460
      Width           =   1000
   End
   Begin VB.TextBox txtPwd 
      Height          =   300
      Left            =   1710
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtLogin 
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   1270
      Width           =   2415
   End
   Begin VB.TextBox txtDbNm 
      Height          =   300
      Left            =   1710
      TabIndex        =   1
      Top             =   860
      Width           =   2415
   End
   Begin VB.TextBox txtServer 
      Height          =   300
      Left            =   1710
      TabIndex        =   0
      Top             =   450
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소(&X)"
      Height          =   400
      Left            =   2040
      TabIndex        =   5
      Top             =   2460
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "연결(&O)"
      Height          =   400
      Left            =   3105
      TabIndex        =   4
      Top             =   2460
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Password :"
      Height          =   210
      Index           =   3
      Left            =   555
      TabIndex        =   9
      Top             =   1725
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Login :"
      Height          =   210
      Index           =   2
      Left            =   555
      TabIndex        =   8
      Top             =   1315
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Database :"
      Height          =   210
      Index           =   1
      Left            =   555
      TabIndex        =   7
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Server :"
      Height          =   210
      Index           =   0
      Left            =   555
      TabIndex        =   6
      Top             =   495
      Width           =   885
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ApplyButton(ByVal pChk As String)
   If pChk = "Onlyreg" Then cmdCancel.Enabled = False
   If pChk = "SetDb" Then cmdCancel.Enabled = True
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   'If CancelIsEnd Then End
End Sub

Private Sub cmdOK_Click()
Dim sTemp As String, sBldCd As String, sBldNm As String, sBldNo As Integer

   If txtServer.Text = "" Or txtDbNm.Text = "" Or txtLogin.Text = "" Or txtPwd.Text = "" Then Exit Sub

   'DB정보가 변경되었거나 기존 DB연결에 실패했을 경우 재연결....
   If txtServer.Text <> SB_ServerNm Or txtDbNm.Text <> SB_DatabaseNm Or _
      txtLogin.Text <> SB_LoginId Or txtPwd.Text <> SB_Password Or SB_ConnStatus = CONNECT_ERROR Then
      If SB_ConnStatus = CONNECT_SUCCESS Then DbConn.DbClose
      Call UnloadForms(frm012Database)
      
      SB_ServerNm = txtServer.Text
      SB_DatabaseNm = txtDbNm.Text
      SB_LoginId = txtLogin.Text
      SB_Password = txtPwd.Text
    
      SB_ConnStatus = DBConnect()
   
      If SB_ConnStatus = CONNECT_ERROR Then
          MsgBox "연결되지 않는 데이타베이스입니다. ", vbCritical + vbOKOnly, "Database"
          txtServer.SetFocus
          Exit Sub
      End If
   End If
   
   SB_ServerNm = txtServer.Text
   SB_DatabaseNm = txtDbNm.Text
   SB_LoginId = txtLogin.Text
   SB_Password = txtPwd.Text

   SaveSetting RegHdSvr, RegSsSvr, RegK1Svr, SB_ServerNm
   SaveSetting RegHdSvr, RegSsSvr, RegK2Svr, SB_DatabaseNm
   SaveSetting RegHdSvr, RegSsSvr, RegK3Svr, SB_LoginId
   SaveSetting RegHdSvr, RegSsSvr, RegK4Svr, SB_Password
   
   medMain.Caption = "SCHWEITZER - LIS " & App.Major & "." & App.Minor & "." & App.Revision & " (" & SB_ServerNm & ":" & SB_DatabaseNm & ")"

   Unload Me
   'If Not CancelIsEnd Then medLogOn.Show 1

End Sub

Private Sub cmdSave_Click()

   If txtServer.Text = "" Or txtDbNm.Text = "" Or txtLogin.Text = "" Or txtPwd.Text = "" Then Exit Sub

   SB_ServerNm = txtServer.Text
   SB_DatabaseNm = txtDbNm.Text
   SB_LoginId = txtLogin.Text
   SB_Password = txtPwd.Text

   SaveSetting RegHdSvr, RegSsSvr, RegK1Svr, SB_ServerNm
   SaveSetting RegHdSvr, RegSsSvr, RegK2Svr, SB_DatabaseNm
   SaveSetting RegHdSvr, RegSsSvr, RegK3Svr, SB_LoginId
   SaveSetting RegHdSvr, RegSsSvr, RegK4Svr, SB_Password

   MsgBox "Registry에 Database정보가 저장되었습니다.", vbInformation, "메세지"
End Sub

Private Sub Form_Load()

   txtServer.Text = SB_ServerNm
   txtDbNm.Text = SB_DatabaseNm
   txtLogin.Text = SB_LoginId
   txtPwd.Text = SB_Password

End Sub


Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK.SetFocus
End Sub
