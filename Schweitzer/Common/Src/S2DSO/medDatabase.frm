VERSION 5.00
Begin VB.Form frmDatabase 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Setup Database"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.ComboBox cboDBType 
      Height          =   300
      ItemData        =   "medDatabase.frx":0000
      Left            =   1800
      List            =   "medDatabase.frx":000D
      Style           =   2  '드롭다운 목록
      TabIndex        =   12
      Top             =   180
      Width           =   2475
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00B9CAD0&
      Caption         =   "&Reg등록"
      Height          =   400
      Left            =   420
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   2445
      Width           =   1000
   End
   Begin VB.TextBox txtPwd 
      BorderStyle     =   0  '없음
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1815
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1905
      Width           =   2415
   End
   Begin VB.TextBox txtLogin 
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   1815
      TabIndex        =   2
      Top             =   1485
      Width           =   2415
   End
   Begin VB.TextBox txtDbNm 
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   1815
      TabIndex        =   1
      Top             =   1050
      Width           =   2415
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      Height          =   255
      Left            =   1815
      TabIndex        =   0
      Top             =   645
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00B9CAD0&
      Caption         =   "취소(&X)"
      Height          =   400
      Left            =   2190
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   2460
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00B9CAD0&
      Caption         =   "연결(&O)"
      Height          =   400
      Left            =   3255
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   2460
      Width           =   1000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "※"
      Height          =   180
      Left            =   4365
      TabIndex        =   13
      Top             =   690
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Type     :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00172C2D&
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   255
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   3
      Left            =   1785
      Shape           =   4  '둥근 사각형
      Top             =   1890
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   2
      Left            =   1785
      Shape           =   4  '둥근 사각형
      Top             =   1470
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   1
      Left            =   1785
      Shape           =   4  '둥근 사각형
      Top             =   1035
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   0
      Left            =   1785
      Shape           =   4  '둥근 사각형
      Top             =   615
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00172C2D&
      Height          =   195
      Index           =   3
      Left            =   465
      TabIndex        =   9
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Login    :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00172C2D&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Database :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00172C2D&
      Height          =   195
      Index           =   1
      Left            =   465
      TabIndex        =   7
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Server   :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00172C2D&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   690
      Width           =   1215
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event RegCanceled()
Public Event RegSaved()

Private mvarOldServerNm As String
Private mvarOldDatabaseNm As String
Private mvarOldLoginId As String
Private mvarOldPassword As String
Private mvarOldDbType As Integer

Private mvarRegChanged As Boolean

Public Property Let OldServerNm(ByVal pValue As String)
    mvarOldServerNm = pValue
End Property
Public Property Let OldDatabaseNm(ByVal pValue As String)
    mvarOldDatabaseNm = pValue
End Property
Public Property Let OldLoginId(ByVal pValue As String)
    mvarOldLoginId = pValue
End Property
Public Property Let OldPassword(ByVal pValue As String)
    mvarOldPassword = pValue
End Property
Public Property Let OldDbType(ByVal pValue As Integer)
    mvarOldDbType = pValue
End Property
Public Property Get RegChanged() As Boolean
    RegChanged = mvarRegChanged
End Property


Public Sub ApplyButton(ByVal pChk As String)
   If pChk = "Onlyreg" Then cmdCancel.Enabled = False
   If pChk = "SetDb" Then cmdCancel.Enabled = True
End Sub

Private Sub cmdCancel_Click()
   RaiseEvent RegCanceled
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sTemp As String, sBldCd As String, sBldNm As String, sBldNo As Integer

    If txtServer.Text = "" Or txtDbNm.Text = "" Or txtLogin.Text = "" Then Exit Sub
    
    'DB정보가 변경되었거나 기존 DB연결에 실패했을 경우 재연결....
    If txtServer.Text <> mvarOldServerNm Or txtDbNm.Text <> mvarOldDatabaseNm Or _
       txtLogin.Text <> mvarOldLoginId Or txtPwd.Text <> mvarOldPassword Then
             
        SaveSetting RegAppName, RegSsSvr, RegK1Svr, txtServer.Text
        SaveSetting RegAppName, RegSsSvr, RegK2Svr, txtDbNm.Text
        SaveSetting RegAppName, RegSsSvr, RegK3Svr, txtLogin.Text
        SaveSetting RegAppName, RegSsSvr, RegK4Svr, txtPwd.Text
        SaveSetting RegAppName, RegSsSvr, RegK5Svr, cboDBType.ListIndex
    
        mvarRegChanged = True
    Else
        
        mvarRegChanged = False
        
    End If
    
    RaiseEvent RegSaved
    Unload Me
    
End Sub

Private Sub cmdSave_Click()

   If txtServer.Text = "" Or txtDbNm.Text = "" Or txtLogin.Text = "" Then Exit Sub

    SaveSetting RegAppName, RegSsSvr, RegK1Svr, txtServer.Text
    SaveSetting RegAppName, RegSsSvr, RegK2Svr, txtDbNm.Text
    SaveSetting RegAppName, RegSsSvr, RegK3Svr, txtLogin.Text
    SaveSetting RegAppName, RegSsSvr, RegK4Svr, txtPwd.Text
    SaveSetting RegAppName, RegSsSvr, RegK5Svr, cboDBType.ListIndex

   MsgBox "Registry에 Database정보가 저장되었습니다.", vbInformation, "메세지"
End Sub

Private Sub Form_Load()
    Call medAlwaysOn(Me, 1)
End Sub

Private Sub txtDbNm_GotFocus()
    With txtDbNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLogin_GotFocus()
    With txtLogin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPwd_GotFocus()
    With txtPwd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK.SetFocus
End Sub


Public Sub SetServerInfo()
    
    txtServer.Text = mvarOldServerNm
    txtDbNm.Text = mvarOldDatabaseNm
    txtLogin.Text = mvarOldLoginId
    txtPwd.Text = mvarOldPassword
    If mvarOldDbType >= 0 Or mvarOldDbType <= cboDBType.ListCount - 1 Then
        cboDBType.ListIndex = mvarOldDbType
    Else
        cboDBType.ListIndex = -1
    End If
End Sub

Private Sub txtServer_GotFocus()
    With txtServer
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
