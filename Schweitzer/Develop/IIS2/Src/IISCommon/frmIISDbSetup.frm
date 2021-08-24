VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIISDbSetup 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Setup Database"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin MSComDlg.CommonDialog cdgSource 
      Left            =   4005
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSource 
      BackColor       =   &H00DBE6E6&
      Caption         =   "..."
      Height          =   345
      Left            =   3998
      Style           =   1  '그래픽
      TabIndex        =   12
      Top             =   555
      Width           =   435
   End
   Begin VB.ComboBox cboDBType 
      Height          =   300
      ItemData        =   "frmIISDbSetup.frx":0000
      Left            =   1935
      List            =   "frmIISDbSetup.frx":0010
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   153
      Width           =   2055
   End
   Begin VB.TextBox txtPwd 
      BorderStyle     =   0  '없음
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1950
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1878
      Width           =   1980
   End
   Begin VB.TextBox txtUid 
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   1950
      TabIndex        =   3
      Top             =   1458
      Width           =   1980
   End
   Begin VB.TextBox txtCatalog 
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   1950
      TabIndex        =   2
      Top             =   1023
      Width           =   1980
   End
   Begin VB.TextBox txtSource 
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      Height          =   255
      Left            =   1950
      TabIndex        =   1
      Top             =   618
      Width           =   1980
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "취소(&X)"
      Height          =   400
      Left            =   2963
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   2448
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저장(&S)"
      Height          =   400
      Left            =   1958
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   2448
      Width           =   1000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Type        :"
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
      Left            =   248
      TabIndex        =   11
      Top             =   225
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   3
      Left            =   1920
      Shape           =   4  '둥근 사각형
      Top             =   1863
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   2
      Left            =   1920
      Shape           =   4  '둥근 사각형
      Top             =   1443
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   1
      Left            =   1920
      Shape           =   4  '둥근 사각형
      Top             =   1008
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0082B9BD&
      BorderWidth     =   2
      Height          =   330
      Index           =   0
      Left            =   1913
      Shape           =   4  '둥근 사각형
      Top             =   585
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Password    :"
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
      Left            =   248
      TabIndex        =   10
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "User ID     :"
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
      Left            =   248
      TabIndex        =   9
      Top             =   1500
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Catalog     :"
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
      Left            =   248
      TabIndex        =   8
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Data Source :"
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
      Left            =   248
      TabIndex        =   7
      Top             =   660
      Width           =   1560
   End
End
Attribute VB_Name = "frmIISDbSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISDbSetup
'   작성자  : 이상대
'   내  용  : 데이터베이스에 연결설정 폼
'   작성일  : 2003-12-02
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAppName    As String               'App Name (레지스트리의 프로젝트명)
Private mDbType     As String               'DB Type (0:Oracle, 1:Sybase, 2:MS-SQL)
Private mSource     As String               'Data Source
Private mCatalog    As String               'Initial Catalog
Private mUid        As String               'User ID
Private mPwd        As String               'Password

Public Event UserExit()                     '사용자가 DB설정을 종료시 발생

Public Property Let AppName(ByVal vData As String)
    mAppName = vData
End Property

Public Property Let DbType(ByVal vData As String)
    mDbType = vData
End Property

Public Property Let Source(ByVal vData As String)
    mSource = vData
End Property

Public Property Let Catalog(ByVal vData As String)
    mCatalog = vData
End Property

Public Property Let Uid(ByVal vData As String)
    mUid = vData
End Property

Public Property Let Pwd(ByVal vData As String)
    mPwd = vData
End Property

Private Sub Form_Load()
    Dim objCom As clsIISCommon
    
    Set objCom = New clsIISCommon
    objCom.mAlwaysOn Me, ccOn
    Set objCom = Nothing
    
    txtSource.Text = mSource
    txtCatalog.Text = mCatalog
    txtUid.Text = mUid
    txtPwd.Text = mPwd
    cboDBType.ListIndex = IIf(mDbType = "", 0, mDbType)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISDbSetup = Nothing
End Sub

Private Sub cmdSave_Click()
    SaveSetting mAppName, cDBSERVER, cDBTYPE, cboDBType.ListIndex
    SaveSetting mAppName, cDBSERVER, cSOURCE, Trim(txtSource.Text)
    SaveSetting mAppName, cDBSERVER, cCATALOG, Trim(txtCatalog.Text)
    SaveSetting mAppName, cDBSERVER, cUID, Trim(txtUid.Text)
    SaveSetting mAppName, cDBSERVER, cPWD, Trim(txtPwd.Text)
    
    If cboDBType.ListIndex = 3 Then
        SaveSetting mAppName, "App", "ClientDb", Trim(txtSource.Text)
    End If
    Unload Me
    DoEvents
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    RaiseEvent UserExit
End Sub

Private Sub cmdSource_Click()
    With cdgSource
        .InitDir = App.Path
        .Filter = "Microsoft Access(*.mdb)|*.mdb"
        .DefaultExt = "*.mdb"
        .ShowOpen
        
        txtSource.Text = .FileName
    End With
End Sub

Private Sub cboDBType_Click()
    '## ORACLE은 Catalog 항목, ACCESS는 Catalog, Uid항목을 입력할 필요없음!!
    Select Case cboDBType.ListIndex
        Case 0
            cmdSource.Enabled = False
            txtCatalog.Text = ""
            txtCatalog.Enabled = False
        Case 3
            cmdSource.Enabled = True
            txtCatalog.Text = ""
            txtUid.Text = ""
            txtCatalog.Enabled = False
            txtUid.Enabled = False
        Case Else
            cmdSource.Enabled = False
            txtCatalog.Enabled = True
            txtUid.Enabled = True
    End Select
End Sub

Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSource_GotFocus()
    With txtSource
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCatalog_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCatalog_GotFocus()
    With txtCatalog
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtUid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtUid_GotFocus()
    With txtUid
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPwd_GotFocus()
    With txtPwd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

