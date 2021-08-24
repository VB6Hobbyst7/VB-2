VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiIISMain 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Schweitzer IIS"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10935
   Icon            =   "mdiIISMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdiIISMain.frx":144A
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   1  '최소화
   Begin VB.Timer tmrMsg 
      Interval        =   5000
      Left            =   1770
      Top             =   1080
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   1035
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picMenu 
      Align           =   1  '위 맞춤
      Height          =   840
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   10875
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      Begin MSComctlLib.Toolbar tbrToolbar 
         Height          =   330
         Left            =   4155
         TabIndex        =   2
         Top             =   0
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
      End
      Begin VB.Label lblMenuNm 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Interface System"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   3975
      End
      Begin VB.Shape shpSubMenu 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   555
         Left            =   105
         Top             =   135
         Width           =   4005
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   645
         Left            =   60
         Top             =   90
         Width           =   4095
      End
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8025
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCommon 
      Left            =   300
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIISMain.frx":29656
            Key             =   "IIS600"
            Object.Tag             =   "Master,Master"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIISMain.frx":2A228
            Key             =   "IIS000"
            Object.Tag             =   "Exit,Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIISMain.frx":2FF8A
            Key             =   "IIS204"
            Object.Tag             =   "검사대상,검사대상조회"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuExit 
         Caption         =   "종료(&X)"
      End
   End
   Begin VB.Menu mnuIIS200 
      Caption         =   "장비 인터페이스(&I)"
      Begin VB.Menu mnuIIS201 
         Caption         =   "검사장비1"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuIIS202 
         Caption         =   "검사장비2"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuIISSEP02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIIS204 
         Caption         =   "검사대상 조회"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuIIS300 
      Caption         =   "결과등록/조회(&R)"
      Begin VB.Menu mnuIIS301 
         Caption         =   "결과등록(&R)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuIIS600 
      Caption         =   "&Master"
      Begin VB.Menu NODE3 
         Caption         =   "검사장비 관련"
         Begin VB.Menu mnuIIS609 
            Caption         =   "검사장비 관리"
         End
         Begin VB.Menu mnuIIS610 
            Caption         =   "검사장비 통신설정"
         End
         Begin VB.Menu mnuIIS611 
            Caption         =   "장비별 검사항목 설정"
         End
         Begin VB.Menu mnuIIS612 
            Caption         =   "사용장비 선택"
         End
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuEditMenu 
         Caption         =   "메뉴편집"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "창(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHor 
         Caption         =   "수평 바둑판식 배열(&H)"
      End
      Begin VB.Menu mnuVer 
         Caption         =   "수직 바둑판식 배열(&V)"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "계단식 배열(&C)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(H)"
      Visible         =   0   'False
      Begin VB.Menu mnuContent 
         Caption         =   "목차(&C)"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "색인(&I)"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Schweitzer IIS 정보(&N)"
      End
   End
End
Attribute VB_Name = "mdiIISMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : mdiIISMain.frm (우리LIS랑 조인할때 사용)
'   작성자  :
'   내  용  : 메인폼
'   작성일  : 2003-12-26
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqpShow As clsIISInstrShow     '장비표시 클래스

Private Sub MDIForm_Initialize()
    Dim objSplash   As clsIISSplash
    Dim strProNm    As String           '프로그램 이름
    Dim strVer      As String           '프로그램 버전
    Dim strTitle    As String           'Title
    
    '## 프로그램 중복실행 방지
    If App.PrevInstance Then
        strTitle = App.Title
        App.Title = "... duplicate instance."
        Call RestoreWindow(strTitle)
        End
    End If
    
    '## Splash화면표시
    strProNm = App.ProductName
    strVer = App.Major & "." & App.Minor & "." & App.Revision
    Set objSplash = New clsIISSplash
    With objSplash
        .ProjectNm = strProNm
        .Version = strVer
        .Message = strProNm & " 프로그램을 실행중입니다..."
        .LoadSplash
        DoEvents
    End With
    
    Call objSplash.SetMsg("프로그램 초기화중입니다...")
    DoEvents
    
    '## 레지정보 로드, IIS.ini 파일 존재확인
    Call GetRegInfo(App.LegalTrademarks & " " & App.FileDescription, App.Path)
    If CheckINIFile = False Then
        Set objSplash = Nothing
        End
    End If
    
    '## 테이블, 컬럼, 코드정보 로드
    Call LoadTableInfo
    Call LoadCodeInfo
    Set mEqpShow = New clsIISInstrShow
    
    Call objSplash.SetMsg("Database에 연결중입니다...")
    DoEvents
    
    '## DB연결
    If DbConnect = False Then
        MsgBox "DB 연결에 문제가 있습니다. 전산실에 문의하세요.", vbCritical, "오류"
        Set objSplash = Nothing
        Call MDIForm_Unload(0)
    End If
    Set objSplash = Nothing
    
    '## 버전정보 표시
    Me.Caption = Me.Caption & " - " & strVer
End Sub

Private Sub MDIForm_Load()
    Dim objMenu     As clsIISHopMenu
    Dim objHop      As clsIISMenuInfo
    Dim frmLogOn    As frmIISLogOn
    
    '## 메인폼 정보를 IISConst 클래스에 저장
    MainFrm = Me
    
    '## 풀다운 메뉴 구성
    Set objMenu = New clsIISHopMenu
    Call objMenu.GetFullMenu
    Set objMenu = Nothing
    
    '## 로그인폼 Show
    Me.Show
    Me.ZOrder 0
    
''EPOC-2 장비 막음
    Set frmLogOn = New frmIISLogOn
    frmLogOn.Show vbModal
    If frmLogOn.IsLogOn = False Then
        Set frmLogOn = Nothing
        Call MDIForm_Unload(0)
    End If
    Set frmLogOn = Nothing
    
'    Call SetUserInfo("001401", "강지훈")
    
    '## 툴바메뉴구성
    Set objHop = New clsIISMenuInfo
    Call objHop.GetToolbar
    Set objHop = Nothing
    
'EPOC-2장비 추가
'    Call mnuIIS201_Click
    
    sbrStatus.Panels(1).Text = HOSPITALNM & " - " & EMPNM
End Sub

Public Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    If UnloadMode = 0 Then
'        Call Form1.Command1_Click
'    End If
'
'    If Form1.strLogOn = True Then
'        Cancel = 0
'        Unload Me
'    Else
'        Cancel = 1
'    End If
    
    If AppExit = False Then Cancel = 1
    Unload Me
End Sub

Public Sub MDIForm_Unload(Cancel As Integer)
    '## 전역, 현재폼에서 사용된 객체의 소멸은 여기에서!!
    Call UnloadObject
    Set mdiIISMain = Nothing
    End
End Sub

Private Sub tbrToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub picMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

'-------------------------------------------------------------------------------------------------'
'                                               메뉴항목
'-------------------------------------------------------------------------------------------------'
'## 종료
Private Sub mnuExit_Click()
    Unload Me
End Sub

'## 검사장비1
Private Sub mnuIIS201_Click()
    Screen.MousePointer = vbHourglass
    Call mEqpShow.InstrShow(mnuIIS201.Tag, mnuIIS201.Caption)
    Screen.MousePointer = vbDefault
End Sub

'## 검사장비2
Private Sub mnuIIS202_Click()
    Screen.MousePointer = vbHourglass
    Call mEqpShow.InstrShow(mnuIIS202.Tag, mnuIIS202.Caption)
    Screen.MousePointer = vbDefault
End Sub

'## 장비오류내역
Private Sub mnuIIS203_Click()
'
End Sub

'## 검사대상 조회
Private Sub mnuIIS204_Click()
    frmIIS204.Show
    frmIIS204.ZOrder 0
End Sub

'## 결과등록
Private Sub mnuIIS301_Click()
    frmIIS301.Show
    frmIIS301.ZOrder 0
End Sub

'## 검사장비 마스터
Private Sub mnuIIS609_Click()
    Call ShowMasterForm("IIS609")
End Sub

'## 검사장비 통신설정
Private Sub mnuIIS610_Click()
    Call ShowMasterForm("IIS610")
End Sub

'## 장비별 검사항목 설정
Private Sub mnuIIS611_Click()
    Call ShowMasterForm("IIS611")
End Sub

'## 사용 검사장비 선택
Private Sub mnuIIS612_Click()
    Call ShowMasterForm("IIS612")
End Sub

'## 원도우 수평정렬
Private Sub mnuHor_Click()
    Me.Arrange vbTileHorizontal
End Sub

'## 원도우 수직정렬
Private Sub mnuVer_Click()
    Me.Arrange vbTileVertical
End Sub

'## 계단식 배열
Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

'## 툴바메뉴 클릭
Private Sub tbrToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "IIS000": Unload Me
        Case "IIS600": Call ShowMasterForm
        Case mnuIIS201.Caption      '## 장비1
            Call mnuIIS201_Click
        Case mnuIIS202.Caption      '## 장비2
            Call mnuIIS202_Click
        Case "IIS204": Call mnuIIS204_Click
    End Select
End Sub

'## Manager
Private Sub ShowMasterForm(Optional ByVal pKey As String = "")
    frmIIS600.Show
    frmIIS600.ZOrder 0
    
    If pKey = "" Then Exit Sub
    frmIIS600.ShowForm (pKey)
End Sub

'-------------------------------------------------------------------------------------------------'
'                                             Pop Up 메뉴항목
'-------------------------------------------------------------------------------------------------'
'## 툴바메뉴구성
Private Sub mnuEditMenu_Click()
    Dim objHop As clsIISMenuInfo
    
    Set objHop = New clsIISMenuInfo
    objHop.ConfigToolbar
    Set objHop = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : IIS.ini 파일의 존재유무 Check
'   반환
'-----------------------------------------------------------------------------'
Private Function CheckINIFile() As Boolean
    Dim strFileNm As String     '경로+IIS.ini
    
    strFileNm = IniPath & "IIS.ini"
    If Dir(strFileNm) = "" Then
        MsgBox "IIS.ini 파일이 없습니다. 개발자에게 문의하세요.", vbInformation, "정보"
        CheckINIFile = False
    Else
        CheckINIFile = True
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 열려있는 모든(자신은 제외)에 Unload 이벤트를 발생시킴
'   반환 : True(종료), False(종료중지)
'-----------------------------------------------------------------------------'
Private Function AppExit() As Boolean
    Dim frmTemp As Form
    Dim intTemp As Integer
    Dim i       As Long
    
    intTemp = MsgBox("프로그램을 종료하시겠습니까?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then
        AppExit = False
        Exit Function
    End If
    
    '## 모든폼에 Unload 이벤트를 발생시킴
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> Me.Name Then
            Unload Forms(i)
        End If
    Next i
    
    '## 장비Dll폼 Unload
    Set mEqpShow = Nothing
    
    AppExit = True

End Function

'-----------------------------------------------------------------------------'
'   기능 : 실행중인 이전 원도우를 활성화 시킴
'   인수 :
'       - pWindowTitle : Window Title
'-----------------------------------------------------------------------------'
Public Sub RestoreWindow(ByVal pWindowTitle As String)
   Dim hWndCtlApp   As Long
   Dim PrevHndl     As Long
    
   hWndCtlApp = FindWindow(vbNullString, pWindowTitle)
   
   If hWndCtlApp Then
        PrevHndl = GetWindow(hWndCtlApp, GW_HWNDPREV)
        If PrevHndl Then
            ShowWindow PrevHndl, SW_RESTORE
            SetForegroundWindow PrevHndl
        End If
        SetForegroundWindow hWndCtlApp
   End If
End Sub
