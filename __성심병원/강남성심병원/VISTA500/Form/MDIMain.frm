VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Dimension VISTA500"
   ClientHeight    =   10530
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15435
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows 기본값
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   2400
      Top             =   1110
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMain 
      Align           =   2  '아래 맞춤
      BorderStyle     =   0  '없음
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   15435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   10155
      Width           =   15435
      Begin HSCotrol.CoolPgBar pgbMain 
         Height          =   360
         Left            =   3075
         TabIndex        =   2
         Top             =   735
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   635
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbMain 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   30
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   21669
               Key             =   "Output"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               TextSave        =   "오전 8:54"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlStatusbar 
      Left            =   1365
      Top             =   1065
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlCoolbar 
      Left            =   225
      Top             =   1065
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   810
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":08CA
            Key             =   "Close"
            Object.Tag             =   "종료"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":11A6
            Key             =   "Logo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":29FA
            Key             =   "Setting"
            Object.Tag             =   "장비설정"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":32D6
            Key             =   "Interface"
            Object.Tag             =   "검사"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":3BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":448E
            Key             =   "Testitem"
            Object.Tag             =   "검사항목"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5646
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":16756
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":16A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1734E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":17C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":18506
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "☞파일(&F)"
      Begin VB.Menu mnuLogin 
         Caption         =   "로그 인 .."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "프린터 설정(&U)..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu mnuWork 
      Caption         =   "작업(&W)"
      Visible         =   0   'False
      Begin VB.Menu mnuOrder 
         Caption         =   "처방(&O)"
      End
      Begin VB.Menu mnuResult 
         Caption         =   "결과 입력(&R)"
      End
      Begin VB.Menu mnuWorkBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRelultPrint 
         Caption         =   "결과 출력(&P)"
      End
      Begin VB.Menu mnuWrokBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSerch 
         Caption         =   "결과 조회"
      End
      Begin VB.Menu mnuStatistics 
         Caption         =   "검사 통계"
      End
   End
   Begin VB.Menu mnuInterface 
      Caption         =   "☞Interface(&I)"
      Begin VB.Menu mnuInterfaceAdvia 
         Caption         =   "Interface ADVIA120"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLink 
         Caption         =   "Communications Link"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "☞설정(&S)"
      Begin VB.Menu mnuTest 
         Caption         =   "검사항목(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetupBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComm 
         Caption         =   "장비 통신(&C)"
      End
      Begin VB.Menu mnuEqpTest 
         Caption         =   "장비 검사항목"
      End
      Begin VB.Menu mnuSetupBar2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRegUser 
         Caption         =   "사용자 등록"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "☞창(&W)"
      Visible         =   0   'False
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "계단식 배열(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "수평 바둑판식 배열(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "수직 바둑판식 배열(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "아이콘 정렬(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "☞도움말(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "목차(&C)"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "찾기(&S)..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "응용 프로그램 정보(&A)..."
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Main From Scale
Private Const MAIN_TOP      As Long = -60
Private Const MAIN_LEFT     As Long = -60
'Private Const MAIN_WIDTH    As Long = 12120
'Private Const MAIN_HEGHT    As Long = 8700
Private Const MAIN_WIDTH    As Long = 15480
Private Const MAIN_HEGHT    As Long = 11400

Private Const TLBKEY_ORDER      As String = "ORDER"
Private Const TLBKEY_RESULT     As String = "RESULT"
Private Const TLBKEY_PRINT      As String = "PRINT"
Private Const TLBKEY_INTERFACE  As String = "INTERFACE"
Private Const TLBKEY_TESTITEM   As String = "TESTITEM"
Private Const TLBKEY_SETTING    As String = "SETTING"
Private Const TLBKEY_LOGIN      As String = "LOGIN"
Private Const TLBKEY_EXIT       As String = "EXIT"
Private Const TLBKEY_USER       As String = "USER"

Private CallForm                As String

Private Sub BHImageButton1_Click()
    frmAbout.Show
End Sub

Private Sub MDIForm_Resize()
    picMain.Height = stbMain.Height
    stbMain.Move 0, 0, picMain.Width, picMain.Height
    pgbMain.Move stbMain.Panels("Output").left + 30, 50, stbMain.Panels("Output").Width - 60, stbMain.Height - 80
End Sub

Private Sub mnuComm_Click()
    Call Load_From(frmComSetup)
End Sub

Private Sub mnuEqpTest_Click()
    Call Load_From(frmTestEqp)
End Sub

Private Sub mnuHelpAbout_Click()
'    Call ShellAbout(hWnd, "LIMAS Emergency System.", _
'                    "LIS Emergency Program" & vbCrLf & "MD SAVER CO LTD", Icon)
    frmAbout.Show vbModal, Me
    
End Sub

Private Sub mnuHelpContents_Click()
'  On Error Resume Next
'  Dim nRet As Integer
'  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'  If Err Then
'    MsgBox Err.Description
'  End If
End Sub

Private Sub mnuHelpSearch_Click()
'  On Error Resume Next
'
'  Dim nRet As Integer
'  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'  If Err Then
'    MsgBox Err.Description
'  End If
End Sub

Private Sub mnuInterfaceAdvia_Click()
'    Call Load_From(frmInterface)
End Sub

Private Sub mnuLink_Click()
    Call Load_From(frmComm)
End Sub

Private Sub mnuLogin_Click()

    Call frmLogin.Show(vbModal, Me)
    
End Sub

Private Sub mnuOrder_Click()
'    Call Load_From(frmOrder)

End Sub

Private Sub mnuRegUser_Click()
    Call Load_From(frmRegUser) '(vbModeless, Me)
End Sub

Private Sub mnuRelultPrint_Click()
'    Call Load_From(frmPrint)
End Sub

Private Sub mnuResult_Click()
'    Call Load_From(frmResult)
End Sub

Private Sub mnuSerch_Click()
'    Call Load_From(frmSerch)
End Sub

Private Sub mnuStatistics_Click()
    Call Load_From(frmStatistics)
End Sub

Private Sub mnuTest_Click()
'    Call Load_From(frmTestItem)
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Call Arrange(vbArrangeIcons)
End Sub

Private Sub mnuWindowTileVertical_Click()
    Call Arrange(vbTileVertical)
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Call Arrange(vbTileHorizontal)
End Sub

Private Sub mnuWindowCascade_Click()
    Call Arrange(vbCascade)
End Sub

Private Sub mnuFileExit_Click()
    Dim rv As Long
    
    If frmComm.comEQP.PortOpen = True Then
        If MsgBox("인터페이스중입니다." & Chr(10) & _
               "작업을 종료하면 받고있거나 검사중인 데이터를 잃게 됩니다" & Chr(10) & _
               "종료하시겠습니까?", vbCritical + vbYesNo, App.Title) = vbYes Then
            End
        End If
    Else
        End
    End If
End Sub

Private Sub mnuFilePrintSetup_Click()
    CallForm = "Private Sub mnuFilePrintSetup_Click()"
    
On Error GoTo ErrorRoutine
     With cdlMain
        .CancelError = False
        .Copies = 1
        .DialogTitle = "Printer Setup"
        
        .Orientation = cdlPortrait '= cdlLandscape
        .PrinterDefault = True
        .flags = cdlPDPrintSetup
        .ShowPrinter
    End With
Exit Sub

ErrorRoutine:
    If Err.Number = cdlCancel Then Exit Sub
    Call ErrMsgProc(CallForm)
End Sub

Private Sub MDIForm_Load()
    Dim tmpSaveDt As String
'    MsgBox Chr(32)
    CallForm = "Private Sub MDIForm_Load()"

On Error GoTo ErrorRouten

'    pnlInstrument.Caption = "Leaders Lab"
            
    INS_CODE = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPCODE)
    INS_NAME = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPNAME)
    Caption = Me.Caption
    
    Call MDIForm_Lock
    Call MDIForm_Tool
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strSaveDt   As String, intCnt   As Integer
    
    sqlDoc = "Select SAVE_DT from INTERFACE001 where EQP_CD = '" & INS_CODE & "'"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    If Not adoRS.EOF Then strSaveDt = Trim$(adoRS(0) & "")
    adoRS.Close:    Set adoRS = Nothing

    tmpSaveDt = Format(DateDiff("d", Val(strSaveDt), Format(Now, "yyyy-mm-dd")), "yyyymmdd")
    
    sqlDoc = "Select count(*) from INTERFACE003 where TransDt <= '" & tmpSaveDt & "'"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    If Not adoRS.EOF Then intCnt = adoRS(0) & ""
    adoRS.Close:    Set adoRS = Nothing
    
    If intCnt > 0 Then
        If MsgBox(strSaveDt + "일 지난 데이타를 삭제하시겠습니까?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(strSaveDt), Format(Now, "YYYY-MM-DD")), "YYYYMMDD")
            
            sqlDoc = "delete from INTERFACE003 where TransDt <= '" & strSaveDt & "'"
            AdoCn_Jet.Execute sqlDoc
        End If
    End If
    
    Call Load_From(frmComm)

Exit Sub

ErrorRouten:
    Call ErrMsgProc(CallForm)
    Resume Next
    
End Sub

Private Sub MDIForm_Lock()
    
    Dim hMenu       As Long
    Dim lngStyle    As Long
    
    CallForm = "Private Sub MDIForm_Lock()"
On Error GoTo ErrorRouten
    
    Call LockWindowUpdate(hwnd)                             'Lock window
    hMenu = GetSystemMenu(hwnd, 0)
    Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)           'Size menu Diseable
    Call DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)       'Max menu Diseable
    lngStyle = GetWindowLong(hwnd, GWL_STYLE)
    lngStyle = lngStyle And (Not (CLng(WS_MAXIMIZEBOX))) And (Not (CLng(WS_THICKFRAME))) 'Max Button Diseable
    lngStyle = lngStyle Or CLng(WS_BORDER)                                               'Size Cursour Diseable
    Call SetWindowLong(hwnd, GWL_STYLE, lngStyle)

    If (GetSystemMetrics(SM_CYSCREEN) <= 600) And (GetSystemMetrics(SM_CXSCREEN) <= 800) Then
        Call Move(MAIN_LEFT, MAIN_TOP, MAIN_WIDTH, MAIN_HEGHT)          'Size Lock
    Else
        Call Move((Screen.Width - MAIN_WIDTH) / 2, (Screen.Height - MAIN_HEGHT) / 2, MAIN_WIDTH, MAIN_HEGHT)
    End If
    Call LockWindowUpdate(0&)
Exit Sub

ErrorRouten:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub MDIForm_Tool()
    
    CallForm = "Private Sub MDIForm_Tool()"
    
On Error GoTo ErrorRouten
    With tlbMain
        .AllowCustomize = False
        Set .ImageList = imlToolbar
        .TextAlignment = tbrTextAlignBottom '= tbrTextAlignRight
        .BorderStyle = ccNone
        .Appearance = cc3D
        .Style = tbrFlat
'        Call .Buttons.Add(, TLBKEY_LOGIN, "로그인", tbrDefault, "Logon")
        Call .Buttons.Add(, "", "", tbrSeparator)
        Call .Buttons.Add(, "", "", tbrSeparator)
'        .Buttons.Add 3, TLBKEY_ORDER, "처   방", tbrDefault, "Order"
'        .Buttons.Add 4, TLBKEY_RESULT, "결과입력", tbrDefault, "Result"
'        .Buttons.Add 5, TLBKEY_PRINT, "결과출력", tbrDefault, "Print"
        Call .Buttons.Add(, TLBKEY_INTERFACE, "", tbrDefault, "Interface")
        Call .Buttons.Add(, TLBKEY_TESTITEM, "", tbrDefault, "Testitem")
        Call .Buttons.Add(, TLBKEY_SETTING, "", tbrDefault, "Setting")
        Call .Buttons.Add(, TLBKEY_USER, "", tbrDefault, "User")
        Call .Buttons.Add(, "", "", tbrSeparator)
        Call .Buttons.Add(, TLBKEY_EXIT, "", tbrDefault, "Close")
        .Refresh
    End With

'    With clbMain
'        Set .ImageList = imlCoolbar
'        With .Bands(1)
'            Set .Child = tlbMain
'            .MinHeight = tlbMain.Height
'        End With
''        With .Bands(2)
''            .Image = "Logo"
''            .MinWidth = 0
''            .MinHeight = tlbMain.Height
''            .Visible = True
''        End With
'        .FixedOrder = False
'        .BandBorders = False
'        .Height = tlbMain.Height
'        .Refresh
'    End With
        
    With stbMain
        .Enabled = False
        .Panels(1).text = CurrUser.CuUserNM
    End With
    
    With pgbMain
        .ForeColor = &H8000000D
    End With
Exit Sub

ErrorRouten:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call mnuFileExit_Click
    
End Sub

Private Sub pnlInstrument_Click()
On Error GoTo ErrorProc:
    Dim Domain As String
    Dim s As Double

    Domain = "http://dglog.anyhelp.net/"

    ShellExecute 0, vbNullString, Domain, vbNullString, vbNullString, 1


    Exit Sub
ErrorProc:
    MsgBox Err.Description
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim X As Long
    Dim Y As Long
    
    X = tlbMain.left + Button.left
'    y = clbMain.Bands(1).Height + tlbMain.Top
    Select Case Button.Key
        Case TLBKEY_ORDER
            Call mnuOrder_Click
        Case TLBKEY_RESULT
            Call mnuResult_Click
        Case TLBKEY_PRINT
            Call mnuRelultPrint_Click
        Case TLBKEY_INTERFACE
            Call mnuLink_Click
        Case TLBKEY_LOGIN
            Call mnuLogin_Click
        Case TLBKEY_EXIT
            Call mnuFileExit_Click
        Case TLBKEY_TESTITEM
            Call mnuEqpTest_Click
        Case TLBKEY_SETTING
            Call mnuComm_Click
        Case TLBKEY_USER
            Call mnuRegUser_Click
    End Select
    
End Sub

Private Sub tlbMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tlbMain_ButtonClick(Button)
End Sub

Private Sub Load_From(ByVal frm As Form)
    
    Dim hMenu       As Long
    Dim lngStyle    As Long
    
    With frm
        .Show
        .SetFocus
        
'        hMenu = GetSystemMenu(.hwnd, 0)
'        Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)           'Size menu Diseable
'        Call DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)       'Max menu Diseable
'        lngStyle = GetWindowLong(.hwnd, GWL_STYLE)
'        lngStyle = lngStyle And (Not (CLng(WS_MAXIMIZEBOX))) And (Not (CLng(WS_THICKFRAME))) 'Max Button Diseable
'        lngStyle = lngStyle Or CLng(WS_BORDER)                                               'Size Cursour Diseable
'        Call SetWindowLong(.hwnd, GWL_STYLE, lngStyle)
    End With
    
End Sub
