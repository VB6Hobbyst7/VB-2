VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "iSmart30"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15360
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   1440
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
      ScaleWidth      =   15360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   10155
      Width           =   15360
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
         Left            =   90
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
               TextSave        =   "오후 12:11"
            EndProperty
         EndProperty
      End
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
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   979
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin BHButton.BHImageButton BHImageButton1 
         Height          =   555
         Left            =   14220
         TabIndex        =   4
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIMain.frx":18DE2
         ImgOutLineSize  =   3
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "☞파일(&F)"
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "프린터 설정(&U)..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu mnuInterface 
      Caption         =   "☞Interface(&I)"
      Begin VB.Menu mnuLink 
         Caption         =   "Communications Link"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "☞설정(&S)"
      Begin VB.Menu mnuComm 
         Caption         =   "장비 통신(&C)"
      End
      Begin VB.Menu mnuEqpTest 
         Caption         =   "장비 검사항목"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "☞창(&W)"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "계단식 배열(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "수평 바둑판식 배열(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "수직 바둑판식 배열(&V)"
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

Private Const MAIN_WIDTH    As Long = 15480
Private Const MAIN_HEGHT    As Long = 11220

Private Const TLBKEY_PRINT      As String = "PRINT"
Private Const TLBKEY_INTERFACE  As String = "INTERFACE"
Private Const TLBKEY_TESTITEM   As String = "TESTITEM"
Private Const TLBKEY_SETTING    As String = "SETTING"
Private Const TLBKEY_EXIT       As String = "EXIT"

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
'    Call Load_From(frmComSetup)
    frmComSetup.Show vbModal
End Sub

Private Sub mnuEqpTest_Click()
    Call Load_From(frmTestEqp)
End Sub

Private Sub mnuLink_Click()
    Call Load_From(frmComm)
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
               "종료하시겠습니까?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
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
        .Flags = cdlPDPrintSetup
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
            
    INS_CODE = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPCODE)
    INS_NAME = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPNAME)
    Caption = Me.Caption & "[" & HOS_NAME & "]"
    
    Call MDIForm_Lock
    Call MDIForm_Tool
    Call Del_OldData    '기간 지난 결과값 삭제
    
    '로그파일 1M 이상이면 내용 삭제
    If GetFileSize(App.Path + "\" + "iSmart30.log") > 1000000 Then
        Open App.Path + "\" + "iSmart30.log" For Output As #1
        Close #1
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
    
Dim btnX As Button

On Error GoTo ErrorRouten
    With tlbMain
        .AllowCustomize = False
        Set .ImageList = imlToolbar
        .TextAlignment = tbrTextAlignBottom '= tbrTextAlignRight
        .BorderStyle = ccNone
        .Appearance = ccFlat
        .Style = tbrFlat
        Set btnX = .Buttons.Add(, "", "", tbrSeparator)
        Set btnX = .Buttons.Add(, TLBKEY_INTERFACE, "", tbrDefault, "Interface")
        btnX.ToolTipText = "검사"
        Set btnX = .Buttons.Add(, TLBKEY_TESTITEM, "", tbrDefault, "Testitem")
         btnX.ToolTipText = "장비검사항목"
        Set btnX = .Buttons.Add(, TLBKEY_SETTING, "", tbrDefault, "Setting")
         btnX.ToolTipText = "장비통신"
        Set btnX = .Buttons.Add(, "", "", tbrSeparator)
        Set btnX = .Buttons.Add(, TLBKEY_EXIT, "", tbrDefault, "Close")
         btnX.ToolTipText = "종료"
        .Refresh
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

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim X As Long
    Dim y As Long
    
    X = tlbMain.left + Button.left
    Select Case Button.Key
        Case TLBKEY_INTERFACE
            Call mnuLink_Click
        Case TLBKEY_EXIT
            Call mnuFileExit_Click
        Case TLBKEY_TESTITEM
            Call mnuEqpTest_Click
        Case TLBKEY_SETTING
            Call mnuComm_Click
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
