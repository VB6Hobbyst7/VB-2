VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Axsym"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.PictureBox picMain 
      Align           =   2  '�Ʒ� ����
      BorderStyle     =   0  '����
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   15240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10455
      Width           =   15240
      Begin MSComctlLib.StatusBar stbMain 
         Height          =   330
         Left            =   15
         TabIndex        =   2
         Top             =   30
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   14843
               Key             =   "Output"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               TextSave        =   "���� 11:19"
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   2370
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlStatusbar 
      Left            =   1470
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   870
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":08CA
            Key             =   "Order"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1B4C
            Key             =   "Logon"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2DCE
            Key             =   "Result"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4050
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":52D2
            Key             =   "Interface"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":6554
            Key             =   "Setting"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":77D6
            Key             =   "Testitem"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8A58
            Key             =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCoolbar 
      Left            =   330
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '�� ����
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   979
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "������ ����(&U)..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "������(&X)"
      End
   End
   Begin VB.Menu mnuInterface 
      Caption         =   "�˻�(&I)"
      Begin VB.Menu mnuInterfaceAxsym 
         Caption         =   "�������̽�"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResult 
         Caption         =   "�����ȸ�׵��"
      End
   End
   Begin VB.Menu mnuWork 
      Caption         =   "���"
      Begin VB.Menu mnuStatistics 
         Caption         =   "�˻� ���"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "����(&S)"
      Begin VB.Menu mnuTest 
         Caption         =   "�˻��׸�(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetupBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComm 
         Caption         =   "��� ���(&C)"
      End
      Begin VB.Menu mnuEqpTest 
         Caption         =   "��� �˻��׸�"
      End
      Begin VB.Menu mnuSetupBar2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRegUser 
         Caption         =   "����� ���"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "â(&W)"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "��ܽ� �迭(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "���� �ٵ��ǽ� �迭(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "���� �ٵ��ǽ� �迭(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "������ ����(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "ã��(&S)..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "���� ���α׷� ����(&A)..."
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
Private Const MAIN_WIDTH    As Long = 12120
Private Const MAIN_HEGHT    As Long = 8700
'Private Const MAIN_WIDTH    As Long = 15480
'Private Const MAIN_HEGHT    As Long = 11220

Private Const TLBKEY_ORDER      As String = "ORDER"
Private Const TLBKEY_RESULT     As String = "RESULT"
Private Const TLBKEY_PRINT      As String = "PRINT"
Private Const TLBKEY_INTERFACE  As String = "INTERFACE"
Private Const TLBKEY_TESTITEM   As String = "TESTITEM"
Private Const TLBKEY_SETTING    As String = "SETTING"
Private Const TLBKEY_LOGIN      As String = "LOGIN"
Private Const TLBKEY_EXIT       As String = "EXIT"

Private CallForm                As String

Private Sub MDIForm_Resize()
    picMain.Height = stbMain.Height
    stbMain.Move 0, 0, picMain.Width, picMain.Height
    'pgbMain.Move stbMain.Panels("Output").left + 30, 50, stbMain.Panels("Output").Width - 60, stbMain.Height - 80
End Sub

Private Sub mnuComm_Click()
    Call Load_From(frmComSetup)
End Sub

Private Sub mnuEqpTest_Click()
    Call Load_From(frmTestEqp)
End Sub

'Private Sub mnuHelpAbout_Click()
'    Call ShellAbout(hWnd, "LIMAS Emergency System.", _
'                    "LIS Emergency Program" & vbCrLf & "MD SAVER CO LTD", Icon)
'    frmAbout.Show vbModal, Me
    
'End Sub

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

Private Sub mnuInterfaceAxsym_Click()
    Call Load_From(frmComm)
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
'    Call Load_From(frmRegUser) '(vbModeless, Me)
End Sub

Private Sub mnuRelultPrint_Click()
    Call Load_From(frmStatistics)
End Sub

Private Sub mnuResult_Click()
    Call Load_From(frmResult)
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
    
    If frmComm.comEQP.PortOpen <> False Then
        If MsgBox("�������̽����Դϴ�." & Chr(10) & _
               "�۾��� �����ϸ� �ް��ְų� �˻����� �����͸� �Ұ� �˴ϴ�" & Chr(10) & _
               "�����Ͻðڽ��ϱ�?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
               
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
    
    CallForm = "Private Sub MDIForm_Load()"

On Error GoTo ErrorRouten
            
    INS_CODE = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPCODE)
    INS_NAME = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPNAME)
    Caption = "LAB(" & Me.Caption & ")"
    
    Call MDIForm_Lock
    Call MDIForm_Tool
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strSaveDt   As String, intCnt   As Integer
    
    sqlDoc = "select SAVE_DT from INTERFACE001 where EQP_CD = '" & INS_CODE & "'"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    If Not adoRS.EOF Then strSaveDt = Trim$(adoRS(0) & "")
    adoRS.Close:    Set adoRS = Nothing

    tmpSaveDt = Format(DateDiff("d", Val(strSaveDt), Format(Now, "yyyy-mm-dd")), "yyyymmdd")
    
    sqlDoc = "select count(*) from INTERFACE003 where TransDt <= '" & tmpSaveDt & "'"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    If Not adoRS.EOF Then intCnt = adoRS(0) & ""
    adoRS.Close:    Set adoRS = Nothing
    
    If intCnt > 0 Then
        If MsgBox(strSaveDt + "�� ���� ����Ÿ�� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(strSaveDt), Format(Now, "YYYY-MM-DD")), "YYYYMMDD")
            
            sqlDoc = "delete from INTERFACE003 where TransDt <= '" & strSaveDt & "'"
            AdoCn_Jet.Execute sqlDoc
        End If
    End If
    
'    Call Load_From(frmComm) 'frmComm
    'Call Load_From(frmComm)

Exit Sub

ErrorRouten:
    Call ErrMsgProc(CallForm)
    Resume Next '-- YEJ Add
    
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
        'Call Move((Screen.Width - MAIN_WIDTH) / 2, (Screen.Height - MAIN_HEGHT) / 2, MAIN_WIDTH, MAIN_HEGHT)
        '-- 1024 * 768
        Call Move(0, 0, 15360, 11520)
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
        .Appearance = ccFlat
        .Style = tbrFlat
'        Call .Buttons.Add(, TLBKEY_LOGIN, "�α���", tbrDefault, "Logon")
        Call .Buttons.Add(, "", "", tbrSeparator)
'        .Buttons.Add 3, TLBKEY_ORDER, "ó   ��", tbrDefault, "Order"
'        .Buttons.Add 4, TLBKEY_RESULT, "����Է�", tbrDefault, "Result"
'        .Buttons.Add 5, TLBKEY_PRINT, "������", tbrDefault, "Print"
        Call .Buttons.Add(, TLBKEY_INTERFACE, "��  ��", tbrDefault, "Interface")
        Call .Buttons.Add(, TLBKEY_RESULT, "�����ȸ", tbrDefault, "Result")
        Call .Buttons.Add(, TLBKEY_PRINT, "������", tbrDefault, "Print")
        Call .Buttons.Add(, TLBKEY_TESTITEM, "�˻��׸�", tbrDefault, "Testitem")
        Call .Buttons.Add(, TLBKEY_SETTING, "��ż���", tbrDefault, "Setting")
        Call .Buttons.Add(, "", "", tbrSeparator)
        Call .Buttons.Add(, TLBKEY_EXIT, "��  ��", tbrDefault, "Close")
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
        .Panels(1).Text = CurrUser.CuUserNM
    End With
    
    'With pgbMain
    '    .ForeColor = &H8000000D
    'End With
Exit Sub

ErrorRouten:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call mnuFileExit_Click
    
End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim x As Long
    Dim y As Long
    
    x = tlbMain.left + Button.left
    'y = clbMain.Bands(1).Height + tlbMain.Top
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
    End Select
    
End Sub

Private Sub tlbMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tlbMain_ButtonClick(Button)
End Sub

Private Sub Load_From(ByVal frm As Form)
    
    Dim hMenu       As Long
    Dim lngStyle    As Long
    
    Gbl_FormName = frm.Name
    
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
