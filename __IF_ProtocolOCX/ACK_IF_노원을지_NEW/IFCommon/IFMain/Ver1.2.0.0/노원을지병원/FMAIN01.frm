VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FMAIN01 
   Caption         =   "   ACK �������̽� ��Ʈ�ѷ� - (ACKICON)"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "FMAIN01.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   6795
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtUID 
      Height          =   270
      Left            =   0
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtUserNm 
      Height          =   300
      Left            =   1200
      TabIndex        =   11
      Top             =   210
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtUserOther 
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPWD 
      Height          =   270
      Left            =   600
      TabIndex        =   9
      Top             =   210
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   60
      Picture         =   "FMAIN01.frx":27A2
      ScaleHeight     =   55
      ScaleMode       =   3  '�ȼ�
      ScaleWidth      =   129
      TabIndex        =   8
      Top             =   90
      Width           =   1965
   End
   Begin Threed.SSCommand cmdIF 
      Height          =   645
      Left            =   2130
      TabIndex        =   0
      ToolTipText     =   "���� �������̽�"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":338A
   End
   Begin VB.ComboBox cboMList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FMAIN01.frx":36A4
      Left            =   2130
      List            =   "FMAIN01.frx":36A6
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   6
      Top             =   30
      Width           =   4650
   End
   Begin Threed.SSCommand cmdRstSrch 
      Height          =   645
      Left            =   2790
      TabIndex        =   1
      ToolTipText     =   "�����ȸ �� ���"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":36A8
   End
   Begin Threed.SSCommand cmdProgCfg 
      Height          =   645
      Left            =   4770
      TabIndex        =   4
      ToolTipText     =   "ȯ�漳��"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":39C2
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   645
      Left            =   6090
      TabIndex        =   7
      ToolTipText     =   "�� ��"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":3CDC
   End
   Begin Threed.SSCommand cmdTestCfg 
      Height          =   645
      Left            =   4110
      TabIndex        =   3
      ToolTipText     =   "�˻��׸� ����"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":3FF6
   End
   Begin Threed.SSCommand cmdDelCfg 
      Height          =   645
      Left            =   5430
      TabIndex        =   5
      ToolTipText     =   "���� ������ �����Ⱓ ����"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":4310
   End
   Begin Threed.SSCommand cmdStatistics 
      Height          =   645
      Left            =   3450
      TabIndex        =   2
      ToolTipText     =   "������� �� �缺�� ��ȸ"
      Top             =   360
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1138
      _StockProps     =   78
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      Picture         =   "FMAIN01.frx":462A
   End
End
Attribute VB_Name = "FMAIN01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iMAX_MACHINE  As Integer
Private msUserInfoYN  As String
Public msVerUserInfo As String
Private msAutoDelYN   As String
Private msAutoIFYN  As String
Private msOSInfo    As String       '2003/1/22 �߰�(yk)
Private Sub DelPreviousData()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim dateTmp As Date
    Dim objLD As Object
    Dim i%
    
    If UCase(msAutoDelYN) = "Y" Then
        For i = 1 To iMAX_MACHINE
            sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachCd(i), "Delete.Interval")
                
            dateTmp = Format(Now - Val(sBuf) - 1, "YYYY-MM-DD")
            
            Set objLD = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachCd(i)), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachCd(i)))
            
            If objLD.Del_IFResult(gsMachCd(i), 2, Format(dateTmp, "YYYYMMDD"), "") = True Then
                ViewMsg "������ ���� �����Ͱ� �����Ǿ����ϴ�!!"""
            End If
            
            Set objLD = Nothing
        Next
    Else
        If MsgBox("������ ���� �����͸� �����Ͻðڽ��ϱ�?", vbYesNo, "���� ������ ���� ����") = vbYes Then
            For i = 1 To iMAX_MACHINE
                sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "Delete.Interval")
                    
                dateTmp = Format(Now - Val(sBuf) - 1, "YYYY-MM-DD")
                
                Set objLD = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachCd(i)), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachCd(i)))
                
                If objLD.Del_IFResult(gsMachCd(i), 2, Format(dateTmp, "YYYYMMDD"), "") = True Then
                    ViewMsg "������ ���� �����Ͱ� �����Ǿ����ϴ�!!"""
                End If
                
                Set objLD = Nothing
            Next
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "DelPreviousData Err - (" & Err.Description & ")"
End Sub

Private Sub GetUserInfo()
    On Error GoTo ErrHandler
    
    Load FLOGIN01
    FLOGIN01.Show vbModal
    
    Exit Sub
    
ErrHandler:
End Sub

Private Sub GetMainini()
    Dim sBuf$
    Dim retval As Long
    Dim i%
    
    'User Information ��뿩��
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("Config", "UserInformation", "N", sBuf, 255, App.Path & "\MainCfg.ini")
    
    If retval = 0 Then
    Else
        msUserInfoYN = LeftH(sBuf, retval)
    End If
    
    'User Information ��뿩��
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("Config", "VersionOfUserInformation", "0100", sBuf, 255, App.Path & "\MainCfg.ini")
    
    If retval = 0 Then
    Else
        msVerUserInfo = LeftH(sBuf, retval)
    End If
    
    'AutoDeleteOfLocalData ��뿩��
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("Config", "AutoDeleteOfLocalData", "Y", sBuf, 255, App.Path & "\MainCfg.ini")
    
    If retval = 0 Then
    Else
        msAutoDelYN = LeftH(sBuf, retval)
    End If
    
    'AutoStartOfInterfaceObject ��뿩��
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("Config", "AutoStartOfInterfaceObject", "Y", sBuf, 255, App.Path & "\MainCfg.ini")
    
    If retval = 0 Then
    Else
        msAutoIFYN = LeftH(sBuf, retval)
    End If


    'OS�� ���� Form ������ ����
    On Error GoTo ErrRtn
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("Config", "OSVersionInformation", "Y", sBuf, 255, App.Path & "\MainCfg.ini")
    
    If retval = 0 Then
    Else
        msOSInfo = LeftH(sBuf, retval)
    End If
    On Error GoTo 0
ErrRtn:

End Sub

Private Sub GetIFini()
    Dim sBuf$
    Dim retval As Long
    Dim i%
    
    For i = 1 To 100
    'Machine Code
        sBuf = String(255, 0)
        retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineCd", "", sBuf, 255, App.Path & "\����ڵ�" & CStr(i) & ".ini")
        
        If retval = 0 Then
        Else
            iMAX_MACHINE = iMAX_MACHINE + 1
            
            ReDim Preserve gsMachCd(iMAX_MACHINE)
            ReDim Preserve gsMachNm(iMAX_MACHINE)
            
            gsMachCd(iMAX_MACHINE) = LeftH(sBuf, retval)
            
            'Machine Name
            sBuf = String(255, 0)
            retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineNm", "", sBuf, 255, App.Path & "\����ڵ�" & CStr(i) & ".ini")
            
            If retval = 0 Then
                MsgBox "����ڵ� ������ �Ǿ� ���� �ʽ��ϴ�. ���α׷��� ����˴ϴ�!!", vbCritical, "����ڵ�" & CStr(i) & ".ini ����"
                End
            End If
            
            gsMachNm(iMAX_MACHINE) = LeftH(sBuf, retval)
            cboMList.AddItem LeftH(sBuf, retval)
        End If
    Next
End Sub

Private Sub RegEditIFini()
    Dim bRetVal As Boolean
    Dim i%
    
    '<-------------------------------------------------------------------------------------->
    For i = 1 To iMAX_MACHINE
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "MachineNm", gsMachNm(i))
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    Next
End Sub

Private Sub InitRegWndTitle()
    Dim i%
    Dim bRetVal As Boolean
    
    For i = 1 To iMAX_MACHINE
        '<-------------------------------------------------------------------------------------->
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "WndTitle.IF", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        '<-------------------------------------------------------------------------------------->
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "WndTitle.RstSrch", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        '<-------------------------------------------------------------------------------------->
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "WndTitle.TestCfg", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        '<-------------------------------------------------------------------------------------->
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "WndTitle.ProgCfg", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        '<-------------------------------------------------------------------------------------->
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(i), "WndTitle.DelCfg", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    Next
End Sub

Private Sub cmdDelCfg_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objDC As Object
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.DelCfg")
    
    Me.MousePointer = vbHourglass
    
    Set objDC = CreateObject("FDC" & Left(fCurVerObject("DelCfg", gsMachCd(cboMList.ListIndex + 1)), 2) & ".FCDC" & fCurVerObject("DelCfg", gsMachCd(cboMList.ListIndex + 1)))
        
    If sBuf = "" Then
        objDC.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objDC.init
        
    Else
        MsgBox "[" & cmdDelCfg.ToolTipText & "] â�� �̹� ���� ���Դϴ�!!"
        
        objDC.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objDC.ShowForm
    End If
    
    Set objDC = Nothing
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
    Set objDC = Nothing
    MsgBox cmdDelCfg.ToolTipText & " ���� ���� - " & Err.Description
End Sub

Private Sub cmdExit_Click()
    Dim sBuf$
    Dim objRS As Object
    
    If MsgBox("ACK �������̽� ��Ʈ�ѷ��� �����ϸ� ��� Interface �۾��� ����˴ϴ�." & vbCrLf & _
              "����Ͻðڽ��ϱ�?" & vbCrLf & vbCrLf & _
              "Interface �۾� ���߿� ������ ��� ���۵����Ͱ� �ս��� �˴ϴ�.", vbQuestion + vbYesNo, _
              "ACK �������̽� ��Ʈ�ѷ� ���� Ȯ��") = vbYes Then
         
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.RstSrch")
    
        gsMachineCd = gsMachCd(cboMList.ListIndex + 1)
        
        If sBuf <> "" Then
            Set objRS = CreateObject("FRS" & Left(fCurVerObject("RstSrch", gsMachCd(cboMList.ListIndex + 1)), 2) & ".FCRS" & fCurVerObject("RstSrch", gsMachCd(cboMList.ListIndex + 1)))
            objRS.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
            objRS.Term
        End If
        
        Set objRS = Nothing
         
        End
    End If
End Sub

Private Sub cmdIF_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objIF As Object
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.IF")
    
    gsMachineCd = gsMachCd(cboMList.ListIndex + 1)
    
    Me.MousePointer = vbHourglass
    
    Set objIF = CreateObject("FIF" & Left(fCurVerObject("IF", gsMachCd(cboMList.ListIndex + 1)), 2) & gsMachCd(cboMList.ListIndex + 1) & ".FCIF" & fCurVerObject("IF", gsMachCd(cboMList.ListIndex + 1)))
    
    If sBuf = "" Then
        objIF.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objIF.init Trim(Me.txtUID), Trim(Me.txtPWD), Trim(Me.txtUserNm), Trim(Me.txtUserOther)
        
    Else
        MsgBox "[" & cmdIF.ToolTipText & "] â�� �̹� ���� ���Դϴ�!!"
        
        objIF.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objIF.ShowForm
    End If
    
    Set objIF = Nothing
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
    Set objIF = Nothing
    MsgBox cmdIF.ToolTipText & " ���� ���� - " & Err.Description
End Sub

Private Sub cmdProgCfg_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objPC As Object
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.ProgCfg")

    gsMachineCd = gsMachCd(cboMList.ListIndex + 1)
    
    Me.MousePointer = vbHourglass
    
    Set objPC = CreateObject("FPC" & Left(fCurVerObject("ProgCfg", gsMachCd(cboMList.ListIndex + 1)), 2) & ".FCPC" & fCurVerObject("ProgCfg", gsMachCd(cboMList.ListIndex + 1)))

    If sBuf = "" Then
        objPC.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objPC.init
    Else
        MsgBox "[" & cmdProgCfg.ToolTipText & "] â�� �̹� ���� ���Դϴ�!!"
        
        objPC.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objPC.ShowForm
    End If
    
    Set objPC = Nothing
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
    Set objPC = Nothing
    MsgBox cmdProgCfg.ToolTipText & " ���� ���� - " & Err.Description
End Sub

Private Sub cmdRstSrch_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objRS As Object
        
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.RstSrch")

    gsMachineCd = gsMachCd(cboMList.ListIndex + 1)
    
    Me.MousePointer = vbHourglass
    
    Set objRS = CreateObject("FRS" & Left(fCurVerObject("RstSrch", gsMachCd(cboMList.ListIndex + 1)), 2) & ".FCRS" & fCurVerObject("RstSrch", gsMachCd(cboMList.ListIndex + 1)))
    
    If sBuf = "" Then
        objRS.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objRS.init Trim(Me.txtUID), Trim(Me.txtPWD), Trim(Me.txtUserNm), Trim(Me.txtUserOther)
    Else
        MsgBox "[" & cmdRstSrch.ToolTipText & "] â�� �̹� ���� ���Դϴ�!!"
        
        objRS.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objRS.ShowForm
    End If
    
    Set objRS = Nothing
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
    Set objRS = Nothing
    MsgBox cmdRstSrch.ToolTipText & " ���� ���� - " & Err.Description
End Sub

Private Sub cmdStatistics_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objSS As Object
        
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.Statistics")

    gsMachineCd = gsMachCd(cboMList.ListIndex + 1)
    
    Set objSS = CreateObject("FST" & Left(fCurVerObject("Statistics", gsMachCd(cboMList.ListIndex + 1)), 2) & ".FCST" & fCurVerObject("Statistics", gsMachCd(cboMList.ListIndex + 1)))
    
    If sBuf = "" Then
        objSS.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objSS.init
    Else
        MsgBox "[" & cmdStatistics.ToolTipText & "] â�� �̹� ���� ���Դϴ�!!"
        
        objSS.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objSS.ShowForm
    End If
    
    Set objSS = Nothing
    
    Exit Sub
    
ErrHandler:
    Set objSS = Nothing
    MsgBox cmdStatistics.ToolTipText & " ���� ���� - " & Err.Description
End Sub

Private Sub cmdTestCfg_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objTC As Object
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachCd(cboMList.ListIndex + 1), "WndTitle.TestCfg")

    gsMachineCd = gsMachCd(cboMList.ListIndex + 1)
    
    Me.MousePointer = vbHourglass
    
    Set objTC = CreateObject("FTC" & Left(fCurVerObject("TestCfg", gsMachCd(cboMList.ListIndex + 1)), 2) & ".FCTC" & fCurVerObject("TestCfg", gsMachCd(cboMList.ListIndex + 1)))

    If sBuf = "" Then
        objTC.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objTC.init
    Else
        MsgBox "[" & cmdTestCfg.ToolTipText & "] â�� �̹� ���� ���Դϴ�!!"
        objTC.SetMachineInfo gsMachCd(cboMList.ListIndex + 1), gsMachNm(cboMList.ListIndex + 1)
        objTC.ShowForm
    End If
    
    Set objTC = Nothing
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
    Set objTC = Nothing
    MsgBox cmdTestCfg.ToolTipText & " ���� ���� - " & Err.Description
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Call GetMainini
    Call GetIFini
    Call RegEditIFini
    Call InitRegWndTitle
        
    If UCase(msUserInfoYN) = "Y" Then
        Call GetUserInfo
    End If
    
    Call DelPreviousData
    
    cboMList.ListIndex = 0
    
    If UCase(msAutoIFYN) = "Y" Then
        cmdIF.DoClick
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("ACK �������̽� ��Ʈ�ѷ��� �����ϸ� ��� Interface �۾��� ����˴ϴ�." & vbCrLf & _
              "����Ͻðڽ��ϱ�?" & vbCrLf & vbCrLf & _
              "Interface �۾� ���߿� ������ ��� ���۵����Ͱ� �ս��� �˴ϴ�.", vbYesNo, _
              "ACK �������̽� ��Ʈ�ѷ� ���� Ȯ��") = vbYes Then
            
        Unload Me
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrHandler
    
    If msOSInfo = "XP" Then
        Me.Height = 1530    'WIN XP
    Else
        Me.Height = 1430    'WIN 98
    End If
    Me.Width = 6915     '6270
    
    Exit Sub
    
ErrHandler:
End Sub
