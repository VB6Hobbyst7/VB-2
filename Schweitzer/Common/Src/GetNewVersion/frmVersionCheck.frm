VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmVersionCheck 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FCEFE9&
   Caption         =   "���׷��̵� ��ƿ��Ƽ"
   ClientHeight    =   2490
   ClientLeft      =   1260
   ClientTop       =   2415
   ClientWidth     =   7005
   Icon            =   "frmVersionCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   7005
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4980
      Top             =   180
   End
   Begin MSComctlLib.ImageList imlNew 
      Left            =   5460
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":06EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":10F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":1B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":250E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":2F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":3926
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":4332
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":4D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":574A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":6156
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":6B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":756E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":7F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":8986
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVersionCheck.frx":9392
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FCEFE9&
      Height          =   530
      Left            =   180
      TabIndex        =   2
      Top             =   1140
      Width           =   6675
      Begin VB.Image Image1 
         Height          =   315
         Left            =   240
         Picture         =   "frmVersionCheck.frx":9D9E
         Stretch         =   -1  'True
         Top             =   135
         Width           =   255
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "#"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   600
         TabIndex        =   3
         Top             =   210
         Width           =   90
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FCEFE9&
      Height          =   720
      Left            =   180
      TabIndex        =   4
      Top             =   1620
      Width           =   6675
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Copy File"
         Height          =   180
         Left            =   4620
         TabIndex        =   10
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblCopyCount 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "0"
         Height          =   180
         Left            =   6015
         TabIndex        =   9
         Top             =   300
         Width           =   90
      End
      Begin VB.Label lblNewCount 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "0"
         Height          =   180
         Left            =   3795
         TabIndex        =   8
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "New File"
         Height          =   180
         Left            =   2520
         TabIndex        =   7
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblFileCount 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "0"
         Height          =   180
         Left            =   1755
         TabIndex        =   6
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Read File"
         Height          =   180
         Left            =   480
         TabIndex        =   5
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Image imgNew 
      Height          =   540
      Left            =   6180
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3405
      TabIndex        =   0
      Top             =   555
      Width           =   135
   End
End
Attribute VB_Name = "frmVersionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private FileCount As Integer
Private NewCount As Integer
Private LogName As String
Private Fd As Integer
Private blnNetCon As Boolean
Private download_ing As Boolean
Private ExeNames As String      '���� �������� ���α׷����� ���ڿ� ����

Const TH32CS_SNAPPROCESS As Long = 2&

Private Sub Form_Load()
    Dim i As Long
    ExeNames = ""
    'CLEAR...
    lblMessage = ""
    lblFile = ""
    prgBar.value = 0
    
    Me.Caption = Me.Caption & App.Major & "." & App.Minor
    Me.Show
    
'    Call medAlwaysOn(Me, 1)
    DoEvents
    
    blnNetCon = False
    blnDownloadMyself = False
    
    Call GetDir
    
    gID = medGetINI("DownLoad", "ID", "C:\Schweitzer\Schweitzer.ini")
    gPWD = medGetINI("DownLoad", "PWD", "C:\Schweitzer\Schweitzer.ini")
    
    Me.Show
    Me.ZOrder 0
    
'    lblMessage.Caption = "��� â�� �ּ�ȭ�ϰ� �ֽ��ϴ�."
'    Call MinimizeAllExcept(App.EXEName)
    
    Me.SetFocus
    
    '2���� �ڵ�����ǵ��� ����.... wooil
'    For i = 1 To 4
'        Call Sleep(500)
'        DoEvents
'    Next i
    If download_ing = False Then cmdDownload_Click
End Sub

Private Sub cmdDownload_Click()
    
    download_ing = True
'    cmdDownload.Enabled = False
    
    'Call MinimizeAllExcept("����üũ")
    Fd = FreeFile
    LogName = App.Path & "\Version.log"
    Open LogName For Output As #Fd
    
    If Dir(gNetDriveChar) <> "" Then Call ConNetDrive(1)  '��Ʈ�p ����̺� ����
    Call ConNetDrive(0)  '��Ʈ�p ����̺� ����
'    Call GetFilesFromServer(gcServerPath & strCommonPath & "\")
'    Call GetFilesFromServer(gcServerPath & strProjectId & "\")


    Call GetFilesFromServer(gcServerPath)
'''''    Call GetFilesFromServer(gNetDriveChar)
    
    If NewCount = 0 Then
        '�ֽŹ��� �ޱ⸦ �� ���ڿ� �ð�,���� write�� ���´�.
'        Call SaveSetting("Schweitzer", "Download", "LastDate", CStr(Format(Now, "yyyyMMddhhmm")))
'        WritePrivateProfileString "Version", "LastDate", CStr(Format(Now, "yyyyMMddhhmm")), App.Path & "\Version.ini"
        Call medSetINI("Version", "LastDate", CStr(Format(Now, "yyyyMMddhhmm")), "C:\Schweitzer\Schweitzer.ini")
        lblMessage.Caption = "���׷��̵� ���� �ʾƵ� �˴ϴ�."
        LogWrite  '##
        Call medSleep(1000)
        Call ExitProgram
    End If
    
'''''    Call FileCopyFromServer

    Call ChkAndKillProcess
    If Not CopyNewVersion Then    '����Copy
        MsgBox "Error �߻�.. ����Ƿ� �����ٶ��ϴ�. (3577)"
        Call ExitProgram
    End If
        
    DoEvents
    lblMessage.Caption = "���׷��̵尡 ���������� �Ϸ�Ǿ����ϴ�."
    LogWrite  '##
    Call medSleep(1000)
    
    If blnNetCon Then Call ConNetDrive(1)  '��Ʈ�p ����̺� ����
    Call medSleep(1000)
    
    lblMessage.Caption = strProjectId & " �ý����� ����˴ϴ�."
    LogWrite  '##
    Close #Fd
    
    '�ֽŹ��� �ޱ⸦ �� ���ڿ� �ð�,���� write�� ���´�.
'    Call SaveSetting("Schweitzer", "Download", "LastDate", CStr(Format(Now, "yyyyMMddhhmm")))

'    WritePrivateProfileString "Version", "LastDate", CStr(Format(Now, "yyyyMMddhhmm")), App.Path & "\Version.ini"
    Call medSetINI("Version", "LastDate", CStr(Format(Now, "yyyyMMddhhmm")), "C:\Schweitzer\Schweitzer.ini")
    If ExeNames <> "" Then
        Dim aryTmp() As String
        Dim i As Long
        
        aryTmp = Split(ExeNames, Chr(19))
        
        For i = LBound(aryTmp) To UBound(aryTmp)
            If aryTmp(i) <> "" Then
                Shell aryTmp(i), vbNormalFocus
            End If
        Next i
    End If

    If Not blnDownloadMyself Then
        Call RestoreAll
'        If gExeFile <> "" Then
'            If Not CheckApp(gExeFile) And _
'               Dir(gExeFile) <> "" Then
'                    '������Ʈ���� ����ȭ���� ��ϵǾ� ���� ��쿡�� �����ϰ��Ѵ�.... wooil
'                    Shell gExeFile, vbNormalFocus                 '���α׷� �⵿
'            End If
'        End If
    Else
        Call DownloadMyself(App.Title)
    End If
    ReleaseNetDir gNetDriveChar
    End
End Sub

Private Sub ChkAndKillProcess()
    Dim i As Integer
    Dim strAppName As String
    Dim strAppPath As String
    Dim ResumeCnt As Integer
    Dim blnCopy As Boolean
    Dim strSysDir As String

    On Error GoTo Err_Trap
    
    ResumeCnt = 0
    
    lblMessage.Caption = "�ý����� ���������� �˻��մϴ�."
    LogWrite  '##
    
    prgBar.Min = 0
    prgBar.Max = FileCount
    
    For i = 1 To FileCount
        DoEvents
        prgBar.value = prgBar.value + 1
        With gsFileInfo(i)
            lblFile = .DestPath & .FileNm
'            LogWrite (.FileNm & "-->" & .flag)
            If .flag And (UCase(Replace(.SvrPath, gcServerPath, "") & .FileNm) <> UCase(App.Title) & ".EXE") Then
                '������ ȭ���� �������̸� ���������Ѵ�.
                Call CheckApp(.DestPath & .FileNm, True)
            End If
        End With
    Next
    Exit Sub
    
Err_Trap:
    If ResumeCnt > 5 Then
        lblMessage.Caption = "Error : " & Err.Description
        LogWrite  '##
        Exit Sub
    End If
    Call medSleep(1000)
    ResumeCnt = ResumeCnt + 1
    On Error GoTo Err_Trap
    Resume

End Sub
Private Sub FileCopyFromServer()

    Dim Resp As VbMsgBoxResult
    
    RegisterServiceProcess GetCurrentProcessId, 1 'Hide app
    If CheckApp(gsAppPath & gsAppName) Then  '���μ����� �������̸�..
        DoEvents
        LogWrite ("�ش� ���α׷��� ���� �������Դϴ�.")  '##
        Resp = MsgBox("�ش� ���α׷��� ���� �������Դϴ�. ���� �����Ͻðڽ��ϱ�?", _
                       vbYesNo + vbQuestion + vbDefaultButton2, "����üũ")
        If Resp = vbYes Then
            LogWrite ("��� ���α׷��� �����ϰ� �� ������ �����մϴ�.")  '##
            Resp = MsgBox("��� ���α׷��� �����ϰ� �� ������ �����մϴ�.", _
                           vbOKCancel + vbExclamation, "����üũ")
            If Resp = vbCancel Then GoTo NoCopy
            Call CheckApp(gsAppPath & gsAppName, True)   '��������
            Call medSleep(3000)
        Else
NoCopy:
            LogWrite ("�� ������ ������� �ʾҽ��ϴ�..")  '##
            Resp = MsgBox("�� ������ ������� �ʾҽ��ϴ�. " & vbCrLf & _
                          "��� ���α׷��� ���� �� ����üũ ���α׷��� �ٽ� �����Ű�ʽÿ�.", _
                          vbExclamation + vbOKOnly, "����üũ")
            Call ExitProgram
        End If
    End If
   
    If Not CopyNewVersion Then    '����Copy
        MsgBox "Error �߻�.. ����Ƿ� �����ٶ��ϴ�. (3577)"
        Call ExitProgram
    End If
    
Finish:
    RegisterServiceProcess GetCurrentProcessId, 0 'Hide app

End Sub

Private Sub DownloadMyself(ByVal strExeNm As String)
    
On Error Resume Next
    
'    If gExeFile <> "" Then Call Shell(gExeFile, vbNormalFocus)
'    Call Shell(gsAppPath & "GetDownloadProgram.EXE " & strProjectId & " " & strExeNm)
    If Dir(App.Path & "\" & "GetDownloadProgram.EXE ") <> "" Then
        Call Shell(App.Path & "\" & "GetDownloadProgram.EXE " & gExeFile, vbNormalFocus)
    End If
End Sub

Private Sub LogWrite(Optional ByVal strText As String = "")
    DoEvents
    If strText = "" Then
        Print #Fd, lblMessage.Caption
    Else
        Print #Fd, strText
    End If
End Sub

Private Sub GetDir()
    Dim strNetDrive As String
    
    lblMessage.Caption = "���׷��̵� ������ ã�� �����ϴ�."
    lblMessage.Refresh
    
    gsWinPath = GetWinDir & "\"     'Windows ����
    gsSysPath = GetSysDir & "\"     'System ����
    
    'Application�� Exeȭ���� �����ϴ� ����
'''''    gsAppPath = GetSetting(RegHdApp, RegSsApp, RegK1App, "")
    
    'Server�� Download ����
'    gNetDrive = S2GetSetting("GNV", "DOWNLOAD", "PATH", "")
    gNetDrive = medGetINI("DownLoad", "Path", "C:\Schweitzer\Schweitzer.ini")
    
    Do
        If gNetDrive <> "" Then Exit Do

        If MsgBox("������ �� �� �����ϴ�. ���� �����Ͻð����ϱ�", vbQuestion + vbYesNo, "��������") = vbNo Then Unload Me

        strNetDrive = InputBox("���׷��̽� ���� : ", "�����Է�", "")
        
        gNetDrive = strNetDrive
        Call medSetINI("DownLoad", "Path", strNetDrive, "C:\Schweitzer\Schweitzer.ini")
'        Call S2SaveSetting("GNV", "DOWNLOAD", "PATH", strNetDrive)
    Loop
    If Mid(gNetDriveChar, Len(gNetDriveChar)) <> "\" Then
        gcServerPath = gNetDriveChar & "\"
    Else
        gcServerPath = gNetDriveChar
    End If
    
'''''
'''''   wooil
'''''
'''''    If Mid(gsAppPath, Len(gsAppPath)) <> "\" Then gsAppPath = gsAppPath & "\"
'''''
'''''    'Server�� Download ����
'''''    gNetDrive = GetSetting(RegHdSet, RegSsSet, RegK1Set, "")
'''''    If Mid(gNetDriveChar, Len(gNetDriveChar)) <> "\" Then gcServerPath = gNetDriveChar & "\"
'''''
'''''    'Application�� ����ȭ�ϸ�
'''''    gsAppName = GetSetting(RegHdApp, RegSsApp, RegK2App, "")
End Sub

Private Sub ConNetDrive(ByVal Index As Integer)
    Dim strConnect As String

On Error GoTo ConNetDrive_error

    If Index = 0 Then
        If gID <> "" Then
            gcServerPath = gNetDrive
    '        If Mid(gcServerPath, Len(gcServerPath)) <> "\" Then gcServerPath = gcServerPath & "\"
            If Mid(gcServerPath, Len(gcServerPath)) = "\" Then gcServerPath = Mid(gcServerPath, 1, Len(gcServerPath) - 1)
            lblMessage.Caption = "���׷��̵� ������ �����ϰ� �ֽ��ϴ�."
            lblMessage.Refresh
            strNetDrive = AttachNetDir(gPWD, gID, gcServerPath, gNetDriveChar)
    '        strConnect = Dir(gcServerPath, vbDirectory)
            If strNetDrive = "" Then
                MsgBox "���׷��̵� ������ ������ �� �����ϴ�.", vbCritical, "���� ���� ����"
                Call ExitProgram
            End If
            If Mid(gcServerPath, Len(gcServerPath)) <> "\" Then gcServerPath = gcServerPath & "\"
        Else
            gcServerPath = gNetDrive
            If Mid(gcServerPath, Len(gcServerPath)) <> "\" Then gcServerPath = gcServerPath & "\"
            lblMessage.Caption = "���׷��̵� ������ �����ϰ� �ֽ��ϴ�."
            lblMessage.Refresh
            strConnect = Dir(gcServerPath, vbDirectory)
            If strNetDrive = "" Then
                MsgBox "���׷��̵� ������ ������ �� �����ϴ�.", vbCritical, "���� ���� ����"
                Call ExitProgram
            End If
        End If
    Else
        ReleaseNetDir gNetDriveChar
    End If
    Exit Sub
    
ConNetDrive_error:
    MsgBox "���׷��̵� ������ ������ �� �����ϴ�.", vbCritical, "���� ���� ����"
    Call ExitProgram



'''''
''''' ������� �ʴ´�.
'''''
'''''    If Index = 0 Then
'''''        lblMessage.Caption = "���׷��̵� ������ �����ϰ� �ֽ��ϴ�."
'''''        LogWrite  '##
'''''        DoEvents
''''''        If Not NetConnect(0, gNetDrive) Then     '��Ʈ�p ����̺� ����
'''''            blnNetCon = False
'''''            gcServerPath = gNetDrive
'''''            If Mid(gcServerPath, Len(gcServerPath)) <> "\" Then gcServerPath = gcServerPath & "\"
''''''            LogWrite ("��Ʈ�p ����̺갡 ���������� ������� �ʾҽ��ϴ�.") '##
''''''            MsgBox "��Ʈ�p ����̺갡 ���������� ������� �ʾҽ��ϴ�." & vbCrLf & _
''''''                        "����ǿ� �����Ͻʽÿ�. (3577)", vbCritical + vbOKOnly, "Error"
''''''            Call ExitProgram
''''''        Else
''''''            blnNetCon = True
''''''        End If
'''''    Else
'''''        If Not blnNetCon Then Exit Sub
'''''        lblMessage.Caption = "���׷��̵� �������� ������ �����ϰ� �ֽ��ϴ�."
'''''        LogWrite  '##
'''''        DoEvents
'''''        If Not NetConnect(1, gNetDrive) Then     '��Ʈ�p ����̺� ����
'''''            LogWrite ("��Ʈ�p ����̺갡 ���������� �������� �ʾҽ��ϴ�.") '##
'''''            MsgBox "��Ʈ�p ����̺갡 ���������� �������� �ʾҽ��ϴ�." & vbCrLf & _
'''''                   "����ǿ� �����Ͻʽÿ�.", vbCritical + vbOKOnly, "Error"
'''''            blnNetCon = False
'''''            Call RestoreAll
'''''            Close #Fd
'''''            End
'''''        End If
'''''    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseNetDir gNetDriveChar
End Sub

Private Sub Timer1_Timer()
    Static iSeq As Integer
    If iSeq = imlNew.ListImages.Count Then iSeq = 0
    iSeq = iSeq + 1
    imgNew.Picture = imlNew.ListImages(iSeq).Picture
    DoEvents
End Sub

Private Function CopyNewVersion() As Boolean
    
    Dim i As Integer
    Dim strAppName As String
    Dim strAppPath As String
    Dim ResumeCnt As Integer
    Dim blnCopy As Boolean
    Dim strSysDir As String

    On Error GoTo Err_Trap
    
    ResumeCnt = 0
    CopyNewVersion = True
    
    lblMessage.Caption = "���׷��̵� �ϰ� �����ϴ�."
    LogWrite  '##
    
    For i = 1 To FileCount
        DoEvents
        prgBar.value = i
        With gsFileInfo(i)
        
            lblFile = .DestPath & .FileNm
            
            'MsgBox UCase(.FileNm) & ", " & UCase(App.Title) & ".EXE"
            If .flag And UCase(Replace(.SvrPath, gcServerPath, "") & .FileNm) <> UCase(App.Title) & ".EXE" Then      '���ο� ����
            
                lblCopyCount = Val(lblCopyCount) + 1
                lblCopyCount.Refresh
                
                Call CheckPath(.DestPath)       '����üũ �� ����
                strAppName = .DestPath & .FileNm
'''''                lblMessage.Caption = .FileNm & " ���� ��.."   '���ϸ�
                LogWrite (strAppName & " ���� ��..") '##
                If .FileExtend = "DLL" Or .FileExtend = "OCX" Then
                    Call ExecCmd(Chr(34) & gsSysPath & "Regsvr32.exe " & Chr(34) & " /u /s " & Chr(34) & strAppName & Chr(34))     ', vbMinimizedNoFocus)
                    blnCopy = CopyFile(.SvrPath, .DestPath, .FileNm, .FileNm)
                    
                    If Not blnCopy Then GoTo Err_Trap
                    Call ExecCmd(Chr(34) & gsSysPath & "Regsvr32.exe " & Chr(34) & " /s " & Chr(34) & strAppName & Chr(34))     ', vbMinimizedNoFocus)
                    LogWrite (strAppName & " Registered") '##
                    DoEvents
                ElseIf .FileExtend = "REG" Then
                    blnCopy = CopyFile(.SvrPath, .DestPath, .FileNm, .FileNm)
                    If Not blnCopy Then GoTo Err_Trap
                    Call ExecCmd("REGEDIT  /s " & Chr(34) & strAppName & Chr(34))       ', vbMinimizedNoFocus)
                    LogWrite (strAppName & " Registered") '##
                Else
                    blnCopy = CopyFile(.SvrPath, .DestPath, .FileNm, .FileNm)
                    If Not blnCopy Then GoTo Err_Trap
'                    FileCopy .SvrPath & .FileNm, strAppName
                    If .FileExtend = "EXE" And (UCase(.FileNm) <> UCase(gsAppName)) Then
                    
'                        ExecCmd (strAppName & " /REGSERVER")
                    End If
                End If
            End If
        End With
    Next
    Exit Function
    
Err_Trap:
    If ResumeCnt > 5 Then
        CopyNewVersion = False
        lblMessage.Caption = "Error : " & Err.Description
        LogWrite  '##
        Exit Function
    End If
    Call medSleep(1000)
    ResumeCnt = ResumeCnt + 1
    On Error GoTo Err_Trap
    Resume
    
End Function

Public Sub CheckPath(ByVal strPath As String)
    Dim i As Long
    Dim strDir As String
    Dim lngPos As Long
    
    On Error GoTo ErrPath
    
    i = 0
    lngPos = InStr(strPath, "\")
    While (lngPos > 0)
        strDir = Mid(strPath, 1, lngPos)
        If Dir(strDir, vbDirectory) = "" Then
            Call MkDir(strDir)
        End If
        lngPos = InStr(lngPos + 1, strPath, "\")
    Wend
ErrPath:
End Sub

Public Function CheckApp(myName As String, Optional ByVal KillFg As Boolean = False) As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    
    On Local Error GoTo Finish
    
    CheckApp = False
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
'''''    lblMessage.Caption = "���μ����� ���������� üũ�ϰ� �ֽ��ϴ�."
    LogWrite (lblMessage & "->" & myName) '##
    DoEvents
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        Dim aryTmp() As String
        Dim strTmp As String
'        Dim aryTmp2() As String
'        Dim strTmp2 As String
        
        
        aryTmp = Split(szExename, "\")
        strTmp = LCase(aryTmp(UBound(aryTmp)))

        '������ó�� ���õ� ���α׷��� ��� Kill
        Select Case strTmp
            Case "s2aps.exe", "s2bbs.exe", "s2lis.exe", "s2iis.exe", _
                    "wardmenu_nurse.exe", "wardmenu_result.exe":
                ExeNames = ExeNames & Chr(19) & szExename
                myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                AppKill = TerminateProcess(myProcess, exitCode)
                Call CloseHandle(myProcess)
                LogWrite (myName & " killed...") '##
                DoEvents
        End Select
        
'        If strTmp = strTmp2 Then
'            '���μ��� ��������
'            If KillFg Then
'                ExeNames = ExeNames & Chr(19) & szExename
'                myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
'                AppKill = TerminateProcess(myProcess, exitCode)
'                Call CloseHandle(myProcess)
'                LogWrite (myName & " killed...") '##
'                DoEvents
'            End If
'            CheckApp = True
'            Exit Do  '�ش� ���μ����� �߰ߵǸ� �׸�ã��
'        End If

        rProcessFound = ProcessNext(hSnapshot, uProcess)
        DoEvents
    Loop

    Call CloseHandle(hSnapshot)
Finish:
End Function

Private Function GetDirectoriesFromServer(ByVal strSvrPath As String) As String
    
    Dim strDirs As String
    Dim strFileNm As String
    
    strDirs = ""
    strFileNm = Dir(strSvrPath, vbDirectory)
    While (strFileNm <> "")
        If strFileNm <> "." And strFileNm <> ".." Then
            strDirs = IIf(strDirs = "", strFileNm, strDirs & ":" & strFileNm)
        End If
        strFileNm = Dir()
    Wend
    GetDirectoriesFromServer = strDirs
End Function


Private Sub GetFilesFromServer(ByVal strSvrPath As String)

    Static IsNetConn As Boolean

    Dim strFileNm As String
    Dim strExtend As String
    Dim strOldVersion As String
    Dim strOldDtTm As String
    Dim strSubDirs As String
    Dim i As Long
    
    On Error GoTo Err_Trap
    
    If IsNetConn = False Then lblMessage.Caption = "���׷��̵� ȭ���� ã�� �����ϴ�."
    LogWrite  '##
    
    strSubDirs = GetDirectoriesFromServer(strSvrPath)
    arySubDirs = Split(strSubDirs, ":")
    For i = LBound(arySubDirs) To UBound(arySubDirs)
        If (GetAttr(strSvrPath & arySubDirs(i)) And vbDirectory) = vbDirectory Then
            Call GetFilesFromServer(strSvrPath & arySubDirs(i) & "\")
        End If
    Next
    
    strFileNm = Dir(strSvrPath)
    While (strFileNm <> "")
        lblFile.Caption = Replace(strSvrPath & strFileNm, gcServerPath, "")
        lblFile.Refresh
'        lblMessage.Caption = "�� File�� Version Check�� �ϰ� �ֽ��ϴ�.. ( " & strFileNm & " )"
'        DoEvents
        
        FileCount = FileCount + 1
        lblFileCount = FileCount
        lblFileCount.Refresh
        
        ReDim Preserve gsFileInfo(FileCount)
        
        With gsFileInfo(FileCount)
            .FileNm = strFileNm
            .FileSize = FileLen(strSvrPath & strFileNm)
            .FileDtTm = FileDateTime(strSvrPath & strFileNm)
            .FileVersion = GetFileVersion(strSvrPath & strFileNm)
            strExtend = UCase(medGetP(strFileNm, 2, "."))
            .FileExtend = strExtend
            .SvrPath = strSvrPath
            
            
'            .flag = True
'            NewCount = NewCount + 1
            
            
'''''''''''''
'''''''''''''   ��� ȭ���� �޴´�.
'''''''''''''
'''''''''''''
'''''''''''''
''''''''''''''''''            .DestPath = gsAppPath & "..\.." & Mid(strSvrPath, Len(gcServerPath))
'''''''''''''
            .DestPath = Replace(strSvrPath, gcServerPath, App.Path & "\")
'''''''''''''
''''''''''''''            Debug.Print .DestPath & " --> " & strFileNm
'''''''''''''
            strOldVersion = GetFileVersion(.DestPath & .FileNm)
            strOldDtTm = GetFileDateTime(.DestPath & .FileNm)
            If .FileVersion <> vbNullString Then
                If .FileVersion > strOldVersion Then   '������
                    .flag = True
                    NewCount = NewCount + 1
                Else
                    GoTo DateCompare
                End If
            Else
DateCompare:
                If .FileDtTm <> strOldDtTm Then  '������ ��
                    .flag = True
                    NewCount = NewCount + 1
                Else
                    .flag = False
                End If
            End If

            If UCase(Replace(.SvrPath, gcServerPath, "") & .FileNm) = UCase(App.Title & ".EXE") Then blnDownloadMyself = True
            LogWrite (.FileNm & vbTab & .FileVersion & vbTab & .FileDtTm & vbTab & .FileSize & vbTab & .DestPath & vbTab & .flag)   '##
        End With
        strFileNm = Dir
        
        lblNewCount = NewCount
                
    Wend
    
    DoEvents
    LogWrite (CStr(NewCount)) '##
    prgBar.Max = FileCount
    Exit Sub
    
Err_Trap:
    LogWrite (Err.Number & " : " & Err.Description)
    Resume Next
End Sub


'Private Sub GetFilesFromServer(ByVal strSvrPath As String)
'
'    Dim strFileNm As String
'    Dim strExtend As String
'    Dim strOldVersion As String
'    Dim strOldDtTm As String
'
'    On Error GoTo Err_Trap
'
'    lblMessage.Caption = "�� File�� Version Check�� �ϰ� �ֽ��ϴ�.."
'    LogWrite  '##
'
'    strFileNm = Dir(strSvrPath)
'    While (strFileNm <> "")
'        lblMessage.Caption = "�� File�� Version Check�� �ϰ� �ֽ��ϴ�.. ( " & strFileNm & " )"
'        DoEvents
'
'        FileCount = FileCount + 1
'        ReDim Preserve gsFileInfo(FileCount)
'
'        With gsFileInfo(FileCount)
'            .FileNm = strFileNm
'            .FileSize = FileLen(strSvrPath & strFileNm)
'            .FileDtTm = FileDateTime(strSvrPath & strFileNm)
'            .FileVersion = GetFileVersion(strSvrPath & strFileNm)
'            strExtend = UCase(medGetP(strFileNm, 2, "."))
'            .FileExtend = strExtend
'            .SvrPath = strSvrPath
'            Select Case strExtend
'                Case "EXE": .DestPath = gsAppPath
'                Case "HLP": .DestPath = gsAppPath & "..\Help\"
'                Case "BMP": .DestPath = gsAppPath & "..\Help\image\"
'                Case "RPT": .DestPath = gsAppPath & "..\RPT\"
'                Case "OCX", "DLL", "LIC":
'                    If UCase(.FileNm) Like "S2*" Then
'                        If UCase(.FileNm) Like "S2" & UCase(strProjectId) & "*" Then
'                            .DestPath = gsAppPath
'                        Else
'                            .DestPath = gsAppPath & "..\..\Common\DLL\"
'                        End If
'                    Else
'                        .DestPath = gsAppPath & "..\..\Common\System\"
'                    End If
'                    '.DestPath = gsSysPath
'                Case Else: .DestPath = gsAppPath & "..\ETC\"
'            End Select
'            strOldVersion = GetFileVersion(.DestPath & .FileNm)
'            strOldDtTm = GetFileDateTime(.DestPath & .FileNm)
'            If chkNewVersion.Value = 0 Then
'                .flag = True
'                NewCount = NewCount + 1
'            Else
'                If .FileVersion <> vbNullString Then
'                    If .FileVersion > strOldVersion Then   '������
'                        .flag = True
'                        NewCount = NewCount + 1
'                    Else
'                        GoTo DateCompare
'                    End If
'                Else
'DateCompare:
'                    If .FileDtTm > strOldDtTm Then  '������ ��
'                        .flag = True
'                        NewCount = NewCount + 1
'                    Else
'                        .flag = False
'                    End If
'                End If
'            End If
'            If UCase(.FileNm) = UCase(App.Title & ".EXE") Then blnDownloadMyself = True
'            LogWrite (.FileNm & vbTab & .FileVersion & vbTab & .FileDtTm & vbTab & .FileSize & vbTab & .DestPath & vbTab & .flag)   '##
'        End With
'        strFileNm = Dir
'
'    Wend
'
'    DoEvents
'    LogWrite (CStr(NewCount)) '##
'    prgBar.Max = FileCount
'    Exit Sub
'
'Err_Trap:
'    LogWrite (Err.Number & " : " & Err.Description)
'    Resume Next
'End Sub


Sub ExitProgram()
    Call RestoreAll
    Call ConNetDrive(1)  '��Ʈ�p ����̺� ����
    Close #Fd
    If blnDownloadMyself Then
        Unload Me
        Call DownloadMyself(App.EXEName)
    End If
    End
End Sub


Function CopyFile(ByVal strSrcDir As String, ByVal strDestDir As String, ByVal strSrcName As String, ByVal strDestName As String) As Boolean
    Const intUNKNOWN% = 0
    Const intCOPIED% = 1
    Const intNOCOPY% = 2
    Const intFILEUPTODATE% = 3

    '
    'VerInstallFile() Flags
    '
    Const VIFF_FORCEINSTALL% = &H1
    Const VIF_TEMPFILE& = &H1
    Const VIF_SRCOLD& = &H4
    Const VIF_DIFFLANG& = &H8
    Const VIF_DIFFCODEPG& = &H10
    Const VIF_DIFFTYPE& = &H20
    Const VIF_WRITEPROT& = &H40
    Const VIF_FILEINUSE& = &H80
    Const VIF_OUTOFSPACE& = &H100
    Const VIF_ACCESSVIOLATION& = &H200
    Const VIF_SHARINGVIOLATION = &H400
    Const VIF_CANNOTCREATE = &H800
    Const VIF_CANNOTDELETE = &H1000
    Const VIF_CANNOTRENAME = &H2000
    Const VIF_OUTOFMEMORY = &H8000&
    Const VIF_CANNOTREADSRC = &H10000
    Const VIF_CANNOTREADDST = &H20000
    Const VIF_BUFFTOOSMALL = &H40000

    Static fIgnoreWarn As Integer             'user warned about ignoring error?

    Dim lRC As Long
    Dim lpTmpNameLen As Long
    Dim intFlags As Integer
    Dim intRESULT As Integer
    Dim fFileAlreadyExisted
    Dim mstrVerTmpName As String                                'temp file name for VerInstallFile API
    Dim intFD As Integer

    On Error Resume Next

    CopyFile = False

    '
    'Setup for VerInstallFile call
    '
    lpTmpNameLen = gintMAX_SIZE
    mstrVerTmpName = String$(lpTmpNameLen, 0)
    'fFileAlreadyExisted = FileExists(strDestDir & strDestName)

    intRESULT = intUNKNOWN
    intFlags = VIFF_FORCEINSTALL

    Do While intRESULT = intUNKNOWN
        'VerInstallFile under Windows 95 does not handle
        '  long filenames, so we must give it the short versions
        '  (32-bit only).
        Dim strShortSrcName As String
        Dim strShortDestName As String
        Dim strShortSrcDir As String
        Dim strShortDestDir As String
        
        strShortSrcName = strSrcName
        strShortSrcDir = strSrcDir
        strShortDestName = strDestName
        strShortDestDir = strDestDir
        If Dir(strDestDir & strDestName) = vbNullString Then
            'If the destination file does not already
            '  exist, we create a dummy with the correct
            '  (long) filename so that we can get its
            '  short filename for VerInstallFile.
            intFD = FreeFile
            Open strDestDir & strDestName For Output Access Write As #intFD
            Close #intFD
        End If
    
        On Error GoTo UnexpectedErr
        
        If Not IsWindowsNT() Then
            Dim strTemp As String
            'This conversion is not necessary under Windows NT
            strShortSrcDir = GetShortPathName(strSrcDir)
            If GetFileName(strSrcName) = strSrcName Then
                strShortSrcName = GetFileName(GetShortPathName(strSrcDir & strSrcName))
            Else
                strTemp = GetShortPathName(strSrcDir & strSrcName)
                strShortSrcName = Mid$(strTemp, Len(strShortSrcDir) + 1)
            End If
            strShortDestDir = GetShortPathName(strDestDir)
            strShortDestName = GetFileName(GetShortPathName(strDestDir & strDestName))
        End If
        On Error Resume Next
            
        
        Call FileCopy(strShortSrcDir & strShortSrcName, strShortDestDir & strShortDestName)
        
        'lRC = VerInstallFile(intFlags, strShortSrcName, strShortDestName, strShortSrcDir, strShortDestDir, 0&, mstrVerTmpName, lpTmpNameLen)
        
        
        'If Err <> 0 Then
            '
            'If the version or file expansion DLLs couldn't be found, then abort setup
            '
        '    ExitSetup frmCopy, gintRET_FATAL
        'End If

        If lRC = 0 Then
            '
            'File was successfully installed, increment reference count if needed
            '
            
            'One more kludge for long filenames: VerInstallFile may have renamed
            'the file to its short version if it went through with the copy.
            'Therefore we simply rename it back to what it should be.
            Name strDestDir & strShortDestName As strDestDir & strDestName
            intRESULT = intCOPIED
            CopyFile = True
        ElseIf lRC And VIF_SRCOLD Then
            '
            'Source file was older, so not copied, the existing version of the file
            'will be used.  Increment reference count if needed
            '
            intRESULT = intFILEUPTODATE
            CopyFile = True
        ElseIf lRC And (VIF_DIFFLANG Or VIF_DIFFCODEPG Or VIF_DIFFTYPE) Then
            '
            'We retry and force installation for these cases.  You can modify the code
            'here to prompt the user about what to do.
            '
            intFlags = VIFF_FORCEINSTALL
        ElseIf lRC And VIF_WRITEPROT Then
            lblMessage.Caption = "��� ������ ���� �����Ǿ� �ֽ��ϴ�."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_FILEINUSE Then
            lblMessage.Caption = "��� ������ ����ϰ� �ֽ��ϴ�. �ٸ� ��� ���� ���α׷��� ���� �ִ��� Ȯ���Ͻʽÿ�."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_OUTOFSPACE Then
            lblMessage.Caption = "��� ����̺��� ������ �����մϴ�."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_ACCESSVIOLATION Then
            lblMessage.Caption = "������ �����ϴ� ���� �׼����� �����Ͽ����ϴ�."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_SHARINGVIOLATION Then
            lblMessage.Caption = "������ �����ϴ� ���� ������ �����Ͽ����ϴ�."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_OUTOFMEMORY Then
            lblMessage.Caption = "������ �����ϴ� �� ����� �޸𸮰� �����մϴ�."
            GoTo UnexpectedErr
        Else
            '
            ' For these cases, we generically report the error and do not install the file
            ' unless this is an SMS install; in which case we abort.
            '
            If lRC And VIF_CANNOTCREATE Then
                lblMessage.Caption = "�ӽ� ������ ���� �� �����ϴ�."
            ElseIf lRC And VIF_CANNOTDELETE Then
                lblMessage.Caption = "���� ��� ������ ������ �� �����ϴ�."
            ElseIf lRC And VIF_CANNOTRENAME Then
                lblMessage.Caption = "�ӽ� ���� �̸��� �ٲ� �� �����ϴ�."
            ElseIf lRC And VIF_CANNOTREADSRC Then
                lblMessage.Caption = "���� ������ ���� �� �����ϴ�."
            ElseIf lRC And VIF_CANNOTREADDST Then
                lblMessage.Caption = "��� ���� �Ӽ��� ���� �� �����ϴ�."
            ElseIf lRC And VIF_BUFFTOOSMALL Then
                lblMessage.Caption = "���� ���� �����Դϴ�."
            End If
            GoTo UnexpectedErr
        End If
    Loop

    '
    'If there was a temp file left over from VerInstallFile, remove it
    '
    If lRC And VIF_TEMPFILE Then
        Kill mstrVerTmpName
    End If

    Exit Function

UnexpectedErr:
    Call LogWrite  '##
    Call ExitProgram
End Function


