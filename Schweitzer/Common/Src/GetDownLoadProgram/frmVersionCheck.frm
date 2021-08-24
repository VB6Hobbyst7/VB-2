VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmVersionCheck 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FCEFE9&
   Caption         =   "Get Download"
   ClientHeight    =   990
   ClientLeft      =   1260
   ClientTop       =   2415
   ClientWidth     =   3765
   Icon            =   "frmVersionCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   3765
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Getting Download Program now......."
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   345
      TabIndex        =   0
      Top             =   225
      Width           =   3075
   End
End
Attribute VB_Name = "frmVersionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FileCount As Integer
Private NewCount As Integer
Private LogName As String
Private Fd As Integer

Private blnNetCon As Boolean

Const TH32CS_SNAPPROCESS As Long = 2&

Private Sub Form_Load()
    Dim i As Long
    
'    Me.Show
    Me.ZOrder 0
    Call medAlwaysOn(Me, 1)
    DoEvents
    Call medSleep(1000)
    DoEvents
    
'    Me.SetFocus
    DoEvents
    
    blnNetCon = False
    
    prgBar.value = 0
    Call GetDir
    
    

    For i = 0 To gNetCount - 1
        Fd = FreeFile
        LogName = App.Path & "\GetDownload.log"
        Open LogName For Output As #Fd
        
        If CheckApp(gsAppPath & gsAppName) Then  '프로세스가 실행중이면..
            Call CheckApp(gsAppPath & gsAppName, True)   '강제종료
            DoEvents
            Call medSleep(3000)
            DoEvents
        End If
        
        If Dir(gNetDriveChar(i)) <> "" Then Call ConNetDrive(1, i)  '네트웤 드라이브 해제
        Call ConNetDrive(0, i)  '네트웤 드라이브 연결
        Call GetFilesFromServer(gcServerPath(i), i)
        
        Call medSleep(1000)
        If NewCount = 0 Then
            lblMessage.Caption = "프로그램을 종료합니다."
            'MsgBox "새로운 버전이 없습니다..프로그램을 종료합니다."
            LogWrite  '##
            medSleep (1000)
            
            GoTo NoData
            
        End If
        Call FileCopyFromServer(i)
        
        DoEvents
        lblMessage.Caption = "작업이 정상적으로 종료되었습니다."
        LogWrite  '##
        Call medSleep(1000)
        
        Call RestoreAll
        If blnNetCon Then Call ConNetDrive(1, i)  '네트웤 드라이브 해제
        Call medSleep(1000)
        
        LogWrite  '##
        Close #Fd
        Exit For

NoData:
        lblMessage = ""
        prgBar.value = 0
        FileCount = 0
        NewCount = 0

        Call ExitProgram(i)
    Next
    
    End
End Sub

Private Sub FileCopyFromServer(ByVal pindex As Long)
    Dim Resp As VbMsgBoxResult

    If CheckApp(gsAppPath & gsAppName) Then  '프로세스가 실행중이면..
       
        DoEvents
        LogWrite ("해당 프로그램이 현재 실행중입니다.")  '##
        Resp = MsgBox("해당 프로그램이 현재 실행중입니다. 지금 종료하시겠습니까?", _
                       vbYesNo + vbQuestion + vbDefaultButton2, "버전체크")

        If Resp = vbYes Then
            MsgBox "3"
            LogWrite ("모든 프로그램을 종료하고 새 파일을 복사합니다.")  '##
            Resp = MsgBox("모든 프로그램을 종료하고 새 파일을 복사합니다.", _
                           vbOKCancel + vbExclamation, "버전체크")
            If Resp = vbCancel Then GoTo NoCopy
            Call CheckApp(gsAppPath & gsAppName, True)   '강제종료
            Call medSleep(3000)
        Else
NoCopy:

            LogWrite ("새 버전이 복사되지 않았습니다..")  '##
            Resp = MsgBox("새 버전이 복사되지 않았습니다. " & vbCrLf & _
                          "모든 프로그램을 종료 후 버전체크 프로그램을 다시 실행시키십시오.", _
                          vbExclamation + vbOKOnly, "버전체크")

            Call ExitProgram(pindex)
        End If
    End If
   
    If Not CopyNewVersion(pindex) Then    '파일Copy
        MsgBox "Error 발생.. 전산실로 연락바랍니다. "
        Call ExitProgram(pindex)
    End If
End Sub

Private Sub DownloadMyself(ByVal strExeNm As String)
'
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
    Dim strAppPath As String
    Dim strNetCount As String
    Dim strMsg As VbMsgBoxResult
    Dim Ret As Long
    Dim strTmp As String
    Dim aryTmp() As String
    Dim strLastDrive As String
    Dim aryNetDrive() As String
    Dim aryNetDriveChar() As String
    Dim aryClientPath() As String
    Dim i As Long
    
    Ret = 0
    strMsg = vbNo
    
    gsWinPath = GetWinDir & "\"     'Windows 폴더
    gsSysPath = GetSysDir & "\"     'System 폴더
    
    
    'gsAppPath = S2GetSetting("GNV", "APP PATH", "PATH", "")
    gsAppPath = medGetINI("DownLoad", "Path", "C:\Schweitzer\Schweitzer.ini")
    
    Do
        If gsAppPath <> "" Then
            strMsg = vbNo
            Exit Do
        End If
        
        strMsg = MsgBox("업그레이드 실행 프로그램의 경로를 알수 없습니다. 지금 설정하시겠습니까?", vbQuestion + vbYesNo, "경로설정")
        
        If strMsg = vbNo Then Exit Sub
        
        strAppPath = InputBox("프로그램 경로 : ", "경로 입력", "")
        
         If strAppPath = "" Then
            strMsg = vbNo
            Exit Sub
        End If
        
        gsAppPath = strAppPath
        Call medSetINI("DownLoad", "Path", strAppPath, "C:\Schweitzer\Schweitzer.ini")
'        Call S2SaveSetting("GNV", "APP PATH", "PATH", strAppPath)
    Loop
    
'
'
'    gNetCount = S2GetSetting("GNV", "SERVER CNT", "CNT", "")
'
'    Do
'        If gNetCount <> "" Then
'            strMsg = vbNo
'            Exit Do
'        End If
'
'        strMsg = MsgBox("업그레이드할 서버의 갯수를 알수 없습니다. 지금 설정하시겠습니까?", vbQuestion + vbYesNo, "서버설정")
'
'        If strMsg = vbNo Then Exit Sub
'
'        strNetCount = InputBox("업그레이스 서버 갯수 : ", "서버갯수 입력", "")
'
'        If strNetCount = "" Then
'            strMsg = vbNo
'            Exit Sub
'        End If
'
'        gNetCount = strNetCount
'
'        Call S2SaveSetting("GNV", "SERVER CNT", "CNT", strNetCount)
'    Loop
'
'    strMsg = vbNo
'
    gNetCount = 1
    ReDim gNetDrive(gNetCount - 1)
    ReDim aryNetDrive(gNetCount - 1)
    ReDim gNetDriveChar(gNetCount - 1)
    ReDim aryNetDriveChar(gNetCount - 1)
    ReDim gClientPath(gNetCount - 1)
    ReDim aryClientPath(gNetCount - 1)

    ReDim gcServerPath(gNetCount - 1)

    For i = 0 To gNetCount - 1
        'Server의 Download 폴더
'        gNetDrive(i) = S2GetSetting("GNV", "SERVER PATH", "PATH" & i + 1, "")
        gNetDrive(i) = medGetINI("DownLoad", "Path", "C:\Schweitzer\Schweitzer.ini")
        If gNetDrive(i) = "" Then
            strMsg = MsgBox("업그레이드할 서버를 알수 없습니다. 지금 설정하시겠습니까?", vbQuestion + vbYesNo, "서버설정")

            If strMsg = vbNo Then Exit Sub

            aryNetDrive(i) = InputBox("업그레이스 서버 : ", "서버 입력", "")

            If aryNetDrive(i) = "" Then Exit Sub

            gNetDrive(i) = aryNetDrive(i)

'            Call S2SaveSetting("GNV", "SERVER PATH", "PATH" & i + 1, aryNetDrive(i))
            Call medSetINI("DownLoad", "Path", aryNetDrive(i), "C:\Schweitzer\Schweitzer.ini")
        End If
    Next
'
    For i = 0 To gNetCount - 1
        'Server의 Download 폴더
        gClientPath(i) = medGetINI("DownLoad", "Path", "C:\Schweitzer\Schweitzer.ini")
        If gClientPath(i) = "" Then
            strMsg = MsgBox("다운로드 경로를 알수 없습니다. 지금 설정하시겠습니까?", vbQuestion + vbYesNo, "경로설정")

            If strMsg = vbNo Then Exit Sub

            aryClientPath(i) = InputBox(gNetDrive(i) & " 다운로드 경로 : ", "경로 입력", "")

            If aryClientPath(i) = "" Then Exit Sub

            gClientPath(i) = aryClientPath(i)

            Call medSetINI("DownLoad", "Path", aryNetDrive(i), "C:\Schweitzer\Schweitzer.ini")
        End If
    Next
            
        
        
    On Error GoTo Errors
    strTmp = String(255, Chr$(0))
    Ret = GetLogicalDriveStrings(255, strTmp)
    
    aryTmp = Split(strTmp, Chr$(0))
    
    For i = LBound(aryTmp) To UBound(aryTmp)
        If aryTmp(i) = "" Then
            strLastDrive = aryTmp(i - 1)
            strLastDrive = Mid(strLastDrive, 1, 1)
            Exit For
        End If
    Next
    For i = 1 To gNetCount
        gNetDriveChar(i - 1) = Chr(Val(Asc(strLastDrive) + i)) & ":"
    Next

    
    For i = 0 To gNetCount - 1
        If Mid(gNetDriveChar(i), Len(gNetDriveChar(i))) <> "\" Then
            gcServerPath(i) = gNetDriveChar(i) & "\"
        Else
            gcServerPath(i) = gNetDriveChar(i)
        End If
    Next
    
    Exit Sub
    
Errors:

'    MsgBox Err.Description
End Sub

Private Sub ConNetDrive(ByVal Index As Integer, ByVal pindex As Long)

    If Index = 0 Then
        gcServerPath = gNetDrive
        If Mid(gcServerPath(pindex), Len(gcServerPath(pindex))) <> "\" Then gcServerPath(pindex) = gcServerPath(pindex) & "\"
        lblMessage.Caption = "업그레이드 서버에 연결하고 있습니다."
        lblMessage.Refresh
        
        '운영체제 체크 필요
        If Dir(gNetDrive(pindex), vbDirectory) = "" Then
        'If Not NetConnect(0, gNetDrive(pindex), pindex) Then      '네트웤 드라이브 연결
            MsgBox gNetDrive(pindex) & "   업그레이드 서버에 연결할 수 없습니다.", vbCritical, "서버 연결 오류"
            blnNetCon = False
            Call ExitProgram(pindex)
        End If
        blnNetCon = True
    Else
        lblMessage.Caption = "업그레이드 서버와 연결을 해제하고 있습니다."
        lblMessage.Refresh
        If Dir(gNetDrive(pindex), vbDirectory) = "" Then
       ' If Not NetConnect(1, gNetDrive(pindex), pindex) Then     '네트웤 드라이브 연결
            MsgBox "업그레이드 서버와 연결을 해제할 수 없습니다.", vbCritical, "서버 연결 오류"
             Call ExitProgram(pindex)
            End
        End If
        
        blnNetCon = False
    End If
End Sub


Private Function CopyNewVersion(ByVal pindex As Long) As Boolean
    
    Dim i As Integer
    Dim strAppName As String
    Dim strAppPath As String
    Dim ResumeCnt As Integer
    Dim blnCopy As Boolean
    Dim strSysDir As String

    On Error GoTo Err_Trap
    
    ResumeCnt = 0
    CopyNewVersion = True
    
    lblMessage.Caption = "파일을 복사하고 있습니다.."
    LogWrite  '##
    
    For i = 1 To FileCount
        DoEvents
        prgBar.value = prgBar.value + 1
        With gsFileInfo(i)
            If .flag And .FileNm <> App.EXEName Then   '새로운 버전
                Call CheckPath(.DestPath)       '폴더체크 및 생성
                strAppName = "C:\Schweitzer\" & .FileNm
                lblMessage.Caption = .FileNm & " 복사 중.."   '파일명
                LogWrite (strAppName & " 복사 중..")   '##
                
                If .FileExtend = "DLL" Or .FileExtend = "OCX" Then
'                    Call ExecCmd(gsSysPath & "Regsvr32.exe /u /s " & strAppName)     ', vbMinimizedNoFocus)
                    blnCopy = CopyFile(.SvrPath, .DestPath, .FileNm, .FileNm)
                    If Not blnCopy Then GoTo Err_Trap
'                    Call ExecCmd(gsSysPath & "Regsvr32.exe /s " & strAppName)     ', vbMinimizedNoFocus)
                    LogWrite (strAppName & "Registered") '##
                    DoEvents
                Else
                    MsgBox "1"
                    FileCopy .SvrPath & .FileNm, strAppName
                    MsgBox "2"
                    If .FileExtend = "EXE" And (.FileNm <> gsAppName) Then
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
    Call medSleep(2000)
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
    
'    lblMessage.Caption = "프로세스가 실행중인지 체크하고 있습니다."
    LogWrite  '##
    DoEvents
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            '프로세스 강제종료
            If KillFg Then
                myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                AppKill = TerminateProcess(myProcess, exitCode)
                Call CloseHandle(myProcess)
                LogWrite (myName & " killed...") '##
                DoEvents
            End If
            CheckApp = True
            Exit Do  '해당 프로세스가 발견되면 그만찾기
        End If
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


Private Sub GetFilesFromServer(ByVal strSvrPath As String, ByVal pindex As Long)

    Dim strFileNm As String
    Dim strExtend As String
    Dim strOldVersion As String
    Dim strOldDtTm As String
    Dim strSubDirs As String
    Dim arySubDirs() As String
    Dim i As Long
    
    On Error GoTo Err_Trap
    
    LogWrite  '##
    
    strSubDirs = GetDirectoriesFromServer(strSvrPath)
    arySubDirs = Split(strSubDirs, ":")
    For i = LBound(arySubDirs) To UBound(arySubDirs)
        If (GetAttr(strSvrPath & arySubDirs(i)) And vbDirectory) = vbDirectory Then
            Call GetFilesFromServer(strSvrPath & arySubDirs(i) & "\", pindex)
        End If
    Next
    
    strFileNm = Dir(strSvrPath)
    While (strFileNm <> "")
        DoEvents
        
        If UCase(strFileNm) <> "GETNEWVERSION.EXE" Then GoTo Skip
        FileCount = FileCount + 1
        ReDim Preserve gsFileInfo(FileCount)
        
        With gsFileInfo(FileCount)
            'strSvrPath
            .FileNm = strFileNm
            .FileSize = FileLen(strSvrPath & strFileNm)
            .FileDtTm = FileDateTime(strSvrPath & strFileNm)
            .FileVersion = GetFileVersion(strSvrPath & strFileNm)
            strExtend = UCase(medGetP(strFileNm, 2, "."))
            .FileExtend = strExtend
            .SvrPath = strSvrPath
'            Select Case strExtend
'                Case "EXE": .DestPath = gsAppPath
'            End Select
            
            If strExtend = "EXE" Then
                .DestPath = Replace(strSvrPath, gcServerPath(pindex), gClientPath(pindex))
            End If
            
            strOldVersion = GetFileVersion("C:\schweitzer\" & .FileNm)
            strOldDtTm = GetFileDateTime("C:\schweitzer\" & .FileNm)
                If .FileVersion <> vbNullString Then
                    If .FileVersion > strOldVersion Then   '버전비교
                        .flag = True
                        NewCount = NewCount + 1
                    Else
                        GoTo DateCompare
                    End If
                Else
DateCompare:
                    If .FileDtTm > strOldDtTm Then  '생성일 비교
                        .flag = True
                        NewCount = NewCount + 1
                    Else
                        .flag = False
                    End If
                End If
            LogWrite (.FileNm & vbTab & .FileVersion & vbTab & .FileDtTm & vbTab & .FileSize & vbTab & .DestPath & vbTab & .flag)   '##
        End With
Skip:
        strFileNm = Dir
        
    Wend
    
    DoEvents
'    If FileCount = 0 Then
    LogWrite (CStr(NewCount)) '##
    
    prgBar.Max = FileCount
    Exit Sub
    
Err_Trap:
    LogWrite (Err.Number & " : " & Err.Description)
    Resume Next
End Sub


Sub ExitProgram(ByVal pindex As Long)
    Call RestoreAll
'    If blnNetCon Then
'        Call ConNetDrive(1, pindex) '네트웤 드라이브 해제
'    End If
'    Close #Fd
    
    If blnNetCon Then
        Close #Fd
        Call ConNetDrive(1, pindex)     '네트웤 드라이브 해제
        
        If gNetCount = pindex + 1 Then
            Unload Me
            End
        End If
    Else
        Close #Fd
        Unload Me
        End
    End If
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
            
        lRC = VerInstallFile(intFlags, strShortSrcName, strShortDestName, strShortSrcDir, strShortDestDir, 0&, mstrVerTmpName, lpTmpNameLen)
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
            lblMessage.Caption = "대상 파일이 쓰기 금지되어 있습니다."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_FILEINUSE Then
            lblMessage.Caption = "대상 파일을 사용하고 있습니다. 다른 모든 응용 프로그램이 닫혀 있는지 확인하십시오."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_OUTOFSPACE Then
            lblMessage.Caption = "대상 드라이브의 공간이 부족합니다."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_ACCESSVIOLATION Then
            lblMessage.Caption = "파일을 복사하는 동안 액세스를 위반하였습니다."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_SHARINGVIOLATION Then
            lblMessage.Caption = "파일을 복사하는 동안 공유를 위반하였습니다."
            GoTo UnexpectedErr
        ElseIf lRC And VIF_OUTOFMEMORY Then
            lblMessage.Caption = "파일을 복사하는 데 사용할 메모리가 부족합니다."
            GoTo UnexpectedErr
        Else
            '
            ' For these cases, we generically report the error and do not install the file
            ' unless this is an SMS install; in which case we abort.
            '
            If lRC And VIF_CANNOTCREATE Then
                lblMessage.Caption = "임시 파일을 만들 수 없습니다."
            ElseIf lRC And VIF_CANNOTDELETE Then
                lblMessage.Caption = "기존 대상 파일을 삭제할 수 없습니다."
            ElseIf lRC And VIF_CANNOTRENAME Then
                lblMessage.Caption = "임시 파일 이름을 바꿀 수 없습니다."
            ElseIf lRC And VIF_CANNOTREADSRC Then
                lblMessage.Caption = "원본 파일을 읽을 수 없습니다."
            ElseIf lRC And VIF_CANNOTREADDST Then
                lblMessage.Caption = "대상 파일 속성을 읽을 수 없습니다."
            ElseIf lRC And VIF_BUFFTOOSMALL Then
                lblMessage.Caption = "내부 복사 오류입니다."
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
    Call ExitProgram(0)
End Function


