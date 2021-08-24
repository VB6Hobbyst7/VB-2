Attribute VB_Name = "ModGlobal"
Option Explicit

'- SMTP
Type typeSMTPItem
    Host As String
    ID As String
    PW As String
    Port As Long
    EMail As String
End Type

'- Mail
Type typeMailItem
    MailListCount As Long
    MailList() As String
End Type

'- SMS
Type typeSMSItem
    SMSListCount As Long
    SMSList() As String
End Type

Type typeFTPItem
    Name As String
    Host As String
    User As String
    Password As String
    Port As String
    RemoteDir As String
    LocalDir As String
    Passive As Integer
End Type


Type typeCodeItem
    Code As String
    Name As String
End Type

'- ���
Type typeSensorItem
    Code As String
    Name As String
    ActionStr As String
    ChannelCount As Integer
    Channel() As typeCodeItem
End Type

'- ������
Type typeSiteItem
    ID As Integer
    Code As String
    Name As String
    Default_tide As String
    FTP As typeFTPItem
    SensorCount As Integer
    SensorList() As typeSensorItem
End Type

'- DB
Type typeDBItem
    ID As String
    PW As String
    Source As String
End Type

Type typeSiteInfo
    SiteCount As Long
    SiteList() As typeSiteItem
End Type

Public Type typeConfigInfo
    SMTP As typeSMTPItem
    Mail As typeMailItem
    SMS As typeSMSItem
    Site As typeSiteInfo
    DB As typeDBItem
    TideConst As Integer
End Type

'�������� ���� ����.
Public ConfigAtt As typeConfigInfo

Public Type typeFileInfo
    FileName As String
    FileType As String
    FileSize As Long
    FileDate As String
End Type

Public Type typeFileList
    ListCount As Long
    List() As typeFileInfo
End Type
Public FileList As typeFileList


Public Type typeReceiveDataItem
    Obs_Date As String
    Referece As String
    Wind_Speed As String
    Max_Wind_Speed As String
    Wind_Direction As String
    Air_Temparature As String
    Air_Pressure As String
    Water_Hieght1 As String
    Salinity As String
    Water_Temparature As String
    Visibility As String
    Water_Height2 As String
End Type

Public Type typeReceiveDataInfo
    ItemCount As Long
    Item() As typeReceiveDataItem
End Type
Public ReceiveData As typeReceiveDataInfo

'---------------------------------------------
Public Type typeTSList
    ID As Integer
    Code As String
    Name As String
End Type

Public Type TypeMirosTideList
    Count As Integer
    TSList() As typeTSList
End Type

Public MTide_TSList As TypeMirosTideList
'---------------------------------------------

Public DBConn As ADODB.Connection
Public DBFlag As Boolean

Public DaeSanCode As String
Public DaeChungDoCode As String
Public bFTPConnected As Boolean

Public Enum SENSORSTATUS
    Normal = 1
    NOT_RECEIVED = 2
    NOT_INSTALL = 3
    Error = 4
End Enum

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lptitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

' API�Լ�
Public Declare Function WaitForSingleObject Lib "kernel32" _
              (ByVal hHandle As Long, _
               ByVal dwMilliseconds As Long) As Long

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
              (ByVal lpApplicationName As Long, _
               ByVal lpCommandLine As String, _
               ByVal lpProcessAttributes As Long, _
               ByVal lpThreadAttributes As Long, _
               ByVal bInheritHandles As Long, _
               ByVal dwCreationFlags As Long, _
               ByVal lpEnvironment As Long, _
               ByVal lpCurrentDriectory As Long, _
               lpStartupInfo As STARTUPINFO, _
               lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' API�Լ� ���
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&
Public Const STARTF_USESHOWWINDOW = &H1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWDEFAULT = 10

'-----------------Grid groupping--------------------
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8
Public Const SS_BORDER_TYPE_OUTLINE = 16
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BDM_CURRENT_ROW = 4

'------------------------------------------
' 1 - SDI��, ���۽� Ʈ���� ǥ��, �ּ�ȭ�� �����, ����Ŭ�� �Ǵ� ������Ŭ��->���̱�� �� Ȱ��ȭ
' 2 - MDI��, ���۽� Ʈ���� ǥ��, �ּ�ȭ�� �����, ����Ŭ�� �Ǵ� ������Ŭ��->���̱�� �� Ȱ��ȭ
' 3 - SDI��, �ּ�ȭ�� Ʈ���̷� �ø��� �����, ����Ŭ�� �Ǵ� ������Ŭ��->���̱�� �� Ȱ��ȭ
' 4 - MDI��, �ּ�ȭ�� Ʈ���̷� �ø��� �����, ����Ŭ�� �Ǵ� ������Ŭ��->���̱�� �� Ȱ��ȭ

Public gExeFlag As Integer ' ���� ����ɼ��� ���� �÷���
Public gIconFlag As Boolean ' ���� ������ ���¸� ���� �÷���

'===================================================================
'-- usefull SQLite
'-------------------------------------------------------------------
Public Const Start As Boolean = True
Public Declare Function QueryPerformanceFrequency& Lib "kernel32" (x@)
Public Declare Function QueryPerformanceCounter& Lib "kernel32" (x@)
''Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function HPTimer#()
Dim x@: Static Frq@
  If Frq = 0 Then QueryPerformanceFrequency Frq
  If QueryPerformanceCounter(x) Then HPTimer = x / Frq
End Function

Public Function Timing(Optional ByVal Start As Boolean) As String
Static T#
  If Start Then T = HPTimer: Exit Function
  Timing = " " & Format$((HPTimer - T) * 1000, "Standard") & "msec"
End Function

Public Function GetTextFromFile(FileName$) As String
Dim FNr&: FNr = FreeFile
  Open FileName For Binary Access Read As FNr
  GetTextFromFile = Space(LOF(FNr))
  Get FNr, , GetTextFromFile: Close FNr
End Function

Public Function FileExists(ByRef FileName As String) As Boolean
  On Error Resume Next
    FileExists = ((GetAttr(FileName) And vbDirectory) <> vbDirectory)
  Err.Clear
End Function
'==========================================================================

Public Sub RunAndWait(RunCommand As String)
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : �ý��� ���� 1��
'" 2. ���������� : �����Լ�
'" 3. ����       : Shell ����� �����Ų �� �ش� Shell�� ������ ����� ������ ����ϴ� �Լ��̴�.
'" 4. �Ķ����   : RunCommand : Shell ���� ���ڿ�
'" 5. �ۼ���     : �赿��
'" 6. �ۼ���     : 2007/08/30
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Dim vProc As PROCESS_INFORMATION
    Dim vStart As STARTUPINFO
    Dim vRv As Long

    vStart.cb = LenB(RunCommand)
    vStart.dwFlags = STARTF_USESHOWWINDOW
    vStart.wShowWindow = SW_SHOWDEFAULT 'SW_SHOWMAXIMIZED

    ' Process ����
    vRv = CreateProcess(0&, RunCommand, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, vStart, vProc)

    Screen.ActiveForm.MousePointer = 11
    DoEvents

    ' ���
    vRv = WaitForSingleObject(vProc.hProcess, INFINITE)

    Screen.ActiveForm.MousePointer = 0
    DoEvents
    ' Process ����
    vRv = CloseHandle(vProc.hProcess)
End Sub

Public Function GetDbConnection(ConnectionString As String) As Boolean
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : �ý��� ���� 1��
'" 2. ���������� : �����Լ�
'" 3. ����       : �����ͺ��̽� ���� ���ؼ� ������ �����Ѵ�.
'" 4. �Ķ����   : ConnectionString : ���� ���ڿ�
'" 5. �ۼ���     : �赿��
'" 6. �ۼ���     : 2007/08/30
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
On Error GoTo ErrorHandler
    DBFlag = False
    
    Set DBConn = New ADODB.Connection
    DBConn.Open ConnectionString
    
    If DBConn.State = adStateOpen Then
        GetDbConnection = True
        DBFlag = True
    Else
        GetDbConnection = False
        DBFlag = False
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
    If 1 = 2 Then
ErrorHandler:
        GetDbConnection = False
        DBFlag = False
    End If
End Function

Public Sub Disconnection2DB()
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : �ý��� ���� 1��
'" 2. ���������� : �����Լ�
'" 3. ����       : �����ͺ��̽� ������ �����Ѵ�.
'" 4. �Ķ����   :
'" 5. �ۼ���     : �赿��
'" 6. �ۼ���     : 2007/08/30
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    If DBConn.State = adStateOpen Then
        DBConn.Close
    End If
    
    If Not DBConn Is Nothing Then
        Set DBConn = Nothing
    End If
End Sub

Public Sub CalScale_10to60(ByVal dec As Double, dd As Integer, mm As Integer, ss As Double)
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : �ý��� ���� 1��
'" 2. ���������� : �����Լ�
'" 3. ����       : WGS-84���� ��/��/�ʷ� ��ȯ�Ѵ�.
'" 4. �Ķ����   :
'" 5. �ۼ���     : �赿��
'" 6. �ۼ���     : 2007/08/30
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Dim dDX As Double, mmX As Double, ssX As Double
    
    dDX = dec
    mmX = (dDX - Int(dDX)) * 60
    ssX = (mmX - Int(mmX)) * 60
    
    dd = Int(dDX)
    mm = Int(mmX)
    ss = Val(Format(ssX, "0.00")) 'Int(ssX + 0.5)

    If ss >= 60 Then
        mm = mm + 1
        ss = 0
    End If
    
    If mm >= 60 Then
        dd = dd + 1
        mm = 0
    End If
End Sub

Public Sub CalScale_60to10(dec As Double, ByVal dd As Integer, ByVal mm As Integer, ByVal ss As Double)
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : �ý��� ���� 1��
'" 2. ���������� : �����Լ�
'" 3. ����       : ��/��/�� ��ǥ���� WGS-84�� ��ȯ�Ѵ�.
'" 4. �Ķ����   :
'" 5. �ۼ���     : �赿��
'" 6. �ۼ���     : 2007/08/30
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Dim mmtemp As Double, sstemp As Double
    
    mmtemp = 1 / 60
    sstemp = 1 / 3600
    
    mmtemp = CDbl(mmtemp * mm)
    sstemp = CDbl(sstemp * ss)
    
    dec = CDbl(dd + mmtemp + sstemp)
End Sub

Public Sub LogWrite(LogString As String)
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : �ý��� ���� 1��
'" 2. ���������� : �����Լ�
'" 3. ����       : �α������� ����Ѵ�.
'" 4. �Ķ����   :
'" 5. �ۼ���     : �赿��
'" 6. �ۼ���     : 2007/08/30
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Dim iFn As Long
    
On Error GoTo ErrorHandler
    
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    If Dir(App.Path & "\Logs\Logs" & Format(Now, "yyyy-mm-dd") & ".txt", vbNormal) = "" Then
        iFn = FreeFile
        Open App.Path & "\Logs\Logs" & Format(Now, "yyyy-mm-dd") & ".txt" For Output As #iFn
            Print #iFn, Format(Now, "yyyy-mm-dd hh:nn:ss") & " >> " & LogString
        Close #iFn
    Else
        iFn = FreeFile
        Open App.Path & "\Logs\Logs" & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #iFn
            Print #iFn, Format(Now, "yyyy-mm-dd hh:nn:ss") & " >> " & LogString
        Close #iFn
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    Else
        Exit Sub
    End If
End Sub

Public Sub KillProcess(PName As String)
'--------------------------------------------------------------------------------------------
'������: ������(chiuoo@enjoyev.net)
'���糯¥: 2007.03.22
'����: KillProcess([�������ϸ�])
'        --> �������ϸ��� �۾�������->���μ����ǿ��� �̹��� �̸��� �ش��ϴ� �̸��� ����
'--------------------------------------------------------------------------------------------
    Dim pgm As String
    Dim wmi As Object
    Dim processes, process
    Dim sQuery As String
    
    pgm = PName
    Set wmi = GetObject("winmgmts:")
    sQuery = "select * from win32_process where name='" & pgm & "'"
    Set processes = wmi.execquery(sQuery)
    
    For Each process In processes
        process.Terminate
    Next
    
    If Not wmi Is Nothing Then Set wmi = Nothing
    If Not processes Is Nothing Then Set processes = Nothing
End Sub

Public Function ProcessCounts(PName As String) As Integer
'--------------------------------------------------------------------------------------------
'���糯¥: 2007.03.22
'����: ProcessCount([�������ϸ�])
'        --> �������ϸ��� �۾�������->���μ����ǿ��� �̹��� �̸��� �ش��ϴ� �̸��� ����
'--------------------------------------------------------------------------------------------
    Dim pgm As String
    Dim wmi As Object
    Dim processes, process
    Dim sQuery As String
    
    pgm = PName
    Set wmi = GetObject("winmgmts:")
    sQuery = "select * from win32_process where name='" & pgm & "'"
    Set processes = wmi.execquery(sQuery)
    
    ProcessCounts = processes.Count
    
    If Not wmi Is Nothing Then Set wmi = Nothing
    If Not processes Is Nothing Then Set processes = Nothing
End Function

