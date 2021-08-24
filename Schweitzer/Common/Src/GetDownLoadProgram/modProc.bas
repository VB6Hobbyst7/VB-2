Attribute VB_Name = "modProc"
Const MAX_PATH& = 260


Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    End Type

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Const WM_CLOSE = &H10
    Private Const GW_CHILD = 5
    Private Const GW_HWNDFIRST = 0
    Private Const GW_HWNDLAST = 1
    Private Const GW_HWNDNEXT = 2
    Private Const GW_HWNDPREV = 3
    Private Const GW_OWNER = 4
    Private Const SW_HIDE = 0
    Private Const SW_SHOWNORMAL = 1
    Private Const SW_NORMAL = 1
    Private Const SW_SHOWMINIMIZED = 2
    Private Const SW_SHOWMAXIMIZED = 3
    Private Const SW_MAXIMIZE = 3
    Private Const SW_SHOWNOACTIVATE = 4
    Private Const SW_SHOW = 5
    Private Const SW_MINIMIZE = 6
    Private Const SW_SHOWMINNOACTIVE = 7
    Private Const SW_SHOWNA = 8
    Private Const SW_RESTORE = 9
    
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
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

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
            hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
            lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
            lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
            ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
            ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
            lpStartupInfo As STARTUPINFO, lpProcessInformation As _
            PROCESS_INFORMATION) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&


Public Function MinimizeAllExcept(frmCaption As String) As Long


    Dim Length As Long
    Dim ListItem As String
    Dim CurrWnd As Long
    
    frmCaption = UCase(frmCaption)
    'Get hWnd of first item in task list
    CurrWnd = GetWindow(Screen.ActiveForm.hwnd, GW_HWNDFIRST)
    'Loop while the hWnd returned by GetWindow is valid.

    'MsgBox "app : " & App.EXEName

    While CurrWnd <> 0
        DoEvents
        'Get the length of task name identified by CurrWnd in the list.
        Length = GetWindowTextLength(CurrWnd)
        'Get task name of the task in the master list.
        ListItem = Space(Length + 1)
        Length = GetWindowText(CurrWnd, ListItem, Length + 1)
        'If there is a task name in the list, add the item to the list.


        If Length > 0 Then


            If IsWindowVisible(CurrWnd) <> 0 Then


                If InStr(UCase(ListItem), UCase(frmVersionCheck.Caption)) > 0 Or InStr(UCase(ListItem), UCase(App.Title)) > 0 Then
                    Call ShowWindow(CurrWnd, SW_SHOWNORMAL)
                    'Do nothing to form with specified caption
                Else
                    'See if it is an icon
                    If IsIconic(CurrWnd) <> 0 Then
                        'Do nothing here either
                    Else
                        'Ok this window is showing, minimize it
                        'MsgBox ListItem
                        MinimizeAllExcept = ShowWindow(CurrWnd, SW_SHOWMINIMIZED)
                        DoEvents
                        'Call WindowHandle(CurrWnd, 4)
                    End If

                End If

            Else
            End If

        End If

        'Get the next task list item in the master list.
        CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
    Wend

End Function



Public Function RestoreAll()

'    Dim Length As Integer
'    Dim ListItem As String
'    Dim CurrWnd
'    'Get hWnd of first item in task list
'    CurrWnd = GetWindow(Screen.ActiveForm.hwnd, GW_HWNDFIRST)
'    'Loop while the hWnd is valid.
'
'
'    While CurrWnd <> 0
'        'Get the length of task name of CurrWnd in the list.
'        Length = GetWindowTextLength(CurrWnd)
'        'Get task name of the task in the master list.
'        ListItem = Space(Length + 1)
'        Length = GetWindowText(CurrWnd, ListItem, Length + 1)
'        'If there is a task name in the list, add the item to the list.
'
'
'        If Length > 0 Then
'
'
'            If IsWindowVisible(CurrWnd) <> 0 Then
'                'Ok this window is showing, restore it
'                RestoreAll = ShowWindow(CurrWnd, SW_NORMAL)
'            End If
'
'        End If
'
'        'Get the next task list item
'        CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
'    Wend

End Function

Sub WindowHandle(win, cas As Long)


    'by storm
    'Case 0 = CloseWindow
    'Case 1 = Show Win
    'Case 2 = Hide Win
    'Case 3 = Max Win
    'Case 4 = Min Win


    Select Case cas
        Case 0:
        Dim X%
        X% = SendMessage(win, WM_CLOSE, 0, 0)
        Case 1:
        X = ShowWindow(win, SW_SHOW)
        Case 2:
        X = ShowWindow(win, SW_HIDE)
        Case 3:
        X = ShowWindow(win, SW_MAXIMIZE)
        Case 4:
        X = ShowWindow(win, SW_MINIMIZE)
    End Select

'any questions e-mail me at storm@n2.com
End Sub


Public Sub ExecCmd(CmdLine$)
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    
    ' Start the shelled application:
    Ret& = CreateProcessA(0&, CmdLine$, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    
    ' Wait for the shelled application to finish:
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret& = CloseHandle(proc.hProcess)
End Sub

Function GetCommandLine(Optional MaxArgs)
   '변수를 선언합니다.
   Dim C, CmdLine, CmdLnLen, InArg, i, NumArgs
   'MaxArgs가 제공되면 참조하십시오.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   '정확한 크기의 배열을 만듭니다.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   '명령줄 인수를 가져옵니다.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   '동시에 한 문자 명령줄을 통과합니다.
   For i = 1 To CmdLnLen
      C = Mid(CmdLine, i, 1)
      '공백 또는 탭 검사.
      If (C <> " " And C <> vbTab) Then
         '공백이나 탭에 관계 없습니다.
         '만일 인수가 준비되면 테스트합니다.
         If Not InArg Then
         '새로운 인수 시작.
         '너무 많은 인수에 대해 테스트합니다.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         '현재 인수에 문자를 연결합니다.
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         '공백이나 탭을 발견.
         'InArg는 False로 설정합니다.
         InArg = False
      End If
   Next i
   '가지고 있는 인수에 충분한 배열 재조정.
   'ReDim Preserve ArgArray(NumArgs)
   '함수 이름에 배열을 반환합니다.
   GetCommandLine = ArgArray()
End Function

