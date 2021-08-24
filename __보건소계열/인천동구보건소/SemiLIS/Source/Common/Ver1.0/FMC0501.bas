Attribute VB_Name = "FMC0501"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDLAST = 1

Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_HIDE = 0
Public Const WM_DESTROY = &H2
Public Const WM_QUIT = &H12

Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_HIDEWINDOW = &H80

'*--------------------------------------------------------
'*      특정 윈도우의 핸들 얻기 (윈도의 Caption 값을 이용
'*      para : 윈도우 Caption 값
'*----------------------------------------------------------
Public Function App_GetMainWindowHandle(ByVal para As String) As Long
    Dim hwnd            As Long
    Dim sWindowText  As String
    Dim App_ret     As Long
    
    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)

    Do Until hwnd = 0
        
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
        
        sWindowText = String(255, 0)
        App_ret = GetWindowText(hwnd, sWindowText, 255)
        sWindowText = LeftH$(sWindowText, App_ret)
        
        'MsgBox sWindowText
        
        If (sWindowText = para) Then
            App_GetMainWindowHandle = hwnd
            Exit Function
        End If
    Loop

    App_GetMainWindowHandle = 0

End Function

Public Function OnlyOneFormPerMenu(ByVal hWndFGM0101 As Long, ByVal para As String) As Integer
    Dim hwnd           As Long
    Dim sWindowText  As String
    Dim App_ret     As Long

    hwnd = GetWindow(hWndFGM0101, GW_HWNDNEXT)

    Do Until hwnd = 0
        
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
        
        sWindowText = String(255, 0)
        App_ret = GetWindowText(hwnd, sWindowText, 255)
        sWindowText = LeftH$(sWindowText, App_ret)

        If (sWindowText Like "*" & para & "*") Then
            App_ret = CloseWindow(hwnd)
            
            If App_ret = 0 Then     'Destroy 실패
                OnlyOneFormPerMenu = 0
            Else        'Destroy 성공
                OnlyOneFormPerMenu = 1
            End If
            
            Exit Function
        End If
    Loop

    OnlyOneFormPerMenu = 1
End Function

Public Function Get_Handle_Of_SomeCap_In_MDI(ByVal hWndFGM0101 As Long, ByVal para As String) As Long
    Dim hwnd           As Long
    Dim sWindowText  As String
    Dim App_ret     As Long
    
    hwnd = hWndFGM0101

    Do Until hwnd = 0
        
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
        
        sWindowText = String(255, 0)
        App_ret = GetWindowText(hwnd, sWindowText, 255)
        sWindowText = LeftH$(sWindowText, App_ret)
        
'        MsgBox sWindowText
        
        If (sWindowText Like "*" & para & "*") Then
            Get_Handle_Of_SomeCap_In_MDI = hwnd
            
            Exit Function
        End If
    Loop

    Get_Handle_Of_SomeCap_In_MDI = 0
End Function

Public Function Get_Handle_Of_SomeCap(ByVal para As String) As Long
    Dim hwnd            As Long
    Dim sWindowText  As String
    Dim App_ret     As Long
    
    hwnd = FindWindow(vbNullString, para)
    
    If hwnd = vbNull Then
        Get_Handle_Of_SomeCap = 0
    Else
        Get_Handle_Of_SomeCap = hwnd
    End If
End Function

