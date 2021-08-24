Attribute VB_Name = "subAppConst"
'-------------------------------
'   Module : subAppConst(subAppConst.bas)
'
'   최근수정일 : 1999-07-27
'   최근수정자 : 김희정
'-------------------------------

Option Explicit

'-------------------------------------*
' Kernel Lib function declare         *
'-------------------------------------*
Declare Function GetCurrentTask Lib "kernel32" () As Long
Declare Function GetModuleUsage Lib "kernel32" (ByVal hModule As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
'-- 한글문자열 길이Check(2000.012.07 K.Y.G추가)
Declare Function AStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

'-------------------------------------*
' User Lib function declare           *
'-------------------------------------*
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function WNetAddConnection Lib "user32" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Declare Function WNetCancelConnection Lib "user32" (ByVal lpszName As String, ByVal bForce As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal pnt As Any) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public App_hWnd As Long

'------------------
'   Set WIndow Position
'   1999-08-27  김희정
'------------------
' SetWindowPos Flags
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
'-----------
'   Get the Network User Name from the Win32 API(mpr.dll)
'-----------
Private Declare Function w32_WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" ( _
  ByVal lpszLocalName As String, _
  ByVal lpszUserName As String, _
  lpcchBuffer As Long) As Long

Private Const NO_ERROR = 0

'-------------
'   Registry관련 API [1999-07-27 김희정 추가]
'                    [1999.12.15 유은자 RegDeleteKey 추가]
'-------------
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const REG_NONE = 0
Public Const REG_VALUETYPE_SZ = 1
Public Const REG_DWORD = 4
Public Const REG_SUCCESS = 0&
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const SYNCHRONIZE = &H100000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
    
Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type



'*----------------------------------------------------
'*  App 함수을 이용하여 얻은 String 값의 쓰레기값 절삭
'*----------------------------------------------------------
Function App_DelString(Code As String) As Integer
    Dim strLen  As Integer

 
    strLen = InStr(Code, Chr(0)) 'Right$(code, 1))
    If strLen <> 0 Then Code = Left$(Code, strLen - 1)

    App_DelString = True

End Function

'*--------------------------------------------------------
'*      특정 윈도우의 핸들 얻기 (윈도의 Caption 값을 이용
'*      para : 윈도우 Caption 값
'*----------------------------------------------------------
Function App_GetWindowHendle(Para As String) As Long
    
    Dim hwnd    As Long
    Dim Windowtext  As String

    Dim App_Ret     As Long

    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)

    Do Until hwnd = 0
        
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
        
        Windowtext = String(255, 0)
        App_Ret = GetWindowText(hwnd, Windowtext, 255)
        Windowtext = Left$(Windowtext, App_Ret)
        
        'MsgBox Windowtext

        If (Windowtext Like "*" + Para + "*") Then
            App_GetWindowHendle = hwnd
            Exit Function
        End If

    Loop

    App_GetWindowHendle = 0

End Function

'*--------------------------------------------------------
'*      특정 윈도우의 핸들 얻기 (윈도의 Caption 값을 이용
'*      para : 윈도우 Caption 값
'*----------------------------------------------------------
Function Get_Me(Para As String, Me_Hwnd As Long) As Integer
    
    Dim hwnd    As Integer
    Dim Windowtext  As String

    Dim App_Ret     As Integer

    Get_Me = False
    
    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)

    Do Until hwnd = 0
        
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
        
        Windowtext = String(255, 0)
        App_Ret = GetWindowText(hwnd, Windowtext, 255)
        Windowtext = Left$(Windowtext, App_Ret)
        
        'MsgBox Windowtext

        If (Windowtext Like "*" + Para + "*") And hwnd <> Me_Hwnd Then
            Get_Me = True
            Exit Function
        End If

    Loop

End Function

Public Function CStringToVBString(psCString As String) As String
' **********
' Purpose:     Convert a C string to a VB string
' Parameters:  (Input Only)
'  psCString - the C string to convert
' Returns:     The converted VB string
' Notes:
'  Returns everything to the left of the first Null character
' **********

   Dim sReturn As String
   Dim iNullCharPos As Integer
   
   iNullCharPos = InStr(psCString, vbNullChar)
   
   If iNullCharPos > 0 Then
      ' return everything left of the null
      sReturn = Left(psCString, iNullCharPos - 1)
   Else
      ' no null, return the original string
      sReturn = psCString
   End If

   CStringToVBString = sReturn
   
End Function
Public Function Get_Term_Id() As String
' **********
' Purpose:     Retrieve the network user name
' Paramters:   None
' Returns:     The indicated name
' Notes:
'  A zero-length string is returned if the function fails
' **********

  Dim lpUserName As String
  Dim lpnLength As Long
  Dim lResult As Long
  
  lpnLength = 256
  lpUserName = Space(lpnLength)
  
  lResult = w32_WNetGetUser(vbNullString, lpUserName, lpnLength)
  
  If lResult = NO_ERROR Then
    Get_Term_Id = CStringToVBString(lpUserName)
  Else
    Get_Term_Id = ""
  End If

End Function

