Attribute VB_Name = "modAPI"
'AlwaysOn 관련 정의
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2   'Not Always top
Public Const HWND_TOPMOST = -1  'Always top
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1


Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Private Const CHUNK_SIZE& = 4096&
Private Const CP_UTF8 As Long = 65001
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                    (ByVal lpApplicationName As String _
                   , ByVal lpKeyName As Any _
                   , ByVal lpString As Any _
                   , ByVal lplFileName As String) As Long
    
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                    (ByVal lpApplicationName As String _
                   , ByVal lpKeyName As Any _
                   , ByVal lpDefault As String _
                   , ByVal lpReturnedString As String _
                   , ByVal nSize As Long _
                   , ByVal lpFileName As String) As Long


