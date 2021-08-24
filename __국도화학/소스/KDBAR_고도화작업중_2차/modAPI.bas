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


'Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'
'
'
'    'Click this button if the project name and the compiled file
'    'name are the same.
'    Dim strFileName As String
'    Dim lngCount As Long
'
'    strFileName = String(255, 0)
'    lngCount = GetModuleFileName(App.hInstance, "Get_TempMaster", 255)
'    strFileName = LEFT(strFileName, lngCount)
'    If UCase(Right(strFileName, 7)) <> "VB5.EXE" Then
'        MsgBox "Compiled Version"
'    Else
'        MsgBox "IDE Version"
'    End If
'
