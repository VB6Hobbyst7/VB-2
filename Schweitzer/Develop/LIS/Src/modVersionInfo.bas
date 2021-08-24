Attribute VB_Name = "modVersionInfo"
Option Explicit

'INI File에Data를 쓰는 API Function
'Public Declare Function WritePrivateProfileString Lib "kernel32" _
'              Alias "WritePrivateProfileStringA" _
'             (ByVal lpApplicationName As String, _
'              ByVal lpKeyName As Any, _
'              ByVal lpString As Any, _
'              ByVal lpFileName As String) As Long
'
'' INI File에서 Data를 읽는 API Function
'Public Declare Function GetPrivateProfileString Lib "kernel32" _
'              Alias "GetPrivateProfileStringA" _
'             (ByVal lpApplicationName As String, _
'              ByVal lpKeyName As Any, _
'              ByVal lpDefault As String, _
'              ByVal lpReturnedString As String, _
'              ByVal nSize As Long, _
'              ByVal lpFileName As String) As Long
'
'
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Public Const EM_GETSEL = &HB0
'Public Const EM_SETSEL = &HB1
'Public Const EM_GETLINECOUNT = &HBA
'Public Const EM_LINEINDEX = &HBB
'Public Const EM_LINELENGTH = &HC1
