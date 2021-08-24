Attribute VB_Name = "modS2CONHOS"
Option Explicit

Const C_PATHNAME As String = "CONHOS.s2"

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function ReadINI(ByVal pSection As String, ByVal pKey As String, ByVal pDefault As String) As String
    Dim P As String
    
    P = Space$(256)
    Call GetPrivateProfileString(pSection, pKey, pDefault, P, 256, App.Path & "\" & C_PATHNAME)
    ReadINI = Mid(Trim(P), 1, Len(Trim(P)) - 1)
End Function

