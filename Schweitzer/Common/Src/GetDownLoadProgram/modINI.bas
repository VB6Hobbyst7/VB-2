Attribute VB_Name = "modINI"
Option Explicit

'# Ini를 다루는 api
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long




' --------------------------------------------------------
' Registry에서 정보를 읽는 부분을 INI로 변경하기 위한 함수
' GetSetting과 SaveSetting과 사용법을 동일하게 만들었다.
' --------------------------------------------------------

Public Function S2SaveSetting(pAppName As String, pSection As String, pKey As String, pSetting As String) As Long
    S2SaveSetting = WritePrivateProfileString(pSection, pKey, pSetting, App.Path & "\" & pAppName & ".ini")
End Function

Public Function S2GetSetting(pAppName As String, pSection As String, pKey As String, Optional pDefault As String = "") As String
    Dim P As String
    
    P = Space$(256)
    Call GetPrivateProfileString(pSection, pKey, pDefault, P, 256, App.Path & "\" & pAppName & ".ini")
    S2GetSetting = Mid(Trim(P), 1, Len(Trim(P)) - 1)
End Function

Public Function medGetINI(ByVal section As String, ByVal key As String, ByVal pathname As String, Optional ByVal defaultvalue As String = "") As String
    Dim P As String
    
    P = Space$(256)
    Call GetPrivateProfileString(section, key, defaultvalue, P, 256, pathname)
    medGetINI = Mid(Trim(P), 1, Len(Trim(P)) - 1)
End Function

Public Sub medSetINI(ByVal section As String, ByVal key As String, ByVal value As String, ByVal pathname As String)
    Call WritePrivateProfileString(section, key, value, pathname)
End Sub

