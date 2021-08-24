Attribute VB_Name = "modSysinfo"
Option Explicit

Public INIPath As String

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Enum Enum_OperatingPlatform
  Platform_Windows_32 = 0
  Platform_Windows_95_98_ME = 1
  Platform_Windows_NT_2K_XP = 2
End Enum

Public Enum Enum_OperatingSystem
  System_Windows_32 = 0
  System_Windows_95 = 1
  System_Windows_98 = 2
  System_Windows_ME = 3
  System_Windows_NT = 4
  System_Windows_2K = 5
  System_Windows_XP = 6
End Enum

Public Function OperatingPlatform() As Enum_OperatingPlatform
    Dim lpVersionInformation As OSVERSIONINFO
    lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
    Call GetVersionExA(lpVersionInformation)
    OperatingPlatform = lpVersionInformation.dwPlatformId
End Function

Public Function OperatingSystem() As Enum_OperatingSystem
    Dim lpVersionInformation As OSVERSIONINFO
    lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
    Call GetVersionExA(lpVersionInformation)
    If (lpVersionInformation.dwPlatformId = Platform_Windows_32) Then
        OperatingSystem = System_Windows_32
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 0) Then
        OperatingSystem = System_Windows_95
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 10) Then
        OperatingSystem = System_Windows_98
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 90) Then
        OperatingSystem = System_Windows_ME
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion < 5) Then
        OperatingSystem = System_Windows_NT
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 0) Then
        OperatingSystem = System_Windows_2K
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 1) Then
        OperatingSystem = System_Windows_XP
    End If
End Function


