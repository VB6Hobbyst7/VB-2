Attribute VB_Name = "modFILETIME"
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As Long, lpLastWriteTime As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
 

Public Function GetFileCreateTime(FileName As String) As String
    Dim FH As Long, FT As FILETIME, ST As SYSTEMTIME
    FH = CreateFile(FileName, &H80000000, &H1 Or &H2, ByVal 0&, 3, &H80, 0)
    If GetFileTime(FH, FT, ByVal 0&, ByVal 0&) Then
        Call FileTimeToSystemTime(FT, ST)
        GetFileCreateTime = CStr(ST.wYear) & " ³â " & CStr(ST.wMonth) & " ¿ù " & CStr(ST.wDay) & " ÀÏ "
    End If
    Call CloseHandle(FH)
End Function
