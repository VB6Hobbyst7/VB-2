Attribute VB_Name = "modGlobal"
Option Explicit

Public mvarDBConn           As Connection
Public mvarObjMyUser        As Object
Public mvarObjSysInfo       As Object
Public mvarIsDBOpen         As Boolean
Public mvarMainFrm          As Object

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

