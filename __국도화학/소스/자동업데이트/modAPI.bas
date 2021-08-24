Attribute VB_Name = "modAPI"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                    (ByVal lpApplicationName As String _
                   , ByVal lpKeyName As Any _
                   , ByVal lpString As Any _
                   , ByVal lplFileName As String) As Long


