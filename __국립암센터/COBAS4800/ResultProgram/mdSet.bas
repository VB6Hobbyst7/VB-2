Attribute VB_Name = "mdSet"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public gFilePath As String
Public gIP As String
Public gPort As String
Public gAutoSend As String

Public res As Integer


Public Function GetSetUp() As Boolean
    Dim db_tmp As String * 100

    GetSetUp = False
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "Port", "", db_tmp, 20, App.Path & "\Interface.ini")
    FrmResult.txtTemp = Trim(db_tmp)
    gPort = Trim(FrmResult.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "IP", "", db_tmp, 20, App.Path & "\Interface.ini")
    FrmResult.txtTemp = Trim(db_tmp)
    gIP = Trim(FrmResult.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "FilePath", "", db_tmp, 200, App.Path & "\Interface.ini")
    FrmResult.txtTemp = Trim(db_tmp)
    gFilePath = Trim(FrmResult.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "AutoSend", "", db_tmp, 200, App.Path & "\Interface.ini")
    FrmResult.txtTemp = Trim(db_tmp)
    gAutoSend = Trim(FrmResult.txtTemp)
    
    GetSetUp = True
    
End Function

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
End Sub




