Attribute VB_Name = "modTLA"
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

Public Cn_Ser   As ADODB.Connection
Public RS_Ser   As ADODB.Recordset
Public SQL      As String

Public gResresh   As Long

Public gGRPCD       As String
Public gEQPCD       As String
Public gWIDTH       As Integer

Public gIP          As String
Public gDB          As String
Public gID          As String
Public gPW          As String


Public Const colHeader = 4




Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
        
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = strName
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Output As FilNum
    Print #FilNum, argSQL
    Close FilNum
    
End Sub
