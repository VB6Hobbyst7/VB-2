VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Set ObjSysInfo(ByVal vData As Object)
    Set mvarObjSysInfo = vData
End Property

Public Property Get ObjSysInfo() As Object
    Set ObjSysInfo = mvarObjSysInfo
End Property

Public Property Set ObjMyUser(ByVal vData As Object)
    Set mvarObjMyUser = vData
End Property

Public Property Get ObjMyUser() As Object
    Set ObjMyUser = mvarObjMyUser
End Property

Public Property Set MainFrm(ByVal vData As Object)
    Set mvarMainFrm = vData
End Property

Public Property Get MainFrm() As Object
    Set MainFrm = mvarMainFrm
End Property

Public Function INIPath() As String
    INIPath = InstallDir & "Schweitzer.ini"
End Function

Public Function InstallDir() As String
    Dim tmpDir As String
    
    tmpDir = GetSetting("Schweitzer2000", "InstallDir", "InstallDir", "")
    If tmpDir <> "" Then
        If Mid(tmpDir, Len(tmpDir), 1) = "\" Then
            tmpDir = tmpDir
        Else
            tmpDir = tmpDir & "\"
        End If
    End If

    InstallDir = tmpDir
End Function

'   DataBase 관련
'
Public Property Get DBConn() As Connection
    Set DBConn = mvarDBConn
End Property

Public Property Set DBConn(ByVal vData As Connection)
    Set mvarDBConn = vData
End Property

Public Property Get IsDBOpen() As Boolean
    IsDBOpen = mvarIsDBOpen
End Property

Public Property Let IsDBOpen(ByVal vData As Boolean)
    mvarIsDBOpen = vData
End Property

Public Function DBConnect() As Long
    
    Dim lngRetry As Long
    
On Error GoTo ConnectError
    
    lngRetry = 0
    IsDBOpen = False

DoAgain:
    lngRetry = lngRetry + 1
    If ConnectDatabase Then
        IsDBOpen = True
        Call GetSystemDate
    Else
        If lngRetry < 3 Then GoTo DoAgain  '연결이 안될 경우 3번까지 재시도..
    End If
    Exit Function

ConnectError:
   
    If Not (mvarDBConn Is Null) Then
        MsgBox mvarDBConn.Errors(0).Number & " : " & mvarDBConn.Errors(0).Description
        mvarDBConn.Close
        Set mvarDBConn = Nothing
    End If
    
    MsgBox "Database 연결에 실패하였습니다. 잠시 후 다시 시도하거나 전산 담당에게 문의하십시오.", vbCritical
    IsDBOpen = False

End Function

Private Function ConnectDatabase() As Boolean
    Dim strMsg As String
    Dim tmpConnectionString As String
    
    Set mvarDBConn = Nothing
    Set mvarDBConn = New Connection
    
    Select Case ObjSysInfo.DBType
        Case 0
            tmpConnectionString = "Provider=msdaora;" & _
                                  "Data Source=" & ObjSysInfo.DatabaseNm & ";" & _
                                  "User Id=" & ObjSysInfo.DBLoginId & ";" & _
                                  "Password=" & ObjSysInfo.DBPassword
        Case 1
            tmpConnectionString = "Provider=Sybase.ASEOLEDBProvider;" & _
                                  "Srvr=" & ObjSysInfo.ServerNm & ",5000;" & _
                                  "Catalog=" & ObjSysInfo.DatabaseNm & ";" & _
                                  "User Id=" & ObjSysInfo.DBLoginId & ";" & _
                                  "Password=" & ObjSysInfo.DBPassword
        Case 2
            tmpConnectionString = "Provider=sqloledb;" & _
                                  "Data Source=" & ObjSysInfo.ServerNm & ";" & _
                                  "Initial Catalog=" & ObjSysInfo.DatabaseNm & ";" & _
                                  "User Id=" & ObjSysInfo.DBLoginId & ";" & _
                                  "Password=" & ObjSysInfo.DBPassword
        Case Else
            tmpConnectionString = ""
    End Select
    
    mvarDBConn.ConnectionString = tmpConnectionString
    
    On Error GoTo ConnectFail

    mvarDBConn.CursorLocation = adUseClient
'    mvarDBConn.ConnectionTimeout = 30
    mvarDBConn.Open mvarDBConn.ConnectionString
    ConnectDatabase = True
    Exit Function
    
ConnectFail:
    ConnectDatabase = False
End Function

Public Function GetSystemDate() As Date
    Dim Rs As Recordset
    Dim tmpGetSysSQL As String
    
    Set Rs = New Recordset
    
    Select Case ObjSysInfo.DBType
        Case 0
            tmpGetSysSQL = "select sysdate from dual"
        Case 1
            tmpGetSysSQL = "select convert(char(10),getdate(),111) + ' ' + convert(char(8),getdate(),108) sysdate" 'syb
        Case 2
            tmpGetSysSQL = "select getdate() sysdate" 'ms7
        Case Else
            tmpGetSysSQL = ""
    End Select
    
    On Error GoTo ErrTrap
    
    Rs.Open tmpGetSysSQL, mvarDBConn
    
    If Rs.EOF = False Then
        Date = Rs.Fields("sysdate").Value & ""
        Time = Rs.Fields("sysdate").Value & ""
        
        GetSystemDate = Rs.Fields("sysdate").Value & ""
    End If
    
    Set Rs = Nothing
    Exit Function
    
ErrTrap:
    Set Rs = Nothing
    
    MsgBox "시스템 날짜를 가져오는 도중 오류가 발생하였습니다." & vbNewLine & vbNewLine & _
           "오류 정보 : " & Err.Description & vbNewLine, vbCritical
End Function

Public Sub DbClose()
    On Error Resume Next
    
    If IsDBOpen Then mvarDBConn.Close
    Set mvarDBConn = Nothing
End Sub
