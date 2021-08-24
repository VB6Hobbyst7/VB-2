Attribute VB_Name = "modDBConnection"
Option Explicit

'########## 데이터베이스 관련 ############################
'DB Type    [ 1 : 오라클, 2 : MSSQL, 3 : Postgres, 4 : Access, 99 : 없음 ]
Public gDBTYPE      As String

'DB Conn    [ 1 : LOCAL, 2 : ERP ]
Public gDBCONN      As String

'로컬데이터베이스
Type LocalDBParameter
    PATH    As String
    UID     As String
    PWD     As String
End Type

Public gLocalDB     As LocalDBParameter

'데이터베이스(오라클)
Type OracleDBParameter
    SID     As String
    UID     As String
    PWD     As String
End Type
Public gORADB        As OracleDBParameter

'데이터베이스(오라클2)
Type OracleDBParameter2
    SID     As String
    UID     As String
    PWD     As String
End Type
Public gORADB2       As OracleDBParameter2


'MS JetDB
Global AdoCn_Local          As ADODB.Connection
Global AdoRs_Local          As ADODB.Recordset
Global AdoCmd_Local         As ADODB.Command
Global AdoParm_Local        As ADODB.Parameter

'공통
Global AdoCn                As ADODB.Connection
Global AdoRs                As ADODB.Recordset
Global AdoCmd               As ADODB.Command
Global AdoParm              As ADODB.Parameter

Global AdoCn2               As ADODB.Connection
Global AdoRs2               As ADODB.Recordset
Global AdoCmd2              As ADODB.Command
Global AdoParm2             As ADODB.Parameter

Public cn_Local_Flag        As Boolean
Public cn_Server_Flag       As Boolean

'########## 데이터베이스 관련 ############################
'-- MDB 연결
Public Function DbConnect_Local() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    DB_Name = gLocalDB.PATH
    UserName = gLocalDB.UID
    Password = gLocalDB.PWD
    
    If DB_Name = "" Then
        DB_Name = App.PATH & "\Database\" & gERP & ".mdb"
    End If
    
    If UserName = "" Then
        UserName = "admin"
    End If
    
    If (DB_Name = "") Or (UserName = "") Then
        DbConnect_Local = False
        Set AdoCn_Local = Nothing
        Exit Function
    End If

    With AdoCn
        .ConnectionTimeout = 25
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Mode").Value = adModeReadWrite
        .Properties("Persist Security Info").Value = False
        .Properties("Data Source").Value = DB_Name
        .Properties("User ID").Value = UserName
        .Properties("Jet OLEDB:Database Password").Value = Password
        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
        
        .Open
        
    End With

    Screen.MousePointer = vbDefault
    DbConnect_Local = True
 Exit Function

ConnectError:

    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn_Local.State <> adStateOpen Then
        DbConnect_Local = False
        Set AdoCn_Local = Nothing
    End If

End Function



'-- ORACLE 연결
Public Function DbConnect_ORACLE() As Boolean
    Dim ServerName As String
    Dim UserName As String
    Dim Password As String
    Dim blnWinNTAuth As Boolean

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    ServerName = gORADB.SID
    UserName = gORADB.UID
    Password = gORADB.PWD

    If (ServerName = "") Then
        DbConnect_ORACLE = False
        Set AdoCn = Nothing
        Exit Function
    End If
    
    With AdoCn
        .ConnectionTimeout = 25
        .Provider = "MSDAORA.1"
        .Properties("Data Source").Value = ServerName
        .Properties("Persist Security Info") = True
        
        If blnWinNTAuth = True Then
            .Properties("Integrated Security").Value = "SSPI"
        Else
            .Properties("User ID").Value = UserName
            .Properties("Password").Value = Password
        End If
        
        Screen.MousePointer = vbHourglass
        .Open
    End With
    
    Screen.MousePointer = vbDefault
    DbConnect_ORACLE = True
 
 Exit Function

ConnectError:
    Screen.MousePointer = vbDefault
    MsgBox " Error No. : " & Err.Number & vbCrLf & _
            " Description : " & Err.Description & vbCrLf & _
            " Source : " & Err.Source & vbCrLf & vbCrLf, vbOKOnly + vbCritical, "Database 연결실패"
            
    If AdoCn.State <> adStateOpen Then
        DbConnect_ORACLE = False
        Set AdoCn = Nothing
    End If

End Function

'-- ORACLE 연결2
Public Function DbConnect_ORACLE2() As Boolean
    Dim ServerName As String
    Dim UserName As String
    Dim Password As String
    Dim blnWinNTAuth As Boolean

    Set AdoCn2 = New ADODB.Connection

On Error GoTo ConnectError

    ServerName = gORADB2.SID
    UserName = gORADB2.UID
    Password = gORADB2.PWD

    If (ServerName = "") Then
        DbConnect_ORACLE2 = False
        Set AdoCn2 = Nothing
        Exit Function
    End If
    
    With AdoCn2
        .ConnectionTimeout = 25
        .Provider = "MSDAORA.1"
        .Properties("Data Source").Value = ServerName
        .Properties("Persist Security Info") = True
        
        If blnWinNTAuth = True Then
            .Properties("Integrated Security").Value = "SSPI"
        Else
            .Properties("User ID").Value = UserName
            .Properties("Password").Value = Password
        End If
        
        Screen.MousePointer = vbHourglass
        .Open
    End With
    
    Screen.MousePointer = vbDefault
    DbConnect_ORACLE2 = True
 
 Exit Function

ConnectError:
    Screen.MousePointer = vbDefault
    MsgBox " Error No. : " & Err.Number & vbCrLf & _
            " Description : " & Err.Description & vbCrLf & _
            " Source : " & Err.Source & vbCrLf & vbCrLf, vbOKOnly + vbCritical, "Database2 연결실패"
            
    If AdoCn.State <> adStateOpen Then
        DbConnect_ORACLE2 = False
        Set AdoCn = Nothing
    End If

End Function


Public Sub DisConnect_Local()

    If cn_Local_Flag = True Then
        AdoCn_Local.Close
    End If
    
End Sub

Public Sub DisConnect_Server()

    If cn_Server_Flag = True Then
        AdoCn.Close
    End If
    
End Sub

'Check Database with oracle server-reset
Public Function DBExists(ByVal AdoCn As ADODB.Connection, ByVal strSql As String) As Boolean
On Error GoTo DBOpenRsError
    Dim recAdoRs As New ADODB.Recordset
    
    Set recAdoRs = New ADODB.Recordset
    With recAdoRs
        .CursorLocation = adUseClient
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
        .Open
    End With
    
    If recAdoRs.EOF = False Then
        DBExists = True
    Else
        DBExists = False
    End If
    
    recAdoRs.Close
    Set recAdoRs = Nothing
    
    Exit Function

DBOpenRsError:
    recAdoRs.Close
    Set recAdoRs = Nothing
    DBExists = False

    Call DBErrorSet(AdoCn, strSql, "DBExists")
End Function


'execute the sql command
Public Function DBExec(ByVal AdoCn As ADODB.Connection, ByVal strSql As String) As Boolean
On Error GoTo DBExecError
    'lock table in row share mode
    With AdoCn
        .BeginTrans
        .Execute strSql
        .CommitTrans
    End With

    DBExec = True
Exit Function
    
DBExecError:
    AdoCn.RollbackTrans
    DBExec = False
    Call DBErrorSet(AdoCn, strSql, "DBExec")
End Function


'Open Sql server-recordset based on the sql query
Public Function GetRecordset(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                            ByVal AdoRs As ADODB.Recordset, _
                            Optional Call_Name As String, _
                            Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                            Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                            Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean
On Error GoTo DBOpenRsError
    With AdoRs
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    GetRecordset = True
Exit Function

DBOpenRsError:
    Set AdoRs = Nothing
    GetRecordset = False
    Call DBErrorSet(AdoCn, strSql, "GetRecordset")
End Function

Public Function GetRecordset2(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                            ByVal AdoRs As ADODB.Recordset, _
                            Optional Call_Name As String, _
                            Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                            Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                            Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean
On Error GoTo DBOpenRsError
    With AdoRs
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn2
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    GetRecordset2 = True
Exit Function

DBOpenRsError:
    Set AdoRs2 = Nothing
    GetRecordset2 = False
    Call DBErrorSet(AdoCn2, strSql, "GetRecordset")
End Function
