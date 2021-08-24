Attribute VB_Name = "modDBConnection"
Option Explicit

'########## 데이터베이스 관련 ############################
'DB Type    [ 1 : 오라클, 2 : MSSQL, 3 : Postgres ]
Public gDBTYPE      As String

'로컬데이터베이스
Type LocalDBParameter
    PATH    As String
    UID     As String
    PWD     As String
End Type

Public gLocalDB     As LocalDBParameter

'병원데이터베이스(오라클)
Type OracleDBParameter
    SID     As String
    UID     As String
    PWD     As String
End Type
Public gORADB        As OracleDBParameter

'병원데이터베이스(MSSQL)
Type MSSQLDBParameter
    IP      As String
    DB      As String
    UID     As String
    PWD     As String
End Type
Public gSQLDB        As MSSQLDBParameter

'병원데이터베이스(Postgres SQL)
Type PGSQLDBParameter
    IP      As String
    DB      As String
    UID     As String
    PWD     As String
End Type
Public gPGSQLDB      As PGSQLDBParameter

'QC데이터베이스(MSSQL)
Type MSSQLDBParameter_QC
    IP      As String
    DB      As String
    UID     As String
    PWD     As String
End Type
Public gSQLDB_QC     As MSSQLDBParameter_QC

'Urin Micro
Type UrinMicro
    WBCM    As String
    RBCM    As String
    EPIC    As String
    BACT    As String
End Type

Public gUrinMic As UrinMicro


'MS JetDB
Global AdoCn_Local          As ADODB.Connection
Global AdoRs_Local          As ADODB.Recordset
Global AdoCmd_Local         As ADODB.Command
Global AdoParm_Local        As ADODB.Parameter


'MS JetDB QC
'Global AdoCn_Local          As ADODB.Connection
'Global AdoRs_Local_QC          As ADODB.Recordset
'Global AdoCmd_Local_QC         As ADODB.Command
'Global AdoParm_Local_QC        As ADODB.Parameter


'ORACLE Server
'''Global AdoCn         As ADODB.Connection
'''Global AdoRs_ORACLE         As ADODB.Recordset
'''Global AdoCmd_ORACLE        As ADODB.Command
'''Global AdoParm_ORACLE       As ADODB.Parameter
'''
''''SQL Server
'''Global AdoCn            As ADODB.Connection
'''Global AdoRs_SQL            As ADODB.Recordset
'''Global AdoCmd_SQL           As ADODB.Command
'''Global AdoParm_SQL          As ADODB.Parameter

'공통
Global AdoCn                As ADODB.Connection
Global AdoRs                As ADODB.Recordset
Global AdoCmd               As ADODB.Command
Global AdoParm              As ADODB.Parameter

'QC
Global AdoCn_QC             As ADODB.Connection
Global AdoRs_QC             As ADODB.Recordset
Global AdoCmd_QC            As ADODB.Command
Global AdoParm_QC           As ADODB.Parameter


Public cn_Local_Flag        As Boolean
Public cn_Server_Flag       As Boolean

'########## 데이터베이스 관련 ############################


'-- MDB 연결
Public Function DbConnect_Local() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean

    Set AdoCn_Local = New ADODB.Connection

On Error GoTo ConnectError

    DB_Name = gLocalDB.PATH
    UserName = gLocalDB.UID
    Password = gLocalDB.PWD
    
    If DB_Name = "" Then
        DB_Name = App.PATH & "\Database\" & gHOSP.MACHNM & ".mdb"
    End If
    
    If UserName = "" Then
        UserName = "admin"
    End If
    
    If (DB_Name = "") Or (UserName = "") Then
        DbConnect_Local = False
        Set AdoCn_Local = Nothing
        Exit Function
    End If

    With AdoCn_Local
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

'-- QC MDB 연결
'Public Function DbConnect_Local_QC() As Boolean
'    Dim DB_Name         As String
'    Dim UserName        As String
'    Dim Password        As String
'    Dim blnWinNTAuth    As Boolean
'
'    Set AdoCn_Local = New ADODB.Connection
'
'On Error GoTo ConnectError
'
'    DB_Name = Mid(gLocalDB.PATH, 1, Len(gLocalDB.PATH) - 4) & "_QC.mdb"
'
'    UserName = gLocalDB.UID
'    Password = gLocalDB.PWD
'
'    If (DB_Name = "") Or (UserName = "") Then
'        DbConnect_Local_QC = False
'        Set AdoCn_Local = Nothing
'        Exit Function
'    End If
'    Dim i As Integer
'
'
'    With AdoCn_Local
'        .ConnectionTimeout = 25
'        .CursorLocation = adUseClient
'        .Provider = "Microsoft.Jet.OLEDB.4.0"
'        .Properties("Mode").Value = adModeReadWrite
'        .Properties("Persist Security Info").Value = False
'        .Properties("Data Source").Value = DB_Name
'        .Properties("User ID").Value = UserName
'        .Properties("Jet OLEDB:Database Password").Value = Password
'        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
'
'        .Open
'
'    End With
'
'    Screen.MousePointer = vbDefault
'    DbConnect_Local_QC = True
' Exit Function
'
'ConnectError:
'
'    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
'           " Description : " & Err.Description & vbCrLf & _
'           "      Source : " & Err.Source & vbCrLf & vbCrLf _
'           , vbCritical, " DB Open Error"
'
'    If AdoCn_Local.State <> adStateOpen Then
'        DbConnect_Local_QC = False
'        Set AdoCn_Local = Nothing
'    End If
'
'End Function

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

    If ServerName = "" Then
        Select Case gEMR
            Case "EONM":        ServerName = "EES"
            Case "AMIS":        ServerName = "sgknm"
        End Select
    End If
    
    If UserName = "" Then
        Select Case gEMR
            Case "EONM":        UserName = "EON_SPP"
            Case "AMIS":        UserName = "scott"
        End Select
    End If
    
    If Password = "" Then
        Select Case gEMR
            Case "EONM":        Password = "EON_SPP"
            Case "AMIS":        Password = "scott001"
        End Select
    End If
    
    If (ServerName = "") Then
        DbConnect_ORACLE = False
        Set AdoCn = Nothing
        Exit Function
    End If
    
    With AdoCn
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"
        .Properties("Data Source").Value = ServerName
'        .Properties("Initial Catalog").Value = DatabaseName
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

'-- MS SQL 연결
Public Function DbConnect_SQL() As Boolean
    
    Dim ServerName As String
    Dim DatabaseName As String
    Dim UserName As String
    Dim Password As String
    Dim blnWinNTAuth As Boolean

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    ServerName = gSQLDB.IP
    DatabaseName = gSQLDB.DB
    UserName = gSQLDB.UID
    Password = gSQLDB.PWD

    If ServerName = "" Then
        Select Case gEMR
            Case "PLIS":        ServerName = "192.168.1.13"
        End Select
    End If
    
    If DatabaseName = "" Then
        Select Case gEMR
            Case "PLIS":        DatabaseName = "plis"
        End Select
    End If
    
    If UserName = "" Then
        Select Case gEMR
            Case "PLIS":        UserName = "sa"
        End Select
    End If
    
    If Password = "" Then
        Select Case gEMR
            Case "PLIS":        Password = "hib"
        End Select
    End If
    
    If (ServerName = "") Or (DatabaseName = "") Then
        DbConnect_SQL = False
        Set AdoCn = Nothing
        Exit Function
    End If
    
    With AdoCn
        .ConnectionTimeout = 25
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = ServerName
        .Properties("Initial Catalog").Value = DatabaseName

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
    DbConnect_SQL = True
    
 Exit Function

ConnectError:
    Screen.MousePointer = vbDefault
    MsgBox " Error No. : " & Err.Number & vbCrLf & _
            " Description : " & Err.Description & vbCrLf & _
            " Source : " & Err.Source & vbCrLf & vbCrLf
            
    If AdoCn.State <> adStateOpen Then
        DbConnect_SQL = False
        Set AdoCn = Nothing
    End If

End Function

'-- PostGresSQL 연결
Public Function DbConnect_PostGres() As Boolean
    
    Dim ServerName As String
    Dim DatabaseName As String
    Dim UserName As String
    Dim Password As String
    Dim blnWinNTAuth As Boolean

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    ServerName = gPGSQLDB.IP
    DatabaseName = gPGSQLDB.DB
    UserName = gPGSQLDB.UID
    Password = gPGSQLDB.PWD

    If (ServerName = "") Or (DatabaseName = "") Then
       DbConnect_PostGres = False
       Set AdoCn = Nothing
       Exit Function
    End If
    
    With AdoCn
        '-- postgress
        .ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;" & _
                            "Data Source=" & ServerName & ";" & _
                            "Initial Catalog=" & DatabaseName & ";" & _
                            "User ID=" & UserName & ";" & _
                            "Password=" & Password
        
        Screen.MousePointer = vbHourglass
        
        .Open
    End With
    
    Screen.MousePointer = vbDefault
    DbConnect_PostGres = True
    
 Exit Function

ConnectError:
    Screen.MousePointer = vbDefault
    MsgBox " Error No. : " & Err.Number & vbCrLf & _
            " Description : " & Err.Description & vbCrLf & _
            " Source : " & Err.Source & vbCrLf & vbCrLf
            
    If AdoCn.State <> adStateOpen Then
        DbConnect_PostGres = False
        Set AdoCn = Nothing
    End If

End Function

'-- MS SQL 연결 QC
Public Function DbConnect_SQL_QC() As Boolean
    
    Dim ServerName As String
    Dim DatabaseName As String
    Dim UserName As String
    Dim Password As String
    Dim blnWinNTAuth As Boolean

    Set AdoCn_QC = New ADODB.Connection

On Error GoTo ConnectError

    ServerName = gSQLDB_QC.IP
    DatabaseName = gSQLDB_QC.DB
    UserName = gSQLDB_QC.UID
    Password = gSQLDB_QC.PWD

    If (ServerName = "") Or (DatabaseName = "") Then
       DbConnect_SQL_QC = False
       Set AdoCn_QC = Nothing
       Exit Function
    End If
    
    With AdoCn_QC
        .ConnectionTimeout = 25
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = ServerName
        .Properties("Initial Catalog").Value = DatabaseName
        
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
    DbConnect_SQL_QC = True
    
 Exit Function

ConnectError:
    Screen.MousePointer = vbDefault
    MsgBox " Error No. : " & Err.Number & vbCrLf & _
            " Description : " & Err.Description & vbCrLf & _
            " Source : " & Err.Source & vbCrLf & vbCrLf
            
    If AdoCn.State <> adStateOpen Then
        DbConnect_SQL_QC = False
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

    'Call DBErrorSet(AdoCn, strSql, Call_Name)
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
    'Call DBErrorSet(AdoCn, strSql, Call_Name)
End Function


