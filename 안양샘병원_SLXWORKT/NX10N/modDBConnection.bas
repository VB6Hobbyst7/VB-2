Attribute VB_Name = "modDBConnection"
Option Explicit

'########## �����ͺ��̽� ���� ############################
'DB Type    [ 1 : ����Ŭ, 2 : MSSQL, 3 : Postgres ]
Public gDBTYPE      As String

'���õ����ͺ��̽�
Type LocalDBParameter
    PATH    As String
    UID     As String
    PWD     As String
End Type

Public gLocalDB     As LocalDBParameter

'���������ͺ��̽�(����Ŭ)
Type OracleDBParameter
    SID     As String
    UID     As String
    PWD     As String
End Type
Public gORADB        As OracleDBParameter

'���������ͺ��̽�(MSSQL)
Type MSSQLDBParameter
    IP      As String
    DB      As String
    UID     As String
    PWD     As String
End Type
Public gSQLDB        As MSSQLDBParameter

'���������ͺ��̽�(Postgres SQL)
Type PGSQLDBParameter
    IP      As String
    DB      As String
    UID     As String
    PWD     As String
End Type
Public gPGSQLDB      As PGSQLDBParameter

'QC�����ͺ��̽�(MSSQL)
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

'����
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

'########## �����ͺ��̽� ���� ############################


'-- MDB ����
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

'-- QC MDB ����
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

'-- ORACLE ����
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
            Case "AMIS":        ServerName = "sgknm"    'POINTNET
            Case "JWINFO":      ServerName = "DJH"
            Case "SAM":         ServerName = "sambaseocs"
        End Select
    End If
    
    If UserName = "" Then
        Select Case gEMR
            Case "EONM":        UserName = "EON_SPP"
            Case "AMIS":        UserName = "scott"      'POINTNET
            Case "JWINFO":      UserName = "dreamer"
            Case "SAM":         UserName = "oras1"
        End Select
    End If
    
    If Password = "" Then
        Select Case gEMR
            Case "EONM":        Password = "EON_SPP"
            Case "AMIS":        Password = "scott001"   'POINTNET
            Case "JWINFO":      Password = "dsdvp"
            Case "SAM":         Password = "oras1"
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
            " Source : " & Err.Source & vbCrLf & vbCrLf, vbOKOnly + vbCritical, "Database �������"
            
    If AdoCn.State <> adStateOpen Then
        DbConnect_ORACLE = False
        Set AdoCn = Nothing
    End If

End Function

'-- MS SQL ����
Public Function DbConnect_SQL() As Boolean
    
    Dim ServerName As String
    Dim DatabaseName As String
    Dim UserName As String
    Dim Password As String
    Dim blnWinNTAuth As Boolean
    
    Dim adoConnectionString As String
    
    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    ServerName = gSQLDB.IP
    DatabaseName = gSQLDB.DB
    UserName = gSQLDB.UID
    Password = gSQLDB.PWD

'MSSQLIP=192.168.0.250
'MSSQLDB = drbitpack
'MSSQLUID = sa
'MSSQLPWD = Bit
    
'MSSQLIP=bitserver
'MSSQLDB=medichart
'MSSQLUID=sa
'MSSQLPWD=bit

'MSSQLIP = LisILib
'MSSQLDB = LISOLIB
'MSSQLUID = IFE13
'MSSQLPWD = E1331E

'MSSQLIP=192.168.0.250
'MSSQLDB = drbitpack
'MSSQLUID = sa
'MSSQLPWD = Bit


'MSSQLIP=192.168.0.228
'MSSQLDB=MEDIPLUS
'MSSQLUID=interface_base
'MSSQLPWD=int_!jj1m

    If ServerName = "" Then
        Select Case gEMR
            Case "SANSOFT":     ServerName = "192.168.0.7\SQLEXPRESS,1433" '��õ����
            Case "BIT70":       ServerName = "192.168.0.10" '��õ����
            Case "EONM":        ServerName = "EES"
            Case "AMIS":        ServerName = "sgknm"
            Case "LABSPEAR":    ServerName = "222.107.187.101,4433" '���Ƿ���� �̳뺣��Ʈ
            Case "MEDICHART":   ServerName = "bitserver"
            Case "BIT":         ServerName = "192.168.0.240" '���ϳ���
            Case "MCC":         ServerName = "192.168.0.228" '���򺴿�
        End Select
    End If
    
    If DatabaseName = "" Then
        Select Case gEMR
            Case "SANSOFT":     DatabaseName = "LabSpearKcLab"
            Case "BIT70":       DatabaseName = "HIB70"
            Case "EONM":        DatabaseName = "EES"
            Case "AMIS":        DatabaseName = "sgknm"
            Case "LABSPEAR":    DatabaseName = "LabSpearKcLab"
            Case "MEDICHART":   DatabaseName = "medichart"
            Case "BIT":         DatabaseName = "drbitpack"
            Case "MCC":         DatabaseName = "MEDIPLUS"
        End Select
    End If
    
    If UserName = "" Then
        Select Case gEMR
            Case "SANSOFT":     UserName = "sa"
            Case "BIT70":       UserName = "sa"
            Case "EONM":        UserName = "EON_SPP"
            Case "AMIS":        UserName = "scott"
            Case "LABSPEAR":    UserName = "interspear"
            Case "MEDICHART":   UserName = "sa"
            Case "BIT":         UserName = "sa"
            Case "MCC":         UserName = "interface_base"
        End Select
    End If
    
    If Password = "" Then
        Select Case gEMR
            Case "SANSOFT":     Password = "1004"
            Case "BIT70":       Password = "hib"
            Case "EONM":        Password = "EON_SPP"
            Case "AMIS":        Password = "scott001"
            Case "LABSPEAR":    Password = "intermanager@!@#$"
            Case "MEDICHART":   Password = "bit"
            Case "BIT":         Password = "bit"
            Case "MCC":         Password = "int_!jj1m"
        End Select
    End If
    

    If (ServerName = "") Or (DatabaseName = "") Then
       DbConnect_SQL = False
       Set AdoCn = Nothing
       Exit Function
    End If
    
    If gEMR = "SCL" Then
        With AdoCn
            adoConnectionString = "dsn=" & ServerName & ";" & _
                                  "uid=" & UserName & ";" & _
                                  "pwd=" & Password & ";" & _
                                  "database=" & DatabaseName & ""
                
            .ConnectionString = adoConnectionString
            .Open
        End With
    Else
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
    End If
    
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

'-- PostGresSQL ����
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

'    PGSQLIP=IF
'    PGSQLDB = postgres
'    PGSQLUID = postgres
'    PGSQLPWD = clinic2013!

    If ServerName = "" Then
        Select Case gEMR
            Case "EASYS":     ServerName = "IF" 'ODBC���� ����� �̸�
        End Select
    End If
    
    If DatabaseName = "" Then
        Select Case gEMR
            Case "EASYS":     DatabaseName = "postgres"
        End Select
    End If
    
    If UserName = "" Then
        Select Case gEMR
            Case "EASYS":     UserName = "postgres"
        End Select
    End If
    
    If Password = "" Then
        Select Case gEMR
            Case "EASYS":     Password = "clinic2013!"
        End Select
    End If
    
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

'-- MS SQL ���� QC
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


