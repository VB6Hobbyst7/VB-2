Attribute VB_Name = "modDBConnection"
Option Explicit

'===========================REGISTREE PATH
'데이타 베이스 부분
'Global Const REG_MSSQLDB    As String = REG_POSITION & "\CONECT_SQL" 'SQL Server
'Global Const REG_ORACLEDB   As String = REG_POSITION & "\CONECT_ORACLE" 'ORACLE
'Global Const REG_JETDB      As String = REG_POSITION & "\CONECT_JET" 'JET DB
'데이타베이스 사용자
Global Const REG_SERVER     As String = "SERVER"
Global Const REG_DATABASE   As String = "DATABASE"
Global Const REG_SERVICE    As String = "SERVICE"
Global Const REG_USER_ID    As String = "USERID"
Global Const REG_PASSWD     As String = "PASSWD"

'===========================DATABASE CONECTION
'MS JetDB
Global AdoCn_Jet            As ADODB.Connection
Global AdoRs_Jet            As ADODB.Recordset
Global AdoCmd_Jet           As ADODB.Command
Global AdoParm_Jet          As ADODB.Parameter
'SQL Server
Global AdoCn_SQL            As ADODB.Connection
Global AdoRs_SQL            As ADODB.Recordset
Global AdoCmd_SQL           As ADODB.Command
Global AdoParm_SQL          As ADODB.Parameter
'ORACLE Server
Global AdoCn_ORACLE         As ADODB.Connection
Global AdoRs_ORACLE         As ADODB.Recordset
Global AdoCmd_ORACLE        As ADODB.Command
Global AdoParm_ORACLE       As ADODB.Parameter

'Public Function DBConnect_MDS() As Boolean ' MS Acess2000 데이터 베이스 붙일때
'    Dim DB_Name         As String
'    Dim UserName        As String
'    Dim Password        As String
'    Dim blnWinNTAuth    As Boolean
'
'    Set AdoCn_Jet = New ADODB.Connection
'
'On Error GoTo ConnectError
'
'    DB_Name = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)
'    UserName = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_USER_ID)
'    Password = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_PASSWD)
'
'    If (DB_Name = "") Or (UserName = "") Then
'        DBConnect_MDS = False
'        Set AdoCn_Jet = Nothing
'        Exit Function
'    Else
'        Call CompactJET(DB_Name)
'    End If
'    Dim i As Integer
'
'    With AdoCn_Jet
'        .ConnectionTimeout = 25
'        .CursorLocation = adUseClient
'        .Provider = "MSDataShape.1"
'        .Properties("Mode").Value = adModeReadWrite
'        .Properties("Persist Security Info").Value = False
'        .Properties("Data Source").Value = DB_Name
'        .Properties("Data Provider").Value = "MICROSOFT.JET.OLEDB.4.0"
'        .Open
'    End With
'
'    Screen.MousePointer = vbDefault
'    DBConnect_MDS = True
' Exit Function
'
'ConnectError:
'
'    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
'           " Description : " & Err.Description & vbCrLf & _
'           "      Source : " & Err.Source & vbCrLf & vbCrLf _
'           , vbCritical, " DB Open Error"
'
'    If AdoCn_Jet.State <> adStateOpen Then
'        DBConnect_MDS = False
'        Set AdoCn_Jet = Nothing
'    End If
'
'End Function

Public Function DbConnect_Jet() As Boolean ' MS Acess2000 데이터 베이스 붙일때
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean

    Set AdoCn_Jet = New ADODB.Connection

On Error GoTo ConnectError

    DB_Name = App.Path & "\Database\DataBase.mdb"   'GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)
    UserName = "admin" 'GetString(HKEY_CURRENT_USER, REG_JETDB, REG_USER_ID)
    Password = "admin" 'GetString(HKEY_CURRENT_USER, REG_JETDB, REG_PASSWD)

    If (DB_Name = "") Or (UserName = "") Then
        DbConnect_Jet = False
        Set AdoCn_Jet = Nothing
        Exit Function
'    Else
'        Call CompactJetDatabase(DB_Name)
    End If
    Dim i As Integer


    With AdoCn_Jet
        .ConnectionTimeout = 25
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Mode").Value = adModeReadWrite
        .Properties("Persist Security Info").Value = False
        .Properties("Data Source").Value = DB_Name
        .Properties("User ID").Value = UserName
        .Properties("Jet OLEDB:Database Password").Value = Password
        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
'        For i = 0 To AdoCn_Jet.Properties.Count - 1
'            Debug.Print AdoCn_Jet.Properties.Item(i).Name & "   " & AdoCn_Jet.Properties.Item(i).Value
'        Next
        .Open
    End With

    Screen.MousePointer = vbDefault
    DbConnect_Jet = True
 Exit Function

ConnectError:

    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn_Jet.State <> adStateOpen Then
        DbConnect_Jet = False
        Set AdoCn_Jet = Nothing
    End If

End Function

Public Sub DisConnect_Jet() ' MS Acess2000 데이터 베이스 끊을때

On Error GoTo ErrorRouten
    
    If Not AdoRs_Jet Is Nothing Then
        Call RsClose
        If AdoCn_Jet.State <> adStateClosed Then AdoCn_Jet.Close
        Set AdoCn_Jet = Nothing
    End If
    
Exit Sub

ErrorRouten:

    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf _
           , vbCritical, " DB Close Error"
    
    If Not AdoCn_Jet Is Nothing Then
        If AdoCn_Jet.State <> adStateClosed Then
            Set AdoCn_Jet = Nothing
        End If
    End If

End Sub

Public Sub RsClose()
    If AdoRs_Jet.State <> adStateClosed Then AdoRs_Jet.Close
    Set AdoRs_Jet = Nothing
End Sub


'Public Function DbConnect_SQL() As Boolean ' MS SQL2000 데이터 베이스 연결
'
'    Dim ServerName As String
'    Dim DatabaseName As String
'    Dim UserName As String
'    Dim Password As String
'    Dim blnWinNTAuth As Boolean
'
'    Set AdoCn_SQL = New ADODB.Connection
'
'On Error GoTo ConnectError
'
'     ServerName = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_SERVER)
'     DatabaseName = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_DATABASE)
'     UserName = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_USER_ID)
'     Password = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_PASSWD)
'
'     If (ServerName = "") Or (DatabaseName = "") Then
'        DbConnect_SQL = False
'        Set AdoCn_SQL = Nothing
'        Exit Function
'     End If
'    With AdoCn_SQL
'        .ConnectionTimeout = 25
'        .Provider = "SQLOLEDB"
'        .Properties("Data Source").Value = ServerName
'        .Properties("Initial Catalog").Value = DatabaseName
'
'        If blnWinNTAuth = True Then
'            .Properties("Integrated Security").Value = "SSPI"
'        Else
'            .Properties("User ID").Value = UserName
'            .Properties("Password").Value = Password
'        End If
'
'        Screen.MousePointer = vbHourglass
'        .Open
'    End With
'
'    Screen.MousePointer = vbDefault
'    DbConnect_SQL = True
' Exit Function
'
'ConnectError:
'    Screen.MousePointer = vbDefault
'    MsgBox " Error No. : " & Err.Number & vbCrLf & _
'            " Description : " & Err.Description & vbCrLf & _
'            " Source : " & Err.Source & vbCrLf & vbCrLf
'
'    If AdoCn_SQL.State <> adStateOpen Then
'        DbConnect_SQL = False
'        Set AdoCn_SQL = Nothing
'    End If
'
'End Function

Public Sub DisConnect_SQL()         ' MS SQL2000 데이터 베이스 끊을때

On Error GoTo ErrorRouten
    
    If Not AdoRs_SQL Is Nothing Then
        Call RsClose
    End If
    
    If Not AdoCn_SQL Is Nothing Then
        AdoCn_SQL.Close
        Set AdoCn_SQL = Nothing
    End If
Exit Sub

ErrorRouten:
    MsgBox "    Error No : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf _
           , vbCritical, " DB Close Error"
    
    If Not AdoCn_SQL Is Nothing Then
        Set AdoCn_SQL = Nothing
    End If
    
End Sub
'
'Public Function DbConnect_ORACLE() As Boolean ' ORACLE 데이터 베이스 연결
'
'    Dim ServerName As String
'    Dim DatabaseName As String
'    Dim UserName As String
'    Dim Password As String
'    Dim blnWinNTAuth As Boolean
'
'    Set AdoCn_ORACLE = New ADODB.Connection
'
'On Error GoTo ConnectError
'
'     ServerName = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_SERVER)
''     DatabaseName = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_DATABASE)
'     UserName = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_USER_ID)
'     Password = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_PASSWD)
'
'     If (ServerName = "") Then
'        DbConnect_ORACLE = False
'        Set AdoCn_ORACLE = Nothing
'        Exit Function
'     End If
'    With AdoCn_ORACLE
'        .ConnectionTimeout = 25
'        '-- 한신메디피아 OraOLEDB로 접속안되어서 MSDAORA로 접속함 ---
''        .Provider = "OraOLEDB.Oracle.1"
'        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
'        '------------------------------------------------------------
'        .Properties("Data Source").Value = ServerName
''        .Properties("Initial Catalog").Value = DatabaseName
'        .Properties("Persist Security Info") = True
'
'        If blnWinNTAuth = True Then
'            .Properties("Integrated Security").Value = "SSPI"
'        Else
'            .Properties("User ID").Value = UserName
'            .Properties("Password").Value = Password
'        End If
'
'        Screen.MousePointer = vbHourglass
'        .Open
'    End With
'    Screen.MousePointer = vbDefault
'    DbConnect_ORACLE = True
' Exit Function
'
'ConnectError:
'    Screen.MousePointer = vbDefault
'    MsgBox " Error No. : " & Err.Number & vbCrLf & _
'            " Description : " & Err.Description & vbCrLf & _
'            " Source : " & Err.Source & vbCrLf & vbCrLf
'
'    If AdoCn_ORACLE.State <> adStateOpen Then
'        DbConnect_ORACLE = False
'        Set AdoCn_ORACLE = Nothing
'    End If
'
'End Function

Public Sub DisConnect_ORACLE()         ' ORACLE 데이터 베이스 끊을때

On Error GoTo ErrorRouten
    
    If Not AdoRs_SQL Is Nothing Then
        Call RsClose
    End If
    
    If Not AdoCn_ORACLE Is Nothing Then
        AdoCn_ORACLE.Close
        Set AdoCn_ORACLE = Nothing
    End If
Exit Sub

ErrorRouten:
    MsgBox "    Error No : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf _
           , vbCritical, " DB Close Error"
    
    If Not AdoCn_ORACLE Is Nothing Then
        Set AdoCn_ORACLE = Nothing
    End If
    
End Sub

'=================================== Jet_DB 압축
'[프로젝트 - 참조]/[Microsoft Access 9.0 Object Library]를 추가.

Public Sub CompactMDB(Location As String, Optional BackupOriginal As Boolean = True)

On Error GoTo CompactErr

    Dim strBackupFile As String
    Dim strTempFile As String

'    If Len(Dir(Location)) Then
'        If BackupOriginal = True Then
'            strBackupFile = GetBackupPath & "Backup.mdb"
'            If Len(Dir(strBackupFile)) Then Kill strBackupFile
'            FileCopy Location, strBackupFile
'        End If
'
'        strTempFile = GetBackupPath & "Temp.mdb"
'
'        If Len(Dir(strTempFile)) Then Kill strTempFile
'
'        DBEngine.CompactDatabase Location, strTempFile
'        Kill Location
'        FileCopy strTempFile, Location
'        Kill strTempFile
'    End If

Exit Sub

CompactErr:

End Sub
'
''[프로젝트 - 참조]/[Micorsoft Jet and Replication Object 2.5 Library]를 추가.
'Public Sub CompactJET(ByVal SourcePath As String)
'    'SourcePath : MDB의 풀패스를 지정해준다.
'    'BackupPath   : 압축되어질 임시 MDB의 풀패스를 지정해준다.
'
'On Error GoTo Errorhandler
'    Dim JET_JRO         As JRO.JetEngine
'    Dim BackupPath      As String
'
'    BackupPath = GetBackupPath & "BACKUP.MDB"
'    'Backup File 이 있으면 삭제 한다.
'    If Len(Dir(BackupPath)) <> 0 Then
'        Kill BackupPath
'    End If
'
'    Set JET_JRO = New JRO.JetEngine
'    Call JET_JRO.CompactDatabase("Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & SourcePath & ";Persist Security Info=False", _
'                                 "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & BackupPath) ' & ";Jet OLEDB:Engine Type=4")
'    Set JET_JRO = Nothing
'
'    '원본 MDB(BackupPath)를 삭제한다.
'    If Len(Dir(SourcePath)) <> 0 Then
'        Kill SourcePath
'    End If
'
'    '압축되어진 임시 MDB(BackupPath)의 이름을 원본 MDB(SourcePath) 이름으로 RENAME시켜준다.
'    Call FileCopy(BackupPath, SourcePath)
'Errorhandler:
'
'End Sub

Public Function GetBackupPath() As String
    Dim BackupPath As String
    Dim pathLangth As String
    
    BackupPath = App.Path
    
    If Right(GetBackupPath, 1) <> "\" Then
        BackupPath = BackupPath & "\" & "Backup\"
    End If
    pathLangth = Dir(BackupPath)
    
    If Len(Dir(BackupPath, vbDirectory)) < 1 Then
        Call MkDir(BackupPath)
    End If
    
    GetBackupPath = BackupPath
    
End Function
'====================================================


'Data Base의 현재일자시간
'Public Function DbSysDate() As Date
'    Dim objData As New clsSQLData
'
'    With objData
'        .SetAdoCn AdoCn
'        DbSysDate = .GetSysDate
'    End With
'
'    Set objData = Nothing
'End Function

'Public Function Write_Log(ByVal DATA As String) As Boolean
'    Dim strTmp() As String
'    Dim strSQL As String
'On Error GoTo ErrorRoutin
'
'    strTmp = Split(DATA, vbTab)
'    Dim objPstmt As New clsPreparedStatem
'
'    AdoCn_Jet.BeginTrans
'    With objPstmt
'        .initPreparedStmt "INSERT INTO Conn_Log ([Date], [Time], [Port_ID], [Remote_IP], [Discription], [Remark]) " & _
'                          " VALUES ( ?, ?, ?, ?, ?, ?) "
'        .setString 1, Format(strTmp(0), "YYYYMMDD")
'        .setString 2, Format(strTmp(0), "HHNNSS")
'        .setString 3, strTmp(1)
'        .setString 4, strTmp(2)
'        .setString 5, strTmp(3)
'        If UBound(strTmp) < 4 Then
'            .setString 6, " "
'        Else
'            .setString 6, strTmp(4)
'        End If
'        AdoCn_Jet.Execute .getPreparedStmt
'    End With
'
'    AdoCn_Jet.CommitTrans
'    Set objPstmt = Nothing
'
'Exit Function
'
'ErrorRoutin:
'    AdoCn_Jet.RollbackTrans
'
'    MsgBox "    Error No : " & Err.Number & vbCrLf & _
'           " Description : " & Err.Description & vbCrLf & _
'           "      Source : " & Err.Source & vbCrLf _
'           , vbCritical, " DB Insert Error"
'
'End Function

'Public Function Read_Log(ByVal CmdTxt As String, ByVal objParam As Scripting.Dictionary) As ADODB.Recordset
'
'    Dim ObjX As Variant
'    Dim strParm() As String
'
'    Set AdoParm_Jet = New ADODB.Parameter
'    Set AdoCmd_Jet = New ADODB.Command
'    Set AdoRs_Jet = New ADODB.Recordset
'
'    With AdoCmd_Jet
'        .ActiveConnection = AdoCn_Jet
'
'        For Each ObjX In objParam
'            strParm = Split(objParam(ObjX), vbTab)
'            Set AdoParm_Jet = .CreateParameter(ObjX, strParm(1), strParm(2), strParm(3), strParm(4))
'            .Parameters.Append AdoParm_Jet
'        Next
'        .CommandType = adCmdTable
'        .CommandText = CmdTxt
'        .Execute
'        AdoRs_Jet.Open AdoCmd_Jet
'    End With
'
'    Set Read_Log = AdoRs_Jet
'
'    Set AdoParm_Jet = Nothing
'    Set AdoCmd_Jet = Nothing
'    Set AdoRs_Jet = Nothing
'
'End Function

'Public Function SetParamter(ByVal ParmNM As String, ByVal dType As DataTypeEnum, _
'                            ByVal Direction As ParameterDirectionEnum, ByVal Size As Long, ByVal Value)
'    SetParamter = ParmNM & vbTab & dType & vbTab & Direction & vbTab & Size & vbTab & Value
'End Function





