Attribute VB_Name = "modDataBase"
Option Explicit

'sql server object for ole
Private mAdoCmd As ADODB.Command

'common sql/oracle method
'=======================================================
'execute the sql command
Public Function DBExec(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, Optional Call_Name As String) As Boolean
On Error GoTo DBExecError
    'lock table in row share mode
    With AdoCn
        .BeginTrans
        .Execute strSql
        .CommitTrans
    End With
    '
    DBExec = True
Exit Function
    
DBExecError:
    AdoCn.RollbackTrans
    DBExec = False
    Call DBErrorSet(AdoCn, strSql, Call_Name)
End Function

'트랜잭션 관리를 하지 않는다. 단지 sql 문장이 정상적으로 처리되었는지 true, false로 전달한다.
Public Function NotTransDBExec(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, Optional Call_Name As String) As Boolean
On Error GoTo NotTransDBEXecError

    AdoCn.Execute strSql
    NotTransDBExec = True
Exit Function

NotTransDBEXecError:
    NotTransDBExec = False
    Call DBErrorSet(AdoCn, strSql, Call_Name)
End Function

'Execute the Multi sql command
Public Function DBMultiExec(ByVal AdoCn As ADODB.Connection, ByRef arySql As Variant, Optional Call_Name As String) As Boolean
    Dim intSqlCnt As Integer

On Error GoTo DBExecError
    'Lock table in row share mode
    With AdoCn
        .BeginTrans
        For intSqlCnt = LBound(arySql) To UBound(arySql)
            .Execute arySql(intSqlCnt)
        Next intSqlCnt
        .CommitTrans
    End With
    '
    DBMultiExec = True
Exit Function
    
DBExecError:
    AdoCn.RollbackTrans
    DBMultiExec = False
    Call DBErrorSet(AdoCn, arySql(intSqlCnt), Call_Name)
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
    Call DBErrorSet(AdoCn, strSql, Call_Name)
End Function

'Check Database with oracle server-reset
Public Function DBExists(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, Optional Call_Name As String) As Boolean
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

    Call DBErrorSet(AdoCn, strSql, Call_Name)
End Function

'문장 양쪽에 Single quote 를 붙인다.
Public Function STS(ByVal strStmt As String) As String
    Dim strTmp As String
    
    strTmp = Replace(strStmt, "'", "''")
    
    STS = "'" & strTmp & "'"
End Function

Public Function Get_Date() As String

On Error GoTo ErrorRoutine

    Dim pAdoRS As ADODB.Recordset
    Set pAdoRS = New ADODB.Recordset
    Set pAdoRS = AdoCn_SQL.Execute("SELECT GETDATE() AS DBDATE")
    Get_Date = Format(pAdoRS("DBDATE"), "YYYYMMDD")
    Set pAdoRS = Nothing
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn_SQL, "SELECT GETDATE()", "Public Function Get_Date() As String")
End Function

Public Function Get_Time() As String

On Error GoTo ErrorRoutine

    Dim pAdoRS As ADODB.Recordset
    Set pAdoRS = New ADODB.Recordset
    Set pAdoRS = AdoCn_SQL.Execute("SELECT GETDATE() AS DBTIME")
    Get_Time = Format(pAdoRS("DBTIME"), "HHMM")
    pAdoRS.Close
    Set pAdoRS = Nothing
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn_SQL, "SELECT GETDATE()", "Public Function Get_Time() As String")
End Function

Public Function CreateTmpTable(AdoCn As ADODB.Connection, desTable As String, Optional sorTable As String) As Boolean
   Dim desAdoRS As ADODB.Recordset
   Dim sorAdoRS As ADODB.Recordset
   
   Dim sorFields As ADODB.Fields
   Dim sorField As ADODB.Field

On Error GoTo ErrHandler

    If TableExists(AdoCn, desTable) Then
        AdoCn.Execute "DROP TABLE " & desTable
        Set desAdoRS = New ADODB.Recordset 'AdoCn.Execute("CREATE TABLE " & desTable)
        Set sorAdoRS = AdoCn.Execute("SELECT TOP 1 * FROM " & sorTable)
    Else
'        Set desAdoRS = New ADODB.Recordset 'AdoCn.Execute("CREATE TABLE " & desTable)
        Set desAdoRS = AdoCn.Execute("CREATE TABLE " & desTable)
        Set sorAdoRS = AdoCn.Execute("SELECT TOP 1 * FROM " & sorTable)
    End If
    
    Set sorFields = sorAdoRS.Fields
    For Each sorField In sorFields
    
    Next
    
    desAdoRS.Open desTable
    CreateTmpTable = True
Exit Function
ErrHandler:
    CreateTmpTable = False
End Function

Public Function TableExists(AdoCn As ADODB.Connection, TableName As String) As Boolean
    Dim AdoRs As ADODB.Recordset
    Set AdoRs = AdoCn.OpenSchema(adSchemaTables)

    AdoRs.MoveFirst
    Call AdoRs.Find("TABLE_NAME = " & STS(TableName))
    
    If AdoRs.EOF Then
        TableExists = False
    Else
        TableExists = True
    End If
    Set AdoRs = Nothing
End Function

Public Function Del_OldData() As Boolean
    Dim strSql      As String
    Dim lngSave     As Long
    Dim pAdoRS      As ADODB.Recordset
    Dim FromDT      As String
 
On Error GoTo ErrorRoutine
    strSql = "SELECT SAVE_DT FROM INTERFACE001 WHERE EQP_CD = " & STS(INS_CODE)
 
    Set pAdoRS = New ADODB.Recordset
    Call GetRecordset(AdoCn_Jet, strSql, pAdoRS, "Public Function Del_OldData() As Boolean")
    If Not pAdoRS Is Nothing Then
        If pAdoRS.EOF Then
            lngSave = 30
        Else
            lngSave = Val(pAdoRS("SAVE_DT") & "")
        End If
        pAdoRS.Close
        Set pAdoRS = Nothing
    Else
        GoTo ErrorRoutine
    End If
    FromDT = Format(CDate(Format(Get_Date, "@@@@/@@/@@")) - lngSave, "YYYYMMDD")
    strSql = "DELETE FROM INTERFACE003 WHERE TRANSDT <= " & STS(FromDT)
    Call DBExec(AdoCn_Jet, strSql, "Public Function Del_OldData() As Boolean")
    Del_OldData = True
Exit Function
 
ErrorRoutine:
    Set pAdoRS = Nothing
    Del_OldData = False
End Function

