Attribute VB_Name = "BasOdbc"
Option Explicit

'Data Access Objects(DAO) Variable
Global DaoWS                    As Workspace            'DAO: WORK SPACE NAME
Global DaoDB                    As Database             'DAO: DB NAME
Global DaoData                  As Recordset            'DAO: QUERY한 데이타 NAME
Global DaoData1                 As Recordset            'DAO: QUERY한 데이타 NAME
Global DaoData2                 As Recordset            'DAO: QUERY한 데이타 NAME
Global DaoData3                 As Recordset            'DAO: QUERY한 데이타 NAME
Global DaoQDef                  As QueryDef             'DAO: 데이타 QUERYDEF
Global DaoTDef                  As TableDef             'DAO: 데이타 TABLEDEF
Global DaoError                 As Error
Global DaoParameter             As Parameter

'Remote Data Objects(RDO) Variable
Global RdoEnv                   As rdoEnvironment       'RDO: 환경변수
Global RdoDB                    As rdoConnection        'RDO: 데이타베이스 NAME
Global RdoQry                   As New rdoQuery         'RDO: QUERY문장 관리
Global RdoSet(9)                As rdoResultset         'RDO: QUERY 한 데이타 NAME
Global RdoSet1                  As rdoResultset         'RDO: QUERY 한 데이타 NAME
Global RdoSet2                  As rdoResultset         'RDO: QUERY 한 데이타 NAME
Global RdoSet3                  As rdoResultset         'RDO: QUERY 한 데이타 NAME
Global RdoPre                   As rdoPreparedStatement 'RDO: 데이타 Prepared NAME
Global RdoErr                   As rdoError             'RDO: Error Define
Global RdoResult                As rdoResultset         'RDO: Result Set

Global RecordCount              As Long                 '읽어들인 Record 총 갯수
Global strSql                   As String               'Query String 변수
Global RtnNOR                   As Long                 'ExecuteSQL의 리턴값 - Number Of Row
Global Result                   As Integer              'SQL 리턴값 True/False
Global GnDBOpenDAO              As Integer              'DAO Open True / False
Global GnDBOpenRDO              As Integer              'RDO Open True / False

Global GstrPassIDnumber         As String               '사용자확인시 필요
Global GstrPassProgramID        As String * 8           '사용자확인시 필요
Global GstrPassWord             As String               '사용자확인시 필요
Global GstrPassGrade            As String               '사용자확인시 필요
Global GstrPassClass            As String               '사용자확인시 필요
Global GstrPassName             As String               '사용자확인시 필요
Global GstrPassPart             As String * 2           '사용자확인시 필요
Global GstrPassDept             As String               '사용자확인시 필요
Global GstrPassRank             As Integer              '사용자확인시 필요


Public Function Convert_String(ArgStr As String) As String
   
   'ODBC를 통하여 Data를 INSERT할경우 DATA의 중간에 "'" 값이 있을경우
   'ERROR가 나옴 이경우를 방지하기위해서는 "'" 값을 "''"값으로 변환하여
   'INSERT하여야 함으로 이 FUNCTION을 사용함
   
    Dim nStart          As Integer
    Dim nPosition       As Integer
    Dim sReturnStr      As String
    
    If InStr(1, ArgStr, "'") = 0 Then       ' 찾는값이 없을경우
        Convert_String = ArgStr
        Exit Function
    End If
    
    nStart = 1
    sReturnStr = ""
    
    Do
        nPosition = InStr(nStart, ArgStr, "'")
        
        If nPosition = 0 Then
            sReturnStr = sReturnStr & Mid$(ArgStr, nStart)
            Exit Do
        Else
            sReturnStr = sReturnStr & Mid$(ArgStr, nStart, nPosition) & "'"
            nStart = nPosition + 1
        End If
    Loop
    
    Convert_String = sReturnStr
End Function


Public Sub DBConnectRDO_ODBCDRV(ArgDRV$, ArgUser$, ArgPwd$)

    Screen.MousePointer = 11        'HOURGLASS
    
On Error GoTo DB_Error_Handler      'Enable Error Trapping.
    rdoEngine.rdoDefaultCursorDriver = rdUseOdbc
    Set RdoEnv = rdoCreateEnvironment("", ArgUser$, ArgPwd$)   '( EnvName, Owner, PassWord )
    Set RdoDB = RdoEnv.OpenConnection(ArgDRV$, rdDriverNoPrompt, False)

On Error GoTo 0                     'Disable Error Trapping.
    
    Screen.MousePointer = 0         'Default
    GnDBOpenRDO = True
    Exit Sub

DB_Error_Handler:
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & vbCrLf & _
           "CAUSE - " & Error(Err), vbExclamation
    MsgBox "프로그램을 종료합니다.", vbInformation
    Screen.MousePointer = 0         'Default
    End
    
End Sub


Public Sub DBConnectRDO(ArgUser$, ArgPwd$)

    Screen.MousePointer = 11        'HOURGLASS
    
On Error GoTo DB_Error_Handler      'Enable Error Trapping.
    rdoEngine.rdoDefaultCursorDriver = rdUseOdbc
    Set RdoEnv = rdoCreateEnvironment("", ArgUser$, ArgPwd$)   '( EnvName, Owner, PassWord )
    Set RdoDB = RdoEnv.OpenConnection("ORA_TWMIS", rdDriverNoPrompt, False)

On Error GoTo 0                     'Disable Error Trapping.
    
    Screen.MousePointer = 0         'Default
    GnDBOpenRDO = True
    Exit Sub

DB_Error_Handler:
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & vbCrLf & _
           "CAUSE - " & Error(Err), vbExclamation
    MsgBox "프로그램을 종료합니다.", vbInformation
    Screen.MousePointer = 0         'Default
    End
    
End Sub


Public Function ExecDAO(ArgQuery As String) As Integer

    Screen.MousePointer = 11            'HOURGLASS

On Error GoTo Error_Handler_EXECUTE

    RtnNOR = DaoDB.ExecuteSQL(ArgQuery)
    If RtnNOR < 1 Then
        MsgBox "실행된 작업이 없습니다.", vbInformation, "참고"
    End If
    
    ExecDAO = True
    Screen.MousePointer = 0             'Default
Exit Function

Error_Handler_EXECUTE:
    
    DaoDB.Rollback
    ExecDAO = False
    
    Call ErrorMsgDAO(ArgQuery)
    Screen.MousePointer = 0             'Default
Exit Function

End Function


Public Function ExecRDO(ArgStrQuery$) As Integer
    
    Screen.MousePointer = 11        'HOURGLASS

On Error GoTo Error_Handler_EXECUTE

    RdoDB.Execute (ArgStrQuery$)
    RecordCount = Val(RdoDB.RowsAffected)
 
    ExecRDO = True

    Screen.MousePointer = 0         'Default
Exit Function

Error_Handler_EXECUTE:
    
    ExecRDO = False
    RecordCount = 0
    Call ErrorMsgRDO(ArgStrQuery$)
    Screen.MousePointer = 0         'Default
Exit Function

End Function


Public Function OpenDAO(ArgQuery As String) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo SQL_Error_Handler   'Enable error trapping.
    
    Set DaoData = DaoDB.OpenRecordset(ArgQuery, dbOpenSnapshot, dbSQLPassThrough)
    
On Error GoTo 0                 'Disable error trapping.
    
    If DaoData.EOF Then
        RecordCount = 0
        DaoData.Close
        OpenDAO = False
    Else
        DaoData.MoveLast
        RecordCount = DaoData.RecordCount
        DaoData.MoveFirst
        OpenDAO = True
    End If
    
    Screen.MousePointer = 0 'Default
    
    Exit Function

SQL_Error_Handler:
    OpenDAO = False
    RecordCount = -1
    Call ErrorMsgDAO(ArgQuery)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function
Public Sub DBConnectDAO(ArgUID As String, ArgPwd As String)

    Dim strConnect          As String
    
    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo DB_Error_Handler ' Enable ERROR Trapping.
    
    'JET/Data Access Objects(DAO) Library를 이용한 데이타베이스 Connect
    strConnect = "ODBC;"
    strConnect = strConnect & "DSN=ORA_TWMIS;"
    strConnect = strConnect & "UID=" & ArgUID & ";"
    strConnect = strConnect & "PWD=" & ArgPwd
    
   'Set DaoWS = CreateWorkspace("", "", "", dbUseODBC)
   'Set DaoDB = DaoWS.OpenDatabase("", False, False, strConnect)    '( DB_Name, Exclusive, ReadOnly, ConnectString)

    Set DaoWS = CreateWorkspace(ArgUID, "Admin", "")
    DBEngine.Workspaces.Append DaoWS
    Set DaoDB = DaoWS.OpenDatabase("", False, False, strConnect)    '( DB_Name, Exclusive, ReadOnly, ConnectString)
    
On Error GoTo 0 ' Disable error trapping.
    
    Screen.MousePointer = 0 'Default
    GnDBOpenDAO = True
    Exit Sub

DB_Error_Handler:
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & vbCrLf & "CAUSE - " & Error(Err), vbExclamation
    MsgBox "프로그램을 종료합니다.", vbInformation
    Screen.MousePointer = 0 'Default
    End
    
End Sub

Public Function OpenDAO3(ArgQuery As String) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo SQL_Error_Handler   'Enable error trapping.
    
    Set DaoData3 = DaoDB.OpenRecordset(ArgQuery, dbOpenSnapshot, dbSQLPassThrough)
    
On Error GoTo 0                 'Disable error trapping.
    
    If DaoData3.EOF Then
        DaoData3.Close
        RecordCount = 0
        OpenDAO3 = False
        Exit Function
    End If
    
    RecordCount = 1
    OpenDAO3 = True
    Screen.MousePointer = 0 'Default
    
    Exit Function

SQL_Error_Handler:
    RecordCount = -1
    OpenDAO3 = False
    Call ErrorMsgDAO(ArgQuery)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function

Public Function OpenRDO(ArgStrQuery$, ArgIndex%) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo DaoErrorHandler ' Enable error trapping.
    
    Set RdoSet(ArgIndex) = RdoDB.OpenResultset(ArgStrQuery$, rdOpenStatic, rdConcurReadOnly)     'Open Table.

On Error GoTo 0     ' Disable error trapping.
    
    If RdoSet(ArgIndex).EOF Then
        RecordCount = 0
        RdoSet(ArgIndex).Close
        Screen.MousePointer = 0 'Default
        OpenRDO = False
        Exit Function
    End If
    
    RecordCount = RdoSet(ArgIndex).RowCount
    RdoSet(ArgIndex).MoveFirst
    
    OpenRDO = True
    Screen.MousePointer = 0 'Default
    Exit Function


DaoErrorHandler:
    OpenRDO = False
    RecordCount = -1
    Call ErrorMsgRDO(ArgStrQuery$)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function
Public Function OpenDAO1(ArgQuery As String) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo SQL_Error_Handler   'Enable error trapping.
    
    Set DaoData1 = DaoDB.OpenRecordset(ArgQuery, dbOpenSnapshot, dbSQLPassThrough)
    
On Error GoTo 0                 'Disable error trapping.
    
    If DaoData1.EOF Then
        RecordCount = 0
        OpenDAO1 = False
        DaoData1.Close
        Exit Function
    End If
    
    RecordCount = 1
    OpenDAO1 = True
    Screen.MousePointer = 0 'Default
    
    Exit Function

SQL_Error_Handler:
    RecordCount = -1
    OpenDAO1 = False
    Call ErrorMsgDAO(ArgQuery)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function
Public Function OpenDAO2(ArgQuery As String) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo SQL_Error_Handler   'Enable error trapping.
    
    Set DaoData2 = DaoDB.OpenRecordset(ArgQuery, dbOpenSnapshot, dbSQLPassThrough)
    
On Error GoTo 0                 'Disable error trapping.
    
    If DaoData2.EOF Then
        DaoData2.Close
        RecordCount = 0
        OpenDAO2 = False
        Exit Function
    End If
    
    RecordCount = 1
    OpenDAO2 = True
    Screen.MousePointer = 0 'Default
    
    Exit Function

SQL_Error_Handler:
    RecordCount = -1
    OpenDAO2 = False
    Call ErrorMsgDAO(ArgQuery)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function

Public Sub ErrorMsgRDO(ArgQuery As String)
    Beep
    
    For Each RdoErr In rdoErrors
        MsgBox "오류코드 - " & RdoErr.Number & vbCrLf & _
               "오류소스 - " & RdoErr.Source & vbCrLf & _
               "오류내용 - " & RdoErr.Description & vbCrLf & vbCrLf & _
               "SQL 문장 - " & ArgQuery _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next RdoErr
    
End Sub
Public Sub ErrorMsgDAO(ArgQuery As String)
    
    Beep
    Set DaoError = DBEngine.Errors(0)
    MsgBox "오류코드 - " & DaoError.Number & Chr$(13) & _
           "오류소스 - " & DaoError.Source & Chr$(13) & _
           "오류내용 - " & DaoError.Description & Chr$(13) & Chr$(13) & _
           "SQL 문 - " & ArgQuery _
           , vbExclamation, "데이타작업중 오류가 발생했습니다."
    
End Sub
Public Function OpenRDO1(ArgStrQuery$) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo DaoErrorHandler ' Enable error trapping.
    
    Set RdoSet1 = RdoDB.OpenResultset(ArgStrQuery$, rdOpenStatic, rdConcurReadOnly)    'Open Table.

On Error GoTo 0     ' Disable error trapping.
    
    If RdoSet1.EOF Then
        RdoSet1.Close
        RecordCount = 0
        OpenRDO1 = False
        Screen.MousePointer = 0 'Default
        Exit Function
    End If
    
    RecordCount = 1
    OpenRDO1 = True
    Screen.MousePointer = 0 'Default
    Exit Function


DaoErrorHandler:
    RecordCount = -1
    OpenRDO1 = False
    Call ErrorMsgRDO(ArgStrQuery$)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function


Public Function OpenRDO2(ArgStrQuery$) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo DaoErrorHandler ' Enable error trapping.
    
    Set RdoSet2 = RdoDB.OpenResultset(ArgStrQuery$, rdOpenStatic, rdConcurReadOnly)    'Open Table.

On Error GoTo 0     ' Disable error trapping.
    
    If RdoSet2.EOF Then
        RdoSet2.Close
        RecordCount = 0
        OpenRDO2 = False
        Screen.MousePointer = 0 'Default
        Exit Function
    End If
    
    RecordCount = 1
    OpenRDO2 = True
    Screen.MousePointer = 0 'Default
    Exit Function


DaoErrorHandler:
    RecordCount = -1
    OpenRDO2 = False
    Call ErrorMsgRDO(ArgStrQuery$)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function


Public Function OpenRDO3(ArgStrQuery$) As Integer

    Screen.MousePointer = 11 'HOURGLASS

On Error GoTo DaoErrorHandler ' Enable error trapping.
    
    Set RdoSet3 = RdoDB.OpenResultset(ArgStrQuery$, rdOpenStatic, rdConcurReadOnly)    'Open Table.

On Error GoTo 0     ' Disable error trapping.
    
    If RdoSet3.EOF Then
        RdoSet3.Close
        RecordCount = 0
        OpenRDO3 = False
        Screen.MousePointer = 0 'Default
        Exit Function
    End If
    
    RecordCount = 1
    OpenRDO3 = True
    Screen.MousePointer = 0 'Default
    Exit Function


DaoErrorHandler:
    RecordCount = -1
    OpenRDO3 = False
    Call ErrorMsgRDO(ArgStrQuery$)
    Screen.MousePointer = 0 'Default
    Exit Function

End Function


