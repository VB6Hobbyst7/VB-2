Attribute VB_Name = "BasOdbc"
Option Explicit

'Data Access Objects(DAO) Variable
Global DaoWS                    As Workspace            'DAO: WORK SPACE NAME
Global DaoDB                    As Database             'DAO: DB NAME
Global DaoData                  As Recordset            'DAO: QUERY�� ����Ÿ NAME
Global DaoData1                 As Recordset            'DAO: QUERY�� ����Ÿ NAME
Global DaoData2                 As Recordset            'DAO: QUERY�� ����Ÿ NAME
Global DaoData3                 As Recordset            'DAO: QUERY�� ����Ÿ NAME
Global DaoQDef                  As QueryDef             'DAO: ����Ÿ QUERYDEF
Global DaoTDef                  As TableDef             'DAO: ����Ÿ TABLEDEF
Global DaoError                 As Error
Global DaoParameter             As Parameter

'Remote Data Objects(RDO) Variable
Global RdoEnv                   As rdoEnvironment       'RDO: ȯ�溯��
Global RdoDB                    As rdoConnection        'RDO: ����Ÿ���̽� NAME
Global RdoQry                   As New rdoQuery         'RDO: QUERY���� ����
Global RdoSet(9)                As rdoResultset         'RDO: QUERY �� ����Ÿ NAME
Global RdoSet1                  As rdoResultset         'RDO: QUERY �� ����Ÿ NAME
Global RdoSet2                  As rdoResultset         'RDO: QUERY �� ����Ÿ NAME
Global RdoSet3                  As rdoResultset         'RDO: QUERY �� ����Ÿ NAME
Global RdoPre                   As rdoPreparedStatement 'RDO: ����Ÿ Prepared NAME
Global RdoErr                   As rdoError             'RDO: Error Define
Global RdoResult                As rdoResultset         'RDO: Result Set

Global RecordCount              As Long                 '�о���� Record �� ����
Global strSql                   As String               'Query String ����
Global RtnNOR                   As Long                 'ExecuteSQL�� ���ϰ� - Number Of Row
Global Result                   As Integer              'SQL ���ϰ� True/False
Global GnDBOpenDAO              As Integer              'DAO Open True / False
Global GnDBOpenRDO              As Integer              'RDO Open True / False

Global GstrPassIDnumber         As String               '�����Ȯ�ν� �ʿ�
Global GstrPassProgramID        As String * 8           '�����Ȯ�ν� �ʿ�
Global GstrPassWord             As String               '�����Ȯ�ν� �ʿ�
Global GstrPassGrade            As String               '�����Ȯ�ν� �ʿ�
Global GstrPassClass            As String               '�����Ȯ�ν� �ʿ�
Global GstrPassName             As String               '�����Ȯ�ν� �ʿ�
Global GstrPassPart             As String * 2           '�����Ȯ�ν� �ʿ�
Global GstrPassDept             As String               '�����Ȯ�ν� �ʿ�
Global GstrPassRank             As Integer              '�����Ȯ�ν� �ʿ�


Public Function Convert_String(ArgStr As String) As String
   
   'ODBC�� ���Ͽ� Data�� INSERT�Ұ�� DATA�� �߰��� "'" ���� �������
   'ERROR�� ���� �̰�츦 �����ϱ����ؼ��� "'" ���� "''"������ ��ȯ�Ͽ�
   'INSERT�Ͽ��� ������ �� FUNCTION�� �����
   
    Dim nStart          As Integer
    Dim nPosition       As Integer
    Dim sReturnStr      As String
    
    If InStr(1, ArgStr, "'") = 0 Then       ' ã�°��� �������
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
    
    MsgBox "����Ÿ���̽��� OPEN �Ҽ� �����ϴ�." & vbCrLf & _
           "CAUSE - " & Error(Err), vbExclamation
    MsgBox "���α׷��� �����մϴ�.", vbInformation
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
    
    MsgBox "����Ÿ���̽��� OPEN �Ҽ� �����ϴ�." & vbCrLf & _
           "CAUSE - " & Error(Err), vbExclamation
    MsgBox "���α׷��� �����մϴ�.", vbInformation
    Screen.MousePointer = 0         'Default
    End
    
End Sub


Public Function ExecDAO(ArgQuery As String) As Integer

    Screen.MousePointer = 11            'HOURGLASS

On Error GoTo Error_Handler_EXECUTE

    RtnNOR = DaoDB.ExecuteSQL(ArgQuery)
    If RtnNOR < 1 Then
        MsgBox "����� �۾��� �����ϴ�.", vbInformation, "����"
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
    
    'JET/Data Access Objects(DAO) Library�� �̿��� ����Ÿ���̽� Connect
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
    MsgBox "����Ÿ���̽��� OPEN �Ҽ� �����ϴ�." & vbCrLf & "CAUSE - " & Error(Err), vbExclamation
    MsgBox "���α׷��� �����մϴ�.", vbInformation
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
        MsgBox "�����ڵ� - " & RdoErr.Number & vbCrLf & _
               "�����ҽ� - " & RdoErr.Source & vbCrLf & _
               "�������� - " & RdoErr.Description & vbCrLf & vbCrLf & _
               "SQL ���� - " & ArgQuery _
               , vbExclamation, "����Ÿ�۾��� ������ �߻��߽��ϴ�."
    Next RdoErr
    
End Sub
Public Sub ErrorMsgDAO(ArgQuery As String)
    
    Beep
    Set DaoError = DBEngine.Errors(0)
    MsgBox "�����ڵ� - " & DaoError.Number & Chr$(13) & _
           "�����ҽ� - " & DaoError.Source & Chr$(13) & _
           "�������� - " & DaoError.Description & Chr$(13) & Chr$(13) & _
           "SQL �� - " & ArgQuery _
           , vbExclamation, "����Ÿ�۾��� ������ �߻��߽��ϴ�."
    
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


