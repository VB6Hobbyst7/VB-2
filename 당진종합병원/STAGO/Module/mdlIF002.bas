Attribute VB_Name = "Module1"
Option Base 1
Option Explicit

Public Const MCODE As String = "507"
Public Const SOH As String = "" 'Chr(1)
Public Const STX As String = "" 'Chr(2)
Public Const ETX As String = "" 'Chr(3)
Public Const EOT As String = "" 'Chr(4)
Public Const ENQ As String = "" 'Chr(5)
Public Const ACK As String = "" 'Chr(6)
Public Const LF As String = vbLf 'Chr(10)
Public Const CR As String = vbCr 'chr(13)
Public Const NAK As String = "" 'Chr(21)
Public Const ETB As String = "" 'Chr(23)

Public adoConnection                    As ADODB.Connection
Public adoRecordset                     As ADODB.Recordset
Public adoCommand                       As ADODB.Command
Public adoConnectionString              As String
Public adoRecordsAffected               As Long

Public cSPName As String
Public cParameter() As String
Public adoParam() As Parameter
'Public cArgument As String
Public lRowCount As Long
Public lAffectedRows As Long

Public dLimit As Double
Public dPerSam As Double
Public lPw As Long

Private nCount As Integer

' Login ���� ( Sp Name : slrtrm10p ) ��������
Public Function adoExecQuery10P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String) As String
    Set adoCommand = New ADODB.Command
    Dim Params1, Params2, Params3 As Parameter
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = RpgName
        .CommandTimeout = 1000
        Set Params1 = .CreateParameter("puserid", adChar, adParamInput, 10, Param1)
        .Parameters.Append Params1
        Set Params2 = .CreateParameter("ppasswd", adChar, adParamInput, 10, Param2)
        .Parameters.Append Params2
        Set Params3 = .CreateParameter("perr", adChar, adParamOutput, 1, Param3)
        .Parameters.Append Params3
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic

    adoExecQuery10P = Params3
    
    Set adoRecordset = Nothing

End Function

' Order ��û ( Sp Name : slrtrm50p ) ��������   >> �ܹ���
Public Function adoExecQuery50P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, ByVal Param6 As String, ByVal Param7 As String) As ADODB.Recordset
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = RpgName
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("pdate", adChar, adParamInput, 8, Param1)
        .Parameters.Append .CreateParameter("pmach", adChar, adParamInput, 3, Param2)
        .Parameters.Append .CreateParameter("pwnof", adInteger, adParamInput, 5, Val(Param3))
        .Parameters.Append .CreateParameter("pwnot", adInteger, adParamInput, 5, Val(Param4))
        .Parameters.Append .CreateParameter("pwcd", adChar, adParamInput, 5, Param5)
        .Parameters.Append .CreateParameter("pgbn", adChar, adParamInput, 1, Param6)
        .Parameters.Append .CreateParameter("perr", adChar, adParamOutput, 1, Param7)
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic
    
    Set adoExecQuery50P = adoRecordset
    
    Set adoRecordset = Nothing

End Function


' Order ��û ( Sp Name : slrtrm51p ) ��������   >> �����
' -- R : ���ڵ� ����
' -- M : ����ڵ� ����
' -- N : ���ڵ� ����
' -- Y : OK

Public Function adoExecQuery51P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String) As ADODB.Recordset
    Dim Params1, Params2, Params3 As Parameter
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = RpgName
        .CommandTimeout = 1000
        Set Params1 = .CreateParameter("pbarc", adChar, adParamInput, 12, Param1)
        .Parameters.Append Params1
        Set Params2 = .CreateParameter("pmach", adChar, adParamInput, 3, Param2)
        .Parameters.Append Params2
        Set Params3 = .CreateParameter("perr", adChar, adParamOutput, 1, Param3)
        .Parameters.Append Params3
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic
    
    Set adoExecQuery51P = adoRecordset
    
    Set adoRecordset = Nothing
    
    strRecordStatus = Params3
    
End Function


'�׿��� ���ڵ�� ��������
'MCLISOLIB.PMCV007RM21
'- INPUT
'1. CHAR(3)  => 223 (����ڵ�)
'2. CHAR(12) => 081291234511 (���ڵ�)   ITF Ÿ�� 12�ڸ���  10�ڸ����(�յ� üũ����Ʈ)
'- OUTPUT
'3. CHAR(8) => 1,2...<= �Ϸù�ȣ �׿��� Ű��
'4. CHAR(1000) => PMPM|PMPT|..... "|"<= ������ (�˻��ڵ�7�ڸ�)    5�ڸ��� 7�ڸ��� ����

Public Function adoExecQuery_Order(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String) As String
    Dim Params1, Params2, Params3, Params4 As Parameter
    Dim Seq, ORD
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = RpgName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, Param1)  '����ڵ�
        .Parameters.Append .CreateParameter("PBAR", adChar, adParamInput, 12, Param2)  '���ڵ��ȣ
        .Parameters.Append .CreateParameter("SEQ", adChar, adParamOutput, 8, "")
        .Parameters.Append .CreateParameter("ORD", adChar, adParamOutput, 1000, "")
        .Execute
        Seq = .Parameters("SEQ").Value
        ORD = .Parameters("ORD").Value
    End With
                
'    ��ũ����Ʈ�� ��ȸ
'    Set adoCommand = New ADODB.Command
'    With adoCommand
'        .ActiveConnection = adoConnection
'        .CommandText = "SQLLIB.PMC309RMS91"
'        .CommandType = adCmdStoredProc
'        .CommandTimeout = 1000
'        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, "224")  '����ڵ�
'        .Parameters.Append .CreateParameter("GDT", adDecimal, adParamInput)  '���ڵ��ȣ
'                                .Parameters("GDT").Precision = 8                       '-- �ڸ���
'                                .Parameters("GDT").NumericScale = 0                    '-- �Ҽ���
'                                .Parameters("GDT").Value = Val("20120328")
'        .Parameters.Append .CreateParameter("BAR", adChar, adParamOutput, 10000, "")
'        .Parameters.Append .CreateParameter("DAT", adChar, adParamOutput, 8000, "")
'        .Parameters.Append .CreateParameter("JNO", adChar, adParamOutput, 5000, "")
'        .Parameters.Append .CreateParameter("NAM", adChar, adParamOutput, 31000, "")
'        .Parameters.Append .CreateParameter("GCD", adChar, adParamOutput, 7000, "")
'        .Parameters.Append .CreateParameter("RLT", adChar, adParamOutput, 13000, "")
'        .Parameters.Append .CreateParameter("CNT", adChar, adParamOutput, 5, "")
'        .Execute
'        Seq = .Parameters("BAR").Value
'        ORD = .Parameters("NAM").Value
'    End With
                

    adoExecQuery_Order = Trim(Seq) & "^" & Trim(ORD)
    Debug.Print Trim(Seq) & "^" & Trim(ORD)

End Function



' ������� ( Sp Name : slrtrm56p ) ��������   >> �����
Public Function adoExecQuery56P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String) As String
    Dim Params1, Params2, Params3 As Parameter
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = RpgName
        .CommandTimeout = 1000
        Set Params1 = .CreateParameter("pbarc", adChar, adParamInput, 12, Param1)
        .Parameters.Append Params1
        Set Params2 = .CreateParameter("pmach", adChar, adParamInput, 3, Param2)
        .Parameters.Append Params2
        Set Params3 = .CreateParameter("perr", adChar, adParamOutput, 1, Param3)
        .Parameters.Append Params3
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic

    adoExecQuery56P = Params3
    
    Set adoRecordset = Nothing

End Function

' ������� ( Sp Name : slrtrm60p ) ��������   >> �����
Public Function adoExecQuery60P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String) As String
    Dim Params1, Params2, Params3 As Parameter
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = RpgName
        .CommandTimeout = 1000
        Set Params1 = .CreateParameter("pdate", adChar, adParamInput, 8, Param1)
        .Parameters.Append Params1
        Set Params2 = .CreateParameter("pmcode", adChar, adParamInput, 3, Param2)
        .Parameters.Append Params2
        Set Params3 = .CreateParameter("perr", adChar, adParamOutput, 1, Param3)
        .Parameters.Append Params3
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic

    adoExecQuery60P = Params3
    
    Set adoRecordset = Nothing

End Function

'�׿��� �������:MCLISOLIB.PMCV027RM21
Public Function adoExecQuery_Result(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, ByVal Param6 As String, ByVal Param7 As String) As String
    Dim Params1, Params2, Params3, Params4, Params5, Params6, Params7 As Parameter
    Dim STS
        
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = RpgName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, Param1)     '����ڵ�
        .Parameters.Append .CreateParameter("BCD", adChar, adParamInput, 12, Param2)    '���ڵ�
        .Parameters.Append .CreateParameter("PORD", adChar, adParamInput, 7, Param3)    '�˻��ڵ�
        .Parameters.Append .CreateParameter("PSEQ", adDecimal, adParamInput)            '�Ϸù�ȣ
                                .Parameters("PSEQ").Precision = 6                       '-- �ڸ���
                                .Parameters("PSEQ").NumericScale = 0                    '-- �Ҽ���
                                .Parameters("PSEQ").Value = Val(Param4)
        .Parameters.Append .CreateParameter("RLT", adDecimal, adParamInput)             '��ġ���
                                .Parameters("RLT").Precision = 9                        '-- �ڸ���
                                .Parameters("RLT").NumericScale = 3                     '-- �Ҽ���
                                .Parameters("RLT").Value = Val(Param5)
        .Parameters.Append .CreateParameter("PCMT", adChar, adParamInput, 40, Param6)   '���ڰ�� (�÷���)
        .Parameters.Append .CreateParameter("ERR", adChar, adParamOutput, 10, "")       '���۰��
        .Parameters.Append .CreateParameter("LOT", adChar, adParamInput, 12, "")    '����Ʈ��ȣ
        .Parameters.Append .CreateParameter("LVL", adChar, adParamInput, 10, "")    '���QC����
        .Parameters.Append .CreateParameter("MSG", adChar, adParamInput, 40, "")   '�ڸ�Ʈ
        .Execute
        STS = Mid(.Parameters("ERR").Value, 1, 1)
    End With

    adoExecQuery_Result = STS
    Debug.Print Trim(STS)

End Function

'�׿���QC �������:MCLISOLIB.PMCV027RMS1
Public Function adoExecQuery_QCResult(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, _
                                      ByVal Param6 As String, ByVal Param7 As String, ByVal Param8 As String, ByVal Param9 As String, ByVal Param10 As String) As String
    Dim Params1, Params2, Params3, Params4, Params5, Params6, Params7 As Parameter
    Dim STS
        
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = RpgName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, Param1)     '����ڵ�
        .Parameters.Append .CreateParameter("BCD", adChar, adParamInput, 12, Param2)     '���ڵ�
        .Parameters.Append .CreateParameter("PORD", adChar, adParamInput, 7, Param3)    '�˻��ڵ�
        .Parameters.Append .CreateParameter("PSEQ", adDecimal, adParamInput)            '�Ϸù�ȣ
                                .Parameters("PSEQ").Precision = 6                       '-- �ڸ���
                                .Parameters("PSEQ").NumericScale = 0                    '-- �Ҽ���
                                .Parameters("PSEQ").Value = Val(Param4)
        .Parameters.Append .CreateParameter("RLT", adDecimal, adParamInput)             '��ġ���
                                .Parameters("RLT").Precision = 9                        '-- �ڸ���
                                .Parameters("RLT").NumericScale = 3                     '-- �Ҽ���
                                .Parameters("RLT").Value = Val(Param5)
        .Parameters.Append .CreateParameter("PCMT", adChar, adParamInput, 40, Param6)   '���ڰ�� (�÷���)
        .Parameters.Append .CreateParameter("ERR", adChar, adParamOutput, 10, "")       '������ȯ��
        .Parameters.Append .CreateParameter("LOT", adChar, adParamInput, 12, Param8)    '����Ʈ��ȣ
        .Parameters.Append .CreateParameter("LVL", adChar, adParamInput, 10, Param9)    '���QC����
        .Parameters.Append .CreateParameter("MSG", adChar, adParamInput, 40, Param10)   '�ڸ�Ʈ
        .Execute
        STS = Mid(.Parameters("ERR").Value, 1, 1)
    End With

    adoExecQuery_QCResult = STS
    Debug.Print Trim(STS)

End Function

Public Function adoExecQuerySQL(ByVal adoParaCnt As Integer) As String
    
'*  Record Set ������Ʈ�� ���� ������ ����.
'    Set adoRecordset = New ADODB.Recordset
'*  Command ������Ʈ�� ���� ������ ����.
    Set adoCommand = New ADODB.Command

    ReDim adoParam(adoParaCnt + 1)
    
'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cSPName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        
'        If adoParaCnt = 1 Then
'            .Parameters(0).Value = cArgument
'        Else
            For nCount = 1 To adoParaCnt
                If UCase(Trim(cParameter(nCount, 1))) = "C" Then
                    Set adoParam(nCount) = .CreateParameter(Trim(cParameter(nCount, 2)), adChar, adParamInput, Val(cParameter(nCount, 3)), Trim(cParameter(nCount, 4)))
                    .Parameters.Append adoParam(nCount)
                ElseIf UCase(Trim(cParameter(nCount, 1))) = "N" Then
                    Set adoParam(nCount) = .CreateParameter(Trim(cParameter(nCount, 2)), adNumeric, adParamInput, , Trim(cParameter(nCount, 4)))
                    .Parameters.Append adoParam(nCount)
                ElseIf UCase(Trim(cParameter(nCount, 1))) = "I" Then
                    Set adoParam(nCount) = .CreateParameter(Trim(cParameter(nCount, 2)), adInteger, adParamInput, , Trim(cParameter(nCount, 4)))
                    .Parameters.Append adoParam(nCount)
                Else
                    Set adoParam(nCount) = .CreateParameter(Trim(cParameter(nCount, 2)), adVarChar, adParamInput, Val(cParameter(nCount, 3)), Trim(cParameter(nCount, 4)))
                    .Parameters.Append adoParam(nCount)
                End If
            Next nCount
'        End If
        Set adoParam(adoParaCnt + 1) = .CreateParameter("perr", adChar, adParamOutput, 1, "")
        .Parameters.Append adoParam(adoParaCnt + 1)
        .Execute
        adoExecQuerySQL = .Parameters(adoParaCnt).Value
        
    End With

'*  Record Set �� ���� �Ӽ��� ����.
'    adoRecordset.CursorLocation = adUseClientBatch
'    adoRecordset.Open adoCommand, , adOpenStatic, adLockReadOnly
'    adoCommand.Execute

End Function

Public Function adoExecQuerySelect(ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, ByVal Param6 As String, ByVal Param7 As String) As ADODB.Recordset
Dim ado_Comm As New ADODB.Command
Dim ado_Parm As New ADODB.Parameter

End Function


Public Function adoTextQueryExc(ByVal cRunQry)
    
'*  Record Set ������Ʈ�� ���� ������ ����.
    Set adoRecordset = New ADODB.Recordset
'*  Command ������Ʈ�� ���� ������ ����.
    Set adoCommand = New ADODB.Command

'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cRunQry
'        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Execute
    End With

'*  Record Set �� ���� �Ӽ��� ����.
    adoRecordset.CursorLocation = adUseClientBatch
    
    adoRecordset.Open adoCommand, , adOpenStatic, adLockOptimistic
      
'    Set adoTextQueryExc = Nothing
'        MsgBox "�˻� �� ����ڰ� �����ϴ�. ��ȸ���ڸ� Ȯ�� �ϼ���. ", vbOKOnly + vbExclamation
'        RecordChk = False
'        Exit Function
'    Else
        Set adoTextQueryExc = AdoRs_SQL
'    End If

    Set AdoRs_SQL = Nothing

End Function

Public Function adoTextQuerySQL(ByVal cRunQry) As Long
    
'*  Record Set ������Ʈ�� ���� ������ ����.
    Set adoRecordset = New ADODB.Recordset
'*  Command ������Ʈ�� ���� ������ ����.
    Set adoCommand = New ADODB.Command

'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cRunQry
'        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
    End With

'*  Record Set �� ���� �Ӽ��� ����.
    adoRecordset.CursorLocation = adUseClientBatch
    
    adoRecordset.Open adoCommand, , adOpenStatic, adLockReadOnly

'*  ���� Affected �� Row ���� ��ȯ.
    adoTextQuerySQL = adoRecordset.RecordCount

End Function

'-- osw make
Public Function adoCountQuerySQL(ByVal cRunQry) As Long
    
'*  Record Set ������Ʈ�� ���� ������ ����.
    Set adoRecordset = New ADODB.Recordset
'*  Command ������Ʈ�� ���� ������ ����.
    Set adoCommand = New ADODB.Command

'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cRunQry
'        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
    End With

'*  Record Set �� ���� �Ӽ��� ����.
    adoRecordset.CursorLocation = adUseClientBatch
    
    adoRecordset.Open adoCommand, , adOpenStatic, adLockReadOnly

'*  ���� Affected �� Row ���� ��ȯ.
    adoCountQuerySQL = adoRecordset.Fields(0).Value

End Function

Public Function fConnPort(cPrgName As String, ctlMSC As MSComm) As Boolean
'����� ��Ʈ������ ����
Dim cSetPort As String

    On Error GoTo Err_ConnPort

    If ctlMSC.PortOpen = True Then ctlMSC.PortOpen = False

    cSetPort = GetSetting(cPrgName, "�Ӽ�", "����", "")
    If cSetPort <> "" Then ctlMSC.Settings = cSetPort

    cSetPort = GetSetting(cPrgName, "����", "��� ��Ʈ", "")
    If cSetPort <> "" Then ctlMSC.CommPort = cSetPort

'    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "")
'    If Handshaking <> "" Then MSCom.Handshaking = Handshaking

    ctlMSC.PortOpen = True
    fConnPort = True

    Exit Function

Err_ConnPort:
    MsgBox "��� ��Ʈ�� �����ϴ� �������� ������ �߻��߽��ϴ�.", vbCritical, "��������"
    fConnPort = False

End Function

Public Sub adoConnectSQLServer(ByVal adoServerName, ByVal adoLoginID, _
                               ByVal adoLoginPassword, ByVal adodefaultDatabaseName)
'*  ******************************************************************************************
'*  �� �� �� : adoConnectSQLServer
'*  ��    �� : ������,�α���ID,�α���PASSWORD,����Ʈ�����ͺ��̽�
'*  �� �� �� : �����
'*  �� �� �� : 2000�� 1�� 18��
'*  ��    �� :
'*  ******************************************************************************************

'*  Connection ������Ʈ�� ����.
    Set adoConnection = New ADODB.Connection
    
'*  Connection �� ���� ODBC Resource ���ڿ� ����.
'    adoConnectionString = "dsn=SMART;" & _
                          "server=" & adoServerName & ";" & _
                          "uid=" & adoLoginID & ";" & _
                          "pwd=" & adoLoginPassword & ";" & _
                          "database=" & adodefaultDatabaseName

    adoConnectionString = "Driver={iSeries Access ODBC Driver};" & _
                          "System=219.252.39.5;" & _
                          "Uid=" & adoLoginID & ";" & _
                          "Pwd=" & adoLoginPassword & ";"


    'Driver={iSeries Access ODBC Driver};System=219.252.39.5;Uid=ODBCUSER;Pwd=I74123;



'*  Connect By Using Active Data Object.
    With adoConnection
        .ConnectionString = adoConnectionString
'        .Properties("PROMPT") = adPromptNever
        .ConnectionTimeout = 60
        .Open
    End With
    
End Sub

Public Sub adoDisconnectSQLServer()
    
'*  ******************************************************************************************
'*  �� �� �� : adoDisconnectSQLServer
'*  �� �� �� : �����
'*  �� �� �� : 2000�� 1�� 18��
'*  ��    �� :
'*  ******************************************************************************************
    
'*  Open �� Connection �� �ݱ�.nn
    adoConnection.Close
    
'*  Open �� Connection �� �Ҵ�� �޸� ���ҵ� �����ϱ�.
    Set adoConnection = Nothing

End Sub

Public Sub adoEndQuerySQL()
    
'*  ******************************************************************************************
'*  �� �� �� : adoEndQuerySQL
'*  �� �� �� : �����
'*  �� �� �� : 2000�� 1�� 18��
'*  ��    �� :
'*  ******************************************************************************************
    
'*  Open �� Record Set �� �Ҵ�� �޸� ���ҵ� �����ϱ�.
    Set adoRecordset = Nothing

'*  Open �� Command �� �Ҵ�� �޸� ���ҵ� �����ϱ�.
    Set adoCommand = Nothing

End Sub





