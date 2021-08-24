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

' Login 정보 ( Sp Name : slrtrm10p ) 가져오기
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

' Order 요청 ( Sp Name : slrtrm50p ) 가져오기   >> 단방향
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


' Order 요청 ( Sp Name : slrtrm51p ) 가져오기   >> 양방향
' -- R : 바코드 오류
' -- M : 장비코드 오류
' -- N : 레코드 없음
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


'네오딘 바코드로 오더생성
'MCLISOLIB.PMCV007RM21
'- INPUT
'1. CHAR(3)  => 223 (장비코드)
'2. CHAR(12) => 081291234511 (바코드)   ITF 타입 12자리중  10자리사용(앞뒤 체크디지트)
'- OUTPUT
'3. CHAR(8) => 1,2...<= 일련번호 네오딘 키값
'4. CHAR(1000) => PMPM|PMPT|..... "|"<= 구분자 (검사코드7자리)    5자리와 7자리가 있음

Public Function adoExecQuery_Order(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String) As String
    Dim Params1, Params2, Params3, Params4 As Parameter
    Dim Seq, ORD
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = RpgName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, Param1)  '장비코드
        .Parameters.Append .CreateParameter("PBAR", adChar, adParamInput, 12, Param2)  '바코드번호
        .Parameters.Append .CreateParameter("SEQ", adChar, adParamOutput, 8, "")
        .Parameters.Append .CreateParameter("ORD", adChar, adParamOutput, 1000, "")
        .Execute
        Seq = .Parameters("SEQ").Value
        ORD = .Parameters("ORD").Value
    End With
                
'    워크리스트로 조회
'    Set adoCommand = New ADODB.Command
'    With adoCommand
'        .ActiveConnection = adoConnection
'        .CommandText = "SQLLIB.PMC309RMS91"
'        .CommandType = adCmdStoredProc
'        .CommandTimeout = 1000
'        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, "224")  '장비코드
'        .Parameters.Append .CreateParameter("GDT", adDecimal, adParamInput)  '바코드번호
'                                .Parameters("GDT").Precision = 8                       '-- 자릿수
'                                .Parameters("GDT").NumericScale = 0                    '-- 소숫점
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



' 결과저장 ( Sp Name : slrtrm56p ) 가져오기   >> 양방향
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

' 결과저장 ( Sp Name : slrtrm60p ) 가져오기   >> 양방향
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

'네오딘 결과전송:MCLISOLIB.PMCV027RM21
Public Function adoExecQuery_Result(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, ByVal Param6 As String, ByVal Param7 As String) As String
    Dim Params1, Params2, Params3, Params4, Params5, Params6, Params7 As Parameter
    Dim STS
        
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = RpgName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, Param1)     '장비코드
        .Parameters.Append .CreateParameter("BCD", adChar, adParamInput, 12, Param2)    '바코드
        .Parameters.Append .CreateParameter("PORD", adChar, adParamInput, 7, Param3)    '검사코드
        .Parameters.Append .CreateParameter("PSEQ", adDecimal, adParamInput)            '일련번호
                                .Parameters("PSEQ").Precision = 6                       '-- 자릿수
                                .Parameters("PSEQ").NumericScale = 0                    '-- 소숫점
                                .Parameters("PSEQ").Value = Val(Param4)
        .Parameters.Append .CreateParameter("RLT", adDecimal, adParamInput)             '수치결과
                                .Parameters("RLT").Precision = 9                        '-- 자릿수
                                .Parameters("RLT").NumericScale = 3                     '-- 소숫점
                                .Parameters("RLT").Value = Val(Param5)
        .Parameters.Append .CreateParameter("PCMT", adChar, adParamInput, 40, Param6)   '문자결과 (플래그)
        .Parameters.Append .CreateParameter("ERR", adChar, adParamOutput, 10, "")       '전송결과
        .Parameters.Append .CreateParameter("LOT", adChar, adParamInput, 12, "")    '장비로트번호
        .Parameters.Append .CreateParameter("LVL", adChar, adParamInput, 10, "")    '장비QC레벨
        .Parameters.Append .CreateParameter("MSG", adChar, adParamInput, 40, "")   '코멘트
        .Execute
        STS = Mid(.Parameters("ERR").Value, 1, 1)
    End With

    adoExecQuery_Result = STS
    Debug.Print Trim(STS)

End Function

'네오딘QC 결과전송:MCLISOLIB.PMCV027RMS1
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
        .Parameters.Append .CreateParameter("DEV", adChar, adParamInput, 3, Param1)     '장비코드
        .Parameters.Append .CreateParameter("BCD", adChar, adParamInput, 12, Param2)     '바코드
        .Parameters.Append .CreateParameter("PORD", adChar, adParamInput, 7, Param3)    '검사코드
        .Parameters.Append .CreateParameter("PSEQ", adDecimal, adParamInput)            '일련번호
                                .Parameters("PSEQ").Precision = 6                       '-- 자릿수
                                .Parameters("PSEQ").NumericScale = 0                    '-- 소숫점
                                .Parameters("PSEQ").Value = Val(Param4)
        .Parameters.Append .CreateParameter("RLT", adDecimal, adParamInput)             '수치결과
                                .Parameters("RLT").Precision = 9                        '-- 자릿수
                                .Parameters("RLT").NumericScale = 3                     '-- 소숫점
                                .Parameters("RLT").Value = Val(Param5)
        .Parameters.Append .CreateParameter("PCMT", adChar, adParamInput, 40, Param6)   '문자결과 (플래그)
        .Parameters.Append .CreateParameter("ERR", adChar, adParamOutput, 10, "")       '오류반환값
        .Parameters.Append .CreateParameter("LOT", adChar, adParamInput, 12, Param8)    '장비로트번호
        .Parameters.Append .CreateParameter("LVL", adChar, adParamInput, 10, Param9)    '장비QC레벨
        .Parameters.Append .CreateParameter("MSG", adChar, adParamInput, 40, Param10)   '코멘트
        .Execute
        STS = Mid(.Parameters("ERR").Value, 1, 1)
    End With

    adoExecQuery_QCResult = STS
    Debug.Print Trim(STS)

End Function

Public Function adoExecQuerySQL(ByVal adoParaCnt As Integer) As String
    
'*  Record Set 오브젝트를 위한 변수의 생성.
'    Set adoRecordset = New ADODB.Recordset
'*  Command 오브젝트를 위한 변수의 생성.
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

'*  Record Set 에 대한 속성의 정의.
'    adoRecordset.CursorLocation = adUseClientBatch
'    adoRecordset.Open adoCommand, , adOpenStatic, adLockReadOnly
'    adoCommand.Execute

End Function

Public Function adoExecQuerySelect(ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, ByVal Param6 As String, ByVal Param7 As String) As ADODB.Recordset
Dim ado_Comm As New ADODB.Command
Dim ado_Parm As New ADODB.Parameter

End Function


Public Function adoTextQueryExc(ByVal cRunQry)
    
'*  Record Set 오브젝트를 위한 변수의 생성.
    Set adoRecordset = New ADODB.Recordset
'*  Command 오브젝트를 위한 변수의 생성.
    Set adoCommand = New ADODB.Command

'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cRunQry
'        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
        .Execute
    End With

'*  Record Set 에 대한 속성의 정의.
    adoRecordset.CursorLocation = adUseClientBatch
    
    adoRecordset.Open adoCommand, , adOpenStatic, adLockOptimistic
      
'    Set adoTextQueryExc = Nothing
'        MsgBox "검사 할 대상자가 없습니다. 조회일자를 확인 하세요. ", vbOKOnly + vbExclamation
'        RecordChk = False
'        Exit Function
'    Else
        Set adoTextQueryExc = AdoRs_SQL
'    End If

    Set AdoRs_SQL = Nothing

End Function

Public Function adoTextQuerySQL(ByVal cRunQry) As Long
    
'*  Record Set 오브젝트를 위한 변수의 생성.
    Set adoRecordset = New ADODB.Recordset
'*  Command 오브젝트를 위한 변수의 생성.
    Set adoCommand = New ADODB.Command

'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cRunQry
'        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
    End With

'*  Record Set 에 대한 속성의 정의.
    adoRecordset.CursorLocation = adUseClientBatch
    
    adoRecordset.Open adoCommand, , adOpenStatic, adLockReadOnly

'*  최종 Affected 된 Row 수를 반환.
    adoTextQuerySQL = adoRecordset.RecordCount

End Function

'-- osw make
Public Function adoCountQuerySQL(ByVal cRunQry) As Long
    
'*  Record Set 오브젝트를 위한 변수의 생성.
    Set adoRecordset = New ADODB.Recordset
'*  Command 오브젝트를 위한 변수의 생성.
    Set adoCommand = New ADODB.Command

'*  Query Stored Procedure.
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandText = cRunQry
'        .CommandType = adCmdStoredProc
        .CommandTimeout = 1000
    End With

'*  Record Set 에 대한 속성의 정의.
    adoRecordset.CursorLocation = adUseClientBatch
    
    adoRecordset.Open adoCommand, , adOpenStatic, adLockReadOnly

'*  최종 Affected 된 Row 수를 반환.
    adoCountQuerySQL = adoRecordset.Fields(0).Value

End Function

Public Function fConnPort(cPrgName As String, ctlMSC As MSComm) As Boolean
'저장된 포트설정을 셋팅
Dim cSetPort As String

    On Error GoTo Err_ConnPort

    If ctlMSC.PortOpen = True Then ctlMSC.PortOpen = False

    cSetPort = GetSetting(cPrgName, "속성", "설정", "")
    If cSetPort <> "" Then ctlMSC.Settings = cSetPort

    cSetPort = GetSetting(cPrgName, "설정", "통신 포트", "")
    If cSetPort <> "" Then ctlMSC.CommPort = cSetPort

'    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "")
'    If Handshaking <> "" Then MSCom.Handshaking = Handshaking

    ctlMSC.PortOpen = True
    fConnPort = True

    Exit Function

Err_ConnPort:
    MsgBox "통신 포트를 설정하는 과정에서 오류가 발생했습니다.", vbCritical, "설정오류"
    fConnPort = False

End Function

Public Sub adoConnectSQLServer(ByVal adoServerName, ByVal adoLoginID, _
                               ByVal adoLoginPassword, ByVal adodefaultDatabaseName)
'*  ******************************************************************************************
'*  함 수 명 : adoConnectSQLServer
'*  인    수 : 서버명,로그인ID,로그인PASSWORD,디폴트데이터베이스
'*  작 성 자 : 백요한
'*  작 성 일 : 2000년 1월 18일
'*  역    할 :
'*  ******************************************************************************************

'*  Connection 오브젝트의 생성.
    Set adoConnection = New ADODB.Connection
    
'*  Connection 을 위한 ODBC Resource 문자열 정의.
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
'*  함 수 명 : adoDisconnectSQLServer
'*  작 성 자 : 백요한
'*  작 성 일 : 2000년 1월 18일
'*  역    할 :
'*  ******************************************************************************************
    
'*  Open 된 Connection 의 닫기.nn
    adoConnection.Close
    
'*  Open 된 Connection 의 할당된 메모리 리소드 해제하기.
    Set adoConnection = Nothing

End Sub

Public Sub adoEndQuerySQL()
    
'*  ******************************************************************************************
'*  함 수 명 : adoEndQuerySQL
'*  작 성 자 : 백요한
'*  작 성 일 : 2000년 1월 18일
'*  역    할 :
'*  ******************************************************************************************
    
'*  Open 된 Record Set 의 할당된 메모리 리소드 해제하기.
    Set adoRecordset = Nothing

'*  Open 된 Command 의 할당된 메모리 리소드 해제하기.
    Set adoCommand = Nothing

End Sub





