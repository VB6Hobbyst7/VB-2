Attribute VB_Name = "modlIF002"
Option Base 1
Option Explicit

'Public Const MCODE As String = "507"
'Public Const SOH As String = "" 'Chr(1)
'Public Const STX As String = "" 'Chr(2)
'Public Const ETX As String = "" 'Chr(3)
'Public Const EOT As String = "" 'Chr(4)
'Public Const ENQ As String = "" 'Chr(5)
'Public Const ACK As String = "" 'Chr(6)
'Public Const LF As String = vbLf 'Chr(10)
'Public Const CR As String = vbCr 'chr(13)
'Public Const NAK As String = "" 'Chr(21)
'Public Const ETB As String = "" 'Chr(23)
'Public Const RS  As String = ""

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

Global strRecordStatus      As String

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
'# Slrtrm50p(pdate : char(8) => 검사일자,
'            pmach : char(3) => 장비코드,
'            pwnof : dec(5)  => 작업번호(from),
'            pwnot : dec(5)  => 작업번호(to),
'            pwcd  : char(5) => 작업코드,
'            pgbn  : char(1) => 구분(0:양방향,1:단방향),
'*******************>>>>>>>>>>>>>    0:양방향으로 전달해야 함  <<<<<<<<<<<<<<<<<<<<<<<*******************
'            perr  : char(1) => 인증확인 및 에러코드)
'       작업번호가 입력되지 않으면 pwnof:0, pwnot:99999 으로, 작업코드가 입력되지 않으면  pwcd:‘     ’로 전송

Public Function adoExecQuery50P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String, ByVal Param4 As String, ByVal Param5 As String, ByVal Param6 As String, ByVal Param7 As String) As ADODB.Recordset
    Dim Params7 As Parameter
    
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
        '.Parameters.Append .CreateParameter("perr", adChar, adParamOutput, 1, Param7)
        Set Params7 = .CreateParameter("perr", adChar, adParamOutput, 1, Param7)
        .Parameters.Append Params7
        
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic
    
    Set adoExecQuery50P = adoRecordset
    
    Set adoRecordset = Nothing

    strRecordStatus = Params7
    
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


Public Function adoExecQuery52P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As String) As ADODB.Recordset
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
    
    Set adoExecQuery52P = adoRecordset
    
    Set adoRecordset = Nothing
    
    strRecordStatus = Params3
    
End Function

' 결과저장 ( Sp Name : slrtrm55p ) 가져오기   >> 단방향(Batch)
Public Function adoExecQuery55P(ByVal RpgName As String, ByVal Param1 As String, ByVal Param2 As String) As String
    Dim Params1, Params2, Params3 As Parameter
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = RpgName
        .CommandTimeout = 1000
        Set Params1 = .CreateParameter("pmach", adChar, adParamInput, 3, Param1)
        .Parameters.Append Params1
        Set Params2 = .CreateParameter("perr", adChar, adParamOutput, 1, Param2)
        .Parameters.Append Params2
    End With
    
    Set adoRecordset = New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open adoCommand, , adOpenStatic, adLockBatchOptimistic

    adoExecQuery55P = Params2
    
    Set adoRecordset = Nothing

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
'    adoRecordset.CursorLocation = adUseClientBatch
    
'    adoRecordset.Open adoCommand, , adOpenStatic, adLockOptimistic
      
'    Set adoTextQueryExc = Nothing
'        MsgBox "검사 할 대상자가 없습니다. 조회일자를 확인 하세요. ", vbOKOnly + vbExclamation
'        RecordChk = False
'        Exit Function
'    Else
'        Set adoTextQueryExc = AdoRs_SQL
'    End If

'    Set AdoRs_SQL = Nothing

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

    adoConnectionString = "dsn=SMART;" & _
                          "uid=" & adoLoginID & ";" & _
                          "pwd=" & adoLoginPassword & ";" & _
                          "database=" & adodefaultDatabaseName


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







