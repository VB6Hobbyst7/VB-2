VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'--------------------------------------------------------------------------------
' 기본 클래스 모듈
'
' Get_UserInfo          사용자 정보를 읽어 온다.
'
' Get_Setting           장비 정보를 읽어 온다.
'
' Get_InstrumentList    장비 리스트를 읽어 온다.

' GetEqp_Property       장비 정보를 읽어 온다.
' LetEqp_Property       장비 정보를 저장 한다.
'
' GetEQP_TestItem       장비 검사정보를 읽어 온다.
' LetEQP_TestItem       장비 검사정보를 저장 한다.
' DelEQP_TestItem       장비 검사정보를 지운다.

' SetEqp_MatchItem      장비 검사코드의 연결정보를 초기화 시킨다.
' LetEqp_MatchItem      장비 검사코드와 마스터 검사코드를 연결해서 저장
' GetEqp_MatchItem      장비 검사코드와 마스터 검사코드의 연결해서 읽어 온다.

' GetEqp_ItemList       장비 검사코드를 읽어온다.
'
' Get_TestItem          장비에서 검사가능한 코드리스트를 읽어온다.
'
' Get_StationList       워크스테이션 리스트를 읽어온다.
'
' Get_RoomList          검사실 리스트를 읽어온다.
'--------------------------------------------------------------------------------

Private AdoCn               As ADODB.Connection
Private adoRS               As ADODB.Recordset
Private Call_Procedure      As String

Public Sub SetAdoCn(ByVal new_AdoCn As ADODB.Connection)
    Set AdoCn = new_AdoCn
End Sub

Private Sub Class_Initialize()
    Call DBErrClear
End Sub

Private Sub Class_Terminate()
    Set AdoCn = Nothing
End Sub

'장비 통신 속성 읽어옴
Public Function Get_EqpProperty(ByVal strEqpCd As String) As ADODB.Recordset
    Dim strSql      As String
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_EqpProperty(ByVal strEqpCd As String) As ADODB.Recordset"
    
    strSql = "SELECT EQP_CD, EQP_NM, CALL_ID, STATION, WORK_ST, SAVE_DT, BUY_DT, TMP_LH, VENDOR, CHARGE, BUY_TEL," & _
            " PRODUCT, MODEL_NO, COM_PORT, COM_SPEED, COM_DATABIT, COM_PARITYBIT, COM_STOPBIT, COM_HANDSHAK," & _
            " COM_INPUTMOD, COM_DTR, COM_EOF, COM_NULDIS, COM_RTS, COM_ECHO, COM_IBS, COM_INLEN, COM_OBS, COM_PTR," & _
            " COM_RTH, COM_STH FROM INTERFACE001 WHERE EQP_CD = " & STS(strEqpCd)
    Set adoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_EqpProperty = adoRS
    Else
        Set Get_EqpProperty = Nothing
    End If
    
    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

'장비 통신 속성을 저장 한다.
Public Function Let_EqpProperty(ByVal Eqp_Propert As Scripting.Dictionary) As Boolean
    Dim objPstmt    As New clsStrState
    Dim strSql      As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Let_EqpProperty(ByVal Eqp_Propert As Scripting.Dictionary) As Boolean"
    
    strSql = "SELECT EQP_CD FROM INTERFACE001 WHERE EQP_CD = " & STS(Eqp_Propert.Item("EQP_CD"))

    With objPstmt
        If DBExists(AdoCn, strSql, Call_Procedure) Then
            .InitPreparedStmt "UPDATE INTERFACE001 SET EQP_NM = ?, STATION = ?, WORK_ST = ?, SAVE_DT = ?, BUY_DT = ?, TMP_LH = ?, VENDOR = ?, CHARGE = ?, BUY_TEL = ?," & _
                              " PRODUCT = ?, MODEL_NO = ?, COM_PORT = ?, COM_SPEED = ?, COM_DATABIT = ?, COM_PARITYBIT = ?, COM_STOPBIT = ?, COM_HANDSHAK = ?," & _
                              " COM_INPUTMOD = ?, COM_DTR = ?, COM_EOF = ?, COM_NULDIS = ?, COM_RTS = ?, COM_ECHO = ?, COM_IBS = ?, COM_INLEN = ?, COM_OBS = ?," & _
                              " COM_PTR = ?, COM_RTH = ?, COM_STH = ?, CALL_ID = ? WHERE EQP_CD = ?"
            .SetString 1, Eqp_Propert.Item("EQP_NM")
            .SetString 2, Eqp_Propert.Item("STATION")
            .SetString 3, Eqp_Propert.Item("WORK_ST")
            .SetString 4, Eqp_Propert.Item("SAVE_DT")
            .SetString 5, Eqp_Propert.Item("BUY_DT")
            .SetString 6, Eqp_Propert.Item("TMP_LH")
            .SetString 7, Eqp_Propert.Item("VENDOR")
            .SetString 8, Eqp_Propert.Item("CHARGE")
            .SetString 9, Eqp_Propert.Item("BUY_TEL")
            .SetString 10, Eqp_Propert.Item("PRODUCT")
            .SetString 11, Eqp_Propert.Item("MODEL_NO")
            .SetString 12, Eqp_Propert.Item("COM_PORT")
            .SetString 13, Eqp_Propert.Item("COM_SPEED")
            .SetString 14, Eqp_Propert.Item("COM_DATABIT")
            .SetString 15, Eqp_Propert.Item("COM_PARITYBIT")
            .SetString 16, Eqp_Propert.Item("COM_STOPBIT")
            .SetString 17, Eqp_Propert.Item("COM_HANDSHAK")
            .SetString 18, Eqp_Propert.Item("COM_INPUTMOD")
            .SetString 19, Eqp_Propert.Item("COM_DTR")
            .SetString 20, Eqp_Propert.Item("COM_EOF")
            .SetString 21, Eqp_Propert.Item("COM_NULDIS")
            .SetString 22, Eqp_Propert.Item("COM_RTS")
            .SetString 23, Eqp_Propert.Item("COM_ECHO")
            .SetString 24, Eqp_Propert.Item("COM_IBS")
            .SetString 25, Eqp_Propert.Item("COM_INLEN")
            .SetString 26, Eqp_Propert.Item("COM_OBS")
            .SetString 27, Eqp_Propert.Item("COM_PTR")
            .SetString 28, Eqp_Propert.Item("COM_RTH")
            .SetString 29, Eqp_Propert.Item("COM_STH")
            .SetString 30, Eqp_Propert.Item("CALL_ID")
            .SetString 31, Eqp_Propert.Item("EQP_CD")
        Else
            .InitPreparedStmt "INSERT INTO INTERFACE001 (EQP_CD, EQP_NM, STATION, WORK_ST, SAVE_DT, BUY_DT, TMP_LH, VENDOR, CHARGE, BUY_TEL," & _
                              " PRODUCT, MODEL_NO, COM_PORT, COM_SPEED, COM_DATABIT, COM_PARITYBIT, COM_STOPBIT, COM_HANDSHAK, COM_INPUTMOD," & _
                              " COM_DTR, COM_EOF, COM_NULDIS, COM_RTS, COM_ECHO, COM_IBS, COM_INLEN, COM_OBS, COM_PTR, COM_RTH, COM_STH, CALL_ID)" & _
                              " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            .SetString 1, Eqp_Propert.Item("EQP_CD")
            .SetString 2, Eqp_Propert.Item("EQP_NM")
            .SetString 3, Eqp_Propert.Item("STATION")
            .SetString 4, Eqp_Propert.Item("WORK_ST")
            .SetString 5, Eqp_Propert.Item("SAVE_DT")
            .SetString 6, Eqp_Propert.Item("BUY_DT")
            .SetString 7, Eqp_Propert.Item("TMP_LH")
            .SetString 8, Eqp_Propert.Item("VENDOR")
            .SetString 9, Eqp_Propert.Item("CHARGE")
            .SetString 10, Eqp_Propert.Item("BUY_TEL")
            .SetString 11, Eqp_Propert.Item("PRODUCT")
            .SetString 12, Eqp_Propert.Item("MODEL_NO")
            .SetString 13, Eqp_Propert.Item("COM_PORT")
            .SetString 14, Eqp_Propert.Item("COM_SPEED")
            .SetString 15, Eqp_Propert.Item("COM_DATABIT")
            .SetString 16, Eqp_Propert.Item("COM_PARITYBIT")
            .SetString 17, Eqp_Propert.Item("COM_STOPBIT")
            .SetString 18, Eqp_Propert.Item("COM_HANDSHAK")
            .SetString 19, Eqp_Propert.Item("COM_INPUTMOD")
            .SetString 20, Eqp_Propert.Item("COM_DTR")
            .SetString 21, Eqp_Propert.Item("COM_EOF")
            .SetString 22, Eqp_Propert.Item("COM_NULDIS")
            .SetString 23, Eqp_Propert.Item("COM_RTS")
            .SetString 24, Eqp_Propert.Item("COM_ECHO")
            .SetString 25, Eqp_Propert.Item("COM_IBS")
            .SetString 26, Eqp_Propert.Item("COM_INLEN")
            .SetString 27, Eqp_Propert.Item("COM_OBS")
            .SetString 28, Eqp_Propert.Item("COM_PTR")
            .SetString 29, Eqp_Propert.Item("COM_RTH")
            .SetString 30, Eqp_Propert.Item("COM_STH")
            .SetString 31, Eqp_Propert.Item("CALL_ID")
        End If
        Call DBExec(AdoCn, .GetPreparedStmt, Call_Procedure)
    End With
    
    Set objPstmt = Nothing
    Let_EqpProperty = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Let_EqpProperty = False
    Call ErrMsgProc(Call_Procedure)
End Function

'서브코드 검색
Public Function Get_TestItem(ByVal strEqp_Cd As String) As Variant
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TestItem() As ADODB.Recordset"

    strSql = "SELECT ls.Coda, SubCoda,  m.LName AS LabEName, Isnull(ls.LabSubEName, '' ) AS LabSubEName " & _
             "FROM neolis..LaLas ls " & _
             "LEFT JOIN neoemr..N_Mother m ON ls.Coda = m.Coda " & _
             "WHERE ls.Coda LIKE '" & strEqp_Cd & "'" & _
             " AND ls. UseYn = '1'" & _
             " AND DeletedYn = '0'" & _
             " AND DYN = '1'"

    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_TestItem = adoRS
    Else
        Set Get_TestItem = Nothing
    End If

'    ---------------------
'     EOF 체크
'    ---------------------
    If adoRS.EOF = True Then ' 마스터에 등록된 장비가 없음
        adoRS.Close
        Set adoRS = Nothing
        Get_TestItem = Null
        Exit Function
    Else
        Get_TestItem = adoRS.GetRows ' SP실행 결과를 반환한다.
        adoRS.Close
        Set adoRS = Nothing
    End If
    
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Get_TestItem = Null
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 리스트
Public Function Get_InstrumentList() As Variant

    Dim strSql      As String
'Set AdoCmd = New ADODB.Command

On Error GoTo ErrTrap
    Call_Procedure = "clsCommon - Public Function  Get_InstrumentList() As ADODB.Recordset"

    strSql = "SELECT MachineCode, MachineName FROM neolis..LabMachinesource " & _
             "ORDER BY MachineName"

    Set adoRS = New ADODB.Recordset

    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_InstrumentList = adoRS
    Else
        Set Get_InstrumentList = Nothing
    End If

    '---------------------
    ' EOF 체크
    '---------------------
    If adoRS.EOF = True Then ' 마스터에 등록된 장비가 없음
        adoRS.Close
        Set adoRS = Nothing
        Get_InstrumentList = Null
        Exit Function
    Else
        Get_InstrumentList = adoRS.GetRows ' SP실행 결과를 반환한다.
        adoRS.Close
        Set adoRS = Nothing
    End If

    Exit Function

ErrTrap:
    If Not adoRS Is Nothing Then Set adoRS = Nothing
    Err.Raise Err.Number, Err.Description

End Function

'워크 스태이션 리스트
Public Function Get_StationList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_StationList() As ADODB.Recordset"
    strSql = "SELECT PARTCODE AS STACD , PARTNAME AS STANM FROM neoemr..Labpartsource ORDER BY PARTCODE"

    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_StationList = adoRS
    Else
        Set Get_StationList = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function


