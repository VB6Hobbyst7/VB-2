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

Private Const FONT_KODIK = "고딕체"
Private Const FONT_GULLIM = "굴림체"
Private Const FONT_BATANG = "바탕체"
Private Const FONT_MOUNGJO = "명조체"
Private Const FONT_TNEWROMAN = "Times New Roman"
Private Const FONT_SERIF = "MS Serif"
Private Const FONT_ARIAL = "Arial"

Private Const FONT_HEIGHT = 200
Private Const FONT_WIDTH = 208


Public Sub SetAdoCn(ByVal new_AdoCn As ADODB.Connection)
    Set AdoCn = new_AdoCn
End Sub

Private Sub Class_Initialize()
    Call DBErrClear
End Sub

Private Sub Class_Terminate()
    Set AdoCn = Nothing
End Sub

'사용자 정보를 가지고온다.
Public Function Get_UserInfo(ByVal USERID As String) As ADODB.Recordset
    Dim strSql  As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_UserInfo(ByVal USERID As String) As ADODB.Recordset"
    
    strSql = "SELECT EMP_ID, EMP_NM, PASSWD, RENK, POWERS, REMARK FROM BASIS007" & _
            " WHERE EMP_ID = " & STS(USERID)
    
    Set adoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_UserInfo = adoRS
    Else
        Set Get_UserInfo = Nothing
    End If
    
    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

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

'장비 검사 코드만 저장
Public Function Let_EqpTestItem(ByVal EqpCD As String, ByVal Testcd As String, ByVal TestNm As String) As Boolean
    Dim objPstmt    As New clsStrState
    Dim strSql      As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function LetEqp_TestItem(ByVal EqpCD As String, ByVal TestCd As String, ByVal TestNm As String) As Boolean"
    
    strSql = "SELECT TESTCD_EQP FROM INTERFACE002 " & _
             "WHERE EQP_CD = " & STS(EqpCD) & " AND TESTCD_EQP = " & STS(Testcd)

    With objPstmt
        If DBExists(AdoCn, strSql, Call_Procedure) Then
            .InitPreparedStmt "UPDATE INTERFACE002 SET TESTNM_EQP = ? WHERE (EQP_CD = ?) AND (TESTCD_EQP = ?)"
            .SetString 1, TestNm
            .SetString 2, EqpCD
            .SetString 3, Testcd
        Else
            .InitPreparedStmt "INSERT INTO INTERFACE002 (EQP_CD, TESTCD_EQP, TESTNM_EQP)" & _
                              " VALUES(?, ?, ?)"
            .SetString 1, EqpCD
            .SetString 2, Testcd
            .SetString 3, TestNm
        End If
        Call DBExec(AdoCn, .GetPreparedStmt, Call_Procedure)
    End With
    
    Set objPstmt = Nothing
    Let_EqpTestItem = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Let_EqpTestItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검사코드 리스트
Public Function Get_EqpTestItem(ByVal EqpCD As String) As ADODB.Recordset
    Dim strSql      As String
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function GetEqp_TestItem(ByVal EqpCD As String) As ADODB.Recordset"
    
    strSql = "SELECT TESTCD_EQP, TESTNM_EQP FROM INTERFACE002 WHERE EQP_CD = " & STS(EqpCD) & " ORDER BY TESTCD "
    Set adoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_EqpTestItem = adoRS
    Else
        Set Get_EqpTestItem = Nothing
    End If
    
    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

'장비 검사코드 지우기
Public Function Del_EqpTestItem(ByVal EqpCD As String, ByVal Testcd As String) As Boolean
    Dim strSql      As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Del_EqpTestItem(ByVal EqpCD As String, ByVal TestCd As String) As Boolean"
    
    strSql = "DELETE FROM INTERFACE002 WHERE EQP_CD = " & STS(EqpCD) & "AND TESTCD_EQP = " & STS(Testcd)
    Call DBExec(AdoCn, strSql, Call_Procedure)
    
    Del_EqpTestItem = True
Exit Function

ErrorTrap:
    Del_EqpTestItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비검사 코드의 LIS검사 코드 지우기
Public Function Set_EqpMatchItem(ByVal EqpCD As String) As Boolean
    Dim strSql      As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function SetEqp_MatchItem(ByVal EqpCd As String) As Boolean"
    
    strSql = "UPDATE INTERFACE002 SET OUT_SEQ = 0, TESTCD = '', TESTNM = '', AUTOVERIFY = '', REMARK = ''," & _
             "                        DELTA = '', DELTAGBN = '', PANICL = '', PANICH = ''" & _
             " WHERE EQP_CD = " & STS(EqpCD)
    Call DBExec(AdoCn, strSql, Call_Procedure)
    
    Set_EqpMatchItem = True
Exit Function

ErrorTrap:
    Set_EqpMatchItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검사 코드 수정
Public Function Let_EqpMatchItem(ByVal MatchItem As Scripting.Dictionary) As Boolean
    Dim objPstmt    As New clsStrState
    Dim strSql      As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function LetEqp_MatchItem(ByVal MatchItem As Scripting.Dictionary) As Boolean"
    
    With objPstmt
        .InitPreparedStmt "UPDATE INTERFACE002 SET TESTNM_EQP = ?, OUT_SEQ = ?, TESTCD = ?, TESTNM = ?, AUTOVERIFY = ?, REMARK = ?," & _
                          "                        REFL = ?, REFH = ?, DELTA = ?, DELTAGBN = ?, PANICL = ?, PANICH = ?" & _
                         " WHERE (EQP_CD = ?) AND (TESTCD_EQP = ?)"
        .SetString 1, MatchItem.Item("TESTNM_EQP")
        .SetString 2, MatchItem.Item("OUT_SEQ"), 1
        .SetString 3, MatchItem.Item("TESTCD")
        .SetString 4, MatchItem.Item("TESTNM")
        .SetString 5, MatchItem.Item("AUTOVERIFY")
        .SetString 6, MatchItem.Item("REMARK")
        .SetString 7, MatchItem.Item("REFL")
        .SetString 8, MatchItem.Item("REFH")
        .SetString 9, MatchItem.Item("DELTA")
        .SetString 10, MatchItem.Item("DELTAGBN")
        .SetString 11, MatchItem.Item("PANICL")
        .SetString 12, MatchItem.Item("PANICH")
        .SetString 13, MatchItem.Item("EQP_CD")
        .SetString 14, MatchItem.Item("TESTCD_EQP")
        Call DBExec(AdoCn, .GetPreparedStmt, Call_Procedure)
        'Call DBExec(AdoCn_Jet, strSql, "Public Function Del_OldData() As Boolean")
    End With
    
    Set objPstmt = Nothing
    Let_EqpMatchItem = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Let_EqpMatchItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검사 코드와 LIS 검사 코드 연결 리스트
Public Function Get_TestItemList(EqpCD As String) As ADODB.Recordset
    
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TestItemList() As ADODB.Recordset"
    strSql = "SELECT RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             " FROM INTERFACE002" & _
             " WHERE (EQP_CD = " & STS(EqpCD) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " ORDER BY TESTCD "

    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_TestItemList = adoRS
    Else
        Set Get_TestItemList = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

'LIS 장비별 검사 코드 리스트
Public Function Get_TestItem(ByVal strEqp_Cd As String) As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TestItem() As ADODB.Recordset"
    
'    strSql = "SELECT DISTINCT LIMAS002.TESTCD, LIMAS001.ABBRNM AS TESTNM FROM LIMAS001, LIMAS002" & _
'            " WHERE  LIMAS001.TESTCD =  LIMAS002.TESTCD" & _
'            " AND LIMAS002.EQUIPCD = " & STS(strEqp_Cd) & _
'            " AND LIMAS002.DELDT = ''"
            
    strSql = "select distinct ORD_CODE, ORD_NAME as TESTNM, SPEC_CODE from SUBOR" & _
             " where PART in ('ER', 'GC', 'ID')"
    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_TestItem = adoRS
    Else
        Set Get_TestItem = Nothing
    End If
    
    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 리스트
Public Function Get_InstrumentList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_InstrumentList() As ADODB.Recordset"
    strSql = "SELECT * FROM INTERFACE001" & _
             " ORDER BY EQP_CD, EQP_NM "

    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_InstrumentList = adoRS
    Else
        Set Get_InstrumentList = Nothing
    End If
    
    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

'워크 스태이션 리스트
Public Function Get_StationList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_StationList() As ADODB.Recordset"
    strSql = "SELECT CDVAL1 AS STACD , FIELD1 AS STANM FROM LIMAS032 WHERE CDDIV = " & STS(IDX_STA) & " ORDER BY FIELD1"
    
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

'검사실 리스트
Public Function Get_RoomList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_RoomList() As ADODB.Recordset"
    strSql = "SELECT CDVAL1 AS ROOMCD , FIELD1 AS ROOMNM FROM LIMAS032 WHERE CDDIV = " & STS(IDX_ROOM) & " ORDER BY FIELD1"
    
    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_RoomList = adoRS
    Else
        Set Get_RoomList = Nothing
    End If
    
    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

'해당 장비를 사용하는 검체 리스트를 가지고 온다.
Public Function Get_SpcItem(strEqpCd As String) As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_SpcItem() As ADODB.Recordset"

    strSql = "SELECT LIMAS002.SPCCD, LIMAS032.FIELD1 FROM LIMAS002, LIMAS032" & _
             " WHERE LIMAS002.SPCCD = LIMAS032.CDVAL1" & _
               " AND LIMAS002.EQUIPCD = " & STS(INS_CODE) & _
               " AND LIMAS032.CDDIV=" & STS(IDX_SPC) & _
             " GROUP BY LIMAS002.SPCCD, LIMAS032.FIELD1" & _
             " ORDER BY LIMAS002.SPCCD"

    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_SpcItem = adoRS
    Else
        Set Get_SpcItem = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검체 코드만 저장
Public Function Let_EqpSpcItem(ByVal EqpCD As String, ByVal SpcCd As String, ByVal SpcNm As String) As Boolean
    Dim objPstmt    As New clsStrState
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function LetEqp_SpcItem(ByVal EqpCD As String, ByVal SpcCd As String, ByVal SpcNm As String) As Boolean"

    strSql = "SELECT SPCCD_EQP FROM INTERFACE004 " & _
             "WHERE EQP_CD = " & STS(EqpCD) & " AND SPCCD_EQP = " & STS(SpcCd)

    With objPstmt
        If DBExists(AdoCn, strSql, Call_Procedure) Then
            .InitPreparedStmt "UPDATE INTERFACE004 SET SPCNM_EQP = ? WHERE (EQP_CD = ?) AND (SPCCD_EQP = ?)"
            .SetString 1, SpcNm
            .SetString 2, EqpCD
            .SetString 3, SpcCd
        Else
            .InitPreparedStmt "INSERT INTO INTERFACE004 (EQP_CD, SPCCD_EQP, SPCNM_EQP)" & _
                              " VALUES(?, ?, ?)"
            .SetString 1, EqpCD
            .SetString 2, SpcCd
            .SetString 3, SpcNm
        End If
        Call DBExec(AdoCn, .GetPreparedStmt, Call_Procedure)
    End With

    Set objPstmt = Nothing
    Let_EqpSpcItem = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Let_EqpSpcItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검체정보를 읽어 온다.
Public Function Get_EqpSpcItem(ByVal EqpCD As String) As ADODB.Recordset
    Dim strSql      As String
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function GetEqp_SPCItem(ByVal EqpCD As String) As ADODB.Recordset"

    strSql = "SELECT SPCCD_EQP, SPCNM_EQP FROM INTERFACE004 WHERE EQP_CD = '" & EqpCD & "' ORDER BY SPCCD_EQP"
    Set adoRS = New ADODB.Recordset

    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_EqpSpcItem = adoRS
    Else
        Set Get_EqpSpcItem = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

'장비 검체정보를 지운다.
Public Function Del_EqpSpcItem(ByVal EqpCD As String, ByVal SpcCd As String) As Boolean
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function DelEqp_SpcItem(ByVal EqpCD As String, ByVal SpcCd As String) As Boolean"

    strSql = "DELETE FROM INTERFACE004 WHERE EQP_CD = '" & EqpCD & "'AND SpcCD_EQP = '" & SpcCd & "'"
    Call DBExec(AdoCn, strSql, Call_Procedure)

    Del_EqpSpcItem = True
Exit Function

ErrorTrap:
    Del_EqpSpcItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검체코드의 연결정보를 초기화 시킨다.
Public Function Set_EqpMatchSpcItem(ByVal EqpCD As String) As Boolean
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function SetEqp_MatchSpcItem(ByVal EqpCd As String) As Boolean"

    strSql = "Delete From INTERFACE004_1 WHERE (EQP_CD = '" & EqpCD & "') "
    Call DBExec(AdoCn, strSql, Call_Procedure)

    Set_EqpMatchSpcItem = True
Exit Function

ErrorTrap:
    Set_EqpMatchSpcItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검체코드와 마스터 검체코드를 연결해서 저장
Public Function Let_EqpMatchSpcItem(ByVal MatchSpc As Scripting.Dictionary) As Boolean
    Dim objPstmt        As New clsStrState
    Dim strSql          As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function LetEqp_MatchSpcItem(ByVal MatchSpc As Scripting.Dictionary) As Boolean"

    With objPstmt
        .InitPreparedStmt "Insert into INTERFACE004_1 (EQP_CD, SpcCd_Eqp, SpcCd, SpcNm ) Values ( ?, ?, ?, ?)"
        .SetString 1, MatchSpc.Item("EQP_CD")
        .SetString 2, MatchSpc.Item("SPCCD_EQP")
        .SetString 3, MatchSpc.Item("SPCCD")
        .SetString 4, MatchSpc.Item("SPCNM")

        strSql = .GetPreparedStmt
        Call DBExec(AdoCn, strSql, Call_Procedure)
    End With

    Set objPstmt = Nothing
    Let_EqpMatchSpcItem = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Let_EqpMatchSpcItem = False
    Call ErrMsgProc(Call_Procedure)

End Function

'장비 검체코드와 마스터 검체코드의 연결해서 읽어 온다.
Public Function Get_EqpMatchSpcItem(ByVal EqpCD As String) As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function GetEqp_MatchSpcItem(ByVal EqpCD As String) As ADODB.Recordset"

    strSql = "SELECT RTRIM(LTRIM(A.EQP_CD)) AS EQP_CD, A.SPCCD_EQP, A.SPCNM_EQP, B.SPCCD, B.SPCNM" & _
            " FROM INTERFACE004 AS A, INTERFACE004_1 AS B " & _
            " WHERE (B.EQP_CD = '" & EqpCD & "') AND ((SPCCD <> '') AND (SPCCD IS NOT NULL))" & _
              " AND (A.EQP_CD = B.EQP_CD) AND (A.SPCCD_EQP = B.SPCCD_EQP)" & _
            " ORDER BY A.SPCCD_EQP"

    Set adoRS = New ADODB.Recordset

    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        Set Get_EqpMatchSpcItem = adoRS
    Else
        Set Get_EqpMatchSpcItem = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

'장비에 맞는 검체종류를 읽어온다.
Public Function Get_EqpSpcType(ByVal SpcCd As String) As String
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "Public Function GetEQP_SpcType(ByVal SpcCd As String) As String"

    strSql = "SELECT SPCCD_EQP" & _
            " FROM INTERFACE004_1" & _
            " WHERE EQP_CD = " & STS(INS_CODE) & " AND SPCCD = " & STS(SpcCd)
    Set adoRS = New ADODB.Recordset

    If GetRecordset(AdoCn, strSql, adoRS, Call_Procedure) Then
        If adoRS Is Nothing Then
            Get_EqpSpcType = ""
        Else
            If adoRS.EOF = False Then
                adoRS.MoveFirst
                Get_EqpSpcType = adoRS.Fields("SPCCD_EQP")
            Else
                Get_EqpSpcType = ""
            End If
        End If
    Else
        Get_EqpSpcType = ""
    End If

Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

Public Function PrintText(ByVal X As Double, _
                          ByVal Y As Double, _
                          ByVal Text As String, _
                          Optional ByVal strFont As String = FONT_KODIK, _
                          Optional ByVal FontSize As Integer = 9, _
                          Optional ByVal Fontbold As Boolean = False) 'Text Print

   Dim oldFontSize As Integer

   Printer.CurrentX = FONT_WIDTH * X
   Printer.CurrentY = FONT_HEIGHT * Y
   Printer.FontSize = FontSize
   Printer.FontName = strFont
   Printer.Fontbold = Fontbold

   Printer.Print Text
   
   Debug.Print Text

End Function
