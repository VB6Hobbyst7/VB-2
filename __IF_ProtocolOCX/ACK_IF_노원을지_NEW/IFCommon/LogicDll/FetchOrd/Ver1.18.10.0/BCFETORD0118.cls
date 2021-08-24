VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCFETORD0218"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCurItemCnt As Integer
Private mPersonCnt As Integer

Private Function ConvLabnoToExpand$(ByVal sComp5$)
    ConvLabnoToExpand = Format(DateAdd("d", Val(sComp5), "1999-01-01"), "YYYYMMDD")
End Function
Public Property Get CurItemCnt() As Integer
    CurItemCnt = mCurItemCnt
End Property

Public Property Get PersonCnt() As Integer
    PersonCnt = mPersonCnt
End Property

Public Function FetchOrderNew(ByVal sMachineCd As String, ByVal ADOCNOrd As ADODB.Connection, _
                            ByVal sGbn As String, ByVal sCWDate As String, ByVal sCJDate As String, _
                            Optional ByVal sCondition As String) As String
    On Error GoTo ErrHandler

    Dim sOneRow$, sJDate$, sJGbn$, sJNo$, sPTNO$, sNm$, sSex$, sTmp$, sTSeq$
    Dim iITcnt%, i%, j%
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq

    Dim ADORS   As New ADODB.Recordset
    Dim SqlStr$
    Dim sRtnVal$

    Dim sEqCd$
    Dim aEqCd() As String
    
    Dim sTotSvrCd$
    Dim sAPTNO$, sSPRDT$, sYear$
    Dim lngTransOK&
    Dim lngTransOK1&
    Dim bTrans  As Boolean
    
    sTotSvrCd = GetTotServerCd

    aEqCd() = Split(gsMachineNm, "-")
    sEqCd = Trim(aEqCd(0))

    FetchOrderNew = ""
    
    sYear = Left(Now, 2)
    sAPTNO = Left(sCondition, 8)
'    sSPRDT = sYear & Right(sCondition, 6)          '바코드 길이 14->13 자리로 수정
    sSPRDT = sYear & Right(sCondition, 5)
    
    If Left(sEqCd, 2) = "ZU" Then
        SqlStr = "SELECT DISTINCT A.WORKCD, A.CODE " _
                & " FROM SPL_SZDATA01T A, SPL_LABACPTLST B " _
                & "WHERE A.ACPTNO = '" & sAPTNO & "' " _
                & "  AND A.BCDDATE = '" & sSPRDT & "' " _
                & "  AND A.ACPTNO = B.PTNO  " _
                & "  AND A.STATUS IN ('1','2') " _
                & "  AND A.CODE in ('" & sTotSvrCd & "') " _
                & "ORDER BY A.CODE "
            
        ADORS.Open SqlStr, ADOCNOrd, adOpenForwardOnly

        If ADORS.EOF = True Then
            ADORS.Close: Set ADORS = Nothing
            Exit Function
        Else
            sRtnVal = ADORS.GetString(adClipString, -1, "|", Chr$(3))
        End If
        ADORS.Close: Set ADORS = Nothing
        
        ADOCNOrd.BeginTrans: bTrans = True
        lngTransOK1 = 0
        
        Do
            sOneRow = GetByOneUserSymbol(sRtnVal, sRtnVal, Chr$(3))
            
            sOneRow = sOneRow & Chr$(124)
            
            sTmp = GetByOne(sOneRow, sOneRow)
            
            If sTmp = "" Then
                Exit Do
            End If
            
            If sTmp <> sEqCd Then
                 sTmp = GetByOne(sOneRow, sOneRow)          '검사코드
                If sTmp = "" Then
                    Exit Do
                End If
                
                SqlStr = "  UPDATE SPL_SZDATA01T  " _
                               & "  SET WORKCD = '" & Trim(sEqCd) & "'" _
                               & "WHERE ACPTNO = '" & Trim(sAPTNO) & "'" _
                               & "  AND BCDDATE = '" & Trim(sSPRDT) & "' " _
                               & "  AND CODE = '" & sTmp & "'" _
                               & "  AND STATUS IN ('1','2') "
                               
                ADOCNOrd.Execute SqlStr, lngTransOK
                lngTransOK1 = lngTransOK1 + lngTransOK
            End If
        Loop
        
        If lngTransOK1 = 0 Then
            ADOCNOrd.RollbackTrans
        Else
            '한 행이라도 UPDATE 된 경우
            ADOCNOrd.CommitTrans
        End If
    End If
            
    '--- 하나로재단用 ORDER 조회   1         2        3       4        5
    SqlStr = "  SELECT DISTINCT A.SPRDT, A.WORKNO, B.PTNM, B.PTSSN, A.CODE " _
            & " FROM SPL_SZDATA01T A, SPL_LABACPTLST B " _
            & "WHERE A.WORKCD = '" & sEqCd & "'" _
            & "  AND A.ACPTNO = '" & sAPTNO & "' " _
            & "  AND A.BCDDATE = '" & sSPRDT & "' " _
            & "  AND A.ACPTNO = B.PTNO  " _
            & "  AND A.STATUS IN ('1','2') " _
            & "ORDER BY A.CODE "

    ADORS.Open SqlStr, ADOCNOrd, adOpenForwardOnly

    If ADORS.EOF = True Then
        ADORS.Close: Set ADORS = Nothing
        Exit Function
    Else
        sRtnVal = ADORS.GetString(adClipString, -1, "|", Chr$(3))
    End If
    ADORS.Close: Set ADORS = Nothing

    sTmp = sRtnVal

    sJDate = GetByOne(sTmp, sTmp)
    sPTNO = GetByOne(sTmp, sTmp)
    sNm = GetByOne(sTmp, sTmp)
    sSex = GetByOne(sTmp, sTmp)
    sSex = Mid(sSex, 7, 1)
    
    If Left(sSex, 1) = "1" Then
        sSex = "M"
    ElseIf Left(sSex, 1) = "2" Then
        sSex = "F"
    End If
    
    sTSeq = ""
    iITcnt = 0

    Do
        sOneRow = GetByOneUserSymbol(sRtnVal, sRtnVal, Chr$(3))

        If sOneRow = "" Then
            Exit Do
        End If

        sOneRow = sOneRow & Chr$(124)

        Call GetByOne(sOneRow, sOneRow)     '접수일자
        Call GetByOne(sOneRow, sOneRow)     'PTID
        Call GetByOne(sOneRow, sOneRow)     '성명
        Call GetByOne(sOneRow, sOneRow)     '성별
        sTmp = GetByOne(sOneRow, sOneRow)   '검사코드

        sTmp = ConvertIFItemInfo(1, sTmp)

        If sTmp = "" Then
        Else
            iITcnt = iITcnt + 1
            sTSeq = sTSeq & sTmp & Chr$(124)
        End If
    Loop

    FetchOrderNew = sJDate & Chr(124) & sPTNO & Chr$(124) & sNm & Chr$(124) & sSex & Chr(124) & CStr(iITcnt) & Chr$(124) & sTSeq

ErrHandler:
    If Err <> 0 Then
        If bTrans = True Then
            ADOCNOrd.RollbackTrans
        End If
        
        If ADORS.State = 1 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        ViewMsg "FetchOrderNew - " & Err.Description
    End If
End Function

Public Function FetchPatInfo(ByVal sMachineCd$, ByVal sCondition$) As String
    On Error GoTo ErrHandler
        
    Dim ADORS As New ADODB.Recordset
    Dim SqlStr$
    Dim sRtnVal$
    Dim sSex$
    
    Dim sTotSvrCd   As String
    
    Dim sEqCd$
    Dim aEqCd() As String
    
    Dim sAPTNO$, sSPRDT$, sYear$

    aEqCd() = Split(gsMachineNm, "-")
    sEqCd = Trim(aEqCd(0))
        
    Call GetTestItem
    Call GetOrdRstCfg
    
    sYear = Left(Now, 2)
    sAPTNO = Left(sCondition, 8)
'    sSPRDT = sYear & Right(sCondition, 6)
    sSPRDT = sYear & Right(sCondition, 5)
    
    sTotSvrCd = GetTotServerCd
                                    '0        1        2       3
        SqlStr = "  SELECT DISTINCT A.SPRDT, A.WORKNO, B.PTNM, B.PTSSN " _
            & " FROM SPL_SZDATA01T A, SPL_LABACPTLST B " _
            & "WHERE A.WORKCD = '" & sEqCd & "'" _
            & "  AND A.ACPTNO = '" & sAPTNO & "' " _
            & "  AND A.BCDDATE = '" & sSPRDT & "' " _
            & "  AND A.ACPTNO = B.PTNO   " _
            & "  AND A.STATUS IN ('1','2') " _
            & "  AND A.CODE in ('" & sTotSvrCd & "') "
    
    ADORS.Open SqlStr, "DSN=" & gOrdCfg.sPath & ";UID=hanaromf;PWD=hanaro;", adOpenForwardOnly
    
    If ADORS.EOF = False Then
        If Mid(Trim(ADORS(3)), 7, 1) = "1" Then
            sSex = "M"
        ElseIf Mid(Trim(ADORS(3)), 7, 1) = "2" Then
            sSex = "F"
        End If
        
        FetchPatInfo = Trim(ADORS(0)) & Chr(124) & Trim(sCondition) & Chr(124) & Trim(ADORS(1)) & Chr(124) & Trim(ADORS(2)) & Chr(124) & sSex$ & Chr(124)
    Else
        FetchPatInfo = ""
    End If
    ADORS.Close: Set ADORS = Nothing
    
ErrHandler:
    If Err <> 0 Then
        If ADORS.State = 1 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        ViewMsg "FetchPatInfo - " & Err.Description
    End If
End Function
Private Function GetTotServerCd() As String

    Dim ii      As Integer
    Dim sStr    As String

    GetTotServerCd = ""
    sStr = ""

    For ii = 1 To giOriginIFItemCnt
        With gIFItem(ii)
            If Trim(.s06) <> "" Then
                If Trim(sStr) = "" Then
                    sStr = Trim(.s06)
                Else
                    sStr = sStr & "','" & Trim(.s06)
                End If
            End If
        End With
    Next ii

    For ii = 1 To giOriginCalItemCnt
        With gCalItem(ii)
            If Trim(.s03) <> "" Then
                If Trim(sStr) = "" Then
                    sStr = Trim(.s03)
                Else
                    sStr = sStr & "','" & Trim(.s03)
                End If
            End If
        End With
    Next ii

    GetTotServerCd = sStr
    
End Function

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub
Private Sub TABLE_LAYOUT()
'<검체번호로 환자정보 읽기>
'테이블:  TC201
'
'필드:
'SpcNo  CHAR(11), /* 검체번호 */
'PtID  CHAR(9), /* 등록번호 */
'PtNm  VARCHAR(20), /* 환자명 */
'Sex  CHAR(1), /* 성별('M':남, 'F':여, 'A':모두, 'E':기타) */
'Age  VARCHAR(6), /* 나이(일령) */
'SpcCd  VARCHAR(10), /* 검체코드 */
'SpcNm  VARCHAR(30), /* 검체명 */
'AllTestNm VARCHAR(800), /* 모든 검사명 */
'StatFg  CHAR(1), /* 응급여부('0':아님, '1':응급) */
'
'<검체번호로 검사정보 읽기>
'테이블:  TC301
'
'필드:
'SpcNo  CHAR(11), /* 검체번호 */
'TestCd  VARCHAR(10), /* 검사코드 */
'Seq  TINYINT, /* 순번 */

'<검사결과를 저장할 테이블>
'테이블:  TC206
'
'필드:
'SpcNo  CHAR(11), /* 검체번호 */
'TestCd  VARCHAR(10), /* 검사코드(LIS프로그램 검사코드) */
'TransDt  CHAR(8), /* 전송일(예 : 20030707) */
'TransTm  CHAR(6), /* 전송시간(예 : 091254) */
'RstVal  VARCHAR(50), /* 전송결과 */
'EquipCd  VARCHAR(10), /* 장비코드(LIS프로그램 기초코드에 등록된 검사장비코드) */
'SpcDiv  CHAR(1), /* 검체구분('G':일반) */
'ErrDes  VARCHAR(50), /* Error Description */

End Sub
