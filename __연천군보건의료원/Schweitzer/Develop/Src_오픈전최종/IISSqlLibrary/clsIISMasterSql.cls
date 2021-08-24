VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMasterSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISMasterSql.cls (전주예수병원)
'   작성자  : 이상대
'   내  용  : IISMaster에서 사용하는 쿼리
'   작성일  : 2004-10-05
'   버  전  :
'       1. 1.1.9: 이상대(2005-05-18)
'       2. 1.2.1: 이상대(2005-05-30)
'-----------------------------------------------------------------------------'

Option Explicit

'------------------------------------------------------------------------------'
'   기능 : 검사코드 입력쿼리
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'       3.pExpireDt  : 폐기일
'       4.pTestNm5   : 검사명(5자리)
'       5.pTestNm10  : 검사명(10자리)
'       6.pTestNm    : 검사명(전체)
'       7.pWorkarea  : Workarea 코드
'       8.pRstType   : 결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require 유형(A: Alternative, R: Require)
'       10.pPanelFg  : 처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
'       11.pDetailFg : Detail 항목여부(Null: 없음, *:단일항목 자코드)
'       12.pRptSeq   : 출력순서
'------------------------------------------------------------------------------'
Public Function InsertTestCd(ByVal pTestCd As String, ByVal pApplyDt As String, _
                             ByVal pExpireDt As String, ByVal pTestNm5 As String, _
                             ByVal pTestNm10 As String, ByVal pTestNm As String, _
                             ByVal pWorkarea As String, ByVal pRstType As String, _
                             ByVal pRstDiv As String, ByVal pPanelFg As String, _
                             ByVal pDetailFg As String, ByVal pRptSeq As Long) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS101 & " (testcd, applydt, expdt, abbrnm5, abbrnm10, " & _
          "     testnm, workarea, rsttype, rstdiv, panelfg, detailfg, rptseq)" & _
          " VALUES ('" & pTestCd & "', '" & pApplyDt & "', " & _
          "         '" & pExpireDt & "', '" & pTestNm5 & "', " & _
          "         '" & pTestNm10 & "', '" & pTestNm & "', " & _
          "         '" & pWorkarea & "', '" & pRstType & "', " & _
          "         '" & pRstDiv & "', '" & pPanelFg & "', " & _
          "         '" & pDetailFg & "', " & CStr(pRptSeq) & ")"
    
    InsertTestCd = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 검사코드 수정쿼리
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'       3.pExpireDt  : 폐기일
'       4.pTestNm5   : 검사명(5자리)
'       5.pTestNm10  : 검사명(10자리)
'       6.pTestNm    : 검사명(전체)
'       7.pWorkarea  : Workarea 코드
'       8.pRstType   : 결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require 유형(A: Alternative, R: Require)
'       10.pPanelFg  : 처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
'       11.pDetailFg : Detail 항목여부(Null: 없음, *:단일항목 자코드)
'       12.pRptSeq   : 출력순서
'------------------------------------------------------------------------------'
Public Function UpdateTestCd(ByVal pTestCd As String, ByVal pApplyDt As String, _
                             ByVal pExpireDt As String, ByVal pTestNm5 As String, _
                             ByVal pTestNm10 As String, ByVal pTestNm As String, _
                             ByVal pWorkarea As String, ByVal pRstType As String, _
                             ByVal pRstDiv As String, ByVal pPanelFg As String, _
                             ByVal pDetailFg As String, ByVal pRptSeq As Long) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS101 & " SET expdt='" & pExpireDt & "', " & _
          "     abbrnm5='" & pTestNm5 & "', " & _
          "     abbrnm10='" & pTestNm10 & "'," & _
          "     testnm='" & pTestNm & "'," & _
          "     workarea='" & pWorkarea & "'," & _
          "     rsttype='" & pRstType & "'," & _
          "     rstdiv='" & pRstDiv & "'," & _
          "     panelfg='" & pPanelFg & "'," & _
          "     detailfg='" & pDetailFg & "'," & _
          "     rptseq=" & CStr(pRptSeq) & _
          " WHERE testcd='" & pTestCd & "' AND applydt='" & pApplyDt & "'"
    
    UpdateTestCd = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 검사코드 삭제쿼리
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'------------------------------------------------------------------------------'
Public Function DeleteTestCd(ByVal pTestCd As String, ByVal pApplyDt As String) As String
    Dim SQL As String

    SQL = " DELETE FROM " & TIIS101 & _
          " WHERE testcd='" & pTestCd & "' AND applydt='" & pApplyDt & "'"
    
    DeleteTestCd = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 지정검체 입력쿼리
'   인수 :
'       1.pTestCd       : 검사코드
'       2.pSpcCd        : 검체코드
'       3.pApplyDt      : 적용일
'       4.pExpireDt     : 폐기일
'       5.pUnit         : 결과단위
'       6.pAvalVal      : 유효숫자
'       7.pPanicFg      : Panic Check (0:No, 1:Yes)
'       8.pPanicFrVal   : Panic From Value
'       9.pPanicToVal   : Panic To Value
'       10.pDeltaFg     : Delta Check (0:No, 1:Yes)
'       11.pDeltaFrVal  : Panic From Value
'       12.pDeltaToVal  : Panic To Value
'------------------------------------------------------------------------------'
Public Function InsertTSpcCd(ByVal pTestCd As String, ByVal pSpcCd As String, _
                             ByVal pApplyDt As String, ByVal pExpireDt As String, _
                             ByVal pUnit As String, ByVal pAvalVal As Long, _
                             ByVal pPanicFg As String, ByVal pPanicFrVal As Single, _
                             ByVal pPanicToVal As Single, ByVal pDeltaFg As String, _
                             ByVal pDeltaFrVal As Long, ByVal pDeltaToVal As Long) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS102 & " (testcd, spccd, applydt, expdt, rstunit, avalval," & _
          "     panicfg, panicfrval, panictoval, deltafg, deltaval, deltaval2)" & _
          " VALUES ('" & pTestCd & "', '" & pSpcCd & "', " & _
          "         '" & pApplyDt & "', '" & pExpireDt & "', " & _
          "         '" & pUnit & "', " & pAvalVal & ", " & _
          "         '" & pPanicFg & "', " & pPanicFrVal & ", " & _
          "          " & pPanicToVal & ", '" & pDeltaFg & "', " & _
          "          " & pDeltaFrVal & ", " & pDeltaToVal & ")"
    
    InsertTSpcCd = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 지정검체 수정쿼리
'   인수 :
'       1.pTestCd       : 검사코드
'       2.pSpcCd        : 검체코드
'       3.pApplyDt      : 적용일
'       4.pExpireDt     : 폐기일
'       5.pUnit         : 결과단위
'       6.pAvalVal      : 유효숫자
'       7.pPanicFg      : Panic Check (0:No, 1:Yes)
'       8.pPanicFrVal   : Panic From Value
'       9.pPanicToVal   : Panic To Value
'       10.pDeltaFg     : Delta Check (0:No, 1:Yes)
'       11.pDeltaFrVal  : Panic From Value
'       12.pDeltaToVal  : Panic To Value
'------------------------------------------------------------------------------'
Public Function UpdateTSpcCd(ByVal pTestCd As String, ByVal pSpcCd As String, _
                             ByVal pApplyDt As String, ByVal pExpireDt As String, _
                             ByVal pUnit As String, ByVal pAvalVal As Long, _
                             ByVal pPanicFg As String, ByVal pPanicFrVal As Single, _
                             ByVal pPanicToVal As Single, ByVal pDeltaFg As String, _
                             ByVal pDeltaFrVal As Long, ByVal pDeltaToVal As Long) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS102 & " SET expdt='" & pExpireDt & "', " & _
          "     rstunit='" & pUnit & "', avalval=" & pAvalVal & ", " & _
          "     panicfg='" & pPanicFg & "', panicfrval=" & pPanicFrVal & ", " & _
          "     panictoval=" & pPanicToVal & ", deltafg='" & pDeltaFg & "', " & _
          "     deltaval=" & pDeltaFrVal & ", deltaval2=" & pDeltaToVal & _
          " WHERE testcd='" & pTestCd & "' AND spccd='" & pSpcCd & "' AND applydt='" & pApplyDt & "'"
    
    UpdateTSpcCd = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 지정검체 삭제쿼리
'   인수 :
'       1.pTestCd       : 검사코드
'       2.pSpcCd        : 검체코드
'       3.pApplyDt      : 적용일
'------------------------------------------------------------------------------'
Public Function DeleteTSpcCd(ByVal pTestCd As String, ByVal pSpcCd As String, _
                             ByVal pApplyDt As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS102 & _
          " WHERE testcd='" & pTestCd & "' AND spccd='" & pSpcCd & "'" & _
          "     AND applydt='" & pApplyDt & "'"
    
    DeleteTSpcCd = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 참고치 입력쿼리
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'       7.pExpireDt : 폐기일
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10.pRefCd   : Alpha결과 참고치
'------------------------------------------------------------------------------'
Public Function InsertReference(ByVal pTestCd As String, ByVal pSpcCd As String, _
                              ByVal pSex As String, ByVal pAgeFr As Long, _
                              ByVal pAgeTo As Long, ByVal pApplyDt As String, _
                              ByVal pExpireDt As String, ByVal pRefFrVal As Single, _
                              ByVal pRefToVal As Single, ByVal pRefCd As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS103 & " (testcd, spccd, applysex, agefrom, ageto, " & _
          "     applydt, expdt, refvalfrom, refvalto, refcd)" & _
          " VALUES ('" & pTestCd & "', '" & pSpcCd & "', " & _
          "         '" & pSex & "', " & CStr(pAgeFr) & ", " & _
          "          " & CStr(pAgeTo) & ", " & pApplyDt & ", " & _
          "         '" & pExpireDt & "', " & CStr(pRefFrVal) & ", " & _
          "          " & CStr(pRefToVal) & ", '" & pRefCd & "')"
    
    InsertReference = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 참고치 수정쿼리
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'       7.pExpireDt : 폐기일
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10.pRefCd   : Alpha결과 참고치
'------------------------------------------------------------------------------'
Public Function UpdateReference(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSex As String, ByVal pAgeFr As Long, _
                    ByVal pAgeTo As Long, ByVal pApplyDt As String, _
                    ByVal pExpireDt As String, ByVal pRefFrVal As Single, _
                    ByVal pRefToVal As Single, ByVal pRefCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS103 & " SET expdt='" & pExpireDt & "', " & _
          "     refvalfrom=" & CStr(pRefFrVal) & ", refvalto=" & CStr(pRefToVal) & ", " & _
          "     refcd='" & pRefCd & "'" & _
          " WHERE testcd='" & pTestCd & "' AND spccd='" & pSpcCd & "'" & _
          "     AND applysex='" & pSex & "' AND agefrom=" & CStr(pAgeFr) & _
          "     AND ageto=" & CStr(pAgeTo) & " AND applydt='" & pApplyDt & "'"
    
    UpdateReference = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 참고치 삭제쿼리
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'------------------------------------------------------------------------------'
Public Function DeleteReference(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSex As String, ByVal pAgeFr As Long, _
                    ByVal pAgeTo As Long, ByVal pApplyDt As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS103 & _
          " WHERE testcd='" & pTestCd & "' AND spccd='" & pSpcCd & "'" & _
          "     AND applysex='" & pSex & "' AND agefrom=" & CStr(pAgeFr) & _
          "     AND ageto=" & CStr(pAgeTo) & " AND applydt='" & pApplyDt & "'"
    
    DeleteReference = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 리스트 조회쿼리(검사코드, 검사명만 조회)
'-----------------------------------------------------------------------------'
Public Function SelectTestCd() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm5 AS TESTNM5" & _
          " FROM " & TIIS101 & _
          " ORDER BY testcd"
    
    SelectTestCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 리스트 조회쿼리(모든정보 조회)
'   인수 :
'       1.pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function SelectTestCdX(ByVal pTestCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT applydt AS APPLYDT, expdt AS EXPIREDT, abbrnm5 AS TESTNM5," & _
          "     abbrnm10 AS TESTNM10, testnm AS TESTNM, workarea AS WORKAREA," & _
          "     rsttype AS RSTTYPE, rstdiv AS RSTDIV, panelfg AS PANELFG," & _
          "     detailfg AS DETAILFG, rptseq AS RPTSEQ" & _
          " FROM " & TIIS101 & _
          " WHERE testcd='" & pTestCd & "'" & _
          " ORDER BY applydt"
    
    SelectTestCdX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사항목의 지정검체 정보 조회쿼리
'   인수 :
'       1.pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function SelectTSpcCd(ByVal pTestCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.spccd AS SPCCD, b.field3 AS SPCNM, a.applydt AS APPLYDT," & _
          "     a.expdt AS EXPIREDT, a.rstunit AS UNIT, a.avalval AS AVALVAL," & _
          "     a.panicfg AS PANICFG, a.panicfrval AS PANICFRVAL, a.panictoval AS PANICTOVAL," & _
          "     a.deltafg AS DELTAFG, a.deltaval AS DELTAFRVAL, a.deltaval2 AS DELTATOVAL" & _
          " FROM " & TIIS102 & " a, " & TIIS002 & " b" & _
          " WHERE a.testcd='" & pTestCd & "' AND b.cdindex='" & CSPCCD & "'" & _
          "     AND a.spccd=b.cdval1" & _
          " ORDER BY a.spccd, a.applydt"
    
    SelectTSpcCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사항목, 검체코드의 참고치 정보를 조회쿼리
'   인수 :
'       1.pTestCd : 검사코드
'       2.pSpcCd  : 검체코드
'-----------------------------------------------------------------------------'
Public Function SelectReference(ByVal pTestCd As String, ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT applysex AS SEX, agefrom AS AGEFR, ageto AS AGETO, applydt AS APPLYDT, " & _
          "     expdt AS EXPIREDT, refvalfrom AS REFFRVAL, refvalto AS REFTOVAL, refcd AS REFCD" & _
          " FROM " & TIIS103 & _
          " WHERE testcd='" & pTestCd & "' AND spccd ='" & pSpcCd & "'" & _
          "     AND applysex IN ('M', 'F', 'B', 'U')" & _
          " ORDER BY applydt, applysex, agefrom, ageto"
    
    SelectReference = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 해당 검사코드, 검체코드, 성별, 나이에 해당하는 참고치 조회쿼리
'   인수 :
'       - pTestCd : 검사코드
'       - pSpcCd  : 검체코드
'       - pSex    : 성별
'       - pAge    : 일령
'------------------------------------------------------------------------------'
Public Function SelectReferenceX(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSex As String, ByVal pAge As Long) As String
    Dim SQL As String
    
    SQL = " SELECT a.refvalfrom AS REFFRVAL, a.refvalto AS REFTOVAL, a.refcd AS REFCD" & _
          " FROM " & TIIS103 & " a" & _
          " WHERE a.testcd='" & pTestCd & "' AND a.spccd='" & pSpcCd & "' AND a.applysex='" & pSex & "'" & _
          "     AND a.agefrom<=" & CStr(pAge) & " AND a.ageto>=" & CStr(pAge) & _
          "     AND a.applydt=(SELECT MAX(applydt) FROM " & TIIS103 & _
          "                    WHERE testcd=a.testcd AND spccd=a.spccd AND applysex=a.applysex" & _
          "                         AND agefrom=a.agefrom AND ageto=a.ageto)" & _
          "     AND (a.expdt IS NULL OR a.expdt='')"

    SelectReferenceX = SQL
End Function

'------------------------------------------------------------------------------'
'   기능 : 해당 검사코드, 검체코드, 성별, 나이에 해당하는 Panic수치 조회쿼리
'   인수 :
'       - pTestCd : 검사코드
'       - pSpcCd  : 검체코드
'       - pSex    : 성별
'       - pAge    : 일령
'------------------------------------------------------------------------------'
Public Function SelectPanic(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSex As String, ByVal pAge As Long) As String
    Dim SQL As String
    
    '## NOTE: QUERY CHECK
    SQL = " SELECT a.panicfrval AS PANICFRVAL, a.panictoval AS PANICTOVAL" & _
          " FROM " & TIIS103 & " a" & _
          " WHERE a.testcd='" & pTestCd & "' AND a.spccd='" & pSpcCd & "' AND a.applysex='" & pSex & "'" & _
          "     AND a.agefrom<=" & CStr(pAge) & " AND a.ageto>=" & CStr(pAge) & _
          "     AND a.applydt=(SELECT MAX(applydt) FROM " & TIIS103 & _
          "                    WHERE testcd=a.testcd AND spccd=a.spccd AND applysex=a.applysex" & _
          "                         AND agefrom=a.agefrom AND ageto=a.ageto)" & _
          "     AND (a.expdt IS NULL OR a.expdt='')"

    SelectPanic = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체 입력쿼리
'   인수 :
'       1.pSpcCd     : 검체코드
'       2.pSpcNm     : 검체명(약어)
'       3.pSpcFullNm : 검체명(전체)
'       4.pSpcBarNm  : 검체명(바코드출력)
'-----------------------------------------------------------------------------'
Public Function InsertSpcCd(ByVal pSpcCd As String, ByVal pSpcNm As String, _
                            ByVal pSpcFullNm As String, ByVal pSpcBarNm As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS002 & " (cdindex, cdval1, field3, field4, field5)" & _
          " VALUES ('" & CSPCCD & "', '" & pSpcCd & "', " & _
          "         '" & pSpcNm & "', '" & pSpcFullNm & "', " & _
          "         '" & pSpcBarNm & "')"
    
    InsertSpcCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체 수정쿼리
'   인수 :
'       1.pSpcCd     : 검체코드
'       2.pSpcNm     : 검체명(약어)
'       3.pSpcFullNm : 검체명(전체)
'       4.pSpcBarNm  : 검체명(바코드출력)
'-----------------------------------------------------------------------------'
Public Function UpdateSpcCd(ByVal pSpcCd As String, ByVal pSpcNm As String, _
                            ByVal pSpcFullNm As String, ByVal pSpcBarNm As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS002 & " SET field3='" & pSpcNm & "', " & _
          "     field4='" & pSpcFullNm & "', " & _
          "     field5='" & pSpcBarNm & "'" & _
          " WHERE cdindex='" & CSPCCD & "' AND cdval1='" & pSpcCd & "'"
    
    UpdateSpcCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체 삭제쿼리
'   인수 :
'       1.pSpcCd     : 검체코드
'-----------------------------------------------------------------------------'
Public Function DeleteSpcCd(ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS002 & _
          " WHERE cdindex='" & CSPCCD & "' AND cdval1='" & pSpcCd & "'"
    
    DeleteSpcCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectSpcCd() As String
    Dim SQL As String
    
    SQL = " SELECT cdval1 AS SPCCD, field3 AS SPCNM," & _
          "        field4 AS SPCFULLNM, field5 AS SPCBARNM" & _
          " FROM " & TIIS002 & _
          " WHERE cdindex='" & CSPCCD & "'" & _
          " ORDER BY cdval1"
    
    SelectSpcCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검체코드에 대한 검체명 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectSpcNm(ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT field3 AS SPCNM" & _
          " FROM " & TIIS002 & _
          " WHERE cdindex='" & CSPCCD & "'"
    
    SelectSpcNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 입력쿼리
'   인수 :
'       1.pWaCd         : Workarea 코드
'       2.pWaEngNm      : Workarea 이름(영어)
'       3.pWaKorNm      : Workarea 이름(한글)
'       4.pWaNumRule    : 접수번호 부여기준
'-----------------------------------------------------------------------------'
Public Function InsertWorkarea(ByVal pWaCd As String, ByVal pWaEngNm As String, _
                               ByVal pWaKorNm As String, ByVal pWaNumRule As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS002 & " (cdindex, cdval1, field1, field2, field3)" & _
          " VALUES ('" & CWACD & "', '" & pWaCd & "', " & _
          "         '" & pWaEngNm & "', '" & pWaNumRule & "', " & _
          "         '" & pWaKorNm & "')"
    
    InsertWorkarea = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 수정쿼리
'   인수 :
'       1.pWaCd         : Workarea 코드
'       2.pWaEngNm      : Workarea 이름(영어)
'       3.pWaKorNm      : Workarea 이름(한글)
'       4.pWaNumRule    : 접수번호 부여기준
'-----------------------------------------------------------------------------'
Public Function UpdateWorkarea(ByVal pWaCd As String, ByVal pWaEngNm As String, _
                               ByVal pWaKorNm As String, ByVal pWaNumRule As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS002 & " SET field1='" & pWaEngNm & "', " & _
          "     field2='" & pWaNumRule & "', " & _
          "     field3='" & pWaKorNm & "'" & _
          " WHERE cdindex='" & CWACD & "' AND cdval1='" & pWaCd & "'"
    
    UpdateWorkarea = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 삭제쿼리
'   인수 :
'       1.pWaCd         : Workarea 코드
'-----------------------------------------------------------------------------'
Public Function DeleteWorkarea(ByVal pWaCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS002 & _
          " WHERE cdindex='" & CWACD & "' AND cdval1='" & pWaCd & "'"
    
    DeleteWorkarea = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectWorkarea() As String
    Dim SQL As String
    
    SQL = " SELECT cdval1 AS WACD, field1 AS WAENGNM, field2 AS WANUMRULE, field3 AS WAKORNM" & _
          " FROM " & TIIS002 & _
          " WHERE cdindex='" & CWACD & "'" & _
          " ORDER BY cdval1"
    
    SelectWorkarea = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목을 입력 쿼리
'   인수 :
'       1.pParentCd : 모코드
'       2.pChildCd  : 자코드
'-----------------------------------------------------------------------------'
Public Function InsertDetailCd(ByVal pParentCd As String, ByVal pChildCd As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS001 & " (cdindex, cdval1, cdval2)" & _
          " VALUES ('" & CDETAILCD & "', '" & pParentCd & "', " & _
          "         '" & pChildCd & "')"
    
    InsertDetailCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목을 삭제 쿼리
'   인수 :
'       1.pParentCd : 모코드
'       2.pChildCd  : 자코드
'-----------------------------------------------------------------------------'
Public Function DeleteDetailCd(ByVal pParentCd As String, ByVal pChildCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS001 & _
          " WHERE cdindex='" & CDETAILCD & "' AND cdval1='" & pParentCd & "'" & _
          "     AND cdval2='" & pChildCd & "'"
    
    DeleteDetailCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목의 모코드에 등록된 자코드를 모두 삭제 쿼리
'   인수 :
'       1.pParentCd : 모코드
'-----------------------------------------------------------------------------'
Public Function DeleteDetailCdAll(ByVal pParentCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS001 & _
          " WHERE cdindex='" & CDETAILCD & "' AND cdval1='" & pParentCd & "'"
    
    DeleteDetailCdAll = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목 모코드 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectDetailParentCd() As String
    Dim SQL As String
    
    SQL = " SELECT a.cdval1 AS TESTCD, b.abbrnm10 AS TESTNM" & _
          " FROM " & TIIS001 & " a, " & TIIS101 & " b" & _
          " WHERE cdindex='" & CDETAILCD & "' AND a.cdval1=b.testcd" & _
          " GROUP BY cdval1, b.abbrnm10" & _
          " ORDER BY cdval1"
    
    SelectDetailParentCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목 자코드 리스트 조회쿼리
'   인수 :
'       1.pParentCd : 모코드
'-----------------------------------------------------------------------------'
Public Function SelectDetailChildCd(ByVal pParentCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.cdval2 AS TESTCD, b.abbrnm10 AS TESTNM" & _
          " FROM " & TIIS001 & " a, " & TIIS101 & " b" & _
          " WHERE cdindex='" & CDETAILCD & "' AND cdval1='" & pParentCd & "'" & _
          "     AND a.cdval2=b.testcd" & _
          " GROUP BY cdval2, b.abbrnm10" & _
          " ORDER BY cdval2"
    
    SelectDetailChildCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 해당 모코드의 코드명을 조회쿼리
'   인수 :
'       1.pParentCd : 모코드
'-----------------------------------------------------------------------------'
Public Function SelectDetailParentNm(ByVal pParentCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE testcd='" & pParentCd & "' AND panelfg='D'" & _
          "     AND (expdt is null OR expdt='')"
    
    SelectDetailParentNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 해당 자코드를 코드명을 조회쿼리
'   인수 :
'       1.pChildCd : 자코드
'-----------------------------------------------------------------------------'
Public Function SelectDatailChildNm(ByVal pChildCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE testcd='" & pChildCd & "' AND detailfg='*'" & _
          "     AND (expdt is null OR expdt='')"
    
    SelectDatailChildNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 상세 모코드 리스트를 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectDetailParentCdX() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm10 TESTNM " & _
          " FROM " & TIIS101 & " WHERE panelfg='D' AND (expdt is null OR expdt='')" & _
          " ORDER BY testcd"
    
    SelectDetailParentCdX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 상세 자코드 리스트를 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectDetailChildCdX() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm10 TESTNM " & _
          " FROM " & TIIS101 & " WHERE detailfg='*' AND (expdt is null OR expdt='')" & _
          " ORDER BY testcd"
    
    SelectDetailChildCdX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 입력 쿼리
'   인수 :
'       1.pParentCd : 그룹코드
'       2.pSeq      : SEQ
'       3.pChildCd  : 상세항목
'-----------------------------------------------------------------------------'
Public Function InsertPanelCd(ByVal pParentCd As String, ByVal pSeq As String, _
                              ByVal pChildCd As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS001 & " (cdindex, cdval1, cdval2, field1)" & _
          " VALUES ('" & CPANELCD & "', '" & pParentCd & "', " & _
          "         '" & pSeq & "', '" & pChildCd & "')"
    
    InsertPanelCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 수정 쿼리
'   인수 :
'       1.pParentCd : 그룹코드
'       2.pSeq      : SEQ
'       3.pChildCd  : 상세항목
'-----------------------------------------------------------------------------'
Public Function UpdatePanelCd(ByVal pParentCd As String, ByVal pSeq As String, _
                              ByVal pChildCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS001 & " SET field1='" & pChildCd & "'" & _
          " WHERE cdindex='" & CPANELCD & "' AND cdval1='" & pParentCd & "'" & _
          "     AND cdval2='" & pSeq & "'"
    
    UpdatePanelCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 삭제 쿼리
'   인수 :
'       1.pParentCd : 그룹코드
'       2.pSeq      : SEQ
'-----------------------------------------------------------------------------'
Public Function DeletePanelCd(ByVal pParentCd As String, ByVal pSeq As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS001 & _
          " WHERE cdindex='" & CPANELCD & "' AND cdval1='" & pParentCd & "'" & _
          "     AND cdval2='" & pSeq & "'"
    
    DeletePanelCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드에 포함된 모든 상세항목 삭제 쿼리
'   인수 :
'       1.pParentCd : 그룹코드
'-----------------------------------------------------------------------------'
Public Function DeletePanelCdAll(ByVal pParentCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS001 & _
          " WHERE cdindex='" & CPANELCD & "' AND cdval1='" & pParentCd & "'"
    
    DeletePanelCdAll = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 리스트 조회 쿼리
'-----------------------------------------------------------------------------'
Public Function SelectPanelList() As String
    Dim SQL As String

    SQL = " SELECT a.cdval1 AS TESTCD, b.abbrnm10 AS TESTNM" & _
          " FROM " & TIIS001 & " a, " & TIIS101 & " b" & _
          " WHERE a.cdindex='" & CPANELCD & "' AND a.cdval1=b.testcd" & _
          " GROUP BY a.cdval1, b.abbrnm10" & _
          " ORDER BY cdval1"
    
    SelectPanelList = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드에 포함된 상세항목 조회 쿼리
'   인수 :
'       1.pParentCd : 그룹코드
'-----------------------------------------------------------------------------'
Public Function SelectPanelChildList(ByVal pParentCd As String) As String
    Dim SQL As String

    SQL = " SELECT a.cdval2 AS SEQ, a.field1 AS TESTCD, b.abbrnm10 AS TESTNM" & _
          " FROM " & TIIS001 & " a, " & TIIS101 & " b" & _
          " WHERE a.cdindex='" & CPANELCD & "' AND a.cdval1='" & pParentCd & "'" & _
          "     AND a.field1=b.testcd" & _
          " ORDER BY cdval2"
    
    SelectPanelChildList = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 그룹코드의 코드명 조회
'   인수 :
'       1.pParentCd : 그룹코드
'-----------------------------------------------------------------------------'
Public Function SelectPanelNm(ByVal pParentCd As String) As String
    Dim SQL As String

    SQL = " SELECT abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE testcd='" & pParentCd & "' AND panelfg='G'" & _
          "     AND (expdt is null OR expdt='')"
    
    SelectPanelNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 상세항목 코드의 코드명 조회
'   인수 :
'       1.pChildCd : 상세코드
'-----------------------------------------------------------------------------'
Public Function SelectPanelChildNm(ByVal pChildCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE testcd='" & pChildCd & "' AND (panelfg='D' or panelfg is null or panelfg='')" & _
          "     AND (detailfg is null OR detailfg='') AND (expdt is null OR expdt='')"
    
    SelectPanelChildNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 그룹코드 리스트 조회
'-----------------------------------------------------------------------------'
Public Function SelectPanelParentCd() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE panelfg='G' AND (expdt is null OR expdt='')" & _
          " ORDER BY testcd"
    
    SelectPanelParentCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 상세모코드, 단일항목 리스트 조회
'-----------------------------------------------------------------------------'
Public Function SelectPanelChildCd() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE (panelfg='D' OR panelfg is null OR panelfg='')" & _
          "     AND (detailfg is null OR detailfg='')" & _
          "     AND (expdt is null OR expdt='')" & _
          " ORDER BY testcd"
    
    SelectPanelChildCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 입력 쿼리
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'       2.pSpcCd    : 검체코드
'-----------------------------------------------------------------------------'
Public Function InsertRepeatCd(ByVal pRepeatCd As String, ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS002 & " (cdindex, cdval1, field1)" & _
          " VALUES ('" & CREPEATCD & "', '" & pRepeatCd & "', " & _
          "         '" & pSpcCd & "')"
    
    InsertRepeatCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 수정 쿼리
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'       2.pSpcCd    : 검체코드
'-----------------------------------------------------------------------------'
Public Function UpdateRepeatCd(ByVal pRepeatCd As String, ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS002 & " SET field1='" & pSpcCd & "'" & _
          " WHERE cdindex='" & CREPEATCD & "' AND cdval1='" & pRepeatCd & "'"
    
    UpdateRepeatCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 삭제 쿼리
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'-----------------------------------------------------------------------------'
Public Function DeleteRepeatCd(ByVal pRepeatCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS002 & _
          " WHERE cdindex='" & CREPEATCD & "' AND cdval1='" & pRepeatCd & "'"
    
    DeleteRepeatCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectRepeatCd() As String
    Dim SQL As String
    
    SQL = " SELECT a.cdval1 AS TESTCD, b.abbrnm10 AS TESTNM, a.field1 AS SPCCD, c.field3 AS SPCNM" & _
          " FROM " & TIIS002 & " a, " & TIIS101 & " b, " & TIIS002 & " c" & _
          " WHERE a.cdindex='" & CREPEATCD & "' AND c.cdindex='" & CSPCCD & "'" & _
          "     AND a.cdval1=b.testcd AND a.field1=c.cdval1" & _
          " ORDER BY a.cdval1"
    
    SelectRepeatCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도 처방코드의 이름 조회쿼리 (그룹코드, 상세모코드에서만 조회)
'-----------------------------------------------------------------------------'
Public Function SelectRepeatNm(ByVal pRepeatCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE testcd='" & pRepeatCd & "' AND (panelfg='G' OR panelfg='D')" & _
          "     AND (expdt is null OR expdt='')"
    
    SelectRepeatNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 코드에 등록된 검체코드중 해당 검체코드의 검체명 조회쿼리
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'       2.pSpcCd    : 검체코드
'-----------------------------------------------------------------------------'
Public Function SelectSpcNmX(ByVal pRepeatCd As String, ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT b.field3 AS SPCNM FROM " & TIIS102 & " a, " & TIIS002 & " b" & _
          " WHERE a.testcd='" & pRepeatCd & "' AND a.spccd='" & pSpcCd & "'" & _
          "     AND b.cdindex='" & CSPCCD & "' AND a.spccd=b.cdval1"
    
    SelectSpcNmX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 그룹코드, 상세모코드 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectRepeatCdX() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm10 AS TESTNM FROM " & TIIS101 & _
          " WHERE (panelfg='G' OR panelfg='D')" & _
          "     AND (expdt is null OR expdt='')" & _
          " ORDER BY testcd"
    
    SelectRepeatCdX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 코드에 등록된 모든 검체조회 쿼리
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'-----------------------------------------------------------------------------'
Public Function SelectSpcCdX(ByVal pRepeatCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.spccd AS SPCCD, b.field3 AS SPCNM" & _
          " FROM " & TIIS102 & " a, " & TIIS002 & " b" & _
          " WHERE a.testcd='" & pRepeatCd & "' AND b.cdindex='" & CSPCCD & "'" & _
          "     AND a.spccd=b.cdval1" & _
          " ORDER BY a.spccd"
    
    SelectSpcCdX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 건물정보 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectLocations() As String
    Dim SQL As String
    
    SQL = " SELECT cdval1 AS LOCCD, field1 AS LOCNM" & _
          " FROM " & TIIS002 & _
          " WHERE cdindex='" & CLOCATIONCD & "' ORDER BY cdval1"
    
    SelectLocations = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 제조회사정보 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectVands() As String
    Dim SQL As String
    
    SQL = " SELECT cdval1 AS VANDCD, text1 AS VANDNM" & _
          " FROM " & TIIS002 & _
          " WHERE cdindex='" & CVANDCD & "' ORDER BY cdval1"
    
    SelectVands = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 리스트 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectEqps() As String
    Dim SQL As String
    
    SQL = " SELECT eqpcd AS EQPCD, eqpnm AS EQPNM, eqpdiv AS EQPDIV, modelnm AS MODELNM," & _
          "     purchdt AS PURCHDT, location AS LOCATION, inusefg AS INUSEFG, vandcd AS VANDCD," & _
          "     templow AS TEMPLOW, temphigh AS TEMPHIGH, tempscale AS TEMPSCALE, remark AS REMARK" & _
          " FROM " & TIIS401 & " ORDER BY eqpcd"
    
    SelectEqps = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 입력 쿼리
'   인수 :
'       1.pEqpCd       : 장비코드
'       2.pEqpNm       : 장비명
'       3.pEqpDiv      : 장비구분(E:일반장비, C:온도관리장비)
'       4.pModelNm     : 모델명, Serial-No
'       5.pPurchDt     : 구매일자
'       6.pLocationCd  : 건물코드
'       7.pInUseFg     : 사용유무(1:사용, 0:사용안함)
'       8.pVandCd      : 건물코드
'       9.pTempLow     : 적정온도(상)
'       10.pTempHigh   : 적정온도(하)
'       11.pTempScale  : 온도구분(C:섭씨, F:화씨)
'       12pRemark      : 비고
'-----------------------------------------------------------------------------'
Public Function InsertEqp(ByVal pEqpCd As String, ByVal pEqpNm As String, ByVal pEqpDiv As String, _
                          ByVal pModelNm As String, ByVal pPurchDt As String, _
                          ByVal pLocationCd As String, ByVal pInUseFg As String, _
                          ByVal pVandCd As String, ByVal pTempLow As Single, _
                          ByVal pTempHigh As Single, ByVal pTempScale As String, _
                          ByVal pRemark As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS401 & " (eqpcd, eqpnm, eqpdiv, modelnm, purchdt, location," & _
          "     inusefg, vandcd, templow, temphigh, tempscale, remark)" & _
          " VALUES ('" & pEqpCd & "', '" & pEqpNm & "', '" & pEqpDiv & "'," & _
          "     '" & pModelNm & "', '" & pPurchDt & "', '" & pLocationCd & "'," & _
          "     '" & pInUseFg & "', '" & pVandCd & "', " & CStr(pTempLow) & ", " & _
          "      " & CStr(pTempHigh) & ", '" & pTempScale & "', '" & pRemark & "')"
    
    InsertEqp = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 수정 쿼리
'   인수 :
'       1.pEqpCd       : 장비코드
'       2.pEqpNm       : 장비명
'       3.pEqpDiv      : 장비구분(E:일반장비, C:온도관리장비)
'       4.pModelNm     : 모델명, Serial-No
'       5.pPurchDt     : 구매일자
'       6.pLocationCd  : 건물코드
'       7.pInUseFg     : 사용유무(1:사용, 0:사용안함)
'       8.pVandCd      : 건물코드
'       9.pTempLow     : 적정온도(상)
'       10.pTempHigh   : 적정온도(하)
'       11.pTempScale  : 온도구분(C:섭씨, F:화씨)
'       12pRemark      : 비고
'-----------------------------------------------------------------------------'
Public Function UpdateEqp(ByVal pEqpCd As String, ByVal pEqpNm As String, ByVal pEqpDiv As String, _
                          ByVal pModelNm As String, ByVal pPurchDt As String, _
                          ByVal pLocationCd As String, ByVal pInUseFg As String, _
                          ByVal pVandCd As String, ByVal pTempLow As Single, _
                          ByVal pTempHigh As Single, ByVal pTempScale As String, _
                          ByVal pRemark As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS401 & " SET eqpnm='" & pEqpNm & "', eqpdiv='" & pEqpDiv & "'," & _
          "     modelnm='" & pModelNm & "', purchdt='" & pPurchDt & "', location='" & pLocationCd & "'," & _
          "     inusefg='" & pInUseFg & "', vandcd='" & pVandCd & "', templow=" & CStr(pTempLow) & "," & _
          "     temphigh=" & CStr(pTempHigh) & ", tempscale='" & pTempScale & "', remark='" & pRemark & "'" & _
          " WHERE eqpcd='" & pEqpCd & "'"
    
    UpdateEqp = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 수정 쿼리
'   인수 :
'       1.pEqpCd      : 장비코드
'-----------------------------------------------------------------------------'
Public Function DeleteEqp(ByVal pEqpCd As String)
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS401 & " WHERE eqpcd='" & pEqpCd & "'"
    
    DeleteEqp = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 입력쿼리
'   인수 :
'       1.pEqpCd      : 장비코드
'       2.pPort       : Port
'       3.pBaudRate   : Baud Rate
'       4.pDatabit    : Data bit
'       5.pStopbit    : Stop bit
'       6.pParitybit  : Parity bit
'       7.pStoredDt   : 보관일자
'-----------------------------------------------------------------------------'
Public Function InsertEqpComm(ByVal pEqpCd As String, ByVal pPort As String, _
                              ByVal pBaudRate As String, ByVal pDatabit As String, _
                              ByVal pStopbit As String, ByVal pParitybit As String, _
                              ByVal pStoredDt As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS402 & " (eqpcd, port, baudrate, databit, stopbit, paritybit, storedate)" & _
          " VALUES ('" & pEqpCd & "', '" & pPort & "', '" & pBaudRate & "'," & _
          "     '" & pDatabit & "', '" & pStopbit & "', '" & pParitybit & "', '" & pStoredDt & "')"
    
    InsertEqpComm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 수정쿼리
'   인수 :
'       1.pEqpCd      : 장비코드
'       2.pPort       : Port
'       3.pBaudRate   : Baud Rate
'       4.pDatabit    : Data bit
'       5.pStopbit    : Stop bit
'       6.pParitybit  : Parity bit
'       7.pStoredDt   : 보관일자
'-----------------------------------------------------------------------------'
Public Function UpdateEqpComm(ByVal pEqpCd As String, ByVal pPort As String, _
                              ByVal pBaudRate As String, ByVal pDatabit As String, _
                              ByVal pStopbit As String, ByVal pParitybit As String, _
                              ByVal pStoredDt As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS402 & " SET port='" & pPort & "', baudrate='" & pBaudRate & "'," & _
          "     databit='" & pDatabit & "', stopbit='" & pStopbit & "'," & _
          "     paritybit='" & pParitybit & "', storedate='" & pStoredDt & "'" & _
          " WHERE eqpcd='" & pEqpCd & "'"
    
    UpdateEqpComm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 삭제쿼리
'   인수 :
'       1.pEqpCd      : 장비코드
'-----------------------------------------------------------------------------'
Public Function DeleteEqpComm(ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS402 & " WHERE eqpcd='" & pEqpCd & "'"
    
    DeleteEqpComm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 통신설정 조회쿼리
'   인수 :
'       1.pEqpCd : 장비코드
'-----------------------------------------------------------------------------'
Public Function SelectEqpComm(ByVal pEqpCd As String) As String
    Dim SQL As String

    SQL = " SELECT b.eqpcd AS EQPCD, b.eqpnm AS EQPNM, a.port AS PORT, a.baudrate AS BAUDRATE," & _
          "     a.databit AS DATABIT, a.stopbit AS STOPBIT, a.paritybit AS PARITYBIT," & _
          "     a.storedate AS STOREDDT" & _
          " FROM " & TIIS402 & " a, " & TIIS401 & " b" & _
          " WHERE a.eqpcd='" & pEqpCd & "' AND b.inusefg='1' AND b.eqpdiv='E'" & _
          "     AND a.eqpcd=b.eqpcd" & _
          " ORDER BY a.eqpcd"

    SelectEqpComm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 사용중인장비의 통신설정 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectEqpComms() As String
    Dim SQL As String
    
    SQL = " SELECT b.eqpcd AS EQPCD, b.eqpnm AS EQPNM, a.port AS PORT, a.baudrate AS BAUDRATE," & _
          "     a.databit AS DATABIT, a.stopbit AS STOPBIT, a.paritybit AS PARITYBIT," & _
          "     a.storedate AS STOREDDT" & _
          " FROM " & TIIS402 & " a, " & TIIS401 & " b" & _
          " WHERE b.inusefg='1' AND b.eqpdiv='E' AND a.eqpcd=b.eqpcd" & _
          " ORDER BY a.eqpcd"
    
    SelectEqpComms = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 사용중인 장비코드, 장비명 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectUsingEqp() As String
    Dim SQL As String
    
    SQL = " SELECT eqpcd AS EQPCD, eqpnm AS EQPNM" & _
          " FROM " & TIIS401 & _
          " WHERE inusefg='1' AND eqpdiv='E' ORDER BY eqpcd"
    
    SelectUsingEqp = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비코드의 장비명 조회쿼리
'   인수 :
'       1.pEpqCd : 장비코드
'-----------------------------------------------------------------------------'
Public Function SelectEqpCd(ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT eqpnm AS EQPNM FROM " & TIIS401 & _
          " WHERE eqpcd='" & pEqpCd & "' AND inusefg='1' AND eqpdiv='E'"
    
    SelectEqpCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드에 검사명 입력쿼리
'   인수 :
'       - pEqpCd    : 장비코드
'       - pIntNm    : 검사명
'       - pIntBase  : 장비기준 검사명
'       - pFrVal    : BorderLine From Value
'       - pToVal    : BorderLine To Value
'-----------------------------------------------------------------------------'
Public Function InsertIntNm(ByVal pEqpCd As String, ByVal pIntNm As String, ByVal pIntBase As String, _
                    ByVal pFrVal As String, ByVal pToVal As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS403 & " (eqpcd, intnm, intbase, frval, toval)" & _
          " VALUES ('" & pEqpCd & "', '" & pIntNm & "', '" & pIntBase & "', '" & pFrVal & "', '" & pToVal & "')"
          
    InsertIntNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드에 검사명 입력쿼리(기존병원용)
'   인수 :
'       - pEqpCd    : 장비코드
'       - pEqpSeq   : 장비순서
'       - pIntBase  : 장비기준 검사명
'       - pIntNm    : 검사명
'       - pPrtOrd   : 출력순서
'       - pFrVal    : BorderLine From Value
'       - pToVal    : BorderLine To Value
'-----------------------------------------------------------------------------'
Public Function InsertIntNmX(ByVal pEqpCd As String, ByVal pEqpSeq As Long, _
                    ByVal pIntBase As String, ByVal pIntNm As String, ByVal pPrtOrd As Long, _
                    ByVal pFrVal As String, ByVal pToVal As String) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "InsertIntNmX", _
        "이 함수대신 InsertIntNm 함수를 사용하세요.", Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드에 검사명 수정쿼리
'   인수 :
'       - pEqpCd    : 장비코드
'       - pIntNm    : 검사명
'       - pIntBase  : 장비기준 검사명
'       - pFrVal    : BorderLine From Value
'       - pToVal    : BorderLine To Value
'-----------------------------------------------------------------------------'
Public Function UpdateIntNm(ByVal pEqpCd As String, ByVal pIntNm As String, ByVal pIntBase As String, _
                    ByVal pFrVal As String, ByVal pToVal As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS403 & _
          "     SET intbase='" & pIntBase & "', frval='" & pFrVal & "', toval='" & pToVal & "'" & _
          " WHERE eqpcd='" & pEqpCd & "' AND intnm='" & pIntNm & "'"
    
    UpdateIntNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드에 검사명 수정쿼리(기존병원용)
'   인수 :
'       - pEqpCd    : 장비코드
'       - pEqpSeq   : 장비기준 검사명 순서
'       - pIntBase  : 장비기준 검사명
'       - pIntNm    : 검사명
'       - pPrtOrd   : 출력순서
'       - pFrVal    : BorderLine From Value
'       - pToVal    : BorderLine To Value
'-----------------------------------------------------------------------------'
Public Function UpdateIntNmX(ByVal pEqpCd As String, ByVal pEqpSeq As Long, _
                    ByVal pIntBase As String, ByVal pIntNm As String, ByVal pPrtOrd As Long, _
                    ByVal pFrVal As String, ByVal pToVal As String) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "UpdateIntNmX", _
        "이 함수대신 UpdateIntNm 함수를 사용하세요.", Now

End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사명 삭제쿼리
'   인수 :
'       - pEqpCd : 장비코드
'       - pIntNm : 검사명
'-----------------------------------------------------------------------------'
Public Function DeleteIntNm(ByVal pEqpCd As String, ByVal pIntNm As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS403 & _
          " WHERE eqpcd='" & pEqpCd & "' AND intnm='" & pIntNm & "'"
          
    DeleteIntNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사명 삭제쿼리(기존병원용)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpSeq : 장비기준 검사명 순서
'-----------------------------------------------------------------------------'
Public Function DeleteIntNmX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "DeleteIntNmX", _
        "이 함수대신 DeleteIntNm 함수를 사용하세요.", Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비에 등록된 모든 검사명 조회쿼리
'   인수 :
'       - pEqpCd : 장비코드
'-----------------------------------------------------------------------------'
Public Function SelectIntNms(ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT eqpcd AS EQPCD, intnm AS INTNM, intbase AS INTBASE, frval AS FRVAL, toval AS TOVAL" & _
          " FROM " & TIIS403 & " WHERE eqpcd='" & pEqpCd & "'"
          
    SelectIntNms = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 검사코드 입력쿼리
'   인수 :
'       - pEqpCd  : 장비코드
'       - pIntNm  : 검사명
'       - pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function InsertIntTestCd(ByVal pEqpCd As String, ByVal pIntNm As String, _
                    ByVal pTestCd As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS404 & " (eqpcd, intnm, testcd)" & _
          " VALUES ('" & pEqpCd & "', '" & pIntNm & "', '" & pTestCd & "')"
          
    InsertIntTestCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 검사코드 입력쿼리(기존병원)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpSeq : 장비기준 검사명 순서
'       - pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function InsertIntTestCdX(ByVal pEqpCd As String, ByVal pEqpSeq As Long, _
                    ByVal pTestCd As String) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "InsertIntTestCdX", _
        "이 함수대신 InsertIntTestCd 함수를 사용하세요.", Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 등록된 해당 검사코드 삭제쿼리
'   인수 :
'       - pEqpCd  : 장비코드
'       - pIntNm  : 검사명
'       - pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function DeleteIntTestCd(ByVal pEqpCd As String, ByVal pIntNm As String, _
                    ByVal pTestCd As String) As String
    Dim SQL As String

    SQL = " DELETE FROM " & TIIS404 & _
          " WHERE eqpcd='" & pEqpCd & "' AND intnm='" & pIntNm & "' AND testcd='" & pTestCd & "'"

    DeleteIntTestCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 등록된 해당 검사코드 삭제쿼리(기존병원용)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpSeq : 장비기준 검사명 순서
'       - pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function DeleteIntTestCdX(ByVal pEqpCd As String, ByVal pEqpSeq As Long, _
                    ByVal pTestCd As String) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "DeleteIntTestCdX", _
        "이 함수대신 DeleteIntTestCd 함수를 사용하세요.", Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사명에 등록된 모든 검사코드 삭제쿼리
'   인수 :
'       - pEqpCd : 장비코드
'       - pIntNm : 장비기준 순서
'-----------------------------------------------------------------------------'
Public Function DeleteIntTestCdAll(ByVal pEqpCd As String, ByVal pIntNm As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS404 & _
          " WHERE eqpcd='" & pEqpCd & "' AND intnm='" & pIntNm & "'"
          
    DeleteIntTestCdAll = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사명에 등록된 모든 검사코드 삭제쿼리(기존병원용)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpSeq : 장비기준 검사명 순서
'-----------------------------------------------------------------------------'
Public Function DeleteIntTestCdAllX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "DeleteIntTestCdAllX", _
        "이 함수대신 DeleteIntTestCdAll 함수를 사용하세요.", Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사명에 등록되 모든 검사코드를 조회쿼리
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pIntNm  : 검사명
'-----------------------------------------------------------------------------'
Public Function SelectIntTestCds(ByVal pEqpCd As String, ByVal pIntNm As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM" & _
          " FROM " & TIIS404 & " a, " & TIIS101 & " b" & _
          " WHERE a.eqpcd='" & pEqpCd & "' AND a.intnm='" & pIntNm & "' AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          " ORDER BY a.testcd"
          
    SelectIntTestCds = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사명에 등록되 모든 검사코드를 조회쿼리(기존병원용)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pIntNm  : 장비기준 검사명 순서
'-----------------------------------------------------------------------------'
Public Function SelectIntTestCdsX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As String
    '## 1.1.9: 이상대(2005-05-18)
    '   - Library.Dll의 소스를 수정할수 없어 골격만 유지!
    Error.SetLog App.EXEName, "clsIISMasterSql", "SelectIntTestCdsX", _
        "이 함수대신 SelectIntTestCds 함수를 사용하세요.", Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비에 등록된 LIS기준 검사코드에 대한 검사명 조회쿼리
'   인수 :
'       - pEqpCd : 장비코드
'-----------------------------------------------------------------------------'
Public Function SelectTestCdIntNms(ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.intnm AS INTNM, b.intbase AS INTBASE," & _
          "     b.frval AS FRVAL, b.toval AS TOVAL" & _
          " FROM " & TIIS404 & " a, " & TIIS403 & " b" & _
          " WHERE a.eqpcd='" & pEqpCd & "' AND a.eqpcd=b.eqpcd AND a.intnm=b.intnm" & _
          " ORDER BY a.intnm"
          
    SelectTestCdIntNms = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비기준 검사명에 등록할수 검사코드를 조회
'          (TIIS101에서 상세자코드, 단일항목 조회쿼리)
'-----------------------------------------------------------------------------'
Public Function SelectForIntTestCd() As String
    Dim SQL As String
    
    SQL = " SELECT testcd AS TESTCD, abbrnm10 AS TESTNM" & _
          " FROM " & TIIS101 & _
          " WHERE (panelfg IS NULL OR panelfg='') AND (expdt IS NULL OR expdt='')" & _
          " ORDER BY testcd"
    
    SelectForIntTestCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사코드의 이름 조회쿼리(단일, 상세자코드 중에서 조회)
'   인수 :
'       1.pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Function SelectForIntTestNm(ByVal pTestCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.abbrnm10 AS TESTNM" & _
          " FROM " & TIIS101 & " a" & _
          " WHERE a.testcd='" & pTestCd & "' AND (a.panelfg IS NULL OR a.panelfg='')" & _
          "     AND (a.expdt IS NULL OR a.expdt='')" & _
          "     AND a.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)"
    
    SelectForIntTestNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿코드 입력 쿼리
'   인수 :
'       1.pCdIndex : 템플릿 인덱스
'       2.pCode    : 템플릿 코드
'       3.pContent : 템플릿 내용
'-----------------------------------------------------------------------------'
Public Function InsertTemplate(ByVal pCdIndex As String, ByVal pCode As String, _
                                ByVal pContent As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS003 & " (cdindex, cdval1, text1)" & _
          " VALUES ('" & pCdIndex & "', '" & pCode & "', '" & pContent & "')"
          
    InsertTemplate = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿코드 수정 쿼리
'   인수 :
'       1.pCdIndex : 템플릿 인덱스
'       2.pCode    : 템플릿 코드
'       3.pContent : 템플릿 내용
'-----------------------------------------------------------------------------'
Public Function UpdateTemplate(ByVal pCdIndex As String, ByVal pCode As String, _
                                ByVal pContent As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS003 & " SET text1='" & pContent & "'" & _
          " WHERE cdindex='" & pCdIndex & "' AND cdval1='" & pCode & "'"
          
    UpdateTemplate = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿코드 삭제 쿼리
'   인수 :
'       1.pCdIndex : 템플릿 인덱스
'       2.pCode    : 템플릿 코드
'-----------------------------------------------------------------------------'
Public Function DeleteTemplate(ByVal pCdIndex As String, ByVal pCode As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS003 & _
          " WHERE cdindex='" & pCdIndex & "' AND cdval1='" & pCode & "'"
          
    DeleteTemplate = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿코드 조회 쿼리
'   인수 :
'       1.pCdIndex : 템플릿 인덱스
'-----------------------------------------------------------------------------'
Public Function SelectTemplate(ByVal pCdIndex As String) As String
    Dim SQL As String
    
    SQL = " SELECT cdval1 AS CODE, text1 AS CONTENT" & _
          " FROM " & TIIS003 & _
          " WHERE cdindex='" & pCdIndex & "' ORDER BY cdval1"
          
    SelectTemplate = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 공통코드 입력쿼리
'   인수 :
'       1.pCdIndex  : 공통코드 인덱스
'       2.pName     : 컨트롤 Name
'       3.pCaption  : 컨트롤 Caption
'       4.pLength   : 컨트롤 Length
'-----------------------------------------------------------------------------'
Public Function InsertComCd(ByVal pCdIndex As String, ByVal pName As String, _
                            ByVal pCaption As String, ByVal pLength As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TCOM001 & " (cdindex, name, caption, length)" & _
          " VALUES('" & pCdIndex & "', '" & pName & "', '" & pCaption & "', '" & pLength & "')"
          
    InsertComCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 공통코드 수정쿼리
'   인수 :
'       1.pCdIndex  : 공통코드 인덱스
'       2.pName     : 컨트롤 Name
'       3.pCaption  : 컨트롤 Caption
'       4.pLength   : 컨트롤 Length
'-----------------------------------------------------------------------------'
Public Function UpdateComCd(ByVal pCdIndex As String, ByVal pName As String, _
                            ByVal pCaption As String, ByVal pLength As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TCOM001 & " SET caption='" & pCaption & "', length='" & pLength & "'" & _
          " WHERE cdindex='" & pCdIndex & "' AND name='" & pName & "'"
          
    UpdateComCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 공통코드 삭제쿼리
'   인수 :
'       1.pCdIndex  : 공통코드 인덱스
'       2.pName     : 컨트롤 Name
'-----------------------------------------------------------------------------'
Public Function DeleteComCd(ByVal pCdIndex As String, ByVal pName As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TCOM001 & " WHERE cdindex='" & pCdIndex & "' AND name='" & pName & "'"

    DeleteComCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 공통코드에 포함된 모든정보 삭제쿼리
'   인수 :
'       1.pCdIndex  : 공통코드 인덱스
'-----------------------------------------------------------------------------'
Public Function DeleteComCdAll(ByVal pCdIndex As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TCOM001 & " WHERE cdindex='" & pCdIndex & "'"
    
    DeleteComCdAll = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 모든 공통코드 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectComCdAll() As String
    Dim SQL As String
    
    SQL = " SELECT cdindex AS CDINDEX, caption AS CAPTION" & _
          " FROM " & TCOM001 & _
          " WHERE name='Form' ORDER BY cdindex"
          
    SelectComCdAll = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 모든 공통코드 조회쿼리
'   인수 :
'       1.pCdIndex : 공통코드 인덱스
'-----------------------------------------------------------------------------'
Public Function SelectComCd(ByVal pCdIndex As String) As String
    Dim SQL As String
    
    SQL = " SELECT cdindex AS CDINDEX, name AS NAME, caption AS CAPTION, length AS LENGTH" & _
          " FROM " & TCOM001 & _
          " WHERE cdindex='" & pCdIndex & "' ORDER BY name"
          
    SelectComCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사코드, 결과의 Alpha결과 판정치 조회 - 1.1.9: 이상대(2005-05-18)
'   인수 :
'       - pTestCd : 검사코드
'       - pResult : 결과
'-----------------------------------------------------------------------------'
Public Function SelectAlphaRef(ByVal pTestCd As String, ByVal pResult As String)
    Dim SQL As String
    
    '## 1.2.1: 이상대(2005-05-30)
    '   - 잘못된 쿼리수정
    SQL = " SELECT field2 AS DECISION FROM " & TIIS001 & _
          " WHERE cdindex='" & CRESULTCD & "'" & _
          "     AND cdval1='" & pTestCd & "' AND cdval2='" & Trim$(pResult) & "'"
          
    SelectAlphaRef = SQL
End Function


