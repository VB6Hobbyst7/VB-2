VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISLibrarySql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISLibrarySql.cls (단국대병원)
'   작성자  : 이상대
'   내  용  : IISLibrary에서 사용하는 쿼리
'   작성일  : 2005-06-24
'   버  전  :
'       1. 1.0.1: 이상대(2005-10-11)
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에 사용하는 접수내역 조회쿼리
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'-----------------------------------------------------------------------------'
Public Function SelectAccInfo(ByVal pSpcYy As String, ByVal pSpcNo As String) As String
    Dim SQL As String
    
    '## 환자ID, 진료과, 병동, 처방의에 Outer조인을 한이유는 QC검체때문에
    '## QC검체는 환자, 진료과, 병동, 처방의가 없음
    SQL = " SELECT a.ptid AS PTID, b." & FPTNM & " AS NAME, " & FJUMIN & " AS SSN, a.ageday AS AGEDAY," & _
          "     a.sex AS SEX, a.workarea AS WORKAREA, a.accdt AS ACCDT, a.accseq AS ACCSEQ," & _
          "     a.stscd AS STSCD, a.spccd AS SPCCD, f.field3 AS SPCNM, a.statfg AS STATFG," & _
          "     a.reqtotcnt AS REQTOTCNT, a.reqinputcnt AS INPUTCNT," & _
          "     a.coldt AS COLDT, a.coltm AS COLTM, a.rcvdt AS RCVDT, a.rcvtm AS RCVTM," & _
          "     a.deptcd AS DEPTCD, c." & FDEPTNM & " AS DEPTNM, d." & FDOCTNM & " AS DOCTNM," & _
          "     a.wardid AS WARDID, e." & FWARDNM & " AS WARDNM," & _
          "     a.hosilid AS HOSILID, a.testdiv AS TESTDIV, a.qcfg AS QCFG, a.buildcd AS BUILDCD" & _
          " FROM " & TIIS203 & " a, " & THIS001 & " b, " & THIS002 & " c, " & THIS004 & " d, " & _
                THIS003 & " e, " & TIIS002 & " f" & _
          " WHERE a.spcyy='" & pSpcYy & "' AND a.spcno=" & pSpcNo & _
          "     AND a.ptid=b." & FPTID & "(+) AND a.deptcd=c." & FDEPTCD & "(+) AND a.orddoct=d." & FDOCTCD & "(+)" & _
          "     AND a.wardid=e." & FWARDCD & "(+) AND a.spccd=f.cdval1(+) AND f.cdindex(+)='" & CSPCCD & "'"

    SelectAccInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에 사용하는 결과내역 조회쿼리(유효숫자, D/P수치도 조회)
'          결과가 없는 항목만 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM, c.field3 AS SPCNM," & _
          "     a.rstunit AS UNIT, a.lastrst AS LASTRST, b.rsttype AS RSTTYPE," & _
          "     a.rstdiv AS RSTDIV, a.detailfg AS DETAILFG, d.avalval AS AVALVAL," & _
          "     d.panicfg AS PANICFG, d.panicfrval AS PANICFRVAL, d.panictoval AS PANICTOVAL," & _
          "     d.deltafg AS DELTAFG, d.deltaval AS DELTAFRVAL, d.deltaval2 AS DELTATOVAL" & _
          " FROM " & TIIS204 & " a, " & TIIS101 & " b, " & TIIS002 & " c, " & TIIS102 & " d" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt IS NULL OR a.vfydt='')" & _
          "     AND (a.rstcd IS NULL OR a.rstcd='')" & _
          "     AND c.cdindex='" & CSPCCD & "' AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          "     AND a.spccd=c.cdval1" & _
          "     AND a.testcd=d.testcd AND a.spccd=d.spccd" & _
          "     AND d.applydt=(SELECT MAX(applydt) FROM " & TIIS102 & _
          "                    WHERE testcd=a.testcd AND spccd=a.spccd)" & _
          " ORDER BY b.rptseq"
          
    SelectResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에 사용하는 특수검사 결과내역 조회쿼리 - 1.0.1: 이상대(2005-10-11)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectSpecialResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM" & _
          " FROM " & TIIS601 & " a, " & TIIS101 & " b" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt IS NULL OR a.vfydt='')" & _
          "     AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          " ORDER BY b.rptseq"
    
    SelectSpecialResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에 사용하는 결과내역 조회쿼리(유효숫자, D/P수치도 조회)
'          보고되지 않음 항목만 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectAutoRerunResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM, c.field3 AS SPCNM," & _
          "     a.rstunit AS UNIT, a.lastrst AS LASTRST, b.rsttype AS RSTTYPE," & _
          "     a.rstdiv AS RSTDIV, a.detailfg AS DETAILFG, d.avalval AS AVALVAL," & _
          "     d.panicfg AS PANICFG, d.panicfrval AS PANICFRVAL, d.panictoval AS PANICTOVAL," & _
          "     d.deltafg AS DELTAFG, d.deltaval AS DELTAFRVAL, d.deltaval2 AS DELTATOVAL" & _
          " FROM " & TIIS204 & " a, " & TIIS101 & " b, " & TIIS002 & " c, " & TIIS102 & " d" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (vfydt IS NULL OR vfydt='')" & _
          "     AND c.cdindex='" & CSPCCD & "' AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          "     AND a.spccd=c.cdval1" & _
          "     AND a.testcd=d.testcd AND a.spccd=d.spccd" & _
          "     AND d.applydt=(SELECT MAX(applydt) FROM " & TIIS102 & _
          "                    WHERE testcd=a.testcd AND spccd=a.spccd)" & _
          " ORDER BY b.rptseq"
          
    SelectAutoRerunResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 전송내역 입력쿼리
'   인수 :
'       - pEqpCd    : 장비코드
'       - pTransDt  : 장비에서 결과 전송일자
'       - pTransTm  : 장비에서 결과 전소시간
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'       - pTransNo  : 장비에서 검체위치
'-----------------------------------------------------------------------------'
Public Function InsertEqpTrans(ByVal pEqpCd As String, ByVal pTransDt As String, _
                    ByVal pTransTm As String, ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pTransNo As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS405 & " (eqpcd, transdt, transtm, spcyy, spcno, transno)" & _
          " VALUES ('" & pEqpCd & "', '" & pTransDt & "', '" & pTransTm & "'," & _
          "     '" & pSpcYy & "', " & pSpcNo & ", '" & pTransNo & "')"
    
    InsertEqpTrans = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 전송내역 삭제쿼리(장비코드, 검체연도, 검체번호가 일치하는것만 삭제)
'   인수 :
'       - pEqpCd    : 장비코드
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'-----------------------------------------------------------------------------'
Public Function DeleteEqpTrans(ByVal pEqpCd As String, ByVal pSpcYy As String, _
                    ByVal pSpcNo As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS405 & _
          " WHERE eqpcd='" & pEqpCd & "' AND spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
    
    DeleteEqpTrans = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 사용하는 상세모코드 결과입력 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pAutoFg   : 장비전송 여부
'       - pEqpCd    : 장비코드
'-----------------------------------------------------------------------------'
Public Function UpdateIISResultDetail(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, _
                    ByVal pAutoFg As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET autofg='" & pAutoFg & "', eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
          
    UpdateIISResultDetail = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 사용하는 상세모코드 결과확인 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pAutoFg   : 장비전송 여부
'       - pEqpCd    : 장비코드
'       - pVfyDt    : 결과확인일자
'       - pVfyTm    : 결과확인시간
'       - pVfyId    : 결과확인자
'-----------------------------------------------------------------------------'
Public Function UpdateIISVfyDetail(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pAutoFg As String, _
                    ByVal pEqpCd As String, ByVal pVfyDt As String, ByVal pVfyTm As String, _
                    ByVal pVfyId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET autofg='" & pAutoFg & "', eqpcd='" & pEqpCd & "'," & _
          "     vfydt='" & pVfyDt & "', vfytm='" & pVfyTm & "', vfyid='" & pVfyId & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISVfyDetail = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 사용하는 단일항목, 상세자코드 결과입력 쿼리
'          결과유형이 Numeric인것만
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pRstVal   : 검사결과1(Numeric)
'       - pRstCd    : 검사결과2(Alphabetic)
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pAutoFg   : 장비전송 여부
'       - pRstType  : 결과유형
'       - pEqpCd    : 장비코드
'-----------------------------------------------------------------------------'
Public Function UpdateIISResultNumeric(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstVal As Single, _
                    ByVal pRstCd As String, ByVal pHLDiv As String, ByVal pDPDiv As String, _
                    ByVal pAutoFg As String, ByVal pRstType As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET rstval=" & CStr(pRstVal) & ", rstcd='" & pRstCd & "'," & _
          "     hldiv='" & pHLDiv & "', dpdiv='" & pDPDiv & "', autofg='" & pAutoFg & "'," & _
          "     rsttype='" & pRstType & "', eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISResultNumeric = SQL
End Function


'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 사용하는 단일항목, 상세자코드 결과입력 쿼리
'          결과유형이 Free인것만
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pRstCd    : 검사결과2(Alphabetic)
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pAutoFg   : 장비전송 여부
'       - pRstType  : 결과유형
'       - pEqpCd    : 장비코드
'-----------------------------------------------------------------------------'
Public Function UpdateIISResultFree(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstCd As String, _
                    ByVal pHLDiv As String, ByVal pDPDiv As String, ByVal pAutoFg As String, _
                    ByVal pRstType As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET rstcd='" & pRstCd & "', hldiv='" & pHLDiv & "'," & _
          "     dpdiv='" & pDPDiv & "', autofg='" & pAutoFg & "'," & _
          "     rsttype='" & pRstType & "', eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISResultFree = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 사용하는 단일항목, 상세자코드 결과확인 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pRstVal   : 검사결과1(Numeric)
'       - pRstCd    : 검사결과2(Alphabetic)
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pAutoFg   : 장비전송 여부
'       - pRstType  : 결과유형
'       - pEqpCd    : 장비코드
'       - pVfyDt    : 결과확인일자
'       - pVfyTm    : 결과확인시간
'       - pVfyId    : 결과확인자
'-----------------------------------------------------------------------------'
Public Function UpdateIISVfy(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstVal As Single, _
                    ByVal pRstCd As String, ByVal pHLDiv As String, ByVal pDPDiv As String, _
                    ByVal pAutoFg As String, ByVal pRstType As String, ByVal pEqpCd As String, _
                    ByVal pVfyDt As String, ByVal pVfyTm As String, ByVal pVfyId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET rstval=" & CStr(pRstVal) & ", rstcd='" & pRstCd & "'," & _
          "     hldiv='" & pHLDiv & "', dpdiv='" & pDPDiv & "', autofg='" & pAutoFg & "'," & _
          "     rsttype='" & pRstType & "', eqpcd='" & pEqpCd & "', vfydt='" & pVfyDt & "'," & _
          "     vfytm='" & pVfyTm & "', vfyid='" & pVfyId & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISVfy = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수내역중 입력된 검사항목개수 수정쿼리
'   인수 :
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'       - pInputCnt : 입력된 검사항목 개수
'-----------------------------------------------------------------------------'
Public Function UpdateInputCnt(ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pInputCnt As Long) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS203 & " SET reqinputcnt=" & CStr(pInputCnt) & _
          " WHERE spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
          
    UpdateInputCnt = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수내역중 총검사항개수, 입력된 검사항목개수 조회쿼리
'   인수 :
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'-----------------------------------------------------------------------------'
Public Function SelectTotCnt(ByVal pSpcYy As String, ByVal pSpcNo As String)
    Dim SQL As String
    
    SQL = " SELECT reqtotcnt AS TOTCNT, reqinputcnt AS INPUTCNT" & _
          " FROM " & TIIS203 & " WHERE spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
          
    SelectTotCnt = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수내역을 결과확인하는 쿼리
'   인수 :
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'       - pStsCd    : 검체상태
'       - pVfyDt    : 결과확인일자
'       - pVfyTm    : 결과확인시간
'       - pVfyId    : 결과확인자
'-----------------------------------------------------------------------------'
Public Function UpdateAccInfoVfy(ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pStsCd As String, ByVal pVfyDt As String, ByVal pVfyTm As String, _
                    ByVal pVfyId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS203 & " SET stscd='" & pStsCd & "', vfydt='" & pVfyDt & "," & _
          "     vfytm='" & pVfyTm & "', vfyid='" & pVfyId & "'" & _
          " WHERE spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
          
    UpdateAccInfoVfy = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 처방내역을 결과확인하는 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pStsCd    : 검체상태
'       - pExamDt   : 결과확인일자
'       - pExamTm   : 결과확인시간
'       - pExamId   : 결과확인자
'-----------------------------------------------------------------------------'
Public Function UpdateOrderVfy(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pStsCd As String, ByVal pExamDt As String, _
                    ByVal pExamTm As String, ByVal pExamId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS202 & " SET stscd='" & pStsCd & "', examdt='" & pExamDt & "'," & _
          "     examtm='" & pExamTm & "', examdoct='" & pExamId & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "' AND accseq=" & CStr(pAccSeq)
          
    UpdateOrderVfy = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb의 접수내역에서 MAX(itemseq)+1을 조회쿼리
'-----------------------------------------------------------------------------'
Public Function SelectMaxSeq() As String
    Dim SQL As String
    
    SQL = " SELECT MAX(itemseq)+1 AS MAXSEQ FROM " & TACC203
    
    SelectMaxSeq = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb에 일반검사 접수내역을 입력하는 쿼리
'   인수 :
'       - pItemSeq : 전송순서
'       - pSpcNo   : 검체번호
'       - pPtId    : 환자ID
'       - pName    : 이름
'       - pSex     : 성별
'       - pAgeDay  : 일령
'       - pDoctNm  : 처방의명
'       - pDeptNm  : 진료과명
'       - pWardNm  : 병동명
'       - pSpcCd   : 검체코드
'       - pSpcNm   : 검체명
'       - pStatFg  : 응급여부
'       - pQcFg    : QC여부
'       - pSpcPos  : 장비에서 검체위치
'       - pEqpCd   : 장비코드
'       - pTransDt : 전송일자
'       - pTransTm : 전송시간
'-----------------------------------------------------------------------------'
Public Function InsertClientAccInfo(ByVal pItemSeq As Long, ByVal pSpcNo As String, _
                    ByVal pPtId As String, ByVal pName As String, ByVal pSex As String, _
                    ByVal pAgeDay As Long, ByVal pDoctNm As String, ByVal pDeptNm As String, _
                    ByVal pWardNm As String, ByVal pSpcCd As String, ByVal pSpcNm As String, _
                    ByVal pStatFg As String, ByVal pQcFg As String, ByVal pSpcPos As String, _
                    ByVal pEqpCd As String, ByVal pTransDt As String, ByVal pTransTm As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC203 & " (itemseq, spcno, ptid, name, sex, ageday, doctnm, deptnm," & _
          "     wardnm, spccd, spcnm, statfg, qcfg, spcpos, eqpcd, transdt, transtm)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pSpcNo & "', '" & pPtId & "', '" & pName & "'," & _
          "     '" & pSex & "', " & CStr(pAgeDay) & ", '" & pDoctNm & "', '" & pDeptNm & "'," & _
          "     '" & pWardNm & "', '" & pSpcCd & "', '" & pSpcNm & "', '" & pStatFg & "'," & _
          "     '" & pQcFg & "', '" & pSpcPos & "', '" & pEqpCd & "', '" & pTransDt & "', '" & pTransTm & "')"
    
    InsertClientAccInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb에 QC검사 접수내역을 입력하는 쿼리
'   인수 :
'       - pItemSeq : 전송순서
'       - pSpcNo   : 검체번호
'       - pPtId    : 환자ID
'       - pQcFg    : QC여부
'       - pCtrlCd  : Control Code
'       - pLevelCd : Level Code
'       - pLotNo   : Lot No
'       - pSpcPos  : 장비에서 검체위치
'       - pEqpCd   : 장비코드
'       - pTransDt : 전송일자
'       - pTransTm : 전송시간
'-----------------------------------------------------------------------------'
Public Function InsertClientQCAccInfo(ByVal pItemSeq As Long, ByVal pSpcNo As String, _
                    ByVal pPtId As String, ByVal pQcFg As String, ByVal pCtrlCd As String, _
                    ByVal pLevelCd As String, ByVal pLotNo As String, ByVal pSpcPos As String, _
                    ByVal pEqpCd As String, ByVal pTransDt As String, ByVal pTransTm As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC203 & " (itemseq, spcno, ptid, qcfg, ctrlcd, levelcd, lotno, spcpos," & _
          "     eqpcd, transdt, transtm)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pSpcNo & "', '" & pPtId & "'," & _
          "     '" & pQcFg & "', '" & pCtrlCd & "', '" & pLevelCd & "', '" & pLotNo & "'," & _
          "     '" & pSpcPos & "', '" & pEqpCd & "', '" & pTransDt & "', '" & pTransTm & "')"
    
    InsertClientQCAccInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb에 결과를 입력하는 쿼리
'   인수 :
'       - pItemSeq  : 전송순서
'       - pIntNm    : 화면표시 검사명
'       - pTestCd   : 검사코드
'       - pResult   : 결과
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'-----------------------------------------------------------------------------'
Public Function InsertClientResult(ByVal pItemSeq As Long, ByVal pIntNm As String, _
                    ByVal pTestCd As String, ByVal pResult As String, ByVal pHLDiv As String, _
                    ByVal pDPDiv As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC204 & " (itemseq, intnm, testcd, result, hldiv, dpdiv)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pIntNm & "', '" & pTestCd & "', '" & pResult & "'," & _
          "     '" & pHLDiv & "', '" & pDPDiv & "')"
          
    InsertClientResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb에 QC결과를 입력하는 쿼리
'   인수 :
'       - pItemSeq  : 전송순서
'       - pIntNm    : 화면표시 검사명
'       - pTestCd   : 검사코드
'       - pResult   : 결과
'       - pRADiv    : QC결과판정 (A/R)
'       - pText     : QC Reject 소견
'-----------------------------------------------------------------------------'
Public Function InsertClientQCResult(ByVal pItemSeq As Long, ByVal pIntNm As String, _
                    ByVal pTestCd As String, ByVal pResult As String, ByVal pRADiv As String, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC204 & " (itemseq, intnm, testcd, result, radiv, comment)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pIntNm & "', '" & pTestCd & "', '" & pResult & "'," & _
          "     '" & pRADiv & "', '" & pText & "')"
    InsertClientQCResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 보관일이 지난 ClientDb의 데이터 삭제
'   인수 :
'       - pEqpCd  : 장비코드
'       - pBaseDt : 삭제 기준일
'-----------------------------------------------------------------------------'
Public Function DeleteHisData(ByVal pEqpCd As String, ByVal pBaseDt As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TACC203 & _
          " WHERE transdt<'" & pBaseDt & "' AND eqpcd='" & pEqpCd & "'"
    
    DeleteHisData = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체명 조회
'   인수 :
'       - pSpcCd : 검체코드
'-----------------------------------------------------------------------------'
Public Function SelectSpcNm(ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT FIELD3 AS SPCNM FROM " & TIIS002 & _
          " WHERE cdindex='" & CSPCCD & "' AND cdval1='" & pSpcCd & "'"
    
    SelectSpcNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 검체연도, 검체순번을 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectBarNo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT spcyy AS SPCYY, spcno AS SPCNO FROM " & TIIS203 & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq)
          
    SelectBarNo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회
'   인수 :
'       - pFromDt : 접수일자 From
'       - pToDt   : 접수일자 To
'-----------------------------------------------------------------------------'
Public Function SelectTargetSpcs(ByVal pEqpCd As String, ByVal pFromDt As String, _
                                 ByVal pToDt As String) As String
    Dim SQL As String

    '## 단일항목, 상세자코드중 결과입력이 안된것만 조회!
    SQL = " SELECT a.spcyy AS SPCYY, a.spcno AS SPCNO, a.workarea AS WORKAREA, a.accdt AS ACCDT," & _
          "     a.accseq AS ACCSEQ, a.ptid AS PTID, c." & FPTNM & " AS NAME, " & FJUMIN & " AS SSN," & _
          "     a.sex AS SEX, a.statfg AS STATFG, a.wardid AS WARDID, a.deptcd AS DEPTCD," & _
          "     d.field3 AS SPCNM, e.testnm AS TESTNM, f.empnm AS RCVNM, a.rcvdt AS RCVDT," & _
          "     a.rcvtm AS RCVTM, g.mesg AS MESG" & _
          " FROM " & TIIS203 & " a, " & TIIS204 & " b, " & THIS001 & " c, " & TIIS002 & " d," & _
                TIIS101 & " e, " & THIS006 & " f, " & TIIS202 & " g" & _
          " WHERE a.rcvdt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' AND a.stscd IN ('2', '3')" & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          "     AND b.testcd IN (SELECT TESTCD FROM " & TIIS404 & " WHERE eqpcd='" & pEqpCd & "')" & _
          "     AND (b.rstcd IS NULL OR b.rstcd='')" & _
          "     AND (b.vfydt IS NULL OR b.vfydt='')" & _
          "     AND b.ptid=g.ptid AND b.orddt=g.orddt AND b.ordno=g.ordno AND b.ordseq=g.ordseq" & _
          "     AND a.ptid=c." & FPTID & " AND d.cdindex='" & CSPCCD & "' AND a.spccd=d.cdval1" & _
          "     AND b.testcd=e.testcd" & _
          "     AND e.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=b.testcd)" & _
          "     AND a.rcvid=f.empid(+)" & _
          " ORDER BY a.workarea, a.accdt, a.accseq"
          
    SelectTargetSpcs = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회
'   인수 :
'       - pEqpCd : 장비코드
'       - pDate  : 조회날짜
'-----------------------------------------------------------------------------'
Public Function SelectTargetSpcsX(ByVal pEqpCd As String, ByVal pDate As String) As String
    Dim SQL As String
    
    '## 1.0.1: 이상대(2005-10-11)
    '   - 특수검사 결과내역도 조회하도록 쿼리수정
    SQL = " SELECT DISTINCT a.spcyy AS SPCYY, a.spcno AS SPCNO," & _
          "     a.workarea||'-'||a.accdt||'-'||to_char(a.accseq) AS ACCNO," & _
          "     a.ptid AS PTID, c." & FPTNM & " AS NAME, a.rcvdt, a.rcvtm" & _
          " FROM " & TIIS203 & " a, " & TIIS204 & " b, " & THIS001 & " c" & _
          " WHERE a.rcvdt='" & pDate & "' AND a.workarea='" & IISCBCWA & "'" & _
          "     AND a.stscd='2' AND AND a.qcfg='0' AND (a.readfg='' OR a.readfg IS NULL)" & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          "     AND (b.vfydt IS NULL OR b.vfydt='')" & _
          "     AND b.testcd IN (SELECT testcd FROM " & TIIS404 & " WHERE eqpcd='" & pEqpCd & "')" & _
          "     AND a.ptid=c." & FPTID
          
    SQL = SQL & " UNION ALL " & _
          " SELECT DISTINCT a.spcyy AS SPCYY, a.spcno AS SPCNO," & _
          "     a.workarea||'-'||a.accdt||'-'||to_char(a.accseq) AS ACCNO," & _
          "     a.ptid AS PTID, c." & FPTNM & " AS NAME, a.rcvdt, a.rcvtm" & _
          " FROM " & TIIS203 & " a, s2lab351 b, " & THIS001 & " c" & _
          " WHERE a.rcvdt='" & pDate & "' AND a.workarea='" & IISCBCWA & "'" & _
          "     AND a.stscd='2' AND (a.readfg='' OR a.readfg IS NULL)" & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          "     AND (b.vfydt IS NULL OR b.vfydt='')" & _
          "     AND b.testcd IN (SELECT testcd FROM " & TIIS404 & " WHERE eqpcd='" & pEqpCd & "')" & _
          "     AND a.ptid=c." & FPTID & _
          " ORDER BY rcvdt, rcvtm"
          
    SelectTargetSpcsX = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사대상 리스트 조회후 조회 Flag 업데이트 쿼리
'-----------------------------------------------------------------------------'
Public Function UpdateTargetSpc(ByVal pSpcYy As String, ByVal pSpcNo As Long) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS203 & " SET readfg='1'" & _
          " WHERE spcyy='" & pSpcYy & "' AND spcno=" & CStr(pSpcNo)
          
    UpdateTargetSpc = SQL
End Function

