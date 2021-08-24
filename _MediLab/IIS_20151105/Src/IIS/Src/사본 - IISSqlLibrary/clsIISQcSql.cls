VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISQcSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISQcSql.cls
'   작성자  : 오세원
'   내  용  : QC에서 사용하는 쿼리
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에 사용하는 QC결과내역 조회쿼리(유효숫자, D/P수치도 조회)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectQCResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    '## 디버깅시 사용할 쿼리
'    SQL = " SELECT a.testcd AS TESTCD, c.abbrnm10 AS TESTNM10, c.testnm AS TESTNM, a.rstunit AS UNIT," & _
'          "     a.rstdiv AS RSTDIV, a.ctrlcd AS CTRLCD, a.levelcd AS LEVELCD, a.lotno AS LOTNO," & _
'          "     b.meanval AS MEANVAL, b.sdval AS SDVAL, b.avalval AS AVALVAL, b.refcd AS REFCD," & _
'          "     b.cvval AS CVVAL, b.minval AS MINVAL, b.maxval AS MAXVAL, b.wmset AS WMSET" & _
'          " FROM " & TIIS307 & " a, " & TIIS304 & " b, " & TIIS101 & " c" & _
'          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
'          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt='' OR a.vfydt IS NULL)" & _
'          "     AND a.ctrlcd=b.ctrlcd AND a.levelcd=b.levelcd AND a.lotno=b.lotno AND a.testcd=b.testcd" & _
'          "     AND a.testcd=c.testcd" & _
'          "     AND c.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
'          " ORDER BY c.rptseq"
          
    '## 실제 사용할 쿼리
    SQL = " SELECT a.testcd AS TESTCD, c.abbrnm10 AS TESTNM10, c.testnm AS TESTNM, a.rstunit AS UNIT," & _
          "     a.rstdiv AS RSTDIV, a.ctrlcd AS CTRLCD, a.levelcd AS LEVELCD, a.lotno AS LOTNO," & _
          "     b.meanval AS MEANVAL, b.sdval AS SDVAL, b.avalval AS AVALVAL, b.refcd AS REFCD," & _
          "     b.cvval AS CVVAL, b.minval AS MINVAL, b.maxval AS MAXVAL, b.wmset AS WMSET" & _
          " FROM " & TIIS307 & " a, " & TIIS304 & " b, " & TIIS101 & " c" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt='' OR a.vfydt IS NULL)" & _
          "     AND (a.rstcd='' OR a.rstcd IS NULL)" & _
          "     AND a.ctrlcd=b.ctrlcd AND a.levelcd=b.levelcd AND a.lotno=b.lotno AND a.testcd=b.testcd" & _
          "     AND a.testcd=c.testcd" & _
          "     AND c.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          " ORDER BY c.rptseq"
    
    SelectQCResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당컨트롤, 검사코드의 과거결과 9개 조회쿼리
'   인수 :
'       - pCtrlCd  : Control Code
'       - pLevelCd : Level Code
'       - pLotNo   : Lot No
'       - pTestCd  : 검사코드
'-----------------------------------------------------------------------------'
Public Function SelectHistoryResult(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                    ByVal pLotNo As String, ByVal pTestCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT rstcd AS RSTCD, rsttype AS RSTTYPE, radiv AS RADIV" & _
          " FROM (SELECT * FROM " & TIIS307 & _
          "       WHERE ctrlcd='" & pCtrlCd & "' AND levelcd='" & pLevelCd & "' AND lotno='" & pLotNo & "'" & _
          "             AND testcd='" & pTestCd & "' AND rsttype='N' AND (vfydt<>'' OR vfydt IS NOT NULL)" & _
          "       ORDER BY vfydt DESC, vfytm DESC)" & _
          " WHERE rownum<=9"
    
    SelectHistoryResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 QC 결과입력 쿼리 (결과유형이 Numeric인것만)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pRstVal   : 검사결과1(Numeric)
'       - pRstCd    : 검사결과2(Alphabetic)
'       - pRstType  : 결과유형 (F:Free, N:Numeric, A:Alpha)
'       - pRADiv    : 결과판정 (A:Accept, R:Reject)
'       - pTxtFg    : Text결과 유무(Null:무, 1:유)
'       - pAutoFg   : 장비전송 여부
'       - pEqpCd    : 장비코드
'-----------------------------------------------------------------------------'
Public Function UpdateQCResultNumeric(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstVal As Single, _
                    ByVal pRstCd As String, ByVal pRstType As String, ByVal pRADiv As String, _
                    ByVal pTxtFg As String, ByVal pAutoFg As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS307 & " SET rstval=" & CStr(pRstVal) & ", rstcd='" & pRstCd & "'," & _
          "     rsttype='" & pRstType & "', radiv='" & pRADiv & "', txtfg='" & pTxtFg & "'," & _
          "     autofg='" & pAutoFg & "', eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"

    UpdateQCResultNumeric = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 QC 결과입력 쿼리 (결과유형이 Free인것만)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pRstCd    : 검사결과2(Alphabetic)
'       - pRstType  : 결과유형 (F:Free, N:Numeric, A:Alpha)
'       - pRADiv    : 결과판정 (A:Accept, R:Reject)
'       - pTxtFg    : Text결과 유무(Null:무, 1:유)
'       - pAutoFg   : 장비전송 여부
'       - pEqpCd    : 장비코드
'-----------------------------------------------------------------------------'
Public Function UpdateQCResultFree(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstCd As String, _
                    ByVal pRstType As String, ByVal pRADiv As String, ByVal pTxtFg As String, _
                    ByVal pAutoFg As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS307 & " SET rstcd='" & pRstCd & "', rsttype='" & pRstType & "'," & _
          "     radiv='" & pRADiv & "', txtfg='" & pTxtFg & "', autofg='" & pAutoFg & "'," & _
          "     eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateQCResultFree = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 QC 소견결과 조회쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Function SelectQCComment(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String) As String
    Dim SQL As String
    
    '## DOIT: QUERY CHECK
    SQL = " SELECT seq AS SEQ FROM " & TIIS308 & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    SelectQCComment = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 QC 소견결과 입력쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Function InsertQCComment(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pSeq As Long, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS308 & " (workarea, accdt, accseq, testcd, seq, text)" & _
          "     VALUES ('" & pWorkarea & "', '" & pAccDt & "', " & CStr(pAccSeq) & "," & _
          "         '" & pTestCd & "', " & CStr(pSeq) & ", '" & pText & "')"
    
    InsertQCComment = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에서 QC 소견결과 수정쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Function UpdateQCComment(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pSeq As Long, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS308 & " SET text='" & pText & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "' AND seq=" & CStr(pSeq)
    
    UpdateQCComment = SQL
End Function
