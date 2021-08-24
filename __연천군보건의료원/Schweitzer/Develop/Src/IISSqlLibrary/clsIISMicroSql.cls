VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMicroSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISMicroSql.cls
'   내  용  : 미생물에서 사용하는 쿼리
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스에 사용하는 미생물 결과내역 조회쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCds  : 장비별 검사항목에 등록된 검사코드
'-----------------------------------------------------------------------------'
Public Function SelectMICResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCds As String) As String
    Dim SQL As String
    
    '   - 한개의 접수번호가 여러개의 Worksheet에 포함될수 있어 Sensi에 대한 결과내역만 조회
    '     하도록 변경
    '   - 장비별 검사항목에 등록된 검사코드에 해당되는것만 조회
    SQL = " SELECT a.wscd AS WSCD, a.wsunit AS WSUNIT, a.selfg AS SELFG, a.genfg AS GENFG," & _
          "     a.scfg AS SCFG, b.testcd AS TESTCD, c.abbrnm10 AS TESTNM10, c.testnm AS TESTNM," & _
          "     b.rsttype AS RSTTYPE, b.mfyseq AS MFYSEQ, b.ptid AS PTID" & _
          " FROM " & TIIS502 & " a, " & TIIS504 & " b, " & TIIS101 & " c" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND a.scfg IN ('" & IISGENSENSI & "','" & IISMIC & "')" & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          "     AND b.stscd IN ('3','4') AND b.rsttype IN ('" & IISGENSENSI & "','" & IISMIC & "')" & _
          "     AND b.testcd IN (" & pTestCds & ")" & _
          "     AND b.testcd=c.testcd" & _
          "     AND c.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & _
          "                    WHERE testcd=b.testcd AND applydt<=b.orddt)"
    
    SelectMICResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : Vitek에서 수신한 균명(약어)를 이용 균코드를 조회쿼리
'   인수 :
'       - pMnmCd : 균명(약어)
'-----------------------------------------------------------------------------'
Public Function SelectMnmCd(ByVal pMnmNm As String) As String
    Dim SQL As String

    SQL = " SELECT cdval1 AS MNMCD FROM " & TIIS002 & _
          " WHERE cdindex='" & CMNMCD & "' AND field1='" & pMnmNm & "'"
    
    SelectMnmCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 결과내역 업데이트 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pSenFg    : 감수성결과 유무
'-----------------------------------------------------------------------------'
Public Function UpdateMicroResult(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pSenFg As String) As String
    Dim SQL As String
    
    '   - 감수성 결과코드만 SenFg를 업데이트
    SQL = " UPDATE " & TIIS504 & " SET senfg='" & pSenFg & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & _
          "     AND rsttype IN ('" & IISGENSENSI & "', '" & IISMIC & "')"
    
    UpdateMicroResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 항생제결과 입력 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'       - pMfySeq   : 수정회수
'       - pSeq      : Seq
'       - pMnmCd    : 균코드
'       - pMicFg    : MIC유무
'       - pMqtCd    : 정도코드
'       - pSCnt     : 항생제 결과수
'       - pResult   : 항생제 결과
'-----------------------------------------------------------------------------'
Public Function InsertMICResult(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pMfySeq As Long, _
                    ByVal pSeq As Long, ByVal pMnmCd As String, ByVal pMicFg As String, _
                    ByVal pMqtCd As String, ByVal pSCnt As String, ByVal pResult As String) As String
    Dim strColumns As String    'INSERT문의 컬럼정의
    Dim SQL        As String
    Dim i          As Long
    
    '## 입력할 항생제개수 만큼 컬럼정의
    For i = 1 To pSCnt
        strColumns = strColumns & ", srst" & CStr(i)
    Next i
    
    SQL = " INSERT INTO " & TIIS505 & " (workarea, accdt, accseq, testcd, mfyseq, seq, mnmcd," & _
          "     micfg, mqtcd, scnt" & strColumns & ")" & _
          " VALUES ('" & pWorkarea & "', '" & pAccDt & "', " & CStr(pAccSeq) & "," & _
          "     '" & pTestCd & "', " & CStr(pMfySeq) & ", " & CStr(pSeq) & "," & _
          "     '" & pMnmCd & "', '" & pMicFg & "', '" & pMqtCd & "', " & CStr(pSCnt) & pResult & ")"
    
    InsertMICResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 항생제결과의 MAX(seq) 조회 쿼리
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectMaxSeq(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT MAX(seq)+1 AS MAXSEQ FROM " & TIIS505 & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq)
    
    SelectMaxSeq = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : Vitek No를 이용해 접수일자(AccDt) 조회쿼리 - 1.1.6:원오세(2005-01-03)
'   인수 :
'       - pWorkarea : Workarea
'       - pMonth    : 접수월
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SelectAccDt(ByVal pWorkarea As String, ByVal pMonth As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT accdt FROM " & TIIS504 & _
          " WHERE workarea='" & pWorkarea & "' AND substr(accdt,5)='" & pMonth & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & _
          " ORDER BY accdt DESC"
          
    SelectAccDt = SQL
End Function

