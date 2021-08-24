VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Const F_OCSRESULT = "oram1.mdresult"          '결과관리테이블
'Private Const F_OCSRSLTHT = "oram1.mdrsltht"          '결과Histoty관리테이블
'Private Const F_OCSRESULTINFOG = "oras1.sg2examt"     '일반건진결과정보관리테이블
'Private Const F_OCSRESULTINFOT = "oras1.su2examt"     '종합건진결과정보관리테이블
'Private Const F_OCSRESULTSTSCD = "oram1.mdexmort"     '검사진행상태관리테이블

'-- Test =========================================================================
Private Const F_OCSRESULT = "mdresult"          '결과관리테이블
Private Const F_OCSRSLTHT = "mdrsltht"          '결과Histoty관리테이블
Private Const F_OCSRESULTINFOG = "sg2examt"     '일반건진결과정보관리테이블
Private Const F_OCSRESULTINFOT = "su2examt"     '종합건진결과정보관리테이블
Private Const F_OCSRESULTINFOM = "mdexmort"     '일반/종건이외결과정보관리테이블
Private Const F_OCSPTAREA = "apipdlst"          '입원환자내역

Private Const F_OCSSLMICWMT = "oras1.slmicwmt"        '미생물검사결과(MASTER)
Private Const F_OCSSLMICWDT = "oras1.slmicwdt"        '미생물검사결과(DETAIL)
Private Const F_OCSSLXSENST = "oras1.slxsenst"        '미생물감수성결과
Private Const F_OCSSLXBACTT = "oras1.slxbactt"        '미생물반응균결과
Private Const F_OCSSLXBACCT = "oras1.slbactct"        '미생물반응균코드
'=================================================================================

'-- Key 값 =======================================================================
Private mvarPtID            As String           '환자번호
Private mvarOrdDt           As String           '처방일자
Private mvarOrdNo           As String           '처방번호
'=================================================================================

Private mvarRstVal          As String           '일반결과값
Private mvarRstText         As String           '텍스트결과값
Private mvarItemVal         As String           '검사항목별결과값
Private mvarRefValFrom      As String           '참고치 From 값
Private mvarRefValTo        As String           '참고치 To 값
Private mvarRstUnit         As String           '검사항목 단위
Private mvarFootNote        As String           'FootfNote
Private mvarVfyDt           As String           '결과확인일자
Private mvarVfyTm           As String           '결과확인시간
Private mvarVfyID           As String           '결과확인자
Private mvarReadDt          As String           '판독일자
Private mvarReadTm          As String           '판독시간
Private mvarReadDate        As String           '판독일자
Private mvarReadDr          As String           '판독의사
Private mvarMfyID           As String           '수정자ID
Private mvarIPAdr           As String           '수정자IP
Private mvarMfyDt           As String           '수정일자
Private mvarMfyTm           As String           '수정시간

'-- 결과전송 시 OCS측 필요정보 ===================================================
Private mvarSpcCd           As String           '검체코드
Private mvarColDt           As String           '채취일자
Private mvarColTm           As String           '채취시간
Private mvarColID           As String           '채취자
Private mvarRcvDt           As String           '접수일자
Private mvarRcvTm           As String           '접수시간
Private mvarRcvID           As String           '접수자

Private mvarSlipCd          As String           'SlipCd(OCS WorkArea)
Private mvarRsltType        As String           'RsltType
Private mvarWrkDiv          As String           'WrkDiv(업무구분 : '3'종합건진, '4'일반건진)
'================================================================================

'-- Status Follow Up ============================================================
Private Const F_Collection = "C"
Private Const F_Accession = "E"
'=================================================================================

'** IP Address 관련 변수/////////////////////////////////////////////////////////////////////////////
Const MAX_IP = 5

Type IPINFO
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize  As Long
    unused1 As Integer
    unused2 As Integer
End Type

Type MIB_IPADDRTABLE
    dEntrys As Long
    mIPInfo(MAX_IP) As IPINFO
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
'////////////////////////////////////////////////////////////////////////////////////////////////////


Public Property Let RstVal(ByVal vData As String)
    mvarRstVal = vData
End Property

Public Property Get RstVal() As String
    RstVal = mvarRstVal
End Property

Public Property Let RstText(ByVal vData As String)
    mvarRstText = vData
End Property

Public Property Get RstText() As String
    RstText = mvarRstText
End Property

Public Property Let ItemVal(ByVal vData As String)
    mvarItemVal = vData
End Property

Public Property Get ItemVal() As String
    ItemVal = mvarItemVal
End Property

Public Property Let RefValFrom(ByVal vData As String)
    mvarRefValFrom = vData
End Property

Public Property Get RefValFrom() As String
    RefValFrom = mvarRefValFrom
End Property

Public Property Let RefValTo(ByVal vData As String)
    mvarRefValTo = vData
End Property

Public Property Get RefValTo() As String
    RefValTo = mvarRefValTo
End Property

Public Property Let RstUnit(ByVal vData As String)
    mvarRstUnit = vData
End Property

Public Property Get RstUnit() As String
    RstUnit = mvarRstUnit
End Property

Public Property Let FootNote(ByVal vData As String)
    mvarFootNote = vData
End Property

Public Property Get FootNote() As String
    FootNote = mvarFootNote
End Property

Public Property Let VfyDt(ByVal vData As String)
    mvarVfyDt = vData
End Property

Public Property Get VfyDt() As String
    VfyDt = mvarVfyDt
End Property

Public Property Let VfyTm(ByVal vData As String)
    mvarVfyTm = vData
End Property

Public Property Get VfyTm() As String
    VfyTm = mvarVfyTm
End Property

Public Property Let VfyID(ByVal vData As String)
    mvarVfyID = vData
End Property

Public Property Get VfyID() As String
    VfyID = mvarVfyID
End Property

Public Property Let ReadDt(ByVal vData As String)
    mvarVfyDt = vData
End Property

Public Property Get ReadDt() As String
    ReadDt = mvarVfyDt
End Property

Public Property Let ReadTm(ByVal vData As String)
    mvarVfyTm = vData
End Property

Public Property Get ReadTm() As String
    ReadTm = mvarVfyTm
End Property

Public Property Let MfyID(ByVal vData As String)
    mvarMfyID = vData
End Property

Public Property Get MfyID() As String
    MfyID = mvarMfyID
End Property

Public Property Let IPAdr(ByVal vData As String)
    mvarIPAdr = vData
End Property

Public Property Get IPAdr() As String
    IPAdr = mvarIPAdr
End Property

Public Property Let MfyDt(ByVal vData As String)
    mvarMfyDt = vData
End Property

Public Property Get MfyDt() As String
    MfyDt = mvarMfyDt
End Property

Public Property Let MfyTm(ByVal vData As String)
    mvarMfyTm = vData
End Property

Public Property Get MfyTm() As String
    MfyTm = mvarMfyTm
End Property


'-- 처방정보 =====================================================================
Public Property Let SpcCd(ByVal vData As String)
    mvarSpcCd = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mvarSpcCd
End Property

Public Property Let ColDt(ByVal vData As String)
    mvarColDt = vData
End Property

Public Property Get ColDt() As String
    ColDt = mvarColDt
End Property

Public Property Let ColTm(ByVal vData As String)
    mvarColTm = vData
End Property

Public Property Get ColTm() As String
    ColTm = mvarColTm
End Property

Public Property Let ColID(ByVal vData As String)
    mvarColID = vData
End Property

Public Property Get ColID() As String
    ColID = mvarColID
End Property

Public Property Let RcvDt(ByVal vData As String)
    mvarRcvDt = vData
End Property

Public Property Get RcvDt() As String
    RcvDt = mvarRcvDt
End Property

Public Property Let RcvTm(ByVal vData As String)
    mvarRcvTm = vData
End Property

Public Property Get RcvTm() As String
    RcvTm = mvarRcvTm
End Property

Public Property Let RcvID(ByVal vData As String)
    mvarRcvID = vData
End Property

Public Property Get RcvID() As String
    RcvID = mvarRcvID
End Property

Public Property Let SlipCd(ByVal vData As String)
    mvarSlipCd = vData
End Property

Public Property Get SlipCd() As String
    SlipCd = mvarSlipCd
End Property

Public Property Let RsltType(ByVal vData As String)
    mvarRsltType = vData
End Property

Public Property Get RsltType() As String
    RsltType = mvarRsltType
End Property

Public Property Let WrkDiv(ByVal vData As String)
    mvarWrkDiv = vData
End Property

Public Property Get WrkDiv() As String
    WrkDiv = mvarWrkDiv
End Property
'=================================================================================



Public Function SqlOCSRESULT_OLD(ByVal pPtid As String, ByVal pOrdDt As String, _
                                 ByVal pOrdNo As Integer, ByVal pTestCd As String) As Boolean
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & F_OCSRESULT & _
             "  where patno = " & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo) & _
             "    and examcode = " & DBS(pTestCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        mvarRstVal = "" & OrdRs.Fields("rslt1").Value
        mvarRstText = "" & OrdRs.Fields("rslt2").Value
        mvarRefValFrom = "" & OrdRs.Fields("rsltlow").Value
        mvarRefValTo = "" & OrdRs.Fields("rsltupp").Value
        mvarRstUnit = "" & OrdRs.Fields("rsltunit").Value
        mvarFootNote = "" & OrdRs.Fields("examtext").Value
        mvarRcvDt = Format(OrdRs.Fields("execdate"), "yyyymmdd")
        mvarRcvTm = Format(OrdRs.Fields("execdate"), "hhmmss")
        mvarVfyDt = Format(OrdRs.Fields("rsltdate"), "yyyymmdd")
        mvarVfyTm = Format(OrdRs.Fields("rsltdate"), "hhmmss")
        mvarVfyID = OrdRs.Fields("reptdr")
        mvarReadDt = Format(OrdRs.Fields("readdate"), "yyyymmdd")
        mvarMfyID = "" & OrdRs.Fields("editid").Value
        mvarIPAdr = "" & OrdRs.Fields("editip").Value
        mvarMfyDt = Format(OrdRs.Fields("editdate"), "yyyymmdd")
        mvarMfyTm = Format(OrdRs.Fields("editdate"), "hhmmss")
        
        mvarRsltType = "" & OrdRs.Fields("rslttype").Value
        mvarSlipCd = "" & OrdRs.Fields("slipcd").Value
        
        SqlOCSRESULT_OLD = True
    Else
        SqlOCSRESULT_OLD = False
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
    
End Function


'%  001. OCS측 결과테이블 RESULT INSERT_UPDATE_FLAG (MDRESULT)
'   'False' = INSERT, 'True' = UPDATE
Public Function SqlOCSRESULT_FLAG(ByVal pPtid As String, ByVal pOrdDt As String, _
                                    ByVal pOrdNo As Integer, ByVal pTestCd As String) As Boolean
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & F_OCSRESULT & _
             "  where patno = " & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo) & _
             "    and examcode = " & DBS(pTestCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
'        mvarRstVal = "" & OrdRs.Fields("rslt1").Value
'        mvarRstText = "" & OrdRs.Fields("rslt2").Value
'        mvarRefValFrom = "" & OrdRs.Fields("rsltlow").Value
'        mvarRefValTo = "" & OrdRs.Fields("rsltupp").Value
'        mvarRstUnit = "" & OrdRs.Fields("rsltunit").Value
'        mvarFootNote = "" & OrdRs.Fields("examtext").Value
'        mvarVfyDt = Format(OrdRs.Fields("rsltdate"), "yyyymmdd")
'        mvarVfyTm = Format(OrdRs.Fields("rsltdate"), "hhmmss")
'        mvarVfyID = OrdRs.Fields("reptdr")
'        mvarReadDt = Format(OrdRs.Fields("readdate"), "yyyymmdd")
'        mvarMfyID = "" & OrdRs.Fields("editid").Value
'        mvarIPAdr = "" & OrdRs.Fields("editip").Value
'        mvarMfyDt = Format(OrdRs.Fields("editdate"), "yyyymmdd")
'        mvarMfyTm = Format(OrdRs.Fields("editdate"), "hhmmss")
'
'        mvarRsltType = "" & OrdRs.Fields("rslttype").Value
'        mvarSlipCd = "" & OrdRs.Fields("slipcd").Value
        
        SqlOCSRESULT_FLAG = True
    Else
        SqlOCSRESULT_FLAG = False
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
    
End Function


'%  001. OCS측 결과테이블 RESULT INSERT (MDRESULT)
Public Function SqlOCSRESULT_INSERT(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                    ByVal pTestCd As String, ByVal pWorkArea As String, ByVal pResult1 As String, _
                                    ByVal pResult2 As String, ByVal pFromRef As String, ByVal pToRef As String, _
                                    ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
                                    ByVal pVfyDt As String, ByVal pVfyId As String, ByVal pReadDt As String, _
                                    ByVal pReadID As String, ByVal pMfyFinalDt As String, ByVal pTestType As String, _
                                    ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strVfyDt  As String
    Dim strReadDt As String
    'Dim strMfyFDt As String
    Dim strMfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If strMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    Select Case pTestCd
        Case "B4064", "S729TB", "X560TB", "CY7512", "B4021D", "B4052A", "B4021AC"
            pResult2 = Replace(pResult2, vbCrLf, "")
                
            If pResult2 = "" Then
                pTestType = "W"
            Else
                pTestType = "T"
            End If
    End Select
    
    SqlOCSRESULT_INSERT = " insert into " & F_OCSRESULT & _
                          " (patno, orddate, ordseqno, examcode, slipcd, rslt1, rslt2, rsltupp, rsltlow, rsltunit, " & _
                          "  examtext, execdate, rsltdate, reptdr, readdate, readdr, chngdate, rslttype, editid, editip, editdate) " & _
                          " values (" & DBS(pPtid) & "," & "TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & "," & DBN(pOrdNo) & "," & _
                          DBS(pTestCd) & "," & DBS(pWorkArea) & "," & DBS(pResult1) & "," & _
                          DBS(pResult2) & "," & DBS(pToRef) & "," & DBS(pFromRef) & "," & _
                          DBS(pRstUnit) & "," & DBS(pRemark) & "," & strRcvDt & "," & _
                          strVfyDt & "," & DBS(pVfyId) & "," & strVfyDt & "," & DBS(pReadID) & "," & _
                          strMfyDt & "," & DBS(pTestType) & "," & DBS(pMfyID) & "," & _
                          DBS(pMfyIP) & "," & strMfyDt & ")"

End Function

'%  001. OCS측 결과테이블 RESULT INSERT (MDRESULT) - NoGrowth
Public Function SqlOCSRESULT_INSERT_NoGrowth(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                             ByVal pTestCd As String, ByVal pWorkArea As String, ByVal pResult1 As String, _
                                             ByVal pResult2 As String, ByVal pFromRef As String, ByVal pToRef As String, _
                                             ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
                                             ByVal pVfyDt As String, ByVal pVfyId As String, ByVal pReadDt As String, _
                                             ByVal pReadID As String, ByVal pMfyFinalDt As String, ByVal pTestType As String, _
                                             ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strVfyDt  As String
    Dim strReadDt As String
    'Dim strMfyFDt As String
    Dim strMfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If strMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSRESULT_INSERT_NoGrowth = " insert into " & F_OCSRESULT & _
                                   " (patno, orddate, ordseqno, examcode, slipcd, rslt1, rslt2, rsltupp, rsltlow, rsltunit, " & _
                                   "  examtext, execdate, rsltdate, reptdr, readdate, readdr, chngdate, rslttype, editid, editip, editdate) " & _
                                   " values (" & DBS(pPtid) & "," & "TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & "," & DBN(pOrdNo) & "," & _
                                   DBS(pTestCd) & "," & DBS(pWorkArea) & "," & DBS(pResult1) & "," & _
                                   DBS(pResult2) & "," & DBS(pToRef) & "," & DBS(pFromRef) & "," & _
                                   DBS(pRstUnit) & "," & DBS(pRemark) & "," & strRcvDt & "," & _
                                   strVfyDt & "," & DBS(pVfyId) & "," & strVfyDt & "," & DBS(pReadID) & "," & _
                                   strMfyDt & "," & DBS(pTestType) & "," & DBS(pMfyID) & "," & _
                                   DBS(pMfyIP) & "," & strMfyDt & ")"

End Function

'%  001. OCS측 결과테이블 RESULT UPDATE (MDRESULT) - NoGrowth
Public Function SqlOCSRESULT_UPDATE_NoGrowth(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                             ByVal pTestCd As String, ByVal pWorkArea As String, ByVal pResult1 As String, _
                                             ByVal pResult2 As String, ByVal pFromRef As String, ByVal pToRef As String, _
                                             ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
                                             ByVal pVfyDt As String, ByVal pVfyId As String, ByVal pReadDt As String, _
                                             ByVal pReadID As String, ByVal pMfyFinalDt As String, ByVal pTestType As String, _
                                             ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strVfyDt  As String
    Dim strReadDt As String
    'Dim strMfyFDt As String
    Dim strMfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If strMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSRESULT_UPDATE_NoGrowth = " update " & F_OCSRESULT & _
                                   "    set rslt1 = " & DBS(pResult1) & ", rslt2 = " & DBS(pResult2) & _
                                   ", execdate = " & strRcvDt & ", " & "slipcd = " & DBS(pWorkArea) & _
                                   ", rsltdate = " & strVfyDt & ", " & "reptdr = " & DBS(pVfyId) & _
                                   ", readdate = " & strVfyDt & ", " & "readdr = " & DBS(pVfyId) & _
                                   "  where patno = " & DBS(pPtid) & _
                                   "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                   "    and ordseqno = " & DBN(pOrdNo) & _
                                   "    and examcode = " & DBS(pTestCd)
    
End Function

'%  002-1. OCS측 결과테이블 RESULT UPDATE(Modify) (MDRESULT)
'수정완료
'온승호
Public Function SqlOCSRESULT_SUB_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                        ByVal pTestCd As String, ByVal pResult1 As String, ByVal pResult2 As String, _
                                        ByVal pRemark As String, ByVal pVfyDt As String, ByVal pVfyId As String) As String
    '-- Date Type을 위한 변수선언
    Dim strVfyDt  As String
    Dim strTmpRst As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 수정일시
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = ""
    End If
    
    strTmpRst = Replace(pResult2, vbCrLf, "")
'    If Len(strTmpRst) > 0 And pTestCd = "B40516A" Then
    If Len(strTmpRst) > 0 Then
        SqlOCSRESULT_SUB_UPDATE = " update " & F_OCSRESULT & _
                                  "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
                                  "        examtext = " & DBS(pRemark) & "," & " rsltdate = " & strVfyDt & "," & _
                                  "        reptdr = " & DBS(pVfyId) & "," & " readdate = " & strVfyDt & "," & _
                                  "        readdr = " & pVfyId & "," & _
                                  "        rslttype = 'T'" & _
                                  "  where patno = " & DBS(pPtid) & _
                                  "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                  "    and ordseqno = " & DBN(pOrdNo) & _
                                  "    and examcode = " & DBS(pTestCd)
    Else
        SqlOCSRESULT_SUB_UPDATE = " update " & F_OCSRESULT & _
                                  "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
                                  "        examtext = " & DBS(pRemark) & "," & " rsltdate = " & strVfyDt & "," & _
                                  "        reptdr = " & DBS(pVfyId) & "," & " readdate = " & strVfyDt & "," & _
                                  "        readdr = " & pVfyId & _
                                  "  where patno = " & DBS(pPtid) & _
                                  "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                  "    and ordseqno = " & DBN(pOrdNo) & _
                                  "    and examcode = " & DBS(pTestCd)
    End If

End Function

'%  002-2. OCS측 결과테이블 RESULT UPDATE(Modify) (MDRESULT)
'수정완료
'온승호
Public Function SqlOCSRESULT_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                    ByVal pTestCd As String, ByVal pWorkArea As String, ByVal pResult1 As String, _
                                    ByVal pResult2 As String, ByVal pFromRef As String, ByVal pToRef As String, _
                                    ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
                                    ByVal pVfyDt As String, ByVal pVfyId As String, ByVal pReadID As String, _
                                    ByVal pMfyFinalDt As String, ByVal pTestType As String, ByVal pMfyID As String, _
                                    ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strMfyDt  As String
    Dim strTmpRst As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = ""
    End If
    
    strTmpRst = Replace(pResult2, vbCrLf, "")
'    If Len(strTmpRst) > 0 And pTestCd = "B40516A" Then
    If Len(strTmpRst) > 0 Then
        SqlOCSRESULT_UPDATE = " update " & F_OCSRESULT & _
                              "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
                              "        examtext = " & DBS(pRemark) & "," & " chngdate = " & strMfyDt & "," & _
                              "        editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & "," & _
                              "        editdate = " & strMfyDt & "," & _
                              "        rslttype = 'T'" & _
                              "  where patno = " & DBS(pPtid) & _
                              "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                              "    and ordseqno = " & DBN(pOrdNo) & _
                              "    and examcode = " & DBS(pTestCd)
    Else
        SqlOCSRESULT_UPDATE = " update " & F_OCSRESULT & _
                              "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
                              "        examtext = " & DBS(pRemark) & "," & " chngdate = " & strMfyDt & "," & _
                              "        editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & "," & _
                              "        editdate = " & strMfyDt & _
                              "  where patno = " & DBS(pPtid) & _
                              "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                              "    and ordseqno = " & DBN(pOrdNo) & _
                              "    and examcode = " & DBS(pTestCd)
    End If
End Function


'%  002-2. OCS측 결과테이블 RESULT UPDATE(Modify) (MDRESULT)
Public Function SqlOCSRESULT_MODIFY(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                    ByVal pTestCd As String, ByVal pResult1 As String, ByVal pResult2 As String, _
                                    ByVal pRemark As String, ByVal pMfyFinalDt As String, ByVal pMfyID As String, _
                                    ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strMfyDt  As String
    Dim strTmpRst As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = ""
    End If
    
    strTmpRst = Replace(pResult2, vbCrLf, "")
'    If Len(strTmpRst) > 0 And pTestCd = "B40516A" Then
    If Len(strTmpRst) > 0 Then
        SqlOCSRESULT_MODIFY = " update " & F_OCSRESULT & _
                              "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
                              "        examtext = " & DBS(pRemark) & "," & " rsltdate = " & strMfyDt & "," & _
                              "        readdate = " & strMfyDt & "," & " chngdate = " & strMfyDt & "," & _
                              "        editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & "," & _
                              "        editdate = " & strMfyDt & "," & _
                              "        rslttype = 'T'" & _
                              "  where patno = " & DBS(pPtid) & _
                              "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                              "    and ordseqno = " & DBN(pOrdNo) & _
                              "    and examcode = " & DBS(pTestCd)
    Else
        SqlOCSRESULT_MODIFY = " update " & F_OCSRESULT & _
                              "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
                              "        examtext = " & DBS(pRemark) & "," & " rsltdate = " & strMfyDt & "," & _
                              "        readdate = " & strMfyDt & "," & " chngdate = " & strMfyDt & "," & _
                              "        editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & "," & _
                              "        editdate = " & strMfyDt & _
                              "  where patno = " & DBS(pPtid) & _
                              "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                              "    and ordseqno = " & DBN(pOrdNo) & _
                              "    and examcode = " & DBS(pTestCd)
    End If

End Function


'%  003. OCS측 결과테이블 RESULT UPDATE() (MDRESULT)
'   미생물 감수성 결과값 업데이트 처리
Public Function SqlOCSRESULT_SEN_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                        ByVal pTestCd As String, ByVal pResult2 As String) As String

    If pTestCd = "B40516A" Then
        SqlOCSRESULT_SEN_UPDATE = " update " & F_OCSRESULT & _
                                  "    set rslt2 = " & DBS(pResult2) & "," & _
                                  "        rslttype = 'T'" & _
                                  "  where patno = " & DBS(pPtid) & _
                                  "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                  "    and ordseqno = " & DBN(pOrdNo) & _
                                  "    and examcode = " & DBS(pTestCd)
    Else
        SqlOCSRESULT_SEN_UPDATE = " update " & F_OCSRESULT & _
                                  "    set rslt2 = " & DBS(pResult2) & _
                                  "  where patno = " & DBS(pPtid) & _
                                  "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                  "    and ordseqno = " & DBN(pOrdNo) & _
                                  "    and examcode = " & DBS(pTestCd)
    End If

End Function

'%  004. OCS측 검사진행상태Flag UPDATE (MDEXMORT)
Public Function SqlOCSSTSFLAG_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                     ByVal pFlag As String, ByVal pVfyDt As String, ByVal pVfyId As String, _
                                     ByVal pMfyID As String, ByVal pMfyIP As String) As String
    '-- Date Type을 위한 변수선언
    Dim strVfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 수정일시
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = ""
    End If
    
    SqlOCSSTSFLAG_UPDATE = " update " & F_OCSRESULTINFOM & _
                           "    set procstat = " & DBS(pFlag) & "," & _
                           "        rsltdate = " & strVfyDt & "," & _
                           "        cofmdr = " & DBS(pVfyId) & "," & " editid = " & DBS(pMfyID) & "," & _
                           "        editip = " & DBS(pMfyIP) & _
                           "  where patno = " & DBS(pPtid) & _
                           "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                           "    and ordseqno = " & DBN(pOrdNo)

End Function

'%  005. OCS측 결과History테이블 RESULT INSERT (MDRSLTHT)
Public Function SqlOCSRSLTHT_INSERT(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                    ByVal pTestCd As String, ByVal pSeqNo As Integer, ByVal pWorkArea As String, ByVal pResult1 As String, _
                                    ByVal pResult2 As String, ByVal pFromRef As String, ByVal pToRef As String, _
                                    ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
                                    ByVal pVfyDt As String, ByVal pVfyId As String, ByVal pReadDt As String, _
                                    ByVal pReadID As String, ByVal pMfyFinalDt As String, ByVal pTestType As String, _
                                    ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strVfyDt  As String
    Dim strMfyFDt As String
    Dim strMfyDt  As String
    Dim strMfyTm  As String
    Dim strReadDt As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
        strMfyFDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
'        strMfyDt = "''"
'        strMfyFDt = "TO_DATE(" & DBS(pMfyFinalDt) & ", 'yyyymmdd hh24:mi:ss')"
        
        '## 5.0.33: 최명길(2005-07-25)
        '여기가 수정 한 곳 By M.G.Choi 2005.07.25
        '============================================================================
            strMfyDt = Format(Date, "yyyymmdd")
            strMfyTm = Format(Date, "hhmmss")
            strMfyDt = "TO_DATE(" & DBS(strMfyDt & strMfyTm) & ", 'yyyymmdd hh24:mi:ss')"
            If strMfyFDt <> "" Then
                strMfyFDt = "TO_DATE(" & DBS(pMfyFinalDt) & ", 'yyyymmdd hh24:mi:ss')"
            Else
                strMfyFDt = strMfyDt
            End If
        '============================================================================
    End If
    
    SqlOCSRSLTHT_INSERT = " insert into " & F_OCSRSLTHT & _
                          " (patno, orddate, ordseqno, examcode, seqno, slipcd, rslt1, rslt2, rsltupp, rsltlow, rsltunit, " & _
                          "  examtext, execdate, rsltdate, reptdr, readdate, readdr, chngdate, rslttype, editid, editip, editdate) " & _
                          " values (" & DBS(pPtid) & "," & "TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & "," & DBN(pOrdNo) & "," & _
                          DBS(pTestCd) & "," & DBN(pSeqNo) & "," & DBS(pWorkArea) & "," & DBS(pResult1) & "," & _
                          DBS(pResult2) & "," & DBS(pFromRef) & "," & DBS(pToRef) & "," & _
                          DBS(pRstUnit) & "," & DBS(pRemark) & "," & strRcvDt & "," & _
                          strVfyDt & "," & DBS(pVfyId) & "," & strReadDt & "," & DBS(pReadID) & "," & _
                          strMfyFDt & "," & DBS(pTestType) & "," & DBS(pMfyID) & "," & _
                          DBS(pMfyIP) & "," & strMfyDt & ")"

End Function

'%  006. OCS측 일반건진 SELECT (SG2EXAMT)
Public Function SqlOCSSG2EXAMT_Find(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer) As Boolean
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & F_OCSRESULTINFOG & _
             "  where patno = " & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo)
                          
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        SqlOCSSG2EXAMT_Find = True
    Else
        SqlOCSSG2EXAMT_Find = False
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  007. OCS측 일반건진 UPDATE (SG2EXAMT)
Public Function SqlOCSSG2EXAMT_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                      ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                      ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                      ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                      ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSG2EXAMT_UPDATE = " update " & F_OCSRESULTINFOG & _
                            "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                            "," & " detail = " & DBS(pRemark) & _
                            "," & " reptdate = " & strVfyDt & _
                            "," & " readdate = " & strReadDt & "," & " readdr = " & DBS(pReadDr) & _
                            "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                            "," & " editdate = " & strMfyDt & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)

End Function

'%  007. OCS측 일반건진 UPDATE (SG2EXAMT)
Public Function SqlOCSSG2EXAMT_UPDATE_New2(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                      ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                      ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                      ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                      ByVal pMfyIP As String, ByVal pMfyDt As String, ByVal pTextRst As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    Dim sLogFile  As String
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSG2EXAMT_UPDATE_New2 = " update " & F_OCSRESULTINFOG & _
                            "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                            "," & " detail = " & DBS(pRemark) & _
                            "," & " reptdate = " & strVfyDt & _
                            "," & " readdate = " & strReadDt & "," & " readdr = " & DBS(pReadDr) & _
                            "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                            "," & " editdate = " & strMfyDt & "," & " rsltcomm = " & DBS(pTextRst) & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)
    On Error Resume Next
                            
'    sLogFile = "\\172.16.10.180\pomis 용\Log\" & Format(Now, "yyyy-mm-dd") & "-SqlOCSSG2EXAMT_UPDATE_NEW2" & ".log"
'    If Len(Dir(sLogFile)) = 0 Then
'        Open sLogFile For Output As #1
'        Close #1
'    End If
'
'    Open sLogFile For Append As #1
'        Print #1, Format(Now, "hh:nn:ss")
'        Print #1, "==========================================="
'        Print #1, SqlOCSSG2EXAMT_UPDATE_New2
'        Print #1, "==========================================="
'    Close #1

End Function

'%  007. OCS측 일반건진 MODIFY (SG2EXAMT)
Public Function SqlOCSSG2EXAMT_MODIFY_SEN(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                          ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                          ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String, _
                                          Optional ByVal pLStsCd As String = "", Optional ByVal pDoneFg As String = "") As String
    '-- Date Type을 위한 변수선언
    Dim strMfyDt  As String '수정일자

    '** 체크사항
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If

    SqlOCSSG2EXAMT_MODIFY_SEN = " update " & F_OCSRESULTINFOG & _
                                "    set result = " & DBS(pResult) & "," & " detail = " & DBS(pRemark) & _
                                "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                                "," & " editdate = " & strMfyDt & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)

End Function

'%  007. OCS측 일반건진 UPDATE (SG2EXAMT)
Public Function SqlOCSSG2EXAMT_UPDATE_NEW(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                          ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                          ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                          ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                          ByVal pMfyIP As String, ByVal pMfyDt As String, Optional ByVal pLStsCd As String = "", _
                                          Optional ByVal pDoneFg As String = "") As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    Dim sLogFile  As String
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSG2EXAMT_UPDATE_NEW = " update " & F_OCSRESULTINFOG & _
                                "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                                "," & " detail = " & DBS(pRemark) & "," & " examdate = " & strRcvDt & _
                                "," & " reptdate = " & strVfyDt & "," & " examdr = " & DBS(pExamDr) & _
                                "," & " readdate = " & strReadDt & "," & " readdr = " & DBS(pReadDr) & _
                                "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                                "," & " editdate = " & strMfyDt & "," & " stscd = " & DBS(pLStsCd) & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)
    On Error Resume Next
                                                                   
'    sLogFile = "\\172.16.10.180\pomis 용\Log\" & Format(Now, "yyyy-mm-dd") & "-SqlOCSSG2EXAMT_UPDATE_NEW" & ".log"
'    If Len(Dir(sLogFile)) = 0 Then
'        Open sLogFile For Output As #1
'        Close #1
'    End If
'
'    Open sLogFile For Append As #1
'        Print #1, Format(Now, "hh:nn:ss")
'        Print #1, "==========================================="
'        Print #1, SqlOCSSG2EXAMT_UPDATE_NEW
'        Print #1, "==========================================="
'    Close #1

End Function

'%  007. OCS측 일반건진 UPDATE (SG2EXAMT)
Public Function SqlOCSSG2EXAMT_UPDATE_NEW3(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                          ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                          ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                          ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                          ByVal pMfyIP As String, ByVal pMfyDt As String, Optional ByVal pLStsCd As String = "", _
                                          Optional ByVal pDoneFg As String = "", Optional ByVal pTextRst As String = "") As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSG2EXAMT_UPDATE_NEW3 = " update " & F_OCSRESULTINFOG & _
                                "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                                "," & " detail = " & DBS(pRemark) & "," & " examdate = " & strRcvDt & _
                                "," & " reptdate = " & strVfyDt & "," & " examdr = " & DBS(pExamDr) & _
                                "," & " readdate = " & strReadDt & "," & " readdr = " & DBS(pReadDr) & _
                                "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                                "," & " editdate = " & strMfyDt & "," & " stscd = " & DBS(pLStsCd) & _
                                "," & " rsltcomm = " & DBS(pTextRst) & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)

End Function


'%  008. OCS측 종합건진 SELECT (SU2EXAMT)
Public Function SqlOCSSU2EXAMT_Find(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer) As Boolean
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & F_OCSRESULTINFOT & _
             "  where patno = " & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo)
                          
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        SqlOCSSU2EXAMT_Find = True
    Else
        SqlOCSSU2EXAMT_Find = False
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  009. OCS측 종합건진 UPDATE (SU2EXAMT)
Public Function SqlOCSSU2EXAMT_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                      ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                      ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                      ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                      ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSU2EXAMT_UPDATE = " update " & F_OCSRESULTINFOT & _
                            "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                            "," & " detail = " & DBS(pRemark) & _
                            "," & " reptdate = " & strVfyDt & _
                            "," & " readdate = " & strVfyDt & "," & " readdr = " & DBS(pReadDr) & _
                            "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                            "," & " editdate = " & strMfyDt & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)
    
End Function

'%  009. OCS측 종합건진 UPDATE (SU2EXAMT)
Public Function SqlOCSSU2EXAMT_UPDATE_New2(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                      ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                      ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                      ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                      ByVal pMfyIP As String, ByVal pMfyDt As String, ByVal pTextRst As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSU2EXAMT_UPDATE_New2 = " update " & F_OCSRESULTINFOT & _
                            "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                            "," & " detail = " & DBS(pRemark) & _
                            "," & " reptdate = " & strVfyDt & _
                            "," & " readdate = " & strVfyDt & "," & " readdr = " & DBS(pReadDr) & _
                            "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                            "," & " editdate = " & strMfyDt & "," & " rsltcomm = " & DBS(pTextRst) & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)
    
End Function

'%  009. OCS측 종합건진 Modify (SU2EXAMT)
Public Function SqlOCSSU2EXAMT_MODIFY_SEN(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                          ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                          ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String, _
                                          Optional ByVal pLStsCd As String, Optional ByVal pDoneFg As String) As String
    '-- Date Type을 위한 변수선언
    Dim strMfyDt  As String '수정일자

    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If

    SqlOCSSU2EXAMT_MODIFY_SEN = " update " & F_OCSRESULTINFOT & _
                                "    set result = " & DBS(pResult) & _
                                "," & " detail = " & DBS(pRemark) & "," & " editid = " & DBS(pMfyID) & _
                                "," & " editip = " & DBS(pMfyIP) & "," & " editdate = " & strMfyDt & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)

End Function

'%  009. OCS측 종합건진 UPDATE (SU2EXAMT)
Public Function SqlOCSSU2EXAMT_UPDATE_New(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                          ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                          ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                          ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                          ByVal pMfyIP As String, ByVal pMfyDt As String, Optional ByVal pLStsCd As String, _
                                          Optional ByVal pDoneFg As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSU2EXAMT_UPDATE_New = " update " & F_OCSRESULTINFOT & _
                                "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                                "," & " detail = " & DBS(pRemark) & "," & " examdate = " & strRcvDt & _
                                "," & " reptdate = " & strVfyDt & "," & " examdr = " & DBS(pExamDr) & _
                                "," & " readdate = " & strVfyDt & "," & " readdr = " & DBS(pReadDr) & _
                                "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                                "," & " editdate = " & strMfyDt & "," & " stscd = " & DBS(pLStsCd) & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)
    
End Function

'%  009. OCS측 종합건진 UPDATE (SU2EXAMT)
Public Function SqlOCSSU2EXAMT_UPDATE_New3(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                          ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                          ByVal pExamDt As String, ByVal pReptDt As String, ByVal pExamDr As String, _
                                          ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                          ByVal pMfyIP As String, ByVal pMfyDt As String, Optional ByVal pLStsCd As String, _
                                          Optional ByVal pDoneFg As String, Optional ByVal pTextRst As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String '검사일자(Interface일자)
    Dim strVfyDt  As String '보고일자
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '   일단 보고일자와 판독일자는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자(Interface 일자)
    If pExamDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pExamDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pReptDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pReptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSU2EXAMT_UPDATE_New3 = " update " & F_OCSRESULTINFOT & _
                                "    set examstat = " & DBS(pStsCd) & "," & " result = " & DBS(pResult) & _
                                "," & " detail = " & DBS(pRemark) & "," & " examdate = " & strRcvDt & _
                                "," & " reptdate = " & strVfyDt & "," & " examdr = " & DBS(pExamDr) & _
                                "," & " readdate = " & strVfyDt & "," & " readdr = " & DBS(pReadDr) & _
                                "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                                "," & " editdate = " & strMfyDt & "," & " stscd = " & DBS(pLStsCd) & _
                                "," & " rsltcomm = " & DBS(pTextRst) & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)
    
End Function

''%  006. OCS측 결과History테이블 RESULT UPDATE(Modify) (MDRSLTHT)
'Public Function SqlOCSRSLTHT_UPDATE(ByVal pPtID As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
'                                    ByVal pTestCd As String, ByVal pWorkArea As String, ByVal pResult1 As String, _
'                                    ByVal pResult2 As String, ByVal pToRef As String, ByVal pFromRef As String, _
'                                    ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
'                                    ByVal pVfyDt As String, ByVal pVfyID As String, ByVal pReadID As String, _
'                                    ByVal pMfyFinalDt As String, ByVal pTestType As String, ByVal pMfyID As String, _
'                                    ByVal pMfyIP As String, ByVal pMfyDt As String) As String
'
'    SqlOCSRSLTHT_UPDATE = " update " & F_OCSRESULT & _
'                          "    set rslt1 = " & DBS(pResult1) & "," & " rslt2 = " & DBS(pResult2) & "," & _
'                          "        examtext = " & DBS(pRemark) & "," & " chngdate = " & DBS(pMfyFinalDt) & "," & _
'                          "        editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & "," & _
'                          "        editdate = TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')" & _
'                          "  where patno = " & DBS(pPtID) & _
'                          "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
'                          "    and ordseqno = " & DBN(pOrdNo) & _
'                          "    and examcode = " & DBS(pTestCd)
'
'End Function

'%  010. OCS측 일건/종건 이외 SELECT (MDEXMORT)
Public Function SqlOCSMDEXMORT_Find(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer) As Boolean
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
'    strSQL = " select * from " & F_OCSRESULTINFOM & _
'             "  where patno = " & DBS(pPtid) & _
'             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
'             "    and ordseqno = " & DBN(pOrdNo)
'
    If Mid(pPtid, 1, 1) = "L" Then
        strSQL = " select * from s2ord999 " & _
                 "  where patno = " & DBS(pPtid) & _
                 "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                 "    and ordseqno = " & DBN(pOrdNo)
    Else
        strSQL = " select * from " & F_OCSRESULTINFOM & _
                 "  where patno = " & DBS(pPtid) & _
                 "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                 "    and ordseqno = " & DBN(pOrdNo)
    End If
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        SqlOCSMDEXMORT_Find = True
    Else
        SqlOCSMDEXMORT_Find = False
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  011. OCS측 OCS측 일건/종건 이외 UPDATE (MDEXMORT)
Public Function SqlOCSMDEXMORT_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                      ByVal pFlag As String, ByVal pVfyDt As String, ByVal pVfyId As String, _
                                      ByVal pMfyID As String, ByVal pMfyIP As String) As String
    '-- Date Type을 위한 변수선언
    Dim strVfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 결과일시
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
                               
    If Mid(pPtid, 1, 1) = "L" Then
        SqlOCSMDEXMORT_UPDATE = " update s2ord999 " & _
                                "    set procstat = " & DBS(pFlag) & "," & _
                                "        readdate = " & strVfyDt & "," & _
                                "        rsltdate = " & strVfyDt & "," & _
                                "        cofmdr = " & DBS(pVfyId) & "," & " editid = " & DBS(pMfyID) & "," & _
                                "        editip = " & DBS(pMfyIP) & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)
    Else
        SqlOCSMDEXMORT_UPDATE = " update " & F_OCSRESULTINFOM & _
                                "    set procstat = " & DBS(pFlag) & "," & _
                                "        readdate = " & strVfyDt & "," & _
                                "        rsltdate = " & strVfyDt & "," & _
                                "        cofmdr = " & DBS(pVfyId) & "," & " editid = " & DBS(pMfyID) & "," & _
                                "        editip = " & DBS(pMfyIP) & _
                                "  where patno = " & DBS(pPtid) & _
                                "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                "    and ordseqno = " & DBN(pOrdNo)

    End If
End Function

'%  011. OCS측 OCS측 일건/종건 이외 UPDATE (MDEXMORT)
Public Function SqlOCSMDEXMORT_UPDATE_NEW(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                      ByVal pFlag As String, ByVal pVfyDt As String, ByVal pVfyId As String, _
                                      ByVal pMfyID As String, ByVal pMfyIP As String, _
                                      ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strVfyDt  As String
    Dim strMfyDt  As String
    
    '** 체크사항
    '-- 결과일시
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSMDEXMORT_UPDATE_NEW = " update " & F_OCSRESULTINFOM & _
                            "    set procstat = " & DBS(pFlag) & "," & _
                            "        readdate = " & strVfyDt & "," & _
                            "        rsltdate = " & strVfyDt & "," & _
                            "        cofmdr = " & DBS(pVfyId) & "," & " editid = " & DBS(pMfyID) & "," & _
                            "        editip = " & DBS(pMfyIP) & "," & " editdate = " & strMfyDt & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)

End Function

'%  012. OCS측 결과History테이블 최종수정순번 가져오기 (MDRSLTHT)
Public Function SqlOCSRESULT_HISTORY(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                     ByVal pTestCd As String) As String
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String

    strSQL = " select max(seqno) as seq from " & F_OCSRSLTHT & _
             "  where patno = " & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo) & _
             "    and examcode = " & DBS(pTestCd) & _
             "  group by patno, orddate, ordseqno, examcode "
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        SqlOCSRESULT_HISTORY = OrdRs.Fields("seq")
    Else
        SqlOCSRESULT_HISTORY = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  009. OCS측 결과전송 시 필요정보 가져오기. (S2ORD102 AND S2LAB201)
Public Sub SqlOCS_Info(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                            ByVal pTestCd As String)
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String

    strSQL = " select a.slipcd, a.wrkdiv, b.* " & _
             "   from " & T_LAB102 & " a, " & T_LAB201 & " b " & _
             "  where a.ptid = " & DBS(pPtid) & _
             "    and a.orddt = " & DBS(pOrdDt) & _
             "    and a.ordno = " & DBN(pOrdNo) & _
             "    and b.workarea = a.workarea " & _
             "    and b.accdt = a.accdt " & _
             "    and b.accseq = a.accseq "
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        mvarSpcCd = "" & OrdRs.Fields("spccd").Value
        mvarColDt = "" & OrdRs.Fields("coldt").Value
        mvarColTm = "" & OrdRs.Fields("coltm").Value
        mvarColID = "" & OrdRs.Fields("colid").Value
        mvarRcvDt = "" & OrdRs.Fields("rcvdt").Value
        mvarRcvTm = "" & OrdRs.Fields("rcvtm").Value
        mvarRcvID = "" & OrdRs.Fields("rcvid").Value
        mvarVfyDt = "" & OrdRs.Fields("vfydt").Value
        mvarVfyTm = "" & OrdRs.Fields("vfytm").Value
        mvarVfyID = "" & OrdRs.Fields("vfyid").Value
        
        mvarSlipCd = "" & OrdRs.Fields("slipcd").Value
        mvarWrkDiv = "" & OrdRs.Fields("wrkdiv").Value
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Sub

'%  009-1. OCS측 결과전송 시 필요정보 가져오기. (S2LAB302 AND S2LAB201 AND S2ORD102)
'**** 중요사항 : 일반검사는 LAB302 기준으로 한다.
Public Sub SqlOCS_Info_New(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
                           
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    '** 수정 전 =======================================================================
'    strSQL = " select c.slipcd, c.wrkdiv, b.* " & _
'             "   from " & T_LAB302 & " a, " & T_LAB201 & " b, " & T_LAB102 & " c " & _
'             "  where a.workarea = " & DBS(pWorkArea) & _
'             "    and a.accdt = " & DBS(pAccDt) & _
'             "    and a.accseq = " & DBN(pAccSeq) & _
'             "    and b.workarea = a.workarea " & _
'             "    and b.accdt = a.accdt " & _
'             "    and b.accseq = a.accseq " & _
'             "    and c.ptid = a.ptid " & _
'             "    and c.orddt = a.orddt " & _
'             "    and c.ordno = a.ordno "
    '==================================================================================
    
    '** 수정 전 =======================================================================
    strSQL = " select b.slipcd, b.wrkdiv, a.* " & _
             "   from " & T_LAB201 & " a, " & T_LAB102 & " b " & _
             "  where a.workarea = " & DBS(pWorkArea) & _
             "    and a.accdt = " & DBS(pAccDt) & _
             "    and a.accseq = " & DBN(pAccSeq) & _
             "    and b.workarea = a.workarea " & _
             "    and b.accdt = a.accdt " & _
             "    and b.accseq = a.accseq "
    '==================================================================================
             
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        mvarSpcCd = "" & OrdRs.Fields("spccd").Value
        mvarColDt = "" & OrdRs.Fields("coldt").Value
        mvarColTm = "" & OrdRs.Fields("coltm").Value
        mvarColID = "" & OrdRs.Fields("colid").Value
        mvarRcvDt = "" & OrdRs.Fields("rcvdt").Value
        mvarRcvTm = "" & OrdRs.Fields("rcvtm").Value
        mvarRcvID = "" & OrdRs.Fields("rcvid").Value
        mvarVfyDt = "" & OrdRs.Fields("vfydt").Value
        mvarVfyTm = "" & OrdRs.Fields("vfytm").Value
        mvarVfyID = "" & OrdRs.Fields("vfyid").Value
        
        mvarSlipCd = "" & OrdRs.Fields("slipcd").Value
        mvarWrkDiv = "" & OrdRs.Fields("wrkdiv").Value
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Sub

'%  009-2. OCS측 결과전송 시 필요정보 가져오기. (S2LAB404 AND S2LAB201 AND S2ORD102)
'**** 중요사항 : 미생물검사는 LAB404 기준으로 한다.
Public Sub SqlOCS_Info_Mic(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
                           
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select c.slipcd, c.wrkdiv, b.* " & _
             "   from " & T_LAB404 & " a, " & T_LAB201 & " b, " & T_LAB102 & " c " & _
             "  where a.workarea = " & DBS(pWorkArea) & _
             "    and a.accdt = " & DBS(pAccDt) & _
             "    and a.accseq = " & DBN(pAccSeq) & _
             "    and b.workarea = a.workarea " & _
             "    and b.accdt = a.accdt " & _
             "    and b.accseq = a.accseq " & _
             "    and c.ptid = a.ptid " & _
             "    and c.orddt = a.orddt " & _
             "    and c.ordno = a.ordno "
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        mvarSpcCd = "" & OrdRs.Fields("spccd").Value
        mvarColDt = "" & OrdRs.Fields("coldt").Value
        mvarColTm = "" & OrdRs.Fields("coltm").Value
        mvarColID = "" & OrdRs.Fields("colid").Value
        mvarRcvDt = "" & OrdRs.Fields("rcvdt").Value
        mvarRcvTm = "" & OrdRs.Fields("rcvtm").Value
        mvarRcvID = "" & OrdRs.Fields("rcvid").Value
        mvarVfyDt = "" & OrdRs.Fields("vfydt").Value
        mvarVfyTm = "" & OrdRs.Fields("vfytm").Value
        mvarVfyID = "" & OrdRs.Fields("vfyid").Value
        
        mvarSlipCd = "" & OrdRs.Fields("slipcd").Value
        mvarWrkDiv = "" & OrdRs.Fields("wrkdiv").Value
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Sub

'%  009-2. OCS측 결과전송 시 미생물 반응균 관련 필요정보 가져오기. (S2LAB201)
Public Function SqlOCS_MICInfo(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String
                           
    Dim OrdRs       As New ADODB.Recordset
    Dim strSQL      As String
    Dim strSpcCd    As String
    Dim strBussDiv  As String
    Dim strDeptCd   As String
    Dim strWardId   As String
    Dim strRcvDt    As String
    Dim strRcvTm    As String
    Dim strVfyDt    As String
    Dim strVfyTm    As String
    Dim strVfyId    As String
    
    strSQL = " select * " & _
             "   from " & T_LAB201 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq)

    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        strSpcCd = "" & OrdRs.Fields("spccd").Value
        strWardId = "" & OrdRs.Fields("wardid").Value
        If strWardId <> "" Then
            strBussDiv = "I"    '입원
        Else
            strBussDiv = "O"    '외래
        End If
        strDeptCd = "" & OrdRs.Fields("deptcd").Value
        strRcvDt = "" & OrdRs.Fields("rcvdt").Value
        strRcvTm = "" & OrdRs.Fields("rcvtm").Value
        strVfyDt = "" & OrdRs.Fields("vfydt").Value
        strVfyTm = "" & OrdRs.Fields("vfytm").Value
        strVfyId = "" & OrdRs.Fields("vfyid").Value
        
        SqlOCS_MICInfo = strSpcCd & vbTab & strBussDiv & vbTab & strDeptCd & vbTab & strWardId & _
                        vbTab & strRcvDt & strRcvTm & vbTab & strVfyDt & strVfyTm
    Else
        SqlOCS_MICInfo = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  000. '3' 종합건진과 '4' 일반검사 구분하기
Public Function SqlOCBussdiv_Flag(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer) As String
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select wrkdiv, slipcd from " & T_LAB102 & _
             "  where ptid = " & DBS(pPtid) & _
             "    and orddt = " & DBS(pOrdDt) & _
             "    and ordno = " & DBN(pOrdNo)
             
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        SqlOCBussdiv_Flag = "" & OrdRs.Fields("wrkdiv").Value
        mvarSlipCd = "" & OrdRs.Fields("slipcd").Value
    Else
        SqlOCBussdiv_Flag = ""
        mvarSlipCd = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  000. OCS측 결과테이블 RESULT UPDATE(FootNote) (MDRESULT)
Public Function SqlOCSRESULT_FOOTNOTE_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer) As String
    
'    SqlOCSRESULT_FOOTNOTE_UPDATE = " update " & F_OCSRESULT & _
'                                   "    set examtext = " & DBS(pRemark) & _
'                                   "  where patno = " & DBS(pPtID) & _
'                                   "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
'                                   "    and ordseqno = " & DBN(pOrdNo)
    MsgBox "SqlOCSRESULT_FOOTNOTE_UPDATE"
End Function

'%  검사항목별 결과코드 사용 시 결과값 및 텍스트 결과 유무를 직접조회하여 처리 한다.
Public Sub ResultValue_Find(ByVal pTestCd As String, ByVal pRstCd As String)
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    Dim strTxtType As String
    
    strSQL = " select a.txttype, b.field1 " & _
             "  from " & T_LAB001 & " a, " & T_LAB031 & " b " & _
             "  where a.testcd = " & DBS(pTestCd) & _
             "    and b.cdval1 (+)= a.testcd " & _
             "    and b.cdindex (+)= '" & LC2_ItemResult & "' " & _
             "    and b.cdval2 (+)= " & DBS(pRstCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        strTxtType = OrdRs.Fields("txttype").Value & ""
        
        If strTxtType = "2" Then
            mvarRsltType = "T"
        Else
            mvarRsltType = "W"
        End If
        
        mvarItemVal = OrdRs.Fields("field1").Value & ""
    Else
        mvarRsltType = ""
        mvarItemVal = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
    
End Sub

'%  검사항목별 결과코드 사용 시 결과값 및 텍스트 결과 유무를 직접조회하여 처리 한다.
'%  NoGrowth 는 별도로 처리 한다.
Public Sub ResultValue_NoGrowth(ByVal pTestCd As String, ByVal pRstCd As String)
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    Dim strTxtType As String
    
    strSQL = " select a.txttype, b.field1 " & _
             "  from " & T_LAB001 & " a, " & T_LAB031 & " b " & _
             "  where a.testcd = " & DBS(pTestCd) & _
             "    and b.cdval1 (+)= a.testcd " & _
             "    and b.cdindex (+)= '" & LC2_ItemResult & "' " & _
             "    and b.cdval2 (+)= " & DBS(pRstCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        strTxtType = OrdRs.Fields("txttype").Value & ""
        
        If strTxtType = "2" Then
            mvarRsltType = "T"
        Else
            mvarRsltType = "W"
        End If
        
        mvarItemVal = OrdRs.Fields("field1").Value & ""
    Else
        mvarRsltType = ""
        mvarItemVal = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
    
End Sub

Public Function NoGrowth_Test(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String) As String

    NoGrowth_Test = " select testcd from " & T_LAB404 & _
                    "  where workarea = " & DBS(pWorkArea) & _
                    "    and accdt    = " & DBS(pAccDt) & _
                    "    and accseq   = " & DBN(pAccSeq) & _
                    "    and testcd in (" & pTestCd & ")"
    
End Function


'%  균명을 조회한다.
Public Function MicroNm_Find(ByVal pMicCd As String, ByVal pQCCd As String) As String
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select a.text1, b.field2 from " & T_LAB032 & " a, " & T_LAB032 & " b " & _
             "  where a.cdindex = " & DBS(LC3_Microbe) & _
             "    and a.cdval1  = " & DBS(pMicCd) & _
             "    and b.cdindex = " & DBS(LC3_Volume) & _
             "    and b.cdval1  = " & DBS(pQCCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        MicroNm_Find = "[" & OrdRs.Fields("text1") & "]" & vbTab & OrdRs.Fields("field2")
    Else
        MicroNm_Find = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  균명을 조회한다.
Public Function QTYNm_Find(ByVal pMicCd As String, ByVal pQCCd As String) As String
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select field2 from " & T_LAB032 & _
             "  where cdindex = " & DBS(LC3_Volume) & _
             "    and cdval1  = " & DBS(pQCCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        QTYNm_Find = OrdRs.Fields("field2").Value & ""
    Else
        QTYNm_Find = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'%  감수성결과의 항생제명을 조회한다.
Public Function SenNm_Find(ByVal pSenCd As String) As String
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_LAB032 & _
             "  where cdindex = " & DBS(LC3_AntiBiotic) & _
             "    and cdval1  = " & DBS(pSenCd)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    '## 5.0.21: 이상대(2005-02-24)
    '   - 코드상 버그수정 (최과장님 미오!!)
    If OrdRs.EOF = False Then
        SenNm_Find = OrdRs.Fields("text1") & ""
    Else
        SenNm_Find = ""
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
End Function

'** 특수검사 처리시작################################################################################
Public Function SpecialTest_Main(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pTestCd As String, ByVal pVfyDt As String, ByVal pVfyTm As String, _
                                ByVal pVfyId As String, ByVal pTxtRst As String, ByVal pFNote As String, _
                                Optional ByVal pFlag As String = "0") As Boolean
    Dim strSQL    As String
    Dim strVfyDt  As String
    Dim strIPAdr  As String
    Dim strOCSSEQ As String
    
    On Error GoTo DBExecError
    
    strVfyDt = pVfyDt & pVfyTm
    
    If Mid$(pAccDt, 1, 1) = "9" Then
       pAccDt = "19" & pAccDt
    Else
       pAccDt = "20" & pAccDt
    End If
    
    '-- Key 정보 Set
    If OrderKey_Info(pWorkArea, pAccDt, pAccSeq) = False Then
        GoTo DBExecError
    End If
    
    '-- OCS 결과전송 정보 Set
    Call SqlOCS_Info_New(pWorkArea, pAccDt, pAccSeq)
    
    Select Case pFlag
        Case "0"
            '-- 결과등록 ======================================================================
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            If mvarWrkDiv = "3" Then
                If SqlOCSSU2EXAMT_Find(mvarPtID, mvarOrdDt, mvarOrdNo) = True Then
                    strSQL = SqlOCSSU2EXAMT_UPDATE(mvarPtID, mvarOrdDt, mvarOrdNo, _
                            "N", "", pFNote, mvarRcvDt & mvarRcvTm, strVfyDt, pVfyId, strVfyDt, pVfyId, "", "", "")
                    
                    Call DBConn.Execute(strSQL)
                    
                Else
                    GoTo DBExecError
                End If
            ElseIf mvarWrkDiv = "4" Then
                    If SqlOCSSG2EXAMT_Find(mvarPtID, mvarOrdDt, mvarOrdNo) = True Then
                        strSQL = SqlOCSSG2EXAMT_UPDATE(mvarPtID, mvarOrdDt, mvarOrdNo, _
                                "N", "", pFNote, mvarRcvDt & mvarRcvTm, strVfyDt, pVfyId, strVfyDt, pVfyId, "", "", "")
                                
                        Call DBConn.Execute(strSQL)
                        
                    Else
                        GoTo DBExecError
                    End If
            Else
                If SqlOCSMDEXMORT_Find(mvarPtID, mvarOrdDt, mvarOrdNo) = True Then
                    strSQL = SqlOCSMDEXMORT_UPDATE(mvarPtID, mvarOrdDt, mvarOrdNo, _
                            "N", strVfyDt, pVfyId, "", "")
                            
                    Call DBConn.Execute(strSQL)
                    
                Else
                    GoTo DBExecError
                End If
                
                '-- 결과저장
                If SqlOCSRESULT_FLAG(mvarPtID, mvarOrdDt, mvarOrdNo, pTestCd) = False Then
                    strSQL = SqlOCSRESULT_INSERT(mvarPtID, mvarOrdDt, mvarOrdNo, _
                            pTestCd, mvarSlipCd, "", pTxtRst, "", "", "", _
                            pFNote, mvarRcvDt & mvarRcvTm, strVfyDt, pVfyId, strVfyDt, pVfyId, "", "T", "", "", "")
                            
                    Call DBConn.Execute(strSQL)
                    
                Else
                    GoTo DBExecError
                End If
            End If
            
            
            '==================================================================================
            
        Case "1"
            '-- 결과수정 ======================================================================
            Dim strMfyDt    As String
            Dim strMfyTm    As String
            
            '-- 전주예수병원IP Address =============
            strIPAdr = LocalIP_Address
            '=======================================
            
            strMfyDt = Format(GetSystemDate, "yyyymmdd")
            strMfyTm = Format(GetSystemDate, "hhmmss")
            
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            If mvarWrkDiv = "3" Then
                If SqlOCSSU2EXAMT_Find(mvarPtID, mvarOrdDt, mvarOrdNo) = True Then
                    strSQL = SqlOCSSU2EXAMT_MODIFY(mvarPtID, mvarOrdDt, mvarOrdNo, _
                            "N", "", pFNote, strVfyDt, pVfyId, pVfyId, strIPAdr, strVfyDt)
                    
                    Call DBConn.Execute(strSQL)
                    
                Else
                    GoTo DBExecError
                End If
            ElseIf mvarWrkDiv = "4" Then
                    If SqlOCSSG2EXAMT_Find(mvarPtID, mvarOrdDt, mvarOrdNo) = True Then
                        strSQL = SqlOCSSG2EXAMT_MODIFY(mvarPtID, mvarOrdDt, mvarOrdNo, _
                                "N", "", pFNote, strVfyDt, pVfyId, pVfyId, strIPAdr, strVfyDt)
                                
                        Call DBConn.Execute(strSQL)
                        
                    Else
                        GoTo DBExecError
                    End If
            Else
                
                If SqlOCSMDEXMORT_Find(mvarPtID, mvarOrdDt, mvarOrdNo) = True Then
                    strSQL = SqlOCSMDEXMORT_UPDATE(mvarPtID, mvarOrdDt, mvarOrdNo, _
                            "N", strVfyDt, pVfyId, pVfyId, strIPAdr)
                            
                    Call DBConn.Execute(strSQL)
                    
                Else
                    GoTo DBExecError
                End If
                
                '-- 결과History내역 순번 가져오기
                strOCSSEQ = SqlOCSRESULT_HISTORY(mvarPtID, mvarOrdDt, mvarOrdNo, pTestCd)
                
                If strOCSSEQ = "" Then
                    strOCSSEQ = "1"
                Else
                    strOCSSEQ = CInt(strOCSSEQ) + 1
                End If
                
                If SqlOCSRESULT_FLAG(mvarPtID, mvarOrdDt, mvarOrdNo, pTestCd) = True Then
                    
                    strSQL = SqlOCSRESULT_UPDATE(mvarPtID, mvarOrdDt, mvarOrdNo, _
                            pTestCd, mvarSlipCd, "", pTxtRst, "", "", "", _
                            pFNote, "", "", "", "", strVfyDt, "T", pVfyId, strIPAdr, strVfyDt)
                            
                    Call DBConn.Execute(strSQL)
                    
                Else
                    GoTo DBExecError
                End If
                
                '-- OCS 결과History테이블 INSERT
                strSQL = SqlOCSRSLTHT_INSERT(mvarPtID, mvarOrdDt, mvarOrdNo, _
                        pTestCd, strOCSSEQ, mvarSlipCd, mvarRstVal, mvarRstText, _
                        mvarRefValFrom, mvarRefValTo, mvarRstUnit, _
                        mvarFootNote, mvarRcvDt & mvarRcvTm, mvarVfyDt & mvarVfyTm, _
                        mvarVfyID, mvarVfyDt & mvarVfyTm, mvarVfyID, _
                        strMfyDt & strMfyTm, mvarRsltType, _
                        mvarMfyID, mvarIPAdr, mvarMfyDt & mvarMfyTm)
                
                Call DBConn.Execute(strSQL)
                
            End If
            
            '==================================================================================
            
    End Select
    
    SpecialTest_Main = True
    
    Exit Function
    
DBExecError:
    SpecialTest_Main = False
    MsgBox "Error No.1"
    
End Function

Private Function OrderKey_Info(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " select ptid, orddt, ordno " & _
             "   from " & T_LAB351 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq)
    
    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If OrdRs.EOF = False Then
        mvarPtID = "" & OrdRs.Fields("ptid").Value
        mvarOrdDt = "" & OrdRs.Fields("orddt").Value
        mvarOrdNo = "" & OrdRs.Fields("ordno").Value
        
        OrderKey_Info = True
    Else
        mvarPtID = ""
        mvarOrdDt = ""
        mvarOrdNo = ""
        
        OrderKey_Info = False
    End If
    
    OrdRs.Close
    Set OrdRs = Nothing
    
End Function

'** 일건 결과수정시 업데이트
Private Function SqlOCSSG2EXAMT_MODIFY(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                       ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                       ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                       ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSG2EXAMT_MODIFY = " update " & F_OCSRESULTINFOG & _
                            "    set readdate = " & strReadDt & "," & " readdr = " & DBS(pReadDr) & _
                            "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                            "," & " editdate = " & strMfyDt & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)
                            
End Function

'** 종건 결과수정시 업데이트
Private Function SqlOCSSU2EXAMT_MODIFY(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                       ByVal pStsCd As String, ByVal pResult As String, ByVal pRemark As String, _
                                       ByVal pReadDt As String, ByVal pReadDr As String, ByVal pMfyID As String, _
                                       ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strReadDt As String '판독일자(결과저장)
    Dim strMfyDt  As String '수정일자
    
    '** 체크사항
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSSU2EXAMT_MODIFY = " update " & F_OCSRESULTINFOT & _
                            "    set readdate = " & strReadDt & "," & " readdr = " & DBS(pReadDr) & _
                            "," & " editid = " & DBS(pMfyID) & "," & " editip = " & DBS(pMfyIP) & _
                            "," & " editdate = " & strMfyDt & _
                            "  where patno = " & DBS(pPtid) & _
                            "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                            "    and ordseqno = " & DBN(pOrdNo)
                            
End Function

'** OCS Status Follow Up ==================================================================
Public Function OCS_Status_Update(ByVal pPtid As String, ByVal pOrdDt As String, _
                                  ByVal pOrdNo As String, ByVal pActDt As String, _
                                  ByVal pActTm As String, ByVal pStsCd As String) As Boolean
    
    Dim rs        As New ADODB.Recordset
    Dim strSQL    As String
    Dim strWrkDiv As String
    
    On Error GoTo ErrMsg
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(pPtid) & _
             "    and orddt = " & DBS(pOrdDt) & _
             "    and ordno = " & DBS(pOrdNo)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strWrkDiv = "" & rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    Select Case pStsCd
        Case "1"        '[B]로 업데이트
            If strWrkDiv = "3" Then         '종건
                Call SU2EXAMT_STS_UPDATE_NEW(pPtid, pOrdDt, pOrdNo, pActDt, pActTm, "B")
            ElseIf strWrkDiv = "4" Then     '일건
                    Call SG2EXAMT_STS_UPDATE_NEW(pPtid, pOrdDt, pOrdNo, pActDt, pActTm, "B")
            Else                            '이외검사
                Call MDEXMORT_STS_UPDATE(pPtid, pOrdDt, pOrdNo, pActDt, pActTm, "B")
            End If
        
        Case "2", "3"   '[E]로 업데이트
            If strWrkDiv = "3" Then         '종건
                Call SU2EXAMT_STS_UPDATE_NEW(pPtid, pOrdDt, pOrdNo, pActDt, pActTm, "E")
            ElseIf strWrkDiv = "4" Then     '일건
                    Call SG2EXAMT_STS_UPDATE_NEW(pPtid, pOrdDt, pOrdNo, pActDt, pActTm, "E")
            Else                            '이외검사
                Call MDEXMORT_STS_UPDATE(pPtid, pOrdDt, pOrdNo, pActDt, pActTm, "E")
            End If
        
        Case Else
            '결과는 결과전송시 Follow Up
            
    End Select
    
    Set rs = Nothing
    
    OCS_Status_Update = True
    
    Exit Function
    
ErrMsg:
    Set rs = Nothing
    OCS_Status_Update = False
    
End Function

''** 일건 Status Update
'Private Sub SG2EXAMT_STS_UPDATE(ByVal pPtID As String, ByVal pOrdDt As String, _
'                                ByVal pOrdNo As String, ByVal pStsCd As String)
'
'    Dim strSQL As String
'
'    strSQL = " update " & F_OCSRESULTINFOG & _
'             "    set examstat = " & DBS(pStsCd) & _
'             "  where patno =" & DBS(pPtID) & _
'             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
'             "    and ordseqno = " & DBN(pOrdNo)
'
'    Call DBConn.Execute(strSQL)
'
'End Sub
'
''** 종건 Status Update
'Private Sub SU2EXAMT_STS_UPDATE(ByVal pPtID As String, ByVal pOrdDt As String, _
'                                ByVal pOrdNo As String, ByVal pStsCd As String)
'
'    Dim strSQL As String
'
'    strSQL = " update " & F_OCSRESULTINFOT & _
'             "    set examstat = " & DBS(pStsCd) & _
'             "  where patno =" & DBS(pPtID) & _
'             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
'             "    and ordseqno = " & DBN(pOrdNo)
'
'    Call DBConn.Execute(strSQL)
'
'End Sub

'** 일건 Status Update
Private Sub SG2EXAMT_STS_UPDATE_NEW(ByVal pPtid As String, ByVal pOrdDt As String, _
                                    ByVal pOrdNo As String, ByVal pActDt As String, _
                                    ByVal pActTm As String, ByVal pStsCd As String)
    
    Dim strSQL   As String
    Dim strTemp  As String
    Dim strActDt As String
    
    strActDt = pActDt & pActTm
    
    If strActDt <> "" Then
        strActDt = "TO_DATE(" & DBS(pActDt & pActTm) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strActDt = "''"
    End If
    
    Select Case pStsCd
        Case "B"
            strTemp = "    set examstat = " & DBS(pStsCd) & _
                      ", colldate = " & strActDt
        
        Case "E"
            '-- 접수와 실시는 진단검사에서는 같은 시점.(접수=실시)
            strTemp = "    set examstat = " & DBS(pStsCd) & _
                      ", acptdate = " & strActDt & " , examdate = " & strActDt & _
                      ", examdr = " & DBS(ObjSysInfo.EmpId)
            
    End Select
    
    strSQL = " update " & F_OCSRESULTINFOG & _
             strTemp & _
             "  where patno =" & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo)
    
    Call DBConn.Execute(strSQL)
    
End Sub

'** 종건 Status Update
Private Sub SU2EXAMT_STS_UPDATE_NEW(ByVal pPtid As String, ByVal pOrdDt As String, _
                                    ByVal pOrdNo As String, ByVal pActDt As String, _
                                    ByVal pActTm As String, ByVal pStsCd As String)
    
    Dim strSQL   As String
    Dim strTemp  As String
    Dim strActDt As String
    
    strActDt = pActDt & pActTm
    
    If strActDt <> "" Then
        strActDt = "TO_DATE(" & DBS(pActDt & pActTm) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strActDt = "''"
    End If
    
    Select Case pStsCd
        Case "B"
            strTemp = "    set examstat = " & DBS(pStsCd) & _
                      ", colldate = " & strActDt
        
        Case "E"
            '-- 접수와 실시는 진단검사에서는 같은 시점.(접수=실시)
            strTemp = "    set examstat = " & DBS(pStsCd) & _
                      ", acptdate = " & strActDt & " , examdate = " & strActDt & _
                      ", examdr = " & DBS(ObjSysInfo.EmpId)
            
    End Select
    
    strSQL = " update " & F_OCSRESULTINFOT & _
             strTemp & _
             "  where patno =" & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo)
    
    Call DBConn.Execute(strSQL)
    
End Sub

'** 일건/종건 이외 Status Update
Private Sub MDEXMORT_STS_UPDATE(ByVal pPtid As String, ByVal pOrdDt As String, _
                                ByVal pOrdNo As String, ByVal pActDt As String, _
                                ByVal pActTm As String, ByVal pStsCd As String)
    
    Dim strSQL   As String
    Dim strTemp  As String
    Dim strActDt As String
    
    strActDt = pActDt & pActTm
    
    If strActDt <> "" Then
        strActDt = "TO_DATE(" & DBS(pActDt & pActTm) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strActDt = "''"
    End If
    
    Select Case pStsCd
        Case "B"
            strTemp = "    set procstat = " & DBS(pStsCd) & _
                      ", colldate = " & strActDt
        
        Case "E"
            '-- 접수와 실시는 진단검사에서는 같은 시점.(접수=실시)
            strTemp = "    set procstat = " & DBS(pStsCd) & _
                      ", acptdate = " & strActDt & " , execdate = " & strActDt
            
    End Select
    
    strSQL = " update " & F_OCSRESULTINFOM & _
             strTemp & _
             "  where patno =" & DBS(pPtid) & _
             "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
             "    and ordseqno = " & DBN(pOrdNo)
                     
    Call DBConn.Execute(strSQL)
    
End Sub


Public Function SqlGetRemark(ByVal pCode As String) As String
    Dim rs      As New ADODB.Recordset
    Dim strSQL  As String
    
    strSQL = " SELECT text1 FROM " & T_LAB034 & _
             " WHERE " & DBW("cdindex = ", LC4_Remark) & _
             " AND   " & DBW("cdval1  = ", pCode)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        SqlGetRemark = rs.Fields("text1").Value & ""
    Else
        SqlGetRemark = ""
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function
'==========================================================================================

'** 일건/종건 이외 결과수정시 업데이트
'Private Function SqlOCSMDEXMORT_Modify() As String
'
'End Function

'** 특수검사 관련 끝#################################################################################




'====================================================================================================
'** 추가루틴 : 전주예수병원 반응균 결과전송 By M.G.Choi 2005.03.02
'-- 미생물검사결과(Master) : False = Insert, True = Update
Public Function SqlSLMICWMT_FLAG(ByVal pAccNo As String) As Boolean
    Dim strSQL  As String
    Dim rs      As New ADODB.Recordset

    strSQL = " select * from " & F_OCSSLMICWMT & _
             "  where micrno = " & DBS(pAccNo)
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        SqlSLMICWMT_FLAG = True
    Else
        SqlSLMICWMT_FLAG = False
    End If
                               
    rs.Close
    Set rs = Nothing
    
End Function

Public Function SqlSLMICWMT_INSERT(ByVal pAccNo As String, ByVal pSpcCd As String, _
                                ByVal pPtid As String, ByVal pPtSect As String, _
                                ByVal pMedDt As String, ByVal pOrdDt As String, _
                                ByVal pDeptcd As String, ByVal pWardId As String, _
                                ByVal pRcvDt As String, ByVal pExecDt As String, _
                                ByVal pVfyDt As String, _
                                Optional ByVal pVfyId As String = "", Optional ByVal pMfyID As String = "") As String
                                
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strExecDt As String
    Dim strVfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 검사일자
    If pExecDt <> "" Then
        strExecDt = "TO_DATE(" & DBS(pExecDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strExecDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    
    
    SqlSLMICWMT_INSERT = " insert into " & F_OCSSLMICWMT & _
                         " (micrno, spccode, patno, patsect, meddate, orddate, meddept, wardno, " & _
                         "  acptdate, execdate, reptdate) " & _
                         " values( " & DBS(pAccNo) & "," & DBS(pSpcCd) & "," & DBS(pPtid) & _
                         "," & DBS(pPtSect) & "," & "TO_DATE(" & DBS(pMedDt) & ", 'yyyymmdd')" & _
                         "," & "TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & "," & DBS(pDeptcd) & _
                         "," & DBS(pWardId) & "," & strRcvDt & "," & strExecDt & "," & strVfyDt & ")"
                                    
                                    
End Function

Public Function SqlSLMICWMT_UPDATE(ByVal pAccNo As String, ByVal pSpcCd As String, _
                                ByVal pPtid As String, ByVal pPtSect As String, _
                                ByVal pMedDt As String, ByVal OrdDt As String, _
                                ByVal pDeptcd As String, ByVal WardId As String, _
                                ByVal pRcvDt As String, ByVal pExecDt As String, _
                                ByVal pVfyDt As String, _
                                Optional ByVal pVfyId As String = "", Optional ByVal pMfyID As String = "") As String
                                
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strExecDt As String
    Dim strVfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 검사일자
    If pExecDt <> "" Then
        strExecDt = "TO_DATE(" & DBS(pExecDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strExecDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    SqlSLMICWMT_UPDATE = " update " & F_OCSSLMICWMT & _
                         "    set acptdate = " & strRcvDt & _
                         ", execdate = " & strExecDt & _
                         ", reptdate = " & strVfyDt & _
                         "  where micrno = " & DBS(pAccNo)
                                     
End Function

'-- 미생물검사결과(Detail) : False = Insert, True = Update
Public Function SqlSLMICWDT_FLAG(ByVal pAccNo As String, ByVal pTestCd As String, _
                                ByVal pOrdSeq As String) As Boolean
    Dim strSQL  As String
    Dim rs      As New ADODB.Recordset

    strSQL = " select * from " & F_OCSSLMICWDT & _
             "  where micrno = " & DBS(pAccNo) & _
             "    and examcode = " & DBS(pTestCd) & _
             "    and ordseqno = " & DBN(pOrdSeq)
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        SqlSLMICWDT_FLAG = True
    Else
        SqlSLMICWDT_FLAG = False
    End If
                               
    rs.Close
    Set rs = Nothing
    
End Function

Public Function SqlSLMICWDT_INSERT(ByVal pAccNo As String, ByVal pTestCd As String, _
                                ByVal pOrdSeq As String, ByVal pRstText As String) As String
                                
    SqlSLMICWDT_INSERT = " insert into " & F_OCSSLMICWDT & _
                         " (micrno, examcode, ordseqno, rslttext) " & _
                         " values( " & DBS(pAccNo) & "," & DBS(pTestCd) & _
                         "," & DBS(pOrdSeq) & "," & DBS(pRstText) & ")"
                                                            
End Function

Public Function SqlSLMICWDT_UPDATE(ByVal pAccNo As String, ByVal pTestCd As String, _
                                ByVal pOrdSeq As String, ByVal pRstText As String) As String
                                
    SqlSLMICWDT_UPDATE = " update " & F_OCSSLMICWDT & _
                       "    set rslttext = " & DBS(pRstText) & _
                       "  where micrno = " & DBS(pAccNo) & _
                       "    and examcode = " & DBS(pTestCd) & _
                       "    and ordseqno = " & DBN(pOrdSeq)
                                    
End Function

'-- 미생물감수성결과 : False = Insert, True = Update
Public Function SqlSLXSENST_FLAG(ByVal pAccNo As String, ByVal pBactCd As String, _
                                ByVal pDrugCd As String) As Boolean
    Dim strSQL  As String
    Dim rs      As New ADODB.Recordset
    
    strSQL = " select * from " & F_OCSSLXSENST & _
             "  where micrno = " & DBS(pAccNo) & _
             "    and bactcode = " & DBS(pBactCd) & _
             "    and drugcode = " & DBN(pDrugCd)
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        SqlSLXSENST_FLAG = True
    Else
        SqlSLXSENST_FLAG = False
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function SqlSLXSENST_DELETE(ByVal pAccNo As String) As String
    SqlSLXSENST_DELETE = " delete from " & F_OCSSLXSENST & _
                         "  where micrno = " & DBS(pAccNo)
End Function

Public Function SqlSLXSENST_INSERT(ByVal pAccNo As String, ByVal pBactCd As String, _
                                ByVal pDrugCd As String, ByVal pSensType As String, _
                                ByVal pSensSign As String, ByVal pBactQty As String) As String
                                
    SqlSLXSENST_INSERT = " insert into " & F_OCSSLXSENST & _
                         " (micrno, bactcode, drugcode, senstype, senssign, bactqty) " & _
                         " values( " & DBS(pAccNo) & "," & DBS(pBactCd) & _
                         "," & DBS(pDrugCd) & "," & DBS(pSensType) & _
                         "," & DBS(pSensSign) & "," & DBS(pBactQty) & ")"
                         
    Debug.Print SqlSLXSENST_INSERT
                                                            
End Function

Public Function SqlSLXSENST_UPDATE(ByVal pAccNo As String, ByVal pBactCd As String, _
                                ByVal pDrugCd As String, ByVal pSensType As String, _
                                ByVal pSensSign As String, ByVal pBactQty As String) As String
                                
    SqlSLXSENST_UPDATE = " update " & F_OCSSLXSENST & _
                         "    set senstype = " & DBS(pSensType) & _
                         ", senssign = " & DBS(pSensSign) & _
                         ", bacctqty = " & DBS(pBactQty) & _
                         "  where micrno = " & DBS(pAccNo) & _
                         "    and bactcode = " & DBS(pBactCd) & _
                         "    and drugcode = " & DBN(pDrugCd)
                        
End Function

'-- 미생물감수성결과 : False = Insert, True = Update
Public Function SqlSLXBACTT_FLAG(ByVal pAccNo As String, ByVal pBactCd As String) As Boolean
    Dim strSQL  As String
    Dim rs      As New ADODB.Recordset
    
    strSQL = " select * from " & F_OCSSLXBACTT & _
             "  where micrno = " & DBS(pAccNo) & _
             "    and bactcode = " & DBS(pBactCd)
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        SqlSLXBACTT_FLAG = True
    Else
        SqlSLXBACTT_FLAG = False
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function SqlSLXBACTT_DELETE(ByVal pAccNo As String) As String
    SqlSLXBACTT_DELETE = " delete from " & F_OCSSLXBACTT & _
                         "  where micrno = " & DBS(pAccNo)
End Function

Public Function SqlSLXBACTT_INSERT(ByVal pAccNo As String, ByVal pBactCd As String) As String
                                
    SqlSLXBACTT_INSERT = " insert into " & F_OCSSLXBACTT & _
                         " (micrno, bactcode) " & _
                         " values( " & DBS(pAccNo) & "," & DBS(pBactCd) & ")"
                                                            
End Function

'Public Function SqlSLXBACTT_UPDATE(ByVal pAccNo As String, ByVal pBactCd As String) As String
'
'    SqlSLXBACTT_UPDATE = " update " & F_OCSSLXBACTT & _
'                       "    set senstype = " & DBS(pSensType) & _
'                       ", senssign = " & DBS(pSensSign) & _
'                       ", bacctqty = " & DBS(pBactQty) & _
'                       "  where micrno = " & DBS(pAccNo) & _
'                       "    and bactcode = " & DBS(pBactCd) & _
'                       "    and drugcode = " & DBN(pDrugCd)
'
'End Function

'** 미생물반응균 관련 끝#############################################################################

'** OCS 결과조회 수정 ###############################################################################
Public Function OCS_INPT_PArea(ByVal pPtid As String, ByVal pINDate As String) As String
    Dim strSQL  As String
    Dim rs      As New Recordset
    
    strSQL = " select wardno1, roomno1 from " & F_OCSPTAREA & _
             "  where patno = " & DBS(pPtid) & _
             "    and admdate = " & DBS(pINDate)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    On Error GoTo ErrMsg
    
    If rs.EOF = False Then
        OCS_INPT_PArea = rs.Fields("wardno1") & "" & "/" & rs.Fields("roomno1") & ""
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrMsg:
    MsgBox "환자 현재위치 오류", vbCritical, "오류확인"
    OCS_INPT_PArea = ""
             
End Function
'** OCS 결과조회 수정 끝#############################################################################


'%  현재 작업중인 로컬 피씨의 IP Address를 확인한다.
Public Function LocalIP_Address() As String
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim Listing As MIB_IPADDRTABLE

    On Error GoTo ErrMsg
    
    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Function
    
    ReDim bBytes(0 To Ret - 1) As Byte
    
    GetIpAddrTable bBytes(0), Ret, False
    
    CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
    
    LocalIP_Address = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
    
    Exit Function
    
ErrMsg:
    MsgBox "IP Address를 가져올 수 없습니다.", vbCritical
    
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    CopyMemory myByte(0), longAddr, 4
    
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    
End Function

Private Sub Class_Initialize()
    mvarPtID = ""
    mvarOrdDt = ""
    mvarOrdNo = ""
    mvarRstVal = ""
    mvarRstText = ""
    mvarRefValFrom = ""
    mvarRefValTo = ""
    mvarRstUnit = ""
    mvarFootNote = ""
    mvarVfyDt = ""
    mvarVfyTm = ""
    mvarVfyID = ""
    mvarMfyID = ""
    mvarIPAdr = ""
    mvarMfyDt = ""
    
    mvarSlipCd = ""
    mvarRsltType = ""
    mvarWrkDiv = ""
End Sub
