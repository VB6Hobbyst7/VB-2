VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlMicRst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------
'1. 미생물 Work List 작성 관련 Sql
'----------------------------------------------------------------------------------

Public Function SQLGetSpcGroup() As String

    'field5 -예외구분:   'W'이면 여성클리닉 Batch GS 이므로 제외시킨다. 2000/4/11 by 김미경
    SQLGetSpcGroup = " SELECT cdval1,field1,field2,field3,field4,field5 FROM " & T_LAB032 & " " & _
                     " WHERE " & DBW("cdindex = ", LC3_SGroup) & _
                     " Order  by field3 asc"
'" AND   (field5 is null  OR  field5 = '')" & _

End Function

Public Function SQLLastWsUnit(ByVal sWsCd As String) As String
        
    SQLLastWsUnit = " SELECT wsunit,fnshdt,fnshtm FROM " & T_LAB401 & " " & _
                    " WHERE " & DBW("wscd  = ", sWsCd) & _
                    " AND    wsunit=(SELECT max(wsunit) FROM " & T_LAB401 & _
                                    " WHERE " & DBW("wscd = ", sWsCd) & ")"

End Function

Public Function SQLGetMedias(ByVal pSGCD As String) As String
    
    SQLGetMedias = " SELECT cdval2 as media FROM " & T_LAB031 & _
                   " WHERE " & DBW("cdindex = ", LC2_SpcMedia) & _
                   " AND   " & DBW("cdval1  = ", pSGCD)

End Function

Public Function SQLGetRstType(ByVal pSG As String) As String
    
    SQLGetRstType = " SELECT field2 FROM " & T_LAB032 & _
                    " WHERE " & DBW("cdindex = ", LC3_MWSKinds) & _
                    " AND   " & DBW("cdval1  = ", pSG)

End Function

Public Function SQLGetWorkList(ByVal sWACD As String, ByVal sSR1 As String, ByVal sSR2 As String, _
                               ByVal sFR As String, ByVal sTO As String, ByVal pRTs As String) As String
    '** 속도를 위해 접수일자/일시 조회조건에서 불리 함 By M.G.Choi 2006.12.21
    Dim strFRcvDt       As String
    Dim strFRcvTm       As String
    Dim strTRcvDt       As String
    Dim strTRcvTm       As String
    
    If Len(sFR) = 14 Then
        strFRcvDt = Mid(sFR, 1, 8)
        strFRcvTm = Mid(sFR, 9, 6)
    Else
        strFRcvDt = Mid(sFR, 1, 8)
        strFRcvTm = Format(Mid(sFR, 9, 6), "0#####")
    End If
    
    If Len(sTO) = 14 Then
        strTRcvDt = Mid(sTO, 1, 8)
        strTRcvTm = Mid(sTO, 9, 6)
    Else
        strTRcvDt = Mid(sTO, 1, 8)
        strTRcvTm = Format(Mid(sTO, 9, 6), "0#####")
    End If
'    '=========================================================================
    
    '** 속도를 위해 접수일자/일시 조회조건에서 불리 함 By M.G.Choi 2006.12.21
    '   원본에 RCVDT 를 한 번 더 강제 조회토록 한다.
    ' 2008.10.23. 양성현 바코드 출력을 위한 조건을 삽입
    SQLGetWorkList = " SELECT  a.workarea,a.accdt,a.accseq,a.sex,a.ageday,a.rcvdt,a.rcvtm,a.spccd,a.ptid,a.wardid," & _
                     "         f." & F_PTNM & " ptnm, a.spcyy || trim( to_char(a.spcno, '000000000') ) as barcode ,c.rsttype, c.testcd, c.orddt, c.ordno, c.ordseq, " & _
                     "         d.testnm,d.abbrnm10,e.field3 as spcnm " & _
                     " FROM  " & T_HIS001 & " f, " & T_LAB001 & " d, " & T_LAB032 & " e, " & _
                                 T_LAB404 & " c, " & T_LAB201 & " a " & _
                     " WHERE " & DBW("a.workarea = ", sWACD) & _
                     " AND     a.accseq between " & DBV("accseq", sSR1) & " AND " & DBV("accseq", sSR2) & _
                     " AND     a.rcvdt between " & DBS(strFRcvDt) & " and " & DBS(strTRcvDt) & _
                     " AND     a.rcvdt" & FUNC_CONCAT & "a.rcvtm between " & DBS(sFR) & " AND " & DBS(sTO) & _
                     " AND   " & DBJ("a.ptid *= f." & F_PTID) & _
                     " AND     a.workarea = c.workarea AND a.accdt = c.accdt AND a.accseq = c.accseq " & _
                     " AND     c.stscd = " & DBV("stscd", enStsCd.StsCd_LIS_Accession) & _
                     " AND     c.rsttype in (" & pRTs & ") " & _
                     " AND     c.testcd = d.testcd " & _
                     " AND    (c.rstdiv = '*' OR c.detailfg = '' OR c.detailfg is null ) " & _
                     " AND     d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " e " & _
                     "                      WHERE e.testcd = d.testcd AND e.applydt <= " & DBV("applydt", Format(Now, CS_DateDbFormat)) & ")" & _
                     " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                     " AND     e.cdval1  = a.spccd " & _
                     " ORDER   BY a.rcvdt, a.rcvtm "
                     
    
    '** 원본 =================================================================================================================================
    '** 여성 클리닉 GS 는 제외한다. - 진료과:OB, 검체:75, 검사항목:B400     2000/4/11 by 김미경
'    SQLGetWorkList = " SELECT  a.workarea,a.accdt,a.accseq,a.sex,a.ageday,a.rcvdt,a.rcvtm,a.spccd,a.ptid,a.wardid," & _
                     "         f." & F_PTNM & " ptnm ,c.rsttype, c.testcd, c.orddt, c.ordno, c.ordseq, " & _
                     "         d.testnm,d.abbrnm10,e.field3 as spcnm " & _
                     " FROM  " & T_HIS001 & " f, " & T_LAB001 & " d, " & T_LAB032 & " e, " & _
                                 T_LAB404 & " c, " & T_LAB201 & " a " & _
                     " WHERE " & DBW("a.workarea = ", sWACD) & _
                     " AND     a.accseq between " & DBV("accseq", sSR1) & " AND " & DBV("accseq", sSR2) & _
                     " AND     a.rcvdt" & FUNC_CONCAT & "a.rcvtm between " & DBS(sFR) & " AND " & DBS(sTO) & _
                     " AND   " & DBJ("a.ptid *= f." & F_PTID) & _
                     " AND     a.workarea = c.workarea AND a.accdt = c.accdt AND a.accseq = c.accseq " & _
                     " AND     c.stscd = " & DBV("stscd", enStsCd.StsCd_LIS_Accession) & _
                     " AND     c.rsttype in (" & pRTs & ") " & _
                     " AND     c.testcd = d.testcd " & _
                     " AND    (c.rstdiv = '*' OR c.detailfg = '' OR c.detailfg is null ) " & _
                     " AND     d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " e " & _
                     "                      WHERE e.testcd = d.testcd AND e.applydt <= " & DBV("applydt", Format(Now, CS_DateDbFormat)) & ")" & _
                     " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                     " AND     e.cdval1  = a.spccd " & _
                     " ORDER   BY a.rcvdt, a.rcvtm "    '** 접수시간순.. 2000/2/17 by 김미경&정미경
                     '<JMK>  OR a.deptcd <> 'CO' 삽입 (2000-06-19)
                     '" AND    (a.deptcd <> 'OB'  OR a.deptcd <> 'CO' OR  a.spccd <> '75'  OR  c.testcd <> 'B400') " & _
    '==========================================================================================================================================

End Function

Public Function SQLGetExcpList(ByVal sWsCd As String, ByVal sWsUn As String) As String
    
    ' 이전 작업에서 제외되었던 검체 로드
    ' 2008.10.23. 양성현 바코드 출력을 위한 조건을 삽입
    
    SQLGetExcpList = " SELECT b.workarea,b.accdt,b.accseq,a.sex,a.ageday,a.rcvdt,a.rcvtm,a.spccd,a.wardid," & _
                     "        a.ptid,f." & F_PTNM & " ptnm, a.spcyy || trim( to_char(a.spcno, '000000000') ) as barcode,c.rsttype,c.testcd,c.orddt,c.ordno,c.ordseq," & _
                     "        d.testnm,d.abbrnm10,e.field3 as spcnm " & _
                     " FROM " & T_LAB201 & " a, " & T_LAB403 & " b, " & T_HIS001 & " f, " & _
                                T_LAB404 & " c, " & T_LAB001 & " d, " & T_LAB032 & " e  " & _
                     " WHERE " & DBW("b.wscd   = ", sWsCd) & _
                     " AND   " & DBW("b.wsunit = ", sWsUn) & _
                     " AND   a.workarea = b.workarea AND a.accdt = b.accdt AND a.accseq = b.accseq " & _
                     " AND   " & DBJ("a.ptid  *= f." & F_PTID) & _
                     " AND   a.workarea = c.workarea AND a.accdt=c.accdt AND a.accseq=c.accseq " & _
                     " AND   c.testcd   = d.testcd " & _
                     " AND  (c.rstdiv   = '*' OR (c.detailfg  =  '' or c.detailfg is null) OR  a.multifg = 'Y') " & _
                     " AND   d.applydt  = (SELECT max(applydt) FROM " & T_LAB001 & " e " & _
                     "                    WHERE e.testcd = d.testcd AND e.applydt <= " & DBS(Format(Now, CS_DateDbFormat)) & ")" & _
                     " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                     " AND   e.cdval1  = a.spccd " & _
                     " ORDER BY a.rcvdt, a.rcvtm "    '** 접수시간순.. 2000/2/17 by 김미경&정미경

End Function

Public Function SQLGetWsUnitNo(ByVal sWA As String, ByVal sSG As String, ByVal sYY As String) As String


    SQLGetWsUnitNo = " SELECT seq FROM " & T_COM099 & " " & _
                     " WHERE " & DBW("noindex = ", COM99_LIS_WSUnit) & _
                     " AND   " & DBW("divcd1  = ", sWA) & _
                     " AND   " & DBW("divcd2  = ", sSG) & _
                     " AND   " & DBW("divcd3  = ", sYY)

End Function

'----------------------------------------------------------------------------------
'2. 미생물 Nogrowth 결과등록 관련 Sql
'----------------------------------------------------------------------------------

Public Function SQLGetNGRstCd(ByVal pCode As String) As String
    'pCode : 미생물 Worksheet코드 - 검체군정보의 field4에 정의
    
'    sqlSG = "SELECT a.* FROM " & T_LAB032 & " a," & T_LAB032 & " b " & _
'            " WHERE a.cdindex='" & LC3_SGroup & "' AND b.cdindex='" & LC3_MWSKinds & "'" & _
'            "   AND b.field2 in ('" & MRT_GSen & "," & MRT_MSen & "','" & MRT_AFC & "','" & MRT_AFS & "') " & _
'            "   AND a.field4=b.cdval1 " & _
'            " ORDER BY a.field3,a.cdval1 asc"
     SQLGetNGRstCd = " SELECT  cdval2, field1 FROM " & T_LAB031 & _
                    " WHERE " & DBW("cdindex=", LC2_MBatchRst) & _
                    " AND   " & DBW("cdval1  = ", pCode)
'                    " AND       cdval1 = (SELECT field4 FROM " & T_LAB032 & _
                                        " WHERE " & DBW("cdindex = ", LC3_SGroup) & _
                                        " AND   " & DBW("cdval1  = ", pCode) & ")"
    
End Function

Public Function SQLGetWScode(ByVal pCase As Long) As String
    
    SQLGetWScode = " SELECT a.cdval1 as spcgrp, a.field1 as spcgrpnm, " & _
                   "        b.cdval1 as wstype, b.field1 as wsname, b.field2 as rsttype " & _
                           " FROM  " & T_LAB032 & " a," & T_LAB032 & " b " & _
                           " WHERE " & DBW("a.cdindex = ", LC3_SGroup) & _
                           " AND   " & DBW("b.cdindex = ", LC3_MWSKinds)
    Select Case pCase
        Case 0
            'ALL
            SQLGetWScode = SQLGetWScode & _
                           " ORDER BY a.field3 "
        Case 1
            'Nogrowth 배치등록...
            SQLGetWScode = SQLGetWScode & _
                           " AND   b.field2 in ('" & MRT_GenSen & "," & _
                                                     MRT_MicSen & "')" & _
                           " AND   a.field4=b.cdval1 " & _
                           " ORDER BY a.field3,a.cdval1 asc"
                                                     
'                                                     DBV("field2", MRT_AFN) & "," & _
'                                                     DBV("field2", MRT_AFS) & ","
        Case 2
            'Stain 결과등록...
            SQLGetWScode = SQLGetWScode & _
                           " AND   b.field2 in (" & DBV("field2", MRT_Stain) & "," & _
                                                    DBV("field2", MRT_AFS) & "," & _
                                                    DBV("field2", MRT_FUS) & "," & _
                                                    DBV("field2", MRT_AFC) & "," & _
                                                    DBV("field2", MRT_FUC) & ")" & _
                           " AND   a.field4=b.cdval1 " & _
                           " ORDER BY a.cdval1 "
                           
        Case 3
            '감수성 결과등록...
            SQLGetWScode = SQLGetWScode & _
                           " AND   b.field2 in ('" & MRT_GenSen & "," & _
                                                     MRT_MicSen & "'" & " ) " & _
                           " AND   a.field4=b.cdval1 " & _
                           " ORDER BY a.field3 "
    End Select
    
End Function

Public Function SQLBarcodeReprint(ByVal pWsCode As String, ByVal pWsUnit As String, _
                    Optional ByVal pRTs As String = "")
    SQLBarcodeReprint = " SELECT distinct b.workarea,b.accdt,b.accseq,a.ptid,f." & F_PTNM & " ptnm," & F_DOB2("f") & " dob," & _
                        "        a.sex,a.ageday,a.coldt,a.coltm,a.rcvdt,a.rcvtm,a.spccd,d.field1 rst," & _
                        "        e.field3 as SpcNm" & _
                        " FROM  " & T_LAB201 & " a, " & T_HIS001 & " f, " & T_LAB404 & " c, " & _
                                    T_LAB031 & " d, " & T_LAB032 & " e, " & T_LAB402 & " b " & _
                        " WHERE " & DBW("b.wscd   = ", pWsCode) & _
                        " AND   " & DBW("b.wsunit = ", pWsUnit) & _
                        " AND   " & DBW("b.genfg  = ", MWS_Ready) & _
                        " AND   " & DBW("b.selfg  <> ", MWS_Final) & _
                        " AND     a.workarea = b.workarea AND a.accdt = b.accdt AND a.accseq = b.accseq " & _
                        " AND     c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                        " AND    (c.senfg ='' or c.senfg is null) " & _
                        " AND     f." & F_PTID & " = a.ptid" & _
                        " AND   " & DBJ(DBW("d.cdindex =", LC2_ItemResult)) & _
                        " AND   " & DBJ("d.cdval1  =* c.testcd") & _
                        " AND   " & DBJ("d.cdval2  =* c.rstcd") & _
                        " AND     c.rsttype in (" & pRTs & ") " & _
                        " AND   " & DBW("c.stscd   < ", enStsCd.StsCd_LIS_FinRst) & _
                        " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                        " AND     e.cdval1 = a.spccd " & _
                        " ORDER BY workarea, accdt, accseq"
End Function

Public Function SQLBarcodeReprint1(ByVal pStart As String, ByVal pEnd As String, _
                    ByVal pWork As String, ByVal pTestCd As String)
'    SQLBarcodeReprint1 = " SELECT distinct b.workarea,b.accdt,b.accseq,a.ptid,f." & F_PTNM & " ptnm," & F_DOB2("f") & " dob," & _
'                        "        a.sex,a.ageday,a.coldt,a.coltm,a.rcvdt,a.rcvtm,a.spccd,d.field1 rst," & _
'                        "        e.field3 as SpcNm" & _
'                        " FROM  " & T_LAB201 & " a, " & T_HIS001 & " f, " & T_LAB404 & " c, " & _
'                                    T_LAB031 & " d, " & T_LAB032 & " e, " & T_LAB402 & " b " & _
'                        " WHERE " & DBW("b.wscd   = ", pWsCode) & _
'                        " AND   " & DBW("b.wsunit = ", pWsUnit) & _
'                        " AND   " & DBW("b.genfg  = ", MWS_Ready) & _
'                        " AND   " & DBW("b.selfg  <> ", MWS_Final) & _
'                        " AND     a.workarea = b.workarea AND a.accdt = b.accdt AND a.accseq = b.accseq " & _
'                        " AND     c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
'                        " AND    (c.senfg ='' or c.senfg is null) " & _
'                        " AND     f." & F_PTID & " = a.ptid" & _
'                        " AND   " & DBJ(DBW("d.cdindex =", LC2_ItemResult)) & _
'                        " AND   " & DBJ("d.cdval1  =* c.testcd") & _
'                        " AND   " & DBJ("d.cdval2  =* c.rstcd") & _
'                        " AND     c.rsttype in (" & pRTs & ") " & _
'                        " AND   " & DBW("c.stscd   < ", enStsCd.StsCd_LIS_FinRst) & _
'                        " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
'                        " AND     e.cdval1 = a.spccd " & _
'                        " ORDER BY workarea, accdt, accseq"


    SQLBarcodeReprint1 = "SELECT DISTINCT a.workarea,                            "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.accdt,                                        "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.accseq,                                       "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.ptid,                                         "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       f.patname ptnm,                                 "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       to_char(f.birtdate, 'yyyymmdd') dob,            "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.sex,                                          "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.ageday,                                       "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.coldt,                                        "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.coltm,                                        "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.rcvdt,                                        "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.rcvtm,                                        "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       a.spccd,                                        "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       d.field1 rst,                                   "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       e.field3 AS SpcNm                               "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "  FROM s2lab201 a,                                     "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       ORAA1.APPATBAT f,                               "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       s2lab302 c,                                     "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       s2lab031 d,                                     "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       s2lab032 e                                      "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + " WHERE a.WORKAREA = '" & pWork & "'                    "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND a.ACCDT BETWEEN  '" & pStart & "'  AND '" & pEnd & "'"
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND c.TESTCD = '" & pTestCd & "'"
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND a.workarea = c.workarea                     "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND a.accdt = c.accdt                           "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND a.accseq = c.accseq                         "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND c.workarea = a.workarea                     "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND c.accdt = a.accdt                           "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND c.accseq = a.accseq                         "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND f.patno = a.ptid                            "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND d.cdindex(+) = 'C110'                       "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND d.cdval1(+) = c.testcd                      "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND d.cdval2(+) = c.rstcd                       "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND a.stscd < '5'                               "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND e.cdindex = 'C215'                          "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       AND e.cdval1 = a.spccd                          "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "ORDER BY workarea,                                     "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       accdt,                                          "
    SQLBarcodeReprint1 = SQLBarcodeReprint1 + "       accseq                                          "

End Function

Public Function SQLReadBatchWorkList(ByVal pCaseFg As Long, ByVal pWsCode As String, _
                                     ByVal pWsUnit As String, Optional ByVal pRTs As String = "") As String
    '튜닝
    Select Case pCaseFg
        Case 1: 'Nogrowth
            SQLReadBatchWorkList = " SELECT distinct b.workarea,b.accdt,b.accseq,a.ptid,f." & F_PTNM & " ptnm," & F_DOB2("f") & " dob," & _
" a.spcyy || trim( to_char(a.spcno, '000000000') ) as barcode , " & _
                                   "        a.sex,a.ageday,a.coldt,a.coltm,a.rcvdt,a.rcvtm,a.spccd,d.field1 rst," & _
                                   "        e.field3 as SpcNm, c.testcd, c.rsttype " & _
                                   " FROM  " & T_LAB201 & " a, " & T_HIS001 & " f, " & T_LAB404 & " c, " & _
                                               T_LAB031 & " d, " & T_LAB032 & " e, " & T_LAB402 & " b " & _
                                   " WHERE " & DBW("b.wscd   = ", pWsCode) & _
                                   " AND   " & DBW("b.wsunit = ", pWsUnit) & _
                                   " AND   " & DBW("b.genfg  = ", MWS_Ready) & _
                                   " AND   " & DBW("b.selfg  <> ", MWS_Final) & _
                                   " AND     a.workarea = b.workarea AND a.accdt = b.accdt AND a.accseq = b.accseq " & _
                                   " AND     c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                                   " AND    (c.senfg ='' or c.senfg is null) " & _
                                   " AND     f." & F_PTID & " = a.ptid" & _
                                   " AND   " & DBJ(DBW("d.cdindex =", LC2_ItemResult)) & _
                                   " AND   " & DBJ("d.cdval1  =* c.testcd") & _
                                   " AND   " & DBJ("d.cdval2  =* c.rstcd") & _
                                   " AND     c.rsttype in (" & pRTs & ") " & _
                                   " AND   " & DBW("c.stscd   < ", enStsCd.StsCd_LIS_FinRst) & _
                                   " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                                   " AND     e.cdval1 = a.spccd " & _
                                   " ORDER BY workarea, accdt, accseq"
        Case 2: 'Growth
            ' Work Sheet (Growth 결과)
            SQLReadBatchWorkList = " SELECT * FROM " & T_LAB402 & _
                                   " WHERE  " & DBW("wscd   = ", pWsCode) & _
                                   " AND    " & DBW("wsunit = ", pWsUnit) & _
                                   " AND    selfg in (" & DBV("selfg", MWS_Holding) & " , " & _
                                                          DBV("selfg", MWS_Growth) & ") " & _
                                   " AND    " & DBW("genfg <> ", MWS_Final) & " AND    " & DBW("selfg <> ", MWS_Final) & _
                                   " ORDER BY workarea, accdt, accseq"
        Case 3: 'Final
            ' Work Sheet (Growth 결과)
            SQLReadBatchWorkList = " SELECT * FROM " & T_LAB402 & _
                                   " WHERE  " & DBW("wscd   = ", pWsCode) & _
                                   " AND    " & DBW("wsunit = ", pWsUnit) & _
                                   " AND   ( " & DBW("selfg  = ", MWS_Final) & _
                                   " OR      " & DBW("genfg  = ", MWS_Final) & ")" & _
                                   " ORDER BY workarea, accdt, accseq"

        Case 4: 'Holding
            SQLReadBatchWorkList = " SELECT distinct b.workarea,b.accdt,b.accseq,a.ptid,f." & F_PTNM & " ptnm," & F_DOB2("f") & " dob," & _
" a.spcyy || trim( to_char(a.spcno, '000000000') ) as barcode , " & _
                                   "        a.sex,a.ageday,a.coldt,a.coltm,a.rcvdt,a.rcvtm,a.spccd," & _
                                   "        d.field1 rst, e.field3 as SpcNm, c.testcd, c.rsttype  " & _
                                   " FROM  " & T_LAB201 & " a, " & T_HIS001 & " f, " & T_LAB404 & " c, " & _
                                               T_LAB031 & " d, " & T_LAB032 & " e, " & T_LAB402 & " b " & _
                                   " WHERE " & DBW("b.wscd   = ", pWsCode) & _
                                   " AND   " & DBW("b.wsunit = ", pWsUnit) & _
                                   " AND   " & DBW("b.genfg  = ", MWS_Holding) & _
                                   " AND   " & DBW("b.selfg < ", MWS_Growth) & _
                                   " AND     a.workarea = b.workarea AND a.accdt = b.accdt AND a.accseq = b.accseq " & _
                                   " AND     c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                                   " AND     c.rsttype in (" & pRTs & ") " & _
                                   " AND   " & DBJ("f." & F_PTID & " =* a.ptid") & _
                                   " AND   " & DBJ(DBW("d.cdindex = ", LC2_MBatchRst)) & _
                                   " AND   " & DBJ("d.cdval2  =* c.rstcd") & _
                                   " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                                   " AND     e.cdval1 = a.spccd " & _
                                   " ORDER BY workarea, accdt, accseq"
                                   

        Case 5: 'Growth & Holding For 감수성결과등록
            ' Work Sheet (Growth 결과)
            
            SQLReadBatchWorkList = " SELECT * FROM " & T_LAB402 & _
                                   " WHERE  " & DBW("wscd   = ", pWsCode) & _
                                   " AND    " & DBW("wsunit = ", pWsUnit) & _
                                   " AND    " & DBW("genfg  = ", MWS_Holding) & _
                                   " AND    " & DBW("selfg <> ", MWS_Final) & _
                                   " ORDER BY workarea, accdt, accseq"
    End Select

End Function

Public Function SQLGetAddSpcInfo(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pSpcGrp As String, ByVal pRTs As String) As String
    
    SQLGetAddSpcInfo = " SELECT distinct a.workarea,a.accdt,a.accseq,a.ptid,f." & F_PTNM & " as ptnm,a.sex,a.ageday,a.rcvdt,a.spccd," & _
                       "        e.field3 as SpcNm, c.testcd, c.rsttype, d.testnm,c.rstcd " & _
                       " FROM  " & T_LAB201 & " a, " & T_HIS001 & " f, " & _
                                   T_LAB404 & " c, " & T_LAB001 & " d, " & T_LAB032 & " e " & _
                       " WHERE " & DBW("a.workarea   = ", pWorkArea) & _
                       " AND   " & DBW("a.accdt   = ", pAccDt) & _
                       " AND   " & DBW("a.accseq  = ", pAccSeq) & _
                       " AND     c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                       " AND    (c.senfg ='' or c.senfg is null) " & _
                       " AND     f." & F_PTID & " = a.ptid" & _
                       " AND     c.rsttype in (" & pRTs & ") " & _
                       " AND   " & DBW("c.stscd   < ", enStsCd.StsCd_LIS_FinRst) & _
                       " AND   " & DBW("e.cdindex = ", LC3_Specimen) & _
                       " AND     e.cdval1 = a.spccd " & _
                       " AND     d.testcd = c.testcd "

End Function

Public Function SQLSaveAddSpc(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pSpcGrp As String, _
                              ByVal pWsUnit As String, ByVal pRTs As String) As String
    
    SQLSaveAddSpc = " INSERT INTO " & T_LAB402 & " (wscd, wsunit, workarea, accdt, accseq, selfg, genfg, scfg) " & _
                    " VALUES ( " & DBV("wscd", pSpcGrp, 1) & DBV("wsunit", pWsUnit, 1) & _
                                    DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & _
                                    DBV("accseq", pAccSeq, 1) & DBV("selfg", MWS_Ready) & _
                                    DBV("genfg", MWS_Ready, 1) & DBV("scfg", pRTs) & ")"
        
        
End Function

Public Function SQLDeleteSpc(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pSpcGrp As String, _
                             ByVal pWsUnit As String) As String
    
    SQLDeleteSpc = " DELETE " & T_LAB402 & _
                   " WHERE  " & DBW("wscd     = ", pSpcGrp) & _
                   " AND    " & DBW("wsunit   = ", pWsUnit) & _
                   " AND    " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq)

        
End Function


Public Function SQLSaveNogrowthBatch(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                     ByVal pTestCds As String, ByVal pResult As String, ByVal pStatus As String, _
                                     ByVal pDate As String, ByVal pTime As String, ByVal pEmpId As String, _
                                     Optional ByVal pSenFg As String = "") As String
    
    SQLSaveNogrowthBatch = " UPDATE " & T_LAB404 & _
                           " SET    " & DBW("rstcd    = ", pResult) & ", " & _
                           "        " & DBW("stscd    = ", pStatus) & ", " & _
                           "        " & DBW("vfydt    = ", pDate) & ", " & _
                           "        " & DBW("vfytm    = ", pTime) & ", " & _
                           "        " & DBW("vfyid    = ", pEmpId) & " " & _
                           " WHERE  " & DBW("workarea = ", pWorkArea) & _
                           " AND    " & DBW("accdt    = ", pAccDt) & _
                           " AND    " & DBW("accseq   = ", pAccSeq) & _
                           " AND    " & DBW("stscd    < ", enStsCd.StsCd_LIS_FinRst) & _
                           " AND    testcd in (" & pTestCds & ") "
                           
'                           " AND    rsttype <> 'G' "

End Function


'----------------------------------------------------------------------------------
'3. 미생물 Stain 결과등록 관련 Sql
'----------------------------------------------------------------------------------

Public Function SQLLoadRemark() As String

    SQLLoadRemark = " SELECT * FROM " & T_LAB034 & _
                    " WHERE " & DBW("cdindex = ", LC4_Remark) & _
                    " ORDER BY cdval1 "

End Function

Public Function SqlGetRemark(ByVal pCode As String) As String
    
    SqlGetRemark = " SELECT text1 FROM " & T_LAB034 & _
                   " WHERE " & DBW("cdindex = ", LC4_Remark) & _
                   " AND   " & DBW("cdval1  = ", pCode)

End Function

Public Function SQLReadWsUnit(ByVal pWsCd As String, Optional ByVal pMfySeq As Boolean = False) As String
    '** 원본 =======================================================================================================
    '## 5.0.20: 이상대(2005-02-11)
    '   - 미생물 연속검사의 경우 마지막 접수번호만 102에 저장되어 102를 접수번호로 조인하면
    '     나머지 접수번호를 표시되지 않어서 102를 PTID, ORDDT, ORDNO, ORDSEQ로 조인하도록 수정
'    SQLReadWsUnit = " SELECT a.wsunit,a.workdt,a.worktm FROM " & T_LAB401 & " a " & _
'                    " WHERE  " & DBW("a.wscd = ", pWsCd) & _
'                    " AND    EXISTS ( SELECT * FROM " & T_LAB102 & " d," & T_LAB404 & " b, " & T_LAB402 & " c " & _
'                                    " WHERE c.wscd     = a.wscd      AND  c.wsunit  = a.wsunit " & _
'                                    " AND   b.workarea = c.workarea  AND  b.accdt   = c.accdt " & _
'                                    " AND   b.accseq   = c.accseq    AND  b.rsttype = c.scfg " & _
'                                    " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
'                                    " AND b.ordseq=d.ordseq AND (d.dcfg='' or d.dcfg is null) "
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("c.genfg >= ", MWS_Final) & _
'                        " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst) & ") "
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("c.genfg < ", MWS_Final) & _
'                        " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) & ") "
'    End If
'    SQLReadWsUnit = SQLReadWsUnit & " Order by a.wsunit desc"
    '===============================================================================================================
    
    '** 변경 By M.G.Choi 2006.12.27 (2단계로 나눔)
    '** 추가 By S.W,Oh   2007.03.27 (쿼리 추가 : "   AND b.workarea is not null   AND  b.accdt is not null AND b.accseq is not null AND b.testcd is not null" & )
    
    
'    SQLReadWsUnit = " SELECT /*+ RULE */ distinct a.wscd, a.wsunit,a.workdt,a.worktm " & _
'                    "   FROM " & T_LAB401 & " a, " & T_LAB402 & " c " & _
'                    " WHERE  " & DBW("a.wscd = ", pWsCd) & _
'                    " AND    c.wscd     = a.wscd      AND  c.wsunit  = a.wsunit " & _
'                    " AND    substr(a.workdt,1,6) >= to_char(TRUNC(ADD_MONTHS(SYSDATE,-3)+1)-TO_CHAR(SYSDATE,'DD'),'YYYYMM') " & _
'                    " AND    EXISTS ( SELECT testcd FROM " & T_LAB404 & " b " & _
'                                    " WHERE b.workarea = c.workarea  AND  b.accdt   = c.accdt " & _
'                                    "   AND b.accseq   = c.accseq    AND  b.rsttype = c.scfg "
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst) & ") " & _
'                        " AND " & DBW("c.genfg >= ", MWS_Final)
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) & ") " & _
'                        " AND " & DBW("c.genfg < ", MWS_Final)
'    End If
'    SQLReadWsUnit = SQLReadWsUnit & " Order by a.wsunit desc"
'
'----------- 2008.10.22 양성현 미생물 결과 및 수정을 위한 조회를 빠르게 하기 위해 인덱스 힌트를 둠.

'    SQLReadWsUnit = " SELECT /*+ INDEX_DESC(a, s2lab401_pk */ /*+ INDEX_DESC(c, s2lab402_pk */ /*+ INDEX_DESC(b, s2lab404_pk */ distinct a.wscd, a.wsunit,a.workdt,a.worktm " & _
'                    "   FROM " & T_LAB401 & " a, " & T_LAB402 & " c join " & T_LAB404 & _
'                    " b on ( b.workarea = c.workarea AND b.accdt  = c.accdt  AND b.accseq  = c.accseq  AND b.rsttype = c.scfg "
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst) & ") "
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) & ") "
'    End If
'
'        SQLReadWsUnit = SQLReadWsUnit & _
'                    " WHERE  " & DBW("a.wscd = ", pWsCd) & _
'                    " AND    c.wscd     = a.wscd      AND  c.wsunit  = a.wsunit " & _
'                    " AND    substr(a.workdt,1,6) >= to_char(TRUNC(ADD_MONTHS(SYSDATE,-6)+1)-TO_CHAR(SYSDATE,'DD'),'YYYYMM') "
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg >= ", MWS_Final)
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg < ", MWS_Final)
'    End If
'    SQLReadWsUnit = SQLReadWsUnit & " Order by a.wsunit desc"
'
' " SELECT /*+ Row(a) +*/ /*+ INDEX_DESC(a, s2lab401_idx5) */ distinct a.wscd, a.wsunit,a.workdt,a.worktm, c.workarea,c.accdt,c.accseq,c.scfg "
'
'    SQLReadWsUnit = " select  /*+ Row(p) +*//*+  AND_EQUALS(s2lab404_pk, s2lab404_idx5) */  distinct p.wscd, p.wsunit,p.workdt,p.worktm from ( " & _
'        " SELECT /*+ Row(a) +*//*+  AND_EQUALS(s2lab401_pk, s2lab401_idx5) */ distinct a.wscd, a.wsunit,a.workdt,a.worktm, c.workarea,c.accdt,c.accseq,c.scfg " & _
'        " FROM " & T_LAB401 & " a, " & T_LAB402 & " c " & _
'        " WHERE " & DBW("a.wscd = ", pWsCd) & " AND c.wscd = a.wscd  AND  c.wsunit  = a.wsunit "
''                   -----------------------
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg >= ", MWS_Final)
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg < ", MWS_Final)
'    End If
''   SQLReadWsUnit = SQLReadWsUnit & " AND substr(a.workdt,1,6) >= (select to_char(ADD_MONTHS(SYSDATE,-6) ,'YYYYMM') || to_char(SYSDATE,'DD') from dual )"
'    SQLReadWsUnit = SQLReadWsUnit & " AND a.WORKDT >= to_char(SYSDATE,'YYYYMM') -6|| to_char(SYSDATE,'DD') " & _
'        " ) p, " & T_LAB404 & " b " & _
'        " where   b.workarea = p.workarea AND b.accdt  = p.accdt  AND b.accseq  = p.accseq  AND b.rsttype = p.scfg "
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst)  '& ") "
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) ' & ") "
'    End If
    
    '## 온승호(2010.03.31)
    '   - 쿼리 속도 수정
    '   - 조회 일정 6개월에서 3개월
    '   - 인덱스 수정
    
    Dim strTmpDt As String

    strTmpDt = Format(DateAdd("m", -1, Now), "yyyymmdd")

    SQLReadWsUnit = " select /*+ Row(p) +*//*+  AND_EQUALS(s2lab404_pk, s2lab404_idx5) */  distinct p.wscd, p.wsunit,p.workdt,p.worktm from ( " & _
        " SELECT /*+ Row(a) +*//*+  AND_EQUALS(s2lab401_pk, s2lab401_idx5) */ distinct a.wscd, a.wsunit,a.workdt,a.worktm, c.workarea,c.accdt,c.accseq,c.scfg " & _
        " FROM " & T_LAB401 & " a, " & T_LAB402 & " c " & _
        " WHERE " & DBW("a.wscd = ", pWsCd) & " AND c.wscd = a.wscd  AND  c.wsunit  = a.wsunit "
'                   -----------------------

    If pMfySeq Then
        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg >= ", MWS_Final)
    Else
        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg < ", MWS_Final)
    End If
'   SQLReadWsUnit = SQLReadWsUnit & " AND substr(a.workdt,1,6) >= (select to_char(ADD_MONTHS(SYSDATE,-6) ,'YYYYMM') || to_char(SYSDATE,'DD') from dual )"
    SQLReadWsUnit = SQLReadWsUnit & " AND a.WORKDT >= '" & strTmpDt & "'" & _
        " ) p, " & T_LAB404 & " b " & _
        " where   b.workarea = p.workarea AND b.accdt  = p.accdt  AND b.accseq  = p.accseq  AND b.rsttype = p.scfg "

    If pMfySeq Then
        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst)  '& ") "
    Else
        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) ' & ") "
    End If
    
    
    SQLReadWsUnit = SQLReadWsUnit & "Order by p.wsunit"   ' & " group by p.wscd, p.wsunit,p.workdt,p.worktm "

    ' Order by a.wsunit desc 를 옮김.
    '===============================================================================================================
'    Debug.Print SQLReadWsUnit
End Function

Public Function SQLReadWsUnit_New(ByVal pWsCd As String, ByVal pMonth As String, Optional ByVal pMfySeq As Boolean = False) As String
    '** 원본 =======================================================================================================
    '## 5.0.20: 이상대(2005-02-11)
    '   - 미생물 연속검사의 경우 마지막 접수번호만 102에 저장되어 102를 접수번호로 조인하면
    '     나머지 접수번호를 표시되지 않어서 102를 PTID, ORDDT, ORDNO, ORDSEQ로 조인하도록 수정
'    SQLReadWsUnit = " SELECT a.wsunit,a.workdt,a.worktm FROM " & T_LAB401 & " a " & _
'                    " WHERE  " & DBW("a.wscd = ", pWsCd) & _
'                    " AND    EXISTS ( SELECT * FROM " & T_LAB102 & " d," & T_LAB404 & " b, " & T_LAB402 & " c " & _
'                                    " WHERE c.wscd     = a.wscd      AND  c.wsunit  = a.wsunit " & _
'                                    " AND   b.workarea = c.workarea  AND  b.accdt   = c.accdt " & _
'                                    " AND   b.accseq   = c.accseq    AND  b.rsttype = c.scfg " & _
'                                    " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
'                                    " AND b.ordseq=d.ordseq AND (d.dcfg='' or d.dcfg is null) "
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("c.genfg >= ", MWS_Final) & _
'                        " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst) & ") "
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("c.genfg < ", MWS_Final) & _
'                        " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) & ") "
'    End If
'    SQLReadWsUnit = SQLReadWsUnit & " Order by a.wsunit desc"
    '===============================================================================================================
    
    '** 변경 By M.G.Choi 2006.12.27 (2단계로 나눔)
    '** 추가 By S.W,Oh   2007.03.27 (쿼리 추가 : "   AND b.workarea is not null   AND  b.accdt is not null AND b.accseq is not null AND b.testcd is not null" & )
    
    
'    SQLReadWsUnit = " SELECT /*+ RULE */ distinct a.wscd, a.wsunit,a.workdt,a.worktm " & _
'                    "   FROM " & T_LAB401 & " a, " & T_LAB402 & " c " & _
'                    " WHERE  " & DBW("a.wscd = ", pWsCd) & _
'                    " AND    c.wscd     = a.wscd      AND  c.wsunit  = a.wsunit " & _
'                    " AND    substr(a.workdt,1,6) >= to_char(TRUNC(ADD_MONTHS(SYSDATE,-3)+1)-TO_CHAR(SYSDATE,'DD'),'YYYYMM') " & _
'                    " AND    EXISTS ( SELECT testcd FROM " & T_LAB404 & " b " & _
'                                    " WHERE b.workarea = c.workarea  AND  b.accdt   = c.accdt " & _
'                                    "   AND b.accseq   = c.accseq    AND  b.rsttype = c.scfg "
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst) & ") " & _
'                        " AND " & DBW("c.genfg >= ", MWS_Final)
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & _
'                        " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) & ") " & _
'                        " AND " & DBW("c.genfg < ", MWS_Final)
'    End If
'    SQLReadWsUnit = SQLReadWsUnit & " Order by a.wsunit desc"
'
'----------- 2008.10.22 양성현 미생물 결과 및 수정을 위한 조회를 빠르게 하기 위해 인덱스 힌트를 둠.

'    SQLReadWsUnit = " SELECT /*+ INDEX_DESC(a, s2lab401_pk */ /*+ INDEX_DESC(c, s2lab402_pk */ /*+ INDEX_DESC(b, s2lab404_pk */ distinct a.wscd, a.wsunit,a.workdt,a.worktm " & _
'                    "   FROM " & T_LAB401 & " a, " & T_LAB402 & " c join " & T_LAB404 & _
'                    " b on ( b.workarea = c.workarea AND b.accdt  = c.accdt  AND b.accseq  = c.accseq  AND b.rsttype = c.scfg "
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst) & ") "
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) & ") "
'    End If
'
'        SQLReadWsUnit = SQLReadWsUnit & _
'                    " WHERE  " & DBW("a.wscd = ", pWsCd) & _
'                    " AND    c.wscd     = a.wscd      AND  c.wsunit  = a.wsunit " & _
'                    " AND    substr(a.workdt,1,6) >= to_char(TRUNC(ADD_MONTHS(SYSDATE,-6)+1)-TO_CHAR(SYSDATE,'DD'),'YYYYMM') "
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg >= ", MWS_Final)
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg < ", MWS_Final)
'    End If
'    SQLReadWsUnit = SQLReadWsUnit & " Order by a.wsunit desc"
'
' " SELECT /*+ Row(a) +*/ /*+ INDEX_DESC(a, s2lab401_idx5) */ distinct a.wscd, a.wsunit,a.workdt,a.worktm, c.workarea,c.accdt,c.accseq,c.scfg "
'
'    SQLReadWsUnit = " select  /*+ Row(p) +*//*+  AND_EQUALS(s2lab404_pk, s2lab404_idx5) */  distinct p.wscd, p.wsunit,p.workdt,p.worktm from ( " & _
'        " SELECT /*+ Row(a) +*//*+  AND_EQUALS(s2lab401_pk, s2lab401_idx5) */ distinct a.wscd, a.wsunit,a.workdt,a.worktm, c.workarea,c.accdt,c.accseq,c.scfg " & _
'        " FROM " & T_LAB401 & " a, " & T_LAB402 & " c " & _
'        " WHERE " & DBW("a.wscd = ", pWsCd) & " AND c.wscd = a.wscd  AND  c.wsunit  = a.wsunit "
''                   -----------------------
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg >= ", MWS_Final)
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("c.genfg < ", MWS_Final)
'    End If
''   SQLReadWsUnit = SQLReadWsUnit & " AND substr(a.workdt,1,6) >= (select to_char(ADD_MONTHS(SYSDATE,-6) ,'YYYYMM') || to_char(SYSDATE,'DD') from dual )"
'    SQLReadWsUnit = SQLReadWsUnit & " AND a.WORKDT >= to_char(SYSDATE,'YYYYMM') -6|| to_char(SYSDATE,'DD') " & _
'        " ) p, " & T_LAB404 & " b " & _
'        " where   b.workarea = p.workarea AND b.accdt  = p.accdt  AND b.accseq  = p.accseq  AND b.rsttype = p.scfg "
'
'    If pMfySeq Then
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst)  '& ") "
'    Else
'        SQLReadWsUnit = SQLReadWsUnit & " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) ' & ") "
'    End If
    
    '## 온승호(2010.03.31)
    '   - 쿼리 속도 수정
    '   - 조회 일정 6개월에서 4개월
    '   - 인덱스 수정
    
    Dim strTmpDt As String

    strTmpDt = Format(DateAdd("m", -(pMonth), Now), "yyyymmdd")

    SQLReadWsUnit_New = " select /*+ Row(p) +*//*+  AND_EQUALS(s2lab404_pk, s2lab404_idx5) */  distinct p.wscd, p.wsunit,p.workdt,p.worktm from ( " & _
        " SELECT /*+ Row(a) +*//*+  AND_EQUALS(s2lab401_pk, s2lab401_idx5) */ distinct a.wscd, a.wsunit,a.workdt,a.worktm, c.workarea,c.accdt,c.accseq,c.scfg " & _
        " FROM " & T_LAB401 & " a, " & T_LAB402 & " c " & _
        " WHERE " & DBW("a.wscd = ", pWsCd) & " AND c.wscd = a.wscd  AND  c.wsunit  = a.wsunit "
'                   -----------------------

    If pMfySeq Then
        SQLReadWsUnit_New = SQLReadWsUnit_New & " AND " & DBW("c.genfg >= ", MWS_Final)
    Else
        SQLReadWsUnit_New = SQLReadWsUnit_New & " AND " & DBW("c.genfg < ", MWS_Final)
    End If
'   SQLReadWsUnit = SQLReadWsUnit & " AND substr(a.workdt,1,6) >= (select to_char(ADD_MONTHS(SYSDATE,-6) ,'YYYYMM') || to_char(SYSDATE,'DD') from dual )"
    SQLReadWsUnit_New = SQLReadWsUnit_New & " AND a.WORKDT >= '" & strTmpDt & "'" & _
        " ) p, " & T_LAB404 & " b " & _
        " where   b.workarea = p.workarea AND b.accdt  = p.accdt  AND b.accseq  = p.accseq  AND b.rsttype = p.scfg "

    If pMfySeq Then
        SQLReadWsUnit_New = SQLReadWsUnit_New & " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst)  '& ") "
    Else
        SQLReadWsUnit_New = SQLReadWsUnit_New & " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst) ' & ") "
    End If
    
    
    SQLReadWsUnit_New = SQLReadWsUnit_New & "Order by p.wsunit"   ' & " group by p.wscd, p.wsunit,p.workdt,p.worktm "

    ' Order by a.wsunit desc 를 옮김.
    '===============================================================================================================
'    Debug.Print SQLReadWsUnit
End Function

Public Function SQLReadWsUnitSub(ByVal pWsCd As String, ByVal pWsUnit As String) As String
    
    SQLReadWsUnitSub = " SELECT  b.workarea, b.accdt, b.accseq " & _
                       "   FROM " & T_LAB402 & " b, " & T_LAB404 & " c " & _
                       "  WHERE b.wscd = " & DBS(pWsCd) & _
                       "    AND b.wsunit = " & DBS(pWsUnit) & _
                       "    AND b.workarea = c.workarea and b.accdt = c.accdt and b.accseq = c.accseq " & _
                       "    AND EXISTS ( SELECT ptid FROM ORAS1.S2ORD102_V d " & _
                       "                  WHERE c.ptid=d.ptid AND c.orddt=d.orddt AND c.ordno=d.ordno " & _
                       "                    AND c.ordseq=d.ordseq AND (d.dcfg='' or d.dcfg is null)) "
    
    SQLReadWsUnitSub = " SELECT  b.workarea, b.accdt, b.accseq " & _
                       "   FROM " & T_LAB402 & " b, " & T_LAB404 & " c,  ORAS1.S2ORD102_V d " & _
                       "  WHERE b.wscd = " & DBS(pWsCd) & _
                       "    AND b.wsunit = " & DBS(pWsUnit) & _
                       "    AND c.workarea = b.workarea and c.accdt = b.accdt and c.accseq = b.accseq " & _
                       "    AND d.ptid=c.ptid AND d.orddt=c.orddt AND d.ordno=c.ordno " & _
                       "    AND d.ordseq=c.ordseq AND (d.dcfg='' or d.dcfg is null) and rownum < 2 "
    
End Function

Public Function SQLReadWSInfo(ByVal pWsCd As String, ByVal pWsUnit As String) As String
    
    SQLReadWSInfo = " SELECT fnshdt,fnshtm,workdt,worktm " & _
                    " FROM  " & T_LAB401 & _
                    " WHERE " & DBW("wscd   = ", pWsCd) & _
                    " AND   " & DBW("wsunit = ", pWsUnit)
End Function

Public Function SQLReadWSList(ByVal pWsCd As String, ByVal pWsUnit As String, _
                              Optional ByVal pMfySeq As Boolean = False) As String

        '복수검체일경우 lab102 에 마지막 번호만 가지고 있으므로, Dc체크가 무의미함
    SQLReadWSList = " SELECT  distinct a.* " & _
                    " FROM  " & T_LAB402 & " a, " & T_LAB404 & " b " & _
                    " WHERE " & DBW("a.wscd   = ", pWsCd) & _
                    " AND   " & DBW("a.wsunit = ", pWsUnit) & _
                    " AND     b.rsttype = a.scfg "

   
    If pMfySeq Then
        SQLReadWSList = SQLReadWSList & _
                        " AND " & DBW("a.genfg >= ", MWS_Final) & _
                        " AND " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_FinRst)
    Else
        SQLReadWSList = SQLReadWSList & _
                        " AND " & DBW("a.genfg < ", MWS_Final) & _
                        " AND " & DBW("b.stscd < ", enStsCd.StsCd_LIS_FinRst)
    End If
    SQLReadWSList = SQLReadWSList & _
                " AND     a.workarea = b.workarea " & _
                " AND     a.accdt    = b.accdt  " & _
                " AND     a.accseq   = b.accseq " & _
                " AND    (b.detailfg = '' OR b.rstdiv = '*' OR b.detailfg is null)"
End Function

Public Function SQLGetStainRstCd(ByVal pTestCd As String) As String
    
    SQLGetStainRstCd = " SELECT * FROM " & T_LAB031 & _
                       " WHERE  " & DBW("cdindex = ", LC2_ItemResult) & _
                       "   AND  " & DBW("cdval1  = ", pTestCd)
End Function

Public Function SQLResultCheck(ByVal pTestCd As String, pRstCd As String) As String
        SQLResultCheck = " SELECT * FROM " & T_LAB031 & " " & _
                         " WHERE " & DBW("cdindex =", LC2_ItemResult) & _
                         " AND   " & DBW("cdval1  =", pTestCd) & _
                         " AND   " & DBW("cdval2  =", pRstCd)
End Function

Public Function SQLVerifyResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pTestCd As String, ByVal pResult As String, ByVal pStatus As String, _
                                ByVal pDate As String, ByVal pTime As String, ByVal pEmpId As String, _
                                Optional ByVal pSenFg As String = "", Optional ByVal pMfySeq As String = "0", _
                                Optional ByVal pRstType As String = "") As String

    
    ' 결과 내역 저장

    SQLVerifyResult = ""
    If pResult = "" Then
        If pStatus >= enStsCd.StsCd_LIS_MidRst Then
            SQLVerifyResult = "UPDATE " & T_LAB404 & " SET "
            SQLVerifyResult = SQLVerifyResult & " " & _
                                DBW("stscd = ", pStatus) & ", " & _
                                DBW("senfg = ", pSenFg) & ", " & _
                                DBW("mfyseq = ", pMfySeq) & ", " & _
                                DBW("vfydt = ", pDate) & ", " & _
                                DBW("vfytm = ", pTime) & ", " & _
                                DBW("vfyid = ", pEmpId)
        End If
    ElseIf pResult <> "" Then
        SQLVerifyResult = "UPDATE " & T_LAB404 & " SET "
        If pStatus < enStsCd.StsCd_LIS_MidRst Then       '결과저장 또는
            
            SQLVerifyResult = SQLVerifyResult & " " & DBW("rstcd = ", pResult)
            
        Else        'If pStatus = enStsCd.StsCd_LIS_FinRst Then  '중간결과/최종결과
            
            SQLVerifyResult = SQLVerifyResult & " " & _
                            DBW("rstcd = ", pResult) & ", " & _
                            DBW("senfg = ", pSenFg) & ", " & _
                            DBW("stscd = ", pStatus) & ", " & _
                            DBW("mfyseq = ", pMfySeq) & ", " & _
                            DBW("vfydt = ", pDate) & ", " & _
                            DBW("vfytm = ", pTime) & ", " & _
                            DBW("vfyid = ", pEmpId)
        End If
    End If
    
    If Trim(SQLVerifyResult) <> "" Then
        If pRstType <> "" Then
            SQLVerifyResult = SQLVerifyResult & ", " & DBW("rsttype = ", pRstType)
        End If
        SQLVerifyResult = SQLVerifyResult & " WHERE " & _
                            DBW("workarea = ", pWorkArea) & " AND " & DBW("accdt  = ", pAccDt) & " AND " & _
                            DBW("accseq   = ", pAccSeq)
        
        
        SQLVerifyResult = SQLVerifyResult & "   AND rsttype in ('" & MRT_MicSen & "','" & MRT_GenSen & "')"
    End If
End Function

Public Function SQLVerifyStainResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pTestCd As String, ByVal pResult As String, ByVal pStatus As String, _
                                ByVal pDate As String, ByVal pTime As String, ByVal pEmpId As String, _
                                Optional ByVal pSenFg As String = "", Optional ByVal pMfySeq As String = "0", _
                                Optional ByVal pRstType As String = "") As String

    
    ' 결과 내역 저장

    SQLVerifyStainResult = ""
    If pResult = "" Then
        If pStatus >= enStsCd.StsCd_LIS_MidRst Then
            SQLVerifyStainResult = "UPDATE " & T_LAB404 & " SET "
            SQLVerifyStainResult = SQLVerifyStainResult & " " & _
                                DBW("rstcd = ", pResult) & ", " & _
                                DBW("stscd = ", pStatus) & ", " & _
                                DBW("senfg = ", pSenFg) & ", " & _
                                DBW("mfyseq = ", pMfySeq) & ", " & _
                                DBW("vfydt = ", pDate) & ", " & _
                                DBW("vfytm = ", pTime)
        End If
    ElseIf pResult <> "" Then
        SQLVerifyStainResult = "UPDATE " & T_LAB404 & " SET "
        If pStatus < enStsCd.StsCd_LIS_MidRst Then       '결과저장 또는
            
            SQLVerifyStainResult = SQLVerifyStainResult & " " & DBW("rstcd = ", pResult) & ", " & _
                                                                DBW("lastrst = ", pResult)
            
        Else        'If pStatus = enStsCd.StsCd_LIS_FinRst Then  '중간결과/최종결과
            
            SQLVerifyStainResult = SQLVerifyStainResult & " " & _
                            DBW("rstcd = ", pResult) & ", " & _
                            DBW("senfg = ", pSenFg) & ", " & _
                            DBW("stscd = ", pStatus) & ", " & _
                            DBW("mfyseq = ", pMfySeq) & ", " & _
                            DBW("vfydt = ", pDate) & ", " & _
                            DBW("vfytm = ", pTime) & ", " & _
                            DBW("vfyid = ", pEmpId) & ", " & _
                            DBW("lastrst = ", pResult) & ", " & _
                            DBW("lastvfydt = ", pDate) & ", " & _
                            DBW("lastvfytm = ", pTime) & ", " & _
                            DBW("lastvfyid = ", pEmpId)
        End If
    End If
    
    If Trim(SQLVerifyStainResult) <> "" Then
        If pRstType <> "" Then
            SQLVerifyStainResult = SQLVerifyStainResult & ", " & DBW("rsttype = ", pRstType)
        End If
        SQLVerifyStainResult = SQLVerifyStainResult & " WHERE " & _
                            DBW("workarea = ", pWorkArea) & " AND " & DBW("accdt  = ", pAccDt) & " AND " & _
                            DBW("accseq   = ", pAccSeq)
        
        '2001-12-03 : Nogrowth결과도 MIC까지 status반영한다.
        If pSenFg = MRT_SenRst Then
            SQLVerifyStainResult = SQLVerifyStainResult & "   AND rsttype in ('" & MRT_MicSen & "','" & MRT_GenSen & "')"
        Else
            SQLVerifyStainResult = SQLVerifyStainResult & "   AND " & DBW("testcd = ", pTestCd)
        End If
    End If
End Function

Public Function SqlUpdateOrder(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                               ByVal pTestCd As String, ByVal pStatus As String, _
                               ByVal pDate As String, ByVal pTime As String, ByVal pEmpId As String, _
                               Optional ByVal pMultiTest As Boolean = False) As String
    
    '** 원본 ===========================================================================
'    SqlUpdateOrder = " UPDATE " & T_LAB102 & _
'                     " SET   " & DBW("stscd    = ", pStatus) & ", " & _
'                                 DBW("examdt   = ", pDate) & ", " & _
'                                 DBW("examtm   = ", pTime) & ", " & _
'                                 DBW("examdoct = ", pEmpId) & _
'                     " WHERE " & DBW("workarea = ", pWorkArea) & _
'                     "  AND  " & DBW("accdt    = ", pAccDt) & _
'                     "  AND  " & DBW("accseq   = ", pAccSeq)
'    If pMultiTest Then
'        SqlUpdateOrder = SqlUpdateOrder & " AND   ordcd   in (" & pTestCd & ")"
'    Else
'        SqlUpdateOrder = SqlUpdateOrder & " AND   " & DBW("ordcd   = ", pTestCd)
'    End If
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strWrkDiv = "" & rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    rs.Close: Set rs = Nothing
    
    If strWrkDiv = "3" Then         '종건
        SqlUpdateOrder = " UPDATE su2examt " & _
                         " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                         " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                         ", examdr = " & DBS(pEmpId) & _
                         " WHERE " & DBW("workarea = ", pWorkArea) & _
                         "  AND  " & DBW("accdt    = ", pAccDt) & _
                         "  AND  " & DBW("accseq   = ", pAccSeq)
        If pMultiTest Then
            SqlUpdateOrder = SqlUpdateOrder & " AND   examcode   in (" & pTestCd & ")"
        Else
            SqlUpdateOrder = SqlUpdateOrder & " AND   examcode = " & DBS(pTestCd)
        End If
    ElseIf strWrkDiv = "4" Then     '일건
            SqlUpdateOrder = " UPDATE sg2examt " & _
                             " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                             " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                             ", examdr = " & DBS(pEmpId) & _
                             " WHERE " & DBW("workarea = ", pWorkArea) & _
                             "  AND  " & DBW("accdt    = ", pAccDt) & _
                             "  AND  " & DBW("accseq   = ", pAccSeq)
            If pMultiTest Then
                SqlUpdateOrder = SqlUpdateOrder & " AND   examcode   in (" & pTestCd & ")"
            Else
                SqlUpdateOrder = SqlUpdateOrder & " AND   examcode = " & DBS(pTestCd)
            End If
    Else
        SqlUpdateOrder = " UPDATE mdexmort " & _
                         " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                         " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                         ", cofmdr = " & DBS(pEmpId) & _
                         " WHERE " & DBW("workarea = ", pWorkArea) & _
                         "  AND  " & DBW("accdt    = ", pAccDt) & _
                         "  AND  " & DBW("accseq   = ", pAccSeq)
        If pMultiTest Then
            SqlUpdateOrder = SqlUpdateOrder & " AND   ordcd   in (" & pTestCd & ")"
        Else
            SqlUpdateOrder = SqlUpdateOrder & " AND   ordcd = " & DBS(pTestCd)
        End If
    End If
    '===================================================================================

End Function


Public Function SqlUpdateOrderStatus(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As String, _
                                     ByVal pOrdSeq As String, ByVal pStatus As String, _
                                     ByVal pDate As String, ByVal pTime As String, ByVal pEmpId As String) As String
    
    '** 원본 ===========================================================================
'    SqlUpdateOrderStatus = " UPDATE " & T_LAB102 & _
'                            " SET   " & DBW("stscd    = ", pStatus) & ", " & _
'                                        DBW("examdt   = ", pDate) & ", " & _
'                                        DBW("examtm   = ", pTime) & ", " & _
'                                        DBW("examdoct = ", pEmpId) & _
'                            " WHERE " & DBW("ptid   = ", pPtid) & _
'                            " AND   " & DBW("orddt  = ", pOrdDt) & _
'                            " AND   " & DBW("ordno  = ", pOrdNo) & _
'                            " AND   " & DBW("ordseq = ", pOrdSeq)
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(pPtid) & _
             "    and orddt = " & DBS(pOrdDt) & _
             "    and ordno = " & DBN(pOrdNo)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strWrkDiv = "" & rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    rs.Close: Set rs = Nothing
    
    If strWrkDiv = "3" Then         '종건
        SqlUpdateOrderStatus = " UPDATE su2examt " & _
                               " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                               " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                               ", examdr = " & DBS(pEmpId) & _
                               " WHERE patno = " & DBS(pPtid) & _
                               " AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                               " AND   ordseqno = " & DBN(pOrdNo)
    ElseIf strWrkDiv = "4" Then     '일건
            SqlUpdateOrderStatus = " UPDATE sg2examt " & _
                                   " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                                   " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                                   ", examdr = " & DBS(pEmpId) & _
                                   " WHERE patno = " & DBS(pPtid) & _
                                   " AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                                   " AND   ordseqno = " & DBN(pOrdNo)
    Else
        SqlUpdateOrderStatus = " UPDATE mdexmort " & _
                           " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                           " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                           ", cofmdr = " & DBS(pEmpId) & _
                           " WHERE patno = " & DBS(pPtid) & _
                           " AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                           " AND   ordseqno = " & DBN(pOrdNo)
    End If
    '===================================================================================
    If Mid(pPtid, 1, 1) = "L" Then
        SqlUpdateOrderStatus = " UPDATE s2ord999 " & _
                           " SET   " & DBW("stscd    = ", pStatus) & ", " & _
                           " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                           ", cofmdr = " & DBS(pEmpId) & _
                           " WHERE patno = " & DBS(pPtid) & _
                           " AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                           " AND   ordseqno = " & DBN(pOrdNo)
    End If

End Function


Public Function SqlGetOrderNumber(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                  ByVal pTestCd As String, Optional ByVal pMultiTest As Boolean = False) As String
        
    SqlGetOrderNumber = " SELECT * FROM " & T_LAB404 & " b " & _
                        "  WHERE " & DBW("b.workarea = ", pWorkArea) & _
                        "  AND   " & DBW("b.accdt    = ", pAccDt) & _
                        "  AND   " & DBW("b.accseq   = ", pAccSeq)
    If pMultiTest Then
        SqlGetOrderNumber = SqlGetOrderNumber & " AND   b.testcd   in (" & pTestCd & ")"
    Else
        SqlGetOrderNumber = SqlGetOrderNumber & " AND   " & DBW("b.testcd   = ", pTestCd)
    End If

End Function

Public Function SqlGetSenOrdNumber(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                  ByVal pTestCd As String, Optional ByVal pMultiTest As Boolean = False) As String
    SqlGetSenOrdNumber = " SELECT * FROM " & T_LAB404 & " b " & _
                        "  WHERE " & DBW("b.workarea = ", pWorkArea) & _
                        "  AND   " & DBW("b.accdt    = ", pAccDt) & _
                        "  AND   " & DBW("b.accseq   = ", pAccSeq) & _
                        "  AND   b.rsttype in ('" & MRT_MicSen & "','" & MRT_GenSen & "')"
End Function

Public Function SqlOrdVerifyCheck(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As String, _
                                  ByVal pOrdSeq As String) As String
        
    SqlOrdVerifyCheck = " SELECT b.workarea, b.accdt, b.accseq FROM " & T_LAB404 & " b " & _
                        "  WHERE " & DBW("b.ptid = ", pPtid) & _
                        "  AND   " & DBW("b.orddt    = ", pOrdDt) & _
                        "  AND   " & DBW("b.ordno    = ", pOrdNo) & _
                        "  AND   " & DBW("b.ordseq   = ", pOrdSeq) & _
                        "  AND   (b.vfydt = ''  or  b.vfydt is null) "
    SqlOrdVerifyCheck = SqlOrdVerifyCheck & " UNION ALL " & _
                        " SELECT b.workarea, b.accdt, b.accseq FROM " & T_LAB201 & " b, " & T_LAB203 & " a " & _
                        "  WHERE " & DBW("a.ptid = ", pPtid) & _
                        "  AND   " & DBW("a.orddt    = ", pOrdDt) & _
                        "  AND   " & DBW("a.ordno    = ", pOrdNo) & _
                        "  AND   " & DBW("a.ordseq   = ", pOrdSeq) & _
                        "  AND    b.workarea = a.workarea " & _
                        "  AND    b.accdt    = a.accdt " & _
                        "  AND    b.accseq   = a.accseq " & _
                        "  AND   " & DBW("b.stscd <= ", enStsCd.StsCd_LIS_Collection)
End Function



Public Function SQLUpdateWoksheet(ByVal pCaseFg As Long, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pWsCd As String, ByVal pWsUnit As String, ByVal pSelFg As String, _
                                  Optional ByVal pH As Long = 0) As String
    
    Dim sTemp As String
    
    Select Case pCaseFg
        Case 1: 'Nogrowth 배치결과등록
            SQLUpdateWoksheet = " UPDATE " & T_LAB402 & _
                                " SET    " & DBW("genfg    = ", pSelFg) & _
                                " WHERE  " & DBW("wscd     = ", pWsCd) & _
                                " AND    " & DBW("workarea = ", pWorkArea) & _
                                " AND    " & DBW("accdt    = ", pAccDt) & _
                                " AND    " & DBW("accseq   = ", pAccSeq)
        Case 2: 'Stain 결과등록
            sTemp = ""
            If pH = 0 Then sTemp = ", " & DBW("genfg = ", MWS_Final)
        
            SQLUpdateWoksheet = " UPDATE " & T_LAB402 & _
                                " SET    " & DBW("selfg    = ", pSelFg) & sTemp & _
                                " WHERE  " & DBW("wscd     = ", pWsCd) & _
                                " AND    " & DBW("workarea = ", pWorkArea) & _
                                " AND    " & DBW("accdt    = ", pAccDt) & _
                                " AND    " & DBW("accseq   = ", pAccSeq)
        Case 3: '감수성결과등록(Nogrowth)
            SQLUpdateWoksheet = " UPDATE " & T_LAB402 & " " & _
                                " SET    " & DBW("selfg    = ", pSelFg) & _
                                " WHERE  " & DBW("workarea = ", pWorkArea) & _
                                " AND    " & DBW("accdt    = ", pAccDt) & _
                                " AND    " & DBW("accseq   = ", pAccSeq) & _
                                " AND  scfg in (" & DBV("scfg", MRT_GenSen) & "," & _
                                                    DBV("scfg", MRT_MicSen) & ") "
    
    End Select

'    If Mid(pPtid, 1, 1) = "L" Then
'        SqlUpdateOrderStatus = " UPDATE s2ord999 " & _
'                           " SET   " & DBW("stscd    = ", pStatus) & ", " & _
'                           " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
'                           ", cofmdr = " & DBS(pEmpId) & _
'                           " WHERE patno = " & DBS(pPtid) & _
'                           " AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
'                           " AND   ordseqno = " & DBN(pOrdNo)
End Function

Public Function SqlGetTestCount(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

    SqlGetTestCount = " SELECT  reqtotcnt, reqinputcnt, stscd " & _
                      " FROM  " & T_LAB201 & " " & _
                      " WHERE " & DBW("workarea = ", pWorkArea) & _
                      " AND   " & DBW("accdt    = ", pAccDt) & _
                      " AND   " & DBW("accseq   = ", pAccSeq)

End Function


Public Function SQLUpdateStatus(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pStatus As String, ByVal pInCnt As Long, ByVal pDate As String, _
                                ByVal pTime As String, ByVal pEmpId As String) As String

    ' 접수 내역에 Status 및 조회정보 반영
    
    Select Case pStatus
        Case enStsCd.StsCd_LIS_MidRst, enStsCd.StsCd_LIS_FinRst
            SQLUpdateStatus = " UPDATE " & T_LAB201 & " " & _
                              " SET    " & DBW("stscd       = ", pStatus) & "," & _
                                           DBW("reqinputcnt = ", pInCnt) & "," & _
                                           DBW("vfydt       = ", pDate) & ", " & _
                                           DBW("vfytm       = ", pTime) & ", " & _
                                           DBW("vfyid       = ", pEmpId) & " " & _
                              " WHERE  " & DBW("workarea    = ", pWorkArea) & _
                              " AND    " & DBW("accdt       = ", pAccDt) & _
                              " AND    " & DBW("accseq      = ", pAccSeq)
    
        Case enStsCd.StsCd_LIS_Modify
            SQLUpdateStatus = " UPDATE " & T_LAB201 & " " & _
                              " SET    " & DBW("stscd       = ", pStatus) & "," & _
                              "        " & DBW("vfydt    = ", pDate) & "," & _
                              "        " & DBW("vfytm    = ", pTime) & "," & _
                              "        " & DBW("vfyid    = ", pEmpId) & _
                              " WHERE  " & DBW("workarea = ", pWorkArea) & _
                              " AND    " & DBW("accdt    = ", pAccDt) & _
                              " AND    " & DBW("accseq   = ", pAccSeq)
        Case Else
            SQLUpdateStatus = " UPDATE " & T_LAB201 & _
                              " SET    " & DBW("reqinputcnt = ", pInCnt) & _
                              " WHERE  " & DBW("workarea    = ", pWorkArea) & _
                              " AND    " & DBW("accdt       = ", pAccDt) & _
                              " AND    " & DBW("accseq      = ", pAccSeq)
    End Select

End Function

Public Function SQLSaveFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pSeq As Long, ByVal pEmpId As String, ByVal pTextRst As String, ByVal pFg As Long) As String
    
    If pFg = 0 Then
        SQLSaveFootNote = " DELETE " & T_LAB304 & _
                          " WHERE  " & DBW("workarea = ", pWorkArea) & _
                          " AND    " & DBW("accdt    = ", pAccDt) & _
                          " AND    " & DBW("accseq   = ", pAccSeq)
    Else
        SQLSaveFootNote = " INSERT INTO " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
                          " VALUES (" & DBV("workarea", pWorkArea) & "," & DBV("accdt", pAccDt) & "," & _
                                        DBV("accseq", pAccSeq) & "," & DBV("seq", pSeq) & "," & _
                                        DBV("vfyid", pEmpId) & "," & DBV("rsttxt", pTextRst) & ")"
    End If

End Function

Public Function SQLSaveRemark(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                              ByVal pSeq As Long, ByVal pRmkCd As String) As String
    
    ' 접수 내역에 Footnote Seq 및 검체 Remark Code 등록
    SQLSaveRemark = " UPDATE " & T_LAB201 & _
                    " SET    " & DBW("footnotefg = ", pSeq) & "," & _
                                 DBW("rmkcd      = ", pRmkCd) & "" & _
                    " WHERE  " & DBW("workarea   = ", pWorkArea) & _
                    " AND    " & DBW("accdt      = ", pAccDt) & _
                    " AND    " & DBW("accseq     = ", pAccSeq)

End Function

Public Function SQLGetPtInfoByLabno(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, Optional ByVal pRstTp As String = "") As String
' 08.11.07 양성현 Stain 결과등록 화면  진료과명, 처방Remark, 처방의, 부서전화번호 등 수정추가
' 08.12.27 양성현 건진의 경우 의사ID가 없는 관계로 (+)를 추가함.
    SQLGetPtInfoByLabno = " SELECT  c.stscd,c.testcd,c.rsttype,c.mfyseq,a.ptid,f." & F_PTNM & " ptnm," & _
                          " a.sex,a.ageday,a.deptcd,a.wardid,a.hosilid,a.bedid,a.majdoct,a.orddoct," & _
                          " a.spccd,a.footnotefg,a.rmkcd,a.rcvdt,b.field4 as spcnm, c.vfydt, g.username as orddrnm, " & _
                          " h.mesg, i.deptnm, i.phone" & _
                          " FROM  " & T_LAB201 & " a, " & T_LAB032 & " b, " & T_LAB404 & " c, " & T_HIS001 & " f, " & T_HIS005 & " g, " & T_LAB102 & " h, " & T_HIS003 & " i  " & _
                          " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                          " AND " & DBW("a.accdt    = ", pAccDt) & _
                          " AND " & DBW("a.accseq   = ", pAccSeq) & _
                          " AND g.userid (+)= a.orddoct and c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                          "  " & _
                          " AND     a.ptid   =  f." & F_PTID & _
                          " AND   " & DBW("b.cdindex = ", LC3_Specimen) & _
                          " AND     b.cdval1 =  a.spccd and h.ptid = a.ptid and h.workarea = a.workarea  and h.accdt = a.accdt and h.accseq = a.accseq and i.dpcd = a.deptcd  "
    
    'If pRstTp <> "" AND Not P_MICSelectedByUser Then SQLGetPtInfoByLabno = SQLGetPtInfoByLabno & " AND " & DBW("c.rsttype = ", pRstTp)
    If pRstTp <> "" Then SQLGetPtInfoByLabno = SQLGetPtInfoByLabno & " AND c.rsttype in (" & pRstTp & ")"
    'AND   ( c.rstdiv = '*' or c.rstdiv = '' or c.rstdiv is null )
End Function

Public Function SQLGetPtInfoByBarNo(ByVal pSpcYY As String, ByVal pSpcNo As String, _
                                    Optional ByVal pRstTp As String = "") As String
' 08.11.07 양성현 Stain 결과등록 화면  진료과명, 처방Remark, 처방의, 부서전화번호 등 수정추가
' 08.12.27 양성현 건진의 경우 의사ID가 없는 관계로 (+)를 추가함.
    SQLGetPtInfoByBarNo = " SELECT  a.workarea, a.accdt, a.accseq, c.stscd,c.testcd,c.rsttype,c.mfyseq,a.ptid,f." & F_PTNM & " ptnm," & _
                          " a.sex,a.ageday,a.deptcd,a.wardid,a.hosilid,a.bedid,a.majdoct,a.orddoct," & _
                          " a.spccd,a.footnotefg,a.rmkcd,b.field4 as spcnm, c.vfydt, g.username as orddrnm, " & _
                          " h.mesg, i.deptnm, i.phone" & _
                          " FROM  " & T_LAB201 & " a, " & T_LAB032 & " b, " & T_LAB404 & " c, " & T_HIS001 & " f, " & T_HIS005 & " g, " & T_LAB102 & " h, " & T_HIS003 & " i  " & _
                          " WHERE " & DBW("a.spcyy = ", pSpcYY) & _
                          " AND " & DBW("a.spcno    = ", pSpcNo) & _
                          " AND g.userid (+)= a.orddoct and c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                          "  " & _
                          " AND     a.ptid   =  f." & F_PTID & _
                          " AND   " & DBW("b.cdindex = ", LC3_Specimen) & _
                          " AND     b.cdval1 =  a.spccd and h.ptid = a.ptid and h.workarea = a.workarea  and h.accdt = a.accdt and h.accseq = a.accseq and i.dpcd = a.deptcd  "
    
    'If pRstTp <> "" AND Not P_MICSelectedByUser Then SQLGetPtInfoByLabno = SQLGetPtInfoByLabno & " AND " & DBW("c.rsttype = ", pRstTp)
    If pRstTp <> "" Then SQLGetPtInfoByBarNo = SQLGetPtInfoByBarNo & " AND c.rsttype in (" & pRstTp & ")"
    'AND   ( c.rstdiv = '*' or c.rstdiv = '' or c.rstdiv is null )
End Function

Public Function SQLGetPtInfoByTestCd(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                     ByVal pAccSeq As String, ByVal pTestCd As String) As String

' 08.11.07 양성현 감수성 결과등록 화면  진료과명, 처방Remark, 처방의, 부서전화번호 등 수정추가
' 08.12.27 양성현 건진의 경우 의사ID가 없는 관계로 (+)를 추가함.

    SQLGetPtInfoByTestCd = " SELECT  c.stscd,c.testcd,c.rsttype,c.mfyseq,a.ptid,f." & F_PTNM & " ptnm," & _
                           "         a.sex,a.ageday,a.deptcd,a.wardid,a.hosilid,a.bedid,a.majdoct,a.orddoct," & _
                           "         a.spccd,a.footnotefg,a.rmkcd,a.rcvdt,b.field4 as spcnm, c.vfydt, g.username as orddrnm,h.mesg, i.deptnm, i.phone, a.bedindt" & _
                           " FROM  " & T_LAB201 & " a, " & T_LAB032 & " b, " & T_LAB404 & " c, " & T_HIS001 & " f, " & T_HIS005 & " g, " & T_LAB102 & " h, " & T_HIS003 & " i " & _
                           " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                           " AND " & DBW("a.accdt    = ", pAccDt) & _
                           " AND " & DBW("a.accseq   = ", pAccSeq) & _
                           " AND g.userid (+)= a.orddoct and c.workarea = a.workarea AND c.accdt = a.accdt AND c.accseq = a.accseq " & _
                           " AND " & DBW("c.testcd   = ", pTestCd) & _
                           " AND a.ptid   =  f." & F_PTID & _
                           " AND " & DBW("b.cdindex = ", LC3_Specimen) & _
                           " AND b.cdval1 =  a.spccd and h.ptid = a.ptid and h.workarea = a.workarea  and h.accdt = a.accdt and h.accseq = a.accseq and i.dpcd = a.deptcd "
End Function

Public Function SqlGetFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String) As String

    SqlGetFootNote = " SELECT * FROM " & T_LAB304 & " " & _
                     " WHERE  " & DBW("workarea = ", pWorkArea) & _
                     " AND    " & DBW("accdt    = ", pAccDt) & _
                     " AND    " & DBW("accseq   = ", pAccSeq)
End Function

Public Function SQLGetMaxNoOfFootnote(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                      ByVal pAccSeq As String) As String

    SQLGetMaxNoOfFootnote = " SELECT max(seq) mfyno FROM " & T_LAB304 & _
                            " WHERE " & DBW("workarea = ", pWorkArea) & _
                            " AND   " & DBW("accdt    = ", pAccDt) & _
                            " AND   " & DBW("accseq   = ", pAccSeq)

End Function

Public Function SQLLoadStainRst(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String, ByVal pRstTp As String, _
                                Optional ByVal pMfyFg As Boolean = False) As String
    
    SQLLoadStainRst = " SELECT a.*, b.testnm, b.abbrnm10, c.field1 rst1, d.field1 rst2 " & _
                      " FROM   " & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB031 & " c, " & T_LAB031 & " d " & _
                      " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                      " AND    " & DBW("a.accdt    = ", pAccDt) & _
                      " AND    " & DBW("a.accseq   = ", pAccSeq) & _
                      " AND    a.rsttype  in  (" & pRstTp & ")"
    If pMfyFg Then
        SQLLoadStainRst = SQLLoadStainRst & _
                      " AND    " & DBW("a.stscd    >= ", enStsCd.StsCd_LIS_FinRst)
    Else
        SQLLoadStainRst = SQLLoadStainRst & _
                      " AND    " & DBW("a.stscd    < ", enStsCd.StsCd_LIS_FinRst) ' 최종확인도 수정 가능하게 처리 온승호 2016.03.28 DBW("a.stscd    < ", enStsCd.StsCd_LIS_FinRst)
    End If
    
    SQLLoadStainRst = SQLLoadStainRst & _
                      " AND    b.testcd  =  a.testcd " & _
                      " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                      " AND    " & DBJ(DBW("c.cdindex  = ", LC2_ItemResult)) & _
                      " AND    " & DBJ("c.cdval1  =* a.testcd ") & _
                      " AND    " & DBJ("c.cdval2  =* a.rstcd ") & _
                      " AND    " & DBJ(DBW("d.cdindex  = ", LC2_ItemResult)) & _
                      " AND    " & DBJ("d.cdval1  =* a.testcd ") & _
                      " AND    " & DBJ("d.cdval2  =* a.lastrst ") & _
                      " ORDER  BY  a.detailfg, b.rptseq "

End Function

'----------------------------------------------------------------------------------
'4. 미생물 Culture 결과등록 관련 Sql
'----------------------------------------------------------------------------------

Public Function SQLLoadMicrobe() As String

    ' 균 코드 로드
    SQLLoadMicrobe = " SELECT * FROM " & T_LAB032 & _
                     " WHERE " & DBW("cdindex = ", LC3_Microbe) & _
                     " ORDER BY cdval1 asc"

End Function

Public Function SQLLoadQuantity() As String
    
    ' 미생물 정도 코드 로드
    SQLLoadQuantity = " SELECT * FROM " & T_LAB032 & _
                      " WHERE " & DBW("cdindex = ", LC3_Volume) & _
                      " ORDER BY cdval1 asc"

End Function

Public Function SQLLoadAntibiotic(ByVal pSenType As String, ByVal pMicGroup As String) As String
    
    ' 균종별 항생제 리스트 로드
    If pSenType = MRT_MicSen Then
        SQLLoadAntibiotic = " SELECT text1 FROM " & T_LAB031 & _
                            " WHERE " & DBW("cdindex = ", LC2_MicroAnti) & _
                            " AND   " & DBW("cdval1  = ", pMicGroup) & _
                            " AND   " & DBW("cdval2  = ", MCD_MSen)
    Else
        SQLLoadAntibiotic = " SELECT text1 FROM " & T_LAB031 & _
                            " WHERE " & DBW("cdindex = ", LC2_MicroAnti) & _
                            " AND   " & DBW("cdval1  = ", pMicGroup) & _
                            " AND   " & DBW("cdval2  = ", MCD_GSen)
    End If

End Function

Public Function SQLLoadTestCdByLabNo(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                     ByVal pAccSeq As String) As String
    SQLLoadTestCdByLabNo = " SELECT a.testcd, a.rsttype, b.abbrnm10, b.rptseq " & _
                           " FROM " & T_LAB404 & " a, " & T_LAB001 & " b " & _
                           " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                           " AND   " & DBW("a.accdt    = ", pAccDt) & _
                           " AND   " & DBW("a.accseq   = ", pAccSeq) & _
                           " AND   a.rsttype in (" & DBV("rsttype", MRT_GenSen) & "," & _
                                                   DBV("rsttype", MRT_MicSen) & ") " & _
                           " AND   b.testcd = a.testcd " & _
                           " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                              " WHERE  testcd = b.testcd) " & _
                           " ORDER BY rptseq"
    
End Function

Public Function SQLGetStainResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String) As String
    
    SQLGetStainResult = " SELECT a.testnm, a.abbrnm10, b.rstcd, b.rsttype, b.rstdiv, c.field1 rstnm" & _
                        " FROM " & T_LAB001 & " a, " & T_LAB031 & " c, " & T_LAB404 & " b " & _
                        " WHERE  a.testcd = b.testcd " & _
                        " AND    a.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                            " WHERE testcd = a.testcd " & _
                                            " AND   applydt <= " & DBV("applydt", Format(Now, CS_DateDbFormat)) & ")" & _
                        " AND  " & DBW("b.workarea = ", pWorkArea) & _
                        " AND  " & DBW("b.accdt    = ", pAccDt) & _
                        " AND  " & DBW("b.accseq   = ", pAccSeq) & _
                        " AND  " & DBW("b.rsttype  = ", MRT_Stain) & _
                        " AND  " & DBJ(DBW("c.cdindex  = ", LC2_ItemResult)) & _
                        " AND  " & DBJ("c.cdval1  =* b.testcd ") & _
                        " AND  " & DBJ("c.cdval2  =* b.rstcd ")


End Function

Public Function SQLGetGrowthResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                   ByVal pAccSeq As String, ByVal pWsCd As String, ByVal pRstTp As String) As String
'pWsCd - 미생물 Worksheet코드("GC")

    SQLGetGrowthResult = " SELECT  a.field1 rst, b.senfg, b.vfydt " & _
                         " FROM  " & T_LAB031 & " a, " & T_LAB404 & " b " & _
                         " WHERE " & DBW("b.workarea = ", pWorkArea) & _
                         " AND   " & DBW("b.accdt    = ", pAccDt) & _
                         " AND   " & DBW("b.accseq   = ", pAccSeq) & _
                         " AND   " & DBW("b.testcd  = ", pRstTp) & _
                         " AND   " & DBJ(DBW("a.cdindex  = ", LC2_MBatchRst)) & _
                         " AND   " & DBJ(DBW("a.cdval1   = ", pWsCd)) & _
                         " AND   " & DBJ("a.cdval2   =* b.rstcd")

End Function

Public Function SQLGetSensiResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pTestCd As String, _
                                  Optional ByVal pMfySeq As String = "0") As String
        
    SQLGetSensiResult = " SELECT a.*, b.field2 mdiv, b.text1 mnm, c.field2 mqt " & _
                        " FROM  " & T_LAB405 & " a, " & T_LAB032 & " b, " & T_LAB032 & " c " & _
                        " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                        " AND   " & DBW("a.accdt    = ", pAccDt) & _
                        " AND   " & DBW("a.accseq   = ", pAccSeq) & _
                        " AND   " & DBW("a.testcd   = ", pTestCd) & _
                        " AND   " & DBW("a.mfyseq   = ", pMfySeq) & _
                        " AND   " & DBJ(DBW("b.cdindex  = ", LC3_Microbe)) & _
                        " AND   " & DBJ("b.cdval1   =* a.mnmcd") & _
                        " AND   " & DBJ(DBW("c.cdindex  = ", LC3_Volume)) & _
                        " AND   " & DBJ("c.cdval1   =* a.mqtcd")
  
End Function

Public Function SQLDelSensiResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pTestCd As String) As String
    SQLDelSensiResult = " DELETE " & T_LAB405 & _
                        " WHERE  " & DBW("workarea = ", pWorkArea) & _
                        " AND    " & DBW("accdt    = ", pAccDt) & _
                        " AND    " & DBW("accseq   = ", pAccSeq) & _
                        " AND    " & DBW("testcd   = ", pTestCd)

End Function

    
Public Function SQLGetCurrentRst(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String) As String
    
    SQLGetCurrentRst = " SELECT * FROM " & T_LAB404 & _
                       " WHERE " & DBW("workarea = ", pWorkArea) & _
                       " AND   " & DBW("accdt    = ", pAccDt) & _
                       " AND   " & DBW("accseq   = ", pAccSeq) & _
                       " AND   testcd in ( " & pTestCd & ") "
'                       " AND   " & DBW("stscd    < ", enStsCd.StsCd_LIS_FinRst)

End Function


Public Function SQLInsertOldData(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pMfySeq As String, _
                                 ByVal pRstCd As String, ByVal pSenFg As String, ByVal pRstDiv As String, _
                                 ByVal pDetailFg As String, ByVal pVfyDt As String, _
                                 ByVal pVfyTm As String, ByVal pVfyId As String) As String
    
    SQLInsertOldData = " INSERT INTO " & T_LAB407 & _
                       " (workarea,accdt,accseq,testcd,mfyseq,rstcd,senfg,rstdiv,detailfg,vfydt,vfytm,vfyid)" & _
                       " VALUES " & _
                       " ( " & _
                                DBV("workarea", pWorkArea) & "," & _
                                DBV("accdt", pAccDt) & "," & _
                                DBV("accseq", pAccSeq) & "," & _
                                DBV("testcd", pTestCd) & "," & _
                                DBV("mfyseq", pMfySeq) & "," & _
                                DBV("rstcd", pRstCd) & "," & _
                                DBV("senfg", pSenFg) & "," & _
                                DBV("rstdiv", pRstDiv) & "," & _
                                DBV("detailfg", pDetailFg) & "," & _
                                DBV("vfydt", pVfyDt) & "," & _
                                DBV("vfytm", pVfyTm) & "," & _
                                DBV("vfyid", pVfyId) & _
                       " )"

End Function


Public Function SqlGetFinList(ByVal pMidVfyDt As String) As String

    SqlGetFinList = " SELECT distinct a.workarea, a.accdt, a.accseq, a.ptid, b." & F_PTNM & " ptnm " & _
                    " FROM " & T_HIS001 & " b, " & T_LAB404 & " a " & _
                    " WHERE " & DBW("a.workarea = ", MIC_WorkArea) & _
                    " AND   " & DBW("a.vfydt = ", pMidVfyDt) & _
                    " AND     a.senfg = 'Y' " & _
                    " AND   " & DBW("a.stscd = ", enStsCd.StsCd_LIS_MidRst) & _
                    " AND     b." & F_PTID & " = a.ptid " & _
                    " order by accdt, accseq "
                    
                    
End Function

Public Function SqlGetFinList_New(ByVal pMidFVfyDt As String, ByVal pMidTVfyDt As String) As String

    SqlGetFinList_New = " SELECT distinct a.workarea, a.accdt, a.accseq, a.ptid, b." & F_PTNM & " ptnm " & _
                        " FROM " & T_HIS001 & " b, " & T_LAB404 & " a " & _
                        " WHERE " & DBW("a.workarea = ", MIC_WorkArea) & _
                        " AND   " & DBW("a.vfydt >= ", pMidFVfyDt) & _
                        " AND   " & DBW("a.vfydt <= ", pMidTVfyDt) & _
                        " AND     a.senfg = 'Y' " & _
                        " AND   " & DBW("a.stscd = ", enStsCd.StsCd_LIS_MidRst) & _
                        " AND     b." & F_PTID & " = a.ptid " & _
                        " order by accdt, accseq "

' 2009.02.19 김대용 선생 컨플레임

' " AND     a.rsttype in ( 'C', 'S') "
'                        " AND     a.rsttype = 'C' " & _

End Function
Public Function SqlGetFinList_Gra(ByVal pMidFVfyDt As String, ByVal pMidTVfyDt As String) As String
'2008.12.18. 양성현 Stain 결과 등록
    SqlGetFinList_Gra = " SELECT distinct a.workarea, a.accdt, a.accseq, a.ptid, b." & F_PTNM & " ptnm " & _
                        " FROM " & T_HIS001 & " b, " & T_LAB404 & " a " & _
                        " WHERE " & DBW("a.workarea = ", MIC_WorkArea) & _
                        " AND   " & DBW("a.vfydt >= ", pMidFVfyDt) & _
                        " AND   " & DBW("a.vfydt <= ", pMidTVfyDt) & _
                        " AND     a.rsttype = 'S' " & _
                        " AND   " & DBW("a.stscd = ", enStsCd.StsCd_LIS_MidRst) & _
                        " AND     b." & F_PTID & " = a.ptid " & _
                        " AND     a.senfg = 'Y' " & _
                        " order by accdt, accseq "
                    
'                        " AND     a.rsttype not in ('C', 'S') " & _

End Function

Public Function SqlSensiAccNo(ByVal pPtid As String, ByVal pVfyDt As String, ByVal pSpcCd As String) As String

    SqlSensiAccNo = " SELECT a.workarea, a.accdt, a.accseq, a.vfydt, a.vfytm " & _
                    " FROM " & T_LAB404 & " a, " & T_LAB201 & " b " & _
                    " WHERE " & DBW("a.ptid = ", pPtid) & _
                    " AND   " & DBW("a.vfydt <= ", pVfyDt) & _
                    " AND     a.senfg = 'Y' " & _
                    " AND   " & DBW("a.stscd >= ", enStsCd.StsCd_LIS_FinRst) & _
                    " AND     b.workarea = a.workarea " & _
                    " AND     b.accdt = a.accdt " & _
                    " AND     b.accseq = a.accseq " & _
                    " AND   " & DBW("b.spccd = ", pSpcCd) & _
                    " order by a.vfydt desc, a.vfytm desc "
                    
'    SqlSensiAccNo = " SELECT a.workarea, a.accdt, a.accseq, a.vfydt, a.vfytm " & _
'                    " FROM " & T_LAB404 & " a, " & T_LAB201 & " b " & _
'                    " WHERE " & DBW("a.ptid = ", pPtid) & _
'                    " AND  (a.vfydt<>'' or a.vfydt is not null ) " & _
'                    " AND     a.senfg = 'Y' " & _
'                    " AND   " & DBW("a.stscd >= ", enStsCd.StsCd_LIS_FinRst) & _
'                    " AND     b.workarea = a.workarea " & _
'                    " AND     b.accdt = a.accdt " & _
'                    " AND     b.accseq = a.accseq " & _
'                    " AND   " & DBW("b.spccd = ", pSpcCd) & _
'                    " order by a.rcvdt desc, a.rcvtm desc "
                    
End Function

'Public Function SqlNoGrowthAccNo(ByVal pPtid As String, ByVal pVfyDt As String, ByVal pSpcCd As String) As String
'
'    SqlNoGrowthAccNo = " SELECT a.workarea, a.accdt, a.accseq, a.vfydt, a.vfytm " & _
'                       " FROM " & T_LAB404 & " a, " & T_LAB201 & " b " & _
'                       " WHERE " & DBW("a.ptid = ", pPtid) & _
'                       " AND   " & DBW("a.vfydt <= ", pVfyDt) & _
'                       " AND     (a.senfg = '' or a.senfg is null)" & _
'                       " AND   " & DBW("a.stscd = ", enStsCd.StsCd_LIS_FinRst) & _
'                       " AND     b.workarea = a.workarea " & _
'                       " AND     b.accdt = a.accdt " & _
'                       " AND     b.accseq = a.accseq " & _
'                       " AND   " & DBW("b.spccd = ", pSpcCd) & _
'                       " order by a.vfydt desc, a.vfytm desc "
'
'End Function


