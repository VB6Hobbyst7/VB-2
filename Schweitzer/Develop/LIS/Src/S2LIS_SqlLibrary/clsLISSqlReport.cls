VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarRiPrint As String

Public Property Let RiPrint(ByVal vNewValue As String)
    mvarRiPrint = vNewValue
End Property

Public Function SqlGetPtInfo(ByVal pWorkArea As String, _
                             ByVal pAccDt As String, ByVal pAccSeq As String) As String
    
    SqlGetPtInfo = " SELECT a.ptid as ptid, b." & F_PTNM & " as ptnm, a.ageday , " & _
                   "        b." & F_SEX & " as sex, a.coldt , a.coltm , " & _
                   "        a.buildcd , d.field1 as buildnm," & _
                   "        a.wardid , a.roomid , a.bedid ," & _
                   "        a.deptcd ," & _
                   "        a.spccd , c.field3 as spcnm " & _
                   " FROM   " & T_LAB201 & " a," & _
                                T_HIS001 & " b," & _
                                T_LAB032 & " c," & _
                                T_LAB032 & " d " & _
                   " WHERE  a.ptid = b." & F_PTID & _
                   " AND  " & DBW("c.cdindex", LC3_Specimen, 2) & _
                   " AND    a.spccd = c.cdval1 " & _
                   " AND  " & DBJ(DBW("d.cdindex", LC3_Buildings, 2)) & _
                   " AND  " & DBJ("d.cdval1 =* a.buildcd") & _
                   " AND  " & DBW("a.workarea", pWorkArea, 2) & _
                   " AND  " & DBW("a.accdt", pAccDt, 2) & _
                   " AND  " & DBW("a.accseq", pAccSeq, 2)
End Function

Public Function SqlGetTmpData(ByVal sTestCd As String, Optional ByVal sEWSfg As String = "1") As String

    SqlGetTmpData = " SELECT a.testcd as TestCd, b.tpdata as TmpData" & _
                    " FROM   " & T_LAB001 & " a, " & T_LAB036 & " b " & _
                    " WHERE  " & DBW("a.testdiv", TST_SpeTest, 2) & _
                    " AND    " & DBW("a.testcd", sTestCd, 2) & _
                    " AND    a.rsttype = b.stcd " & _
                    " AND    " & DBW("b.wsfg", sEWSfg, 2)
                     
End Function
'-----------
'결과지 출력
'-----------
Public Function GetLisReportList(ByVal pLastDt As String, ByVal pVfyDt As String, ByVal pChkLoad As String, _
                                 ByVal pBussDiv As String, ByVal pWardId As String, ByVal pDoctId As String) As String
    '
    'pChkLoad : 출력Option & Col_Div & TestDiv
    
    Dim strWardWhere    As String
    Dim strChkLoad      As String
    Dim strTestDiv      As String
    Dim strQuery        As String
    
    strChkLoad = medGetP(pChkLoad, 1, COL_DIV)
    strTestDiv = medGetP(pChkLoad, 2, COL_DIV)
    
    If pBussDiv = enBussDiv.BussDiv_OutPatient Or pBussDiv = enBussDiv.BussDiv_HealthCaer Then
        '진료과조건
        If Trim(pWardId) <> CS_AllCaption Then
            strWardWhere = strWardWhere & " AND " & DBW("a.deptcd = ", pWardId)
        End If
    Else
        '병동
        '병동조건(병동이 바뀔수 있으므로처방난 병동을 기준으로 한다.)
        '결국 결과등록시의 병동으로 한다.
        If Trim(pWardId) <> CS_AllCaption Then
            strWardWhere = strWardWhere & " AND " & DBW("a.wardid  = ", pWardId)
        End If
    End If
    
    '주치의별 조건
    If Trim(pDoctId) <> CS_AllCaption Then
        strWardWhere = strWardWhere & " AND (" & DBW("a.majdoct = ", pDoctId) & " or " & DBW("a.orddoct=", pDoctId) & " )"
    End If

    GetLisReportList = " SELECT  a.deptcd,a.majdoct, '5' stscd, '0' as testdiv, " & _
                        "        c.testcd,c.rsttype,c.rstval,c.rstcd,c.rstunit,c.hldiv,c.dpdiv, " & _
                        "        c.vfydt,c.vfytm,c.vfyid,c.ptid,c.mfyfg, '' as mfydt, a.wardid,a.hosilid," & _
                        "        a.wardid " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.hosilid as location " & _
                        " FROM " & T_LAB032 & " d, " & T_LAB302 & " c, " & T_LAB201 & " a " & _
                        " WHERE  " & DBW("a.vfydt = ", pVfyDt) & _
                        " AND    " & DBW("a.stscd = ", enStsCd.StsCd_LIS_FinRst) & _
                        " AND    a.wardid is not null "
    GetLisReportList = GetLisReportList & _
                        " AND    " & DBW("d.cdindex", LC3_ReportTesctCd, 2) & _
                        " AND    c.testcd=d.cdval1 " & _
                        " AND    c.spccd=d.field2 " & _
                        " AND    c.vfydt = a.vfydt " & _
                        " AND    c.ptid  = a.ptid " & _
                        " AND  ( c.mfyfg = '0' or c.mfyfg is null ) " & strWardWhere
                        
    '일반검사 수정
    GetLisReportList = GetLisReportList & " UNION ALL " & _
                        " SELECT a.deptcd,a.majdoct, '6' stscd, '0' as testdiv, " & _
                        "        d.testcd,d.rsttype,d.rstval,d.rstcd,d.rstunit,d.hldiv,c.dpdiv, " & _
                        "        c.vfydt,c.vfytm,c.vfyid,c.ptid,'1' mfyfg, d.mfydt , a.wardid,a.hosilid," & _
                        "        a.wardid " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.hosilid as location " & _
                        "  FROM " & T_LAB032 & " e, " & T_LAB308 & " d, " & T_LAB302 & " c, " & T_LAB201 & " a " & _
                        "  WHERE " & DBW("a.vfydt = ", pVfyDt) & _
                        "  AND   " & DBW("a.stscd = ", enStsCd.StsCd_LIS_Modify) & _
                        "  AND    ( a.wardid is not null AND a.wardid <> ' ') "
    
    GetLisReportList = GetLisReportList & _
                        " AND    " & DBW("e.cdindex", LC3_ReportTesctCd, 2) & _
                        " AND    c.testcd=e.cdval1 " & _
                        " AND    c.spccd=e.field2 " & _
                        " AND    d.mfydt = a.vfydt " & _
                        " AND    c.ptid  = a.ptid " & _
                        " AND    c.workarea = d.workarea " & _
                        " AND    c.accdt = d.accdt " & _
                        " AND    c.accseq = d.accseq " & strWardWhere
   

    GetLisReportList = GetLisReportList & " order by hosilid,ptid, deptcd, testdiv, vfydt, vfytm "
    
End Function

Public Function GetTestReportList() As String
    
    GetTestReportList = " SELECT a.cdval1,a.field1, b.detailfg ,b.panelfg,b.rptseq " & _
                        "   FROM " & T_LAB032 & " a, " & T_LAB001 & " b " & _
                        "  WHERE " & DBW("cdindex", LC3_ReportTesctCd, 2) & _
                        "    AND a.cdval1=b.testcd " & _
                        "  order by rptseq,cdval1"
End Function

Public Function GetVfyTestCd(ByVal pTestCd As String, ByVal pRstCd As String) As String
    
    GetVfyTestCd = " SELECT field1,field4 " & _
                   "   FROM " & T_LAB031 & _
                   "  WHERE " & DBW("cdindex", LC2_ItemResult, 2) & _
                   "    AND " & DBW("cdval1", pTestCd, 2) & _
                   "    AND " & DBW("cdval2", pRstCd, 2)
End Function

'-- 2007.06.28 osw
Public Function GetAccLAbNoLIS201_New(ByVal pPtid As String, ByVal pVfyDt As String, _
                                      ByVal pVfyToDt As String) As String

    '일반검사 결과확인
    GetAccLAbNoLIS201_New = " SELECT distinct a.deptcd location, a.deptcd, a.wardid,a.roomid as hosilid, a.majdoct, a.ptid," & _
                        "     '" & enTestDiv.TST_RouTest & "' as testdiv, " & _
                        "     '" & enStsCd.StsCd_LIS_FinRst & "' stscd , c.vfydt, c.vfytm, c.vfyid, c.vfydt as mfydt, c.mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB302 & " c " & _
                        " WHERE  " & DBW("c.ptid  = ", pPtid) & _
                        " AND    c.vfydt BETWEEN " & DBS(pVfyDt) & " AND " & DBS(pVfyToDt)
    
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & _
                        " AND    c.mfyfg = '0' " & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    '일반검사 수정
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & " union " & _
                        " SELECT distinct  a.deptcd location, a.deptcd, a.wardid,a.roomid as hosilid, a.majdoct, a.ptid," & _
                        "     '" & enTestDiv.TST_RouTest & "' as testdiv, " & _
                        "     '" & enStsCd.StsCd_LIS_Modify & "' stscd, c.vfydt, c.vfytm, c.vfyid, d.mfydt as mfydt, '1' mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB302 & " c, " & T_LAB308 & " d " & _
                        " WHERE  d.mfydt BETWEEN " & DBS(pVfyDt) & " AND " & DBS(pVfyToDt) & _
                        " AND    " & DBW("d.ptid  = ", pPtid)
    
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & _
                        " AND    c.workarea = d.workarea " & _
                        " AND    c.accdt    = d.accdt " & _
                        " AND    c.accseq   = d.accseq " & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    
    '기타검사 결과확인/수정
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & " union " & _
                        " SELECT distinct a.deptcd location, a.deptcd, a.wardid, a.roomid as hosilid,a.majdoct, a.ptid," & _
                        "        '" & enTestDiv.TST_SpeTest & "' as testdiv, " & _
                        "        c.stscd, c.vfydt, c.vfytm, c.vfyid, c.vfydt as mfydt, " & FUNC_CONVERT("char", "c.mfyseq") & " mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB351 & " c " & _
                        " WHERE  c.vfydt BETWEEN " & DBS(pVfyDt) & " AND " & DBS(pVfyToDt) & _
                        " AND    " & DBW("c.ptid  = ", pPtid)
                        
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    
    '미생물검사 결과확인/수정
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & " union " & _
                        " SELECT distinct a.deptcd location, a.deptcd, a.wardid,a.roomid as hosilid, a.majdoct, a.ptid," & _
                        " '" & enTestDiv.TST_MicTest & "' as testdiv, " & _
                        "        c.stscd, c.vfydt, c.vfytm, c.vfyid, c.vfydt as mfydt, " & FUNC_CONVERT("char", "c.mfyseq") & " mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB404 & " c " & _
                        " WHERE  c.vfydt BETWEEN " & DBS(pVfyDt) & " AND " & DBS(pVfyToDt) & _
                        " AND    " & DBW("c.ptid  = ", pPtid)
                        
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    
    GetAccLAbNoLIS201_New = GetAccLAbNoLIS201_New & " order by deptcd, majdoct, ptid, testdiv, vfydt, vfytm "
        
End Function

Public Function GetAccLAbNoLIS201(ByVal pPtid As String, ByVal pVfyDt As String) As String

    '일반검사 결과확인
    GetAccLAbNoLIS201 = " SELECT distinct a.deptcd location, a.deptcd, a.wardid,a.roomid as hosilid, a.majdoct, a.ptid," & _
                        "     '" & enTestDiv.TST_RouTest & "' as testdiv, " & _
                        "     '" & enStsCd.StsCd_LIS_FinRst & "' stscd , c.vfydt, c.vfytm, c.vfyid, c.vfydt as mfydt, c.mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB302 & " c " & _
                        " WHERE  " & DBW("c.ptid  = ", pPtid) & _
                        " AND    " & DBW("c.vfydt = ", pVfyDt)
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & _
                        " AND    c.mfyfg = '0' " & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    '일반검사 수정
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & " union " & _
                        " SELECT distinct  a.deptcd location, a.deptcd, a.wardid,a.roomid as hosilid, a.majdoct, a.ptid," & _
                        "     '" & enTestDiv.TST_RouTest & "' as testdiv, " & _
                        "     '" & enStsCd.StsCd_LIS_Modify & "' stscd, c.vfydt, c.vfytm, c.vfyid, d.mfydt as mfydt, '1' mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB302 & " c, " & T_LAB308 & " d " & _
                        " WHERE  " & DBW("d.mfydt = ", pVfyDt) & _
                        " AND    " & DBW("d.ptid  = ", pPtid) '
    
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & _
                        " AND    c.workarea = d.workarea " & _
                        " AND    c.accdt    = d.accdt " & _
                        " AND    c.accseq   = d.accseq " & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    
    '기타검사 결과확인/수정
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & " union " & _
                        " SELECT distinct a.deptcd location, a.deptcd, a.wardid, a.roomid as hosilid,a.majdoct, a.ptid," & _
                        "        '" & enTestDiv.TST_SpeTest & "' as testdiv, " & _
                        "        c.stscd, c.vfydt, c.vfytm, c.vfyid, c.vfydt as mfydt, " & FUNC_CONVERT("char", "c.mfyseq") & " mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB351 & " c " & _
                        " WHERE  " & DBW("c.vfydt = ", pVfyDt) & _
                        " AND    " & DBW("c.ptid  = ", pPtid)
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    
    '미생물검사 결과확인/수정
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & " union " & _
                        " SELECT distinct a.deptcd location, a.deptcd, a.wardid,a.roomid as hosilid, a.majdoct, a.ptid," & _
                        " '" & enTestDiv.TST_MicTest & "' as testdiv, " & _
                        "        c.stscd, c.vfydt, c.vfytm, c.vfyid, c.vfydt as mfydt, " & FUNC_CONVERT("char", "c.mfyseq") & " mfyfg " & _
                        " FROM " & T_LAB201 & " a, " & T_LAB404 & " c " & _
                        " WHERE  " & DBW("c.vfydt = ", pVfyDt) & _
                        " AND    " & DBW("c.ptid  = ", pPtid)
                        
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & _
                        " AND    a.workarea = c.workarea " & _
                        " AND    a.accdt    = c.accdt " & _
                        " AND    a.accseq   = c.accseq "
    
    GetAccLAbNoLIS201 = GetAccLAbNoLIS201 & " order by deptcd, majdoct, ptid, testdiv, vfydt, vfytm "
        
End Function


Public Function UpdatePrtDt(ByVal pPrtDt As String, ByVal pPrtTm As String, ByVal pPrtId As String) As String
    '결과보고 일경우 최초 레포트 일자 ,결과 출력 정보 Update
    UpdatePrtDt = " update " & T_COM099 & _
                  " set " & _
                              DBW("remark", pPrtDt, 2) & _
                  " WHERE " & DBW("noindex", COM99_LIS_Report, 2)
End Function


Public Function UpdateFinalRptFg(ByVal pPtid As String, ByVal pStartDt As String, _
                                 ByVal pToDt As String, ByVal pLastDt As String, _
                                 ByVal pLastTm As String, ByVal pTestDiv As String) As String
    
    '결과보고 일경우 최초 레포트 일자 ,결과 출력 정보 Update
    Dim strTblNm As String
    
    Select Case pTestDiv
    Case "0": strTblNm = T_LAB302
    Case "1": strTblNm = T_LAB351
    Case "2": strTblNm = T_LAB404
    End Select
    
    UpdateFinalRptFg = " update " & strTblNm & _
                       " set " & _
                                   DBW("rptfg", "Y", 3) & _
                                   DBW("rptdt", Format(Now, CS_DateDbFormat), 3) & _
                                   DBW("rpttm", Format(Now, CS_TimeDbFormat), 3) & _
                                   DBW("rptid", ObjMyUser.EmpId, 2) & _
                       " WHERE " & DBW("ptid", pPtid, 2) & _
                       "   AND " & DBW("vfydt <= ", pToDt) & _
                       "   AND (rptfg = ''  or  rptfg is null) " & _
                       "   AND vfydt " & FUNC_CONCAT & " vfytm <= '" & pLastDt & pLastTm & "'"

                       '"   AND " & DBW("vfydt >= ", pStartDt)
End Function

Public Function UpdatePrintRptFg(ByVal pPtid As String, ByVal pDeptcd As String, _
                                 ByVal pToDt As String, ByVal pLastDt As String, _
                                 ByVal pLastTm As String, ByVal pTestDiv As String, _
                                 ByVal pBussDiv As String) As String
    
    '결과보고 일경우 최초 레포트 일자 ,결과 출력 정보 Update
    Dim strTblNm As String
    
    Select Case pTestDiv
    Case "0": strTblNm = T_LAB302
    Case "1": strTblNm = T_LAB351
    Case "2": strTblNm = T_LAB404
    End Select
    
    
    UpdatePrintRptFg = " update " & strTblNm & _
                       " set " & _
                                   DBW("rptfg", "Y", 3) & _
                                   DBW("rptdt", Format(Now, CS_DateDbFormat), 3) & _
                                   DBW("rpttm", Format(Now, CS_TimeDbFormat), 3) & _
                                   DBW("rptid", ObjMyUser.EmpId, 2) & _
                       " WHERE " & DBW("ptid", pPtid, 2) & _
                       "   AND " & DBW("vfydt <= ", pToDt) & _
                       "   AND (rptfg = ''  or  rptfg is null) " & _
                       "   AND vfydt " & FUNC_CONCAT & " vfytm <= '" & pLastDt & pLastTm & "'"
    If pTestDiv <> "0" Then
        UpdatePrintRptFg = UpdatePrintRptFg & _
                       "   AND " & DBW("stscd >= ", enStsCd.StsCd_LIS_FinRst)
    End If
    
    If pBussDiv = "1" Then  '외래
       ' UpdatePrintRptFg = UpdatePrintRptFg & _
                       "   AND exists (SELECT * FROM " & T_LAB201 & _
                                    "   WHERE workarea = " & strTblNm & ".workarea " & _
                                    "   AND   accdt    = " & strTblNm & ".accdt    " & _
                                    "   AND   accseq   = " & strTblNm & ".accseq   " & _
                                    "   AND   " & DBW("deptcd = ", pDeptcd) & ")"
    Else    '병동
        'UpdatePrintRptFg = UpdatePrintRptFg & _
                       "   AND exists (SELECT * FROM " & T_LAB201 & _
                                    "   WHERE workarea = " & strTblNm & ".workarea " & _
                                    "   AND   accdt    = " & strTblNm & ".accdt    " & _
                                    "   AND   accseq   = " & strTblNm & ".accseq   " & _
                                    "   AND   " & DBW("wardid = ", pDeptCd) & ")"
    End If

                       '"   AND " & DBW("vfydt >= ", pStartDt)
End Function


Public Function SqlGetESignInfo(ByVal pEmpId As String) As String

    SqlGetESignInfo = " SELECT * FROM " & T_LAB032 & _
                      " WHERE  " & DBW("cdindex", LC3_ElectronicSign, 2) & _
                      "   AND  " & DBW("cdval1", pEmpId, 2)

End Function

Public Function SqlGetLastReportDt() As String

    SqlGetLastReportDt = " SELECT remark as lastdt FROM " & T_COM099 & _
                         " WHERE  " & DBW("noindex", COM99_LIS_Report, 2) & _
                         "   AND  divcd1 = '0'  AND  divcd2 = '0'  AND  divcd3 = '0'"

End Function

'Public Function Get_EmpNm(ByVal EmpId As String) As String
''직원이름을 가지고온다.
'    Get_EmpNm = GetEmpName(EmpId)
'End Function



Public Function RiReportList(ByVal FrDT As String, ByVal toDt As String, ByVal Bussdiv As String, _
                             ByVal RptFg As String, Optional ByVal DeptCd As String = "", Optional ByVal DoctId As String = "")
    Dim sSQL As String
    Dim SDEPT As String
    Dim sDoct As String
    Dim SLOCATION As String
    
    SDEPT = "": sDoct = ""
    Select Case Bussdiv
        Case "1"
            SLOCATION = "DeptCd"
            If DeptCd <> "" Then SDEPT = " AND " & DBW("F.DEPTCD=", DeptCd)
        Case "2"
            SLOCATION = "WardId"
            If DeptCd <> "" Then SDEPT = " AND " & DBW("F.WARDID=", DeptCd)
    End Select
    
    If DoctId <> "" Then sDoct = " AND " & DBW("F.MAJDOCT=", DoctId)
    
    sSQL = " SELECT DISTINCT A." & SLOCATION & " AS LOCATION ,A.DEPTCD ,A.MAJDOCT ,A.PTID, B.NAME AS PTNM," & _
                            F_DOB2("B") & " as dob, " & F_SEX2("B") & " as sex, '0' as testdiv," & _
          "                '5' stscd , a.vfydt, a.vfytm, a.vfyid, '' as mfydt, '' mfyfg" & _
          " FROM " & T_HIS001 & " B," & T_LAB201 & " A," & T_LAB302 & " C" & _
          " WHERE " & DBW("A.VFYDT>=", FrDT) & _
          " AND " & DBW("A.VFYDT<=", toDt) & _
          " AND EXISTS (SELECT * FROM S2ORD102 E,S2ORD101 F" & _
          "             WHERE " & DBW("C.VFYDT>=", FrDT) & _
          "               AND " & DBW("C.VFYDT<=", toDt) & _
          "               AND C.WORKAREA=A.WORKAREA " & _
          "               AND C.ACCDT=A.ACCDT " & _
          "               AND C.ACCSEQ=A.ACCSEQ" & _
          "               AND (C.MFYFG='0' OR C.MFYFG IS NULL)"
          
    '재출력/출력 대상조회
    If RptFg <> "" Then
        sSQL = sSQL & " AND " & DBW("C.RPTFG=", "Y")
    Else
        sSQL = sSQL & " AND (C.RPTFG='' OR C.RPTFG IS NULL)"
    End If
    
    If SDEPT <> "" Then sSQL = sSQL & SDEPT
    If sDoct <> "" Then sSQL = sSQL & sDoct
    
    sSQL = sSQL & "               AND C.WORKAREA IN('RI','OR') " & _
                  "               AND A.WORKAREA=E.WORKAREA " & _
                  "               AND A.ACCDT=E.ACCDT " & _
                  "               AND A.ACCSEQ=E.ACCSEQ" & _
                  "               AND E.PTID=F.PTID " & _
                  "               AND E.ORDDT=F.ORDDT " & _
                  "               AND E.ORDNO=F.ORDNO " & _
                  "               AND " & DBW("F.BUSSDIV=", Bussdiv) & ")" & _
                  " AND A.PTID=B." & F_PTID
    
    sSQL = sSQL & " UNION"
    
    sSQL = sSQL & " SELECT DISTINCT A." & SLOCATION & "  LOCATION ,A.DEPTCD ,A.MAJDOCT ,A.PTID, B.NAME AS PTNM," & _
                            F_DOB2("B") & " as dob, " & F_SEX2("B") & " as sex, '0' as testdiv," & _
          "                '6' stscd , a.vfydt, a.vfytm, a.vfyid, '' as mfydt, '' mfyfg" & _
          " FROM TB_IDMAST B,S2LAB201 A,S2LAB302 C" & _
          " WHERE " & DBW("A.VFYDT>=", FrDT) & _
          " AND " & DBW("A.VFYDT<=", toDt) & _
          " AND EXISTS (SELECT * FROM S2ORD102 E,S2ORD101 F,S2LAB308 G" & _
          "             WHERE " & DBW("C.VFYDT>=", FrDT) & _
          "             AND " & DBW("C.VFYDT<=", toDt) & _
          "             AND C.WORKAREA=A.WORKAREA " & _
          "             AND C.ACCDT=A.ACCDT " & _
          "             AND C.ACCSEQ=A.ACCSEQ" & _
          "             AND C.WORKAREA=G.WORKAREA" & _
          "             AND C.ACCDT=G.ACCDT" & _
          "             AND C.ACCSEQ=G.ACCSEQ" & _
          "             AND (C.MFYFG='0' OR C.MFYFG IS NULL)" '& _
          "             AND (C.RPTFG='' OR C.RPTFG IS NULL)"
    '재출력/출력 대상조회
    If RptFg <> "" Then
        sSQL = sSQL & " AND " & DBW("C.RPTFG=", "Y")
    Else
        sSQL = sSQL & " AND (C.RPTFG='' OR C.RPTFG IS NULL)"
    End If
    '진료과/병동
    If SDEPT <> "" Then sSQL = sSQL & SDEPT
    '주치의별
    If sDoct <> "" Then sSQL = sSQL & sDoct
    
    sSQL = sSQL & _
          " AND C.WORKAREA IN('RI','OR') " & _
          " AND A.WORKAREA=E.WORKAREA" & _
          " AND A.ACCDT=E.ACCDT" & _
          " AND A.ACCSEQ=E.ACCSEQ" & _
          " AND E.PTID=F.PTID" & _
          " AND E.ORDDT=F.ORDDT" & _
          " AND E.ORDNO=F.ORDNO" & _
          " AND " & DBW("F.BUSSDIV=", Bussdiv) & " )" & _
          " AND A.PTID=B." & F_PTID
      
    RiReportList = sSQL
End Function

'임상병리 출력을 위한 조회리스트

Public Function LABReportList(ByVal FrDT As String, ByVal toDt As String, ByVal Bussdiv As String, _
                             ByVal RptFg As String, ByVal TestDiv As String, Optional ByVal DeptCd As String = "", Optional ByVal DoctId As String = "")
    Dim sSQL As String
    Dim SDEPT As String
    Dim sDoct As String
    Dim SLOCATION As String
    Dim sQuery    As String
    
    

    '성바오로병원에서 이렇게 해봤당.
    SDEPT = "": sDoct = ""
    Select Case Bussdiv
        Case "1"
            SLOCATION = "deptcd"
            If DeptCd <> "" Then SDEPT = " AND " & DBW("a.deptcd=", DeptCd)
            If DoctId <> "" Then sDoct = " AND " & DBW("a.majdoct=", DoctId)
        Case "2"
            SLOCATION = "wardid"
            If DeptCd <> "" Then SDEPT = " AND " & DBW("a.wardid=", DeptCd)
            If DoctId <> "" Then sDoct = " AND " & DBW("a.majdoct=", DoctId)
    End Select
    
    If TestDiv = "3" Or TestDiv = "0" Then
        '일반검사(수정 말구)
        sSQL = " SELECT DISTINCT a." & SLOCATION & " AS LOCATION ,a.deptcd ,a.roomid as HOSILID,a.majdoct ,a.ptid, " & _
               "     '0' as testdiv," & _
               "     '5' stscd , a.vfydt, a.vfytm, a.vfyid, '' as mfydt, '' mfyfg" & _
               " FROM " & T_LAB302 & " c," & T_LAB201 & " a" & _
               " WHERE " & DBW("a.vfydt>=", FrDT) & _
               " AND " & DBW("a.vfydt<=", toDt)
        
        sSQL = sSQL & "  AND c.workarea=a.workarea " & _
                      "  AND c.accdt=a.accdt " & _
                      "  AND c.accseq=a.accseq" & _
                      "  AND (c.mfyfg='0' OR c.mfyfg IS NULL)"
        
        '** 예수병원에서는 Remark 처리 한다.
        '=======================================================================================
'        Select Case Bussdiv
'            Case "1":
'                      sSQL = sSQL & " AND (a.wardid is null or a.wardid='HC' or a.wardid='HDR') "
'            Case "2":
'                      sSQL = sSQL & " AND a.wardid not in (' ' ,'HC','HDR') "
'        End Select
        '=======================================================================================
        
        If SDEPT <> "" Then sSQL = sSQL & SDEPT
        If sDoct <> "" Then sSQL = sSQL & sDoct
        
        '재출력/출력 대상조회
        If RptFg <> "" Then
            sSQL = sSQL & " AND " & DBW("c.rptfg=", "Y") & " AND " & DBW("c.rptdt=", toDt)
        Else
            sSQL = sSQL & " AND (c.rptfg=' ' OR c.rptfg IS NULL)"
        End If


        '일반검사 수정
        sSQL = sSQL & " UNION"
        
        sSQL = sSQL & " SELECT DISTINCT a." & SLOCATION & "  LOCATION ,a.deptcd ,a.roomid as HOSILID,a.majdoct ,a.ptid, " & _
                      "                '0' as testdiv," & _
                      "                '6' stscd , a.vfydt, a.vfytm, a.vfyid,  d.mfydt,  c.mfyfg" & _
                      " FROM " & T_LAB308 & " d," & T_LAB302 & " c," & T_LAB201 & " a" & _
                      " WHERE " & DBW("a.vfydt>=", FrDT) & _
                      " AND " & DBW("a.vfydt<=", toDt)
              
        sSQL = sSQL & "  AND c.workarea=a.workarea " & _
                      "  AND c.accdt=a.accdt " & _
                      "  AND c.accseq=a.accseq AND c.workarea=d.workarea AND c.accdt=d.accdt AND c.accseq=d.accseq AND c.testcd=d.testcd " & _
                      "  AND (c.mfyfg='1')"
                      
        '** 예수병원에서는 Remark 처리 한다.
        '=======================================================================================
'        Select Case Bussdiv
'            Case "1":
'                      sSQL = sSQL & " AND (a.wardid is null or a.wardid='HC' or a.wardid='HDR') "
'            Case "2":
'                      sSQL = sSQL & " AND a.wardid not in (' ' ,'HC','HDR') "
'        End Select
        '=======================================================================================
        
        '진료과/병동
        If SDEPT <> "" Then sSQL = sSQL & SDEPT
        '의사별
        If sDoct <> "" Then sSQL = sSQL & sDoct
        '재출력/출력 대상조회
        If RptFg <> "" Then
            sSQL = sSQL & " AND " & DBW("c.rptfg=", "Y") & " AND " & DBW("c.rptdt=", toDt)
        Else
            sSQL = sSQL & " AND (c.rptfg=' ' OR c.rptfg IS NULL)"
        End If
    End If
    
    '기타검사
    If TestDiv = "3" Or TestDiv = "1" Then
        If sSQL = "" Then
            sSQL = sSQL
        Else
            sSQL = sSQL & " UNION "
        End If
        
        sSQL = sSQL & " SELECT DISTINCT a." & SLOCATION & "  LOCATION ,a.deptcd ,a.roomid as HOSILID,a.majdoct ,a.ptid," & _
                      "                '1' as testdiv," & _
                      "                a.stscd , a.vfydt, a.vfytm, a.vfyid, '' as mfydt, " & FUNC_CONVERT("char", "c.mfyseq") & " mfyfg " & _
                      " FROM " & T_LAB351 & " c," & T_LAB201 & " a" & _
                      " WHERE " & DBW("a.vfydt>=", FrDT) & _
                      " AND " & DBW("a.vfydt<=", toDt)
              
        sSQL = sSQL & "  AND c.workarea=a.workarea " & _
                      "  AND c.accdt=a.accdt " & _
                      "  AND c.accseq=a.accseq" '& _
                      "  AND (c.mfyfg='0' OR c.mfyfg IS NULL)"
        
        '** 예수병원에서는 Remark 처리 한다.
        '=======================================================================================
'        Select Case Bussdiv
'            Case "1":
'                      sSQL = sSQL & " AND (a.wardid is null or a.wardid='HC' or a.wardid='HDR') "
'            Case "2":
'                      sSQL = sSQL & " AND a.wardid not in (' ' ,'HC','HDR') "
'        End Select
        '=======================================================================================
              
        '진료과/병동
        If SDEPT <> "" Then sSQL = sSQL & SDEPT
        
        '의사별
        If sDoct <> "" Then sSQL = sSQL & sDoct

        '재출력/출력 대상조회
        If RptFg <> "" Then
            sSQL = sSQL & " AND " & DBW("c.rptfg=", "Y") & " AND " & DBW("c.rptdt=", toDt)
        Else
            sSQL = sSQL & " AND (c.rptfg=' ' OR c.rptfg IS NULL)"
        End If
    End If
    '미생물 결과 혹은 수정
    If TestDiv = "3" Or TestDiv = "2" Then
    
        If sSQL = "" Then
            sSQL = sSQL
        Else
            sSQL = sSQL & " UNION "
        End If
        sSQL = sSQL & " SELECT DISTINCT a." & SLOCATION & "  LOCATION ,a.deptcd ,a.roomid as HOSILID,a.majdoct ,a.ptid, " & _
                      "     '2' as testdiv," & _
                      "     c.stscd , a.vfydt, a.vfytm, a.vfyid, '' as mfydt,  " & FUNC_CONVERT("char", "c.mfyseq") & " mfyfg " & _
                      " FROM " & T_LAB404 & " c," & T_LAB201 & " a" & _
                      " WHERE " & DBW("a.vfydt>=", FrDT) & _
                      " AND " & DBW("a.vfydt<=", toDt) '
              
        sSQL = sSQL & "  AND c.workarea=a.workarea " & _
                      "  AND c.accdt=a.accdt " & _
                      "  AND c.accseq=a.accseq" & _
                      "  AND c.stscd>='4'"
                      
                      '& _
                      "  AND (c.mfyfg='0' OR c.mfyfg IS NULL)"
        
        '** 예수병원에서는 Remark 처리 한다.
        '=======================================================================================
'        Select Case Bussdiv
'            Case "1":
'                      sSQL = sSQL & " AND (a.wardid is null or a.wardid='HC' or a.wardid='HDR') "
'            Case "2":
'                      sSQL = sSQL & " AND a.wardid not in (' ' ,'HC','HDR') "
'        End Select
        '=======================================================================================
              
        '진료과/병동
        If SDEPT <> "" Then sSQL = sSQL & SDEPT
        
        '의사별
        If sDoct <> "" Then sSQL = sSQL & sDoct
              
        '재출력/출력 대상조회
        If RptFg <> "" Then
            sSQL = sSQL & " AND " & DBW("c.rptfg=", "Y") & " AND " & DBW("c.rptdt=", toDt)
        Else
            sSQL = sSQL & " AND (c.rptfg=' ' OR c.rptfg IS NULL)"
        End If
    End If

    LABReportList = sSQL & " order by LOCATION,ptid" ', deptcd, testdiv, vfydt, vfytm "
 
    
End Function

