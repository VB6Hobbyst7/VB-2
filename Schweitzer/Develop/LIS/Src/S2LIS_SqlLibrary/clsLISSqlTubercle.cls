VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlTubercle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------
'항산성균 결과등록 SQL
'결핵균 검사 결과서
'레포트, 조회
'--------------------------

Public Function SQLGetSenDt(ByVal pLoadSeq As String, Optional ByVal pSendDt As String) As String

    Select Case pLoadSeq
        Case "1":
            SQLGetSenDt = " SELECT distinct senddt FROM " & T_LAB205 & _
                          "  WHERE " & DBW("testcd", P_AFBSENSCD, 2) & _
                          "    AND senddt is not null " & _
                          "  order by senddt desc"
        Case "2":
            SQLGetSenDt = " SELECT * FROM " & T_LAB205 & _
                          "  WHERE " & DBW("testcd", P_AFBSENSCD, 2) & _
                          "    AND " & DBW("senddt", pSendDt, 2) & _
                          "  order by workarea,accdt,accseq"
    End Select

End Function


Public Function SQLGetTBTestCd() As String

   
    SQLGetTBTestCd = " SELECT * FROM " & T_LAB032 & _
                     "  WHERE " & DBW("cdindex", LC3_TUBERCLE, 2) & _
                     " order by cdval1 "

End Function


Public Function SQLAFPSensHeaderNOLoad(ByVal pField As String, ByVal pNumber As String) As String

   
    SQLAFPSensHeaderNOLoad = " SELECT * FROM " & T_LAB360 & _
                             "  WHERE " & DBW(pField, pNumber, 2)

End Function

Public Function SQLAFPSensHeaderLoad(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

   
    SQLAFPSensHeaderLoad = " SELECT * FROM " & T_LAB360 & _
                           "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                           "    AND " & DBW("accdt", pAccDt, 2) & _
                           "    AND " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SQLAFPSensBodyLoad(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

   
    SQLAFPSensBodyLoad = " SELECT * FROM " & T_LAB361 & _
                         "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                         "    AND " & DBW("accdt", pAccDt, 2) & _
                         "    AND " & DBW("accseq", pAccSeq, 2) & _
                         "  order by drugcd,seq"
                         

End Function



Public Function SQLAFPCultureRstLoad(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

   
    SQLAFPCultureRstLoad = " SELECT b.rstcd,a.field1 AS rstnm " & _
                           "   FROM " & T_LAB031 & " a," & T_LAB302 & " b " & _
                           "  WHERE " & DBW("b.workarea", pWorkArea, 2) & _
                           "    AND " & DBW("b.accdt", pAccDt, 2) & _
                           "    AND " & DBW("b.accseq", pAccSeq, 2) & _
                           "    AND " & DBW("b.testcd", P_AFPCULTCD, 2) & _
                           "    AND " & DBW("a.cdindex", LC2_ItemResult, 2) & _
                           "    AND b.testcd=a.cdval1 AND b.rstcd = a.cdval2 "

End Function



Public Function SQLGetPtInformation(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

   
    SQLGetPtInformation = " SELECT c.ptid,c.orddoct,c.majdoct,c.coldt,c.coltm,c.colid,c.rcvdt,c.rcvtm,c.rcvid,c.vfydt,c.vfytm,c.vfyid, " & _
                          "        c.spccd,c.deptcd,c.wardid,c.hosilid,c.bedid, a.orddt,a.ordtm,d.field3 as spcnm   " & _
                          "   FROM " & T_LAB032 & " d ," & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB201 & " c  " & _
                          "  WHERE " & DBW("c.workarea", pWorkArea, 2) & _
                          "    AND " & DBW("c.accdt", pAccDt, 2) & _
                          "    AND " & DBW("c.accseq", pAccSeq, 2) & _
                          "    AND " & DBW("d.cdindex", LC3_Specimen, 2) & _
                          "    AND c.workarea=b.workarea AND c.accdt=b.accdt AND c.accseq=b.accseq " & _
                          "    AND b.ptid = a.ptid AND b.orddt = a.orddt  AND b.ordno = a.ordno " & _
                          "    AND c.spccd=d.cdval1"

End Function


Public Function SQLInsertAFBHeader(ByVal strWorkArea As String, ByVal strAccDt As String, ByVal strAccSeq As String, _
                                   ByVal strTBNo As String, ByVal strRcvDt As String, ByVal strBacNo As String, _
                                   ByVal strBacRstCd As String, ByVal strDILaw As String, ByVal strDGrow As String, _
                                   ByVal strRALaw As String, ByVal strRGrow As String, ByVal strRemark As String, _
                                   ByVal strVfyDt As String, ByVal strVfyTm As String, ByVal strVfyId As String) As String

   
    SQLInsertAFBHeader = " insert into " & T_LAB360 & "(workarea,accdt,accseq,tbno,rcvdt,bacno, " & _
                                                      "bacrstcd,dilaw,dgrow,ralaw,rgrow,remark,vfydt,vfytm,vfyid) " & _
                         " values(" & DBV("workarea", strWorkArea, 1) & DBV("accdt", strAccDt, 1) & DBV("accseq", strAccSeq, 1) & _
                                      DBV("tbno", strTBNo, 1) & DBV("rcvdt", strRcvDt, 1) & DBV("bacno", strBacNo, 1) & _
                                      DBV("bacrstcd", strBacRstCd, 1) & DBV("dilaw", strDILaw, 1) & DBV("dgrow", strDGrow, 1) & _
                                      DBV("ralaw", strRALaw, 1) & DBV("rgrow", strRGrow, 1) & DBV("remark", strRemark, 1) & _
                                      DBV("vfydt", strVfyDt, 1) & DBV("vfytm", strVfyTm, 1) & DBV("vfyid", strVfyId) & ") "
    


End Function


Public Function SQLInsertAFBBody(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                 ByVal pDrugCd As String, ByVal pSeq As String, ByVal pRstValue As String) As String

   
    SQLInsertAFBBody = " insert into " & T_LAB361 & "(workarea,accdt,accseq,drugcd,seq,rstvalue) " & _
                       " values(" & DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                                    DBV("drugcd", pDrugCd, 1) & DBV("seq", pSeq, 1) & DBV("rstvalue", pRstValue) & ") "

End Function

Public Function SQLGetAccInfor(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

   
    SQLGetAccInfor = " SELECT reqtotcnt, reqinputcnt FROM " & T_LAB201 & _
                     "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                     "    AND " & DBW("accdt", pAccDt, 2) & _
                     "    AND " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SQLAccInforUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pReqInputCnt As Integer, _
                                  ByVal pVfyDt As String, ByVal pVfyTm As String, _
                                  ByVal pVfyId As String, Optional ByVal chkModify As Boolean = False) As String
    If chkModify = False Then
        SQLAccInforUpdate = " update " & T_LAB201 & _
                            "    set " & DBW("reqinputcnt", pReqInputCnt, 3) & _
                                         DBW("vfydt", pVfyDt, 3) & _
                                         DBW("vfytm", pVfyTm, 3) & _
                                         DBW("vfyid", pVfyId, 3) & _
                                         DBW("stscd", StsCd_LIS_FinRst, 2) & _
                            "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                            "    AND " & DBW("accdt", pAccDt, 2) & _
                            "    AND " & DBW("accseq", pAccSeq, 2)
    Else
        SQLAccInforUpdate = " update " & T_LAB201 & _
                            "    set " & DBW("stscd", StsCd_LIS_Modify, 2) & _
                            "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                            "    AND " & DBW("accdt", pAccDt, 2) & _
                            "    AND " & DBW("accseq", pAccSeq, 2)
    End If

End Function

Public Function SQLOrderInforUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, ByVal pVfyDt As String, _
                                    ByVal pVfyTm As String, ByVal pVfyId As String, Optional ByVal chkModify As Boolean = False) As String
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBS(pAccSeq)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strWrkDiv = "" & rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    rs.Close: Set rs = Nothing
    
    If chkModify = False Then
        '** 원본 ==============================================================================
'        SQLOrderInforUpdate = " update " & T_LAB102 & _
'                              "    set " & DBW("examdt", pVfyDt, 3) & _
'                                           DBW("examtm", pVfyTm, 3) & _
'                                           DBW("examdoct", pVfyId, 3) & _
'                                           DBW("stscd", StsCd_LIS_FinRst, 2) & _
'                              "  WHERE " & DBW("workarea", pWorkArea, 2) & _
'                              "    AND " & DBW("accdt", pAccDt, 2) & _
'                              "    AND " & DBW("accseq", pAccSeq, 2)
        '=======================================================================================
        
        '** 전주예수병원 추가 루틴 =============================================================
        If strWrkDiv = "3" Then         '종건
            SQLOrderInforUpdate = " update su2examt " & _
                                  "    set " & _
                                  " examdate = TO_DATE(" & DBS(pVfyDt & pVfyTm) & ", 'yyyymmdd hh24:mi:ss') " & _
                                  ", examdr = " & DBS(pVfyId) & _
                                  ", " & DBW("stscd", StsCd_LIS_FinRst, 2) & _
                                  "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                                  "    AND " & DBW("accdt", pAccDt, 2) & _
                                  "    AND " & DBW("accseq", pAccSeq, 2)
        ElseIf strWrkDiv = "4" Then     '일건
                SQLOrderInforUpdate = " update sg2examt " & _
                                      "    set " & _
                                      " examdate = TO_DATE(" & DBS(pVfyDt & pVfyTm) & ", 'yyyymmdd hh24:mi:ss') " & _
                                      ", examdr = " & DBS(pVfyId) & _
                                      ", " & DBW("stscd", StsCd_LIS_FinRst, 2) & _
                                      "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                                      "    AND " & DBW("accdt", pAccDt, 2) & _
                                      "    AND " & DBW("accseq", pAccSeq, 2)
        Else
            SQLOrderInforUpdate = " update mdexmort " & _
                                  "    set " & _
                                  " rsltdate = TO_DATE(" & DBS(pVfyDt & pVfyTm) & ", 'yyyymmdd hh24:mi:ss') " & _
                                  ", cofmdr = " & DBS(pVfyId) & _
                                  ", " & DBW("stscd", StsCd_LIS_FinRst, 2) & _
                                  "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                                  "    AND " & DBW("accdt", pAccDt, 2) & _
                                  "    AND " & DBW("accseq", pAccSeq, 2)
        End If
        '=======================================================================================
        
    Else
        '** 원본 ==============================================================================
'        SQLOrderInforUpdate = " update " & T_LAB102 & _
'                              "    set " & DBW("stscd", StsCd_LIS_Modify, 2) & _
'                              "  WHERE " & DBW("workarea", pWorkArea, 2) & _
'                              "    AND " & DBW("accdt", pAccDt, 2) & _
'                              "    AND " & DBW("accseq", pAccSeq, 2) & _
'                              "    AND " & DBW("ordcd", P_AFBSENSCD, 2)
        '=======================================================================================
        
        '** 전주예수병원 추가 루틴 =============================================================
        If strWrkDiv = "3" Then         '종건
            SQLOrderInforUpdate = " update su2examt " & _
                                  "    set " & DBW("stscd", StsCd_LIS_Modify, 2) & _
                                  "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                                  "    AND " & DBW("accdt", pAccDt, 2) & _
                                  "    AND " & DBW("accseq", pAccSeq, 2) & _
                                  "    AND " & DBW("ordcd", P_AFBSENSCD, 2)
        ElseIf strWrkDiv = "4" Then     '일건
                SQLOrderInforUpdate = " update sg2examt " & _
                                      "    set " & DBW("stscd", StsCd_LIS_Modify, 2) & _
                                      "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                                      "    AND " & DBW("accdt", pAccDt, 2) & _
                                      "    AND " & DBW("accseq", pAccSeq, 2) & _
                                      "    AND " & DBW("ordcd", P_AFBSENSCD, 2)
        Else
            SQLOrderInforUpdate = " update mdexmort " & _
                                  "    set " & DBW("stscd", StsCd_LIS_Modify, 2) & _
                                  "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                                  "    AND " & DBW("accdt", pAccDt, 2) & _
                                  "    AND " & DBW("accseq", pAccSeq, 2) & _
                                  "    AND " & DBW("ordcd", P_AFBSENSCD, 2)
        End If
        '=======================================================================================
        
    End If

End Function

Public Function SQLOutLabUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String) As String

   
    SQLOutLabUpdate = " update " & T_LAB205 & _
                      "    set " & DBW("stscd", StsCd_LIS_FinRst, 2) & _
                      "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                      "    AND " & DBW("accdt", pAccDt, 2) & _
                      "    AND " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SQLVfyInforUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pVfyDt As String, _
                                  ByVal pVfyTm As String, ByVal pVfyId As String, ByVal pValue As String, _
                                  Optional pChkModify As Boolean = False) As String
    

    SQLVfyInforUpdate = " update " & T_LAB302 & _
                        "    set " & DBW("rstcd", pValue, 3) & _
                                     DBW("vfydt", pVfyDt, 3) & _
                                     DBW("vfytm", pVfyTm, 3) & _
                                     DBW("vfyid", pVfyId, 2) & _
                        "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                        "    AND " & DBW("accdt", pAccDt, 2) & _
                        "    AND " & DBW("accseq", pAccSeq, 2) & _
                        "    AND " & DBW("testcd", P_AFBSENSCD, 2)
End Function


Public Function SQLUpdateAFBHeader(ByVal strWorkArea As String, ByVal strAccDt As String, ByVal strAccSeq As String, _
                                   ByVal strTBNo As String, ByVal strBacNo As String, _
                                   ByVal strBacRstCd As String, ByVal strDILaw As String, ByVal strDGrow As String, _
                                   ByVal strRALaw As String, ByVal strRGrow As String, ByVal strRemark As String, _
                                   ByVal strMfyDt As String, ByVal strMfyTm As String, ByVal strMfyId As String) As String

   
    SQLUpdateAFBHeader = " update " & T_LAB360 & _
                         "    set " & DBW("tbno", strTBNo, 3) & _
                                      DBW("bacno", strBacNo, 3) & _
                                      DBW("bacrstcd", strBacRstCd, 3) & _
                                      DBW("dilaw", strDILaw, 3) & _
                                      DBW("dgrow", strDGrow, 3) & _
                                      DBW("ralaw", strRALaw, 3) & _
                                      DBW("rgrow", strRGrow, 3) & _
                                      DBW("remark", strRemark, 3) & _
                                      DBW("mfyfg", "1", 3) & _
                                      DBW("mfydt", strMfyDt, 3) & _
                                      DBW("mfytm", strMfyTm, 3) & _
                                      DBW("mfyid", strMfyId, 2) & _
                         "  WHERE " & DBW("workarea", strWorkArea, 2) & _
                         "    AND " & DBW("accdt", strAccDt, 2) & _
                         "    AND " & DBW("accseq", strAccSeq, 2)
    
    
    
End Function

Public Function SQLUpdateRptAFBHeader(ByVal strWorkArea As String, ByVal strAccDt As String, ByVal strAccSeq As String, _
                                   ByVal strRptDt As String, ByVal strRptTm As String, ByVal strRptId As String) As String

   
    SQLUpdateRptAFBHeader = " update " & T_LAB360 & _
                            "    set " & DBW("rptfg", "1", 3) & _
                                         DBW("rptdt", strRptDt, 3) & _
                                         DBW("rpttm", strRptTm, 3) & _
                                         DBW("rptid", strRptId, 2) & _
                            "  WHERE " & DBW("workarea", strWorkArea, 2) & _
                            "    AND " & DBW("accdt", strAccDt, 2) & _
                            "    AND " & DBW("accseq", strAccSeq, 2)
    
    
End Function


Public Function SQLUpdateAFBBody(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                 ByVal pDrugCd As String, ByVal pSeq As String, ByVal pRstValue As String) As String

   
    SQLUpdateAFBBody = " update " & T_LAB361 & _
                       "    set " & DBW("rstvalue", pRstValue, 2) & _
                       "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                       "    AND " & DBW("accdt", pAccDt, 2) & _
                       "    AND " & DBW("accseq", pAccSeq, 2) & _
                       "    AND " & DBW("drugcd", pDrugCd, 2) & _
                       "    AND " & DBW("seq", pSeq, 2)
                       
                       

End Function


'------------------------------
' 항산성균 감수성 결과 일괄출력
'------------------------------

Public Function SQLGetAFBSensReport(ByVal pVfyDt As String, ByVal pLastDay As String, ByVal pWorkArea As String, _
                                    ByVal pBussDiv As String, Optional pField As String = "", _
                                    Optional pDeptcd As String = "", Optional pMajDoct As String = "", Optional pChkSql As String = "1") As String
' SELECT c.wardid,c.deptcd,b.ptid,a.vfydt,a.vfytm,
'b.rcvdt,b.rcvtm,c.majdoct,a.bacrstcd,a.dilaw,a.dgrow,
' a.ralaw , a.rgrow, a.Remark
'FROM S2ORD101 c,S2LAB360 a,S2ORD102 b
'WHERE  b.examdt >= '20020826'    AND  b.examdt <= '20020828'    AND  b.workarea  ='51'
'AND  b.stscd > '4'    AND  b.ordcd  ='B4064'    AND c.bussdiv  ='1'
'AND b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq
' AND b.ptid = c.ptid AND b.orddt =c.orddt AND b.ordno = c.ordno     AND a.rptfg IS NULL
    Select Case pChkSql
        Case "1":
            SQLGetAFBSensReport = " SELECT c.wardid,c.deptcd,b.ptid,a.vfydt,a.vfytm," & _
                                  "        b.stscd,b.rcvdt,b.rcvtm,c.majdoct,a.bacrstcd,a.dilaw, " & _
                                  "        a.dgrow,a.ralaw,a.rgrow,a.remark,a.workarea,a.accdt,a.accseq " & _
                                  "   FROM " & T_LAB101 & " c," & T_LAB360 & " a," & T_LAB102 & " b " & _
                                  "  WHERE " & DBW("b.examdt >= ", pLastDay) & _
                                  "    AND " & DBW("b.examdt <= ", pVfyDt) & _
                                  "    AND " & DBW("b.workarea", pWorkArea, 2) & _
                                  "    AND " & DBW("b.stscd > ", StsCd_LIS_MidRst) & _
                                  "    AND " & DBW("b.ordcd", P_AFBSENSCD, 2) & _
                                  "    AND " & DBW("c.bussdiv", pBussDiv, 2) & _
                                  "    AND b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                                  "    AND b.ptid = c.ptid AND b.orddt =c.orddt AND b.ordno = c.ordno " & _
                                  "    AND a.rptfg is null "

        Case "2":
            SQLGetAFBSensReport = " SELECT c.wardid,c.deptcd,b.ptid,a.vfydt,a.vfytm," & _
                                  "        b.stscd,b.rcvdt,b.rcvtm,c.majdoct,a.bacrstcd,a.dilaw, " & _
                                  "        a.dgrow,a.ralaw,a.rgrow,a.remark,a.workarea,a.accdt,a.accseq " & _
                                  "   FROM " & T_LAB101 & " c," & T_LAB360 & " a," & T_LAB102 & " b " & _
                                  "  WHERE " & DBW("b.examdt >= ", pLastDay) & _
                                  "    AND " & DBW("b.examdt <= ", pVfyDt) & _
                                  "    AND " & DBW("b.workarea", pWorkArea, 2) & _
                                  "    AND " & DBW("b.stscd > ", StsCd_LIS_MidRst) & _
                                  "    AND " & DBW("b.ordcd", P_AFBSENSCD, 2) & _
                                  "    AND " & DBW("c.bussdiv", pBussDiv, 2) & _
                                  "    AND b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                                  "    AND b.ptid = c.ptid AND b.orddt =c.orddt AND b.ordno = c.ordno " & _
                                  "    AND " & DBW("a.rptfg", "1", 2)
       
        End Select
    
    
    If pDeptcd <> "" Then SQLGetAFBSensReport = SQLGetAFBSensReport & " AND " & DBW(pField, pDeptcd, 2)
    If pMajDoct <> "" Then SQLGetAFBSensReport = SQLGetAFBSensReport & " AND " & DBW("c.majdoct", pMajDoct, 2)
    
    
    SQLGetAFBSensReport = SQLGetAFBSensReport & " order by ptid,deptcd,vfydt,vfytm "
    
End Function



Public Function SQLGetAFBSensPtIdReport(ByVal pPtid As String, ByVal pVfyDt As String, _
                                        ByVal pWorkArea As String) As String
    
    SQLGetAFBSensPtIdReport = " SELECT b.wardid,b.deptcd,b.ptid,a.vfydt,a.vfytm," & _
                              "        b.stscd,b.rcvdt,b.rcvtm,b.majdoct,a.bacrstcd,a.dilaw, " & _
                              "        a.dgrow,a.ralaw,a.rgrow,a.remark,a.workarea,a.accdt,a.accseq " & _
                              "   FROM " & T_LAB360 & " a," & T_LAB201 & " b " & _
                              "  WHERE " & DBW("b.vfydt", pVfyDt, 2) & _
                              "    AND " & DBW("b.workarea", pWorkArea, 2) & _
                              "    AND " & DBW("b.stscd", StsCd_LIS_FinRst, 2) & _
                              "    AND " & DBW("b.ptid", pPtid, 2) & _
                              "    AND b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq "
                              

   SQLGetAFBSensPtIdReport = SQLGetAFBSensPtIdReport & " UNION ALL " & _
                              " SELECT b.wardid,b.deptcd,b.ptid,a.vfydt,a.vfytm," & _
                              "        b.stscd,b.rcvdt,b.rcvtm,b.majdoct,a.bacrstcd,a.dilaw, " & _
                              "        a.dgrow,a.ralaw,a.rgrow,a.remark,a.workarea,a.accdt,a.accseq " & _
                              "   FROM " & T_LAB360 & " a," & T_LAB201 & " b " & _
                              "  WHERE " & DBW("b.workarea", pWorkArea, 2) & _
                              "    AND " & DBW("b.stscd = ", StsCd_LIS_Modify) & _
                              "    AND " & DBW("b.ptid", pPtid, 2) & _
                              "    AND " & DBW("a.mfydt", pVfyDt, 2) & _
                              "    AND b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq "
    

    
    
    SQLGetAFBSensPtIdReport = SQLGetAFBSensPtIdReport & " order by ptid,deptcd,vfydt,vfytm "
    
End Function
