VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlETest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SqlGetAppendText(ByVal pTestCd As String, Optional ByVal pApdCd As String = "") As String
    
    SqlGetAppendText = " SELECT * FROM " & T_LAB031 & _
                       " WHERE " & DBW("cdindex ", LC2_ItemResult, 2) & _
                       " AND   " & DBW("cdval1  ", pTestCd, 2)
    If pApdCd <> "" Then
        SqlGetAppendText = SqlGetAppendText & " AND   " & DBW("cdval2  ", pApdCd, 2)
    End If
    
End Function

Public Function SqlGetRemark(Optional ByVal pRmkCd As String = "") As String
    
    SqlGetRemark = " SELECT cdval1, text1 FROM " & T_LAB034 & _
                   " WHERE " & DBW("cdindex ", LC4_Remark, 2)
    If pRmkCd <> "" Then
        SqlGetRemark = SqlGetRemark & " AND " & DBW("cdval1 ", pRmkCd, 2)
    End If
    SqlGetRemark = SqlGetRemark & " ORDER BY cdval1 asc"
    
End Function

Public Function SqlGetTemplateRst(ByVal sRType As String, Optional ByVal sTCode As String = "") As String
    
    SqlGetTemplateRst = " SELECT tpcd,tpnm,tpdata FROM " & T_LAB036 & _
                        " WHERE " & DBW("stcd", sRType, 2)
    If sTCode <> "" Then SqlGetTemplateRst = SqlGetTemplateRst & " AND   " & DBW("tpcd", sTCode, 2)
    
    SqlGetTemplateRst = SqlGetTemplateRst & _
                        " ORDER BY tpcd "
    
End Function

Public Function SqlGetVfyDate(ByVal sWorkArea As String, ByVal sAccDt As String, ByVal sAccSeq As String) As String
    
    SqlGetVfyDate = " SELECT vfydt, vfytm FROM " & T_LAB351 & _
                    " WHERE " & DBW("workarea = ", sWorkArea) & _
                    " AND   " & DBW("accdt    = ", sAccDt) & _
                    " AND   " & DBW("accseq   = ", sAccSeq)

End Function

Public Function SqlLoadTestList() As String
    
    SqlLoadTestList = " SELECT testcd, testnm, rsttype FROM " & T_LAB001 & " a " & _
                      " WHERE " & DBW("a.testdiv", TST_SpeTest, 2) & _
                      " AND   a.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                        " WHERE  testcd  = a.testcd " & _
                                        " AND    applydt <= '" & Format(Now, CS_DateDbFormat) & "') "

End Function

Public Function SqlLoadResultByLabNo(ByVal sTestCd As String, ByVal sWorkArea As String, _
                                     ByVal sAccDt As String, ByVal sAccSeq As String) As String
    
    SqlLoadResultByLabNo = " SELECT stscd, mfyseq FROM " & T_LAB351 & " b " & _
                           " WHERE " & DBW("b.testcd  ", sTestCd, 2) & _
                           " AND   " & DBW("b.workarea", sWorkArea, 2) & _
                           " AND   " & DBW("b.accdt   ", sAccDt, 2) & _
                           " AND   " & DBW("b.accseq  ", sAccSeq, 2)

End Function

Public Function SqlGetDataByLabNo(ByVal sTestCd As String, ByVal sWorkArea As String, _
                                  ByVal sAccDt As String, ByVal sAccSeq As String) As String
    
    SqlGetDataByLabNo = " SELECT a.ptid,f." & F_PTNM & " as ptnm,a.sex,a.ageday,a.deptcd,a.wardid,a.roomid,a.bedid,a.spccd," & _
                        "        a.stscd,b.valfg,a.rmkcd,a.footnotefg,a.rcvdt,b.testcd,c.field4 as spcnm, a.orddoct, a.majdoct " & _
                        " FROM   " & T_LAB201 & " a, " & T_HIS001 & " f, " & T_LAB351 & " b, " & T_LAB032 & " c " & _
                        " WHERE  " & DBW("a.workarea", sWorkArea, 2) & _
                        " AND    " & DBW("a.accdt", sAccDt, 2) & _
                        " AND    " & DBW("a.accseq", sAccSeq, 2) & _
                        " AND    a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
                        " AND    " & DBW("b.testcd", sTestCd, 2) & _
                        " AND    " & DBW("c.cdindex", LC3_Specimen, 2) & _
                        " AND    c.cdval1 = a.spccd" & _
                        " AND    a.ptid   = f." & F_PTID
End Function

Public Function SqlGetLastLabNo(ByVal pPtid As String, ByVal pTestCd As String) As String
    '' 최근결과순서 수정 || vfytm
    SqlGetLastLabNo = " SELECT * FROM " & T_LAB351 & _
                      " WHERE " & DBW("ptid  ", pPtid, 2) & _
                      " AND   " & DBW("testcd", pTestCd, 2) & _
                      " AND   " & DBW("stscd>", enStsCd.StsCd_LIS_MidRst) & _
                      " ORDER BY vfydt || vfytm DESC"
End Function

Public Function SqlGetLastLabNoData(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String) As String
    
    SqlGetLastLabNoData = " SELECT * FROM " & T_LAB201 & _
                      " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                      "   AND  " & DBW("accdt", pAccDt, 2) & _
                      "   AND  " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SqlGetResultText(ByVal sWorkArea As String, ByVal sAccDt As String, _
                                 ByVal sAccSeq As String, ByVal sTestCd As String, _
                                 Optional ByVal sMfySeq As String = "") As String
        
    SqlGetResultText = " SELECT * FROM " & T_LAB353 & _
                       "  WHERE " & DBW("workarea", sWorkArea, 2) & _
                       "    AND " & DBW("accdt", sAccDt, 2) & _
                       "    AND " & DBW("accseq", sAccSeq, 2) & _
                       "    AND " & DBW("testcd", sTestCd, 2)
    If sMfySeq <> "" Then SqlGetResultText = SqlGetResultText & "    AND " & DBW("mfyseq", sMfySeq, 2)

End Function

Public Function SqlGetLabNoList(ByVal pTestCd As String) As String
    
    SqlGetLabNoList = " SELECT * FROM " & T_LAB351 & _
                      "  WHERE " & DBW("testcd", pTestCd, 2) & _
                      "    AND " & DBW("stscd ", enStsCd.StsCd_LIS_MidRst, 2)

End Function

Public Function SqlGetAccList(ByVal pTestCd As String) As String


    SqlGetAccList = " SELECT a.*,b.rcvdt,b.rcvtm FROM " & T_LAB201 & " b," & T_LAB351 & " a " & _
                    "  WHERE " & DBW("a.testcd", pTestCd, 2) & _
                    "    AND " & DBW("a.stscd < ", enStsCd.StsCd_LIS_MidRst) & _
                    "    AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
                    " order by rcvdt,rcvtm"



'    SqlGetAccList = " SELECT * FROM " & T_LAB351 & _
'                    "  WHERE " & DBW("testcd", pTestCd, 2) & _
'                    "    AND " & DBW("stscd < ", enStsCd.StsCd_LIS_MidRst)
    
'    SqlGetAccList = " SELECT a.* FROM " & T_LAB301 & " b," & T_LAB351 & " a" & _
'                    "  WHERE " & DBW("testcd", pTestCd, 2) & _
'                    "    AND " & DBW("stscd < ", enStsCd.StsCd_LIS_MidRst) & _
'                    "    AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
'                    "  ORDER BY b.workdt,b.workcd,b.workseq"

End Function

Public Function SqlGetResultData(ByVal sWorkArea As String, ByVal sAccDt As String, _
                                 ByVal sAccSeq As String) As String
    
    SqlGetResultData = " SELECT a.testcd, a.testnm, b.stscd, a.rsttype, b.mfyseq " & _
                       " FROM " & T_LAB001 & " a, " & T_LAB351 & " b " & _
                       " WHERE a.testcd = b.testcd " & _
                       " AND " & DBW("b.workarea", sWorkArea, 2) & _
                       " AND " & DBW("b.accdt", sAccDt, 2) & _
                       " AND " & DBW("b.accseq", sAccSeq, 2) & _
                       " AND   a.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                       "                WHERE testcd = a.testcd " & _
                       "                AND applydt <= '" & Format(Now, CS_DateDbFormat) & "')"

End Function

Public Function SqlReadFootNote(ByVal sWorkArea As String, ByVal sAccDt As String, _
                                ByVal sAccSeq As String) As String
       
    SqlReadFootNote = " SELECT * FROM " & T_LAB304 & _
                      "  WHERE " & DBW("workarea", sWorkArea, 2) & _
                      "    AND " & DBW("accdt", sAccDt, 2) & _
                      "    AND " & DBW("accseq", sAccSeq, 2)
End Function


Public Function SqlGetCurrRst(ByVal sWorkArea As String, ByVal sAccDt As String, _
                              ByVal sAccSeq As String, ByVal sTestCd As String) As String
    
    SqlGetCurrRst = " SELECT * FROM " & T_LAB351 & _
                    "  WHERE " & DBW("workarea", sWorkArea, 2) & _
                    "   AND  " & DBW("accdt", sAccDt, 2) & _
                    "   AND  " & DBW("accseq", sAccSeq, 2) & _
                    "   AND  " & DBW("testcd", sTestCd, 2) & _
                    "   AND  " & DBW("stscd>=", enStsCd.StsCd_LIS_FinRst)

End Function

Public Function SqlUpdateETest(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByVal pTestCd As String, _
                               ByVal pStatus As String, ByVal pMfyCnt As Integer, ByVal pValRst As String, _
                               ByVal pTxtRst As String, ByVal pDate As String, pTime As String, pEmpId As String)

    SqlUpdateETest = "UPDATE " & T_LAB351 & _
                     "   SET " & _
                                 DBW("valfg", pValRst, 3) & _
                                 DBW("txtfg", pTxtRst, 3) & _
                                 DBW("stscd", pStatus, 3) & _
                                 DBW("mfyseq", pMfyCnt, 3) & _
                                 DBW("vfydt", pDate, 3) & _
                                 DBW("vfytm", pTime, 3) & _
                                 DBW("vfyid", pEmpId, 2) & _
                     " WHERE " & DBW("workarea", pWorkArea, 2) & _
                     "   AND " & DBW("accdt", pAccDt, 2) & _
                     "   AND " & DBW("accseq", pAccSeq, 2) & _
                     "   AND " & DBW("testcd", pTestCd, 2)

End Function


Public Function SqlUpdateOrder(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByVal pTestCd As String, _
                               ByVal pStatus As String, ByVal pDate As String, _
                               ByVal pTime As String, ByVal pEmpId As String) As String
    
    '** 원본 ===========================================================================
'    SqlUpdateOrder = "UPDATE " & T_LAB102 & _
'                     "   SET " & _
'                                 DBW("stscd", pStatus, 3) & _
'                                 DBW("examdt", pDate, 3) & _
'                                 DBW("examtm", pTime, 3) & _
'                                 DBW("examdoct", pEmpId, 2) & _
'                     " WHERE EXISTS " & _
'                     "  (SELECT * FROM " & T_LAB351 & " b " & _
'                     "   WHERE " & DBW("b.workarea", pWorkArea, 2) & _
'                     "   AND   " & DBW("b.accdt", pAccDt, 2) & _
'                     "   AND   " & DBW("b.accseq", pAccSeq, 2) & _
'                     "   AND   " & DBW("b.testcd", pTestCd, 2) & _
'                     "   AND " & T_LAB102 & ".ptid = b.ptid " & _
'                     "   AND " & T_LAB102 & ".orddt = b.orddt " & _
'                     "   AND " & T_LAB102 & ".ordno = b.ordno " & _
'                     "   AND " & T_LAB102 & ".ordseq = b.ordseq) "
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
        SqlUpdateOrder = "UPDATE su2examt " & _
                         "   SET " & _
                                     DBW("stscd", pStatus, 3) & _
                         " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                         ", examdr = " & DBS(pEmpId) & _
                         " WHERE EXISTS " & _
                         "  (SELECT * FROM " & T_LAB351 & " b " & _
                         "   WHERE " & DBW("b.workarea", pWorkArea, 2) & _
                         "   AND   " & DBW("b.accdt", pAccDt, 2) & _
                         "   AND   " & DBW("b.accseq", pAccSeq, 2) & _
                         "   AND   " & DBW("b.testcd", pTestCd, 2) & _
                         "   AND mdexmort.patno = b.ptid " & _
                         "   AND mdexmort.orddt = TO_DATE(b.orddt, 'yyyymmdd') " & _
                         "   AND mdexmort.ordno = b.ordno "
    ElseIf strWrkDiv = "4" Then     '일건
            SqlUpdateOrder = "UPDATE sg2examt " & _
                             "   SET " & _
                                         DBW("stscd", pStatus, 3) & _
                             " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                             ", examdr = " & DBS(pEmpId) & _
                             " WHERE EXISTS " & _
                             "  (SELECT * FROM " & T_LAB351 & " b " & _
                             "   WHERE " & DBW("b.workarea", pWorkArea, 2) & _
                             "   AND   " & DBW("b.accdt", pAccDt, 2) & _
                             "   AND   " & DBW("b.accseq", pAccSeq, 2) & _
                             "   AND   " & DBW("b.testcd", pTestCd, 2) & _
                              "   AND mdexmort.patno = b.ptid " & _
                             "   AND mdexmort.orddt = TO_DATE(b.orddt, 'yyyymmdd') " & _
                             "   AND mdexmort.ordno = b.ordno "
    Else
        SqlUpdateOrder = "UPDATE mdexmort " & _
                         "   SET " & _
                                     DBW("stscd", pStatus, 3) & _
                         " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                         ", cofmdr = " & DBS(pEmpId) & _
                         " WHERE EXISTS " & _
                         "  (SELECT * FROM " & T_LAB351 & " b " & _
                         "   WHERE " & DBW("b.workarea", pWorkArea, 2) & _
                         "   AND   " & DBW("b.accdt", pAccDt, 2) & _
                         "   AND   " & DBW("b.accseq", pAccSeq, 2) & _
                         "   AND   " & DBW("b.testcd", pTestCd, 2) & _
                         "   AND mdexmort.patno = b.ptid " & _
                         "   AND mdexmort.orddt = TO_DATE(b.orddt, 'yyyymmdd') " & _
                         "   AND mdexmort.ordno = b.ordno "
    End If
    '===================================================================================

End Function

Public Function SqlSetOrdStatus(ByVal pPtid As String, ByVal pOrdDt As String, _
                                ByVal pOrdNo As String, ByVal pOrdSeq As String, _
                                ByVal pStatus As String, ByVal pDate As String, _
                                ByVal pTime As String, ByVal pEmpId As String) As String
    
    '** 원본 ===========================================================================
'    SqlSetOrdStatus = "UPDATE " & T_LAB102 & _
'                      "   SET " & _
'                                    DBW("stscd", pStatus, 3) & _
'                                    DBW("examdt", pDate, 3) & _
'                                    DBW("examtm", pTime, 3) & _
'                                    DBW("examdoct", pEmpId, 2) & _
'                      " WHERE   " & DBW("ptid", pPtid, 2) & _
'                      "   AND   " & DBW("orddt", pOrdDt, 2) & _
'                      "   AND   " & DBW("ordno", pOrdNo, 2) & _
'                      "   AND   " & DBW("ordseq", pOrdSeq, 2)
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
        SqlSetOrdStatus = "UPDATE su2examt " & _
                          "   SET " & _
                                        DBW("stscd", pStatus, 3) & _
                          " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                          ", examdr = " & DBS(pEmpId) & _
                          " WHERE   patno = " & DBS(pPtid) & _
                          "   AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                          "   AND   ordseqno = " & DBN(pOrdNo)
    ElseIf strWrkDiv = "4" Then     '일건
            SqlSetOrdStatus = "UPDATE sg2examt " & _
                              "   SET " & _
                                            DBW("stscd", pStatus, 3) & _
                              " examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                              ", examdr = " & DBS(pEmpId) & _
                              " WHERE   patno = " & DBS(pPtid) & _
                              "   AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                              "   AND   ordseqno = " & DBN(pOrdNo)
    Else
        SqlSetOrdStatus = "UPDATE mdexmort " & _
                          "   SET " & _
                                        DBW("stscd", pStatus, 3) & _
                          " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                          ", cofmdr = " & DBS(pEmpId) & _
                          " WHERE   patno = " & DBS(pPtid) & _
                          "   AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                          "   AND   ordseqno = " & DBN(pOrdNo)
    End If
    '===================================================================================
        If Mid(pPtid, 1, 1) = "L" Then
        SqlSetOrdStatus = "UPDATE s2ord999 " & _
                          "   SET " & _
                                        DBW("stscd", pStatus, 3) & _
                          " rsltdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') " & _
                          ", cofmdr = " & DBS(pEmpId) & _
                          " WHERE   patno = " & DBS(pPtid) & _
                          "   AND   orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                          "   AND   ordseqno = " & DBN(pOrdNo)
        End If
End Function

Public Function SqlGetOrderSeq(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByVal pTestCd As String) As String
                               
    SqlGetOrderSeq = " SELECT * FROM " & T_LAB351 & " b " & _
                     "   WHERE " & DBW("b.workarea", pWorkArea, 2) & _
                     "   AND   " & DBW("b.accdt", pAccDt, 2) & _
                     "   AND   " & DBW("b.accseq", pAccSeq, 2) & _
                     "   AND   " & DBW("b.testcd", pTestCd, 2)

End Function


Public Function SqlSaveValResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String, _
                                 ByVal pMfySeq As String, ByVal pResult As String, _
                                 ByVal pFlag As String) As String
                               
    If pFlag = "1" Then
        SqlSaveValResult = " DELETE " & T_LAB353 & _
                           "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                           "    AND " & DBW("accdt", pAccDt, 2) & _
                           "    AND " & DBW("accseq", pAccSeq, 2) & _
                           "    AND " & DBW("testcd", pTestCd, 2) & _
                           "    AND " & DBW("mfyseq", pMfySeq, 2)
    Else
        SqlSaveValResult = " INSERT INTO " & T_LAB353 & _
                           " (workarea,accdt,accseq,testcd,mfyseq,txtrst) " & _
                           " VALUES " & _
                           " ( " & DBV("workarea", pWorkArea, 1) & _
                                   DBV("accdt", pAccDt, 1) & _
                                   DBV("accseq", pAccSeq, 1) & _
                                   DBV("testcd", pTestCd, 1) & _
                                   DBV("mfyseq", pMfySeq, 1) & _
                                   DBV("txtrst", pResult) & _
                           " ) "
    End If
    
End Function

Public Function SqlSaveSpeResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String, _
                                 ByVal pMfySeq As String, ByVal pResult As String, _
                                 ByVal pRstCd1 As String, ByVal pRstCd2 As String, _
                                 ByVal pRstCd3 As String, ByVal pFlag As String) As String
                               
    If pFlag = "1" Then
        SqlSaveSpeResult = " DELETE " & T_LAB353 & _
                           "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                           "    AND " & DBW("accdt", pAccDt, 2) & _
                           "    AND " & DBW("accseq", pAccSeq, 2) & _
                           "    AND " & DBW("testcd", pTestCd, 2) & _
                           "    AND " & DBW("mfyseq", pMfySeq, 2)
    Else
        SqlSaveSpeResult = " INSERT INTO " & T_LAB353 & _
                           " (workarea,accdt,accseq,testcd,mfyseq,txtrst,rstcd1,rstcd2,rstcd3) " & _
                           " VALUES " & _
                           " ( " & DBV("workarea", pWorkArea, 1) & _
                                   DBV("accdt", pAccDt, 1) & _
                                   DBV("accseq", pAccSeq, 1) & _
                                   DBV("testcd", pTestCd, 1) & _
                                   DBV("mfyseq", pMfySeq, 1) & _
                                   DBV("txtrst", pResult, 1) & _
                                   DBV("rstcd1", pRstCd1, 1) & _
                                   DBV("rstcd2", pRstCd2, 1) & _
                                   DBV("rstcd3", pRstCd3) & _
                           " ) "
    End If
    
End Function

Public Function SqlDeleteFootnote(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String) As String

    SqlDeleteFootnote = " DELETE " & T_LAB304 & " " & _
                        "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                        "    AND " & DBW("accdt", pAccDt, 2) & _
                        "    AND " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SqlGetMfySeq(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String) As String
    
    SqlGetMfySeq = " SELECT max(seq) mfyno FROM " & T_LAB304 & _
                   " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                   "   AND  " & DBW("accdt", pAccDt, 2) & _
                   "   AND  " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SqlInsertFootnote(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pMfyNo As Long, _
                                  ByVal pEmpId As String, ByVal pFootNote As String) As String
    
    SqlInsertFootnote = " INSERT INTO " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
                        " VALUES (" & DBV("workarea", pWorkArea, 1) & _
                                      DBV("accdt", pAccDt, 1) & _
                                      DBV("accseq", pAccSeq, 1) & _
                                      DBV("seq", pMfyNo, 1) & _
                                      DBV("vfyid", pEmpId, 1) & _
                                      DBV("rsttxt", pFootNote) & ")"

End Function


Public Function SqlUpdateAccData(ByVal sWorkArea As String, ByVal sAccDt As String, _
                                 ByVal sAccSeq As String, ByVal sFSeq As Long, _
                                 ByVal sRemarkCd As String) As String
    
    SqlUpdateAccData = " UPDATE " & T_LAB201 & _
                       "    SET " & _
                                    DBW("footnotefg", sFSeq, 3) & _
                                    DBW("rmkcd", sRemarkCd, 2) & _
                       "  WHERE " & DBW("workarea", sWorkArea, 2) & _
                       "    AND " & DBW("accdt", sAccDt, 2) & _
                       "    AND " & DBW("accseq", sAccSeq, 2)

End Function

Public Function SqlReadVerifyList(ByVal sDeptCd As String, ByVal sDate As String, _
                                  ByVal sPtid As String, ByVal sMfyFg As String) As String
    
    SqlReadVerifyList = " SELECT * FROM " & T_LAB202 & _
                        "  WHERE " & DBW("deptcd", sDeptCd, 2) & _
                        "    AND " & DBW("vfydt", sDate, 2) & _
                        "    AND " & DBW("ptid", sPtid, 2) & _
                        "    AND " & DBW("mfyfg", sMfyFg, 2)

End Function

Public Function SqlUpdateVerifyList(ByVal sDeptCd As String, ByVal sDate As String, ByVal sTime As String, _
                                    ByVal sPtid As String, ByVal sMfyFg As String, ByVal sEmpId As String) As String
    
    SqlUpdateVerifyList = " UPDATE " & T_LAB202 & " " & _
                          "    SET " & _
                                       DBW("vfytm", sTime, 3) & _
                                       DBW("vfyid", sEmpId, 3) & _
                                       DBW("donefg", "", 3) & _
                                       DBW("doneid", "0", 2) & _
                          " WHERE  " & DBW("deptcd", sDeptCd, 2) & _
                          "   AND  " & DBW("vfydt", sDate, 2) & _
                          "   AND  " & DBW("ptid", sPtid, 2) & _
                          "   AND  " & DBW("mfyfg", sMfyFg, 2)

End Function

Public Function SqlInsertVerifyList(ByVal sDeptCd As String, ByVal sDate As String, ByVal sTime As String, _
                                    ByVal sPtid As String, ByVal sMfyFg As String, ByVal sEmpId As String, _
                                    ByVal sMajDoct As String, ByVal sBussDiv As String) As String
        
    SqlInsertVerifyList = " INSERT INTO " & T_LAB202 & _
                          " (deptcd,vfydt,ptid,mfyfg,vfytm,vfyid,donefg,doneid,majdoct,bussdiv)" & _
                          " VALUES " & _
                          " ( " & _
                                DBV("deptcd", sDeptCd, 1) & _
                                DBV("vfydt", sDate, 1) & _
                                DBV("ptid", sPtid, 1) & _
                                DBV("mfyfg", sMfyFg, 1) & _
                                DBV("vfytm", sTime, 1) & _
                                DBV("vfyid", sEmpId, 1) & _
                                DBV("donefg", "", 1) & _
                                DBV("doneid", "0", 1) & _
                                DBV("majdoct", sMajDoct, 1) & _
                                DBV("bussdiv", sBussDiv) & _
                          " ) "
End Function

Public Function SqlGetTestCount(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String) As String
    
    SqlGetTestCount = " SELECT reqtotcnt,reqinputcnt,stscd FROM " & T_LAB201 & _
                      " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                      "   AND  " & DBW("accdt", pAccDt, 2) & _
                      "   AND  " & DBW("accseq", pAccSeq, 2)

End Function

Public Function SqlUpdateAccStatus(ByVal sWorkArea As String, ByVal sAccDt As String, _
                                   ByVal sAccSeq As String, ByVal sStatus As String, _
                                   ByVal iInputCount As Integer, _
                                   Optional ByVal pDate As String, _
                                   Optional ByVal pTime As String, _
                                   Optional ByVal pEmpId As String) As String
    
    If sStatus >= enStsCd.StsCd_LIS_FinRst Then
        SqlUpdateAccStatus = " UPDATE " & T_LAB201 & " " & _
                             "    SET " & _
                                          DBW("stscd", sStatus, 3) & _
                                          DBW("reqinputcnt", iInputCount, 3) & _
                                          DBW("vfydt", pDate, 3) & _
                                          DBW("vfytm", pTime, 3) & _
                                          DBW("vfyid", pEmpId, 2) & _
                             "  WHERE " & DBW("workarea", sWorkArea, 2) & _
                             "    AND " & DBW("accdt", sAccDt, 2) & _
                             "    AND " & DBW("accseq", sAccSeq, 2)
    Else
        SqlUpdateAccStatus = " UPDATE " & T_LAB201 & _
                             "    SET " & DBW("reqinputcnt", iInputCount, 2) & _
                             "  WHERE " & DBW("workarea", sWorkArea, 2) & _
                             "    AND " & DBW("accdt", sAccDt, 2) & _
                             "    AND " & DBW("accseq", sAccSeq, 2)
    End If

End Function

'%  47. LD, CPK 등 관련검사의 결과 조회
'%       - Calling FROM [ frm202AccDataEntry, frm204WSDataEntry ] :
Public Function SqlGetRelTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As String) As String
      
    SqlGetRelTest = " SELECT a.testcd, b.testnm, b.abbrnm5, a.rstcd, a.rstunit, a.rstdiv, a.detailfg, a.spccd, a.vfydt, a.vfytm, " & _
                    "        c.empnm as VfyNm, a.hldiv, a.dpdiv, b.rptseq, g.field1 rstnm, " & _
                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & _
                             FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & _
                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
                    "        f.coldt, f.coltm, f.ageday, f.sex " & _
                    " FROM   " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " a, " & _
                                 T_LAB031 & " g, " & T_LAB031 & " e, " & T_LAB201 & " f, " & T_LAB351 & " d " & _
                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
                    " AND    " & DBW("d.accdt", AccDt, 2) & _
                    " AND    " & DBW("d.accseq", accseq, 2) & _
                    " AND    " & DBW("e.cdindex", LC2_RelTest, 2) & _
                    " AND    e.cdval1 = d.testcd " & _
                    " AND    f.workarea = a.workarea   AND  f.accdt    = a.accdt  AND  f.accseq   = a.accseq" & _
                    " AND    a.ptid   = d.ptid " & _
                    " AND    a.testcd = e.cdval2 " & _
                    " AND    a.vfydt <> ' '" & _
                    " AND    a.vfydt" & FUNC_CONCAT & "a.vfytm = (SELECT max(vfydt" & FUNC_CONCAT & "vfytm) FROM " & T_LAB302 & " WHERE ptid = a.ptid AND testcd = a.testcd) " & _
                    " AND    " & DBJ("c.empid =* a.vfyid") & _
                    " AND    " & DBJ(DBW("g.cdindex = ", LC2_ItemResult)) & _
                    " AND    " & DBJ("g.cdval1 =* a.testcd") & _
                    " AND    " & DBJ("g.cdval2 =* a.rstcd") & _
                    " AND    b.testcd = a.testcd " & _
                    " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " order  by b.rptseq "
               
End Function

Public Function SqlGetSpeRstCode(ByVal TestCd As String, Optional ByVal RstType As String) As String
    
    SqlGetSpeRstCode = " SELECT a.cdval2 rstcd, a.field1 rstnm " & _
                    " FROM   " & T_LAB031 & " a " & _
                    " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
                    " AND    " & DBW("a.cdval1", TestCd, 2) & " order by rstcd"
'    SqlGetSpeRstCode = " SELECT a.cdval2 rstcd, a.text1 rstnm " & _
                    " FROM   " & T_LAB031 & " a " & _
                    " WHERE  " & DBW("a.cdindex", LC2_SpeAddTemp, 2) & _
                    " AND    a.cdval1 like '<#" & RstType & "%'"
'    SqlGetSpeRstCode = " SELECT a.cdval2 rstcd, a.text1 rstnm " & _
                    " FROM   " & T_LAB031 & " a, " & T_LAB031 & " b " & _
                    " WHERE  " & DBW("b.cdindex", LC2_SpeTemp, 2) & _
                    " AND    " & DBW("b.cdval1", RstType, 2) & _
                    " AND    " & DBW("a.cdindex", LC2_SpeAddTemp, 2) & _
                    " AND    a.cdval1=b.cdval2 "

End Function
