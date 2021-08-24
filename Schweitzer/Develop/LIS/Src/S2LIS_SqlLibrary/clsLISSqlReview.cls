VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#Const APS_REVIEW = True
'#Const BBS_REVIEW = True
'#Const POC_REVIEW = False
'#Const CMT_REVIEW = True

'%  32. 환자Id, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
'%       - Lab101 : 처방Header내역
'%       - Lab102 : 처방Body내역
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Private mvarRiPrint As String

Private mvarTestCd As String

Public Property Let RiPrint(ByVal vNewValue As String)
    mvarRiPrint = vNewValue
End Property

Public Property Let TestCd(ByVal vNewValue As String)
    mvarTestCd = vNewValue
End Property
   
'%  32. 환자Id, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
'%       - Lab101 : 처방Header내역
'%       - Lab102 : 처방Body내역
'%       - Calling FROM [ frm401ResultView ] : 결과조회

Public Function SqlQueryOrders(ByVal PtId As String, ByVal QueryKey As String, ByVal FromDate As String, _
                                  ByVal ToDate As String, ByVal pResultReviewFlag As String, Optional ByVal lngCase As Long = 0) As String
    
    'lngCase : 0-전체, 1-임상, 2-해부, 3-혈액, 7-미생물
    
    Dim strSql1 As String, strSql2 As String, strSql3 As String, strSql4 As String, strSql5 As String, strSql6 As String, strSql8 As String, strSql9 As String
    
    If pResultReviewFlag = "7" Then
        
        strSql5 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
                  "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
                  "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
                  "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
                  "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
                  "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
                  "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, b.orddtm as orddtm, " & _
                  "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
                  "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct, null as rstseq  " & _
                  " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
                               T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
                  " WHERE  " & DBW("a.ptid", PtId, 2) & _
                  " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
                  " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
                  " AND    a.ptid    =  b.ptid  AND    a.orddt   =  b.orddt  AND    a.ordno   =  b.ordno " & _
                  " AND    " & DBW("b.orddiv", "L", 2) & _
                  " AND    c.testcd  =  a.ordcd " & _
                  " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
                                      "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
                  " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
                  " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
                  " AND    " & DBJ("g.empid   =* a.examdoct") & _
                  " AND    exists (SELECT * FROM " & T_LAB404 & " WHERE ptid = a.ptid AND orddt = a.orddt " & _
                                                                " AND ordno = a.ordno AND ordseq = a.ordseq )" 'AND senfg = 'Y'
                  
                       
    End If
    
                          
'    strSql1 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
'              "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
'              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
'              "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
'              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
'              "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
'              "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, " & _
'              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
'              "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct " & _
'              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
'                           T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
'              " WHERE  " & DBW("a.ptid", PtId, 2) & _
'              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
'              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
'              " AND    a.ptid    =  b.ptid  " & _
'              " AND    a.orddt   =  b.orddt " & _
'              " AND    a.ordno   =  b.ordno " & _
'              " AND    " & DBW("b.orddiv", "L", 2) & _
'              " AND    c.testcd  =  a.ordcd " & _
'              " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
'                                  "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
'              " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
'              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
'              " AND    " & DBJ("g.empid   =* a.examdoct")
                       
     strSql1 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
              "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
              "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
              "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, b.orddtm as orddtm, " & _
              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct, null as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
                           T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
              " AND    a.ptid    =  b.ptid  " & _
              " AND    a.orddt   =  b.orddt " & _
              " AND    a.ordno   =  b.ordno " & _
              " AND    " & DBW("b.orddiv", "L", 2) & _
              " AND    c.testcd  =  a.ordcd " & _
              " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
                                  "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
              " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
              " AND    " & DBJ("g.empid   =* a.examdoct")
                       

'    strSql3 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
'              "        a.ordno, a.ordseq, a.ordcd, c.testnm, '' panelfg, '' grpfg, a.spccd, a.statfg, " & _
'              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
'              "        a.stscd, a.dcfg, a.mesg, '혈액' SpcNm, '' MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
'              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
'              "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
'              "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, " & _
'              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, '' testdiv, '*' as Verified, " & _
'              "        '' spcyy, 0 spcno, a.unitqty, b.reqdt, b.reqtm, b.majdoct " & _
'              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & _
'                           T_BBS001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a" & _
'              " WHERE  " & DBW("a.ptid", PtId, 2) & _
'              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
'              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
'              " AND    a.ptid    =  b.ptid  " & _
'              " AND    a.orddt   =  b.orddt " & _
'              " AND    a.ordno   =  b.ordno " & _
'              " AND    " & DBW("b.orddiv", BBS_ORDDIV, 2) & _
'              " AND    c.testcd  =  a.ordcd   " & _
'              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct ") & _
'              " AND    " & DBJ("g.empid   =* a.examdoct")
                    
    strSql3 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
              "        a.ordno, a.ordseq, a.ordcd, c.testnm, '' panelfg, '' grpfg, a.spccd, a.statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
              "        a.stscd, a.dcfg, a.mesg, '혈액' SpcNm, '' MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(d.vfydt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(d.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(d.vfydt,7,2) as ExamDt, " & _
              "        " & FUNC_SUBSTR & "(d.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(d.vfytm,3,2) as ExamTm, g.empnm as ExamNm, '' as orddtm, " & _
              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, '' testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, a.unitqty, b.reqdt, b.reqtm, b.majdoct, d.rstseq as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & _
                           T_BBS001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a, " & T_BBS302 & " d " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
              " AND    a.ptid    =  b.ptid  " & _
              " AND    a.orddt   =  b.orddt " & _
              " AND    a.ordno   =  b.ordno " & _
              " AND    " & DBW("b.orddiv", BBS_ORDDIV, 2) & _
              " AND    c.testcd  =  a.ordcd   " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct ") & _
              " AND    " & DBJ("g.empid   =* d.vfyid") & _
              " AND    a.WORKAREA = d.WORKAREA" & _
              " AND    a.ACCDT = d.ACCDT" & _
              " AND    a.ACCSEQ = d.ACCSEQ"
                    
    strSql4 = " SELECT distinct " & FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as OrdDate, " & _
              "        99 ordno, 1 ordseq, a.testcd ordcd, c.testnm, c.panelfg, c.grpfg, b.field2 spccd, '' statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.vfydt rcvdt, '00:00' as RcvTm, " & _
                       DBV("stscd", enStsCd.StsCd_LIS_FinRst) & " stscd, '' dcfg, '' mesg, d.field3 SpcNm, '' MultiFg, a.vfydt orddt, 'P' orddiv, a.wardid, '' hosilid, " & _
                       " '00:00' as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(a.vfydt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as ExamDt, " & _
                       " '00:00' as ExamTm, '' as ExamNm, '' as orddtm, " & _
              "        '' workarea, '' accdt, 0 accseq, " & DBV("stscd", enStsCd.StsCd_LIS_FinRst) & " as RstIn, " & DBV("testdiv", enTestDiv.TST_RouTest) & " testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, 0 unitqty, '' reqdt, '' reqtm, '' majdoct, null as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB032 & " d, " & _
                           T_LAB001 & " c, " & T_LAB032 & " b, " & T_LAB195 & " a  " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a.vfydt >= " & DBV(QueryKey, FromDate) & _
              " AND    a.vfydt <= " & DBV(QueryKey, ToDate) & _
              " AND    " & DBW("b.cdindex = ", LC3_POCTestCd) & _
              " AND    b.cdval1  =  a.testcd " & _
              " AND    c.testcd  =  a.testcd   " & _
              " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
              " AND    d.cdval1  =  b.field2 " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* a.majdoct ")
                    
    strSql6 = " SELECT  distinct " & FUNC_SUBSTR & "(b.rptdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.rptdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.rptdt,7,2) as OrdDate, " & _
              "         1 ordno, 0 ordseq, null ordcd, '임상병리 종합검증/판독' testnm, null panelfg, null grpfg, null spccd, null statfg, " & _
              "         g.empnm as DoctNm, null rcvdt, null RcvTm, '8' stscd, null dcfg, null mesg, " & _
              "         null SpcNm, null as MultiFg, a.bedindt as orddt, 'C' orddiv, a.wardid, a.hosilid, null OrdTm, " & _
              "         " & FUNC_SUBSTR & "(a.rptdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,7,2) as ExamDt, " & _
              "         null ExamTm, g.empnm as ExamNm, '' as orddtm, " & _
              "         null workarea, null accdt, 0 accseq, null  RstIn, '' testdiv, '*' as Verified, '' spcyy, 0 spcno, 0 unitqty, '' reqdt, '' reqtm, Null majdoct, null as rstseq  " & _
              " FROM    " & T_LAB501 & " a, " & T_LAB502 & " b, " & T_LAB015 & " g " & _
              " WHERE   " & DBW("a.ptid = ", PtId) & _
              " AND     " & DBW("a.bedindt>=", FromDate) & _
              " AND     " & DBW("a.bedindt<=", ToDate) & _
              " AND     a.donefg = '2' " & _
              " AND     b.ptid  =  a.ptid" & _
              " AND     b.rptdt =  a.rptdt " & _
              " AND     " & DBJ("g.empid =* a.rptid")
              
     strSql8 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
              "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
              "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
              "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, b.orddtm as orddtm, " & _
              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct, null as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
                           T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
              " AND    a.ptid    =  b.ptid  " & _
              " AND    a.orddt   =  b.orddt " & _
              " AND    a.ordno   =  b.ordno " & _
              " AND    " & DBW("b.orddiv", "L", 2) & _
              " AND    c.testcd  =  a.ordcd " & _
              " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
                                  "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
              " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
              " AND    " & DBJ("g.empid   =* a.examdoct") & _
              " AND    a.ORDCD = 'B1114G'  "
                       
     strSql9 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
              "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
              "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
              "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, b.orddtm as orddtm, " & _
              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct, null as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
                           T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
              " AND    a.ptid    =  b.ptid  " & _
              " AND    a.orddt   =  b.orddt " & _
              " AND    a.ordno   =  b.ordno " & _
              " AND    " & DBW("b.orddiv", "L", 2) & _
              " AND    c.testcd  =  a.ordcd " & _
              " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
                                  "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
              " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
              " AND    " & DBJ("g.empid   =* a.examdoct") & _
              " AND    a.ORDCD in ('C6001','C6005','27BN1')  "
              
    Select Case pResultReviewFlag
        Case 0:
            SqlQueryOrders = strSql1 & " UNION ALL " & strSql3 & " UNION ALL " & strSql6        '& " UNION ALL " & strSql4
        Case 1: SqlQueryOrders = strSql1 & " UNION ALL " & strSql6     '임상
        Case 2: SqlQueryOrders = strSql2    '해부
        Case 3: SqlQueryOrders = strSql1  '임상/해부
        Case 4: SqlQueryOrders = strSql3  '혈액
        Case 5: SqlQueryOrders = strSql1  '임상/혈액
        Case 6: SqlQueryOrders = strSql1  '해부/혈액
        Case 7: SqlQueryOrders = strSql5  '미생물
        Case 8: SqlQueryOrders = strSql8  'BMA
        Case 9: SqlQueryOrders = strSql9  '염색체
    End Select
    SqlQueryOrders = SqlQueryOrders & " Order  by orddt desc , ordno, ordseq "

End Function

'%  32. 환자Id, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
'%       - Lab101 : 처방Header내역
'%       - Lab102 : 처방Body내역
'%       - Calling FROM [ frm401ResultView ] : 결과조회

Public Function SqlQueryOrders_WM(ByVal PtId As String, ByVal QueryKey As String, ByVal FromDate As String, _
                                  ByVal ToDate As String, ByVal pResultReviewFlag As String, Optional ByVal lngCase As Long = 0) As String
    
    'lngCase : 0-전체, 1-임상, 2-해부, 3-혈액, 7-미생물
    
    Dim strSql1 As String, strSql2 As String, strSql3 As String, strSql4 As String, strSql5 As String, strSql6 As String
    
    If pResultReviewFlag = "7" Then
        
        strSql5 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
                  "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
                  "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
                  "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
                  "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
                  "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
                  "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, b.orddtm as orddtm, " & _
                  "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
                  "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct, null as rstseq  " & _
                  " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
                               T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
                  " WHERE  " & DBW("a.ptid", PtId, 2) & _
                  " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
                  " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
                  " AND    a.ptid    =  b.ptid  AND    a.orddt   =  b.orddt  AND    a.ordno   =  b.ordno " & _
                  " AND    " & DBW("b.orddiv", "L", 2) & _
                  " AND    c.testcd  =  a.ordcd " & _
                  " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
                                      "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
                  " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
                  " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
                  " AND    " & DBJ("g.empid   =* a.examdoct") & _
                  " AND    exists (SELECT * FROM " & T_LAB404 & " WHERE ptid = a.ptid AND orddt = a.orddt " & _
                                                                " AND ordno = a.ordno AND ordseq = a.ordseq )" 'AND senfg = 'Y'
                  
                       
    End If
                       
     strSql1 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
              "        a.ordno, a.ordseq, a.ordcd, c.testnm, c.panelfg, c.grpfg, a.spccd, a.statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
              "        a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(a.examdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examdt,7,2) as ExamDt, " & _
              "        " & FUNC_SUBSTR & "(a.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, b.orddtm as orddtm, " & _
              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, 0 unitqty, b.reqdt, b.reqtm, b.majdoct, null as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_LAB032 & " d, " & _
                           T_LAB001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
              " AND    a.ptid    =  b.ptid  " & _
              " AND    a.orddt   =  b.orddt " & _
              " AND    a.ordno   =  b.ordno " & _
              " AND    " & DBW("b.orddiv", "L", 2) & _
              " AND    c.testcd  =  a.ordcd " & _
              " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & _
                                  "  WHERE testcd = c.testcd AND applydt <= a.orddt) " & _
              " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & " AND    d.cdval1 = a.spccd " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct") & _
              " AND    " & DBJ("g.empid   =* a.examdoct") & _
              " AND    a.ordcd in ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E')     "
                    
    strSql3 = " SELECT " & FUNC_SUBSTR & "(a.orddt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.orddt,7,2) as OrdDate, " & _
              "        a.ordno, a.ordseq, a.ordcd, c.testnm, '' panelfg, '' grpfg, a.spccd, a.statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.rcvdt, " & FUNC_SUBSTR & "(a.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rcvtm,3,2) as RcvTm, " & _
              "        a.stscd, a.dcfg, a.mesg, '혈액' SpcNm, '' MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
              "        " & FUNC_SUBSTR & "(b.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.ordtm,3,2) as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(d.vfydt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(d.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(d.vfydt,7,2) as ExamDt, " & _
              "        " & FUNC_SUBSTR & "(d.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(d.vfytm,3,2) as ExamTm, g.empnm as ExamNm, '' as orddtm, " & _
              "        a.workarea, a.accdt, a.accseq, a.stscd as RstIn, '' testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, a.unitqty, b.reqdt, b.reqtm, b.majdoct, d.rstseq as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB015 & " g, " & _
                           T_BBS001 & " c, " & T_LAB101 & " b, " & T_LAB102 & " a, " & T_BBS302 & " d " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a." & QueryKey & " >= " & DBV(QueryKey, FromDate) & _
              " AND    a." & QueryKey & " <= " & DBV(QueryKey, ToDate) & _
              " AND    a.ptid    =  b.ptid  " & _
              " AND    a.orddt   =  b.orddt " & _
              " AND    a.ordno   =  b.ordno " & _
              " AND    " & DBW("b.orddiv", BBS_ORDDIV, 2) & _
              " AND    c.testcd  =  a.ordcd   " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* b.orddoct ") & _
              " AND    " & DBJ("g.empid   =* d.vfyid") & _
              " AND    a.WORKAREA = d.WORKAREA" & _
              " AND    a.ACCDT = d.ACCDT" & _
              " AND    a.ACCSEQ = d.ACCSEQ"
                    
    strSql4 = " SELECT distinct " & FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as OrdDate, " & _
              "        99 ordno, 1 ordseq, a.testcd ordcd, c.testnm, c.panelfg, c.grpfg, b.field2 spccd, '' statfg, " & _
              "        f." & F_DOCTNM & " as DoctNm, a.vfydt rcvdt, '00:00' as RcvTm, " & _
                       DBV("stscd", enStsCd.StsCd_LIS_FinRst) & " stscd, '' dcfg, '' mesg, d.field3 SpcNm, '' MultiFg, a.vfydt orddt, 'P' orddiv, a.wardid, '' hosilid, " & _
                       " '00:00' as OrdTm, " & _
              "        " & FUNC_SUBSTR & "(a.vfydt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as ExamDt, " & _
                       " '00:00' as ExamTm, '' as ExamNm, '' as orddtm, " & _
              "        '' workarea, '' accdt, 0 accseq, " & DBV("stscd", enStsCd.StsCd_LIS_FinRst) & " as RstIn, " & DBV("testdiv", enTestDiv.TST_RouTest) & " testdiv, '*' as Verified, " & _
              "        '' spcyy, 0 spcno, 0 unitqty, '' reqdt, '' reqtm, '' majdoct, null as rstseq " & _
              " FROM   " & T_HIS005 & " f, " & T_LAB032 & " d, " & _
                           T_LAB001 & " c, " & T_LAB032 & " b, " & T_LAB195 & " a  " & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    a.vfydt >= " & DBV(QueryKey, FromDate) & _
              " AND    a.vfydt <= " & DBV(QueryKey, ToDate) & _
              " AND    " & DBW("b.cdindex = ", LC3_POCTestCd) & _
              " AND    b.cdval1  =  a.testcd " & _
              " AND    c.testcd  =  a.testcd   " & _
              " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
              " AND    d.cdval1  =  b.field2 " & _
              " AND    " & DBJ("f." & F_DOCTID & " =* a.majdoct ")
                    
    strSql6 = " SELECT  distinct " & FUNC_SUBSTR & "(b.rptdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.rptdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(b.rptdt,7,2) as OrdDate, " & _
              "         1 ordno, 0 ordseq, null ordcd, '임상병리 종합검증/판독' testnm, null panelfg, null grpfg, null spccd, null statfg, " & _
              "         g.empnm as DoctNm, null rcvdt, null RcvTm, '8' stscd, null dcfg, null mesg, " & _
              "         null SpcNm, null as MultiFg, a.bedindt as orddt, 'C' orddiv, a.wardid, a.hosilid, null OrdTm, " & _
              "         " & FUNC_SUBSTR & "(a.rptdt,1,4)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,7,2) as ExamDt, " & _
              "         null ExamTm, g.empnm as ExamNm, '' as orddtm, " & _
              "         null workarea, null accdt, 0 accseq, null  RstIn, '' testdiv, '*' as Verified, '' spcyy, 0 spcno, 0 unitqty, '' reqdt, '' reqtm, Null majdoct, null as rstseq  " & _
              " FROM    " & T_LAB501 & " a, " & T_LAB502 & " b, " & T_LAB015 & " g " & _
              " WHERE   " & DBW("a.ptid = ", PtId) & _
              " AND     " & DBW("a.bedindt>=", FromDate) & _
              " AND     " & DBW("a.bedindt<=", ToDate) & _
              " AND     a.donefg = '2' " & _
              " AND     b.ptid  =  a.ptid" & _
              " AND     b.rptdt =  a.rptdt " & _
              " AND     " & DBJ("g.empid =* a.rptid")
    
    Select Case pResultReviewFlag
        Case 0:
            SqlQueryOrders_WM = strSql1 & " UNION ALL " & strSql3 & " UNION ALL " & strSql6        '& " UNION ALL " & strSql4
        Case 1: SqlQueryOrders_WM = strSql1    '임상
        Case 2: SqlQueryOrders_WM = strSql2    '해부
        Case 3: SqlQueryOrders_WM = strSql1  '임상/해부
        Case 4: SqlQueryOrders_WM = strSql3  '혈액
        Case 5: SqlQueryOrders_WM = strSql1  '임상/혈액
        Case 6: SqlQueryOrders_WM = strSql1  '해부/혈액
        Case 7: SqlQueryOrders_WM = strSql5  '미생물
    End Select
    SqlQueryOrders_WM = SqlQueryOrders_WM & " Order  by orddt desc , ordno, ordseq "

End Function

'%  32-1. 연속검사인 경우 환자Id, 처방일(채혈일)을 기준으로 연속검사내역을 검색한다.
'%       - Lab203 : 연속검사내역
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlMultiTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long) As String
    SqlMultiTest = " SELECT  distinct a.workarea, a.accdt, a.accseq, a.spcseq " & _
                   " FROM   " & T_LAB203 & " a, " & T_LAB203 & " b " & _
                   " WHERE  " & DBW("b.workarea", WorkArea, 2) & _
                   " AND    " & DBW("b.accdt", AccDt, 2) & _
                   " AND    " & DBW("b.accseq", accseq, 2) & _
                   " AND    a.ptid = b.ptid " & _
                   " AND    a.orddt = b.orddt " & _
                   " AND    a.ordno = b.ordno " & _
                   " AND    a.ordseq = b.ordseq " & _
                   " ORDER BY spcseq "

End Function


Public Function SqlGetWsUnit(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long) As String
    SqlGetWsUnit = " SELECT  distinct a.wscd, a.wsunit " & _
                   " FROM   " & T_LAB402 & " a " & _
                   " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
                   " AND    " & DBW("a.accdt", AccDt, 2) & _
                   " AND    " & DBW("a.accseq", accseq, 2)

End Function


'%  33. Lab No를 기준으로 결과내역을 검색한다.
'%       - Lab302 : 결과내역
'%       - Lab001 : 검사항목 마스터
'%       - Lab303 : Text 결과 내역
'%       - Lab031 : 결과코드명 (C110-Field1)
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlQueryResults(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long, ByVal TestDiv As String, _
                                Optional ByVal Sex As Variant, Optional ByVal AgeDay As Variant) As String
    
    Dim strCond As String
    
    Select Case TestDiv
        Case "0":
            '## 5.0.13: 이상대(2004-12-30)
            '   - 최근결과가 검사항목별 결과코드에 등록된 코드이면 결과명도 조회
            SqlQueryResults = " SELECT  a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.statfg, a.rstdiv, " & _
                              "         a.lastrst, h.field1 as lastrstval, " & FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                              "         a.lastvfyid, e.empnm as LastVfyNm, a.vfydt, a.vfytm, a.vfyid, f.empnm as VfyNm, " & _
                              "         a.attrcd, a.mfyfg, a.grpfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, " & _
                              "         a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                              "         c.rsttxt as TextResult, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, g.text1 notice,g.text2 as applydate " & _
                              " FROM    " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB015 & " e, " & T_LAB034 & " g, " & _
                                            T_LAB015 & " f, " & T_LAB031 & " h, " & T_LAB303 & " c, " & T_LAB302 & " a " & _
                              " WHERE   " & DBW("a.workarea", WorkArea, 2) & _
                              " AND     " & DBW("a.accdt", AccDt, 2) & _
                              " AND     " & DBW("a.accseq", accseq, 2) & _
                              " AND     b.testcd = a.testcd " & _
                              " AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
                              " AND     " & DBJ("c.workarea =* a.workarea") & _
                              " AND     " & DBJ("c.accdt    =* a.accdt") & _
                              " AND     " & DBJ("c.accseq   =* a.accseq") & _
                              " AND     " & DBJ("c.testcd   =* a.testcd") & _
                              " AND     " & DBJ(DBW("d.cdindex", LC2_ItemResult, 2)) & _
                              " AND     " & DBJ("d.cdval1   =* a.testcd") & _
                              " AND     " & DBJ("d.cdval2   =* a.rstcd") & _
                              " AND     " & DBJ("e.empid    =* a.lastvfyid") & _
                              " AND     " & DBJ("f.empid    =* a.vfyid") & _
                              " AND     " & DBJ(DBW("g.cdindex", LC4_ClinicalNotice, 2)) & _
                              " AND     " & DBJ("g.cdval1   =* a.testcd")
            SqlQueryResults = SqlQueryResults & _
                              " AND     " & DBJ(DBW("h.cdindex=*", LC2_ItemResult)) & _
                              " AND     " & DBJ("a.testcd*=h.cdval1") & _
                              " AND     " & DBJ("a.lastrst*=h.cdval2") & _
                              " ORDER BY RptSeq, a.testcd "
        Case "1":
            SqlQueryResults = " SELECT  a.testcd, a.rsttype, a.valfg, a.txtfg, a.mfyseq, a.stscd, " & _
                              "         a.vfydt, a.vfytm, a.vfyid, e.empnm as VfyNm, " & _
                              "         a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                              "         d.txtrst as TextResult " & _
                              " FROM    " & T_LAB001 & " b, " & T_LAB353 & " d, " & T_LAB015 & " e, " & T_LAB351 & " a " & _
                              " WHERE   " & DBW("a.workarea", WorkArea, 2) & _
                              " AND     " & DBW("a.accdt", AccDt, 2) & _
                              " AND     " & DBW("a.accseq", accseq, 2) & _
                              " AND     b.testcd = a.testcd " & _
                              " AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
                              " AND     " & DBJ("d.workarea =* a.workarea") & _
                              " AND     " & DBJ("d.accdt    =* a.accdt") & _
                              " AND     " & DBJ("d.accseq   =* a.accseq") & _
                              " AND     " & DBJ("d.testcd   =* a.testcd") & _
                              " AND     " & DBJ("d.mfyseq   =* a.mfyseq") & _
                              " AND     " & DBJ("e.empid    =* a.vfyid") & _
                              " ORDER BY testcd "
        Case "2":
            
            SqlQueryResults = " SELECT  a.testcd, a.rstcd, a.lastrst, " & FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                              "         a.rstdiv, a.senfg, a.stscd, a.lastvfyid, e.empnm as LastVfyNm, a.vfydt, a.vfytm, a.vfyid, f.empnm as VfyNm, " & _
                              "         a.mfyseq, a.rsttype, a.detailfg, a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                              "         c.rsttxt as TextResult, d.field1 as RstCdNm1, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, '' as RstCdNm2 "  'g.field1 as RstCdNm2 "
            SqlQueryResults = SqlQueryResults & " FROM   " & T_LAB001 & " b, " & T_LAB303 & " c, " & T_LAB031 & " d, " & _
                                                             T_LAB015 & " e, " & T_LAB015 & " f, " & T_LAB404 & " a "     '& T_LAB031 & " g "
            SqlQueryResults = SqlQueryResults & _
                              " WHERE   " & DBW("a.workarea", WorkArea, 2) & _
                              " AND     " & DBW("a.accdt", AccDt, 2) & _
                              " AND     " & DBW("a.accseq", accseq, 2) & _
                              " AND     b.testcd = a.testcd " & _
                              " AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
                              " AND     " & DBJ("c.workarea =* a.workarea") & _
                              " AND     " & DBJ("c.accdt    =* a.accdt") & _
                              " AND     " & DBJ("c.accseq   =* a.accseq") & _
                              " AND     " & DBJ("c.testcd   =* a.testcd") & _
                              " AND     " & DBJ(DBW("d.cdindex", LC2_ItemResult, 2)) & _
                              " AND     " & DBJ("d.cdval1   =* a.testcd") & _
                              " AND     " & DBJ("d.cdval2   =* a.rstcd") & _
                              " AND     " & DBJ("e.empid    =* a.lastvfyid") & _
                              " AND     " & DBJ("f.empid    =* a.vfyid") & _
                              " ORDER BY RptSeq, a.testcd "
    End Select
End Function

'%  34. Text Result
'%       - Lab303  : Text결과 내역 검색
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlGetRstText(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long, ByVal TestCd As String) As String
    SqlGetRstText = " SELECT a.rsttxt " & _
                    " FROM   " & T_LAB303 & " a " & _
                    " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
                    " AND    " & DBW("a.accdt", AccDt, 2) & _
                    " AND    " & DBW("a.accseq", accseq, 2) & _
                    " AND    " & DBW("a.testcd", TestCd, 2)
End Function

'%  34-1. Supplemental Report
'%       - 305 : Text결과 수정내역 검색
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlGetSuppText(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long, ByVal TestCd As String) As String
    SqlGetSuppText = " SELECT a.rsttxt, b.empnm, a.seq, " & FUNC_SUBSTR & "(a.mfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.mfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.mfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.mfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.mfytm,3,2) as MfyDtTm " & _
                     " FROM   " & T_LAB305 & " a, " & T_LAB015 & " b " & _
                     " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
                     " AND    " & DBW("a.accdt", AccDt, 2) & _
                     " AND    " & DBW("a.accseq", accseq, 2) & _
                     " AND    " & DBW("a.testcd", TestCd, 2) & _
                     " AND    " & DBJ("b.empid =* a.mfyid") & _
                     " ORDER BY seq "
End Function

'%  34-2. Supplemental Report
'%       - Lab353 : 기타검사 Text결과 수정내역 검색
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlETextRst(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, ByVal pTestCd As String) As String
    SqlETextRst = " SELECT txtrst as TextResult FROM " & T_LAB353 & _
                  " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                  " AND    " & DBW("accdt", pAccDt, 2) & _
                  " AND    " & DBW("accseq", pAccSeq, 2) & _
                  " AND    " & DBW("testcd", pTestCd, 2) & _
                  " AND    " & DBW("mfyseq > ", "0")
End Function

'%  35. 미생물 감수성결과 (LAB405)
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlSenResult(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long, ByVal TestCd As String) As String
    SqlSenResult = " SELECT a.seq, b.text1 as MicroNm, c.field2 as GrowthQty, a.micfg, a.scnt, a.srst1, a.srst2, a.srst3, a.srst4, a.srst5, a.srst6, a.srst7, " & _
                   "        a.srst8, a.srst9, a.srst10, a.srst11, a.srst12, a.srst13, a.srst14, a.srst15, a.srst16, a.srst17, a.srst18, a.srst19, a.srst20, " & _
                   "        a.srst21, a.srst22, a.srst23, a.srst24, a.srst25 " & _
                   " FROM   " & T_LAB404 & " d, " & T_LAB405 & " a, " & T_LAB032 & " b, " & T_LAB032 & " c " & _
                   " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
                   " AND    " & DBW("d.accdt", AccDt, 2) & _
                   " AND    " & DBW("d.accseq", accseq, 2) & _
                   " AND    " & DBW("d.testcd", TestCd, 2) & _
                   " AND    a.workarea = d.workarea " & _
                   " AND    a.accdt = d.accdt " & _
                   " AND    a.accseq = d.accseq " & _
                   " AND    a.testcd = d.testcd  " & _
                   " AND    a.mfyseq = d.mfyseq  " & _
                   " AND    " & DBJ(DBW("b.cdindex", LC3_Microbe, 2)) & _
                   " AND    " & DBJ("b.cdval1 =* a.mnmcd") & _
                   " AND    " & DBJ(DBW("c.cdindex", LC3_Volume, 2)) & _
                   " AND    " & DBJ("c.cdval1 =* a.mqtcd") & _
                   " ORDER BY seq "
End Function


'%  7-1. Get Reference Range for a Patient
'%      - Calling FROM [frm401ResultReview]
Public Function SqlGetReference(ByVal strTestCd As String, _
                                              ByVal strSpcCd As String, _
                                              ByVal strApplyDt As String, _
                                              ByVal strApplySex As String, _
                                              Optional ByVal intAge As Variant, _
                                              Optional ByVal intAgeFrom As Variant, _
                                              Optional ByVal intAgeTo As Variant) As String


' 2012.04.12. 참고치값을 읽어오는데 폐기된거 제외


' 2009.04.01. 양성현 참고치값을 읽어오는데 초기참고치만 읽어오는 문제가 있어서 하기의 내용을 추가함.
' " and applydt >= ( SELECT MAX(c.applydt) FROM  " & T_LAB005 & " c WHERE " & DBW("c.testcd", strTestCd, 2) & " AND " & DBW("c.spccd", strSpcCd, 2) & " AND " & DBW("applydt < ", strApplyDt) & ") AND "
 

'    SqlGetReference = " SELECT * " & _
'                      " FROM  " & T_LAB005 & " " & _
'                      " WHERE " & DBW("testcd", strTestCd, 2) & _
'                      " AND   " & DBW("spccd", strSpcCd, 2) & _
'                      " AND   " & DBW("applydt <= ", strApplyDt) & _
'                      " AND   " & DBW("applysex", strApplySex, 2)
'
'    If Not (IsMissing(intAgeFrom) Or IsMissing(intAgeTo)) Then
'       SqlGetReference = SqlGetReference & _
'                      " AND   " & DBW("agefrom", intAgeFrom, 2) & _
'                      " AND   " & DBW("ageto", intAgeTo, 2)
'    ElseIf Not IsMissing(intAge) Then
'       SqlGetReference = SqlGetReference & _
'                      " AND   " & DBW("agefrom <= ", intAge) & _
'                      " AND   " & DBW("ageto >= ", intAge)
'    End If
'
'    SqlGetReference = SqlGetReference & " AND (expdt is null or expdt = '') "
'    SqlGetReference = SqlGetReference & " ORDER BY applydt desc"
'
   
    SqlGetReference = " SELECT * " & _
                      " FROM  " & T_LAB005 & " " & _
                      " WHERE " & DBW("testcd", strTestCd, 2) & _
                      " AND   " & DBW("spccd", strSpcCd, 2) & _
                      " AND   " & DBW("applydt <= ", strApplyDt) & _
                      " and applydt >= ( SELECT MAX(c.applydt) FROM  " & T_LAB005 & " c WHERE " & DBW("c.testcd", strTestCd, 2) & _
                      " AND " & DBW("c.spccd", strSpcCd, 2) & " AND " & DBW("applydt < ", strApplyDt) & ") AND " & DBW("applysex", strApplySex, 2)

    If Not (IsMissing(intAgeFrom) Or IsMissing(intAgeTo)) Then
       SqlGetReference = SqlGetReference & _
                      " AND   " & DBW("agefrom", intAgeFrom, 2) & _
                      " AND   " & DBW("ageto", intAgeTo, 2)
    ElseIf Not IsMissing(intAge) Then
       SqlGetReference = SqlGetReference & _
                      " AND   " & DBW("agefrom <= ", intAge) & _
                      " AND   " & DBW("ageto >= ", intAge)
    End If
    
    SqlGetReference = SqlGetReference & " AND (expdt is null or expdt = '') "
    SqlGetReference = SqlGetReference & " ORDER BY applydt desc"


End Function
   

'%  37. 같은 Lab No 중 확인된 결과존재 여부 (LAB302)
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlVerifiedResult1(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String

    SqlVerifiedResult1 = " SELECT * FROM " & T_LAB302 & _
                         " WHERE " & DBW("workarea", pWorkArea, 2) & _
                         " AND   " & DBW("accdt", pAccDt, 2) & _
                         " AND   " & DBW("accseq", pAccSeq, 2) & _
                         " AND   vfydt <> ' ' " '& _
                         " AND   " & DBW("vfydt <> ", "")

End Function

'%  38. 같은 처방번호,처방Seq로 발생한 항목중 결과확인 상태인 데이타 존재여부 (LAB302)
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlVerifiedResult2(ByVal pPtid As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, ByVal pOrdSeq As Integer) As String
            
    SqlVerifiedResult2 = " SELECT * FROM " & T_LAB302 & _
                         " WHERE " & DBW("ptid", pPtid, 2) & _
                         " AND   " & DBW("orddt", pOrdDt, 2) & _
                         " AND   " & DBW("ordno", pOrdNo, 2) & _
                         " AND   " & DBW("ordseq", pOrdSeq, 2) & _
                         " AND   vfydt <> ' ' " '& _
                         " AND   " & DBW("vfydt <> ", "")

End Function


'%  39. Remark
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlGetRemark(ByVal pRmkCd As String) As String

    SqlGetRemark = " SELECT text1 as Remark FROM " & T_LAB034 & _
                   " WHERE " & DBW("cdindex", LC4_Remark, 2) & _
                   " AND   " & DBW("cdval1", pRmkCd, 2)
End Function


'%  40. Foot Note
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlGetFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String
   
    SqlGetFootNote = " SELECT rsttxt as FootNote " & _
                     " FROM  " & T_LAB304 & _
                     " WHERE " & DBW("workarea", pWorkArea, 2) & _
                     " AND   " & DBW("accdt", pAccDt, 2) & _
                     " AND   " & DBW("accseq", pAccSeq, 2)

End Function


'%  41. 기타검사 Value Result
'%       - Calling FROM [ frm401ResultView ] : 결과조회
Public Function SqlEValueRst(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pMfyFg As String) As String

    SqlEValueRst = " SELECT seq, head, valrst, unit " & _
                   " FROM  " & T_LAB352 & _
                   " WHERE " & DBW("workarea", pWorkArea, 2) & _
                   " AND   " & DBW("accdt", pAccDt, 2) & _
                   " AND   " & DBW("accseq", pAccSeq, 2) & _
                   " AND   " & DBW("testcd", pTestCd, 2) & _
                   " AND   " & DBW("mfyseq", pMfyFg, 2) & _
                   " ORDER BY seq "

End Function
   
Public Function SqlQueryAllResults(ByVal PtId As String, ByVal QueryKey As String, ByVal FromDate As String, ByVal ToDate As String, _
                                   ByVal TestDiv As String, Optional ByVal Sex As Variant, Optional ByVal AgeDay As Variant, _
                                   Optional ByVal pReportFg As Boolean = False) As String

    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim Result4 As String   '해부병리
    Dim Result5 As String   '혈액은행
    Dim Result6 As String   'POC
    Dim Result7 As String   '종합검증
    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim tmpCond As String
    Dim tmpNotice As String
    
'    If dbconn.Whatsthis = dbconn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then ' 사이베이스
        MsgBox "사이베이스용으로 컨버트 하는 부분 수정 요함", vbCritical
'        tmpNotice = FUNC_CONVERT("varchar(255)", "e.text1") & " notice, " & FUNC_CONVERT("varchar(255)", "e.text2") & " applydate, "
    Else
        tmpNotice = " e.text1 notice,e.text2 as applydate, "
    End If
    
    '급한대로 Sex를 재발행여부로 사용하자...나중에 고치고...2001.7.6.김미경
    
    tmpStr = "  " & FUNC_SUBSTR & "(x." & QueryKey & ",3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x." & QueryKey & ",5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x." & QueryKey & ",7,2) as KeyDate "
    If pReportFg Then
        tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                  "  AND (a.rptfg = ''  or  a.rptfg is null) "
    Else
        tmpCond = ""
    End If

    If pReportFg Then
        tmpOrderby = " ORDER BY KeyDate, workarea, accdt, accseq, RptSeq, ordcd "
    Else
        tmpOrderby = " ORDER BY KeyDate desc, workarea, accdt, accseq, RptSeq, ordcd "
    End If

    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x." & QueryKey & " >= '" & FromDate & "' AND  x." & QueryKey & "<= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    
    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice,null applydate, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x." & QueryKey & " >= '" & FromDate & "' AND  x." & QueryKey & "<= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd")
                   '"AND      f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    
    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, null applydate," & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x." & QueryKey & " >= '" & FromDate & "' AND  x." & QueryKey & "<= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, null applydate, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x." & QueryKey & " >= '" & FromDate & "' AND  x." & QueryKey & "<= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    Result5 = " SELECT x.workarea as WorkArea, x.accdt as AccDt, x.accseq as AccSeq, x.ordcd testcd, '' rstcd, '' rstunit, " & _
              "        '' hldiv, '' dpdiv, '' spccd, '' rstdiv, x.statfg, '' lastrst, " & DBV("lastvfyid", "0") & " as lastvfyid, x.examdt VfyDt, " & _
              "        x.examtm VfyTm, x.examdoct VfyId, '5' as testdiv, '' footnotefg, '' rmkcd, " & _
              "        '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, '혈액' SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, x.ptid, x.orddt, x.ordno, x.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, '' DetailItem, " & _
              "        k.deptcd, '' sex, 0 ageday, k.orddoct, c.colid, c.rcvid, null notice, null applydate," & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        x.unitqty, k.reqdt, k.reqtm, '' as RstCdNm, 0 as RptSeq, " & tmpStr & _
              " FROM " & T_BBS001 & " b, " & T_BBS201 & " c, " & T_LAB101 & " k, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x." & QueryKey & " >= '" & FromDate & "' AND  x." & QueryKey & "<= '" & ToDate & "' " & _
              " AND    c.ptid = x.ptid " & _
              " AND    c.spcyy = (SELECT max(spcyy) FROM " & T_BBS201 & " WHERE ptid = x.ptid) " & _
              " AND    c.spcno = (SELECT max(spcno) FROM " & T_BBS201 & _
                                " WHERE ptid = c.ptid AND spcyy = c.spcyy) " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = x.ordcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_BBS001 & " WHERE testcd = x.ordcd AND applydt <= x.orddt ) "
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    
    Result6 = " SELECT '' as WorkArea, '' as AccDt, 0 as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        '' hldiv, '' dpdiv, c.field2 spccd, '' rstdiv, '' statfg, '' lastrst, null lastvfyid, a.vfydt, " & _
              "        a.vfytm, a.vfyid, '9' as testdiv, '' footnotefg, '' rmkcd, " & _
              "        '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        '' dcfg, '' mesg, a.ptid, a.vfydt orddt, 99 ordno, 1 ordseq, a.testcd ordcd, 'P' orddiv, " & _
              "        a.wardid, '' hosilid, '' roomid, " & _
                       FUNC_SUBSTR & "(b.testnm,1,16) " & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        a.deptcd, '' sex, 0 ageday, a.majdoct orddoct, a.vfyid colid, a.vfyid rcvid, null notice,null applydate, " & _
                       FUNC_SUBSTR & "(a.vfydt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
              "        0 unitqty, '' reqdt, '' reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as KeyDate " & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " c, " & T_LAB195 & " a " & _
              " WHERE  " & DBW("a.ptid = ", PtId) & _
              " AND    a.vfydt >= '" & FromDate & "' AND  a.vfydt <= '" & ToDate & "' " & _
              " AND    a.rstcd  is not null " & _
              " AND    a.rstcd  <> ' ' " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.vfydt ) " & _
              " AND    " & DBW("c.cdindex = ", LC3_POCTestCd) & _
              " AND    c.cdval1  =  a.testcd " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* c.field2")
      
    Result7 = " SELECT '' as WorkArea, '' as AccDt, 0 as AccSeq, '' testcd, '' rstcd, '' rstunit, " & _
              "        '' hldiv, '' dpdiv, '' spccd, '' rstdiv, '' statfg, '' lastrst, null lastvfyid, a.rptdt vfydt, " & _
              "        '' vfytm, a.rptid vfyid, '8' as testdiv, '' footnotefg, '' rmkcd, " & _
              "        '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, '' SpcNm, '' as SenFg, " & _
              "        '' dcfg, '' mesg, a.ptid, a.bedindt as  orddt, 1 ordno, 1 ordseq, '' ordcd, 'C' orddiv, " & _
              "        a.wardid, a.hosilid, '' roomid, " & _
              "        '임상병리 종합검증/판독' as TestLongNm, '종합검증' as TestShortNm, '' as DetailItem, " & _
              "        a.deptcd, '' sex, 0 ageday, null orddoct, null colid, null rcvid, null notice,null applydate, " & _
                       FUNC_SUBSTR & "(a.bedindt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.bedindt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.bedindt,7,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
              "        '' as ColDtTm, " & _
              "        '' as RcvDtTm, " & _
                       FUNC_SUBSTR & "(a.rptdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,7,2)  as VfyDtTm, " & _
              "        0 unitqty, '' reqdt, '' reqtm, '' as RstCdNm, 1 as RptSeq, " & _
                       FUNC_SUBSTR & "(a.rptdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.rptdt,7,2) as KeyDate " & _
              " FROM " & T_LAB502 & " b, " & T_LAB501 & " a " & _
              " WHERE  " & DBW("a.ptid = ", PtId) & _
              " AND    " & DBW("a.bedindt >=", FromDate) & _
              " AND    " & DBW("a.bedindt <=", ToDate) & _
              " AND    a.donefg  = '2' " & _
              " AND     b.ptid  =  a.ptid" & _
              " AND     b.rptdt =  a.rptdt "
      
    Select Case TestDiv
        Case "0":  SqlQueryAllResults = result1 & " UNION ALL " & Result6     '일반결과만..
        Case "1":  SqlQueryAllResults = result2    '특수검사결과만..
        Case "2":  SqlQueryAllResults = result3    '미생물결과만..
        
        Case "3":  '모두...
            Select Case P_ResultReviewFlag  'S2LIS_Const의 clsLISParameters에 정의되어 있음
                Case 0:
                        SqlQueryAllResults = result1 & " UNION ALL " & result2 & " UNION ALL " & Result7 & " UNION ALL " & _
                                             result3 & " UNION ALL " & Result5     '임상/해부/혈액
                Case 1: SqlQueryAllResults = result1 & " UNION ALL " & result2 & " UNION ALL " & result3 & " UNION ALL " & Result7    '임상
                Case 2: SqlQueryAllResults = Result4    '해부
                Case 3: SqlQueryAllResults = result1 & " UNION ALL " & result2 & " UNION ALL " & Result7 & " UNION ALL " & _
                                             result3    '임상/해부
                Case 4: SqlQueryAllResults = Result5    '혈액
                Case 5: SqlQueryAllResults = result1 & " UNION ALL " & result2 & " UNION ALL " & _
                                             result3 & " UNION ALL " & Result7 & " UNION ALL " & Result5  '임상/혈액
                Case 6: SqlQueryAllResults = Result5   '해부/혈액
            End Select
    End Select
    SqlQueryAllResults = SqlQueryAllResults & tmpOrderby
                           
End Function

Public Function SqlResultsForReport(ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String, _
                                    ByVal TestDiv As String, Optional ByVal pReprintFg As Boolean = False) As String

    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim Result4 As String   '해부병리
    Dim Result5 As String   '혈액은행
    Dim Result6 As String   'POC
    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim tmpCond As String
    
    
    'tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    tmpStr = "  " & FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(c.coldt,7,2) as KeyDate "
    
    If pReprintFg Then  '재발행
        tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                  "  AND  a.rptfg <> ' '  AND  a.rptfg is not null "
    Else    '신규
        tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                  "  AND (a.rptfg = ''    or   a.rptfg is null) "
    End If
    
    tmpOrderby = " ORDER BY KeyDate, ColDtTm, workarea, accdt, accseq, RptSeq, ordcd "

    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd")
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd")
                   '"AND      f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    Result6 = " SELECT '' as WorkArea, '' as AccDt, 0 as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        '' hldiv, '' dpdiv, c.field2 spccd, '' rstdiv, '' statfg, '' lastrst, null lastvfyid, a.vfydt, " & _
              "        a.vfytm, a.vfyid, '9' as testdiv, '' footnotefg, '' rmkcd, " & _
              "        '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        '' dcfg, '' mesg, a.ptid, a.vfydt orddt, 99 ordno, 1 ordseq, a.testcd ordcd, 'P' orddiv, " & _
              "        a.wardid, '' hosilid, '' roomid, " & _
                       FUNC_SUBSTR & "(b.testnm,1,16) " & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        a.deptcd, '' sex, 0 ageday, a.majdoct orddoct, a.vfyid colid, a.vfyid rcvid, " & _
                       FUNC_SUBSTR & "(a.vfydt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
              "        0 unitqty, '' reqdt, '' reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as KeyDate " & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " c, " & T_LAB195 & " a " & _
              " WHERE  " & DBW("a.ptid = ", PtId) & _
              " AND    a.vfydt >= '" & FromDate & "' AND  a.vfydt <= '" & ToDate & "' " & _
              " AND    a.rstcd  is not null " & _
              " AND    a.rstcd  <> ' ' " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.vfydt ) " & _
              " AND    " & DBW("c.cdindex = ", LC3_POCTestCd) & _
              " AND    c.cdval1  =  a.testcd " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* c.field2")
      
    Select Case TestDiv
        Case "0":  SqlResultsForReport = result1 & " UNION ALL " & Result6     '일반결과만..
        Case "1":  SqlResultsForReport = result2    '특수검사결과만..
        Case "2":  SqlResultsForReport = result3    '미생물결과만..
    End Select
    
    SqlResultsForReport = SqlResultsForReport & tmpOrderby
                           
End Function

Public Function SqlQueryERForReport(ByVal pPtid As String, ByVal pVfyDt As String, _
                                    ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, _
                                    Optional ByVal pReprintFg As Boolean = False, _
                                    Optional ByVal pBatchFg As Boolean = False, _
                                    Optional ByVal pWardMenu As Boolean = False) As String

    Dim result1 As String
    Dim tmpOrderby As String
    Dim tmpNotice As String
    Dim tmpStr As String
    Dim tmpCond As String
    Dim strQuery As String

'    SELECT Case
'        Case "05"
'            If mvarRiPrint = "" Then
'                strQuery = " AND a.workarea not in('OT','OR','RI')"
'            Else
'                strQuery = " AND a.workarea  in('OR','RI')"
'            End If
'    End SELECT
    
    tmpOrderby = " ORDER BY KeyDate, Seq, workarea, accdt, accseq, RptSeq, ordcd "
    tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    
    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        '5' stscd, a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid", pPtid, 2) & _
              " AND    " & DBW("x.examdt", pVfyDt, 2) & " AND    " & DBW("x.workarea", pWorkArea, 2) & " AND " & DBW("x.accdt", pAccDt, 2) & " AND " & DBW("x.accseq", pAccSeq, 2) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
'    result1 = result1 & " AND a.testcd not in(" & P_NoResultReport & ")"
    result1 = result1 & strQuery
    
    result1 = result1 & " AND " & DBW("a.testcd<>", P_NoResultReport)
   
   
    SqlQueryERForReport = result1
    
    SqlQueryERForReport = SqlQueryERForReport & tmpOrderby
                           
End Function

Public Function SqlQueryImageForReport(ByVal pPtid As String, ByVal pVfyDt As String, _
                                    ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, ByVal pTestDiv As String) As String

    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim Result4 As String   '해부병리
    Dim Result5 As String   '혈액은행
    Dim Result6 As String   'POC
    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim tmpCond As String
    Dim tmpNotice As String
     Dim strQuery As String
     
'    If dbconn.Whatsthis = dbconn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then 'sybase
        MsgBox "사이베이스용으로 소스 수정 필요", vbCritical
'        tmpNotice = FUNC_CONVERT("varchar(255)", "e.text1") & " notice, "
    Else
        tmpNotice = " e.text1 notice, "
    End If
   
'    SELECT Case
'        Case "05"
'            If mvarRiPrint = "" Then
'                strQuery = " AND a.workarea not in('OT','OR','RI')"
'            Else
'                strQuery = " AND a.workarea  in('OR','RI')"
'            End If
'    End SELECT
    
    tmpOrderby = " ORDER BY KeyDate, Seq, workarea, accdt, accseq, RptSeq, ordcd "
    tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    

    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        x.stscd, a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid", pPtid, 2) & _
              " AND    " & DBW("x.examdt", pVfyDt, 2) & " AND    " & DBW("x.workarea", pWorkArea, 2) & " AND " & DBW("x.accdt", pAccDt, 2) & " AND " & DBW("x.accseq", pAccSeq, 2) & " AND " & DBW("x.ordcd", mvarTestCd, 2) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    
    If P_NoResultReport <> "" Then
        result1 = result1 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result1 = result1 & strQuery
    End If
    
    'result1 = result1 & " AND " & DBW("a.testcd<>", P_NoResultReport)
    '기타검사
    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid", pPtid, 2) & _
              " AND    " & DBW("x.examdt", pVfyDt, 2) & " AND    " & DBW("x.workarea", pWorkArea, 2) & " AND " & DBW("x.accdt", pAccDt, 2) & " AND " & DBW("x.accseq", pAccSeq, 2) & " AND " & DBW("x.ordcd", mvarTestCd, 2) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea")
                   '"AND      f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result2 = result2 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result2 = result2 & strQuery
    End If
    
    '미생물결과
    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, a. stscd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid,  " & tmpNotice & " " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid", pPtid, 2) & _
              " AND    " & DBW("x.examdt", pVfyDt, 2) & " AND    " & DBW("x.workarea", pWorkArea, 2) & " AND " & DBW("x.accdt", pAccDt, 2) & " AND " & DBW("x.accseq", pAccSeq, 2) & " AND " & DBW("x.ordcd", mvarTestCd, 2) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result3 = result3 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result3 = result3 & strQuery
    End If
    '미생물결과
    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid,   " & tmpNotice & " " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid", pPtid, 2) & _
              " AND    " & DBW("x.examdt", pVfyDt, 2) & " AND    " & DBW("x.workarea", pWorkArea, 2) & " AND " & DBW("x.accdt", pAccDt, 2) & " AND " & DBW("x.accseq", pAccSeq, 2) & " AND " & DBW("x.ordcd", mvarTestCd, 2) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result3 = result3 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result3 = result3 & strQuery
    End If
    
  
      
    Select Case pTestDiv
        Case "0":  SqlQueryImageForReport = result1   '일반결과만..
        Case "1":  SqlQueryImageForReport = result2    '특수검사결과만..
        Case "2":  SqlQueryImageForReport = result3    '미생물결과만..
    End Select
    
    SqlQueryImageForReport = SqlQueryImageForReport & tmpOrderby
                           
End Function


Public Function SqlQueryForReport(ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String, _
                                  ByVal TestDiv As String, Optional ByVal pReprintFg As Boolean = False, _
                                  Optional ByVal pBatchFg As Boolean = False, _
                                  Optional ByVal pWardMenu As Boolean = False) As String

    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim Result4 As String   '해부병리
    Dim Result5 As String   '혈액은행
    Dim Result6 As String   'POC
    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim tmpCond As String
    Dim tmpNotice As String
    
'    If dbconn.Whatsthis = dbconn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then 'sybase
        MsgBox "사이베이스용으로 소스 수정 필요", vbCritical
'        tmpNotice = FUNC_CONVERT("varchar(255)", "e.text1") & " notice, "
    Else
        tmpNotice = " e.text1 notice, "
    End If
    
    tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    If pWardMenu Then
        tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null "
    Else
        If pBatchFg Then    '일괄재발행(결과출력된것만...)
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                      "  AND  a.rptfg <> ' '  AND  a.rptfg is not null  AND " & DBW("a.rptdt = ", ToDate)
        ElseIf pReprintFg Then  '재발행
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null "
        Else    '신규(결과출력안된것만...)
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                      "  AND (a.rptfg = ''    or   a.rptfg is null) "
        End If
        
    End If
    
    Dim strQuery As String

'    SELECT Case
'        Case "05"
'            If mvarRiPrint = "" Then
'                strQuery = " AND a.workarea not in('OT','OR','RI')"
'            Else
'                strQuery = " AND a.workarea  in('OR','RI')"
'            End If
'    End SELECT
    tmpOrderby = " ORDER BY KeyDate, Seq, workarea, accdt, accseq, RptSeq, ordcd "

    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        '5' stscd, a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    
    If P_NoResultReport <> "" Then
        result1 = result1 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result1 = result1 & strQuery
    End If
    
    'result1 = result1 & " AND " & DBW("a.testcd<>", P_NoResultReport)
    '기타검사
    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea")
                   '"AND      f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result2 = result2 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result2 = result2 & strQuery
    End If
    
    '미생물결과
    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, a. stscd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid,  " & tmpNotice & " " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result3 = result3 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result3 = result3 & strQuery
    End If
    '미생물결과
    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid,   " & tmpNotice & " " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result3 = result3 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result3 = result3 & strQuery
    End If
    
    '혈당검사 (성모자애 병원에서만 사용하구 있음)
    Result6 = " SELECT '' as WorkArea, '' as AccDt, 0 as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        '' hldiv, '' dpdiv, c.field2 spccd, '' rstdiv, '' statfg, '' lastrst, null lastvfyid, a.vfydt, " & _
              "        a.vfytm, a.vfyid, '9' as testdiv, '' footnotefg, '' rmkcd, '5'  stscd, " & _
              "        '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        '' dcfg, '' mesg, a.ptid, a.vfydt orddt, 99 ordno, 1 ordseq, a.testcd ordcd, 'P' orddiv, " & _
              "        a.wardid, '' hosilid, '' roomid, " & _
                       FUNC_SUBSTR & "(b.testnm,1,16) " & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        99 Seq, a.deptcd, '' sex, 0 ageday, a.majdoct orddoct, a.vfyid colid, a.vfyid rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(a.vfydt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
              "        0 unitqty, '' reqdt, '' reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as KeyDate " & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " c, " & T_LAB195 & " a " & _
              " WHERE  " & DBW("a.ptid = ", PtId) & _
              " AND    a.vfydt >= '" & FromDate & "' AND  a.vfydt <= '" & ToDate & "' " & _
              " AND    a.rstcd  is not null " & _
              " AND    a.rstcd  <> ' ' " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.vfydt ) " & _
              " AND    " & DBW("c.cdindex = ", LC3_POCTestCd) & _
              " AND    c.cdval1  =  a.testcd " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* c.field2")
      
    Select Case TestDiv
        Case "0":  SqlQueryForReport = result1 & " UNION ALL " & Result6     '일반결과만..
        Case "1":  SqlQueryForReport = result2    '특수검사결과만..
        Case "2":  SqlQueryForReport = result3    '미생물결과만..
    End Select
    
    SqlQueryForReport = SqlQueryForReport & tmpOrderby
                           
End Function

'== 예수병원 추가루틴 =======================================================================
Public Function SqlQueryForReport_Special(ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String, _
                                  ByVal TestDiv As String, Optional ByVal pReprintFg As Boolean = False, _
                                  Optional ByVal pBatchFg As Boolean = False, _
                                  Optional ByVal pWardMenu As Boolean = False) As String

    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim Result4 As String   '해부병리
    Dim Result5 As String   '혈액은행
    Dim Result6 As String   'POC
    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim tmpCond As String
    Dim tmpNotice As String
    
'    If dbconn.Whatsthis = dbconn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then 'sybase
        MsgBox "사이베이스용으로 소스 수정 필요", vbCritical
'        tmpNotice = FUNC_CONVERT("varchar(255)", "e.text1") & " notice, "
    Else
        tmpNotice = " e.text1 notice, "
    End If
    
    tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    If pWardMenu Then
        tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null "
    Else
        If pBatchFg Then    '일괄재발행(결과출력된것만...)
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                      "  AND  a.rptfg <> ' '  AND  a.rptfg is not null  AND " & DBW("a.rptdt = ", ToDate)
        ElseIf pReprintFg Then  '재발행
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null "
        Else    '신규(결과출력안된것만...)
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                      "  AND (a.rptfg = ''    or   a.rptfg is null) "
        End If
        
    End If
    
    Dim strQuery As String

'    SELECT Case
'        Case "05"
'            If mvarRiPrint = "" Then
'                strQuery = " AND a.workarea not in('OT','OR','RI')"
'            Else
'                strQuery = " AND a.workarea  in('OR','RI')"
'            End If
'    End SELECT
    tmpOrderby = " ORDER BY KeyDate, Seq, workarea, accdt, accseq, RptSeq, ordcd "

    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        '5' stscd, a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                  '"AND      e.empid =* a.lastvfyid   AND   f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    
    If P_NoResultReport <> "" Then
        result1 = result1 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result1 = result1 & strQuery
    End If
    
    'result1 = result1 & " AND " & DBW("a.testcd<>", P_NoResultReport)
    '기타검사
    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea")
                   '"AND      f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result2 = result2 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result2 = result2 & strQuery
    End If
    
    '미생물결과
    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, a. stscd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid,  " & tmpNotice & " " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result3 = result3 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result3 = result3 & strQuery
    End If
    '미생물결과
    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid,   " & tmpNotice & " " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & T_LAB034 & " e, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", PtId) & _
              " AND    x.examdt >= '" & FromDate & "' AND  x.examdt <= '" & ToDate & "' " & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")
                   '"AND      e.empid =* a.lastvfyid   AND    f.empid =* a.vfyid   AND    g.empl_no =* k.orddoct   AND    h.empid =* c.colid   AND    i.empid =* c.rcvid "
    If P_NoResultReport <> "" Then
        result3 = result3 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    If strQuery <> "" Then
        result3 = result3 & strQuery
    End If
    
    '혈당검사 (성모자애 병원에서만 사용하구 있음)
    Result6 = " SELECT '' as WorkArea, '' as AccDt, 0 as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        '' hldiv, '' dpdiv, c.field2 spccd, '' rstdiv, '' statfg, '' lastrst, null lastvfyid, a.vfydt, " & _
              "        a.vfytm, a.vfyid, '9' as testdiv, '' footnotefg, '' rmkcd, '5'  stscd, " & _
              "        '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        '' dcfg, '' mesg, a.ptid, a.vfydt orddt, 99 ordno, 1 ordseq, a.testcd ordcd, 'P' orddiv, " & _
              "        a.wardid, '' hosilid, '' roomid, " & _
                       FUNC_SUBSTR & "(b.testnm,1,16) " & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        99 Seq, a.deptcd, '' sex, 0 ageday, a.majdoct orddoct, a.vfyid colid, a.vfyid rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(a.vfydt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
              "        0 unitqty, '' reqdt, '' reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & _
                       FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2) as KeyDate " & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " c, " & T_LAB195 & " a " & _
              " WHERE  " & DBW("a.ptid = ", PtId) & _
              " AND    a.vfydt >= '" & FromDate & "' AND  a.vfydt <= '" & ToDate & "' " & _
              " AND    a.rstcd  is not null " & _
              " AND    a.rstcd  <> ' ' " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.vfydt ) " & _
              " AND    " & DBW("c.cdindex = ", LC3_POCTestCd) & _
              " AND    c.cdval1  =  a.testcd " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* c.field2")
      
    Select Case TestDiv
        Case "0":  SqlQueryForReport_Special = result1 & " UNION ALL " & Result6     '일반결과만..
        Case "1":  SqlQueryForReport_Special = result2    '특수검사결과만..
        Case "2":  SqlQueryForReport_Special = result3    '미생물결과만..
    End Select
    
    SqlQueryForReport_Special = SqlQueryForReport_Special & tmpOrderby
                           
End Function
'============================================================================================

Public Function SqlQueryForMicReport(ByVal sWorkArea As String, ByVal sAccDt As String, ByVal sAccSeq As String) As String

    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim tmpCond As String
    
    
    tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    
    tmpOrderby = " ORDER BY KeyDate, Seq, workarea, accdt, accseq, RptSeq, ordcd "

    SqlQueryForMicReport = _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("a.workarea = ", sWorkArea) & _
              " AND    " & DBW("a.accdt    = ", sAccDt) & _
              " AND    " & DBW("a.accseq   = ", sAccSeq) & _
              " AND    x.ptid = a.ptid  AND  x.orddt = a.orddt  AND  x.ordno = a.ordno  AND  x.ordseq = a.ordseq  " & _
              " AND    a.senfg = 'Y'    AND  a.vfydt <> ' '     AND  a.vfydt is not null " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea")
      
    SqlQueryForMicReport = SqlQueryForMicReport & tmpOrderby
                           
End Function



Public Function SqlResultsForCmt(ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String) As String

    SqlResultsForCmt = " SELECT  a.workarea, a.accdt, a.accseq, a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, '' as mnmcd, " & _
                       "         a.orddt, a.vfydt, a.rstdiv, a.detailfg, j.field3 as SpcNm, b.testnm as TestLongNm, " & _
                       "         " & FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
                       "         d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, '1' as orddiv " & _
                       " FROM    " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB032 & " j " & _
                       " WHERE   " & DBW("a.ptid", PtId, 2) & _
                       " AND     " & DBW("a.orddt >= ", FromDate) & _
                       " AND     " & DBW("a.orddt <= ", ToDate) & _
                       " AND    (a.hldiv <> ' '  or a.dpdiv <> ' ') " & _
                       " AND     a.vfydt <> ' '" & _
                       " AND     b.testcd = a.testcd  " & _
                       " AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                       "                      WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
                       " AND     " & DBJ(DBW("d.cdindex", LC2_ItemResult, 2)) & _
                       " AND     " & DBJ("d.cdval1 =* a.testcd") & "  AND  " & DBJ("d.cdval2 =* a.rstcd") & _
                       " AND     " & DBJ(DBW("j.cdindex", LC3_Specimen, 2)) & _
                       " AND     " & DBJ("j.cdval1 =* a.spccd")
                       
'                       " ORDER BY a.workarea, a.vfydt, RptSeq, a.testcd "
                       
    SqlResultsForCmt = SqlResultsForCmt & " UNION ALL " & _
                       " SELECT  distinct a.workarea, a.accdt, a.accseq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv,  e.spccd, c.mnmcd, " & _
                       "         a.orddt, a.vfydt, a.rstdiv, a.detailfg, j.field3 as SpcNm, b.testnm as TestLongNm, " & _
                       "         " & FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
                       "         '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, '2' as orddiv " & _
                       " FROM    " & T_LAB404 & " a, " & T_LAB405 & " c, " & T_LAB201 & " e, " & _
                       "         " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB032 & " j " & _
                       " WHERE   " & DBW("a.ptid", PtId, 2) & _
                       " AND     " & DBW("a.orddt >= ", FromDate) & _
                       " AND     " & DBW("a.orddt <= ", ToDate) & _
                       " AND     a.vfydt <> ' '" & _
                       " AND     a.workarea = c.workarea and a.accdt = c.accdt and a.accseq = c.accseq  and a.testcd = c.testcd " & _
                       " AND     a.workarea = e.workarea and a.accdt = e.accdt and a.accseq = e.accseq " & _
                       " AND     b.testcd = a.testcd  " & _
                       " AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                       "                      WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
                       " AND     " & DBJ(DBW("j.cdindex", LC3_Specimen, 2)) & _
                       " AND     " & DBJ("j.cdval1 =* e.spccd") & _
                       " ORDER BY workarea, vfydt, RptSeq, testcd "
                       
                       '** By M.G.Choi 2006.10.22
                       '" ORDER BY workarea, vfydt, RptSeq, testcd "
      
End Function

Public Function SqlResultsForCmt_NEW(ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String) As String

    SqlResultsForCmt_NEW = " SELECT  a.workarea, a.accdt, a.accseq, a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, " & _
                       "         a.orddt, a.vfydt, a.rstdiv, a.detailfg, j.field3 as SpcNm, b.testnm as TestLongNm, " & _
                       "         " & FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, " & _
                       "         d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq " & _
                       " FROM    " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB032 & " j " & _
                       " WHERE   " & DBW("a.ptid", PtId, 2) & _
                       " AND     " & DBW("a.orddt >= ", FromDate) & _
                       " AND     " & DBW("a.orddt <= ", ToDate) & _
                       " AND     a.vfydt <> ' '" & _
                       " AND     b.testcd = a.testcd  " & _
                       " AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                       "                      WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
                       " AND     " & DBJ(DBW("d.cdindex", LC2_ItemResult, 2)) & _
                       " AND     " & DBJ("d.cdval1 =* a.testcd") & "  AND  " & DBJ("d.cdval2 =* a.rstcd") & _
                       " AND     " & DBJ(DBW("j.cdindex", LC3_Specimen, 2)) & _
                       " AND     " & DBJ("j.cdval1 =* a.spccd") & _
                       " ORDER BY a.workarea, a.vfydt, RptSeq, a.testcd "
      
End Function

'%  47. LD, CPK 등 관련검사의 결과 조회
'%       - Calling FROM [ frm202AccDataEntry, frm204WSDataEntry ] :
Public Function SqlGetRelTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As String) As String
    Dim sVfyDt As String
    
    '## 5.0.13: 이상대(2004-12-30)
    '   - 검사항목별 결과코드에 등록된 결과값을 조회
    '   - 특수검사는 조회하는 의미가 없는거 같어 제외
    '## 5.0.14: 이상대(2004-12-31)
    '   - 정확한 인덱스가 타지않어 RULE Base로 변경!
    sVfyDt = Format(DateAdd("d", -60, GetSystemDate), "YYYYMMDD")
    
    '## 일반검사
    SqlGetRelTest = " SELECT /*+ RULE */ a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
                    "        e.cdval1 as testcd,'0' as testdiv, trim(b.abbrnm5) abbrnm5, a.rstcd, f.field1 as rstval, a.rstunit, c.empnm as VfyNm, a.hldiv, a.dpdiv, b.rptseq, " & _
                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, a.spccd, k.field3 as spcnm " & _
                    " FROM   " & T_LAB302 & " d, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " a, " & T_LAB031 & " e, " & T_LAB031 & " f, " & T_LAB032 & " k  " & _
                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
                    " AND    " & DBW("d.accdt", AccDt, 2) & _
                    " AND    " & DBW("d.accseq", accseq, 2) & _
                    " AND    " & DBW("e.cdindex", LC2_RelTest, 2) & _
                    " AND    e.cdval1 = d.testcd " & _
                    " AND    a.ptid = d.ptid " & _
                    " AND    a.testcd = e.cdval2 " & _
                    " AND    a.vfydt" & FUNC_CONCAT & "a.vfytm = (SELECT max(vfydt" & FUNC_CONCAT & "vfytm) FROM " & T_LAB302 & " WHERE ptid = a.ptid AND testcd = a.testcd) " & _
                    " AND    " & DBJ("c.empid =* a.vfyid") & _
                    " AND    b.testcd = a.testcd " & _
                    " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) and k.cdindex = '" & LC3_Specimen & "' and k.cdval1 = a.spccd " & _
                    " AND    " & DBJ("f.cdindex=*'" & LC2_ItemResult & "'") & _
                    " AND    " & DBJ("a.testcd*=f.cdval1") & _
                    " AND    " & DBJ("a.rstcd*=f.cdval2")
                    
                    '** 관련검사항목 조회 안되는 문제로 제외 By M.G.Choi 2007.08.22
                    '" AND    " & DBW("a.vfydt>=", sVfyDt) & _

    '## 특수검사
'    SqlGetRelTest = SqlGetRelTest & " UNION ALL " & _
'                    " SELECT a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
'                    "        e.cdval1 as testcd,'1' as testdiv, b.abbrnm5, '특수검사' as rstcd, '' as rstunit, c.empnm as VfyNm, '' as hldiv, '' as dpdiv, b.rptseq, " & _
'                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
'                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm " & _
'                    " FROM   " & T_LAB351 & " a, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " d, " & T_LAB031 & " e " & _
'                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
'                    " AND    " & DBW("d.accdt", AccDt, 2) & _
'                    " AND    " & DBW("d.accseq", AccSeq, 2) & _
'                    " AND    " & DBW("e.cdindex", LC2_RelTest, 2) & _
'                    " AND    e.cdval1 = d.testcd " & _
'                    " AND    a.ptid = d.ptid " & _
'                    " AND    a.testcd = e.cdval2 " & _
'                    " AND    " & DBW("a.vfydt<>", "<>") & _
'                    "  " & _
'                    " AND    " & DBJ("c.empid =* a.vfyid") & _
'                    " AND    b.testcd = a.testcd " & _
'                    " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " "

    '## 미생물
'    SqlGetRelTest = SqlGetRelTest & " UNION ALL " & _
'                    " SELECT /*+ RULE */ a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
'                    "        e.cdval1 as testcd,'2' as testdiv, trim(b.abbrnm5) abbrnm5, a.rstcd as rstcd, g.field1 as rstval, '' as rstunit, c.empnm as VfyNm, '' as hldiv, '' as dpdiv, b.rptseq, " & _
'                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
'                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, d.spccd, k.field3 as spcnm " & _
'                    " FROM   " & T_LAB201 & " f," & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " d, " & T_LAB031 & " e, " & T_LAB031 & " g, " & T_LAB032 & " k " & _
'                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
'                    " AND    " & DBW("d.accdt", AccDt, 2) & _
'                    " AND    " & DBW("d.accseq", AccSeq, 2) & _
'                    " AND    " & DBW("e.cdindex", LC2_RelTest, 2) & _
'                    " AND    e.cdval1 = d.testcd " & _
'                    " AND    a.ptid = d.ptid " & _
'                    " AND    a.testcd = e.cdval2 " & _
'                    " AND    " & DBW("f.vfydt>", sVfyDt) & _
'                    " AND    " & DBW("f.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
'                    " AND    a.workarea=f.workarea AND a.accdt=f.accdt AND a.accseq=f.accseq " & _
'                    " AND    " & DBJ("c.empid =* a.vfyid") & _
'                    " AND    b.testcd = a.testcd " & _
'                    " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) and k.cdindex = '" & LC3_Specimen & "' and k.cdval1 = d.spccd " & _
'                    " AND    " & DBJ("g.cdindex=*'" & LC2_ItemResult & "'") & _
'                    " AND    " & DBJ("a.testcd*=g.cdval1") & _
'                    " AND    " & DBJ("a.rstcd*=g.cdval2") & _
'                    " ORDER BY VfyDtTm desc,rptseq"

    SqlGetRelTest = SqlGetRelTest & " UNION ALL " & _
                    " SELECT /*+ RULE */ a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
                    "        e.cdval1 as testcd,'2' as testdiv, trim(b.abbrnm5) abbrnm5, a.rstcd as rstcd, g.field1 as rstval, '' as rstunit, c.empnm as VfyNm, '' as hldiv, '' as dpdiv, b.rptseq, " & _
                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, d.spccd, k.field3 as spcnm " & _
                    " FROM   " & T_LAB201 & " f," & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " d, " & T_LAB031 & " e, " & T_LAB031 & " g, " & T_LAB032 & " k " & _
                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
                    " AND    " & DBW("d.accdt", AccDt, 2) & _
                    " AND    " & DBW("d.accseq", accseq, 2) & _
                    " AND    " & DBW("e.cdindex", LC2_RelTest, 2) & _
                    " AND    e.cdval1 = d.testcd " & _
                    " AND    a.ptid = d.ptid " & _
                    " AND    a.testcd = e.cdval2 " & _
                    " AND    " & DBW("f.vfydt>", sVfyDt) & _
                    " AND    f.stscd in ('4','5','6')" & _
                    " AND    a.workarea=f.workarea AND a.accdt=f.accdt AND a.accseq=f.accseq " & _
                    " AND    " & DBJ("c.empid =* a.vfyid") & _
                    " AND    b.testcd = a.testcd " & _
                    " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) and k.cdindex = '" & LC3_Specimen & "' and k.cdval1 = d.spccd " & _
                    " AND    " & DBJ("g.cdindex=*'" & LC2_ItemResult & "'") & _
                    " AND    " & DBJ("a.testcd*=g.cdval1") & _
                    " AND    " & DBJ("a.rstcd*=g.cdval2") & _
                    " ORDER BY VfyDtTm desc,rptseq"
End Function


Public Function SqlGetMicRelTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As String) As String
    Dim sVfyDt As String
    
    '## 5.0.13: 이상대(2004-12-30)
    '   - 검사항목별 결과코드에 등록된 결과값을 조회
    '   - 특수검사는 조회하는 의미가 없는거 같어 제외
    '## 5.0.14: 이상대(2004-12-31)
    '   - 정확한 인덱스가 타지않어 RULE Base로 변경!
    ' 090520 양성현 검체명 추가
    sVfyDt = Format(DateAdd("d", -60, GetSystemDate), "YYYYMMDD")
    
    '## 일반검사
    SqlGetMicRelTest = " SELECT /*+ RULE */ distinct a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
                    "        a.testcd as testcd,'0' as testdiv, b.abbrnm5, a.rstcd, f.field1 as rstval, a.rstunit, c.empnm as VfyNm, a.hldiv, a.dpdiv, b.rptseq, " & _
                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, h.field4 as spcnm " & _
                    " FROM " & T_LAB404 & " d, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " a, " & T_LAB031 & " e, " & T_LAB031 & " f, " & T_LAB201 & " g," & T_LAB032 & " h" & _
                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
                    " AND " & DBW("d.accdt", AccDt, 2) & _
                    " AND " & DBW("d.accseq", accseq, 2) & _
                    " AND " & DBW("e.cdindex", LC2_RelTest, 2) & _
                    " AND e.cdval1 = d.testcd " & _
                    " AND a.ptid = d.ptid " & _
                    " AND a.testcd = e.cdval2 " & _
                    " AND a.vfydt" & FUNC_CONCAT & "a.vfytm = (SELECT max(vfydt" & FUNC_CONCAT & "vfytm) FROM " & T_LAB302 & " WHERE ptid = a.ptid AND testcd = a.testcd) " & _
                    " AND " & DBJ("c.empid =* a.vfyid") & _
                    " AND b.testcd = a.testcd " & _
                    " AND b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " AND " & DBJ("f.cdindex=*'" & LC2_ItemResult & "'") & _
                    " AND " & DBJ("a.testcd*=f.cdval1") & _
                    " AND " & DBJ("a.rstcd*=f.cdval2") & _
                    " and g.workarea = d.workarea and g.accdt = d.accdt and g.accseq = d.accseq AND " & DBW("h.cdindex = ", LC3_Specimen) & " AND h.cdval1 =  g.spccd "
                    
                    '** 관련검사항목 조회 안되는 문제로 제외 By M.G.Choi 2007.08.22
                    '" AND    " & DBW("a.vfydt>=", sVfyDt) & _

    '## 특수검사
'    SqlGetMicRelTest = SqlGetMicRelTest & " UNION ALL " & _
'                    " SELECT a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
'                    "        e.cdval1 as testcd,'1' as testdiv, b.abbrnm5, '특수검사' as rstcd, '' as rstunit, c.empnm as VfyNm, '' as hldiv, '' as dpdiv, b.rptseq, " & _
'                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
'                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm " & _
'                    " FROM   " & T_LAB351 & " a, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " d, " & T_LAB031 & " e " & _
'                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
'                    " AND    " & DBW("d.accdt", AccDt, 2) & _
'                    " AND    " & DBW("d.accseq", AccSeq, 2) & _
'                    " AND    " & DBW("e.cdindex", LC2_RelTest, 2) & _
'                    " AND    e.cdval1 = d.testcd " & _
'                    " AND    a.ptid = d.ptid " & _
'                    " AND    a.testcd = e.cdval2 " & _
'                    " AND    " & DBW("a.vfydt<>", "<>") & _
'                    "  " & _
'                    " AND    " & DBJ("c.empid =* a.vfyid") & _
'                    " AND    b.testcd = a.testcd " & _
'                    " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " "

    '## 미생물
    SqlGetMicRelTest = SqlGetMicRelTest & " UNION ALL " & _
                    " SELECT /*+ RULE */ distinct a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " a.testcd as labno," & _
                    "        e.cdval1 as testcd,'2' as testdiv, b.abbrnm5, a.rstcd as rstcd, g.field1 as rstval, '' as rstunit, c.empnm as VfyNm, '' as hldiv, '' as dpdiv, b.rptseq, " & _
                             FUNC_SUBSTR & "(a.vfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & _
                             FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyDtTm, h.field4 as spcnm " & _
                    " FROM " & T_LAB201 & " f," & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB404 & " d, " & T_LAB031 & " e, " & T_LAB031 & " g, " & T_LAB032 & " h" & _
                    " WHERE  " & DBW("d.workarea", WorkArea, 2) & _
                    " AND " & DBW("d.accdt", AccDt, 2) & _
                    " AND " & DBW("d.accseq", accseq, 2) & _
                    " AND " & DBW("e.cdindex", LC2_RelTest, 2) & _
                    " AND e.cdval1 = d.testcd " & _
                    " AND a.ptid = d.ptid " & _
                    " AND a.testcd = e.cdval2 " & _
                    " AND a.vfydt||a.vfytm = (SELECT max(vfydt||vfytm) FROM s2lab404 WHERE ptid = a.ptid AND testcd = a.testcd) " & _
                    " AND (a.rstcd is not null or a.rstcd <> '') " & _
                    " AND a.workarea=f.workarea AND a.accdt=f.accdt AND a.accseq=f.accseq " & _
                    " AND " & DBJ("c.empid =* a.vfyid") & _
                    " AND b.testcd = a.testcd " & _
                    " AND b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " AND " & DBJ("g.cdindex=*'" & LC2_ItemResult & "'") & _
                    " AND " & DBJ("a.testcd*=g.cdval1") & _
                    " AND " & DBJ("a.rstcd*=g.cdval2") & " AND " & DBW("h.cdindex = ", LC3_Specimen) & " AND h.cdval1 =  f.spccd" & _
                    " ORDER BY VfyDtTm desc,rptseq"
End Function

Public Function SqlGetTAT(ByVal pTestCd As String, ByVal pSpcCd As String) As String

    SqlGetTAT = " SELECT a.tats FROM " & T_LAB004 & " a " & _
                " WHERE  " & DBW("a.testcd", pTestCd, 2) & _
                " AND    " & DBW("a.spccd", pSpcCd, 2) & _
                " AND    a.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                " WHERE  testcd = a.testcd AND spccd= a.spccd)"

End Function


Public Function SqlGetBuildNoForTAT(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

    SqlGetBuildNoForTAT = " SELECT a.field2 as BldNo FROM " & T_LAB032 & " a, " & T_LAB201 & " b " & _
                          " WHERE  " & DBW("b.workarea", pWorkArea, 2) & _
                          " AND    " & DBW("b.accdt   ", pAccDt, 2) & _
                          " AND    " & DBW("b.accseq  ", pAccSeq, 2) & _
                          " AND    " & DBW("a.cdindex ", LC3_Buildings, 2) & _
                          " AND    a.cdval1   = b.buildcd"
End Function

Public Function SqlGetCollectData(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

    SqlGetCollectData = " SELECT * FROM " & T_LAB201 & " a " & _
                        " WHERE  " & DBW("a.workarea", pWorkArea, 2) & _
                        " AND    " & DBW("a.accdt", pAccDt, 2) & _
                        " AND    " & DBW("a.accseq", pAccSeq, 2)
End Function

    
Public Function SqlArchiveVerifiedList(ByVal pDeptcd As String, ByVal pYesterday As String) As String
    
    SqlArchiveVerifiedList = " Delete FROM " & T_LAB202 & _
                             " WHERE  " & DBW("deptcd = ", pDeptcd) & _
                             " AND    " & DBW("vfydt <  ", pYesterday)

End Function
Public Function SqlUpdateVerifiedList(ByVal pDeptcd As String, ByVal pVerifyDt As String, _
                                      ByVal pPtid As String, ByVal pMfyFg As String, _
                                      ByVal pEmpId As String) As String
    SqlUpdateVerifiedList = " Update " & T_LAB202 & _
                            " set    " & DBW("donefg ", "1", 3) & _
                                         DBW("doneid ", pEmpId, 2) & _
                            " WHERE  " & DBW("deptcd ", pDeptcd, 2) & _
                            " AND    " & DBW("vfydt  ", pVerifyDt, 2) & _
                            " AND    " & DBW("ptid   ", pPtid, 2) & _
                            " AND    " & DBW("mfyfg  ", pMfyFg, 2)

End Function

Public Function SqlGetVerifiedList(ByVal pDeptcd As String, ByVal pDoneFg As String) As String
    

    SqlGetVerifiedList = " SELECT a.deptcd, a.ptid, a.vfydt, " & _
                         "        " & FUNC_SUBSTR & "(a.vfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.vfytm,3,2) as VfyTm, " & _
                         "        a.mfyfg, a.donefg, b." & F_PTNM & " as PtNm, '" & lis_orddiv & "' as orddiv " & _
                         " FROM   " & T_LAB202 & " a, " & T_HIS001 & " b " & _
                         " WHERE  " & DBW("a.vfydt  >= ", Format(DateAdd("d", P_VerifyListInterval, Now), CS_DateDbFormat)) & _
                         " AND    " & DBW("a.deptcd  = ", pDeptcd) & _
                         " AND   (" & DBW("a.donefg <> ", pDoneFg) & "  or  a.donefg is null ) " & _
                         " AND    b." & F_PTID & " = a.ptid "
    SqlGetVerifiedList = SqlGetVerifiedList & " ORDER BY vfydt, VfyTm"
    
End Function

Public Function SqlPtListByDoctor(ByVal pDoctId As String) As String

    
    SqlPtListByDoctor = " SELECT   a." & F_INPTID & " ptid, a." & F_PTWARDID & " wardid, a." & F_PTROOMID & " roomid, " & _
                        "          b." & F_PTNM & " as PtNm, '" & lis_orddiv & "' as orddiv " & _
                        " FROM   " & T_HIS001 & " b, " & T_HIS002 & " a " & _
                        " WHERE    a." & F_BEDOUTDT & " is null " & _
                        " AND    " & DBW("a." & F_MAJDOCT & "  = ", pDoctId) & _
                        " AND      b." & F_PTID & " = a." & F_INPTID & _
                        " ORDER BY a." & F_INPTID
End Function

Public Function SqlGetStartDate(ByVal pDeptcd As String) As String

    SqlGetStartDate = " SELECT field1 FROM " & T_LAB032 & _
                      " WHERE " & DBW("cdindex", LC3_StartDate, 2) & _
                      " AND   " & DBW("cdval1", pDeptcd, 2)

End Function

Public Function SqlGetWorkAreaList() As String
    
    SqlGetWorkAreaList = " SELECT cdval1, field1 " & _
                         " FROM  " & T_LAB032 & _
                         " WHERE " & DBW("cdindex", lc3_workarea, 2)
              
End Function

Public Function SqlGetItemList(ByVal pWorkArea As String, ByVal pSpcCd As String) As String
    '수정. detailfg ='D'항목이 안 나옴
    SqlGetItemList = " SELECT   a.testcd, a.testnm, a.rptseq, " & _
                     "          a.testdiv, a.workarea, a.panelfg, c.field3 as SpcNm " & _
                     " FROM   " & T_LAB001 & " a, " & T_LAB004 & " b, " & T_LAB032 & " c " & _
                     " WHERE  " & DBW("a.workarea", pWorkArea, 2) & _
                     " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
                     " AND    b.testcd = a.testcd " & _
                     " AND    " & DBW("b.spccd", pSpcCd, 2) & _
                     " AND    " & DBW("c.cdindex", LC3_Specimen, 2) & _
                     " AND    " & DBW("c.cdval1", pSpcCd, 2) & _
                     " ORDER BY a.rptseq"
'    SqlGetItemList = " SELECT   a.testcd, a.testnm, a.rptseq, " & _
'                     "          a.testdiv, a.workarea, a.panelfg, c.field3 as SpcNm " & _
'                     " FROM   " & T_LAB001 & " a, " & T_LAB004 & " b, " & T_LAB032 & " c " & _
'                     " WHERE  " & DBW("a.workarea", pWorkArea, 2) & _
'                     " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
'                     " AND   (a.panelfg = ''  or  a.panelfg is null) " & _
'                     " AND    b.testcd = a.testcd " & _
'                     " AND    " & DBW("b.spccd", pSpcCd, 2) & _
'                     " AND    " & DBW("c.cdindex", LC3_Specimen, 2) & _
'                     " AND    " & DBW("c.cdval1", pSpcCd, 2) & _
'                     " ORDER BY a.rptseq"
 
End Function

Public Function SqlGetOldResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String, ByVal pTestCd As String) As String
    
    '## 5.0.13: 이상대(2004-12-30)
    '   - 검사항목별 결과코드에 등록된 코드를 결과명도 조회하도록 쿼리수정
    SqlGetOldResult = " SELECT  distinct a.testcd, a.rstcd, d.field1 as rstval, a.vfydt," & _
                      "         a.vfytm, a.vfyid, b.empnm, c.abbrnm5, a.seq mfyseq " & _
                      " FROM  " & T_LAB308 & " a, " & T_LAB015 & " b, " & T_LAB001 & " c, " & T_LAB031 & " d" & _
                      " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                      " AND   " & DBW("a.accdt    = ", pAccDt) & _
                      " AND   " & DBW("a.accseq   = ", pAccSeq) & _
                      " AND   " & DBW("a.testcd   = ", pTestCd) & _
                      " AND   b.empid = a.vfyid  AND  c.testcd = a.testcd " & _
                      " AND   " & DBJ("d.cdindex=*'" & LC2_ItemResult & "'") & _
                      " AND   " & DBJ("a.testcd*=d.cdval1") & _
                      " AND   " & DBJ("a.rstcd*=d.cdval2")
                      
    SqlGetOldResult = SqlGetOldResult & " UNION ALL " & _
                      " SELECT  distinct a.testcd, a.rstcd, d.field1 as rstval, a.vfydt," & _
                      "         a.vfytm, a.vfyid, b.empnm, c.abbrnm5, a.mfyseq " & _
                      " FROM  " & T_LAB407 & " a, " & T_LAB015 & " b, " & T_LAB001 & " c, " & T_LAB031 & " d " & _
                      " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                      " AND   " & DBW("a.accdt    = ", pAccDt) & _
                      " AND   " & DBW("a.accseq   = ", pAccSeq) & _
                      " AND   " & DBW("a.testcd   = ", pTestCd) & _
                      " AND   b.empid = a.vfyid  AND  c.testcd = a.testcd " & _
                      " AND   " & DBJ("d.cdindex=*'" & LC2_ItemResult & "'") & _
                      " AND   " & DBJ("a.testcd*=d.cdval1") & _
                      " AND   " & DBJ("a.rstcd*=d.cdval2") & _
                      " ORDER BY mfyseq DESC "
End Function

'----------------------------------------------------------------------------------
'1. 미생물 Work List 작성 관련 Sql
'----------------------------------------------------------------------------------

Public Function SqlGetCumItem(ByVal pCumCd As String, Optional ByVal iCase As Integer = 0) As String


    If iCase = 0 Then
        SqlGetCumItem = " SELECT a.cdval2 as TestCd, a.field2 as DeptCd, a.field3 as PWD, " & FUNC_CONVERT("num", "a.field4") & " as RptSeq, a.field5, b.testnm, b.panelfg, b.testdiv " & _
                        " FROM  " & T_LAB031 & " a, " & T_LAB001 & " b " & _
                        " WHERE " & DBW("a.cdindex", LC2_CumItem, 2) & _
                        " AND   " & DBW("a.cdval1", pCumCd, 2) & _
                        " AND   b.testcd = a.cdval2 " & _
                        " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                          " WHERE testcd = b.testcd) " & _
                        " ORDER BY RptSeq"
    Else
        SqlGetCumItem = " SELECT a.cdval2 as TestCd, " & FUNC_CONVERT("num", "a.field4") & " as Seq, 1 as Seq1, " & _
                                "a.field5 as SpcCd, b.testnm, b.panelfg, b.testdiv, b.workarea, c.field3 as SpcNm " & _
                        " FROM  " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB032 & " c " & _
                        " WHERE " & DBW("a.cdindex", LC2_CumItem, 2) & _
                        " AND   " & DBW("a.cdval1", pCumCd, 2) & _
                        " AND   b.testcd = a.cdval2 " & _
                        " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                           " WHERE testcd = b.testcd) " & _
                        " AND  (" & DBW("b.panelfg != ", PN_Group) & " or b.panelfg is null) " & _
                        " AND   " & DBW("c.cdindex", LC3_Specimen, 2) & _
                        " AND   c.cdval1 = a.field5 "
        SqlGetCumItem = SqlGetCumItem & " UNION ALL " & _
                        " SELECT c.field1 as TestCd, " & FUNC_CONVERT("num", "a.field4") & " as Seq, d.rptseq as Seq1, " & _
                                "a.field5 as SpcCd, d.testnm, d.panelfg, d.testdiv, d.workarea, e.field3 as SpcNm " & _
                        " FROM  " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB031 & " c, " & _
                                    T_LAB001 & " d, " & T_LAB032 & " e " & _
                        " WHERE " & DBW("a.cdindex", LC2_CumItem, 2) & _
                        " AND   " & DBW("a.cdval1", pCumCd, 2) & _
                        " AND   b.testcd = a.cdval2 " & _
                        " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                           " WHERE testcd = b.testcd) " & _
                        " AND   " & DBW("b.panelfg", PN_Group, 2) & _
                        " AND   " & DBW("c.cdindex", LC2_Panel, 2) & _
                        " AND   c.cdval1 = b.testcd " & _
                        " AND   d.testcd = c.field1 " & _
                        " AND   d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                           " WHERE testcd = d.testcd) " & _
                        " AND   " & DBW("e.cdindex", LC3_Specimen, 2) & _
                        " AND   e.cdval1 = a.field5 " & _
                        " AND  (e.field1=''  or  e.field1 is null)"
        SqlGetCumItem = SqlGetCumItem & " UNION ALL " & _
                        " SELECT c.field1 as TestCd, " & FUNC_CONVERT("num", "a.field4") & " as Seq, d.rptseq as Seq1, " & _
                                "f.field1 as SpcCd, d.testnm, d.panelfg, d.testdiv, d.workarea, g.field3 as SpcNm " & _
                        " FROM  " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB031 & " c, " & _
                                    T_LAB001 & " d, " & T_LAB032 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g " & _
                        " WHERE " & DBW("a.cdindex", LC2_CumItem, 2) & _
                        " AND   " & DBW("a.cdval1", pCumCd, 2) & _
                        " AND   b.testcd = a.cdval2 " & _
                        " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                          " WHERE testcd = b.testcd) " & _
                        " AND   " & DBW("b.panelfg", PN_Group, 2) & _
                        " AND   " & DBW("c.cdindex", LC2_Panel, 2) & _
                        " AND   c.cdval1 = b.testcd " & _
                        " AND   d.testcd = c.field1 " & _
                        " AND   d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                           " WHERE testcd = d.testcd) " & _
                        " AND   " & DBW("e.cdindex", LC3_Specimen, 2) & _
                        " AND   e.cdval1 = a.field5 " & _
                        " AND   " & DBW("e.field1", "Y", 2) & _
                        " AND   " & DBW("f.cdindex", LC2_MultiSpc, 2) & _
                        " AND   f.cdval1 = a.field5 " & _
                        " AND   f.cdval2 = c.cdval2 " & _
                        " AND   " & DBW("g.cdindex", LC3_Specimen, 2) & _
                        " AND   g.cdval1 = f.field1 " & _
                        " ORDER BY Seq, Seq1, TestCd"
                        
    End If
    
End Function

Public Function SqlGetCumDetail(ByVal pTestCd As String) As String

    SqlGetCumDetail = " SELECT a.cdval2 as TestCd, b.testnm, b.workarea, b.panelfg, b.testdiv, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq " & _
                      " FROM  " & T_LAB031 & " a, " & T_LAB001 & " b " & _
                      " WHERE " & DBW("a.cdindex", LC2_Detail, 2) & _
                      " AND   " & DBW("a.cdval1", pTestCd, 2) & _
                      " AND   b.testcd = a.cdval2 " & _
                      " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                         " WHERE testcd = b.testcd) " & _
                      " ORDER BY RptSeq, TestCd"
End Function

Public Function SqlCumResult(ByVal pPtid As String, ByVal pFromDt As String, ByVal pToDt As String, _
                             ByVal pTestCd As String, ByVal pPanelFg As String, ByVal pTestDiv As String, _
                             ByVal pSpcCd As String, ByVal pWorkArea As String)
    Select Case pPanelFg
    Case PN_Detail:
        SqlCumResult = ""
    Case "":
        '## 5.0.25: 이상대(2005-05-30)
        '   - 속도를 위해 Rule Base로 변경
        '   - Lab303(Text결과내역)과 조인되는 부분 제거
        SqlCumResult = " SELECT /*+ RULE S2LAB201 (S2LAB201_IDX5) */ b.coldt, " & FUNC_SUBSTR & "(b.coltm, 1, 4) as ColTm, " & _
                                "b.rcvdt, " & FUNC_SUBSTR & "(b.rcvtm, 1, 4) as RcvTm, " & _
                                "a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.statfg, " & _
                                "a.workarea, a.accdt, a.accseq, b.footnotefg, b.rmkcd, " & _
                                "b.deptcd, b.wardid, b.hosilid, " & _
                                "a.rstdiv, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, " & _
                                "d.field1 as RstCdNm " & _
                       " FROM  " & T_LAB302 & " a, " & T_LAB201 & " b, " & T_LAB031 & " d "
        SqlCumResult = SqlCumResult & _
                       " WHERE " & DBW("b.ptid", pPtid, 2) & _
                       " AND   " & DBW("b.coldt >= ", pFromDt) & _
                       " AND   " & DBW("b.workarea", pWorkArea, 2) & _
                       " AND   " & DBW("b.rcvdt >= ", pFromDt) & _
                       " AND   " & DBW("b.stscd >= ", enStsCd.StsCd_LIS_Accession) & _
                       " AND   a.workarea = b.workarea " & _
                       " AND   a.accdt = b.accdt " & _
                       " AND   a.accseq = b.accseq " & _
                       " AND   " & DBW("a.testcd", pTestCd, 2) & _
                       " AND   " & DBW("a.spccd", pSpcCd, 2) & _
                       " AND   a.vfydt <> ' '" & _
                       " AND   " & DBJ(DBW("d.cdindex", LC2_ItemResult, 2)) & _
                       " AND   " & DBJ("d.cdval1   =* a.testcd") & _
                       " AND   " & DBJ("d.cdval2   =* a.rstcd")
    End Select
'''    If pTestCd = "C3720" Then
'''       MsgBox "aa"
'''    End If
    Debug.Print SqlCumResult
End Function

Public Function SqlCumOldResult(ByVal pPtid As String, ByVal pFromDt As String, ByVal pToDt As String, _
                             ByVal pTestCd As String, ByVal pPanelFg As String, ByVal pTestDiv As String, _
                             ByVal pSpcCd As String, ByVal pWorkArea As String)
    Dim strTblNm As String
    Dim strRcvDt As String
    
    strRcvDt = Format(DateAdd("d", -2, Format(pToDt, CS_DateMask)), CS_DateDbFormat)
    strTblNm = Choose(Val(pTestDiv) + 1, T_LAB302, T_LAB351, T_LAB405)
    
    Select Case pPanelFg
    Case PN_Detail:
        SqlCumOldResult = " SELECT "
    Case "":
        SqlCumOldResult = " SELECT  a.vfydt as ColDt, " & FUNC_SUBSTR & "(a.vfytm, 1, 4) as ColTm, " & _
                                "a.testcd, a.rstcd, a.rstunit, a.hidiv as hldiv, '' as dpdiv, a.spccd, '' as statfg, " & _
                                "a.workarea, '' as accdt, 0 as accseq, a.mesg as footnote, '' as rmkcd, " & _
                                "a.deptcd, '' as wardid, '' as hosilid, " & _
                                "'' as rstdiv, '' as txtfg, '' as rsttype, '' as detailfg, '' as eqpcd, " & _
                                "'' as TextResult, '' as RstCdNm " & _
                          " FROM  " & T_LAB999 & " a "
        SqlCumOldResult = SqlCumOldResult & _
                          " WHERE " & DBW("a.ptid", pPtid, 2) & _
                          " AND   " & DBW("a.testcd", pTestCd, 2) & _
                          " AND   " & DBW("a.spccd", pSpcCd, 2) & _
                          " AND   " & DBW("a.vfydt >= ", pFromDt)
        SqlCumOldResult = SqlCumOldResult & " UNION ALL " & _
                          " SELECT  a.vfydt as ColDt, " & FUNC_SUBSTR & "(a.vfytm, 1, 4) as ColTm, " & _
                                  "rtrim(a.testcd) " & FUNC_CONCAT & " rtrim(a.testseq) as testcd, a.rstcd, a.rstunit, a.hidiv as hldiv, '' as dpdiv, a.spccd, '' as statfg, " & _
                                  "a.workarea, '' as accdt, 0 as accseq, a.mesg as footnote, '' as rmkcd, " & _
                                  "a.deptcd, '' as wardid, '' as hosilid, " & _
                                  "'' as rstdiv, '' as txtfg, '' as rsttype, '' as detailfg, '' as eqpcd, " & _
                                  "'' as TextResult, '' as RstCdNm " & _
                          " FROM  " & T_LAB999 & " a "
        SqlCumOldResult = SqlCumOldResult & _
                          " WHERE " & DBW("a.ptid", pPtid, 2) & _
                          " AND   rtrim(a.testcd) " & FUNC_CONCAT & " rtrim(a.testseq) = " & DBS(pTestCd) & _
                          " AND   " & DBW("a.spccd", pSpcCd, 2) & _
                          " AND   " & DBW("a.vfydt >= ", pFromDt)
    End Select
    
End Function


Public Function SqlGetPOCResult(ByVal pPtid As String, ByVal pVfyDt As String) As String

    SqlGetPOCResult = " SELECT a.vfytm, a.vfyid, a.testcd, a.rstcd, a.rstunit, c.testnm, c.abbrnm10, b.empnm vfynm " & _
                      " FROM  " & T_LAB195 & " a, " & T_LAB015 & " b, " & T_LAB001 & " c " & _
                      " WHERE " & DBW("a.ptid", pPtid, 2) & _
                      " AND   " & DBW("a.vfydt", pVfyDt, 2) & _
                      " AND   a.rstcd  is not null " & _
                      " AND   a.rstcd  <> ' ' " & _
                      " AND   b.empid   = a.vfyid " & _
                      " AND   c.testcd  = a.testcd " & _
                      " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                        " WHERE testcd = c.testcd) " & _
                      " ORDER BY a.vfytm"
End Function


Public Function GetAssignCnt(ByVal qWorkArea As String, ByVal qAccdt As String, qaccseq As String) As Long
    
    Dim sSQL As String
    Dim aRs   As Recordset
    
    sSQL = " SELECT assigncnt,assigncancelcnt,deliverycnt FROM " & T_BBS203 & _
           " WHERE " & _
                     DBW("workarea=", qWorkArea) & _
           " AND " & DBW("accdt=", qAccdt) & _
           " AND " & DBW("accseq=", qaccseq)
    
    Set aRs = New Recordset
    aRs.Open sSQL, DBConn
    
    If Not aRs.EOF Then
        GetAssignCnt = Val(aRs.Fields("assigncnt").Value & "")
    Else
        GetAssignCnt = 0
    End If
    Set aRs = Nothing

End Function

Public Function GetDeliveryCnt(ByVal qWorkArea As String, ByVal qAccdt As String, qaccseq As String) As String
    Dim sSQL As String
    Dim rs   As Recordset
    
    sSQL = " SELECT assigncnt,assigncancelcnt,deliverycnt,retcnt,expcnt FROM " & T_BBS203 & _
           " WHERE " & _
                     DBW("workarea=", qWorkArea) & _
           " AND " & DBW("accdt=", qAccdt) & _
           " AND " & DBW("accseq=", qaccseq)
        
    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        GetDeliveryCnt = rs.Fields("assigncnt").Value & "" & COL_DIV & _
                         rs.Fields("assigncancelcnt").Value & "" & COL_DIV & _
                         rs.Fields("deliverycnt").Value & "" & COL_DIV & _
                         rs.Fields("retcnt").Value & "" & COL_DIV & _
                         rs.Fields("expcnt").Value & ""
    Else
        GetDeliveryCnt = ""
    End If
    Set rs = Nothing
End Function

Public Function GetDeliveryCnt_New(ByVal qWorkArea As String, ByVal qAccdt As String, qaccseq As String, qRstSeq As String) As String
    Dim sSQL As String
    Dim rs   As Recordset
    
    sSQL = " SELECT a.assigncnt,a.assigncancelcnt,a.deliverycnt,a.retcnt,a.expcnt,b.deliverydt,b.deliverytm FROM " & T_BBS203 & " a, " & T_BBS402 & " b" & _
           " WHERE " & _
                     DBW("a.workarea=", qWorkArea) & _
           " AND " & DBW("a.accdt=", qAccdt) & _
           " AND " & DBW("a.accseq=", qaccseq) & _
           " AND b.rstseq = '" & qRstSeq & "' " & _
           " AND a.workarea = b.workarea " & _
           " AND a.accdt = b.accdt " & _
           " AND a.accseq = b.AccSeq "

    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        GetDeliveryCnt_New = rs.Fields("assigncnt").Value & "" & COL_DIV & _
                         rs.Fields("assigncancelcnt").Value & "" & COL_DIV & _
                         rs.Fields("deliverycnt").Value & "" & COL_DIV & _
                         rs.Fields("retcnt").Value & "" & COL_DIV & _
                         rs.Fields("expcnt").Value & "" & COL_DIV & _
                         rs.Fields("deliverydt").Value & "" & COL_DIV & _
                         rs.Fields("deliverytm").Value & ""
    Else
        GetDeliveryCnt_New = ""
    End If
    Set rs = Nothing
End Function

Public Function GetDeptInfo(ByVal PtId As String) As String
    Dim sSQL As String
    Dim rs   As Recordset
    Dim strOrdDt As String
    Dim strOrdNO As String
    
    
    GetDeptInfo = ""
    sSQL = " SELECT max(orddt) as orddt1" & _
           " FROM " & T_LAB101 & _
           " WHERE " & DBW("ptid=", PtId)

    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        strOrdDt = rs.Fields("orddt1").Value & ""
        sSQL = " SELECT max(ordno) as ordno1" & _
               " FROM " & T_LAB101 & _
               " WHERE " & DBW("ptid=", PtId) & _
               " AND " & DBW("orddt=", strOrdDt)
        Set rs = Nothing
        Set rs = New Recordset
        
        rs.Open sSQL, DBConn
        
        strOrdNO = rs.Fields("ordno1").Value & ""
        
        sSQL = " SELECT b." & F_DEPTNM & " as deptnm, c." & F_DOCTNM & " as doctnm" & _
               " FROM " & T_LAB101 & " a," & T_HIS003 & " b," & T_HIS005 & " c" & _
               " WHERE " & DBW("a.ptid=", PtId) & _
               " AND " & DBW("a.orddt=", strOrdDt) & _
               " AND " & DBW("a.ordno=", strOrdNO) & _
               " AND  a.deptcd=b." & F_DEPTCD & _
               " AND  a.orddoct=c." & F_DOCTID
        Set rs = Nothing
        Set rs = New Recordset
        rs.Open sSQL, DBConn
        
        If Not rs.EOF Then
            GetDeptInfo = rs.Fields("deptnm").Value & "" & COL_DIV & rs.Fields("doctnm").Value & ""
        End If
    End If
    Set rs = Nothing
End Function

'%  추가 MGChoi 2002.08.27
'%  병동Id, 퇴원일자를 기준으로 환자리스트를 가져온다.
'%       - His002 : 재원환자마스터
'%  조건식 : 1. 전체병동 여부
'            2. 금일검사결과 여부
Public Function SqlWard_PtList(ByVal pWardId As String, ByVal pOutDate As String, _
                               ByVal pAllFlag As Boolean, Optional RstDate As String = "") As String
    
    If pAllFlag = False Then
        If RstDate <> "" Then
            SqlWard_PtList = " SELECT distinct a." & F_INPTID & " as hospno " _
                           & "   FROM " & T_HIS002 & " a" & "," & T_LAB102 & " b"
                   
            SqlWard_PtList = SqlWard_PtList & "  WHERE " & DBW("a." & F_PTWARDID, pWardId, 2) _
                                            & "    AND (" & FUNC_CONVERT("char", "a." & F_BEDOUTDT, "yyyymmdd") & " is null)" _
                                            & "    AND " & DBW("b.examdt", RstDate, 2) _
                                            & "    AND " & DBW("b.stscd", StsCd_LIS_FinRst, 2) _
                                            & "    AND a." & F_INPTID & " = b.ptid "
        Else
            SqlWard_PtList = " SELECT " & F_INPTID & " as hospno FROM " & T_HIS002
                           
            SqlWard_PtList = SqlWard_PtList & "  WHERE " & DBW(F_PTWARDID, pWardId, 2) _
                                            & "    AND (" & FUNC_CONVERT("char", F_BEDOUTDT, "yyyymmdd") & " is null )"
        End If
    Else
        If RstDate <> "" Then
            SqlWard_PtList = " SELECT distinct a." & F_INPTID & " as hospno " _
                           & "   FROM " & T_HIS002 & " a" & "," & T_LAB102 & " b"
                           
            SqlWard_PtList = SqlWard_PtList & "  WHERE (" & FUNC_CONVERT("char", "a." & F_BEDOUTDT, "yyyymmdd") & " is null )" _
                                            & "    AND " & DBW("b.examdt", RstDate, 2) _
                                            & "    AND " & DBW("b.stscd", StsCd_LIS_FinRst, 2) _
                                            & "    AND a." & F_INPTID & " = b.ptid"
        Else
            SqlWard_PtList = " SELECT " & F_INPTID & " as hospno FROM " & T_HIS002
            
            SqlWard_PtList = SqlWard_PtList & "  WHERE ( " & FUNC_CONVERT("char", F_BEDOUTDT, "yyyymmdd") & " is null )"
        End If
    End If
End Function

'-- 성바오로병원 결과지 출력...FeedBack
Public Function SqlGet_LAB404(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String) As String

    SqlGet_LAB404 = " SELECT rptfg, rptdt, rpttm, rptid " & _
                    "   FROM " & T_LAB404 & _
                    "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                    "    AND " & DBW("accdt", pAccDt, 2) & _
                    "    AND " & DBW("accseq", pAccSeq, 2) & _
                    "    AND " & DBW("testcd", pTestCd, 2)
End Function

'-- 이미지 데이타 내역
Public Function SqlGetImageData(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, Optional pTestCd As String = "", _
                               Optional pSeq As String = "") As String

    SqlGetImageData = " SELECT * " & _
                    "   FROM " & T_LAB310 & _
                    "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                    "    AND " & DBW("accdt", pAccDt, 2) & _
                    "    AND " & DBW("accseq", pAccSeq, 2) & _
                    "    AND prtfg ='1'"
    
    If pTestCd <> "" Then SqlGetImageData = SqlGetImageData & "    AND " & DBW("testcd", pTestCd, 2)
    If pSeq <> "" Then SqlGetImageData = SqlGetImageData & "    AND " & DBW("seq", pSeq, 2)
End Function

Public Function SqlSpecialReport(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String, _
                                 ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String, _
                                 ByVal TestDiv As String, Optional ByVal pReprintFg As Boolean = False, _
                                 Optional ByVal pBatchFg As Boolean = False, _
                                 Optional ByVal pWardMenu As Boolean = False) As String

    Dim result2     As String
    Dim tmpStr      As String
    Dim tmpOrderby  As String
    Dim tmpCond     As String
    Dim tmpNotice   As String
    
    tmpNotice = " e.text1 notice, "
    tmpStr = "  " & FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(x.examdt,7,2) as KeyDate "
    tmpOrderby = " ORDER BY KeyDate, Seq, workarea, accdt, accseq, RptSeq, ordcd "
    
    If pWardMenu Then
        tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null "
    Else
        If pBatchFg Then    '일괄재발행(결과출력된것만...)
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                      "  AND  a.rptfg <> ' '  AND  a.rptfg is not null  AND " & DBW("a.rptdt = ", ToDate)
        ElseIf pReprintFg Then  '재발행
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null "
        Else    '신규(결과출력안된것만...)
            tmpCond = "  AND  a.vfydt <> ' '  AND  a.vfydt is not null " & _
                      "  AND (a.rptfg = ''    or   a.rptfg is null) "
        End If
    End If

    '기타검사
    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd,  a.stscd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, " & FUNC_CONVERT("num", "i.field5") & " Seq, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, '' notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB032 & " i, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("a.workarea=", pWorkArea) & _
              " AND    " & DBW("a.accdt=", pAccDt) & _
              " AND    " & DBW("a.accseq=", pAccSeq) & _
              " AND    " & DBW("a.testcd=", pTestCd) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & tmpCond & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd") & _
              " AND  " & DBJ(DBW("i.cdindex = ", lc3_workarea)) & "   AND " & DBJ("i.cdval1 =* a.workarea")
    
    If P_NoResultReport <> "" Then
        result2 = result2 & " AND a.testcd not in(" & P_NoResultReport & ")"
    End If
    
    SqlSpecialReport = result2 & tmpOrderby
End Function

