VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISSqlAccession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'%  22. 접수번호로 처방Body 검색
'%      - Calling FROM [clsAccession]
Public Function SqlOrderItems(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As Integer, Optional ByVal pMultiFg As String) As String
   Dim strtmp As String
   
   ' ** 예수병원 속도로 인한 수정
   Dim strSQL   As String
   Dim Rs       As New ADODB.Recordset
   Dim strPtID  As String
     
   strSQL = " select DiSTINCT ptid from " & T_LAB102 & " where workarea = " & DBS(pWorkArea) & " and accdt = " & DBS(pAccDt) & " and accseq = " & DBN(pAccSeq)
   
   Set Rs = DbCon.Execute(strSQL, , adCmdText)
    
    If Rs.EOF = False Then
        strPtID = Rs.Fields("ptid").Value & ""
    End If
    
    Rs.Close: Set Rs = Nothing
   '============================
   
'  " FROM " & T_LAB203 & " d, " & T_LAB102 & " a, " & T_LAB201 & " e, " & T_LAB001 & " b, " & T_LAB004 & " c "
'  " FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB203 & " d, " & T_LAB201 & " e "
'
    '복수검체(연속검사)
    SqlOrderItems = " SELECT /*+ INDEX(d, s2lab203_idx1) *//*+ INDEX(a, s2ord102_pk) *//*+ INDEX(e, s2lab201_idx1) */ a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd, a.spccd, a.attrcd, a.statfg, " & _
                    "        b.testdiv, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                    "        e.testdiv as L201TestDiv, e.multifg, e.spccd as ColSpcCd, d.spcseq, " & _
                    "        '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, " & DBV("vfyid", "0") & " as LastVfyId " & _
                    " FROM " & T_LAB203 & " d, " & T_LAB102 & " a, " & T_LAB201 & " e, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                    " WHERE   " & DBW("d.workarea =", pWorkArea) & _
                    " AND     " & DBW("d.accdt   =", pAccDt) & _
                    " AND     " & DBW("d.accseq  =", pAccSeq) & _
                    " AND    e.workarea = d.workarea " & _
                    " AND    e.accdt   =  d.accdt " & _
                    " AND    e.accseq  =  d.accseq " & _
                    " AND    e.multifg <> ' ' " & _
                    " AND    a.ptid    =  d.ptid " & _
                    " AND    a.orddt   =  d.orddt " & _
                    " AND    a.ordno   =  d.ordno " & _
                    " AND    a.ordseq  =  d.ordseq " & _
                    " AND   (a.dcfg = '' or a.dcfg is null) " & _
                    " AND    b.testcd  =  a.ordcd " & _
                    " AND    b.applydt =  (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " AND    c.testcd  =  a.ordcd " & _
                    " AND    c.spccd   =  a.spccd " & _
                    " AND    c.applydt =  (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                   " AND   spccd = c.spccd) "
   'Else
   
   '** 예수병원 추가 (일반검사 302 FULL SCAN)=========
   strtmp = " AND    a.ptid = " & DBS(strPtID)
             
'    strtmp = " AND    d.workarea (+)= a.workarea " & _
'             " AND    d.accdt    (+)= a.accdt " & _
'             " AND    d.accseq   (+)= a.accseq "
    '=================================================
'" FROM " & T_LAB102 & " a, " & T_LAB201 & " e, " & T_LAB302 & " d, " & T_LAB001 & " b, " & T_LAB004 & " c "
'" FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB302 & " d, " & T_LAB201 & " e "
    '일반검사
    SqlOrderItems = SqlOrderItems & " UNION ALL " & _
                    " SELECT  /*+ INDEX(a, s2ord102_idx1) *//*+ INDEX(e, s2lab201_idx1) *//*+ INDEX(d, s2lab302_pk) */ a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd, a.spccd, a.attrcd, a.statfg, " & _
                    "        b.testdiv, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                    "        e.testdiv as L201TestDiv, e.multifg, e.spccd as ColSpcCd, '' as SpcSeq, " & _
                    "        d.rstcd as LastRst, d.vfydt as LastVfyDt, d.vfytm as LastVfyTm, d.vfyid as LastVfyId " & _
                    " FROM " & T_LAB102 & " a, " & T_LAB201 & " e, " & T_LAB302 & " d, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                    " WHERE   " & DBW("a.workarea =", pWorkArea) & _
                    " AND     " & DBW("a.accdt    =", pAccDt) & _
                    " AND     " & DBW("a.accseq   =", pAccSeq) & _
                    " AND   (a.dcfg = '' or a.dcfg is null) " & _
                    " AND    e.workarea = a.workarea " & _
                    " AND    e.accdt    = a.accdt " & _
                    " AND    e.accseq   = a.accseq " & strtmp & _
                    " AND   (e.multifg  = '' or  e.multifg is null ) " & _
                    " AND    b.testcd   = a.ordcd " & _
                    " AND    b.applydt  = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " AND    c.testcd   = a.ordcd  AND    c.spccd    = a.spccd " & _
                    " AND    c.applydt  = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                   " AND spccd = c.spccd) " & _
                    " AND    d.ptid     = a.ptid   AND    d.testcd = a.ordcd   AND  d.spccd = a.spccd " & _
                    " AND    d.vfydt" & FUNC_CONCAT & "d.vfytm = (SELECT /*+ INDEX_DESC(" & T_LAB302 & ", s2lab302_idx2) */ max(vfydt" & FUNC_CONCAT & "vfytm) FROM " & T_LAB302 & " WHERE ptid = a.ptid  " & _
                                                                  " AND  testcd = a.ordcd  AND  spccd = a.spccd  AND vfydt <> ' ' " & _
                                                                  " AND  vfydt <= " & DBV("vfydt", Format(Now, CS_DateDbFormat)) & ") " & _
                    " AND    exists (SELECT /*+ INDEX_DESC(" & T_LAB302 & ", s2lab302_idx2) */ * FROM " & T_LAB302 & " WHERE ptid = a.ptid AND  testcd = a.ordcd  AND spccd = a.spccd " & _
                                                                  " AND   vfydt <> ' ' AND vfydt <= " & DBV("vfydt", Format(Now, CS_DateDbFormat)) & ") "
'" FROM " & T_LAB102 & " a, " & T_LAB201 & " e, " & T_LAB001 & " b, " & T_LAB004 & " c "
'" FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB201 & " e "
    SqlOrderItems = SqlOrderItems & " UNION ALL " & _
                    " SELECT /*+ INDEX(a, s2ord102_idx1) *//*+ INDEX(e, s2lab201_idx1) */ a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd, a.spccd, a.attrcd, a.statfg, " & _
                    "        b.testdiv, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                    "        e.testdiv as L201TestDiv, e.multifg, e.spccd as ColSpcCd, '' as SpcSeq, " & _
                    "        '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, " & DBV("vfyid", "0") & " as LastVfyId " & _
                    " FROM " & T_LAB102 & " a, " & T_LAB201 & " e, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                    " WHERE   " & DBW("a.workarea =", pWorkArea) & _
                    " AND     " & DBW("a.accdt    =", pAccDt) & _
                    " AND     " & DBW("a.accseq   =", pAccSeq) & _
                    " AND   (a.dcfg = '' or a.dcfg is null) " & _
                    " AND    e.workarea = a.workarea " & _
                    " AND    e.accdt    = a.accdt " & _
                    " AND    e.accseq   = a.accseq " & _
                    " AND   (e.multifg  = '' or e.multifg is null)" & _
                    " AND    b.testcd   = a.ordcd " & _
                    " AND    b.applydt  = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                    " AND    c.testcd   = a.ordcd " & _
                    " AND    c.spccd    = a.spccd " & _
                    " AND    c.applydt  = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                   " AND spccd = c.spccd) " & _
                    " AND    not exists   (SELECT /*+ INDEX(" & T_LAB302 & ", s2lab302_idx1) */ * FROM " & T_LAB302 & " WHERE ptid = a.ptid " & _
                                                                        " AND  testcd = a.ordcd AND spccd = a.spccd " & _
                                                                        " AND vfydt <> ' ' " & _
                                                                        " AND vfydt <= " & DBV("vfydt", Format(Now, CS_DateDbFormat)) & ") "
  
    
End Function

'%  22. 접수번호로 처방Body 검색
'%      - Calling FROM [clsAccession]
Public Function SqlOrderItems1(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, Optional ByVal ItemCase As String) As String
   
    Dim sSelectClause(3) As String
    Dim sFromClause(3) As String
    Dim sWhereClause(3) As String
    
    Select Case ItemCase
    Case "Normal":
        sSelectClause(1) = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, a.ordcd as TestCd2, a.spccd as OrdSpc, " & _
                           "        a.attrcd, a.statfg, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, " & _
                           "        c.rstunit, d.spccd, d.multifg, d.testdiv, 'N' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "  e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "  '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
        
        sFromClause(1) = " FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB201 & " d "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
        
        sWhereClause(1) = " WHERE " & DBW("a.workarea =", pWorkArea) & _
                          " AND   " & DBW("a.accdt    =", pAccDt) & _
                          " AND   " & DBW("a.accseq   =", pAccSeq) & _
                          " AND   d.workarea = a.workarea   AND    d.accdt = a.accdt    AND    d.accseq = a.accseq   " & _
                          " AND  (d.multifg  = ''  or d.multifg is null) " & _
                          " AND   b.testcd   = a.ordcd      AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                          " AND   " & DBW("b.panelfg  =", PN_Normal) & _
                          " AND   c.testcd   = a.ordcd   AND    c.spccd = a.spccd " & _
                          " AND   c.applydt  = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                        " AND spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                          " AND   e.ptid  = a.ptid   AND    e.testcd = a.ordcd   AND   e.spccd = a.spccd " & _
                          " AND   e.vfydt = (SELECT max(vfydt) FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = a.ordcd AND spccd = a.spccd AND vfydt <> '' AND vfydt < d.rcvdt) " & _
                          " AND   exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = a.ordcd AND  spccd = a.spccd  AND vfydt <> '' AND vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Detail":
        sSelectClause(1) = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, f.cdval2 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                           "        b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                           "        d.spccd, d.multifg, d.testdiv, 'Y' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & " e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & " '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
                                    
        sFromClause(1) = " FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                    T_LAB201 & " d, " & T_LAB031 & " f, " & T_LAB001 & " g "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
                                    
        sWhereClause(1) = " WHERE  " & DBW("a.workarea =", pWorkArea) & _
                          " AND    " & DBW("a.accdt =", pAccDt) & _
                          " AND    " & DBW("a.accseq = ", pAccSeq) & _
                          " AND    b.testcd = a.ordcd " & _
                          " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                          " AND    " & DBW("b.panelfg =", PN_Detail) & _
                          " AND    d.workarea = a.workarea   AND    d.accdt = a.accdt   AND    d.accseq = a.accseq  " & _
                          " AND   (d.multifg = '' or  d.multifg is null) " & _
                          " AND    " & DBW("f.cdindex =", LC2_Detail) & _
                          " AND    f.cdval1 = a.ordcd " & _
                          " AND    g.testcd = f.cdval2  " & _
                          " AND    g.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = g.testcd) " & _
                          " AND    c.testcd = f.cdval2   AND    c.spccd = a.spccd " & _
                          " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                        " AND spccd    = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                          " AND    e.ptid = a.ptid   AND    e.testcd = f.cdval2  AND e.spccd = d.spccd " & _
                          " AND    e.vfydt = (SELECT max(vfydt) FROM " & T_LAB302 & " WHERE ptid = a.ptid  " & _
                                                                        " AND  testcd = f.cdval2 AND spccd = d.spccd AND vfydt <> '' AND vfydt < d.rcvdt) " & _
                          " AND    exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  " & _
                                                                        " AND  testcd = f.cdval2  AND spccd = d.spccd " & _
                                                                        " AND (vfydt <> '' or vfydt is null)" & _
                                                                        " AND  vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "AND    not exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = a.ordcd AND vfydt <> '' AND vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " UNION ALL " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Group1":   '일반 그룹처방
        sSelectClause(1) = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, f.field1 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                           "        b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                           "        d.spccd, d.multifg, d.testdiv, 'N' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "  e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "  '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
                                    
        sFromClause(1) = " FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                    T_LAB201 & " d, " & T_LAB031 & " f, " & T_LAB001 & " h "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
                                    
        sWhereClause(1) = " WHERE " & DBW("a.workarea =", pWorkArea) & _
                          " AND   " & DBW("a.accdt =", pAccDt) & _
                          " AND   " & DBW("a.accseq =", pAccSeq) & _
                          " AND   d.workarea = a.workarea   AND    d.accdt = a.accdt   AND    d.accseq = a.accseq   " & _
                          " AND  (d.multifg = ''  or  d.multifg is null)  AND   h.testcd = a.ordcd " & _
                          " AND   h.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = h.testcd) " & _
                          " AND   " & DBW("h.panelfg =", PN_Group) & _
                          " AND   " & DBW("f.cdindex =", LC2_Panel) & _
                          " AND   f.cdval1 = h.testcd " & _
                          " AND   b.testcd = f.field1 " & _
                          " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                          " AND   " & DBW("b.panelfg =", PN_Normal) & _
                          " AND   c.testcd = b.testcd   AND    c.spccd = d.spccd " & _
                          " AND   c.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                       " AND   spccd  = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                          " AND    e.ptid = a.ptid   AND    e.testcd = f.field1  AND  e.spccd = d.spccd " & _
                          " AND    e.vfydt = (SELECT max(vfydt) FROM " & T_LAB302 & " WHERE ptid   = a.ptid  " & _
                                                                                    " AND   testcd = f.field1 AND spccd = d.spccd " & _
                                                                                    " AND   vfydt <> ' ' " & _
                                                                                    " AND   vfydt < d.rcvdt) " & _
                          " AND    exists (SELECT * FROM " & T_LAB302 & " WHERE ptid   = a.ptid  " & _
                                                                        " AND   testcd = f.field1 AND spccd = d.spccd " & _
                                                                        " AND   vfydt <> ' ' " & _
                                                                        " AND   vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "AND    not exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = f.field1 AND vfydt <> '' AND vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " UNION ALL " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Group2":   '그룹처방이면서 상세항목
        sSelectClause(1) = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, f.field1 as Testcd1, g.cdval2 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                           "        b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                           "        d.spccd, d.multifg, d.testdiv, 'Y' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & " e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & " '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
                                    
        sFromClause(1) = " FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB031 & " g, " & _
                                    T_LAB201 & " d, " & T_LAB031 & " f, " & T_LAB001 & " h, " & T_LAB001 & " i "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
                                    
        sWhereClause(1) = " WHERE  " & DBW("a.workarea =", pWorkArea) & _
                          " AND    " & DBW("a.accdt    =", pAccDt) & _
                          " AND    " & DBW("a.accseq   =", pAccSeq) & _
                          " AND    d.workarea = a.workarea   AND    d.accdt = a.accdt   AND    d.accseq = a.accseq   " & _
                          " AND   (d.multifg = '' or d.multifg is null)  AND    i.testcd = a.ordcd " & _
                          " AND    i.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = i.testcd) " & _
                          " AND    " & DBW("i.panelfg =", PN_Group) & _
                          " AND    " & DBW("f.cdindex =", LC2_Panel) & "   AND    f.cdval1 = i.testcd " & _
                          " AND    h.testcd = f.field1 " & _
                          " AND    h.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = h.testcd) " & _
                          " AND    " & DBW("h.panelfg =", PN_Detail) & _
                          " AND    " & DBW("g.cdindex =", LC2_Detail) & "  AND    g.cdval1 = h.testcd " & _
                          " AND    b.testcd = g.cdval2 " & _
                          " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                          " AND    c.testcd = g.cdval2   AND    d.spccd = a.spccd " & _
                          " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd AND spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                          " AND    e.ptid = a.ptid   AND    e.testcd = g.cdval2 AND e.spccd = d.spccd " & _
                          " AND    e.vfydt = (SELECT max(vfydt) FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = g.cdval2 AND spccd = d.spccd " & _
                                                                                    " AND vfydt <> ' ' AND vfydt < d.rcvdt) " & _
                          " AND    exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid   AND  testcd = g.cdval2 AND spccd = d.spccd " & _
                                                                        " AND vfydt <> ' ' AND vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "AND    not exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = g.cdval2 AND vfydt <> '' AND vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " UNION ALL " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Group3":    '그룹이면서 연속검사(복수검체)
        sSelectClause(1) = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, f.field1 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                           "        b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                           "        e.spccd, e.multifg, e.testdiv, 'N' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "  g.rstcd as LastRst, g.vfydt as LastVfyDt, g.vfytm as LastVfyTm, g.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "  '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "

        sFromClause(1) = " FROM " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                    T_LAB203 & " d, " & T_LAB201 & " e, " & T_LAB031 & " f "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " g "
        sFromClause(3) = sFromClause(1)
        
        sWhereClause(1) = " WHERE  " & DBW("d.workarea =", pWorkArea) & _
                          " AND    " & DBW("d.accdt =", pAccDt) & _
                          " AND    " & DBW("d.accseq =", pAccSeq) & _
                          " AND    e.workarea = a.workarea " & _
                          " AND    e.accdt = a.accdt " & _
                          " AND    e.accseq = a.accseq  " & _
                          " AND    e.multifg <> ' ' " & _
                          " AND    a.ptid = d.ptid " & _
                          " AND    a.orddt = d.orddt " & _
                          " AND    a.ordno = d.ordno " & _
                          " AND    a.ordseq = d.ordseq " & _
                          " AND    " & DBW("f.cdindex =", LC2_Panel) & _
                          " AND    f.cdval1 = a.ordcd    AND    f.cdval2 = e.multifg " & _
                          " AND    b.testcd = f.field1 " & _
                          " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                          " AND    c.testcd = f.field1   AND    c.spccd = e.spccd " & _
                          " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd " & _
                                                                                        " AND   spccd  = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                          " AND    g.ptid = a.ptid   AND    g.testcd = f.field1 AND g.spccd = d.spccd " & _
                          " AND    g.vfydt = (SELECT max(vfydt) FROM " & T_LAB302 & " WHERE ptid = a.ptid  " & _
                                                                                    " AND  testcd = f.field1 AND spccd = d.spccd " & _
                                                                                    " AND  vfydt <> ' ' " & _
                                                                                    " AND  vfydt < e.rcvdt) " & _
                          " AND    exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  " & _
                                                                        " AND  testcd = f.field1 AND spccd = d.spccd " & _
                                                                        " AND  vfydt <> ' ' " & _
                                                                        " AND  vfydt < e.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "AND    not exists (SELECT * FROM " & T_LAB302 & " WHERE ptid = a.ptid  AND  testcd = f.field1 AND vfydt <> '' AND vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " UNION ALL " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    End Select
        
    SqlOrderItems1 = SqlOrderItems1 & " Order by Testcd1, TestCd2 "
                               
 End Function

'%  23. 최근결과 검색 (환자ID - 검사항목)
'%      - Calling FROM [clsAccession]
Public Function SqlLastResult(ByVal pPID As String, ByVal pTcd As String) As String

    SqlLastResult = " SELECT rstcd, vfydt, vfytm, vfyid FROM " & T_LAB302 & _
                    " WHERE  " & DBW("ptid   = ", pPID) & _
                    " AND    " & DBW("testcd = ", pTcd) & _
                    " AND    vfydt <> ' '  AND  vfydt is not null " & _
                    " Order  By vfydt desc, vfytm desc"

End Function

Public Function SqlMicLastResult(ByVal pPID As String, ByVal pTcd As String, ByVal pSCd As String) As String

'---------------
'최근결과 갖구 오는데 이상해서 바깠음...
'Modify By legends 2003/07/30
'    If Trim(pSCd) <> "" Then
'        SqlMicLastResult = " SELECT a.rstcd, a.vfydt, a.vfytm, a.vfyid " & _
'                           " FROM " & T_LAB031 & " c, " & T_LAB201 & " b, " & T_LAB404 & " a " & _
'                           " WHERE  " & DBW("a.ptid   = ", pPID) & _
'                           " AND    (a.vfydt <> ''  AND  a.vfydt is not null) " & _
'                           " AND    b.workarea = a.workarea " & _
'                           " AND    b.accdt = a.accdt " & _
'                           " AND    b.accseq = a.accseq " & _
'                           " AND    " & DBW("b.spccd = ", pSCd) & _
'                           " AND    " & DBW("c.cdindex = ", LC2_ItemResult) & _
'                           " AND    " & DBW("c.cdval1  = ", pTCd) & _
'                           " AND    c.cdval1 = a.testcd "
'    Else
'        SqlMicLastResult = " SELECT a.rstcd, a.vfydt, a.vfytm, a.vfyid " & _
'                           " FROM " & T_LAB031 & " b, " & T_LAB404 & " a " & _
'                           " WHERE  " & DBW("a.ptid   = ", pPID) & _
'                           " AND    (a.vfydt <> ''  or  a.vfydt is not null) " & _
'                           " AND    " & DBW("b.cdindex = ", LC2_ItemResult) & _
'                           " AND    " & DBW("b.cdval1  = ", pTCd) & _
'                           " AND    b.cdval1 = a.testcd "
'
'    End If
'    SqlMicLastResult = SqlMicLastResult & " Order  By vfydt desc, vfytm desc"
'--------------

 SqlMicLastResult = " SELECT a.rstcd, a.vfydt, a.vfytm, a.vfyid ,b.rcvdt,b.rcvtm FROM " & T_LAB404 & " a, " & T_LAB201 & " b " & _
                    " WHERE " & DBW("a.ptid=", pPID) & _
                    " AND    (a.vfydt <> ''  or  a.vfydt is not null) " & _
                    " AND " & DBW("a.testcd=", pTcd) & _
                    " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                    " AND a.workarea=b.workarea " & _
                    " AND a.accdt=b.accdt " & _
                    " AND a.accseq=b.accseq " & _
                    " AND " & DBW("b.spccd=", pSCd) & _
                    " Order  By b.rcvdt desc, b.rcvtm desc"
End Function



'%  27. 처방Body내역 Update
'%      - Calling FROM [clsAccession]
Public Function SqlStatusUpdate(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                ByVal pOrdSeq As Integer, ByVal pStsCd As String, _
                                ByVal pDate As String, ByVal pTime As String, Optional ByVal pDoneFg As Variant) As String
    Dim tmpStr As String
    Dim tmpStr1 As String
    
    '** 원본 ==========================================================================
'    If pStsCd = enStsCd.StsCd_LIS_Accession Then   '접수
'        tmpStr = "  " & DBW("rcvdt =", pDate) & ", " & _
'                 "  " & DBW("rcvtm =", pTime)
'    Else
'        tmpStr = ""
'    End If
'
'    If IsMissing(pDoneFg) Then
'        tmpStr1 = ""
'    Else
'        tmpStr1 = "  " & DBW("donefg =", pDoneFg) & ", "
'    End If
'    SqlStatusUpdate = " Update " & T_LAB102 & " " & _
'                      " Set    " & DBW("stscd  =", pStsCd) & ", " & tmpStr1 & tmpStr & _
'                      " WHERE  " & DBW("ptid   =", pPtid) & _
'                      " AND    " & DBW("orddt  =", pOrdDt) & _
'                      " AND    " & DBW("ordno  =", pOrdNo) & _
'                      " AND    " & DBW("ordseq =", pOrdSeq) & _
'                      " AND    " & DBW("stscd <= ", pStsCd)
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim Rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(pPtId) & _
             "    and orddt = " & DBS(pOrdDt) & _
             "    and ordno = " & DBN(pOrdNo)
             
    Set Rs = DbCon.Execute(strSQL, , adCmdText)

    If Rs.EOF = False Then
        strWrkDiv = "" & Rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    Rs.Close: Set Rs = Nothing
    
    If pStsCd = enStsCd.StsCd_LIS_Accession Then   '접수
        tmpStr = " acptdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') "
        
        '-- 2014.07.29 osw
        'tmpStr = tmpStr & ", execdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') "
        'tmpStr = tmpStr & ", examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') "
        If strWrkDiv = "3" Or strWrkDiv = "4" Then         '종건/일건
            tmpStr = tmpStr & ", examdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') "
        Else
            tmpStr = tmpStr & ", procstat = 'E' "
            tmpStr = tmpStr & ", execdate = TO_DATE(" & DBS(pDate & pTime) & ", 'yyyymmdd hh24:mi:ss') "
        End If
    Else
        tmpStr = ""
    End If
    
    If IsMissing(pDoneFg) Then
        tmpStr1 = ""
    Else
        tmpStr1 = "  " & DBW("donefg =", pDoneFg) & ", "
    End If

'SELECT * FROM sg2examt  -- examdate
'SELECT * FROM mdexmort  -- execdate
    
    If strWrkDiv = "3" Then         '종건
        '** 종건 Update
        SqlStatusUpdate = " Update   su2examt " & _
                          " Set    " & DBW("stscd  =", pStsCd) & ", " & tmpStr1 & tmpStr & _
                          " WHERE    patno = " & DBS(pPtId) & _
                          " AND      orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                          " AND      ordseqno = " & DBN(pOrdNo) & _
                          " AND      " & DBW("stscd <= ", pStsCd)
    ElseIf strWrkDiv = "4" Then     '일건
        '** 일건 Update
        SqlStatusUpdate = " Update   sg2examt " & _
                          " Set    " & DBW("stscd  =", pStsCd) & ", " & tmpStr1 & tmpStr & _
                          " WHERE    patno = " & DBS(pPtId) & _
                          " AND      orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                          " AND      ordseqno = " & DBN(pOrdNo) & _
                          " AND      " & DBW("stscd <= ", pStsCd)
    Else
        'mdexmort
        SqlStatusUpdate = " Update mdexmort " & _
                          " Set    " & DBW("stscd  =", pStsCd) & ", " & tmpStr1 & tmpStr & _
                          " WHERE  patno = " & DBS(pPtId) & _
                          " AND    orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                          " AND    ordseqno = " & DBN(pOrdNo) & _
                          " AND    " & DBW("stscd <= ", pStsCd)
    End If
    '===================================================================================

End Function



'%  27-2. 처방Header내역 Update (DoneFg)
'%      - Calling FROM [clsAccession]
Public Function SqlHeaderUpdate(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                ByVal pDoneFg As String, ByVal pFlag As Integer) As String
    If pFlag = 1 Then
        SqlHeaderUpdate = " SELECT * FROM " & T_LAB102 & " " & _
                          " WHERE     " & DBW("ptid   =", pPtId) & _
                          " AND       " & DBW("orddt  =", pOrdDt) & _
                          " AND       " & DBW("ordno  =", pOrdNo) & _
                          " AND       " & DBW("donefg <", pDoneFg)
    Else
        '** 원본 ==========================================================================
'        SqlHeaderUpdate = " Update " & T_LAB101 & " " & _
'                          " Set       " & DBW("donefg =", pDoneFg) & _
'                          " WHERE     " & DBW("ptid   =", pPtid) & _
'                          " AND       " & DBW("orddt  =", pOrdDt) & _
'                          " AND       " & DBW("ordno  =", pOrdNo)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        Dim strWrkDiv   As String
        Dim strSQL      As String
        Dim Rs          As New ADODB.Recordset
        
        strSQL = " select wrkdiv from " & T_LAB102 & _
                 "  where ptid = " & DBS(pPtId) & _
                 "    and orddt = " & DBS(pOrdDt) & _
                 "    and ordno = " & DBS(pOrdNo)
                 
        Set Rs = DbCon.Execute(strSQL, , adCmdText)
        If Rs.EOF = False Then
            strWrkDiv = "" & Rs.Fields("wrkdiv").Value
        Else
            strWrkDiv = ""
        End If
        
        Rs.Close: Set Rs = Nothing
        
        If strWrkDiv = "3" Then         '종건
            '** 종건 Update
            SqlHeaderUpdate = " Update    su2examt " & _
                              " Set       " & DBW("donefg =", pDoneFg) & _
                              " WHERE     patno = " & DBS(pPtId) & _
                              " AND       orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                              " AND       ordseqno = " & DBN(pOrdNo)
        ElseIf strWrkDiv = "4" Then     '일건
                '** 일건 Update
                SqlHeaderUpdate = " Update    sg2examt " & _
                                  " Set       " & DBW("donefg =", pDoneFg) & _
                                  " WHERE     patno = " & DBS(pPtId) & _
                                  " AND       orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                                  " AND       ordseqno = " & DBN(pOrdNo)
        Else
            SqlHeaderUpdate = " Update    mdexmort " & _
                              " Set       " & DBW("donefg =", pDoneFg) & _
                              " WHERE     patno = " & DBS(pPtId) & _
                              " AND       orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                              " AND       ordseqno = " & DBN(pOrdNo)
        End If
        '===================================================================================
    End If

End Function


'%  28. 연속검사인 경우 채혈접수내역의 검체Seq를 Key로 해당항목 검색, 또는 전체 검색
'%       - Calling FROM [clsAccession]
'%       -CdIndex = 'C101' : 그룹처방(Field1), 'C103' : 상세항목(CdVal2)
Public Function SqlPanelTest(ByVal PtId As String, ByVal cdindex As String, ByVal TestCd As String, _
                             ByVal SpcCd As String, Optional ByVal MultiFg As Variant) As String
    
    Dim tmpStr As String
    Dim TestCdField As String
   
    If Not IsMissing(MultiFg) Then tmpStr = " AND      " & DBW("a.cdval2 =", MultiFg)
    
    If cdindex = LC2_Panel Then
       TestCdField = "a.field1"
    ElseIf cdindex = LC2_Detail Then
       TestCdField = "a.cdval2"
    End If
' ' & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB302 & " d "
'   & T_LAB302 & " d, " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c "
    SqlPanelTest = " SELECT /*+ INDEX_DESC(d, s2lab302_idx2) */ " & TestCdField & " as testcd, b.rstdiv, b.rsttype, b.panelfg, b.outlabcd, c.labelcnt,  c.statflags, " & _
                   "          c.rstunit, d.rstcd as LastRst, d.vfydt as LastVfyDt, d.vfytm as LastVfyTm, d.vfyid as LastVfyId " & _
                   " FROM   " & T_LAB302 & " d, " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                   " WHERE  " & DBW("a.cdindex = ", cdindex) & _
                   " AND    " & DBW("a.cdval1  = ", TestCd) & tmpStr & _
                   " AND    b.testcd  = " & TestCdField & _
                   " AND    b.applydt =   (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                   " AND    c.testcd  = " & TestCdField & _
                   " AND    " & DBW("c.spccd   = ", SpcCd) & _
                   " AND    c.applydt =   (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd AND spccd = c.spccd) " & _
                   " AND    " & DBW("d.ptid    = ", PtId) & _
                   " AND    d.testcd = " & TestCdField & " AND  " & DBW("d.spccd   = ", SpcCd) & _
                   " AND    d.vfydt" & FUNC_CONCAT & "d.vfytm = (SELECT /*+ INDEX_DESC(" & T_LAB302 & ", s2lab302_idx2) */ max(vfydt" & FUNC_CONCAT & "vfytm) FROM " & T_LAB302 & _
                                             " WHERE ptid   = d.ptid " & _
                                             " AND   testcd = d.testcd  AND  spccd = d.spccd " & _
                                             " AND   vfydt  <> ' '  AND  vfydt is not null) " & _
                   " AND    exists (SELECT /*+ INDEX_DESC(" & T_LAB302 & ", s2lab302_idx2) */ ptid FROM " & T_LAB302 & _
                                  " WHERE  " & DBW("ptid   = ", PtId) & _
                                  " AND    testcd = " & TestCdField & " AND  " & DBW("spccd   = ", SpcCd) & _
                                  " AND    vfydt <> ' '  AND  vfydt is not null ) "
    SqlPanelTest = SqlPanelTest & " UNION ALL " & _
                   " SELECT  " & TestCdField & " as testcd, b.rstdiv, b.rsttype, b.panelfg, b.outlabcd, c.labelcnt,  c.statflags, " & _
                   "          c.rstunit, '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, " & DBV("vfyid", "0") & " as LastVfyId " & _
                   " FROM   " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                   " WHERE  " & DBW("a.cdindex = ", cdindex) & _
                   " AND    " & DBW("a.cdval1  = ", TestCd) & tmpStr & _
                   " AND    b.testcd  = " & TestCdField & _
                   " AND    b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                   " AND    c.testcd  = " & TestCdField & _
                   " AND    " & DBW("c.spccd   = ", SpcCd) & _
                   " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = c.testcd AND spccd = c.spccd) " & _
                   " AND    not exists  (SELECT  /*+ INDEX_DESC(" & T_LAB302 & ", s2lab302_idx2) */ ptid FROM " & T_LAB302 & _
                                      "  WHERE " & DBW("ptid   = ", PtId) & _
                                      "  AND   testcd = " & TestCdField & " AND  " & DBW("spccd   = ", SpcCd) & _
                                      "  AND   vfydt <> ' '  AND  vfydt is not null) "

End Function


Public Function SqlGetTotCnt(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

    SqlGetTotCnt = " SELECT count(*) as TotCnt FROM " & T_LAB302 & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq) & _
                   " AND   (detailfg = '' or detailfg is null or rstdiv = '*') "
    SqlGetTotCnt = SqlGetTotCnt & " UNION ALL " & _
                   " SELECT count(*) as TotCnt FROM " & T_LAB351 & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq)
    SqlGetTotCnt = SqlGetTotCnt & " UNION ALL " & _
                   " SELECT count(*) as TotCnt FROM " & T_LAB404 & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq) & _
                   " AND   (detailfg = '' or detailfg is null  or rstdiv = '*') "
    
End Function

Public Function SqlGetInputCnt(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

    SqlGetInputCnt = " SELECT count(*) as TotCnt FROM " & T_LAB302 & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq) & _
                   " AND   (detailfg = '' or detailfg is null or rstdiv = '*') " & _
                   " AND    vfydt is not null  AND  vfydt <> ' '"
    SqlGetInputCnt = SqlGetInputCnt & " UNION ALL " & _
                   " SELECT count(*) as TotCnt FROM " & T_LAB351 & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq) & _
                   " AND    vfydt is not null  AND  vfydt <> ' '"
    SqlGetInputCnt = SqlGetInputCnt & " UNION ALL " & _
                   " SELECT count(*) as TotCnt FROM " & T_LAB404 & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq) & _
                   " AND   (detailfg = '' or detailfg is null  or rstdiv = '*') " & _
                   " AND    vfydt is not null  AND  vfydt <> ' '"
    
End Function



Public Function CreateSql_Result(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByRef MyResult As Object) As String
   
    With MyResult
        Select Case .TestDiv
        Case TST_RouTest, TST_AboTest:    '"0", "3":  '일반검사/ABO검사
            CreateSql_Result = "Insert into " & T_LAB302 & _
                               "(workarea, accdt, accseq, testcd, rstval, rstcd, rstunit, " & _
                               "hldiv, dpdiv, spccd, statfg, lastrst, lastvfydt, lastvfytm, lastvfyid, " & _
                               "vfydt, vfytm, vfyid, attrcd, mfyfg, grpfg, txtfg, autofg, rsttype, rstdiv, " & _
                               "detailfg, eqpcd, ptid, orddt, ordno, ordseq) " & _
                               "Values (" & _
                                DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                                DBV("testcd", .TestCd, 1) & DBV("rstval", .RstVal, 1) & DBV("rstcd", .RstCd, 1) & _
                                DBV("rstunit", .RstUnit, 1) & DBV("hldiv", .HLDiv, 1) & DBV("dpdiv", .DPDiv, 1) & _
                                DBV("spccd", .SpcCd, 1) & DBV("statfg", .StatFg, 1) & DBV("lastrst", .LastRst, 1) & _
                                DBV("lastvfydt", .LastVfyDt, 1) & DBV("lastvfytm", .LastVfyTm, 1) & _
                                DBV("lastvfyid", .LastVfyId1, 1) & DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & _
                                DBV("vfyid", .VfyId1, 1) & DBV("attrcd", .AttrCd, 1) & DBV("mfyfg", .MfyFg, 1) & _
                                DBV("grpfg", .GrpFg, 1) & DBV("txtfg", .TxtFg, 1) & DBV("autofg", .AutoFg, 1) & _
                                DBV("rsttype", .RstType, 1) & DBV("rstdiv", .RstDiv, 1) & DBV("detailfg", .DetailFg, 1) & _
                                DBV("eqpcd", .EqpCd, 1) & DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & _
                                DBV("ordno", .OrdNo, 1) & DBV("ordseq", .OrdSeq) & " )"
      
        Case TST_SpeTest: '"1":   '기타검사
            CreateSql_Result = "Insert into " & T_LAB351 & _
                               "(workarea, accdt, accseq, testcd, txtfg, valfg, stscd, rsttype, " & _
                               "mfyseq, vfydt, vfytm, vfyid, ptid, orddt, ordno, ordseq) " & _
                               "Values (" & _
                                DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                                DBV("testcd", .TestCd, 1) & DBV("txtfg", .TxtFg, 1) & DBV("valfg", .ValFg, 1) & _
                                DBV("stscd", StsCd_LIS_Accession, 1) & DBV("rsttype", .RstType, 1) & _
                                DBV("mfyseq", .MfyFg, 1) & DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & _
                                DBV("vfyid", .VfyId1, 1) & DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & _
                                DBV("ordno", .OrdNo, 1) & DBV("ordseq", .OrdSeq) & ")"
        
        Case TST_MicTest: '"2":   '미생물검사
            CreateSql_Result = "Insert into " & T_LAB404 & _
                               "(workarea, accdt, accseq, testcd, rstcd, senfg, stscd, rsttype,mfyseq, rstdiv, " & _
                               "detailfg, vfydt, vfytm, vfyid, lastrst, lastvfydt, lastvfytm, lastvfyid, ptid, " & _
                               "orddt, ordno, ordseq) " & _
                               "Values (" & _
                                DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                                DBV("testcd", .TestCd, 1) & DBV("rstcd", .RstCd, 1) & DBV("senfg", .SenFg, 1) & _
                                DBV("stscd", StsCd_LIS_Accession, 1) & DBV("rsttype", .RstType, 1) & _
                                DBV("mfyseq", .MfyFg, 1) & DBV("rstdiv", .RstDiv, 1) & DBV("detailfg", .DetailFg, 1) & _
                                DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & DBV("vfyid", .VfyId1, 1) & _
                                DBV("lastrst", .LastRst, 1) & DBV("lastvfydt", .LastVfyDt, 1) & _
                                DBV("lastvfytm", .LastVfyTm, 1) & DBV("lastvfyid", .LastVfyId1, 1) & _
                                DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & _
                                DBV("ordno", .OrdNo, 1) & DBV("ordseq", .OrdSeq) & ")"
            
        End Select
    End With
   
End Function

Public Function CreateSql_Result_Temp(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByRef MyResult As Object) As String
   
    With MyResult
        Select Case .TestDiv
        Case TST_RouTest, TST_AboTest:    '"0", "3":  '일반검사/ABO검사
            CreateSql_Result_Temp = "Insert into S2LAB320 " & _
                               "(workarea, accdt, accseq, testcd, rstval, rstcd, rstunit, " & _
                               "hldiv, dpdiv, spccd, statfg, lastrst, lastvfydt, lastvfytm, lastvfyid, " & _
                               "vfydt, vfytm, vfyid, attrcd, mfyfg, grpfg, txtfg, autofg, rsttype, rstdiv, " & _
                               "detailfg, eqpcd, ptid, orddt, ordno, ordseq) " & _
                               "Values (" & _
                                DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                                DBV("testcd", .TestCd, 1) & DBV("rstval", .RstVal, 1) & DBV("rstcd", .RstCd, 1) & _
                                DBV("rstunit", .RstUnit, 1) & DBV("hldiv", .HLDiv, 1) & DBV("dpdiv", .DPDiv, 1) & _
                                DBV("spccd", .SpcCd, 1) & DBV("statfg", .StatFg, 1) & DBV("lastrst", .LastRst, 1) & _
                                DBV("lastvfydt", .LastVfyDt, 1) & DBV("lastvfytm", .LastVfyTm, 1) & _
                                DBV("lastvfyid", .LastVfyId1, 1) & DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & _
                                DBV("vfyid", .VfyId1, 1) & DBV("attrcd", .AttrCd, 1) & DBV("mfyfg", .MfyFg, 1) & _
                                DBV("grpfg", .GrpFg, 1) & DBV("txtfg", .TxtFg, 1) & DBV("autofg", .AutoFg, 1) & _
                                DBV("rsttype", .RstType, 1) & DBV("rstdiv", .RstDiv, 1) & DBV("detailfg", .DetailFg, 1) & _
                                DBV("eqpcd", .EqpCd, 1) & DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & _
                                DBV("ordno", .OrdNo, 1) & DBV("ordseq", .OrdSeq) & " )"
      
'''        Case TST_SpeTest: '"1":   '기타검사
'''            CreateSql_Result_Temp = "Insert into " & T_LAB351 & _
'''                               "(workarea, accdt, accseq, testcd, txtfg, valfg, stscd, rsttype, " & _
'''                               "mfyseq, vfydt, vfytm, vfyid, ptid, orddt, ordno, ordseq) " & _
'''                               "Values (" & _
'''                                DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
'''                                DBV("testcd", .TestCd, 1) & DBV("txtfg", .TxtFg, 1) & DBV("valfg", .ValFg, 1) & _
'''                                DBV("stscd", StsCd_LIS_Accession, 1) & DBV("rsttype", .RstType, 1) & _
'''                                DBV("mfyseq", .MfyFg, 1) & DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & _
'''                                DBV("vfyid", .VfyId1, 1) & DBV("ptid", .Ptid, 1) & DBV("orddt", .OrdDt, 1) & _
'''                                DBV("ordno", .OrdNo, 1) & DBV("ordseq", .OrdSeq) & ")"
'''
'''        Case TST_MicTest: '"2":   '미생물검사
'''            CreateSql_Result_Temp = "Insert into " & T_LAB404 & _
'''                               "(workarea, accdt, accseq, testcd, rstcd, senfg, stscd, rsttype,mfyseq, rstdiv, " & _
'''                               "detailfg, vfydt, vfytm, vfyid, lastrst, lastvfydt, lastvfytm, lastvfyid, ptid, " & _
'''                               "orddt, ordno, ordseq) " & _
'''                               "Values (" & _
'''                                DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
'''                                DBV("testcd", .TestCd, 1) & DBV("rstcd", .RstCd, 1) & DBV("senfg", .SenFg, 1) & _
'''                                DBV("stscd", StsCd_LIS_Accession, 1) & DBV("rsttype", .RstType, 1) & _
'''                                DBV("mfyseq", .MfyFg, 1) & DBV("rstdiv", .RstDiv, 1) & DBV("detailfg", .DetailFg, 1) & _
'''                                DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & DBV("vfyid", .VfyId1, 1) & _
'''                                DBV("lastrst", .LastRst, 1) & DBV("lastvfydt", .LastVfyDt, 1) & _
'''                                DBV("lastvfytm", .LastVfyTm, 1) & DBV("lastvfyid", .LastVfyId1, 1) & _
'''                                DBV("ptid", .Ptid, 1) & DBV("orddt", .OrdDt, 1) & _
'''                                DBV("ordno", .OrdNo, 1) & DBV("ordseq", .OrdSeq) & ")"
            
        End Select
    End With
   
End Function

Public Function CreateSql_ABOResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, ByRef MyResult As Object) As String
'** 2001.2.14 추가 : ABO 검사는 h7bbs303에 결과내역을 생성한다.

    With MyResult
        CreateSql_ABOResult = "Insert into " & T_BBS303 & _
                              "(workarea, accdt, accseq, seq, spccd, statfg, abo1, rh1, abo2, rh2, abosub, rhsub, " & _
                              "vfyid1, vfydt1, vfytm1, mfyfg1, vfyid2, vfydt2, vfytm2, mfyfg2, vfydt, ptid) " & _
                              "Values (" & _
                               DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                               DBV("seq", 0, 1) & DBV("spccd", .SpcCd, 1) & DBV("statfg", .StatFg, 1) & _
                               DBV("abo1", "", 1) & DBV("rh1", "", 1) & DBV("abo2", "", 1) & DBV("rh2", "", 1) & _
                               DBV("abosub", "", 1) & DBV("rhsub", "", 1) & DBV("vfyid1", "", 1) & _
                               DBV("vfydt1", "", 1) & DBV("vfytm1", "", 1) & DBV("mfyfg1", "", 1) & _
                               DBV("vfyid2", "", 1) & DBV("vfydt2", "", 1) & _
                               DBV("vfytm2", "", 1) & DBV("mfyfg2", "", 1) & DBV("vfydt", "", 1) & DBV("ptid", .PtId) & ")"
    End With
    
End Function

Public Function CreateSql_OutLab(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByRef MyOutLab As Object) As String
    
    With MyOutLab
        CreateSql_OutLab = "Insert into " & T_LAB205 & _
                           "(workarea, accdt, accseq, testcd, ptid, spccd, outlabcd," & _
                           "stscd, rcvdt, senddt, chargedt, sendid) " & _
                           "Values ( " & _
                           DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & _
                           DBV("testcd", .TestCd, 1) & DBV("ptid", .PtId, 1) & DBV("spccd", .SpcCd, 1) & _
                           DBV("outlabcd", .OutLabCd, 1) & DBV("stscd", StsCd_LIS_Accession, 1) & _
                           DBV("rcvdt", .RcvDt, 1) & DBV("senddt", .SendDt, 1) & DBV("chargedt", .ChargeDt, 1) & _
                           DBV("sendid", .SendId) & " )"
   End With
                              
End Function


Public Function CreateSql_ColUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, ByVal pStsCd As String, ByVal pReqTotCnt As Integer, _
                                    ByVal pRcvDt As String, ByVal pRcvTm As String, ByVal pRcvId As String)
    
    CreateSql_ColUpdate = " Update " & T_LAB201 & _
                          " Set    " & DBW("stscd", StsCd_LIS_Accession, 3) & DBW("reqtotcnt", pReqTotCnt, 3) & _
                          "        " & DBW("rcvdt", pRcvDt, 3) & DBW("rcvtm", pRcvTm, 3) & _
                          "        " & DBW("rcvid", pRcvId, 2) & _
                          " WHERE  " & DBW("workarea =", pWorkArea) & _
                          " AND    " & DBW("accdt    =", pAccDt) & _
                          " AND    " & DBW("accseq   =", pAccSeq)
                         
End Function


'%  43. 일반접수시 처방 Display
'%       - Calling FROM [ frm155Accession ] :
Public Function SqlOrdersForAccess(ByVal QueryOption As Integer, Optional ByVal Key1 As Variant, _
                                   Optional ByVal Key2 As Variant, Optional ByVal Key3 As Variant) As String

    Dim tmpStr As String

    If QueryOption = 1 Then
        tmpStr = " WHERE   " & DBW("a.workarea =", Key1) & _
                 " AND     " & DBW("a.accdt    =", Key2) & _
                 " AND     " & DBW("a.accseq   =", Key3)
    Else
        tmpStr = " WHERE   " & DBW("a.spcyy =", Key1) & _
                 " AND     " & DBW("a.spcno =", Key2)
    End If
                 
' 08.10.23. 양성현 검체 보관방법 명칭추가 변경
    
    SqlOrdersForAccess = " SELECT a.workarea, a.accdt, a.accseq, a.ptid, b." & F_PTNM & " as PtNm, c." & F_DEPTNM & " as DeptNm, " & _
                         "        a.wardid" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.hosilid as Location, a.stscd, a.statfg, a.storecd, " & _
                         "        d.field3 as SpcNm, e.orddt, e.ordcd, f.testnm ,e.mesg, g.field1 as storenm" & _
                         " FROM " & T_LAB201 & " a, " & T_HIS001 & " b, " & T_HIS003 & " c, " & _
                                    T_LAB032 & " d, " & T_LAB102 & " e, " & T_LAB001 & " f, " & T_LAB032 & " g " & _
                           tmpStr & _
                         " AND   (a.multifg  =  '' or a.multifg is null) " & _
                         " AND    b." & F_PTID & "  =  a.ptid " & _
                         " AND    " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                         " AND    " & DBW("d.cdindex  =", LC3_Specimen) & _
                         " AND    d.cdval1   =  a.spccd " & _
                         " AND    e.workarea =  a.workarea " & _
                         " AND    e.accdt    =  a.accdt " & _
                         " AND    e.accseq   =  a.accseq " & _
                         " AND   (e.dcfg     =  '' or e.dcfg is null ) " & _
                         " AND    f.testcd   =  e.ordcd " & _
                         " AND    f.applydt  = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = f.testcd) " & _
                         " and    g.cdindex = 'C229' and a.storecd = g.cdval1 "
      
      SqlOrdersForAccess = SqlOrdersForAccess & " UNION ALL " & _
                         " SELECT a.workarea, a.accdt, a.accseq, a.ptid, b." & F_PTNM & " as ptnm, c." & F_DEPTNM & " as DeptNm, " & _
                         "        a.wardid" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.hosilid as Location, a.stscd, a.statfg, a.storecd, " & _
                         "        d.field3 as SpcNm, e.orddt, e.ordcd, f.testnm,e.mesg, h.field1 as storenm " & _
                         " FROM " & T_LAB201 & " a, " & T_HIS001 & " b, " & T_HIS003 & " c, " & _
                                    T_LAB032 & " d, " & T_LAB102 & " e, " & T_LAB001 & " f, " & T_LAB203 & " g, " & T_LAB032 & " h " & _
                           tmpStr & _
                         " AND    a.multifg <> ' ' " & _
                         " AND    b." & F_PTID & "  =  a.ptid " & _
                         " AND    " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                         " AND    " & DBW("d.cdindex  =", LC3_Specimen) & _
                         " AND    d.cdval1   =  a.spccd " & _
                         " AND    a.workarea =  g.workarea " & _
                         " AND    a.accdt    =  g.accdt " & _
                         " AND    a.accseq   =  g.accseq " & _
                         " AND    e.ptid     =  g.ptid " & _
                         " AND    e.orddt    =  g.orddt " & _
                         " AND    e.ordno    =  g.ordno " & _
                         " AND    e.ordseq   =  g.ordseq " & _
                         " AND   (e.dcfg     =  '' or e.dcfg is null ) " & _
                         " AND    f.testcd   =  e.ordcd " & _
                         " AND    f.applydt  = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = f.testcd) " & _
                         " and    h.cdindex = 'C229' and a.storecd = h.cdval1 " & _
                         " Order by orddt "

End Function



'----------------------------------------------------------------------------------
'3. 접수취소에 관련된 SQL문
'----------------------------------------------------------------------------------

Public Function SQLGetCancelReason() As String

    SQLGetCancelReason = " SELECT cdval1, text1 FROM " & T_LAB034 & _
                         " WHERE  " & DBW("cdindex =", LC4_CancelReason)

End Function


'%  48. 접수번호로 Status Check
'%       - Calling FROM [ frm108AccCancel ] :
Public Function SqlCheckStatus(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String) As String
      
    SqlCheckStatus = " SELECT a.stscd, a.ptid, a.orddoct, a.deptcd, a.wardid, a.roomid, a.bedid, a.hosilid, a.spccd, " & _
                     "        a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, a.multifg, b.field3 as SpcNm " & _
                     " FROM  " & T_LAB201 & " a, " & T_LAB032 & " b " & _
                     " WHERE " & DBW("a.workarea", WorkArea, 2) & _
                     " AND   " & DBW("a.accdt", AccDt, 2) & _
                     " AND   " & DBW("a.accseq", AccSeq, 2) & _
                     " AND   " & DBW("b.cdindex", LC3_Specimen, 2) & _
                     " AND   b.cdval1 = a.spccd "
            
End Function


'%  49. 접수번호로 결과내역 검색
'%       - Calling FROM [ frm108AccCancel ] :
Public Function SqlAccOrder(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                            Optional ByVal pMultiFg As Boolean = False) As String
      
    If Not pMultiFg Then
        SqlAccOrder = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, a.examdt as vfydt, b.testnm " & _
                        " FROM  " & T_LAB102 & " a, " & T_LAB001 & " b " & _
                        " WHERE " & DBW("a.workarea", pWorkArea, 2) & _
                        " AND   " & DBW("a.accdt", pAccDt, 2) & _
                        " AND   " & DBW("a.accseq", pAccSeq, 2) & _
                        " AND   b.testcd = a.ordcd " & _
                        " ORDER BY orddt, ordno, ordseq"
    Else
        SqlAccOrder = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, a.examdt as vfydt, b.testnm " & _
                        " FROM  " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB203 & " c " & _
                        " WHERE " & DBW("c.workarea", pWorkArea, 2) & _
                        " AND   " & DBW("c.accdt", pAccDt, 2) & _
                        " AND   " & DBW("c.accseq", pAccSeq, 2) & _
                        " AND   a.ptid = c.ptid  AND a.orddt = c.orddt AND a.ordno = c.ordno AND a.ordseq = c.ordseq " & _
                        " AND   b.testcd = a.ordcd " & _
                        " ORDER BY orddt, ordno, ordseq"
    End If
'    End If
               
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수취소시 처방내역 검색
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : AccDt
'       - pAccSeq   : AccSeq
'-----------------------------------------------------------------------------'
Public Function SqlAccOrderX(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim SQL As String
    
    '## 복수검체가 아닌 처방조회
    SQL = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, a.examdt as vfydt, a.stscd, b.testnm, '' as multifg " & _
            " FROM  " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c" & _
            " WHERE " & DBW("a.workarea", pWorkArea, 2) & _
            " AND   " & DBW("a.accdt", pAccDt, 2) & _
            " AND   " & DBW("a.accseq", pAccSeq, 2) & _
            " AND   b.testcd = a.ordcd " & _
            " AND   a.workarea=c.workarea AND a.accdt=c.accdt AND a.accseq=c.accseq" & _
            " AND   (c.multifg IS NULL OR c.multifg='')"
    
    '## 복수검체 처방조회
    SQL = SQL & " UNION ALL " & _
            " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, a.examdt as vfydt, a.stscd, b.testnm, 'Y' as multifg " & _
            " FROM  " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB203 & " c " & _
            " WHERE " & DBW("c.workarea", pWorkArea, 2) & _
            " AND   " & DBW("c.accdt", pAccDt, 2) & _
            " AND   " & DBW("c.accseq", pAccSeq, 2) & _
            " AND   a.ptid = c.ptid  AND a.orddt = c.orddt AND a.ordno = c.ordno AND a.ordseq = c.ordseq " & _
            " AND   b.testcd = a.ordcd " & _
            " ORDER BY orddt, ordno, ordseq"
            
    SqlAccOrderX = SQL
End Function

Public Function SqlDelRstTable(ByVal pTblNm As String, ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    
    SqlDelRstTable = " delete FROM  " & pTblNm & _
                     " WHERE " & DBW("workarea", pWorkArea, 2) & _
                     " AND   " & DBW("accdt", pAccDt, 2) & _
                     " AND   " & DBW("accseq", pAccSeq, 2)
    
End Function

Public Function SqlDelRstTableForMic(ParamArray pData()) As String
'미생물 워크시트 내역 삭제
'Coding By Legends 2003/10/15

'공통
'pdata(0):테이블명
'401 지울때
'pdata(1) : wscd, pdata(2) : wsunit
'402 지울때
'pdata(1) : workarea, pdata(2) : accdt, pdata(3) : accseq

    Dim strTblNm As String
    
    Select Case pData(0)
        Case T_LAB401
            SqlDelRstTableForMic = " delete FROM  " & pData(0) & _
                                   " WHERE " & DBW("wscd=", pData(1)) & _
                                   " AND " & DBW("wsunit=", pData(2))
        Case T_LAB402
            SqlDelRstTableForMic = " delete FROM  " & pData(0) & _
                                   " WHERE " & DBW("workarea=", pData(1)) & _
                                   " AND   " & DBW("accdt=", pData(2)) & _
                                   " AND   " & DBW("accseq=", pData(3))
    End Select
End Function

'-----------------------------------------------------------------------------'
'   기능 : Worksheet 내역의 부분삭제 - 이상대(2004-12-13)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pPtId     : 환자ID
'       - pOrdDt    : 처방일자
'       - pOrdNo    : 처방번호
'       - pOrdSeq   : 처방순번
'-----------------------------------------------------------------------------'
Public Function SqlDelRstTableForMicPart(ByVal pWorkArea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As String, ByVal pPtId As String, ByVal pOrdDt As String, _
                    ByVal pOrdNo As String, ByVal pOrdSeq As String) As String
    Dim SQL As String

    SQL = " DELETE FROM " & T_LAB402 & " a" & _
          " WHERE EXISTS (SELECT * FROM " & T_LAB404 & " b" & _
          "             WHERE " & DBW("b.workarea", pWorkArea, 2) & _
          "             AND " & DBW("b.accdt", pAccDt, 2) & _
          "             AND " & DBW("b.accseq", pAccSeq, 2) & _
          "             AND " & DBW("b.orddt", pOrdDt, 2) & _
          "             AND " & DBW("b.ordno", pOrdNo, 2) & _
          "             AND " & DBW("b.ordseq", pOrdSeq, 2) & _
          "             AND b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq" & _
          "             AND b.rsttype=a.scfg)"

    SqlDelRstTableForMicPart = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkSheet Code, Unit으로 WorkSheet 내역조회
'-----------------------------------------------------------------------------'
Public Function SqlGetMicBody(ByVal pWsCd As String, ByVal pWsUnit As String)
    SqlGetMicBody = " SELECT workarea, accdt, accseq FROM " & T_LAB402 & _
                    " WHERE " & DBW("wscd=", pWsCd) & _
                    " AND " & DBW("wsunit=", pWsUnit)
End Function

Public Function SqlDelRstPart(ByVal pTblNm As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String, _
                              ByVal pPtId As String, ByVal pOrdDt As String, _
                              ByVal pOrdNo As String, ByVal pOrdSeq As String) As String

    Dim strTblNm As String
    
    Select Case pTblNm
    Case T_LAB205, T_LAB305
        SqlDelRstPart = " delete FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2) & _
                        " AND   exists (  SELECT * FROM " & T_LAB302 & _
                                        " WHERE " & DBW("ptid = ", pPtId) & _
                                        " AND   " & DBW("orddt = ", pOrdDt) & _
                                        " AND   " & DBW("ordno = ", pOrdNo) & _
                                        " AND   " & DBW("ordseq = ", pOrdSeq) & _
                                        " AND   testcd = " & pTblNm & ".testcd) "
    
    Case T_LAB405, T_LAB407
        SqlDelRstPart = " delete FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2) & _
                        " AND   exists (  SELECT * FROM " & T_LAB404 & _
                                        " WHERE workarea = " & pTblNm & ".workarea" & _
                                        " AND   accdt    = " & pTblNm & ".accdt" & _
                                        " AND   accseq   = " & pTblNm & ".accseq" & _
                                        " AND   " & DBW("ptid = ", pPtId) & _
                                        " AND   " & DBW("orddt = ", pOrdDt) & _
                                        " AND   " & DBW("ordno = ", pOrdNo) & _
                                        " AND   " & DBW("ordseq = ", pOrdSeq) & _
                                        " AND   testcd = " & pTblNm & ".testcd) "
    
    Case T_LAB302, T_LAB351, T_LAB404
        SqlDelRstPart = " delete FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2) & _
                        " AND   " & DBW("ptid = ", pPtId) & _
                        " AND   " & DBW("orddt = ", pOrdDt) & _
                        " AND   " & DBW("ordno = ", pOrdNo) & _
                        " AND   " & DBW("ordseq = ", pOrdSeq)
                        
    Case T_LAB360, T_LAB361
        SqlDelRstPart = " delete FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2)
    End Select
    
End Function

Public Function SqlPartDataExists(ByVal pTblNm As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String, _
                              ByVal pPtId As String, ByVal pOrdDt As String, _
                              ByVal pOrdNo As String, ByVal pOrdSeq As String) As String

    Dim strTblNm As String
    
    Select Case pTblNm
    Case T_LAB205, T_LAB305
        SqlPartDataExists = " SELECT * FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2) & _
                        " AND   exists (  SELECT * FROM " & T_LAB302 & _
                                        " WHERE " & DBW("ptid = ", pPtId) & _
                                        " AND   " & DBW("orddt = ", pOrdDt) & _
                                        " AND   " & DBW("ordno = ", pOrdNo) & _
                                        " AND   " & DBW("ordseq = ", pOrdSeq) & _
                                        " AND   testcd = " & pTblNm & ".testcd) "
    
    Case T_LAB405, T_LAB407
        SqlPartDataExists = " SELECT *  FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2) & _
                        " AND   exists (  SELECT * FROM " & T_LAB404 & _
                                        " WHERE workarea = " & pTblNm & ".workarea" & _
                                        " AND   accdt    = " & pTblNm & ".accdt" & _
                                        " AND   accseq   = " & pTblNm & ".accseq" & _
                                        " AND   " & DBW("ptid = ", pPtId) & _
                                        " AND   " & DBW("orddt = ", pOrdDt) & _
                                        " AND   " & DBW("ordno = ", pOrdNo) & _
                                        " AND   " & DBW("ordseq = ", pOrdSeq) & _
                                        " AND   testcd = " & pTblNm & ".testcd) "
    
    Case T_LAB302, T_LAB351, T_LAB404
        SqlPartDataExists = " SELECT *  FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2) & _
                        " AND   " & DBW("ptid = ", pPtId) & _
                        " AND   " & DBW("orddt = ", pOrdDt) & _
                        " AND   " & DBW("ordno = ", pOrdNo) & _
                        " AND   " & DBW("ordseq = ", pOrdSeq)
                        
    Case T_LAB360, T_LAB361
        SqlPartDataExists = " SELECT *  FROM  " & pTblNm & _
                        " WHERE " & DBW("workarea", pWorkArea, 2) & _
                        " AND   " & DBW("accdt", pAccDt, 2) & _
                        " AND   " & DBW("accseq", pAccSeq, 2)
    End Select
    
End Function

Public Function SqlDataExists(ByVal pTblNm As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String) As String

    SqlDataExists = " SELECT * FROM " & pTblNm & _
                    " WHERE   " & DBW("workarea= ", pWorkArea) & _
                    " AND     " & DBW("accdt   = ", pAccDt) & _
                    " AND     " & DBW("accseq  = ", pAccSeq)

End Function

Public Function SqlDataExistsForMic(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String) As String
'미생물 워크시트 내역 존재여부 판단
'Coding By Legends 2003/10/15

    SqlDataExistsForMic = " SELECT * FROM " & T_LAB401 & " a, " & T_LAB402 & " b " & _
                          " WHERE " & DBW("b.workarea=", pWorkArea) & _
                          " AND " & DBW("b.accdt=", pAccDt) & _
                          " AND " & DBW("b.accseq=", pAccSeq) & _
                          " AND a.wscd=b.wscd " & _
                          " AND a.wsunit=b.wsunit "
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 처방에 대한 Worksheet 존재하는지 조회 - 이상대(2004-12-13)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pPtId     : 환자ID
'       - pOrdDt    : 처방일자
'       - pOrdNo    : 처방번호
'       - pOrdSeq   : 처방순번
'-----------------------------------------------------------------------------'
Public Function SqlDataExistForMicPart(ByVal pWorkArea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As String, ByVal pPtId As String, ByVal pOrdDt As String, _
                    ByVal pOrdNo As String, ByVal pOrdSeq As String) As String
    Dim SQL As String

    SQL = " SELECT DISTINCT b.wscd, b.wsunit FROM " & T_LAB404 & " a, " & T_LAB402 & " b" & _
          " WHERE " & DBW("a.workarea", pWorkArea, 2) & _
          " AND " & DBW("a.accdt", pAccDt, 2) & _
          " AND " & DBW("a.accseq", pAccSeq, 2) & _
          " AND " & DBW("a.ptid", pPtId, 2) & _
          " AND " & DBW("a.orddt", pOrdDt, 2) & _
          " AND " & DBW("a.ordno", pOrdNo, 2) & _
          " AND " & DBW("a.ordseq", pOrdSeq, 2) & _
          " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          " AND a.rsttype=b.scfg"

    SqlDataExistForMicPart = SQL
End Function
'-----------------------------------------------------------------------------'
'   기능 : 접수취소시 LAB201 업데이트 - 5.0.32: 이상대(2005-08-26)
'          SqlCancel201 대용함수
'   인수 :
'       - pWorkarea, pAccDt, pAccSeq : 접수번호
'       - pStsCd : 7(처방상태로 취소), 1(채혈상태로 취소)
'       - pEmpId : 취소자 직원ID
'-----------------------------------------------------------------------------'
Public Function SqlCancel201X(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pStsCd As String, _
                             ByVal pEmpId As String) As String
    Dim SQL As String
    Dim strCancelDt As String
    Dim strCancelTm As String
    
    strCancelDt = Format$(GetSystemDate, "yyyymmdd")
    strCancelTm = Format$(GetSystemDate, "HhmmSs")

    '## 처방, 채혈상태로 취소시 entdt, enttm, entid에 취소일시, 취소자를 입력(전체선택 취소시만)
    If pStsCd = "7" Then        '## 처방상태로 취소
        SQL = " UPDATE " & T_LAB201 & _
              " SET    " & DBW("stscd ", pStsCd, 3) & _
                    DBW("reqinputcnt ", 0, 3) & _
              "     vfydt = '', vfytm = '', vfyid = '', " & _
              "     coldt = '', coltm = '', colid = '', " & _
              "     rcvdt = '', rcvtm = '', rcvid = '', " & _
                    DBW("entdt", strCancelDt, 3) & _
                    DBW("enttm", strCancelTm, 3) & _
                    DBW("entid", pEmpId, 3) & _
              "     footnotefg='0', rmkcd='' " & _
              " WHERE  " & DBW("workarea = ", pWorkArea) & _
              " AND    " & DBW("accdt    = ", pAccDt) & _
              " AND    " & DBW("accseq   = ", pAccSeq)
    
    ElseIf pStsCd = "1" Then    '## 채혈상태로 취소
        SQL = " UPDATE " & T_LAB201 & _
              " SET    " & DBW("stscd ", pStsCd, 3) & _
                    DBW("reqinputcnt ", 0, 3) & _
              "     vfydt = '', vfytm = '', vfyid = '', " & _
              "     rcvdt = '', rcvtm = '', rcvid = '', " & _
                    DBW("entdt", strCancelDt, 3) & _
                    DBW("enttm", strCancelTm, 3) & _
                    DBW("entid", pEmpId, 3) & _
              "     footnotefg='0', rmkcd='' " & _
              " WHERE  " & DBW("workarea = ", pWorkArea) & _
              " AND    " & DBW("accdt    = ", pAccDt) & _
              " AND    " & DBW("accseq   = ", pAccSeq)
    End If
    SqlCancel201X = SQL
End Function


Public Function SqlCancel201(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pStsCd As String) As String

    SqlCancel201 = " update " & T_LAB201 & _
                   " set    " & DBW("stscd ", pStsCd, 3) & _
                                DBW("reqinputcnt ", 0, 3) & _
                              " vfydt = '', vfytm = '', vfyid = 0, " & _
                              " footnotefg='0', rmkcd='' " & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq)

End Function


Public Function SqlUpdate201(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pStsCd As String, _
                             ByVal pTotCnt As Long, ByVal pInputCnt As Long) As String

    Dim strVfyDt As String
    Dim strVfyTm As String
    
    If pStsCd = enStsCd.StsCd_LIS_FinRst Then
        strVfyDt = Format(Now, CS_DateDbFormat)
        strVfyTm = Format(Now, CS_TimeDbFormat)
    Else
        strVfyDt = ""
        strVfyTm = ""
    End If
    SqlUpdate201 = " update " & T_LAB201 & _
                   " set    " & DBW("stscd ", pStsCd, 3) & _
                                DBW("reqtotcnt ", pTotCnt, 3) & _
                                DBW("reqinputcnt ", pInputCnt, 3) & _
                                DBW("vfydt ", strVfyDt, 3) & _
                                DBW("vfytm ", strVfyTm, 2) & _
                   " WHERE  " & DBW("workarea = ", pWorkArea) & _
                   " AND    " & DBW("accdt    = ", pAccDt) & _
                   " AND    " & DBW("accseq   = ", pAccSeq)

End Function


Public Function SqlInsertReason(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pEmpId As String, ByVal pReason As String) As String

    SqlInsertReason = " insert into " & T_LAB304 & _
                      " values (" & DBV("workarea", pWorkArea, 1) & _
                                    DBV("accdt", pAccDt, 1) & _
                                    DBV("accseq", pAccSeq, 1) & _
                                    DBV("seq", "1", 1) & _
                                    DBV("vfyid", pEmpId, 1) & _
                                    DBV("rsttxt", pReason) & _
                               ")"

    SqlInsertReason = " insert into " & T_LAB304 & _
                             " ( workarea, accdt, accseq, seq,  vfyid, rsttxt ) " & _
                      " values (" & DBV("workarea", pWorkArea, 1) & _
                                    DBV("accdt", pAccDt, 1) & _
                                    DBV("accseq", pAccSeq, 1) & _
                                    DBV("seq", "1", 1) & _
                                    DBV("vfyid", pEmpId, 1) & _
                                    DBV("rsttxt", pReason) & _
                              ")"

End Function

Public Function SqlUpdateReason(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pEmpId As String, ByVal pReason As String) As String

    Dim Rs As Recordset
    Dim sSQL As String
    Dim sReason As String
    
    
    sSQL = "SELECT rsttxt FROM " & T_LAB304 & _
           " WHERE  " & DBW("workarea = ", pWorkArea) & _
                    " AND    " & DBW("accdt = ", pAccDt) & _
                    " AND    " & DBW("accseq = ", pAccSeq) & _
                    " AND    " & DBW("seq = ", "1")
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        sReason = Rs.Fields("rsttxt").Value & "" & "," & pReason
    Else
        sReason = pReason
    End If
    
    Set Rs = Nothing
    
    
    SqlUpdateReason = " update " & T_LAB304 & " set " & _
                                   DBW("vfyid ", pEmpId, 3) & _
                                   DBW("rsttxt=", sReason) & _
                      " WHERE  " & DBW("workarea = ", pWorkArea) & _
                      " AND    " & DBW("accdt = ", pAccDt) & _
                      " AND    " & DBW("accseq = ", pAccSeq) & _
                      " AND    " & DBW("seq = ", "1")

End Function


Public Function SqlUpdateOrdH(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As String, _
                              ByVal pFlag As Long) As String

    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim Rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(pPtId) & _
             "    and orddt = " & DBS(pOrdDt) & _
             "    and ordno = " & DBS(pOrdNo)
             
    Rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        strWrkDiv = "" & Rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    Rs.Close: Set Rs = Nothing
    
    '처방Header Update
    If pFlag = 1 Then
        '** 원본 ==========================================================================
'        SqlUpdateOrdH = " update " & T_LAB101 & _
'                        " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Order) & _
'                        " WHERE  " & DBW("ptid   = ", pPtid)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        If strWrkDiv = "3" Then         '종건
            SqlUpdateOrdH = " update su2examt " & _
                            " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Order) & _
                            " WHERE  patno = " & DBS(pPtId)
        ElseIf strWrkDiv = "4" Then     '일건
                SqlUpdateOrdH = " update sg2examt " & _
                                " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Order) & _
                                " WHERE  patno = " & DBS(pPtId)
        Else
            SqlUpdateOrdH = " update mdexmort " & _
                            " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Order) & _
                            " WHERE  patno = " & DBS(pPtId)
        End If
        '===================================================================================
    ElseIf pFlag = 3 Then
        '** 원본 ==========================================================================
'        SqlUpdateOrdH = " update " & T_LAB101 & _
'                        " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Order) & _
'                        " WHERE  " & DBW("ptid   = ", pPtid)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        SqlUpdateOrdH = " update mdexmort " & _
                        " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Order) & _
                        " WHERE  patno = " & DBS(pPtId)
        '===================================================================================
    Else
        '** 원본 ==========================================================================
'        SqlUpdateOrdH = " update " & T_LAB101 & _
'                        " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Collection) & _
'                        " WHERE  " & DBW("ptid   = ", pPtid)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        If strWrkDiv = "3" Then         '종건
            SqlUpdateOrdH = " update su2examt " & _
                            " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Collection) & _
                            " WHERE  patno = " & DBS(pPtId)
        ElseIf strWrkDiv = "4" Then     '일건
                SqlUpdateOrdH = " update sg2examt " & _
                                " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Collection) & _
                                " WHERE  patno = " & DBS(pPtId)
        Else
            SqlUpdateOrdH = " update mdexmort " & _
                            " set    " & DBW("donefg = ", enStsCd.StsCd_LIS_Collection) & _
                            " WHERE  patno = " & DBS(pPtId)
        End If
        '===================================================================================
    End If
    SqlUpdateOrdH = SqlUpdateOrdH & _
                    " AND    orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                    " AND    ordseqno = " & DBN(pOrdNo)

End Function

Public Function SqlUpdateOrdB(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As String, _
                              ByVal pOrdSeq As String, ByVal pFlag As Long) As String
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim Rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(pPtId) & _
             "    and orddt = " & DBS(pOrdDt) & _
             "    and ordno = " & DBS(pOrdNo)
             
    Rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        strWrkDiv = "" & Rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    Rs.Close: Set Rs = Nothing
    
    '처방Body Update
    If pFlag = 1 Then
        '** 원본 ==========================================================================
'        SqlUpdateOrdB = " update " & T_LAB102 & _
'                        " set   stscd ='0', donefg='0', rcvdt='',   rcvtm='', " & _
'                        "       examdt='',  examtm='',  examdoct=0, workarea='', accdt='', accseq=0 " & _
'                        " WHERE " & DBW("ptid = ", pPtid)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        If strWrkDiv = "3" Then         '종건
            SqlUpdateOrdB = " update su2examt " & _
                            " set   stscd ='0', donefg='0', acptdate ='', " & _
                            "       examdate = '', examdr = '', workarea='', accdt='', accseq=0, examstat = '', " & _
                            "       colldate = '', readdate = '', readdr = '', reptdate = ''," & _
                            "       result = '', rsltcomm = '', rsltsumm = '', detail = ''" & _
                            " WHERE patno = " & DBS(pPtId)
        ElseIf strWrkDiv = "4" Then     '일건
                SqlUpdateOrdB = " update sg2examt " & _
                            " set   stscd ='0', donefg='0', acptdate ='', " & _
                            "       examdate = '',  examdr = '', workarea='', accdt='', accseq=0, examstat = '', " & _
                            "       colldate = '', readdate = '', readdr = '', reptdate = ''," & _
                            "       result = '', rsltcomm = '', rsltsumm = '', detail = ''" & _
                            " WHERE patno = " & DBS(pPtId)
        Else
            '-- 2014.07.29 osw
            SqlUpdateOrdB = " update mdexmort " & _
                            " set   stscd ='0', donefg='0', acptdate ='', " & _
                            "       execdate = '',  cofmdr = '', workarea='', accdt='', accseq=0, procstat = '', " & _
                            "       colldate = '', readdate = '', rsltdate = '' " & _
                            " WHERE patno = " & DBS(pPtId)
            
            'SqlUpdateOrdB = " update mdexmort " & _
                            " set   stscd ='0', donefg='0', acptdate ='', " & _
                            "       examdate = '',  cofmdr = '', workarea='', accdt='', accseq=0, procstat = '', " & _
                            "       colldate = '', readdate = '', rsltdate = '' " & _
                            " WHERE patno = " & DBS(pPtId)
        
        End If
        '===================================================================================
        
    ElseIf pFlag = 3 Then
        '** 원본 ==========================================================================
'        SqlUpdateOrdB = " update " & T_LAB102 & _
'                        " set   dcfg='1',stscd ='0',donefg='0',rcvdt='',   rcvtm='', " & _
'                        "       examdt='',  examtm='',  examdoct=0, workarea='', accdt='', accseq=0 " & _
'                        " WHERE " & DBW("ptid = ", pPtid)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        SqlUpdateOrdB = " update mdexmort " & _
                        " set   dcfg='1',stscd ='0',donefg='0', acptdate='', " & _
                        "       execdate = '',  cofmdr = 0, workarea='', accdt='', accseq=0 " & _
                        " WHERE patno = " & DBS(pPtId)
        '===================================================================================
    Else
        '** 원본 ==========================================================================
'        SqlUpdateOrdB = " update " & T_LAB102 & _
'                        " set   stscd ='1', donefg='1', rcvdt='',   rcvtm='' " & _
'                        " WHERE " & DBW("ptid =", pPtid)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        If strWrkDiv = "3" Then         '종건
            SqlUpdateOrdB = " update su2examt " & _
                            " set    stscd ='1', donefg='1', acptdate='', examstat = 'B', " & _
                            "        examdate = '', readdate = '', examdr = '', readdr = '', " & _
                            "        reptdate = '', result = '', rsltcomm = '', rsltsumm = '', detail = '' " & _
                            " WHERE patno = " & DBS(pPtId)
        ElseIf strWrkDiv = "4" Then     '일건
                SqlUpdateOrdB = " update sg2examt " & _
                                " set    stscd ='1', donefg='1', acptdate='', examstat = 'B', " & _
                                "        examdate = '', readdate = '', examdr = '', readdr = '', " & _
                                "        reptdate = '', result = '', rsltcomm = '', rsltsumm = '', detail = '' " & _
                                " WHERE patno = " & DBS(pPtId)
        Else
            SqlUpdateOrdB = " update mdexmort " & _
                            " set    stscd ='1', donefg='1', acptdate='', procstat = 'B', " & _
                            "        execdate = '', readdate = '', rsltdate = '', cofmdr = '' " & _
                            " WHERE patno = " & DBS(pPtId)
        End If
        '===================================================================================
    End If
    SqlUpdateOrdB = SqlUpdateOrdB & _
                    " AND    orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & _
                    " AND    ordseqno = " & DBN(pOrdNo)

End Function

Public Function OCSResultDelete(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As String, _
                                Optional ByVal pOrdSeq As String = "1", Optional ByVal pFlag As Long = 0) As String
    
    OCSResultDelete = " delete from mdresult " & _
                     "  where patno = " & DBS(pPtId) & _
                     "    and orddate = TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd') " & _
                     "    and ordseqno = " & DBN(pOrdNo)
    
End Function

Public Function SqlRackNoState(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
            
    SqlRackNoState = " SELECT * FROM s2lab206 " & _
                      "  WHERE workarea = '" & pWorkArea & "'" & _
                      "    AND accdt    = '" & pAccDt & "'" & _
                      "    AND accseq   = '" & pAccSeq & "'"
                      
End Function

Public Function SqlRackNoInsert(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pRcvDt As String, ByVal pRcvTm As String, ByVal pRackNo As String) As String
            
'    SqlRackNoInsert = " insert into s2lab206 " & _
'                      " values(" & DBV("workarea", pWorkArea, 1) & _
'                                   DBV("accdt", pAccDt, 1) & _
'                                   DBV("accseq", pAccSeq, 1) & _
'                                   DBV("rcvdt", pRcvDt, 1) & _
'                                   DBV("rcvtm", pRcvTm, 1) & _
'                                   DBV("rackno", pRackNo)
    
    SqlRackNoInsert = " insert into s2lab206 " & _
                      " values('" & pWorkArea & "', '" & pAccDt & "',  '" & pAccSeq & "', " & _
                      "'" & pRcvDt & "', '" & pRcvTm & "', '" & pRackNo & "')"
                      
End Function

Public Function SqlRackNoUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal pRackNo As String) As String
            
'    SqlRackNoInsert = " insert into s2lab206 " & _
'                      " values(" & DBV("workarea", pWorkArea, 1) & _
'                                   DBV("accdt", pAccDt, 1) & _
'                                   DBV("accseq", pAccSeq, 1) & _
'                                   DBV("rcvdt", pRcvDt, 1) & _
'                                   DBV("rcvtm", pRcvTm, 1) & _
'                                   DBV("rackno", pRackNo)
    
    SqlRackNoUpdate = " update s2lab206 " & _
                      "    set rackno = '" & pRackNo & "'" & _
                      "  WHERE workarea = '" & pWorkArea & "'" & _
                      "    AND accdt    = '" & pAccDt & "'" & _
                      "    AND accseq   = '" & pAccSeq & "'"
                      
End Function

Public Function SqlRackNoMax(ByVal pRcvDt As String, ByVal pRcvTm As String) As String
            
    SqlRackNoMax = " SELECT max(rackno) as cnt FROM s2lab206 " & _
                   "  WHERE " & DBW("rcvdt = ", pRcvDt)
End Function
'-----------------------------------------------------------------------------'
'   기능 : 연속검사 결과내역 삭제쿼리 - 5.0.34: 이상대(2005-09-01)
'   인수 :
'       - pPtId     : 환자ID
'       - pOrdDt    : 처방일
'       - pOrdNo    : 처방번호
'       - pOrdSeq   : 처방순번
'-----------------------------------------------------------------------------'
Public Function SqlDelMultiResult(ByVal pPtId As String, ByVal pOrdDt As String, _
                                  ByVal pOrdNo As String, ByVal pOrdSeq As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & T_LAB302 & _
          " WHERE " & DBW("ptid=", pPtId) & _
          "     AND " & DBW("orddt=", pOrdDt) & _
          "     AND " & DBW("ordno=", pOrdNo) & _
          "     AND " & DBW("ordseq=", pOrdSeq)

    SqlDelMultiResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 연속검사내역 삭제쿼리 - 5.0.36: 이상대(2005-09-07)
'   인수 :
'       - pPtId     : 환자ID
'       - pOrdDt    : 처방일
'       - pOrdNo    : 처방번호
'       - pOrdSeq   : 처방순번
'-----------------------------------------------------------------------------'
Public Function SqlDelMultiHistory(ByVal pPtId As String, ByVal pOrdDt As String, _
                                  ByVal pOrdNo As String, ByVal pOrdSeq As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & T_LAB203 & _
          " WHERE " & DBW("ptid=", pPtId) & _
          "     AND " & DBW("orddt=", pOrdDt) & _
          "     AND " & DBW("ordno=", pOrdNo) & _
          "     AND " & DBW("ordseq=", pOrdSeq)

    SqlDelMultiHistory = SQL
End Function


Public Function SqlPtBloodType(ByVal pPtId As String) As String
    Dim SQL As String
    
    SQL = " Select  c.ABO || c.RH as Blood , max(b.vfydt || b.vfytm) as vfydtm From " & T_BBS202 & " a, " & T_BBS302 & " b, " & T_BBS401 & " c " & _
          " WHERE " & DBW("a.ptid=", pPtId) & _
          "   AND b.workarea = a.workarea and b.accdt = a.accdt and b.accseq = a.accseq " & _
          "   AND c.bldsrc = b.bldsrc and c.bldyy = b.bldyy and c.bldno = b.bldno and c.compocd = b.compocd" & _
          " group by  c.ABO , c.RH"

    SqlPtBloodType = SQL
End Function


