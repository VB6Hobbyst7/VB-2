VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarWorkarea As String

Public Property Let WorkArea(ByVal vData As String)
    mvarWorkarea = vData
End Property

Public Function GetBuildCd() As String
'FROM LAB032 공통코드 마스터

    GetBuildCd = "SELECT cdval1 as BuildCd, field1 as BuildNm FROM " & T_LAB032 & _
                 " WHERE " & _
                    DBW("cdindex=", LC3_Buildings) & _
                 " order by BuildCd "
End Function

Public Function GetWACd() As String
'FROM LAB032 공통코드 마스터

    GetWACd = " SELECT cdval1 as WACd , field1 as WANm " & _
                " FROM " & T_LAB032 & _
                " WHERE " & DBW("cdindex=", lc3_workarea) & " order by WACd"
End Function

Public Function GetSpcList() As String
'FROM LAB032 공통코드 마스터

    GetSpcList = " SELECT cdval1 as spccd , field4 as spcnm " & _
                " FROM " & T_LAB032 & _
                " WHERE " & DBW("cdindex=", LC3_Specimen) & " order by spcnm"
End Function

Public Function GetSpcListByTest(ByVal sTestCd As String) As String
'FROM LAB004 지정검체 마스터

    GetSpcListByTest = " SELECT a.spccd , b.field4 as spcnm " & _
                        " FROM  " & T_LAB032 & " b, " & T_LAB004 & " a " & _
                        " WHERE " & DBW("a.testcd=", sTestCd) & _
                        " AND   " & DBW("b.cdindex=", LC3_Specimen) & _
                        " AND   b.cdval1 = a.spccd " & _
                        " order by spcnm"
End Function

Public Function GetEqpCd(Optional ByVal pOption As Boolean = True, _
                         Optional ByVal pBldCd As String = "") As String
'FROM LAB006
'pOption True : ListIndex=0 False : ListIndex > 0

    If pOption Then
        GetEqpCd = " SELECT  eqpcd, eqpnm FROM " & T_LAB006
    Else
        GetEqpCd = " SELECT  eqpcd, eqpnm FROM " & T_LAB006 & _
                            " WHERE " & DBW("location=", pBldCd)
    End If
End Function

Public Function SqlStaticItem() As String
    
    SqlStaticItem = " SELECT a.cdval1 as testcd, " & FUNC_CONVERT("num", "a.field1") & " as seq, b.testnm, b.workarea, a.field2 as wanm " & _
                    " FROM  " & T_LAB032 & " a, " & T_LAB001 & " b " & _
                    " WHERE " & DBW("a.cdindex", LC3_StaticItem, 2) & _
                    " AND   b.testcd = a.cdval1 " & _
                    " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                    "                    WHERE  testcd=b.testcd) " & _
                    " ORDER BY seq"

End Function


Public Function SqlTestCount(ByVal intTable As Integer, ByVal strColDt As String, _
                             ByVal strFrDt As String, ByVal strToDt As String, ByVal BldCd As String) As String
    
    Dim strTblNm As String
    Dim strtmp   As String
    
    '"AND   c.stscd > '" & StsCd_LIS_Accession & "' "    '검사수행된것만...
    '** 임상병리과(CL01)에서 낸 처방은 제외..

    'strTblNm = Choose(intTable, T_LAB302, T_LAB351, T_LAB404)
    
'    If dbconn.Whatsthis = dbconn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then   'sybase
        strtmp = T_LAB201 & " c ( index h7lab201_index2), "
'    ElseIf dbconn.Whatsthis = dbconn.ThisIsOracle Then
    ElseIf ObjSysInfo.dbtype = 0 Then 'Oracle
        strtmp = T_LAB201 & " c , "
    Else
        strtmp = T_LAB201 & " c , "
    End If
    
    
'    SqlTestCount = "SELECT  b.groupcd, a.field1 as WAnm, b.itemseq,b.testcd, b.testnm, c.rcvdt, count(d.testcd) as Cnt " & _
                "FROM " & T_LAB032 & " a, " & T_LAB001 & " b, " & strtmp & strTblNm & " d " & _
                "WHERE c.rcvdt>='" & strFrDt & "' " & _
                "AND   c.rcvdt<='" & strToDt & "' " & _
                "AND   c.buildcd = '" & BldCd & "' " & _
                "AND   d.workarea=c.workarea " & _
                "AND   d.accdt=c.accdt " & _
                "AND   d.accseq=c.accseq " & _
                "AND   b.testcd=d.testcd " & _
                "AND   (b.detailfg='' or b.detailfg is null) " & _
                "AND   b.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                "WHERE testcd=d.testcd) " & _
                "AND   a.cdindex='" & LC3_StaticGroup & "' " & _
                "AND   " & DBJ("a.cdval1=*b.groupcd") & _
                " group by b.groupcd, a.field1, b.itemseq, b.testcd, b.testnm, c.rcvdt " & _
                "having count(*) > 0 " & _
                "order by b.groupcd, b.itemseq, b.testcd, c.rcvdt"
    
    ' 20017.05.16 속도문제로 rcvdt => accdt 로 변경
    
    SqlTestCount = "SELECT  b.groupcd, a.field1 as WAnm, b.itemseq,b.testcd, b.testnm, c.rcvdt, count(d.testcd) as Cnt " & _
                "FROM " & T_LAB032 & " a, " & T_LAB001 & " b, " & strtmp & T_LAB302 & " d " & _
                "WHERE c.accdt>='" & strFrDt & "' " & _
                "AND   c.accdt<='" & strToDt & "' " & _
                "AND   c.buildcd = '" & BldCd & "' " & _
                "AND   d.workarea=c.workarea " & _
                "AND   d.accdt=c.accdt " & _
                "AND   d.accseq=c.accseq " & _
                "AND   b.testcd=d.testcd " & _
                "AND   (b.detailfg='' or b.detailfg is null) " & _
                "AND   b.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                "WHERE testcd=d.testcd) " & _
                "AND   a.cdindex='" & LC3_StaticGroup & "' " & _
                "AND   " & DBJ("a.cdval1=*b.groupcd") & _
                " group by b.groupcd, a.field1, b.itemseq, b.testcd, b.testnm, c.rcvdt " & _
                "having count(*) > 0 "
    SqlTestCount = SqlTestCount & " UNION ALL "
    
    SqlTestCount = SqlTestCount & "SELECT  b.groupcd, a.field1 as WAnm, b.itemseq,b.testcd, b.testnm, c.rcvdt, count(d.testcd) as Cnt " & _
                "FROM " & T_LAB032 & " a, " & T_LAB001 & " b, " & strtmp & T_LAB351 & " d " & _
                "WHERE c.accdt>='" & strFrDt & "' " & _
                "AND   c.accdt<='" & strToDt & "' " & _
                "AND   c.buildcd = '" & BldCd & "' " & _
                "AND   d.workarea=c.workarea " & _
                "AND   d.accdt=c.accdt " & _
                "AND   d.accseq=c.accseq " & _
                "AND   b.testcd=d.testcd " & _
                "AND   (b.detailfg='' or b.detailfg is null) " & _
                "AND   b.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                "WHERE testcd=d.testcd) " & _
                "AND   a.cdindex='" & LC3_StaticGroup & "' " & _
                "AND   " & DBJ("a.cdval1=*b.groupcd") & _
                " group by b.groupcd, a.field1, b.itemseq, b.testcd, b.testnm, c.rcvdt " & _
                "having count(*) > 0 "
    
    SqlTestCount = SqlTestCount & " UNION ALL "
    
    SqlTestCount = SqlTestCount & "SELECT  b.groupcd, a.field1 as WAnm, b.itemseq,b.testcd, b.testnm, c.rcvdt, count(d.testcd) as Cnt " & _
                "FROM " & T_LAB032 & " a, " & T_LAB001 & " b, " & strtmp & T_LAB404 & " d " & _
                "WHERE c.rcvdt>='" & strFrDt & "' " & _
                "AND   c.rcvdt<='" & strToDt & "' " & _
                "AND   c.buildcd = '" & BldCd & "' " & _
                "AND   d.workarea=c.workarea " & _
                "AND   d.accdt=c.accdt " & _
                "AND   d.accseq=c.accseq " & _
                "AND   b.testcd=d.testcd " & _
                "AND   (b.detailfg='' or b.detailfg is null) " & _
                "AND   b.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                "WHERE testcd=d.testcd) " & _
                "AND   a.cdindex='" & LC3_StaticGroup & "' " & _
                "AND   " & DBJ("a.cdval1=*b.groupcd") & _
                " group by b.groupcd, a.field1, b.itemseq, b.testcd, b.testnm, c.rcvdt " & _
                "having count(*) > 0 "
                
    SqlTestCount = SqlTestCount & " order by groupcd, itemseq, testcd, rcvdt"
End Function



'Public Function GetAccCnt(ByVal pStart As String, ByVal pEnd As String)
'FROM LAB201, LAB302, LAB001, HIS003
'
'    Dim SqlStmt As String
'
'    '일반검사
'    SqlStmt = "SELECT a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
'    SqlStmt = SqlStmt & "FROM  " & T_HIS003 & " c, " & T_LAB001 & " d, " & T_LAB302 & " b, " & T_LAB201 & " a "
'    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
'    '** 임상병리과(CL01)에서 낸 처방은 제외...
'    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
'    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
'    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
'    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
'    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
'    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
'                                          "WHERE  testcd = d.testcd) "
'    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
'    SqlStmt = SqlStmt & " GROUP BY a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & ", b.testcd, d.testnm "
'    '기타검사
'    SqlStmt = SqlStmt & "UNION ALL SELECT a.buildcd, a.workarea, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
'    SqlStmt = SqlStmt & "FROM  " & T_HIS003 & " c, " & T_LAB001 & " d, " & T_LAB351 & " b, " & T_LAB201 & " a "
'    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
'    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
'    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
'    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
'    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
'    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
'                                          "WHERE  testcd = d.testcd) "
'    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
'    SqlStmt = SqlStmt & " GROUP BY a.buildcd, a.workarea, c." & F_DEPTNM & ", b.testcd, d.testnm "
'    '미생물검사
'    SqlStmt = SqlStmt & "UNION ALL SELECT a.buildcd, a.workarea, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
'    SqlStmt = SqlStmt & "FROM  " & T_HIS003 & " c, " & T_LAB001 & " d, " & T_LAB404 & " b, " & T_LAB201 & " a "
'    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
'    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
'    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
'    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
'    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
'    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
'                                          "WHERE  testcd = d.testcd) "
'    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
'    SqlStmt = SqlStmt & " GROUP BY a.buildcd, a.workarea, c." & F_DEPTNM & ", b.testcd, d.testnm "
'    SqlStmt = SqlStmt & " ORDER BY buildcd, workarea, eqpcd, DeptNm, testcd, testnm "
'
'    GetAccCnt = SqlStmt
    
'End Function

Public Function GetSensiResult(ByVal pVfyDt As String, ByVal pMic As String) As String
'FROM LAB405 , LAB201
    
    GetSensiResult = " SELECTb.* " & _
                       " FROM " & T_LAB404 & " a," & T_LAB405 & " b" & _
                      " WHERE " & DBW("a.vfydt=", pVfyDt) & _
                    " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                    " AND " & DBW("a.senfg=", "Y") & _
                    " AND b.workarea=a.workarea" & " AND b.accdt=a.accdt" & " AND b.accseq=a.accseq" & " AND b.testcd=a.testcd" & " AND b.mfyseq=a.mfyseq "
    If pMic <> "" Then GetSensiResult = GetSensiResult & " AND b.mnmcd in (" & pMic & ")"
    
End Function

Public Function GetSensiData(ByVal pOption As Boolean, ByVal pVfyDt As String, ByVal pMic As String, ByVal pWardId As String) As String
'FROM LAB405 , LAB201
    Dim sSQL    As String
    Dim sQuery  As String
    
    If pOption Then '병동
        sQuery = IIf(pWardId <> "", " AND " & DBW("wardid=", Trim(Mid(pWardId, 1, 5))) & ")", ")")
        
        sSQL = sSQL & _
                    " AND exists (SELECT * FROM " & T_LAB201 & _
                                " WHERE workarea = a.workarea AND accdt = a.accdt AND accseq = a.accseq"
        sSQL = sSQL & sQuery
    Else            '진료과
        sQuery = IIf(pWardId <> "", " AND " & DBW("deptcd=", Trim(Mid(pWardId, 1, 5))) & ")", ")")
    
        sSQL = sSQL & " AND exists (SELECT * FROM " & T_LAB201 & " WHERE workarea = a.workarea AND accdt = a.accdt AND accseq = a.accseq"
        sSQL = sSQL & sQuery
                                
    End If
    
    GetSensiData = sSQL
End Function


Public Function GetSensiResult_New(ByVal pVfyDtF As String, ByVal pVfyDtT As String, ByVal pMic As String, ByVal sOptDept As String, ByVal pWardId As String) As String
'FROM LAB405 , LAB201
    Dim sSQL    As String
    Dim sQuery  As String

    sSQL = " SELECTb.* " & _
                       " FROM " & T_LAB404 & " a," & T_LAB405 & " b" & _
                      " WHERE " & DBW("a.vfydt>=", pVfyDtF) & " AND " & DBW("a.vfydt<=", pVfyDtT) & _
                    " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                    " AND " & DBW("a.senfg=", "Y") & _
                    " AND b.workarea=a.workarea" & " AND b.accdt=a.accdt" & " AND b.accseq=a.accseq" & " AND b.testcd=a.testcd" & " AND b.mfyseq=a.mfyseq "
    
    If pMic <> "" Then sSQL = sSQL & " AND b.mnmcd in (" & pMic & ")"
    
    Select Case sOptDept
        Case "1"    '병동
            sQuery = IIf(pWardId <> "", " AND " & DBW("wardid=", Trim(Mid(pWardId, 1, 5))) & ")", ")")
            sSQL = sSQL & " AND exists (SELECT * FROM " & T_LAB201 & " WHERE workarea = a.workarea AND accdt = a.accdt AND accseq = a.accseq"
        Case "2"    '진료과
            sQuery = IIf(pWardId <> "", " AND " & DBW("deptcd=", Trim(Mid(pWardId, 1, 5))) & ")", ")")
            sSQL = sSQL & " AND exists (SELECT * FROM " & T_LAB201 & " WHERE workarea = a.workarea AND accdt = a.accdt AND accseq = a.accseq"
    End Select

    sSQL = sSQL & sQuery
    GetSensiResult_New = sSQL
End Function



Public Function GetSpecies() As String
'FROM LAB032    균종

    GetSpecies = " SELECT cdval1 mkind FROM " & T_LAB032 & _
                 " WHERE " & DBW("cdindex=", LC3_Species) & " " & _
                 " ORDER BY cdval1 asc"
End Function

Public Function GetMicrobe(ByVal pOption As Boolean, Optional pMKind As String) As String
'FROM LAB032 균
'pOption

    Dim sqlMic As String
    
'    If dbconn.Whatsthis = dbconn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then 'sybase
        If pOption Then
            sqlMic = " SELECT cdval1 mic, text1 micnm FROM " & T_LAB032 & _
                     " WHERE " & DBW("cdindex=", LC3_Microbe) & " "
        Else
            sqlMic = " SELECT cdval1 mic, text1 micnm FROM " & T_LAB032 & _
                     " WHERE " & DBW("cdindex=", LC3_Microbe) & _
                     " AND " & DBW("field2=", pMKind) & " "
        End If
    Else
        If pOption Then
            sqlMic = " SELECT cdval1 mic, text1 micnm FROM " & T_LAB032 & _
                     " WHERE " & DBW("cdindex=", LC3_Microbe) & " " & _
                     " ORDER BY micnm asc"
        Else
            sqlMic = " SELECT cdval1 mic, text1 micnm FROM " & T_LAB032 & _
                     " WHERE " & DBW("cdindex=", LC3_Microbe) & _
                     " AND " & DBW("field2=", pMKind) & " " & _
                     " ORDER BY micnm asc"
        End If
    End If
    
    GetMicrobe = sqlMic
End Function

Public Function GetAntiBiotic(Optional ByVal pSCd As String = "") As String
'FROM LAB032  항생제
    If pSCd = CS_AllCaption Then
        GetAntiBiotic = " SELECT cdval1 anti, text1 antinm FROM " & T_LAB032 & _
                        " WHERE " & DBW("cdindex = ", LC3_AntiBiotic) & _
                        " ORDER BY cdval1 asc"
    Else
        GetAntiBiotic = " SELECT cdval2 anti FROM " & T_LAB031 & _
                        " WHERE " & DBW("cdindex = ", LC2_MicroAnti) & _
                        " AND   " & DBW("cdval1  = ", pSCd) & _
                        " ORDER BY cdval2 asc"
    End If
End Function

Public Function GetCntofWard(ByVal pAccYr As String, ByVal pAccDt As String) As String
'FROM LAB201

    Dim SqlStmt As String
    
    SqlStmt = "SELECT a.wardid as Wardid, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM   " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE a.workarea = '06' AND  " & DBW("a.accdt =", pAccYr & Mid(pAccDt, 4, 2))
    SqlStmt = SqlStmt & "AND a.wardid <> '' AND a.accseq in ( SELECT accseq FROM " & T_LAB402 & "  "
    SqlStmt = SqlStmt & "                                      WHERE wscd = 'BL' AND workarea = '06' "
    SqlStmt = SqlStmt & "                                        AND  " & DBW("accdt =", pAccYr & Mid(pAccDt, 4, 2))
    SqlStmt = SqlStmt & "                                        AND (scfg = 'S' or scfg = 'C')  ) "
    SqlStmt = SqlStmt & "Group by a.wardid  Order by a.wardid"
    
    GetCntofWard = SqlStmt
End Function

Public Function GetAbnormalLst(ByVal sStartDt As String, ByVal SendDt As String, _
                               ByVal sCondQuery As String) As String
    '## 5.0.12: 이상대(2004-12-29)
    '   - 보고자, 보고일시 조회할수 있도록 쿼리수정
    '## 5.0.25: 이상대(2005-05-31)
    '   - 검사항목별 결과코드에 등록된 결과명을 조회하도록 쿼리수정 /*+ RULE */ /*+ S2LAB201_IDX7 */
'    GetAbnormalLst = " SELECT /*+ S2LAB201_IDX7 */ a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
'                  "        a.hosilid as roomid , a.bedid as bedid, b.testcd as testcd, b.rstcd as rstcd, h.field1 as rstnm, b.lastrst, " & _
'                  "        b.hldiv as hldiv , b.dpdiv as dpdiv , g.empnm, b.vfydt, b.vfytm, c.rsttxt as rsttxt, " & _
'                  "        d." & F_PTNM & " as ptnm, e.testnm as testnm, f.field1 as BuildNm " & _
'                  " FROM " & T_LAB302 & " b, " & T_LAB304 & " c, " & T_HIS001 & " d, " & _
'                        T_LAB001 & " e, " & T_LAB032 & " f, " & T_COM006 & " g, " & T_LAB031 & " h, " & T_LAB201 & " a " & _
'                  " WHERE " & _
'                            DBW("a.accdt >= ", sStartDt) & _
'                  " AND " & DBW("a.accdt <= ", SendDt) & _
'                  " AND   b.workarea = a.workarea" & _
'                  " AND   b.accdt    = a.accdt" & _
'                  " AND   b.accseq   = a.accseq" & _
'                  " AND " & DBJ("c.workarea =* a.workarea") & _
'                  " AND " & DBJ("c.accdt    =* a.accdt") & _
'                  " AND " & DBJ("c.accseq   =* a.accseq") & _
'                  " AND   e.testcd = b.testcd" & _
'                  " AND   e.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = e.testcd)" & _
'                  " AND   a.ptid = d." & F_PTID & _
'                  " AND   b.vfydt <> ' ' " & _
'                  " AND   b.vfydt is not null " & _
'                  " AND   ( " & sCondQuery & " )" & _
'                  " AND   " & DBJ(DBW("f.cdindex = ", LC3_Buildings)) & _
'                  " AND   " & DBJ("f.cdval1  =* a.buildcd") & _
'                  " AND   b.vfyid=g.empid"
                  
    GetAbnormalLst = " SELECT /*+ S2LAB201_IDX7 */ a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
                  "        a.hosilid as roomid , a.bedid as bedid, b.testcd as testcd, b.rstcd as rstcd, h.field1 as rstnm, b.lastrst, " & _
                  "        b.hldiv as hldiv , b.dpdiv as dpdiv , g.empnm, b.vfydt, b.vfytm, c.rsttxt as rsttxt, " & _
                  "        d." & F_PTNM & " as ptnm, e.testnm as testnm, f.field1 as BuildNm " & _
                  " FROM " & T_LAB302 & " b, " & T_LAB304 & " c, " & T_HIS001 & " d, " & _
                        T_LAB001 & " e, " & T_LAB032 & " f, " & T_COM006 & " g, " & T_LAB031 & " h, " & T_LAB201 & " a " & _
                  " WHERE " & _
                            DBW("a.accdt >= ", sStartDt) & _
                  " AND " & DBW("a.accdt <= ", SendDt) & _
                  " AND   b.workarea = a.workarea" & _
                  " AND   b.accdt    = a.accdt" & _
                  " AND   b.accseq   = a.accseq" & _
                  " AND " & DBJ("c.workarea =* a.workarea") & _
                  " AND " & DBJ("c.accdt    =* a.accdt") & _
                  " AND " & DBJ("c.accseq   =* a.accseq") & _
                  " AND   e.testcd = b.testcd" & _
                  " AND   a.ptid = d." & F_PTID & _
                  " AND   b.vfydt <> ' ' " & _
                  " AND   b.vfydt is not null " & _
                  " AND   ( " & sCondQuery & " )" & _
                  " AND   " & DBJ(DBW("f.cdindex = ", LC3_Buildings)) & _
                  " AND   " & DBJ("f.cdval1  =* a.buildcd") & _
                  " AND   b.vfyid=g.empid"
                  
    GetAbnormalLst = GetAbnormalLst & " AND   " & DBJ(DBW("h.cdindex=", LC2_ItemResult)) & _
                  " AND   " & DBJ("h.cdval1=*b.testcd") & _
                  " AND   " & DBJ("h.cdval2=*b.rstcd")
End Function

'-----------------------------------------------------------------------------'
'   기능 : Workarea별로 AbNormal List 조회 - 5.0.12:이상대(2004-12-29)
'   인수 :
'       - pStartDt   : 검색 From Date
'       - pEndDt     : 검색 To Date
'       - pCondQuery : 검색조건
'       - pWorkarea  : Workarea
'-----------------------------------------------------------------------------'
Public Function GetAbnormalLstX(ByVal pStartDt As String, ByVal pEndDt As String, _
                               ByVal pCondQuery As String, Optional ByVal pWorkArea As String = "") As String
    Dim SQL As String
    
    '## 5.0.25: 이상대(2005-05-31)
    '   - 검사항목별 결과코드에 등록된 결과명을 조회하도록 쿼리수정 /*+ RULE */
'    SQL = " SELECT  a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
'            "        a.hosilid as roomid , a.bedid as bedid, b.testcd as testcd, b.rstcd as rstcd, h.field1 as rstnm, b.lastrst, " & _
'            "        b.hldiv as hldiv , b.dpdiv as dpdiv , g.empnm, b.vfydt, b.vfytm, c.rsttxt as rsttxt, " & _
'            "        d." & F_PTNM & " as ptnm, e.testnm as testnm, f.field1 as BuildNm " & _
'            " FROM " & T_LAB302 & " b, " & T_LAB304 & " c, " & T_HIS001 & " d, " & T_LAB001 & " e, " & _
'                    T_LAB032 & " f, " & T_COM006 & " g, " & T_LAB031 & " h, " & T_LAB201 & " a " & _
'            " WHERE " & DBW("a.workarea=", pWorkArea) & _
'            " AND " & DBW("a.accdt >= ", pStartDt) & _
'            " AND " & DBW("a.accdt <= ", pEndDt) & _
'            " AND   b.workarea = a.workarea" & _
'            " AND   b.accdt    = a.accdt" & _
'            " AND   b.accseq   = a.accseq" & _
'            " AND " & DBJ("c.workarea =* a.workarea") & _
'            " AND " & DBJ("c.accdt    =* a.accdt") & _
'            " AND " & DBJ("c.accseq   =* a.accseq") & _
'            " AND   e.testcd = b.testcd" & _
'            " AND   e.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = e.testcd)" & _
'            " AND   a.ptid = d." & F_PTID & _
'            " AND   b.vfydt <> ' ' " & _
'            " AND   b.vfydt is not null " & _
'            " AND  ( " & pCondQuery & " )" & _
'            " AND " & DBJ(DBW("f.cdindex = ", LC3_Buildings)) & _
'            " AND " & DBJ("f.cdval1  =* a.buildcd") & _
'            " AND b.vfyid=g.empid"
            
    SQL = " SELECT  a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
            "        a.hosilid as roomid , a.bedid as bedid, b.testcd as testcd, b.rstcd as rstcd, h.field1 as rstnm, b.lastrst, " & _
            "        b.hldiv as hldiv , b.dpdiv as dpdiv , g.empnm, b.vfydt, b.vfytm, c.rsttxt as rsttxt, " & _
            "        d." & F_PTNM & " as ptnm, e.testnm as testnm, f.field1 as BuildNm " & _
            " FROM " & T_LAB302 & " b, " & T_LAB304 & " c, " & T_HIS001 & " d, " & T_LAB001 & " e, " & _
                    T_LAB032 & " f, " & T_COM006 & " g, " & T_LAB031 & " h, " & T_LAB201 & " a " & _
            " WHERE " & DBW("a.workarea=", pWorkArea) & _
            " AND " & DBW("a.accdt >= ", pStartDt) & _
            " AND " & DBW("a.accdt <= ", pEndDt) & _
            " AND   b.workarea = a.workarea" & _
            " AND   b.accdt    = a.accdt" & _
            " AND   b.accseq   = a.accseq" & _
            " AND " & DBJ("c.workarea =* a.workarea") & _
            " AND " & DBJ("c.accdt    =* a.accdt") & _
            " AND " & DBJ("c.accseq   =* a.accseq") & _
            " AND   e.testcd = b.testcd" & _
            " AND   a.ptid = d." & F_PTID & _
            " AND   b.vfydt <> ' ' " & _
            " AND   b.vfydt is not null " & _
            " AND  ( " & pCondQuery & " )" & _
            " AND " & DBJ(DBW("f.cdindex = ", LC3_Buildings)) & _
            " AND " & DBJ("f.cdval1  =* a.buildcd") & _
            " AND b.vfyid=g.empid"
            
    SQL = SQL & " AND   " & DBJ(DBW("h.cdindex=", LC2_ItemResult)) & _
                " AND   " & DBJ("h.cdval1=*b.testcd") & _
                " AND   " & DBJ("h.cdval2=*b.rstcd")
                  
    GetAbnormalLstX = SQL
End Function

'------------------------------
'WorkArea별로 검사항목 불러오기
'------------------------------
Public Function GetWAvsTest(ByVal WorkArea As String) As String
    '## 5.0.12: 이상대(2004-12-29)
    '   - 그룹항목을 조회하지 않도록 쿼리수정
    
    '상세대표코드
    GetWAvsTest = " SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
                  " WHERE " & _
                            DBW("workarea=", WorkArea) & _
                  " AND " & DBW("panelfg=", "D")
                  
    '단일항목코드
    GetWAvsTest = GetWAvsTest & " union  SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
                  " WHERE " & DBW("workarea=", WorkArea) & _
                  " AND panelfg is null"
                  
    '** 원본 ==============================================================
'    '상세대표코드
'    GetWAvsTest = " SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
'                  " WHERE " & _
'                            DBW("workarea=", WorkArea) & _
'                  " AND " & DBW("panelfg=", "D") & _
'                  " AND " & DBW("testdiv<>", "3")
'
'    '단일항목코드
'    GetWAvsTest = GetWAvsTest & " union  SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
'                  " WHERE " & DBW("workarea=", WorkArea) & _
'                  " AND panelfg is null" & _
'                  " AND " & DBW("testdiv<>", "3")
End Function

Public Function GetAccTest(Optional pTestCd As String = "") As String
    '상세대표코드
    GetAccTest = " SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
                  " WHERE  " & _
                             DBW("panelfg=", "D")
                             
    If pTestCd <> "" Then GetAccTest = GetAccTest & " AND " & DBW("testcd", pTestCd, 2)
                
    '단일항목코드
    GetAccTest = GetAccTest & _
                                " union  SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
                                " WHERE " & _
                                        " (detailfg is null or detailfg ='')"
    
    If pTestCd <> "" Then GetAccTest = GetAccTest & " AND " & DBW("testcd", pTestCd, 2)
End Function

Public Function GetAccTest_Wm(Optional pTestCd As String = "") As String
    '상세대표코드
    GetAccTest_Wm = " SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
                  " WHERE workarea = '14' and " & _
                             DBW("panelfg=", "D")
                             
    If pTestCd <> "" Then GetAccTest_Wm = GetAccTest_Wm & " AND " & DBW("testcd", pTestCd, 2)
                
    '단일항목코드
    GetAccTest_Wm = GetAccTest_Wm & _
                                " union  SELECT testcd ,abbrnm10, testdiv FROM " & T_LAB001 & _
                                " WHERE workarea = '14' and " & _
                                        " (detailfg is null or detailfg ='')"
    
    If pTestCd <> "" Then GetAccTest_Wm = GetAccTest_Wm & " AND " & DBW("testcd", pTestCd, 2)
End Function

'-----------------------------------------------------------------------
'검사항목별로 접수일시,처방일시,결과일시를 불러온다(결과상태 이상인것만)
'-----------------------------------------------------------------------
Public Function TurnAroundDateRecord(ByVal Fdt As String, ByVal Tdt As String, ByVal TestCd As String, _
                                     ByVal qStscd As String, ByVal qOrddiv As String) As Recordset
'    Dim sSQL As String
'
'    If WA = "07" Or WA = "08" Or WA = "09" Then
'        sSQL = " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
'               " FROM " & T_LAB404 & " b," & T_LAB201 & " a" & _
'               " WHERE " & _
'                         DBW("a.rcvdt>=", Fdt) & _
'               " AND " & DBW("a.rcvdt<=", Tdt) & _
'               " AND " & DBW("a.stscd>", qStscd) & _
'               " AND " & DBW("b.testcd=", TestCd) & _
'               " AND a.workarea=b.workarea" & _
'               " AND a.accdt=b.accdt" & _
'               " AND a.accseq=b.accseq" & _
'               " order by a.rcvdt,a.rcvtm"
'    Else
'        '임상병리
'        sSQL = " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
'               " FROM " & T_LAB302 & " b," & T_LAB201 & " a" & _
'               " WHERE " & _
'                         DBW("a.rcvdt>=", Fdt) & _
'               " AND " & DBW("a.rcvdt<=", Tdt) & _
'               " AND " & DBW("a.stscd>", qStscd) & _
'               " AND " & DBW("b.testcd=", TestCd) & _
'               " AND a.workarea=b.workarea" & _
'               " AND a.accdt=b.accdt" & _
'               " AND a.accseq=b.accseq" & _
'               " order by a.rcvdt,a.rcvtm"
'    End If
'
'    Set TurnAroundDateRecord = OpenRecordSet(sSQL)
End Function

Public Function TurnAroundNewDateRecord(ByVal Fdt As String, ByVal Tdt As String, ByVal TestCd As String, _
                                        ByVal qStscd As String, ByVal qOrddiv As String, ByVal WA As String) As Recordset
    
End Function

Public Function GetTurnAroundDateRecord(ByVal Fdt As String, ByVal Tdt As String, ByVal TestCd As String, _
                                        ByVal qStscd As String, ByVal qOrddiv As String, ByVal StatFg As String, _
                                        ByVal ErCheck As String) As Recordset
    Dim sSQL As String
    Dim strtmp As String
    
    If StatFg = "1" Then        '응급
        strtmp = " AND " & DBW("a.statfg = ", "1")
    ElseIf StatFg = "2" Then    '비응급
        strtmp = " AND " & DBW("a.statfg = ", "0")
    Else
        strtmp = ""
    End If
    
    If ErCheck = "1" Then
        strtmp = strtmp & " AND a.deptcd in (" & P_ErDeptCd & ")"
    End If
    
    sSQL = " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
           " FROM " & T_LAB302 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
                     DBW("a.rcvdt>=", Fdt) & _
           " AND " & DBW("a.rcvdt<=", Tdt) & _
           " AND " & DBW("a.stscd>", qStscd) & strtmp & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("b.testcd=", TestCd) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq"
           
    sSQL = sSQL & " UNION ALL " & _
           " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
           " FROM " & T_LAB404 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
                     DBW("a.rcvdt>=", Fdt) & _
           " AND " & DBW("a.rcvdt<=", Tdt) & _
           " AND " & DBW("a.stscd>", qStscd) & strtmp & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("b.testcd=", TestCd) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq"
    
    sSQL = sSQL & " UNION ALL " & _
           " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
           " FROM " & T_LAB351 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
                     DBW("a.rcvdt>=", Fdt) & _
           " AND " & DBW("a.rcvdt<=", Tdt) & _
           " AND " & DBW("a.stscd>", qStscd) & strtmp & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("b.testcd=", TestCd) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq" & _
           " order by rcvdt,rcvtm"
    
    Set GetTurnAroundDateRecord = New Recordset
    GetTurnAroundDateRecord.Open sSQL, DBConn
End Function

'-----------------------------------------------------------------------------'
'   기능 : GetTurnAroundDateRecord 함수에 Workarea조건을 추가 - 5.0.12: 이상대(2004-12-29)
'-----------------------------------------------------------------------------'
Public Function GetTurnAroundDateRecordX(ByVal Fdt As String, ByVal Tdt As String, ByVal TestCd As String, _
                                        ByVal qStscd As String, ByVal qOrddiv As String, ByVal StatFg As String, _
                                        ByVal ErCheck As String, ByVal pWorkArea As String) As Recordset
    Dim sSQL    As String
    Dim strtmp  As String
    
    If StatFg = "1" Then        '응급
        strtmp = " AND " & DBW("a.statfg = ", "1")
    ElseIf StatFg = "2" Then    '비응급
        strtmp = " AND " & DBW("a.statfg = ", "0")
    Else
        strtmp = ""
    End If
    
    If ErCheck = "1" Then
'''        strtmp = strtmp & " AND a.deptcd in (" & P_ErDeptCd & ")"
        strtmp = strtmp & " AND a.deptcd in ('EM','ER')"   '2014-11-05 PSK 응급일경우
    End If
    
    sSQL = " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
           " FROM " & T_LAB302 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
                     DBW("a.rcvdt>=", Fdt) & _
           " AND " & DBW("a.rcvdt<=", Tdt) & _
           " AND " & DBW("a.workarea=", pWorkArea) & _
           " AND " & DBW("a.stscd>", qStscd) & strtmp & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("b.testcd=", TestCd) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq"
           
    sSQL = sSQL & " UNION ALL " & _
           " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
           " FROM " & T_LAB404 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
                     DBW("a.rcvdt>=", Fdt) & _
           " AND " & DBW("a.rcvdt<=", Tdt) & _
           " AND " & DBW("a.workarea=", pWorkArea) & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("b.stscd>", qStscd) & strtmp & _
           " AND " & DBW("b.testcd=", TestCd) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq"
    
    sSQL = sSQL & " UNION ALL " & _
           " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,a.coldt,a.coltm " & _
           " FROM " & T_LAB351 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
                     DBW("a.rcvdt>=", Fdt) & _
           " AND " & DBW("a.rcvdt<=", Tdt) & _
           " AND " & DBW("a.workarea=", pWorkArea) & _
           " AND " & DBW("a.stscd>", qStscd) & strtmp & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("b.testcd=", TestCd) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq" & _
           " order by rcvdt,rcvtm"
    
    Debug.Print sSQL
    Set GetTurnAroundDateRecordX = New Recordset
    GetTurnAroundDateRecordX.Open sSQL, DBConn
End Function

Function GetPtTATTime(ByVal Fdt As String, ByVal Tdt As String, ByVal TestCd As String, _
                      ByVal WorkArea As String, ByVal StsCd As String, ByVal ORDDIV As String) As Recordset

End Function

Public Function GetOverTATTime(ByVal Fdt As String, ByVal Tdt As String, ByVal TestCd As String, _
                             ByVal StatFg As String, ByVal StsCd As String, _
                             ByVal ORDDIV As String, ByVal ErCheck As String) As Recordset
    Dim sSQL As String
    Dim strtmp As String
    
    
        If StatFg = "1" Then        '응급
            strtmp = " AND " & DBW("a.statfg = ", "1")
        ElseIf StatFg = "2" Then    '비응급
            strtmp = " AND " & DBW("a.statfg = ", "0")
        Else
            strtmp = ""
        End If
        
        If ErCheck = "1" Then
            strtmp = strtmp & " AND a.deptcd in (" & P_ErDeptCd & ")"
        End If
        
    
        sSQL = " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,b.vfyid,b.testcd,b.spccd,b.ptid,b.workarea,b.accdt,b.accseq," & _
               " c.abbrnm10,d." & F_PTNM & " as ptnm," & F_SSN2("d") & " as ssn," & _
               " e.bussdiv,e.wardid,e.hosilid,e.deptcd" & _
               " FROM " & T_LAB101 & " e," & T_HIS001 & " d," & T_LAB001 & " c," & T_LAB201 & " a," & T_LAB302 & " b" & _
               " WHERE " & _
                         DBW("a.rcvdt>=", Fdt) & _
               " AND " & DBW("a.rcvdt<=", Tdt) & _
               " AND " & DBW("a.stscd>", StsCd) & strtmp & _
               " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
               " AND " & DBW("b.testcd=", TestCd) & _
               " AND a.workarea=b.workarea" & _
               " AND a.accdt=b.accdt" & _
               " AND a.accseq=b.accseq" & _
               " AND a.workarea=c.workarea" & _
               " AND b.testcd=c.testcd" & _
               " AND b.ptid=e.ptid" & _
               " AND b.orddt=e.orddt" & _
               " AND b.ordno=e.ordno" & _
               " AND b.ptid=d." & F_PTID
   
        sSQL = sSQL & " UNION ALL " & _
               " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,b.vfyid,b.testcd,a.spccd,b.ptid,b.workarea,b.accdt,b.accseq," & _
               " c.abbrnm10,d." & F_PTNM & " as ptnm," & F_SSN2("d") & " as ssn," & _
               " e.bussdiv,e.wardid,e.hosilid,e.deptcd" & _
               " FROM " & T_LAB101 & " e," & T_HIS001 & " d," & T_LAB001 & " c," & T_LAB201 & " a," & T_LAB351 & " b" & _
               " WHERE " & _
                         DBW("a.rcvdt>=", Fdt) & _
               " AND " & DBW("a.rcvdt<=", Tdt) & _
               " AND " & DBW("a.stscd>", StsCd) & strtmp & _
               " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
               " AND " & DBW("b.testcd=", TestCd) & _
               " AND a.workarea=b.workarea" & _
               " AND a.accdt=b.accdt" & _
               " AND a.accseq=b.accseq" & _
               " AND a.workarea=c.workarea" & _
               " AND b.testcd=c.testcd" & _
               " AND b.ptid=e.ptid" & _
               " AND b.orddt=e.orddt" & _
               " AND b.ordno=e.ordno" & _
               " AND b.ptid=d." & F_PTID
   
'    If WorkArea = "07" Or WorkArea = "08" Or WorkArea = "09" Then
        sSQL = sSQL & " UNION ALL " & _
               " SELECT a.rcvdt,a.rcvtm,b.vfydt,b.vfytm,b.vfyid,'' as spccd ,b.testcd,b.ptid,b.workarea,b.accdt,b.accseq," & _
               " c.abbrnm10,d." & F_PTNM & " as ptnm," & F_SSN2("d") & " as ssn," & _
               " e.bussdiv,e.wardid,e.hosilid,e.deptcd" & _
               " FROM " & T_LAB101 & " e," & T_HIS001 & " d," & T_LAB001 & " c," & T_LAB201 & " a," & T_LAB404 & " b" & _
               " WHERE " & _
                         DBW("a.rcvdt>=", Fdt) & _
               " AND " & DBW("a.rcvdt<=", Tdt) & _
               " AND " & DBW("a.stscd>", StsCd) & strtmp & _
               " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
               " AND " & DBW("b.testcd=", TestCd) & _
               " AND a.workarea=b.workarea" & _
               " AND a.accdt=b.accdt" & _
               " AND a.accseq=b.accseq" & _
               " AND a.workarea=c.workarea" & _
               " AND b.testcd=c.testcd" & _
               " AND b.ptid=e.ptid" & _
               " AND b.orddt=e.orddt" & _
               " AND b.ordno=e.ordno" & _
               " AND b.ptid=d." & F_PTID & _
               " order by workarea,accdt,accseq,testcd"
'    Else
'
'    End If
    
    Set GetOverTATTime = New Recordset
    GetOverTATTime.Open sSQL, DBConn

End Function

Public Function GetCaseStudy1(ByVal sWorkArea As String, ByVal sAccDt As String, _
                            ByVal sFromSeq As String, ByVal sToSeq As String, _
                            ByVal sTestCd As String, ByVal sRstCd As String, _
                            ByVal sPtid As String, ByVal pStart As String, ByVal pEnd As String, _
                            ByVal sDeptCd As String, ByVal sTestDiv As String) As String

    Dim strDeptSql  As String
    Dim strAccSql   As String
    Dim strRstSql   As String
    Dim strPtntSql  As String
    Dim strWorkArea As String
    
    '## 5.0.25: 이상대(2005-05-31)
    '   - 접수일자 기준에서 보고일자 기준으로 검색
    '   - 일반검사의 경우 결과코드가 아닌 결과명도 조회하도록 수정
    strDeptSql = "": strAccSql = "": strPtntSql = "": strRstSql = ""
    If Trim(sDeptCd) <> "" Then
        strDeptSql = " AND  " & DBW("a.deptcd = ", sDeptCd)
    End If
    If Trim(sPtid) <> "" Then
        strPtntSql = " AND  " & DBW("a.ptid = ", sPtid)
    End If
    If Trim(sRstCd) <> "" Then
        strRstSql = " AND  b.rstcd in (" & sRstCd & ") "
    End If
'    If Trim(sWorkArea) <> "" Then
'        '-- 원본
'        strAccSql = " AND  " & DBW("a.workarea >= ", sWorkArea) & " AND  " & DBW("a.workarea <= ", sWorkArea)
'    End If
    If Trim(sAccDt) <> "" Then
        strAccSql = strAccSql & " AND  " & DBW("a.accdt = ", sAccDt)
    End If
    If Trim(sFromSeq) <> "" Then
        strAccSql = strAccSql & " AND  " & DBW("a.accseq >= ", sFromSeq)
    End If
    If Trim(sToSeq) <> "" Then
        strAccSql = strAccSql & " AND  " & DBW("a.accseq <= ", sToSeq)
    End If
    If Trim(sTestCd) <> "" Then
        strAccSql = strAccSql & " AND " & DBW("b.testcd=", sTestCd)
    End If
    
    Select Case sTestDiv
    Case TST_RouTest, TST_AboTest:
        '일반검사
        '## 5.0.27: 텍스트 결과도 조회할수 있도록 쿼리수정
        '## 5.0.31: 이상대(2005-07-21)
        '   - 텍스트 결과를 조회하기 위해 조인시 Outer 조인으로 수정
        If Trim(sWorkArea) <> "" Then
            '-- 원본
            strWorkArea = DBW("a.workarea >= ", sWorkArea) & " AND  " & DBW("a.workarea <= ", sWorkArea)
        End If
        
        '/*+ INDEX S2LAB201(S2LAB201_IDX4) **/
        
        GetCaseStudy1 = " SELECT /*+ RULE */ distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                       "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, a.stscd, j.rsttxt as txtrst, " & _
                       "        b.rstcd rstcd1, h.field1 as rstnm1, '' rstcd2, '' rstcd3, g.field3 as spcnm, " & _
                       "        b.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                       "        a.vfydt, a.vfytm, i.empnm " & _
                       " FROM  " & T_LAB303 & " j, " & T_COM006 & " i, " & T_LAB031 & " h, " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                   T_LAB302 & " b, " & T_LAB201 & " a " & _
                       " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                       strDeptSql & strPtntSql & _
                       "  AND  b.workarea = a.workarea " & "  AND  b.accdt    = a.accdt " & "  AND  b.accseq   = a.accseq " & _
                       strAccSql & strRstSql & _
                       "  AND  b.vfydt is not null" & _
                       "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                       "  AND  e." & F_PTID & " = b.ptid " & _
                       "  AND  d.testcd = b.testcd " & _
                       "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                              "WHERE  testcd = d.testcd) " & _
                       "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                       "  AND  g.cdval1 = a.spccd " & _
                       "  AND  " & DBJ("h.cdindex=*'" & LC2_ItemResult & "'") & _
                       "  AND  " & DBJ("h.cdval1=*b.testcd") & "  AND  " & DBJ("h.cdval2=*b.rstcd") & _
                       "  AND  " & DBJ("b.workarea*=j.workarea") & _
                       "  AND  " & DBJ("b.accdt*=j.accdt") & _
                       "  AND  " & DBJ("b.accseq*=j.accseq") & _
                       "  AND  " & DBJ("b.testcd*=j.testcd")
        
        GetCaseStudy1 = GetCaseStudy1 & " AND " & DBJ("i.empid=*a.vfyid")
    
    Case TST_SpeTest:
        '기타검사
        If Trim(sWorkArea) <> "" Then
            strWorkArea = DBW("a.workarea = ", sWorkArea) & ""
        End If
        If Trim(sRstCd) <> "" Then
            strRstSql = " AND  (b.rstcd1 in (" & sRstCd & ") or b.rstcd2 in (" & sRstCd & ")  or  b.rstcd3 in (" & sRstCd & "))"
        End If

'--2010.10.12
'--온승호
'--PBS 조회조건 수정

'        GetCaseStudy1 = " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
'                       "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, f.testcd, d.testnm, f.stscd, b.txtrst, " & _
'                       "        b.rstcd1, '' rstnm1, b.rstcd2, b.rstcd3, g.field3 as spcnm, " & _
'                       "        f.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
'                       "        a.vfydt, a.vfytm, '' as empnm " & _
'                       " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
'                                   T_LAB353 & " b, " & T_LAB351 & " f, " & T_LAB201 & " a " & _
'                       " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & _
'                       "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
'                       strDeptSql & strPtntSql & _
'                       "  AND  f.workarea = a.workarea " & _
'                       "  AND  f.accdt    = a.accdt " & _
'                       "  AND  f.accseq   = a.accseq " & _
'                       strAccSql & strRstSql & _
'                       "  AND  f.vfydt is not null" & _
'                       "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
'                       "  AND  e." & F_PTID & " = f.ptid " & _
'                       "  AND  b.workarea = f.workarea " & _
'                       "  AND  b.accdt    = f.accdt   AND  b.accseq   = f.accseq   AND  b.testcd   = f.testcd " & _
'                       "  AND  b.mfyseq = (SELECT max(mfyseq) FROM " & T_LAB353 & " WHERE workarea=b.workarea AND accdt=b.accdt AND accseq=b.accseq) " & _
'                       "  AND  d.testcd = f.testcd " & _
'                       "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
'                                                              "WHERE  testcd = d.testcd) " & _
'                       "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
'                       "  AND  g.cdval1 = a.spccd "
        GetCaseStudy1 = " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                       "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, f.testcd, d.testnm, f.stscd, b.txtrst, " & _
                       "        b.rstcd1, '' rstnm1, b.rstcd2, b.rstcd3, g.field3 as spcnm, " & _
                       "        f.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                       "        a.vfydt, a.vfytm, '' as empnm " & _
                       " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                   T_LAB353 & " b, " & T_LAB351 & " f, " & T_LAB201 & " a " & _
                       " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & _
                       "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                       strDeptSql & strPtntSql & _
                       "  AND  f.workarea = a.workarea " & _
                       "  AND  f.accdt    = a.accdt " & _
                       "  AND  f.accseq   = a.accseq " & _
                       strAccSql & strRstSql & _
                       "  AND  f.vfydt is not null" & _
                       "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                       "  AND  e." & F_PTID & " = f.ptid " & _
                       "  AND  b.workarea = f.workarea " & _
                       "  AND  b.accdt    = f.accdt   AND  b.accseq   = f.accseq   AND  b.testcd   = f.testcd " & _
                       "  AND  b.mfyseq = (SELECT max(mfyseq) FROM " & T_LAB353 & " WHERE workarea=b.workarea AND accdt=b.accdt AND accseq=b.accseq) " & _
                       "  AND  d.testcd = f.testcd " & _
                       "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                       "  AND  g.cdval1 = a.spccd "
    
    Case TST_MicTest:
        '미생물검사
        If Trim(sWorkArea) <> "" Then
            strWorkArea = DBW("a.workarea = ", sWorkArea) & ""
        End If
        GetCaseStudy1 = " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                       "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, b.stscd, '' txtrst, " & _
                       "        b.rstcd rstcd1, '' rstnm1, '' rstcd2, '' rstcd3, g.field3 as spcnm, " & _
                       "        b.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                       "        a.vfydt, a.vfytm, '' as empnm " & _
                       " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                   T_LAB404 & " b, " & T_LAB201 & " a " & _
                       " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & _
                       "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                       strDeptSql & strPtntSql & _
                       "  AND  b.workarea = a.workarea " & _
                       "  AND  b.accdt    = a.accdt " & _
                       "  AND  b.accseq   = a.accseq " & _
                       strAccSql & strRstSql & _
                       "  AND  b.vfydt is not null" & _
                       "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                       "  AND  e." & F_PTID & " = b.ptid " & _
                       "  AND  d.testcd = b.testcd " & _
                       "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                              "WHERE  testcd = d.testcd) " & _
                       "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                       "  AND  g.cdval1 = a.spccd "
    
    Case Else
        '** workarea가 TB 일경우 조회조건을 접수일로 변경(vfydt => rcvdt)
        If sWorkArea = "TB" Then
            '** 추가 By M.G.choi 2007.12.05 (일반검사 + 미생물검사 + 특수검사)
            If Trim(sWorkArea) <> "" Then
                '-- 원본
                strWorkArea = DBW("a.workarea >= ", sWorkArea) & " AND  " & DBW("a.workarea <= ", sWorkArea)
            End If
            
            '/*+ INDEX S2LAB201(S2LAB201_IDX4) **/

            GetCaseStudy1 = " SELECT /*+ RULE */ distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                           "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, a.stscd, j.rsttxt as txtrst, " & _
                           "        b.rstcd rstcd1, h.field1 as rstnm1, '' rstcd2, '' rstcd3, g.field3 as spcnm, " & _
                           "        b.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                           "        a.vfydt, a.vfytm, i.empnm " & _
                           " FROM  " & T_LAB303 & " j, " & T_COM006 & " i, " & T_LAB031 & " h, " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                       T_LAB302 & " b, " & T_LAB201 & " a " & _
                           " WHERE " & strWorkArea & " AND " & DBW("a.rcvdt>=", pStart) & " " & "  AND  " & DBW("a.rcvdt<=", pEnd) & " " & _
                           strDeptSql & strPtntSql & _
                           "  AND  b.workarea = a.workarea " & "  AND  b.accdt    = a.accdt " & "  AND  b.accseq   = a.accseq " & _
                           strAccSql & strRstSql & _
                           "  AND  b.vfydt is not null" & _
                           "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                           "  AND  e." & F_PTID & " = b.ptid " & _
                           "  AND  d.testcd = b.testcd " & _
                           "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                                  "WHERE  testcd = d.testcd) " & _
                           "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                           "  AND  g.cdval1 = a.spccd " & _
                           "  AND  " & DBJ("h.cdindex=*'" & LC2_ItemResult & "'") & _
                           "  AND  " & DBJ("h.cdval1=*b.testcd") & "  AND  " & DBJ("h.cdval2=*b.rstcd") & _
                           "  AND  " & DBJ("b.workarea*=j.workarea") & _
                           "  AND  " & DBJ("b.accdt*=j.accdt") & _
                           "  AND  " & DBJ("b.accseq*=j.accseq") & _
                           "  AND  " & DBJ("b.testcd*=j.testcd")
            
            GetCaseStudy1 = GetCaseStudy1 & " AND " & DBJ("i.empid=*a.vfyid")
            
            GetCaseStudy1 = GetCaseStudy1 & " union all "
            
            '기타검사
            If Trim(sWorkArea) <> "" Then
                strWorkArea = DBW("a.workarea = ", sWorkArea) & ""
            End If
            If Trim(sRstCd) <> "" Then
                strRstSql = " AND  (b.rstcd1 in (" & sRstCd & ") or b.rstcd2 in (" & sRstCd & ")  or  b.rstcd3 in (" & sRstCd & "))"
            End If
            GetCaseStudy1 = GetCaseStudy1 & " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                           "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, f.testcd, d.testnm, f.stscd, b.txtrst, " & _
                           "        b.rstcd1, '' rstnm1, b.rstcd2, b.rstcd3, g.field3 as spcnm, " & _
                           "        f.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                           "        a.vfydt, a.vfytm, '' as empnm " & _
                           " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                       T_LAB353 & " b, " & T_LAB351 & " f, " & T_LAB201 & " a " & _
                           " WHERE " & strWorkArea & " AND " & DBW("a.rcvdt>=", pStart) & " " & _
                           "  AND  " & DBW("a.rcvdt<=", pEnd) & " " & _
                           strDeptSql & strPtntSql & _
                           "  AND  f.workarea = a.workarea " & _
                           "  AND  f.accdt    = a.accdt " & _
                           "  AND  f.accseq   = a.accseq " & _
                           strAccSql & strRstSql & _
                           "  AND  f.vfydt is not null" & _
                           "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                           "  AND  e." & F_PTID & " = f.ptid " & _
                           "  AND  b.workarea = f.workarea " & _
                           "  AND  b.accdt    = f.accdt   AND  b.accseq   = f.accseq   AND  b.testcd   = f.testcd " & _
                           "  AND  b.mfyseq = (SELECT max(mfyseq) FROM " & T_LAB353 & " WHERE workarea=b.workarea AND accdt=b.accdt AND accseq=b.accseq) " & _
                           "  AND  d.testcd = f.testcd " & _
                           "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                                  "WHERE  testcd = d.testcd) " & _
                           "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                           "  AND  g.cdval1 = a.spccd "
            
            
            GetCaseStudy1 = GetCaseStudy1 & " union all "
            
            '미생물검사
            If Trim(sWorkArea) <> "" Then
                strWorkArea = DBW("a.workarea = ", sWorkArea) & ""
            End If
            GetCaseStudy1 = GetCaseStudy1 & " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                           "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, b.stscd, '' txtrst, " & _
                           "        b.rstcd rstcd1, '' rstnm1, '' rstcd2, '' rstcd3, g.field3 as spcnm, " & _
                           "        b.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                           "        a.vfydt, a.vfytm, '' as empnm " & _
                           " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                       T_LAB404 & " b, " & T_LAB201 & " a " & _
                           " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & _
                           "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                           strDeptSql & strPtntSql & _
                           "  AND  b.workarea = a.workarea " & _
                           "  AND  b.accdt    = a.accdt " & _
                           "  AND  b.accseq   = a.accseq " & _
                           strAccSql & strRstSql & _
                           "  AND  b.vfydt is not null" & _
                           "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                           "  AND  e." & F_PTID & " = b.ptid " & _
                           "  AND  d.testcd = b.testcd " & _
                           "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                                  "WHERE  testcd = d.testcd) " & _
                           "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                           "  AND  g.cdval1 = a.spccd "
        Else
            '** 추가 By M.G.choi 2007.12.05 (일반검사 + 미생물검사 + 특수검사)
            If Trim(sWorkArea) <> "" Then
                '-- 원본
                strWorkArea = DBW("a.workarea >= ", sWorkArea) & " AND  " & DBW("a.workarea <= ", sWorkArea)
            End If
            
            '/*+ INDEX S2LAB201(S2LAB201_IDX4) **/

            GetCaseStudy1 = " SELECT /*+ RULE */ distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                           "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, a.stscd, j.rsttxt as txtrst, " & _
                           "        b.rstcd rstcd1, h.field1 as rstnm1, '' rstcd2, '' rstcd3, g.field3 as spcnm, " & _
                           "        b.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                           "        a.vfydt, a.vfytm, i.empnm " & _
                           " FROM  " & T_LAB303 & " j, " & T_COM006 & " i, " & T_LAB031 & " h, " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                       T_LAB302 & " b, " & T_LAB201 & " a " & _
                           " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                           strDeptSql & strPtntSql & _
                           "  AND  b.workarea = a.workarea " & "  AND  b.accdt    = a.accdt " & "  AND  b.accseq   = a.accseq " & _
                           strAccSql & strRstSql & _
                           "  AND  b.vfydt is not null" & _
                           "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                           "  AND  e." & F_PTID & " = b.ptid " & _
                           "  AND  d.testcd = b.testcd " & _
                           "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                                  "WHERE  testcd = d.testcd) " & _
                           "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                           "  AND  g.cdval1 = a.spccd " & _
                           "  AND  " & DBJ("h.cdindex=*'" & LC2_ItemResult & "'") & _
                           "  AND  " & DBJ("h.cdval1=*b.testcd") & "  AND  " & DBJ("h.cdval2=*b.rstcd") & _
                           "  AND  " & DBJ("b.workarea*=j.workarea") & _
                           "  AND  " & DBJ("b.accdt*=j.accdt") & _
                           "  AND  " & DBJ("b.accseq*=j.accseq") & _
                           "  AND  " & DBJ("b.testcd*=j.testcd")
            
            GetCaseStudy1 = GetCaseStudy1 & " AND " & DBJ("i.empid=*a.vfyid")
            
            GetCaseStudy1 = GetCaseStudy1 & " union all "
            
            '기타검사
            If Trim(sWorkArea) <> "" Then
                strWorkArea = DBW("a.workarea = ", sWorkArea) & ""
            End If
            If Trim(sRstCd) <> "" Then
                strRstSql = " AND  (b.rstcd1 in (" & sRstCd & ") or b.rstcd2 in (" & sRstCd & ")  or  b.rstcd3 in (" & sRstCd & "))"
            End If
            GetCaseStudy1 = GetCaseStudy1 & " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                           "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, f.testcd, d.testnm, f.stscd, b.txtrst, " & _
                           "        b.rstcd1, '' rstnm1, b.rstcd2, b.rstcd3, g.field3 as spcnm, " & _
                           "        f.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                           "        a.vfydt, a.vfytm, '' as empnm " & _
                           " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                       T_LAB353 & " b, " & T_LAB351 & " f, " & T_LAB201 & " a " & _
                           " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & _
                           "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                           strDeptSql & strPtntSql & _
                           "  AND  f.workarea = a.workarea " & _
                           "  AND  f.accdt    = a.accdt " & _
                           "  AND  f.accseq   = a.accseq " & _
                           strAccSql & strRstSql & _
                           "  AND  f.vfydt is not null" & _
                           "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                           "  AND  e." & F_PTID & " = f.ptid " & _
                           "  AND  b.workarea = f.workarea " & _
                           "  AND  b.accdt    = f.accdt   AND  b.accseq   = f.accseq   AND  b.testcd   = f.testcd " & _
                           "  AND  b.mfyseq = (SELECT max(mfyseq) FROM " & T_LAB353 & " WHERE workarea=b.workarea AND accdt=b.accdt AND accseq=b.accseq) " & _
                           "  AND  d.testcd = f.testcd " & _
                           "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                                  "WHERE  testcd = d.testcd) " & _
                           "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                           "  AND  g.cdval1 = a.spccd "
            
            
            GetCaseStudy1 = GetCaseStudy1 & " union all "
            
            '미생물검사
            If Trim(sWorkArea) <> "" Then
                strWorkArea = DBW("a.workarea = ", sWorkArea) & ""
            End If
            GetCaseStudy1 = GetCaseStudy1 & " SELECT distinct a.buildcd, a.workarea, a.accdt, a.accseq, a.spccd, a.deptcd, a.wardid, a.hosilid, " & _
                           "        a.coldt, a.coltm, a.rcvdt, a.rcvtm, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, b.stscd, '' txtrst, " & _
                           "        b.rstcd rstcd1, '' rstnm1, '' rstcd2, '' rstcd3, g.field3 as spcnm, " & _
                           "        b.ptid, e." & F_PTNM & " ptnm, " & F_SEX2("e") & " sex, " & F_DOB2("e") & " dob, " & _
                           "        a.vfydt, a.vfytm, '' as empnm " & _
                           " FROM  " & T_LAB032 & " g, " & T_HIS001 & " e, " & T_HIS003 & " c, " & T_LAB001 & " d, " & _
                                       T_LAB404 & " b, " & T_LAB201 & " a " & _
                           " WHERE " & strWorkArea & " AND " & DBW("a.vfydt>=", pStart) & " " & _
                           "  AND  " & DBW("a.vfydt<=", pEnd) & " " & _
                           strDeptSql & strPtntSql & _
                           "  AND  b.workarea = a.workarea " & _
                           "  AND  b.accdt    = a.accdt " & _
                           "  AND  b.accseq   = a.accseq " & _
                           strAccSql & strRstSql & _
                           "  AND  b.vfydt is not null" & _
                           "  AND  " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                           "  AND  e." & F_PTID & " = b.ptid " & _
                           "  AND  d.testcd = b.testcd " & _
                           "  AND  d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                                  "WHERE  testcd = d.testcd) " & _
                           "  AND  " & DBW("g.cdindex=", LC3_Specimen) & _
                           "  AND  g.cdval1 = a.spccd "
        End If
    End Select
End Function


Public Function GetAccCnt_Bussdiv(ByVal pStart As String, ByVal pEnd As String, ByVal InOut As String)
'FROM LAB201, LAB302, LAB001, HIS003

    Dim SqlStmt As String
        
    '일반검사
    '** 원본 ============================================================================================================
'    SqlStmt = "SELECT a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
    '====================================================================================================================
    
    '** 예수병원 변경루틴 By M.G.Choi 2004.12.28 ========================================================================
    SqlStmt = "SELECT /*+ rule */ " & _
              "       a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
    '====================================================================================================================
    
    SqlStmt = SqlStmt & "FROM  " & T_HIS003 & " c, " & T_LAB001 & " d, " & T_LAB302 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    
    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = b.testcd) "
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If

    Select Case InOut
        Case "1": SqlStmt = SqlStmt & " AND (a.wardid is null  or a.wardid='HC' or a.wardid='HDR') "
        Case "2": SqlStmt = SqlStmt & " AND a.wardid not in (' ' ,'HC','HDR') "
    End Select
    
    SqlStmt = SqlStmt & " GROUP BY a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & ", b.testcd, d.testnm "
    '기타검사
    SqlStmt = SqlStmt & "UNION ALL SELECT a.buildcd, a.workarea, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & T_HIS003 & " c, " & T_LAB001 & " d, " & T_LAB351 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    'SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = b.testcd) "
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If

    Select Case InOut
        Case "1": SqlStmt = SqlStmt & " AND (a.wardid is null or a.wardid='HC' or a.wardid='HDR') "
        Case "2": SqlStmt = SqlStmt & " AND a.wardid not in (' ' ,'HC','HDR') "
    End Select
    
    SqlStmt = SqlStmt & " GROUP BY a.buildcd, a.workarea, c." & F_DEPTNM & ", b.testcd, d.testnm "
    '미생물검사
    SqlStmt = SqlStmt & "UNION ALL SELECT a.buildcd, a.workarea, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & T_HIS003 & " c, " & T_LAB001 & " d, " & T_LAB404 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = b.testcd) "
    
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If
    
    Select Case InOut
        Case "1": SqlStmt = SqlStmt & " AND (a.wardid is null or a.wardid='HC' or a.wardid='HDR') "
        Case "2": SqlStmt = SqlStmt & " AND a.wardid not in (' ' ,'HC','HDR') "
    End Select
    
    SqlStmt = SqlStmt & " GROUP BY a.buildcd, a.workarea, c." & F_DEPTNM & ", b.testcd, d.testnm "
    SqlStmt = SqlStmt & " ORDER BY buildcd, workarea, eqpcd, DeptNm, testcd, testnm "
    
    GetAccCnt_Bussdiv = SqlStmt
    
End Function

'[환자건수]: 입퇴원환자 건수 구하기
Public Function GetPatientCount(ByVal Bussdiv As String, ByVal StartDt As String, ByVal EndDt As String) As String
    Dim sSQL    As String
    
    sSQL = " SELECT a.deptcd,a.wardid,a.orddoct,count(a.orddoct) as cnt " & _
           " FROM " & T_LAB101 & " a" & _
           " WHERE " & DBW("a.orddiv=", lis_orddiv) & _
           " AND " & DBW("a.bussdiv=", Bussdiv) & _
           " AND " & DBW("a.donefg>=", enStsCd.StsCd_LIS_Order) & _
           " AND " & DBW("a.orddt>=", StartDt) & _
           " AND " & DBW("a.orddt<=", EndDt) & _
           " GROUP BY a.deptcd,a.wardid,a.orddoct " & _
           " ORDER BY a.deptcd,a.wardid,a.orddoct"
    GetPatientCount = sSQL
End Function
'[WorkArea별 검사항목코드]: itemseq>0인항목만 통계처리한다.
Public Function GetWorkareaTestItem(ByVal sWorkArea As String) As String
    Dim sSQL    As String
    
    sSQL = " SELECT testcd, testnm " & _
           " FROM " & T_LAB001 & _
           " WHERE " & DBW("workarea=", sWorkArea)
    
    If P_ItemSeqFG = True Then
        sSQL = sSQL & " AND " & DBW("itemseq>", "0")
    End If
    
    sSQL = sSQL & " order by testcd "
    GetWorkareaTestItem = sSQL
End Function
'[WorkLoad 검색]
Public Function GetEtcWorkLoad(Optional ByVal sCdval1 As String = "")
    Dim sSQL As String
    
    sSQL = " SELECT cdval1,field1,field3 FROM " & T_LAB032 & _
           " WHERE " & DBW("cdindex=", LC3_EtcWorkLoad)
    If sCdval1 <> "" Then
        sSQL = sSQL & " AND " & DBW("cdval1=", sCdval1)
    End If
    GetEtcWorkLoad = sSQL
End Function
'[WorkLoad에서 사용하는 검사항목리스트]
Public Function GetWorkLoadPopUpSQL(Optional ByVal sTestCd As String = "") As String
    Dim sSQL As String
    
    sSQL = " SELECT a.testcd,a.abbrnm10,a.workarea as testnm " & _
           " FROM " & T_LAB001 & " a " & _
           " WHERE (a.expdt='' or a.expdt is null)" & _
           " AND a.applydt = ( SELECT max(a.applydt) FROM " & T_LAB001 & " c " & _
                        "                              WHERE c.testcd = testcd ) "
    If sTestCd <> "" Then
        sSQL = sSQL & " AND " & DBW("testcd=", sTestCd)
    End If
    sSQL = sSQL & " order by a.testcd"
    
    GetWorkLoadPopUpSQL = sSQL
End Function
'[WorkLoad Item Delete]
Public Function DeleteWorkLoadItem(ByVal cdindex As String, ByVal sTestCd As String) As String
    DeleteWorkLoadItem = "delete FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", cdindex) & " AND " & DBW("cdval1=", sTestCd)
End Function
'[WorkLoad Item Insert]
Public Function InsertWorkLoadItem(ByVal cdindex As String, ByVal CdVal1 As String, ByVal Field1 As String, _
                                   ByVal Field2 As String, ByVal Field3 As String, ByVal Field4 As String, ByVal Field5 As String) As String
    InsertWorkLoadItem = "insert into " & T_LAB032 & " (cdindex,cdval1,field1,field2,field3,field4,field5)" & _
                         " values (" & _
                         DBV("cdindex", cdindex, 1) & DBV("cdval1", CdVal1, 1) & _
                         DBV("field1", Field1, 1) & DBV("field2", Field2, 1) & _
                         DBV("field3", Field3, 1) & DBV("field4", Field4, 1) & _
                         DBV("field5", Field5) & _
                         ")"
End Function
'[WorkLoad Item GetSQL]
Public Function GetWorkLoadItemSQL(ByVal cdindex As String) As String
    GetWorkLoadItemSQL = "SELECT * FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", cdindex)
End Function
'[WorkLoad Unit Query SQL]
Public Function GetWorkLoadUnitQuery(ByVal strEmpID As String, ByVal StrStart As String, ByVal StrEnd As String, ByVal StrQC As String) As String
    Dim SSQL1       As String
    Dim SSQL2       As String
    Dim SSQL3       As String
    Dim SSQL4       As String
    
    '일반검사
    SSQL1 = " SELECT c.cdval1 as testcd,d.empnm,c.field1 as testnm,c.field3 as unit ," & _
            " ' ' qcfg,count(c.cdval1) as cnt" & _
            " FROM " & T_COM006 & " d," & T_LAB032 & " c," & T_LAB201 & " a " & _
            " WHERE" & _
                      DBW("a.vfydt >=", StrStart) & _
            " AND " & DBW("a.vfydt<=", StrEnd) & _
            " AND " & DBW("a.vfyid=", strEmpID) & _
            " AND " & DBW("c.cdindex=", LC3_ItemWorkLoad) & _
            " AND exists( SELECT * FROM " & T_LAB302 & " z" & _
                        " WHERE z.workarea=a.workarea AND z.accdt=a.accdt AND z.accseq=a.accseq" & _
                        " AND  (z.rstdiv<>'*' ) AND c.cdval1=z.testcd)" & _
            " AND " & DBW("a.qcfg<>", "1") & _
            " AND d.empid=a.vfyid" & _
            " group by c.cdval1,d.empnm,c.field1,c.field3"
    '특수검사
    SSQL2 = " SELECT c.cdval1 as testcd,d.empnm,c.field1 as testnm,c.field3 as unit ," & _
            " ' ' qcfg,count(c.cdval1) as cnt" & _
            " FROM " & T_COM006 & " d," & T_LAB032 & " c," & T_LAB201 & " a " & _
            " WHERE" & _
                      DBW("a.vfydt >=", StrStart) & _
            " AND " & DBW("a.vfydt<=", StrEnd) & _
            " AND " & DBW("a.vfyid=", strEmpID) & _
            " AND " & DBW("c.cdindex=", LC3_ItemWorkLoad) & _
            " AND exists( SELECT * FROM " & T_LAB351 & " z" & _
                        " WHERE z.workarea=a.workarea AND z.accdt=a.accdt AND z.accseq=a.accseq" & _
                        " AND c.cdval1=z.testcd)" & _
            " AND " & DBW("a.qcfg<>", "1") & _
            " AND d.empid=a.vfyid" & _
            " group by c.cdval1,d.empnm,c.field1,c.field3"
    
    '미생물검사
    SSQL3 = " SELECT c.cdval1 as testcd,d.empnm,c.field1 as testnm,c.field3 as unit ," & _
            " ' ' qcfg,count(c.cdval1) as cnt" & _
            " FROM " & T_COM006 & " d," & T_LAB032 & " c," & T_LAB201 & " a " & _
            " WHERE" & _
                      DBW("a.vfydt >=", StrStart) & _
            " AND " & DBW("a.vfydt<=", StrEnd) & _
            " AND " & DBW("a.vfyid=", strEmpID) & _
            " AND " & DBW("c.cdindex=", LC3_ItemWorkLoad) & _
            " AND exists( SELECT * FROM " & T_LAB404 & " z" & _
                        " WHERE z.workarea=a.workarea AND z.accdt=a.accdt AND z.accseq=a.accseq" & _
                        " AND  (z.rstdiv<>'*') AND c.cdval1=z.testcd)" & _
            " AND " & DBW("a.qcfg<>", "1") & _
            " AND d.empid=a.vfyid" & _
            " group by c.cdval1,d.empnm,c.field1,c.field3"
    
    'QC
    SSQL4 = " SELECT c.cdval1 as testcd,d.empnm,c.field1 as testnm,c.field5 as unit ," & _
            " '1' qcfg,count(c.cdval1) as cnt" & _
            " FROM " & T_COM006 & " d," & T_LAB032 & " c," & T_LAB201 & " a " & _
            " WHERE" & _
                      DBW("a.vfydt >=", StrStart) & _
            " AND " & DBW("a.vfydt<=", StrEnd) & _
            " AND " & DBW("a.vfyid=", strEmpID) & _
            " AND " & DBW("c.cdindex=", LC3_ItemWorkLoad) & _
            " AND exists( SELECT * FROM " & T_LAB026 & " z" & _
                        " WHERE z.workarea=a.workarea AND z.accdt=a.accdt AND z.accseq=a.accseq" & _
                        " AND  (z.rstdiv<>'*') AND c.cdval1=z.testcd)" & _
            " AND " & DBW("a.qcfg=", "1") & _
            " AND d.empid=a.vfyid" & _
            " group by c.cdval1,d.empnm,c.field1, c.field5"
    
    GetWorkLoadUnitQuery = SSQL1 & " UNION ALL " & SSQL2 & " UNION ALL " & SSQL3
    If StrQC <> "" Then
        GetWorkLoadUnitQuery = GetWorkLoadUnitQuery & " UNION ALL " & SSQL4
    End If
    GetWorkLoadUnitQuery = GetWorkLoadUnitQuery & " order by testcd"
End Function
'[BloodCulture Count]
Public Function GetBloodCultureCount(ByVal strWorkArea As String, ByVal strYear As String, ByVal StrFrom As String, _
                                         ByVal strTo As String, ByVal SrcGroup As String) As String
    Dim sSQL As String
    sSQL = " SELECT wardid ,count(*) as cnt FROM " & T_LAB201 & _
           " WHERE " & DBW("workarea=", strWorkArea) & _
           " AND " & DBW("accdt=", strYear) & _
           " AND " & DBW("accseq>=", StrFrom) & _
           " AND " & DBW("accseq<=", strTo) & _
           " AND " & DBW("stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND spccd in (SELECT cdval1 FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", LC3_Specimen) & _
           " AND " & DBW("field2=", SrcGroup) & ")" & _
           " AND (wardid <>'' or wardid is not null)" & _
           " group by wardid" & _
           " order by wardid"
    GetBloodCultureCount = sSQL
End Function
'[BloodCulture Detail]
Public Function GetBloodCultureDetailSQL(ByVal strWorkArea As String, ByVal strYear As String, ByVal StrFrom As String, _
                                         ByVal strTo As String, ByVal strWardId As String) As String
    
    Dim sSQL As String
    
    sSQL = " SELECT a.ptid,b." & F_PTNM & " as ptnm,a.coldt,a.coltm,a.rcvdt,a.rcvtm,a.vfydt,a.vfytm,a.stscd,c.field3 as spcnm," & _
           " a.workarea " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accdt" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accseq as labno" & _
           " FROM " & T_LAB201 & " a," & T_HIS001 & " b," & T_LAB032 & " c" & _
           " WHERE " & _
                     DBW("a.workarea=", strWorkArea) & _
           " AND " & DBW("a.accdt=", strYear) & _
           " AND " & DBW("a.accseq>=", StrFrom) & _
           " AND " & DBW("a.accseq<=", strTo) & _
           " AND " & DBW("a.stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("a.wardid=", strWardId) & _
           " AND a.ptid=b." & F_PTID & _
           " AND " & DBW("cdindex=", LC3_Specimen) & _
           " AND a.spccd=c.cdval1"
    GetBloodCultureDetailSQL = sSQL
    
End Function

'[BloodCulture 접수번호 범위]
Public Function GetBloodCultureAccSeqRange(ByVal SrcGroup As String)
    GetBloodCultureAccSeqRange = " SELECT * FROM " & T_LAB032 & _
                                 " WHERE " & DBW("cdindex=", LC3_SGroup) & _
                                 " AND " & DBW("cdval1=", SrcGroup)
End Function
'[검체,검체군별 배양균 통계]
Public Function GetSpcSensiSQL(ByVal qStart As String, ByVal qEnd As String, Optional ByVal sPtid As String = "", _
                               Optional ByVal sWard As String = "", Optional ByVal sDoct As String = "", _
                               Optional ByVal sSpcCd As String = "", Optional ByVal sMncd As String = "") As String
                               
    Dim sSQL    As String
    

    sSQL = " SELECT distinct a.workarea,a.accdt,a.accseq,a.buildcd as BuildCd ,f.field2 as SpcGrp, a.spccd,f.field5 as spcnm, e.field1 as SpcBac, e.text1 as SpcFull, " & _
           "        a.ptid,a.wardid,a.hosilid,a.deptcd,a.majdoct,c." & F_PTNM & " as ptnm ,g.field2 as qtnm" & _
           " FROM " & T_LAB032 & " g," & T_HIS001 & " c," & T_LAB032 & " f," & T_LAB032 & " e," & T_LAB405 & " b," & T_LAB201 & " a" & _
           " WHERE  " & DBW("a.workarea >= ", MIC_WorkArea) & _
           " AND    " & DBW("a.workarea <= ", MIC_WorkArea) & _
           " AND    " & DBW("a.rcvdt >=", qStart) & _
           " AND    " & DBW("a.rcvdt <=", qEnd) & _
           " AND    " & DBW("a.stscd<", enStsCd.StsCd_LIS_Cancel)
           
    If sPtid <> "" Then sSQL = sSQL & " AND " & DBW("a.ptid=", sPtid)
    If sWard <> "" Then sSQL = sSQL & " AND " & DBW("a.wardid=", sWard)
    If sDoct <> "" Then sSQL = sSQL & " AND " & DBW("a.majdoct=", sDoct)
    If sSpcCd <> "" Then sSQL = sSQL & " AND" & DBW("a.spccd=", sSpcCd)
    If sMncd <> "" Then sSQL = sSQL & " AND " & DBW("b.mnmcd=", sMncd)
    
    sSQL = sSQL & _
           " AND    a.workarea = b.workarea" & _
           " AND    a.accdt = b.accdt" & _
           " AND    a.accseq = b.accseq" & _
           " AND    " & DBW("e.cdindex =", LC3_Microbe) & _
           " AND    e.cdval1 = b.mnmcd" & _
           " AND    " & DBW("f.cdindex = ", LC3_Specimen) & _
           " AND    f.cdval1 = a.spccd" & _
           " AND    " & DBW("g.cdindex=", LC3_Volume) & _
           " AND    g.cdval1=b.mqtcd" & _
           " AND    c." & F_PTID & "=a.ptid" & _
           " ORDER BY ptid,workarea,accdt,accseq,BuildCd,wardid, SpcGrp, SpcBac ,spccd"
    GetSpcSensiSQL = sSQL
End Function

Public Function GetNewResultLst(ByVal pStartDt As String, ByVal pEndDt As String, Optional ByVal pStartTM As String = "", Optional ByVal pEndTM As String = "", Optional ByVal pWorkArea As String = "") As String
    Dim SQL, strTemp As String
    strTemp = ""
    If pWorkArea <> "" Then strTemp = " AND " & DBW("a.workarea=", pWorkArea)
    
    '## 5.0.25: 이상대(2005-05-31)
    '   - 검사항목별 결과코드에 등록된 결과명을 조회하도록 쿼리수정
    SQL = " SELECT /*+ INDEX(a s2lab201_idx4 */a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, e.testnm, a.wardid as wardid, a.hosilid as roomid, " & _
            "        a.ptid as ptid, d." & F_PTNM & " as ptnm, g.empnm, b.vfydt, b.vfytm, c.rsttxt as rsttxt, a.vfyid as vid, a.vfydt as vdt, a.vfytm as vtm, a.COLDT, a.COLTM, a.COLID, a.RCVDT, a.RCVTM, a.DEPTCD, f.FIELD3 as spcnm " & _
            " FROM " & T_LAB302 & " b, " & T_LAB304 & " c, " & T_HIS001 & " d, " & T_LAB001 & " e, " & T_LAB032 & " f, " & _
                     T_COM006 & " g, " & T_LAB201 & " a " & _
            " WHERE " & DBW("a.vfydt >= ", pStartDt) & " AND " & DBW("a.vfydt <= ", pEndDt) & strTemp & _
              " AND a.stscd >= '5' and b.workarea = a.workarea AND b.accdt    = a.accdt AND b.accseq   = a.accseq" & _
              " AND " & DBJ("c.workarea =* a.workarea") & _
              " AND " & DBJ("c.accdt    =* a.accdt") & _
              " AND " & DBJ("c.accseq   =* a.accseq") & _
              " AND e.testcd = b.testcd" & _
              " AND e.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = e.testcd)" & _
              " AND d." & F_PTID & " = a.ptid " & _
              " AND a.vfydt <> ' ' " & _
              " AND a.vfydt is not null " & _
              " AND g.empid = a.vfyid AND f.CDINDEX = 'C215' AND a.SPCCD = f.CDVAL1"

    SQL = " select * from ( " & SQL & " ) z where z.vdt||z.vtm >=  '" & pStartDt & pStartTM & "'  and z.vdt||z.vtm <= '" & pEndDt & pEndTM & "'"
                      
    GetNewResultLst = SQL & " order by vdt,vtm desc"

'            " AND ( " & pCondQuery & " )" & _


End Function


Public Function GetNewResultLst_New(ByVal pStartDt As String, ByVal pEndDt As String, Optional ByVal pStartTM As String = "", Optional ByVal pEndTM As String = "", Optional ByVal pWorkArea As String = "") As String
    Dim SQL, strTemp As String
    strTemp = ""
    If pWorkArea <> "" Then strTemp = " AND " & DBW("a.workarea=", pWorkArea)
    
    '## 5.0.25: 이상대(2005-05-31)
    '   - 검사항목별 결과코드에 등록된 결과명을 조회하도록 쿼리수정
    SQL = " SELECT /*+ INDEX(a s2lab201_idx4 */a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, e.testnm, a.wardid as wardid, a.hosilid as roomid, " & _
            "        a.ptid as ptid, d." & F_PTNM & " as ptnm, g.empnm, b.vfydt, b.vfytm, c.rsttxt as rsttxt, a.vfyid as vid, a.vfydt as vdt, a.vfytm as vtm, a.COLDT, a.COLTM, a.COLID, a.RCVDT, a.RCVTM, a.DEPTCD, f.FIELD3 as spcnm " & _
            " FROM " & T_LAB302 & " b, " & T_LAB304 & " c, " & T_HIS001 & " d, " & T_LAB001 & " e, " & T_LAB032 & " f, " & _
                     T_COM006 & " g, " & T_LAB201 & " a " & _
            " WHERE " & DBW("a.vfydt >= ", pStartDt) & " AND " & DBW("a.vfydt <= ", pEndDt) & strTemp & _
              " AND a.stscd >= '5' and b.workarea = a.workarea AND b.accdt    = a.accdt AND b.accseq   = a.accseq" & _
              " AND " & DBJ("c.workarea =* a.workarea") & _
              " AND " & DBJ("c.accdt    =* a.accdt") & _
              " AND " & DBJ("c.accseq   =* a.accseq") & _
              " AND e.testcd = b.testcd" & _
              " AND e.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = e.testcd)" & _
              " AND d." & F_PTID & " = a.ptid " & _
              " AND a.vfydt <> ' ' " & _
              " AND a.vfydt is not null " & _
              " AND g.empid = a.vfyid AND f.CDINDEX = 'C215' AND a.SPCCD = f.CDVAL1 "

    SQL = " select * from ( " & SQL & " ) z where z.vdt >=  '" & pStartDt & "'  and z.vdt <= '" & pEndDt & "'"
                      
    GetNewResultLst_New = SQL & " order by vdt,vtm desc"

'            " AND ( " & pCondQuery & " )" & _


End Function

Public Function GetAccCnt_Bussdiv_New(ByVal pStart As String, ByVal pEnd As String, ByVal InOut As String, ByVal StatFg As String, ByVal StatPM As String)
'FROM LAB201, LAB302, LAB001, HIS003

    Dim SqlStmt As String
        
    '일반검사
    SqlStmt = "SELECT a.workarea, b.testcd, d.testnm, c.tats, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & T_LAB004 & " c, " & T_LAB001 & " d, " & T_LAB302 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd and applydt<=a.rcvdt) "
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If
    
    Select Case StatFg
        Case 1
            SqlStmt = SqlStmt & " AND c.statfg='1'"
        Case 2
            SqlStmt = SqlStmt & " AND c.statfg='0'"
    End Select
    
    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
        SqlStmt = SqlStmt & " AND (a.wardid <>'' or wardid is not null) "
    End If
    
    SqlStmt = SqlStmt & " AND (tats = '' OR tats IS not null)  AND d.TESTCD = c.TESTCD AND b.SPCCD = c.SPCCD AND c.tats <> ';'        AND c.tats <> ';;'  "
    SqlStmt = SqlStmt & " GROUP BY a.workarea, b.testcd, d.testnm, c.tats "
    '기타검사
    SqlStmt = SqlStmt & "UNION ALL SELECT a.workarea, b.testcd, d.testnm, c.tats, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & T_LAB004 & " c, " & T_LAB001 & " d, " & T_LAB351 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd and applydt<=a.rcvdt) "
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If
    
    Select Case StatFg
        Case 1
            SqlStmt = SqlStmt & " AND c.statfg='1'"
        Case 2
            SqlStmt = SqlStmt & " AND c.statfg='0'"
    End Select
    
    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
        SqlStmt = SqlStmt & " AND (a.wardid <>'' or wardid is not null) "
    End If
    
    SqlStmt = SqlStmt & " AND (tats = '' OR tats IS not null)  AND d.TESTCD = c.TESTCD AND c.tats <> ';'        AND c.tats <> ';;'  "
    SqlStmt = SqlStmt & " GROUP BY a.workarea, b.testcd, d.testnm, c.tats "
    '미생물검사
'    SqlStmt = SqlStmt & "UNION ALL SELECT a.workarea, b.testcd, d.testnm, c.tats, count(*) as Cnt "
'    SqlStmt = SqlStmt & "FROM  " & T_LAB004 & " c, " & T_LAB001 & " d, " & T_LAB404 & " b, " & T_LAB201 & " a "
'    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
'
'    If mvarWorkarea <> "" Then
'        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
'        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
'    End If
'
'    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
'    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
'    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
'    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
'    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
'                                          "WHERE  testcd = d.testcd and applydt<=a.rcvdt) "
'
'    If P_ItemSeqFG = True Then
'        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
'    Else
'        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
'    End If
'
'    Select Case StatFg
'        Case 1
'            SqlStmt = SqlStmt & " AND c.statfg='1'"
'        Case 2
'            SqlStmt = SqlStmt & " AND c.statfg='0'"
'    End Select
'
'    If StatPM = 1 Then
'        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
'    End If
'
'    SqlStmt = SqlStmt & " AND (tats = '' OR tats IS not null)  AND d.TESTCD = c.TESTCD AND c.tats <> ';'        AND c.tats <> ';;'  "
'    SqlStmt = SqlStmt & " GROUP BY a.workarea, b.testcd, d.testnm, c.tats "
    SqlStmt = SqlStmt & " ORDER BY workarea, testcd, testnm "
    
    GetAccCnt_Bussdiv_New = SqlStmt
    
End Function

Public Function GetAccCnt_Bussdiv_New1(ByVal pStart As String, ByVal pEnd As String, ByVal InOut As String, ByVal StatFg As String, ByVal StatPM As String, ByVal StatEM As String)
'FROM LAB201, LAB302, LAB001, HIS003

    Dim SqlStmt As String
        
    '일반검사
    SqlStmt = "SELECT a.workarea, b.testcd, d.testnm, c.tats, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & T_LAB004 & " c, " & T_LAB001 & " d, " & T_LAB302 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("b.vfydt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("b.vfydt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND a.spccd = c.spccd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd and applydt<=a.rcvdt) "
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If
    
    Select Case InOut
        Case 1
            SqlStmt = SqlStmt & " AND (a.wardid <>'' or wardid is not null) "
            SqlStmt = SqlStmt & " AND a.deptcd <> 'EM'"
    End Select
    
    Select Case StatFg
        Case 1
            SqlStmt = SqlStmt & " AND a.statfg='1'"
    End Select
    
    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
'        SqlStmt = SqlStmt & " AND (a.wardid <>'' or wardid is not null) "
        SqlStmt = SqlStmt & " AND (a.wardid ='' or wardid is null) "
        SqlStmt = SqlStmt & " AND a.deptcd <> 'EM'"
    End If
    
    If StatEM = 1 Then
        SqlStmt = SqlStmt & " AND a.deptcd = 'EM'"
    End If
    
    SqlStmt = SqlStmt & " AND (tats = '' OR tats IS not null)  AND d.TESTCD = c.TESTCD AND b.SPCCD = c.SPCCD AND c.tats <> ';'        AND c.tats <> ';;' AND c.tats <> ';;;'  "
    SqlStmt = SqlStmt & " GROUP BY a.workarea, b.testcd, d.testnm, c.tats "
    '기타검사
    SqlStmt = SqlStmt & "UNION ALL SELECT a.workarea, b.testcd, d.testnm, c.tats, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & T_LAB004 & " c, " & T_LAB001 & " d, " & T_LAB351 & " b, " & T_LAB201 & " a "
    SqlStmt = SqlStmt & "WHERE " & DBW("b.vfydt>=", Format(pStart, CS_DateDbFormat)) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("b.vfydt<=", Format(pEnd, CS_DateDbFormat)) & " "
    
    If mvarWorkarea <> "" Then
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
    End If
    
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND a.spccd = c.spccd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd and applydt<=a.rcvdt) "
    If P_ItemSeqFG = True Then
        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
    Else
        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
        SqlStmt = SqlStmt & " AND a.deptcd <> 'EM'"
    End If
    
    Select Case InOut
        Case 1
            SqlStmt = SqlStmt & " AND (a.wardid <>'' or wardid is not null) "
    End Select
    
    Select Case StatFg
        Case 1
            SqlStmt = SqlStmt & " AND a.statfg='1'"
    End Select
    
    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
'        SqlStmt = SqlStmt & " AND (a.wardid <>'' or wardid is not null) "
        SqlStmt = SqlStmt & " AND (a.wardid ='' or wardid is null) "
        SqlStmt = SqlStmt & " AND a.deptcd <> 'EM'"
    End If
    
    If StatEM = 1 Then
        SqlStmt = SqlStmt & " AND a.deptcd = 'EM'"
    End If
    
    SqlStmt = SqlStmt & " AND (tats = '' OR tats IS not null)  AND d.TESTCD = c.TESTCD AND c.tats <> ';' AND c.tats <> ';;'  AND c.tats <> ';;;' "
    SqlStmt = SqlStmt & " GROUP BY a.workarea, b.testcd, d.testnm, c.tats "
    '미생물검사
'    SqlStmt = SqlStmt & "UNION ALL SELECT a.workarea, b.testcd, d.testnm, c.tats, count(*) as Cnt "
'    SqlStmt = SqlStmt & "FROM  " & T_LAB004 & " c, " & T_LAB001 & " d, " & T_LAB404 & " b, " & T_LAB201 & " a "
'    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, CS_DateDbFormat)) & " "
'    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, CS_DateDbFormat)) & " "
'
'    If mvarWorkarea <> "" Then
'        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea>=", mvarWorkarea)
'        SqlStmt = SqlStmt & "AND   " & DBW("a.workarea<=", mvarWorkarea)
'    End If
'
'    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
'    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
'    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
'    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
'    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
'                                          "WHERE  testcd = d.testcd and applydt<=a.rcvdt) "
'
'    If P_ItemSeqFG = True Then
'        SqlStmt = SqlStmt & "AND (d.itemseq>0) "
'    Else
'        SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
'    End If
'
'    Select Case StatFg
'        Case 1
'            SqlStmt = SqlStmt & " AND c.statfg='1'"
'        Case 2
'            SqlStmt = SqlStmt & " AND c.statfg='0'"
'    End Select
'
'    If StatPM = 1 Then
'        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
'    End If
'
'    SqlStmt = SqlStmt & " AND (tats = '' OR tats IS not null)  AND d.TESTCD = c.TESTCD AND c.tats <> ';'        AND c.tats <> ';;'  "
'    SqlStmt = SqlStmt & " GROUP BY a.workarea, b.testcd, d.testnm, c.tats "
    SqlStmt = SqlStmt & " ORDER BY workarea, testcd, testnm "
    
    GetAccCnt_Bussdiv_New1 = SqlStmt
    
End Function
Public Function GetAccCnt_ResultTime(ByVal pStart As String, ByVal pEnd As String, ByVal strTestCd As String, ByVal StatFg As String, ByVal StatPM As String)
    Dim SQL As String
        
    SQL = "SELECT a.vfydt, a.vfytm, b.TATS, c.RCVDT, c.RCVTM, c.PTID, c.workarea, c.ACCDT, c.ACCSEQ, a.TestCd, c.wardid, b.statfg FROM S2LAB302 a, S2LAB004 b, S2LAB201 c"
    SQL = SQL + " WHERE a.vfydt BETWEEN '" & pStart & "' AND '" & pEnd & "'                         "
    SQL = SQL + "   AND a.testcd = '" & strTestCd & "'                                              "
    SQL = SQL + "   AND a.TESTCD = b.TESTCD                                                                "
    SQL = SQL + "   AND a.SPCCD = b.SPCCD                                                                  "
    SQL = SQL + "   AND a.WORKAREA = c.WORKAREA                                                            "
    SQL = SQL + "   AND a.ACCDT = c.ACCDT                                                                  "
    SQL = SQL + "   AND a.ACCSEQ = c.ACCSEQ                                                                "
    SQL = SQL + "   AND (b.TATS <> '' OR b.TATS IS not null)                                               "
    SQL = SQL + "   AND b.TATS <> ';'                                                                      "
    SQL = SQL + "   AND b.TATS <> ';;'                                                                      "
    
    Select Case StatFg
        Case 1
            SQL = SQL & " AND b.statfg='1'"
        Case 2
            SQL = SQL & " AND b.statfg='0'"
    End Select
    
'    If StatPM = 1 Then
'        SQL = SQL & " AND c.deptcd <> 'PM' AND c.deptcd <> 'HC'"
'    End If

    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
        SQL = SQL & " AND (c.wardid <>'' or wardid is not null)"
    End If

    GetAccCnt_ResultTime = SQL
    
End Function

Public Function GetAccCnt_ResultTime1(ByVal pStart As String, ByVal pEnd As String, ByVal strTestCd As String, ByVal StatFg As String, ByVal StatPM As String, ByVal StatEM As String)
    Dim SQL As String
        
    SQL = "SELECT a.vfydt, a.vfytm, b.TATS, c.RCVDT, c.RCVTM, c.PTID, c.workarea, c.ACCDT, c.ACCSEQ, a.TestCd, c.wardid, c.statfg FROM S2LAB302 a, S2LAB004 b, S2LAB201 c"
    SQL = SQL + " WHERE a.vfydt BETWEEN '" & pStart & "' AND '" & pEnd & "'                         "
    SQL = SQL + "   AND a.testcd = '" & strTestCd & "'                                              "
    SQL = SQL + "   AND a.TESTCD = b.TESTCD                                                                "
    SQL = SQL + "   AND a.SPCCD = b.SPCCD                                                                  "
    SQL = SQL + "   AND a.WORKAREA = c.WORKAREA                                                            "
    SQL = SQL + "   AND a.ACCDT = c.ACCDT                                                                  "
    SQL = SQL + "   AND a.ACCSEQ = c.ACCSEQ                                                                "
    SQL = SQL + "   AND (b.TATS <> '' OR b.TATS IS not null)                                               "
    SQL = SQL + "   AND b.TATS <> ';'                                                                      "
    SQL = SQL + "   AND b.TATS <> ';;'                                                                      "
    
    Select Case StatFg
        Case 1
            SQL = SQL & " AND c.statfg='1'"
    End Select
    
    If StatEM = 1 Then
        SQL = SQL & " AND c.deptcd = 'EM'"
    End If

    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
        SQL = SQL & " AND (c.wardid <>'' or wardid is not null)"
    End If

    GetAccCnt_ResultTime1 = SQL
    
End Function

Public Function GetAccCnt_ResultTime2(ByVal pStart As String, ByVal pEnd As String, ByVal strTestCd As String, ByVal sInout As String, ByVal StatFg As String, ByVal StatPM As String, ByVal StatEM As String)
    Dim SQL As String
        
    SQL = "SELECT a.vfydt, a.vfytm, b.TATS, c.RCVDT, c.RCVTM, c.PTID, c.workarea, c.COLDT, c.COLTM, c.COLID, c.ACCDT, c.ACCSEQ, a.TestCd, c.DEPTCD, c.wardid, c.statfg FROM S2LAB302 a, S2LAB004 b, S2LAB201 c"
    SQL = SQL + " WHERE a.vfydt BETWEEN '" & pStart & "' AND '" & pEnd & "'                         "
    SQL = SQL + "   AND a.testcd = '" & strTestCd & "'                                              "
    SQL = SQL + "   AND a.TESTCD = b.TESTCD                                                                "
    SQL = SQL + "   AND a.SPCCD = b.SPCCD                                                                  "
    SQL = SQL + "   AND a.WORKAREA = c.WORKAREA                                                            "
    SQL = SQL + "   AND a.ACCDT = c.ACCDT                                                                  "
    SQL = SQL + "   AND a.ACCSEQ = c.ACCSEQ                                                                "
    SQL = SQL + "   AND (b.TATS <> '' OR b.TATS IS not null)                                               "
    SQL = SQL + "   AND b.TATS <> ';'                                                                      "
    SQL = SQL + "   AND b.TATS <> ';;'                                                                     "
    SQL = SQL + "   AND b.TATS <> ';;;'                                                                    "
    
    Select Case sInout
        Case 1
            SQL = SQL & " AND (c.wardid <>'' or wardid is not null)"
            SQL = SQL & " AND c.deptcd <> 'EM'"
    End Select
    
    Select Case StatFg
        Case 1
            SQL = SQL & " AND c.statfg='1'"
    End Select
    
    If StatEM = 1 Then
        SQL = SQL & " AND c.deptcd = 'EM'"
    End If

    If StatPM = 1 Then
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM' AND a.deptcd <> 'HC'"
''        SqlStmt = SqlStmt & " AND a.deptcd <> 'PM'"
        SQL = SQL & " AND (c.wardid ='' or wardid is null)"
        SQL = SQL & " AND c.deptcd <> 'EM'"
    End If

    GetAccCnt_ResultTime2 = SQL
    
End Function
