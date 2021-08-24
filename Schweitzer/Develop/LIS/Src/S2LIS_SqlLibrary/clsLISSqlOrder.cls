VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'%  8. Get High Frequency Test List
'%      - intDiv : 1 -  (다빈도처방 리스트), 2 - (전체 처방 리스트)
'%      - Calling FROM [modOrder]
'%      - 검사명,검사코드,검체코드,응급여부,WorkArea,보관구분,아침채혈여부,
'%         검사구분,복수검체여부,접수번호부여기준,검체군

Public Function SqlItemList(ByVal intDiv As Integer) As String
    If intDiv = 1 Then
        'SqlItemList = "SELECT a.testnm, a.abbrnm5, b.cdval1 as testcd, b.field1 as spccd, c.statfg, a.workarea, " & _
                    "          c.storecd, c.rndfg, c.labelcnt, c.statflags, a.testdiv, d.field1 as MultiFg, d.field2 as SpcGrp "
        SqlItemList = " SELECT a.testnm, b.cdval1 as testcd " & _
                      " FROM  " & T_LAB001 & " a, " & T_LAB032 & " b " & _
                      " WHERE " & DBW("b.cdindex", LC3_HighItem, 2) & _
                      " AND   b.cdval1 = a.testcd " & _
                      " AND   a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                      "                     WHERE testcd = a.testcd ) " & _
                      " AND   (a.detailfg is null or " & DBW("a.detailfg", "", 2) & " ) "
                      
    Else
        SqlItemList = " SELECT a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
                      "          b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg, c.field2 as SpcGrp " & _
                      " FROM  " & T_LAB001 & " a, " & T_LAB004 & " b, " & T_LAB032 & " c " & _
                      " WHERE a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                      "                              WHERE testcd = a.testcd ) " & _
                      " AND  (a.detailfg = ''  or  a.detailfg is null) " & _
                      " AND   a.testcd = b.testcd " & _
                      " AND   " & DBW("b.seq", "1", 2) & _
                      " AND  (b.expdt = ''  or  b.expdt is null)" & _
                      " AND   " & DBW("c.cdindex", LC3_Specimen, 2) & _
                      " AND   c.cdval1 = b.spccd  "
    End If
End Function


'%  13. Get Specimen List
'%      - Calling FROM [frm101Order]
'%      - 검체코드,검체명,응급여부,보관구분,복수검체여부,검체군,우선순위
Public Function SqlSpecList(ByVal ParaTestCd As String) As String
    
    SqlSpecList = " SELECT a.spccd, b.field3 as SpcNm, a.statfg, a.storecd, a.rndfg, a.labelcnt, a.statflags, " & _
                  "        b.field1 as MultiFg, b.field2 as SpcGrp, b.field5 as SpcAbbr, a.seq,  " & _
                  "        c.field2 as LabDiv, d.field2 as LabRange, '1' InsurFg " & _
                  " FROM  " & T_LAB004 & " a, " & T_LAB032 & " b ," & T_LAB032 & " c, " & T_LAB032 & " d, " & T_LAB001 & " e " & _
                  " WHERE " & DBW("a.testcd", ParaTestCd, 2) & _
                  " AND   a.applydt = ( SELECT max(applydt) FROM " & T_LAB004 & _
                  "                     WHERE  testcd = a.testcd  " & _
                  "                     AND    spccd  = a.spccd ) " & _
                  " AND   (a.expdt is null  or " & DBW("a.expdt", "", 2) & " ) " & _
                  " AND   " & DBW("b.cdindex", LC3_Specimen, 2) & _
                  " AND   b.cdval1  = a.spccd " & _
                  " AND   " & DBW("e.testcd", ParaTestCd, 2) & _
                  " AND   e.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                  "                     WHERE  testcd = e.testcd ) " & _
                  " AND   " & DBW("c.cdindex", lc3_workarea, 2) & _
                  " AND   c.cdval1  = e.workarea " & _
                  " AND   " & DBJ(DBW("d.cdindex", LC3_SGroup, 2)) & _
                  " AND   " & DBJ("d.cdval1  =* b.field2") & _
                  " Order by a.seq "

End Function

'%  14. 직원정보 검색
'%      - Calling FROM [frm101Order]
Public Function SqlLAB015Read(ByVal paraEmpId As String, ByVal Index As Integer) As String
    If Index = 1 Then
        SqlLAB015Read = " SELECT * " & _
                        " FROM  " & T_LAB015 & " a " & _
                        " WHERE " & DBW("a.logonid", paraEmpId, 2)
    Else
        SqlLAB015Read = " SELECT * " & _
                        " FROM  " & T_LAB015 & " a " & _
                        " WHERE " & DBW("a.empid", paraEmpId, 2)
    End If
End Function

'%  15. 가장 마지막 처방번호
'%      - Calling FROM [frm101Order]
Public Function SqlLastOrdNo(ByVal paraPtId As String, ByVal paraDate As String) As String
    
    SqlLastOrdNo = " SELECT ordno " & _
                   " FROM  " & T_LAB101 & _
                   " WHERE " & DBW("ptid", paraPtId, 2) & _
                   " AND   " & DBW("orddt", paraDate, 2) & _
                   " Order by ordno desc"
                   
'    If objsysinfo.dbtype = 0 Then 'Oracle
''    If dbconn.Whatsthis = dbconn.ThisIsOracle Then
'        SqlLastOrdNo = SqlLastOrdNo & " for update "
'    End If
    
End Function


Public Function CreateSqlH(ByVal objOrder As Object, ByRef OrdH As OrderHeader) As String
    Dim strOrdFg    As String
    
    '임상병리 백업처방은
    strOrdFg = "1"
    With objOrder
        CreateSqlH = "Insert into " & T_LAB101 & _
                     "      (ptid, orddt, ordtm, ordno, bussdiv, bedindt, reqdt, reqtm, " & _
                     "       deptcd, orddoct, majdoct, entid, entdt, enttm, orddiv, repeatfg, " & _
                     "       orgaccno, sporddiv, donefg, receptno, wardid, roomid,ordfg, hosilid) " & _
                     "Values (" _
                            & DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & DBV("ordtm", .OrdTm, 1) _
                            & DBV("ordno", OrdH.OrdNo, 1) & DBV("bussdiv", .Bussdiv, 1) & DBV("bedindt", .Bedindt, 1) & DBV("reqdt", OrdH.ReqDt, 1) _
                            & DBV("reqtm", OrdH.ReqTm, 1) & DBV("deptcd", .DeptCd, 1) & DBV("orddoct", .OrdDoct, 1) & DBV("majdoct", .MajDoct, 1) & DBV("entid", .EntId, 1) _
                            & DBV("entdt", .EntDt, 1) & DBV("enttm", .EntTm, 1) _
                            & DBV("orddiv", .ORDDIV, 1) & DBV("repeatfg", .RepeatFg, 1) & DBV("orgaccno", .OrgAccNo, 1) & DBV("sporddiv", .SpOrdDiv, 1) & DBV("donefg", .DoneFg, 1) _
                            & DBV("receptno", .ReceptNo, 1) & DBV("wardid", .WardId, 1) & DBV("roomid", .ROOMID, 1) _
                            & DBV("ordfg", strOrdFg, 1) & DBV("hosilid", .HosilId) & " ) "
    End With

End Function

Public Function CreateSqlB(ByVal objOrder As Object, ByRef OrdH As OrderHeader, ByVal BSeq As Long) As String
    Dim strPaydt As String
    '임상병리과 백업처방은
    strPaydt = "0"
    With OrdH.OrdBody(BSeq)
        CreateSqlB = "Insert into " & T_LAB102 & _
                     "      (ptid, orddt, ordno, ordseq, ordcd, spccd, storecd, dcfg, dcdt, dcno, " & _
                     "       attrcd, examdt, examtm, examdoct, stscd, statfg, insdiv,paydt, donefg) " & _
                     "Values (" _
                            & DBV("ptid", objOrder.PtId, 1) & DBV("orddt", objOrder.OrdDt, 1) & DBV("ordno", OrdH.OrdNo, 1) & DBV("ordseq", BSeq, 1) & DBV("ordcd", .OrdCd, 1) _
                            & DBV("spccd", .SpcCd, 1) & DBV("storecd", .StoreCd, 1) & DBV("dcfg", .DcFg, 1) & DBV("dcdt", .DcDt, 1) & DBV("dcno", .DcNo, 1) & DBV("attrcd", .AttrCd, 1) _
                            & DBV("examdt", .ExamDt, 1) & DBV("examtm", .ExamTm, 1) & DBV("examdoct", .ExamDoct, 1) & DBV("stscd", .StsCd, 1) & DBV("statfg", .StatFg, 1) _
                            & DBV("insdiv", .InsDiv, 1) & DBV("paydt", strPaydt, 1) & DBV("donefg", .DoneFg) & " ) "
    End With
    
End Function

Public Function CreateSqlHeader(ByVal objOrder As Object, ByVal OrdH As Object) As String
    Dim strOrdFg    As String
    
    '임상병리 백업처방은
    strOrdFg = "1"
    
    With objOrder
'        CreateSqlHeader = "Insert into " & T_LAB101 & _
'                        "      (ptid, orddt, ordtm, ordno, bussdiv, bedindt, reqdt, reqtm, " & _
'                        "       deptcd, orddoct, majdoct, entid, entdt, enttm, orddiv, repeatfg, " & _
'                        "       orgaccno, sporddiv, donefg, receptno, wardid, roomid,ordfg, hosilid) " & _
'                        "Values (" _
'                               & DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & DBV("ordtm", .OrdTm, 1) _
'                               & DBV("ordno", OrdH.Fields("OrdNo"), 1) & DBV("bussdiv", .Bussdiv, 1) & DBV("bedindt", .Bedindt, 1) & DBV("reqdt", OrdH.Fields("ReqDt"), 1) _
'                               & DBV("reqtm", OrdH.Fields("ReqTm"), 1) & DBV("deptcd", .DeptCd, 1) & DBV("orddoct", .OrdDoct, 1) & DBV("majdoct", .MajDoct, 1) & DBV("entid", .EntId, 1) _
'                               & DBV("entdt", .EntDt, 1) & DBV("enttm", .EntTm, 1) _
'                               & DBV("orddiv", .ORDDIV, 1) & DBV("repeatfg", .RepeatFg, 1) & DBV("orgaccno", .OrgAccNo, 1) & DBV("sporddiv", .SpOrdDiv, 1) & DBV("donefg", .DoneFg, 1) _
'                               & DBV("receptno", .ReceptNo, 1) & DBV("wardid", .WardId, 1) & DBV("roomid", .ROOMID, 1) _
'                               & DBV("ordfg", strOrdFg, 1) & DBV("hosilid", .HosilId) & " ) "
'    CreateSqlHeader = " insert into oram1.mdexmort (patno,orddate,ordseqno,ordtime,patsect,meddate,hopedate, " & _
'                      " meddept,orddr,chadr,editid,editdate,wardno,roomno,patsite, ordgrp,slipcd,ordtype,ordkind,ordcd ) values (" & _
'                      "'" & .PtId & "'," & "to_date(" & .OrdDt & ",'yyyy-mm-dd'), " & OrdH.Fields("OrdNo") & ",to_date(" & .OrdDt & .OrdTm & ",'yyyy-mm-dd hh24:mi:ss'),'" & IIf(.Bussdiv = "1", "O", "I") & "'," & IIf(.Bedindt = "", "''", "to_date(" & .Bedindt & ",'yyyy-mm-dd')") & _
'                      "," & "to_date(" & OrdH.Fields("ReqDt") & OrdH.Fields("Reqtm") & ",'yyyy-mm-dd hh24:mi:ss') " & _
'                      ",'" & .DeptCd & "','" & .OrdDoct & "','" & .MajDoct & "','" & .EntId & "'," & "to_date(" & .OrdDt & .OrdTm & ",'yyyy-mm-dd hh24:mi:ss') , " & _
'                      "'" & .WardId & "','" & .ROOMID & "','Y','C1','L','1','1','1')"
    
    CreateSqlHeader = " insert into S2ORD999 (patno,orddate,ordseqno,ordtime,patsect,meddate,hopedate, " & _
                      " meddept,orddr,chadr,editid,editdate,wardno,roomno,patsite, ordgrp,slipcd,ordtype,ordkind,ordcd ) values (" & _
                      "'" & .PtId & "'," & "to_date(" & .OrdDt & ",'yyyy-mm-dd'), " & OrdH.Fields("OrdNo") & ",to_date(" & .OrdDt & .OrdTm & ",'yyyy-mm-dd hh24:mi:ss'),'" & IIf(.Bussdiv = "1", "O", "I") & "'," & IIf(.Bedindt = "", "''", "to_date(" & .Bedindt & ",'yyyy-mm-dd')") & _
                      "," & "to_date(" & OrdH.Fields("ReqDt") & OrdH.Fields("Reqtm") & ",'yyyy-mm-dd hh24:mi:ss') " & _
                      ",'" & .DeptCd & "','" & .OrdDoct & "','" & .MajDoct & "','" & .EntId & "'," & "to_date(" & .OrdDt & .OrdTm & ",'yyyy-mm-dd hh24:mi:ss') , " & _
                      "'" & .WardId & "','" & .ROOMID & "','Y','C1','L','1','1','1')"
    End With

End Function

Public Function CreateSqlBody(ByVal objOrder As Object, ByVal OrdB As Object) As String
    Dim strPaydt As String
    '임상병리과 백업처방은
    strPaydt = "0"
    
    With OrdB
'        CreateSqlBody = "Insert into " & T_LAB102 & _
'                        "      (ptid, orddt, ordno, ordseq, ordcd, spccd, storecd, dcfg, dcdt, dcno, " & _
'                        "       attrcd, examdt, examtm, examdoct, stscd, statfg, insdiv,paydt, donefg) " & _
'                        "Values (" _
'                               & DBV("ptid", objOrder.PtId, 1) & DBV("orddt", objOrder.OrdDt, 1) & DBV("ordno", OrdB.Fields("OrdNo"), 1) _
'                               & DBV("ordseq", OrdB.Fields("BSeq"), 1) & DBV("ordcd", OrdB.Fields("OrdCd"), 1) _
'                               & DBV("spccd", OrdB.Fields("SpcCd"), 1) & DBV("storecd", OrdB.Fields("StoreCd"), 1) _
'                               & DBV("dcfg", OrdB.Fields("DcFg"), 1) & DBV("dcdt", OrdB.Fields("DcDt"), 1) _
'                               & DBV("dcno", OrdB.Fields("DcNo"), 1) & DBV("attrcd", OrdB.Fields("AttrCd"), 1) _
'                               & DBV("examdt", OrdB.Fields("ExamDt"), 1) & DBV("examtm", OrdB.Fields("ExamTm"), 1) _
'                               & DBV("examdoct", OrdB.Fields("ExamDoct"), 1) & DBV("stscd", OrdB.Fields("StsCd"), 1) _
'                               & DBV("statfg", OrdB.Fields("StatFg"), 1) _
'                               & DBV("insdiv", OrdB.Fields("InsDiv"), 1) & DBV("paydt", strPaydt, 1) & DBV("donefg", OrdB.Fields("DoneFg")) & " ) "
'        CreateSqlBody = " update oram1.mdexmort set ordcd='" & .Fields("OrdCd") & "', spccode1='" & .Fields("SpcCd") & "', storecd='" & .Fields("StoreCd") & "',stscd='0',statfg='" & IIf(.Fields("StatFg") = "1", "Y", "") & "', rcpdate=to_date(" & objOrder.OrdDt & ",'yyyy-mm-dd')" & ",donefg='0'" & _
'                        " where patno='" & objOrder.PtId & "'" & _
'                        " and orddate=to_date(" & objOrder.OrdDt & ",'yyyy-mm-dd')" & _
'                        " and ordseqno=" & .Fields("OrdNo")
    
        CreateSqlBody = " update S2ORD999 set ordcd='" & .Fields("OrdCd") & "', spccode1='" & .Fields("SpcCd") & "', storecd='" & .Fields("StoreCd") & "',stscd='0',statfg='" & IIf(.Fields("StatFg") = "1", "Y", "") & "', rcpdate=to_date(" & objOrder.OrdDt & ",'yyyy-mm-dd')" & ",donefg='0'" & _
                        " where patno='" & objOrder.PtId & "'" & _
                        " and orddate=to_date(" & objOrder.OrdDt & ",'yyyy-mm-dd')" & _
                        " and ordseqno=" & .Fields("OrdNo")
    End With
    
End Function
