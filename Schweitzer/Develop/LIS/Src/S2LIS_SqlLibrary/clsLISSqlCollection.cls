VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#Const APS_COLLECT = True
'#Const BBS_COLLECT = True

Public Function SqlGetWardColSeq(ByVal pWorkDt As String, ByVal pWardId As String, ByVal pWorkTm As String)

    SqlGetWardColSeq = "SELECT max(seq) as ColSeq FROM " & T_LAB204 & _
                       " WHERE  " & DBW("workdt =", pWorkDt) & _
                       " AND    " & DBW("wardid =", pWardId) & _
                       " AND    " & DBW("worktm =", pWorkTm)

End Function


'%  42. 복수검체
'%       - Calling FROM [ clsCollection ] :
Public Function SqlGetMultiSpc(ByVal pSpcCd As String) As String
   
   SqlGetMultiSpc = "SELECT a.cdval2 as SpcSeq, a.field1 as SpcCd, b.field5 as BarNm  " & _
                    " FROM   " & T_LAB031 & " a, " & T_LAB032 & " b " & _
                    " WHERE  " & DBW("a.cdindex =", LC2_MultiSpc) & _
                    " AND    " & DBW("a.cdval1 =", pSpcCd) & _
                    " AND    " & DBW("b.cdindex =", LC3_Specimen) & _
                    " AND    b.cdval1 = a.field1 " & _
                    " order by SpcSeq "

End Function

'%  42-1. 복수검체 갯수
'%       - Calling FROM [ clsCollection ] :
Public Function SqlMultiSpcCnt1(ByVal SpcCd As String) As String

   SqlMultiSpcCnt1 = " SELECT Count(*) as Count " & _
                     " FROM   " & T_LAB031 & _
                     " WHERE  " & DBW("cdindex =", LC2_MultiSpc) & _
                     " AND    " & DBW("cdval1 =", SpcCd)

End Function

'%  42-2. 이미 생성된 복수검체 갯수
'%       - Calling FROM [ clsCollection ] :
Public Function SqlMultiSpcCnt2(ByVal PtId As String, ByVal OrdDt As String, _
                                ByVal OrdNo As Integer, ByVal OrdSeq As Integer) As String

    SqlMultiSpcCnt2 = " SELECT * " & _
                      " FROM    " & T_LAB203 & " " & _
                      " WHERE   " & DBW("ptid", PtId, 2) & _
                      " AND     " & DBW("orddt", OrdDt, 2) & _
                      " AND     " & DBW("ordno", OrdNo, 2) & _
                      " AND     " & DBW("ordseq", OrdSeq, 2)

End Function
   
'%  29. 병동 일괄 채혈을 위해 병동별로 처방내역을 검색한다. (날짜만 비교, 12.29)
'%       - Calling FROM [frm151Round]
Public Function SqlWardOrder(ByVal ReqDt As String, ByVal ReqTm As String, ByVal WardId As String) As String
   
   '2001/06/27
   '혈액은행 irradiation 처방은 제외한다..
   'dbw("b.orddiv<>","Z") 추가..
    SqlWardOrder = " SELECT  b.wardid as WardId, a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                             F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                             F_DOB2("a") & " as Dob, b.bedindt, b.deptcd, " & _
                    "        b.orddoct, b.majdoct, b.roomid, '' as bedid, b.hosilid, b.orddiv  " & _
                    " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                    " WHERE  " & DBW("b.wardid", WardId, 2) & _
                    " AND    " & DBW("b.donefg", "0", 2) & _
                    " AND    b.reqdt||b.reqtm<='" & ReqDt & ReqTm & "'" & _
                    " AND    " & DBW("b.orddiv<>", "Z") & _
                    " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                    " AND    a." & F_PTID & " = b.ptid " & _
                    " AND    c.ptid = b.ptid " & _
                    " AND   c.orddt = b.orddt " & _
                    " AND   c.ordno = b.ordno  " & _
                    " AND    " & DBW("c.stscd =", enStsCd.StsCd_LIS_Order) & _
                    " AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                    " Order  By WardId,hosilid, PtId, c.statfg, c.spccd "
End Function

Public Function SqlOutOrder(ByVal ReqDt As String, ByVal ReqTm As String, ByVal DeptCd As String) As String
    Dim strtmp  As String
    
'    If P_PayDtUsed Then
'        strtmp = " AND (c.paydt<>'' or c.paydt is not null)"
'    End If
    
    SqlOutOrder = " SELECT  b.deptcd , a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                             F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                             F_DOB2("a") & " as Dob, d.field5 spcnm, " & _
                    "        b.orddoct, b.majdoct, b.orddiv,c.paydt, c.inout as inout  " & _
                    " FROM   " & T_LAB032 & " d, " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                    " WHERE  " & DBW("b.donefg", "0", 2) & _
                    " AND    " & DBW("b.reqdt <= ", ReqDt) & _
                    " AND    " & DBW("b.orddiv=", "L") & _
                    " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                    " AND    " & DBW("b.deptcd", DeptCd, 2) & _
                    " AND    a." & F_PTID & " = b.ptid " & _
                    " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & strtmp & _
                    " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                    " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
                    " AND    d.cdval1 = c.spccd " & _
                    " AND    (c.inout <> '' or c.inout is not null)"
    
    SqlOutOrder = SqlOutOrder & " UNION ALL " & _
                    " SELECT  b.deptcd , a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                              F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                              F_DOB2("a") & " as Dob, '혈액' as  spcnm, " & _
                    "         b.orddoct, b.majdoct, b.orddiv ,c.paydt, c.inout as inout " & _
                    " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                    " WHERE  " & DBW("b.donefg", "0", 2) & _
                    " AND    " & DBW("b.reqdt <= ", ReqDt) & _
                    " AND    " & DBW("b.orddiv=", "B") & _
                    " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                    " AND    " & DBW("b.deptcd", DeptCd, 2) & _
                    " AND    a." & F_PTID & " = b.ptid " & _
                    " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                    " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                    " AND    (c.inout <> '' or c.inout is not null)"
                    
    SqlOutOrder = SqlOutOrder & " Order  By inout, PtId, statfg, spccd "

End Function

'-------------------------------------------------------------
'** 예수병원 추가루틴
Public Function SqlOutOrder_New(ByVal ReqDt As String, ByVal ReqTm As String, _
                                ByVal DeptCd As String, ByVal pCorpCd As String) As String
    Dim strtmp  As String
    
    SqlOutOrder_New = " SELECT  b.deptcd , a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                             F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                             F_DOB2("a") & " as Dob, d.field5 spcnm, " & _
                    "        b.orddoct, b.majdoct, b.orddiv,c.paydt, c.inoutseq as inoutseq  " & _
                    " FROM   " & T_LAB032 & " d, " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                    " WHERE  " & DBW("b.donefg", "0", 2) & _
                    " AND    " & DBW("b.reqdt <= ", ReqDt) & _
                    " AND    " & DBW("b.orddiv=", "L") & _
                    " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                    " AND    " & DBW("b.deptcd", DeptCd, 2) & _
                    " AND    a." & F_PTID & " = b.ptid " & _
                    " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & strtmp & _
                    " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                    " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
                    " AND    d.cdval1 = c.spccd " & _
                    " and    c.inout = " & DBS(pCorpCd) & _
                    " AND    (c.inout <> '' or c.inout is not null)"
    

    SqlOutOrder_New = SqlOutOrder_New & " Order  By inoutseq, PtId, spccd "
End Function
'-------------------------------------------------------------


Public Function SqlOrderForMornCol(ByVal ReqDt As String, ByVal ReqTm As String, ByVal WardId As String) As String

' 임상병리 아침채혈
' 기준 : 작업시점을 기준으로 발생되어 있는 모든처방
'        임상병리일반검사 중 검체가 혈액인 처방 (임상병리 마스터에서 정의 : 검사-검체)
'        응급처방이 하나라도 포함되어 있으면 표시 --> 간호사채혈
'        비응급처방만 임상병리사가 채혈함

    SqlOrderForMornCol = " SELECT  b.wardid as WardId, a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                                   F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                                   F_DOB2("a") & " as Dob, b.bedindt, b.deptcd, " & _
                         "        b.orddoct, b.majdoct, b.roomid, '' as bedid, b.hosilid, b.orddiv  " & _
                         " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c, " & T_LAB004 & " d " & _
                         " WHERE  " & DBW("b.wardid", WardId, 2) & _
                         " AND    " & DBW("b.donefg", "0", 2) & _
                         " AND    " & DBW("b.reqdt < ", ReqDt) & _
                         " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                         " AND    " & DBW("b.orddiv", lis_orddiv, 2) & _
                         " AND    a." & F_PTID & " = b.ptid " & _
                         " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                         " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                         " AND    d.testcd = c.ordcd " & _
                         " AND    d.spccd  = c.spccd " & _
                         " AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                                             " WHERE testcd = d.testcd  AND  spccd = d.spccd)  " & _
                         " AND    " & DBW("c.statfg<>", "1") & _
                         " AND    d.rndfg = '1' "

    SqlOrderForMornCol = SqlOrderForMornCol & " UNION ALL " & _
                         " SELECT  b.wardid as WardId, a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                                   F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                                   F_DOB2("a") & " as Dob, b.bedindt, b.deptcd, " & _
                         "        b.orddoct, b.majdoct, b.roomid, '' as bedid, b.hosilid, b.orddiv  " & _
                         " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c, " & T_LAB004 & " d " & _
                         " WHERE  " & DBW("b.wardid", WardId, 2) & _
                         " AND    " & DBW("b.donefg", "0", 2) & _
                         " AND    " & DBW("b.reqdt = ", ReqDt) & _
                         " AND    " & DBW("b.reqtm <= ", ReqTm) & _
                         " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                         " AND    " & DBW("b.orddiv", lis_orddiv, 2) & _
                         " AND    a." & F_PTID & " = b.ptid " & _
                         " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                         " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                         " AND    d.testcd = c.ordcd " & _
                         " AND    d.spccd  = c.spccd " & _
                         " AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                                             " WHERE testcd = d.testcd  AND  spccd = d.spccd)  " & _
                         " AND    " & DBW("c.statfg<>", "1") & _
                         " AND    d.rndfg = '1' " & _
                         " Order  By WardId,hosilid, PtId, statfg, spccd "

End Function

Public Function CreateSql_Collection(ByVal objCol As Object) As String

    With objCol
        CreateSql_Collection = _
                    "Insert into " & T_LAB201 & _
                    "    (spcyy, spcno, ptid, sex, ageday, bedindt, deptcd, orddoct, majdoct, " & _
                    "     workarea, accdt, accseq, stscd, reqtotcnt, reqinputcnt, " & _
                    "     vfydt, vfytm, vfyid, coldt, coltm, colid, rcvdt, rcvtm, rcvid, " & _
                    "     entdt, enttm, entid, spccd, multifg, orgaccno, wardid, roomid, bedid, " & _
                    "     footnotefg, storecd, rptfg, testdiv, qcfg, rmkcd, statfg, buildcd, orgbuildcd, hosilid) " & _
                    "Values  (" & DBV("spcyy", .SpcYy, 1) & DBV("spcno", .SpcNo, 1) & DBV("ptid", .PtId, 1) _
                    & DBV("sex", .Sex, 1) & DBV("ageday", .AgeDay, 1) & DBV("bedindt", .Bedindt, 1) _
                    & DBV("deptcd", .DeptCd, 1) & DBV("orddoct", .OrdDoct, 1) & DBV("majdoct", .MajDoct, 1) _
                    & DBV("workarea", .WorkArea, 1) & DBV("accdt", .AccDt, 1) & DBV("accseq", .accseq, 1) _
                    & DBV("stscd", .StsCd, 1) & DBV("reqtotcnt", .ReqTotCnt, 1) & DBV("reqinputcnt", .ReqInputCnt, 1) _
                    & DBV("vfydt", .VfyDt, 1) & DBV("vfytm", .VfyTm, 1) & DBV("vfyid", .VfyID, 1) _
                    & DBV("coldt", .ColDt, 1) & DBV("coltm", .ColTm, 1) & DBV("colid", .ColID, 1) _
                    & DBV("rcvdt", .RcvDt, 1) & DBV("rcvtm", .RcvTm, 1) & DBV("rcvid", .RcvID, 1) _
                    & DBV("entdt", .EntDt, 1) & DBV("enttm", .EntTm, 1) & DBV("entid", .EntId, 1) _
                    & DBV("spccd", .SpcCd, 1) & DBV("multifg", .MultiFg, 1) & DBV("orgaccno", .OrgAccNo, 1) _
                    & DBV("wardid", .WardId, 1) & DBV("roomid", .ROOMID, 1) & DBV("bedid", .BedID, 1) _
                    & DBV("footnotefg", .FootNoteFg, 1) & DBV("storecd", .StoreCd, 1) & DBV("rptfg", .RptFg, 1) _
                    & DBV("testdiv", .TestDiv, 1) & DBV("qcfg", .QcFg, 1) & DBV("rmkcd", .RmkCd, 1) _
                    & DBV("statfg", .StatFg, 1) & DBV("buildcd", .BuildCd, 1) _
                    & DBV("orgbuildcd", .OrgBuildCd, 1) & DBV("hosilid", .HosilId) & ") "
    End With
    
End Function


Public Function CreateSql_MultiSpc(ByVal objCol As Object) As String

    With objCol
        CreateSql_MultiSpc = _
                    "Insert into " & T_LAB203 & " ( ptid, orddt, ordno, ordseq, spcseq, spccd, workarea, accdt, accseq ) " & _
                    "Values ( " & DBV("ptid", .PtId, 1) & DBV("orddt", .OrdDt, 1) & DBV("ordno", .OrdNo, 1) _
                                & DBV("ordseq", .OrdSeq, 1) & DBV("spcseq", .SpcSeq, 1) & DBV("spccd", .SpcCd, 1) _
                                & DBV("workarea", .WorkArea, 1) & DBV("accdt", .AccDt, 1) & DBV("accseq", .accseq) & " )"
    End With
    
End Function

Public Function CreateSql_WardCol(ByVal objCol As Object) As String

    With objCol
        CreateSql_WardCol = "Insert into " & T_LAB204 & " ( workdt, wardid, worktm, seq, workarea, accdt, accseq, " & _
                                         " colid, buildcd, orgbuildcd, spcyy, spcno, orddiv, mornfg ) " & _
                            "Values ( " & DBV("workdt", .WorkDt, 1) & DBV("wardid", .WardId, 1) _
                                        & DBV("worktm", .WorkTm, 1) & DBV("workseq", .WorkSeq, 1) _
                                        & DBV("workarea", .WorkArea, 1) & DBV("accdt", .AccDt, 1) _
                                        & DBV("accseq", .accseq, 1) & DBV("colid", .ColID, 1) _
                                        & DBV("buildcd", .BuildCd, 1) & DBV("orgbuildcd", .OrgBuildCd, 1) _
                                        & DBV("spcyy", .SpcYy, 1) & DBV("spcno", .SpcNo, 1) _
                                        & DBV("orddiv", lis_orddiv, 1) & DBV("mornfg", .MornFg) _
                                        & " )"
    End With
    
End Function

Public Function CreateSql_UpdateLabNo(ByVal objCol As Object, ByVal PtId As String, ByVal OrdDt As String, _
                                      ByVal OrdNo As Integer, ByVal OrdSeq As String, _
                                      ByVal StsCd As String, Optional ByVal DoneFg As Variant) As String
    Dim tmpStr As String
    
    '** 원본 ==========================================================================
'    If IsMissing(DoneFg) Then
'        tmpStr = " "
'    Else
'        tmpStr = ",  " & DBW("donefg", DoneFg, 2)
'    End If
'    With objCol
'        CreateSql_UpdateLabNo = " Update   " & T_LAB102 & " " & _
'                                " Set      " & DBW("workarea", .WorkArea, 3) & _
'                                "          " & DBW("accdt", .AccDt, 3) & _
'                                "          " & DBW("accseq", .AccSeq, 3) & _
'                                "          " & DBW("stscd", StsCd, 2) & tmpStr & _
'                                " WHERE    " & DBW("ptid", PtId, 2) & _
'                                " AND      " & DBW("orddt", OrdDt, 2) & _
'                                " AND      " & DBW("ordno", OrdNo, 2) & _
'                                " AND      " & DBW("ordseq", OrdSeq, 2) & _
'                                " AND      " & DBW("stscd <= ", StsCd)
'    End With
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(PtId) & _
             "    and orddt = " & DBS(OrdDt) & _
             "    and ordno = " & DBS(OrdNo)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strWrkDiv = "" & rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    rs.Close: Set rs = Nothing
    
    If IsMissing(DoneFg) Then
        tmpStr = " "
    Else
        tmpStr = ",  " & DBW("donefg", DoneFg, 2)
    End If
    
    With objCol
        If strWrkDiv = "3" Then         '종건
            '** 종건 Update
            CreateSql_UpdateLabNo = " Update   su2examt " & _
                                    " Set      " & DBW("workarea", .WorkArea, 3) & _
                                    "          " & DBW("accdt", .AccDt, 3) & _
                                    "          " & DBW("accseq", .accseq, 3) & _
                                    "          " & DBW("stscd", StsCd, 2) & tmpStr & _
                                    " WHERE    patno = " & DBS(PtId) & _
                                    " AND      orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd') " & _
                                    " AND      ordseqno = " & DBN(OrdNo) & _
                                    " AND      " & DBW("stscd <= ", StsCd)
            
        ElseIf strWrkDiv = "4" Then     '일건
                '** 일건 Update
                CreateSql_UpdateLabNo = " Update   sg2examt " & _
                                        " Set      " & DBW("workarea", .WorkArea, 3) & _
                                        "          " & DBW("accdt", .AccDt, 3) & _
                                        "          " & DBW("accseq", .accseq, 3) & _
                                        "          " & DBW("stscd", StsCd, 2) & tmpStr & _
                                        " WHERE    patno = " & DBS(PtId) & _
                                        " AND      orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd') " & _
                                        " AND      ordseqno = " & DBN(OrdNo) & _
                                        " AND      " & DBW("stscd <= ", StsCd)
        Else                            '이외검사
            CreateSql_UpdateLabNo = " Update  mdexmort " & _
                                    " Set      " & DBW("workarea", .WorkArea, 3) & _
                                    "          " & DBW("accdt", .AccDt, 3) & _
                                    "          " & DBW("accseq", .accseq, 3) & _
                                    "          " & DBW("stscd", StsCd, 2) & tmpStr & _
                                    " WHERE    patno = " & DBS(PtId) & _
                                    " AND      orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd') " & _
                                    " AND      ordseqno = " & DBN(OrdNo) & _
                                    " AND      " & DBW("stscd <= ", StsCd)
        End If
        
        If Mid(PtId, 1, 1) = "L" Then
            CreateSql_UpdateLabNo = " Update  S2ORD999 " & _
                                    " Set      " & DBW("workarea", .WorkArea, 3) & _
                                    "          " & DBW("accdt", .AccDt, 3) & _
                                    "          " & DBW("accseq", .accseq, 3) & _
                                    "          " & DBW("stscd", StsCd, 2) & tmpStr & _
                                    " WHERE    patno = " & DBS(PtId) & _
                                    " AND      orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd') " & _
                                    " AND      ordseqno = " & DBN(OrdNo) & _
                                    " AND      " & DBW("stscd <= ", StsCd)
        End If
    End With
    '===================================================================================
End Function

Public Function CreateSql_HeaderUpdate(ByVal PtId As String, ByVal OrdDt As String, ByVal OrdNo As Integer, _
                                       ByVal DoneFg As String, ByVal Flag As Integer)
    If Flag = 1 Then
        CreateSql_HeaderUpdate = " SELECT * FROM " & T_LAB102 & _
                                 " WHERE   " & DBW("ptid", PtId, 2) & _
                                 " AND     " & DBW("orddt", OrdDt, 2) & _
                                 " AND     " & DBW("ordno", OrdNo, 2) & _
                                 " AND   ( dcfg = '' or dcfg is null )  " & _
                                 " AND     " & DBW("donefg <", DoneFg)
    Else
        '** 원본 ==========================================================================
'        CreateSql_HeaderUpdate = " Update " & T_LAB101 & _
'                                 " Set     " & DBW("donefg", DoneFg, 2) & _
'                                 " WHERE   " & DBW("ptid", PtId, 2) & _
'                                 " AND     " & DBW("orddt", OrdDt, 2) & _
'                                 " AND     " & DBW("ordno", OrdNo, 2)
        '===================================================================================
        
        '** 전주 예수 병원 변경 루틴========================================================
        Dim strWrkDiv   As String
        Dim strSQL      As String
        Dim rs          As New ADODB.Recordset
        
        strSQL = " select wrkdiv from " & T_LAB102 & _
                 "  where ptid = " & DBS(PtId) & _
                 "    and orddt = " & DBS(OrdDt) & _
                 "    and ordno = " & DBS(OrdNo)
                 
        rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
        
        If rs.EOF = False Then
            strWrkDiv = "" & rs.Fields("wrkdiv").Value
        Else
            strWrkDiv = ""
        End If
        
        rs.Close: Set rs = Nothing
        
        If strWrkDiv = "3" Then         '종건
            '** 종건 Update
            CreateSql_HeaderUpdate = " Update su2examt " & _
                                     " Set     " & DBW("donefg", DoneFg, 2) & _
                                     " WHERE  patno = " & DBS(PtId) & _
                                     " AND    orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd')" & _
                                     " AND    ordseqno = " & DBN(OrdNo)
        ElseIf strWrkDiv = "4" Then     '일건
                '** 일건 Update
                CreateSql_HeaderUpdate = " Update sg2examt " & _
                                         " Set     " & DBW("donefg", DoneFg, 2) & _
                                         " WHERE  patno = " & DBS(PtId) & _
                                         " AND    orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd')" & _
                                         " AND    ordseqno = " & DBN(OrdNo)
        Else                            '이외검사
            CreateSql_HeaderUpdate = " Update mdexmort " & _
                                     " Set     " & DBW("donefg", DoneFg, 2) & _
                                     " WHERE  patno = " & DBS(PtId) & _
                                     " AND    orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd')" & _
                                     " AND    ordseqno = " & DBN(OrdNo)
        End If
        
        If Mid(PtId, 1, 1) = "L" Then
            CreateSql_HeaderUpdate = " Update s2ord999 " & _
                                     " Set     " & DBW("donefg", DoneFg, 2) & _
                                     " WHERE  patno = " & DBS(PtId) & _
                                     " AND    orddate = TO_DATE(" & DBS(OrdDt) & ", 'yyyymmdd')" & _
                                     " AND    ordseqno = " & DBN(OrdNo)
        
        End If
        '===================================================================================
   End If
End Function

Public Function CreateSql_SpcNo(ByVal pSpcYY As String, ByVal pFlag As Integer, _
                                Optional ByVal pSeq As Variant) As String
'** 추후 적용시점에는 아래 SQL문을 사용 (table : h7com099)
    Select Case pFlag
    Case 1:
        CreateSql_SpcNo = " SELECT seq " & _
                          " FROM   " & T_COM099 & _
                          " WHERE  " & DBW("noindex =", COM99_LIS_Specimen) & _
                          " AND    " & DBW("divcd1  =", lis_orddiv) & _
                          " AND    " & DBW("divcd2  =", pSpcYY) & _
                          " AND    " & DBW("divcd3  =", "0")
    Case 2:
        CreateSql_SpcNo = " Insert Into " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq) " & _
                          " Values (" & DBV("noindex", COM99_LIS_Specimen, 1) & DBV("divcd1", lis_orddiv, 1) _
                                      & DBV("divcd2", pSpcYY, 1) & DBV("divcd3", "0", 1) & DBV("seq", pSeq) & ")"
    Case 3:
        CreateSql_SpcNo = " Update " & T_COM099 & _
                          " Set    " & DBW("seq =", pSeq) & _
                          " WHERE  " & DBW("noindex =", COM99_LIS_Specimen) & _
                          " AND    " & DBW("divcd1  =", lis_orddiv) & _
                          " AND    " & DBW("divcd2  =", pSpcYY) & _
                          " AND    " & DBW("divcd3  =", "0")
    Case 4:
        CreateSql_SpcNo = " Update " & T_COM099 & _
                          " Set    seq = seq " & _
                          " WHERE  " & DBW("noindex =", COM99_LIS_Specimen) & _
                          " AND    " & DBW("divcd1  =", lis_orddiv) & _
                          " AND    " & DBW("divcd2  =", pSpcYY) & _
                          " AND    " & DBW("divcd3  =", "0")
    End Select

'** 길병원시스템을 교육용으로 병행하므로 h7lab201의 Dup에러를 방지하기 위해
'** 한시적으로 아래 SQL문을 사용 (table : h7lab099)
'   SELECT Case pFlag
'   Case 1:
'      CreateSql_SpcNo = " SELECT seq " & _
'                        " FROM   " & T_LAB099 & _
'                        " WHERE cdindex = " & DBStr("01") & _
'                        " AND   divcd1  = " & DBStr("SPCNO") & _
'                        " AND   divcd2  = " & DBStr(pSpcYY) & _
'                        " AND   divcd3  = " & DBStr("0")
'   Case 2:
'      CreateSql_SpcNo = " Insert Into " & T_LAB099 & " (cdindex, divcd1, divcd2, divcd3, seq) " & _
'                        " Values (" & DBStr("01", 1) & DBStr("SPCNO", 1) & DBStr(pSpcYY, 1) & DBStr("0", 1) & DBNum(pSeq) & ")"
'   Case 3:
'      CreateSql_SpcNo = " Update " & T_LAB099 & " Set seq = " & pSeq & _
'                        " WHERE cdindex = " & DBStr("01") & _
'                        " AND   divcd1  = " & DBStr("SPCNO") & _
'                        " AND   divcd2  = " & DBStr(pSpcYY) & _
'                        " AND   divcd3  = " & DBStr("0 ")
'   Case 4:
'      CreateSql_SpcNo = " Update " & T_LAB099 & " Set seq = seq " & _
'                        " WHERE cdindex = " & DBStr("01") & _
'                        " AND   divcd1  = " & DBStr("SPCNO") & _
'                        " AND   divcd2  = " & DBStr(pSpcYY) & _
'                        " AND   divcd3  = " & DBStr("0 ")
'   End SELECT

End Function


Public Function CreateSql_LabNo(ByVal WorkArea As String, ByVal AccDt As String, ByVal SpcGrp As String, _
                                ByVal Flag As Integer, Optional ByVal Seq As Variant) As String

'** 추후 적용시점에는 아래 SQL문을 사용 (table : h7com099)

    Select Case Flag
        Case 1:
           CreateSql_LabNo = " SELECT seq " & _
                             " FROM   " & T_COM099 & " " & _
                             " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                             " AND    " & DBW("divcd1 =", WorkArea) & _
                             " AND    " & DBW("divcd2 =", AccDt) & _
                             " AND    " & DBW("divcd3 =", SpcGrp)
        Case 2:
           CreateSql_LabNo = "Insert Into " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq) " & _
                             "Values (" & DBV("noindex", COM99_LIS_LabNo, 1) & DBV("divcd1", WorkArea, 1) _
                                        & DBV("divcd2", AccDt, 1) & DBV("divcd3", SpcGrp, 1) & DBV("seq", Seq) & ")"
        Case 3:
           CreateSql_LabNo = " Update " & T_COM099 & _
                             " Set    " & DBW("seq =", Seq) & _
                             " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                             " AND    " & DBW("divcd1 =", WorkArea) & _
                             " AND    " & DBW("divcd2 =", AccDt) & _
                             " AND    " & DBW("divcd3 =", SpcGrp)
        Case 4:
           CreateSql_LabNo = " Update " & T_COM099 & " Set seq = seq " & _
                             " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                             " AND    " & DBW("divcd1 =", WorkArea) & _
                             " AND    " & DBW("divcd2 =", AccDt) & _
                             " AND    " & DBW("divcd3 =", SpcGrp)
    End Select
    Debug.Print CreateSql_LabNo
End Function

Public Function CreateSql_SeqNo(ByVal WorkArea As String, ByVal RcvDt As String, ByVal SpcGrp As String, _
                                ByVal Flag As Integer, ByVal AccNo As String, Optional ByVal Seq As Variant) As String

    Select Case Flag
        Case 1:
           CreateSql_SeqNo = " SELECT seq " & _
                             " FROM   " & T_COM099 & " " & _
                             " WHERE  " & DBW("noindex =", COM99_LIS_SeqNo) & _
                             " AND    " & DBW("divcd1 =", WorkArea) & _
                             " AND    " & DBW("divcd2 =", RcvDt) & _
                             " AND    " & DBW("divcd3 =", SpcGrp)
        Case 2:
           CreateSql_SeqNo = "Insert Into " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq, remark) " & _
                             "Values (" & DBV("noindex", COM99_LIS_SeqNo, 1) & DBV("divcd1", WorkArea, 1) _
                                        & DBV("divcd2", RcvDt, 1) & DBV("divcd3", SpcGrp, 1) & DBV("seq", Seq, 1) _
                                        & DBV("remark", AccNo) & ")"
        Case 3:
           CreateSql_SeqNo = " Update " & T_COM099 & _
                             " Set    " & DBW("seq =", Seq) & _
                             " WHERE  " & DBW("noindex =", COM99_LIS_SeqNo) & _
                             " AND    " & DBW("divcd1 =", WorkArea) & _
                             " AND    " & DBW("divcd2 =", RcvDt) & _
                             " AND    " & DBW("divcd3 =", SpcGrp)
        Case 4:
           CreateSql_SeqNo = " Update " & T_COM099 & " Set seq = seq " & _
                             " WHERE  " & DBW("noindex =", COM99_LIS_SeqNo) & _
                             " AND    " & DBW("divcd1 =", WorkArea) & _
                             " AND    " & DBW("divcd2 =", RcvDt) & _
                             " AND    " & DBW("divcd3 =", SpcGrp)
                             
        Case 5:
            CreateSql_SeqNo = " SELECT max(seq) as seq " & _
                              "   FROM " & T_COM099 & _
                              "  WHERE " & DBW("noindex =", COM99_LIS_SeqNo) & _
                              "   AND  " & DBW("divcd1 =", WorkArea) & _
                              "   AND  " & DBW("divcd2 =", RcvDt) & _
                              " group by noindex, divcd1, divcd2 "
    End Select
    
End Function

Public Function SqlLoadOrderDate(ByVal pKeyValue As String, ByVal pFlag As Long) As String
' 2009.01.09 양성현 미래처방의 처리 현재일,REQDT, 처방의,주치의 추가

    If pFlag = 0 Then
        '전체처방
        SqlLoadOrderDate = " SELECT a.orddt, b.bedindt, b.orddoct, b.majdoct, to_char(SYSDATE,'YYYYMMDD') as today FROM " & T_LAB101 & " b," & T_LAB102 & " a " & _
                           " WHERE " & DBW("b.ptid =", pKeyValue) & _
                           " AND   " & DBW("b.donefg =", enStsCd.StsCd_LIS_Order) & _
                           " AND   " & DBW("b.bussdiv=", enBussDiv.BussDiv_OutPatient) & _
                           " AND      a.ptid=b.ptid" & _
                           " AND      a.orddt=b.orddt" & _
                           " AND      a.ordno=b.ordno" & _
                           " AND     (a.dcfg is null or a.dcfg = '')" & _
                           " GROUP BY a.orddt, b.bedindt, b.orddoct, b.majdoct" & _
                           " ORDER BY orddt desc"
                           
        '2014-01-27 실제 오더미존재하는 데이타까지 DisPlay 됨.
        'SqlLoadOrderDate = " SELECT a.orddt, b.bedindt, b.orddoct, b.majdoct, to_char(SYSDATE,'YYYYMMDD') as today FROM " & T_LAB101 & " b," & T_LAB102 & " a " & _
                           " WHERE " & DBW("b.ptid =", pKeyValue) & _
                           " AND   " & DBW("b.donefg =", enStsCd.StsCd_LIS_Order) & _
                           " AND   " & DBW("b.bussdiv=", enBussDiv.BussDiv_OutPatient) & _
                           " AND      a.ptid=b.ptid" & _
                           " AND      a.orddt=b.orddt" & _
                           " AND      a.ordno=b.ordno" & _
                           " AND     (a.dcfg is null or a.dcfg = '')" & _
                           " AND    ( (SELECT count(*) FROM s2lab001 WHERE testcd = a.ordcd AND applydt <= a.orddt) > 0 " & _
                           " OR       (SELECT count(*) FROM s2bbs001 WHERE testcd = a.ordcd AND applydt <= a.orddt) > 0 ) " & _
                           " GROUP BY a.orddt, b.bedindt, b.orddoct, b.majdoct" & _
                           " ORDER BY orddt desc"
    ElseIf pFlag = 1 Then
        '## 수납된 환자만
        
        '** By M.G.Choi 핵의학 처방으로 인해 처방일이 조회되는 경우 제외 함
        '   핵의학 처방상태의 판단은 접수일(RCVDT) 여부로 처리 한다.
        SqlLoadOrderDate = " SELECT a.orddt, b.bedindt, b.orddoct, b.majdoct, to_char(SYSDATE,'YYYYMMDD') as today FROM " & T_LAB101 & " b," & T_LAB102 & " a " & _
                           " WHERE " & DBW("b.ptid =", pKeyValue) & _
                           " AND   " & DBW("b.donefg =", enStsCd.StsCd_LIS_Order) & _
                           " AND   " & DBW("b.bussdiv=", enBussDiv.BussDiv_OutPatient) & _
                           " AND      a.ptid=b.ptid" & _
                           " AND      a.orddt=b.orddt" & _
                           " AND      a.ordno=b.ordno" & _
                           " AND (a.paydt<>'' or a.paydt is not null)" & _
                           " AND     (a.dcfg is null or a.dcfg = '')" & _
                           " AND (a.rcvdt is null or a.rcvdt = '')" & _
                           " GROUP BY a.orddt, b.bedindt, b.orddoct, b.majdoct" & _
                           " ORDER BY orddt desc"
                           
        '## 원본 ############################################################################
'        SqlLoadOrderDate = " SELECT a.orddt FROM " & T_LAB101 & " b," & T_LAB102 & " a " & _
'                           " WHERE " & DBW("b.ptid =", pKeyValue) & _
'                           " AND   " & DBW("b.donefg =", enStsCd.StsCd_LIS_Order) & _
'                           " AND   " & DBW("b.bussdiv=", enBussDiv.BussDiv_OutPatient) & _
'                           " AND      a.ptid=b.ptid" & _
'                           " AND      a.orddt=b.orddt" & _
'                           " AND      a.ordno=b.ordno" & _
'                           " AND (a.paydt<>'' or a.paydt is not null)" & _
'                           " AND     (a.dcfg is null or a.dcfg = '')" & _
'                           " GROUP BY a.orddt" & _
'                           " ORDER BY orddt desc"
        '####################################################################################
    Else
        SqlLoadOrderDate = " SELECT ptid, orddt, b.bedindt, b.orddoct, b.majdoct, to_char(SYSDATE,'YYYYMMDD') as today FROM " & T_LAB101 & _
                           " WHERE   " & DBW("receptno =", pKeyValue) & _
                           " group  by ptid, orddt, bedindt, orddoct, majdoct "
    End If
    
    Debug.Print SqlLoadOrderDate
End Function

Public Function SqlGetBatchDept() As String
    SqlGetBatchDept = " SELECT  a.cdval1 deptcd, b." & F_DEPTNM & _
                      " FROM  " & T_HIS003 & " b, " & T_LAB032 & " a " & _
                      " WHERE " & DBW("a.cdindex = ", LC3_BatchColDept) & _
                      " AND    trim( b." & F_DEPTCD & ") = a.cdval1"
End Function
   
Public Function SqlReadOutOrder(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                Optional ByVal DeptCd As String = "") As String

    Dim strtmp  As String
    
    If P_PayDtUsed Then
        strtmp = " AND (b.paydt<>'' or b.paydt is not null) "
    End If
    
    SqlReadOutOrder = "SELECT c.testcd, c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                    "        d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.paydt, " & _
                    "        b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                    "        e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, f.labelcnt, '' as ErrFg, '' wardid, '' roomid, '' bedid " & _
                    " FROM   " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_HIS005 & " e, " & _
                                 T_LAB004 & " f, " & T_LAB102 & " b, " & T_LAB101 & " a " & _
                    " WHERE  " & DBW("a.ptid   = ", PtId) & _
                    " AND    " & DBW("a.donefg = ", enStsCd.StsCd_LIS_Order) & _
                    " AND    " & DBW("a.orddt <= ", ReqDt) & _
                    " AND    " & DBW("a.orddiv = ", lis_orddiv)
    If DeptCd <> "" Then
        SqlReadOutOrder = SqlReadOutOrder & " AND   " & DBW("a.deptcd = ", DeptCd)
    End If
    
    SqlReadOutOrder = SqlReadOutOrder & _
                    " AND    b.ptid  = a.ptid " & _
                    " AND    b.orddt = a.orddt " & _
                    " AND    b.ordno = a.ordno " & _
                    " AND    " & DBW("b.donefg = ", enStsCd.StsCd_LIS_Order) & strtmp & _
                    " AND   ( b.dcfg = '' or b.dcfg is null ) " & _
                    " AND    c.testcd = b.ordcd " & _
                    " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                        " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
                    " AND    d.cdval1  = b.spccd " & _
                    " AND    " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                    " AND    f.testcd  = b.ordcd AND f.spccd = b.spccd " & _
                    " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                                        " WHERE testcd = f.testcd   AND     spccd = f.spccd ) " & _
                    " Order By ColTm, orddt, ordno, ordseq "
                    
End Function

Public Function SqlReadOrderER(ByVal PtId As String, ByVal ReqDt As String, ByVal ReceptNo As String) As String

    Dim tmpStr As String
    
    SqlReadOrderER = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, b.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                             "          d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                             "          a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, f.labelcnt " & _
                             " FROM   " & T_LAB101 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                             " WHERE  " & DBW("a.ptid", PtId, 2) & _
                             " AND    " & DBW("a.donefg", "0", 2) & _
                             " AND    " & DBW("a.orddt", ReqDt, 2) & _
                             " AND    " & DBW("a.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                             " AND    " & DBW("a.deptcd", "ER", 2) & _
                             " AND    " & DBW("a.receptno", ReceptNo, 2) & _
                             " AND    b.ptid = a.ptid " & _
                             " AND    b.orddt = a.orddt " & _
                             " AND    b.ordno = a.ordno " & _
                             " AND    " & DBW("b.donefg", "0", 2) & _
                             " AND  ( b.dcfg = '' or b.dcfg is null ) " & _
                             " AND    c.testcd = b.ordcd " & _
                             " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                             " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & _
                             " AND    d.cdval1 = b.spccd " & _
                             " AND    " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                             " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                             " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd   AND     spccd = f.spccd ) " & _
                             " Order By ColTm, orddt, ordno, ordseq "

End Function



Public Function SqlGsBatchOrder(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                           Optional ByVal StatFg As String = "", Optional ByVal Bussdiv As String = "", Optional ByVal ReceptNo As String = "") As String

    Dim tmpStr As String
    
    '여성클리닉 환자들의 Gram Stain 일괄검사처리를 위한 Query 문
    '--> 검사항목-'B400', 검체-'75', 진료과-'OB'
    tmpStr = ""
    If StatFg <> "" Then tmpStr = tmpStr & "AND    b.statfg = '" & StatFg & "' "   '응급여부
    
        SqlGsBatchOrder = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, b.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                      "       d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, " & _
                      "       b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                      "       e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, f.labelcnt, '' as ErrFg " & _
                      " FROM  " & T_LAB101 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & _
                                  T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                      " WHERE " & DBW("a.ptid", PtId, 2) & _
                      " AND   " & DBW("a.donefg", "0", 2) & _
                      " AND   " & DBW("a.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                      " AND   " & DBW("a.orddt", ReqDt, 2) & _
                      " AND   a.deptcd in (" & DBV("deptcd", "OB") & "," & DBV("deptcd", "CO") & ") " & _
                      " AND   b.ptid = a.ptid   AND  b.orddt = a.orddt   AND    b.ordno = a.ordno " & _
                      " AND   " & DBW("b.ordcd", "B400", 2) & _
                      " AND   " & DBW("b.donefg", "0", 2) & _
                      " AND   " & DBW("b.spccd", "75", 2) & _
                      " AND  (b.dcfg or b.dcfg is null ) " & _
                      tmpStr & _
                      " AND   c.testcd = b.ordcd " & _
                      " AND   c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                      " AND   " & DBW("d.cdindex", LC3_Specimen, 2) & _
                      " AND    d.cdval1 = b.spccd " & _
                      " AND   " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                      " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                      " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd   AND     spccd = f.spccd ) "
        SqlGsBatchOrder = SqlGsBatchOrder & " UNION ALL " & _
                      " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, '' as spccd, b.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                      "       '' as SpcNm, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, " & _
                      "       b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                      "       e." & F_DOCTNM & " as DoctNm, a.deptcd,  '' as statflags, 0 as labelcnt, 'Y' as ErrFg " & _
                      " FROM  " & T_LAB101 & " a, " & T_LAB102 & " b,  " & _
                                  T_LAB001 & " c, " & T_HIS005 & " e " & _
                      " WHERE " & DBW("a.ptid", PtId, 2) & _
                      " AND   " & DBW("a.donefg", "0", 2) & _
                      " AND   " & DBW("a.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                      " AND   " & DBW("a.orddt", ReqDt, 2) & _
                      " AND   a.deptcd in (" & DBV("deptcd", "OB") & "," & DBV("deptcd", "CO") & ") " & _
                      " AND   b.ptid = a.ptid   AND  b.orddt = a.orddt   AND    b.ordno = a.ordno " & _
                      " AND   " & DBW("b.ordcd", "B400", 2) & _
                      " AND   " & DBW("b.donefg", "0", 2) & _
                      " AND   " & DBW("b.spccd", "75", 2) & _
                      " AND  (b.dcfg='' or b.dcfg is null ) " & _
                      tmpStr & _
                      " AND    c.testcd = b.ordcd " & _
                      " AND    c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                      " AND   " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                      " AND    NOT EXISTS ( SELECT * FROM " & T_LAB004 & " WHERE testcd = b.ordcd AND spccd = b.spccd " & _
                      "                     AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                      " Order By ColTm, orddt, ordno, ordseq "
End Function

'병동채혈/접수 POCT만 검색
Public Function SqlReadWardOrder_New(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                 Optional ByVal StatFg As String = "", Optional ByVal Bussdiv As String = "", _
                                 Optional ByVal ReceptNo As String = "", Optional ByVal ORDDIV As String = "") As String
 
    Dim tmpStr      As String
    Dim tmpStr1     As String
    Dim strSQL(2)   As String

    tmpStr = "": tmpStr1 = ""
    
    If StatFg <> "" Then tmpStr = tmpStr & " AND   " & DBW("b.statfg = ", StatFg)       '응급여부
    If Bussdiv <> "" Then tmpStr = tmpStr & " AND  " & DBW("a.bussdiv = ", Bussdiv)     '외래병동구분
    
    If Bussdiv = enBussDiv.BussDiv_OutPatient Then  '외래
        tmpStr = tmpStr & " AND  " & DBW("a.orddt = ", ReqDt)
    Else
'        tmpStr1 = " AND a.reqdt||a.reqtm>='" & ReqDt & ReqTm & "'"
        tmpStr1 = " AND a.reqdt='" & ReqDt & "'"
    End If
    
    If ReceptNo <> "" Then tmpStr = tmpStr & " AND (b.paydt<>'' or b.paydt is not null)  "
    
    strSQL(1) = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv,  a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("num", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB004 & " f, " & T_LAB102 & " b, " & T_LAB101 & " a " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND   " & DBW("a.donefg =", enStsCd.StsCd_LIS_Order) & _
                " AND  " & DBW("a.orddiv = ", lis_orddiv) & tmpStr1 & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND   " & DBW(" b.donefg =", enStsCd.StsCd_LIS_Order) & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) " & _
                " AND    b.ordcd in ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E','C3710A') "
    
    '혈액은행 : Phersis검사는 제외(testdiv <> '3')...
    strSQL(2) = " SELECT c.testnm, c.abbrnm5, '3' as testdiv, 'XM' as workarea, b.spccd, '' as storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        '혈액' as SpcNm, '혈액' as SpcNm5, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, a.deptcd,  '' as statflags, " & _
                "        1 as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_BBS001 & " c, " & _
                           T_LAB102 & " b, " & T_LAB101 & " a " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND   " & DBW("a.donefg =", enStsCd.StsCd_LIS_Order) & _
                " AND   " & DBW("a.orddiv = ", BBS_ORDDIV) & tmpStr1 & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND   " & DBW(" b.donefg =", enStsCd.StsCd_LIS_Order) & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd = b.ordcd " & _
                " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_BBS001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') "
 
    Select Case ORDDIV
        Case "B": SqlReadWardOrder_New = strSQL(2)
        Case "L": SqlReadWardOrder_New = strSQL(1)
        Case "W":
            If P_IncludeBBSSystem Then
                SqlReadWardOrder_New = strSQL(1) & " UNION ALL " & strSQL(2) '& " UNION ALL " & strSQL(1)
            Else
                SqlReadWardOrder_New = strSQL(1)
            End If
    End Select

    SqlReadWardOrder_New = SqlReadWardOrder_New & " Order By ColTm, orddt, ordno, ordcd "                          '<< D/C 처방 제외 >>
    
' 수정 : 채혈실 요구로 처방일자, 희망채혈일시,검체명 순으로
'    SqlReadWardOrder = SqlReadWardOrder & " Order By orddt, reqdt DESC, SpcNm ASC"                          '<< D/C 처방 제외 >>
    Debug.Print SqlReadWardOrder_New
End Function

Public Function SqlReadWardOrder(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                 Optional ByVal StatFg As String = "", Optional ByVal Bussdiv As String = "", _
                                 Optional ByVal ReceptNo As String = "", Optional ByVal ORDDIV As String = "") As String
 
    Dim tmpStr      As String
    Dim tmpStr1     As String
    Dim strSQL(2)   As String

    tmpStr = "": tmpStr1 = ""
    
    If StatFg <> "" Then tmpStr = tmpStr & " AND   " & DBW("b.statfg = ", StatFg)       '응급여부
    If Bussdiv <> "" Then tmpStr = tmpStr & " AND  " & DBW("a.bussdiv = ", Bussdiv)     '외래병동구분
    
    If Bussdiv = enBussDiv.BussDiv_OutPatient Then  '외래
        tmpStr = tmpStr & " AND  " & DBW("a.orddt = ", ReqDt)
    Else
        tmpStr1 = " AND a.reqdt||a.reqtm<='" & ReqDt & ReqTm & "'"
    End If
    
    If ReceptNo <> "" Then tmpStr = tmpStr & " AND (b.paydt<>'' or b.paydt is not null)  "
    
    strSQL(1) = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv,  a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("num", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB004 & " f, " & T_LAB102 & " b, " & T_LAB101 & " a " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND   " & DBW("a.donefg =", enStsCd.StsCd_LIS_Order) & _
                " AND  " & DBW("a.orddiv = ", lis_orddiv) & tmpStr1 & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND   " & DBW(" b.donefg =", enStsCd.StsCd_LIS_Order) & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) "
    
    
    '혈액은행 : Phersis검사는 제외(testdiv <> '3')...
    strSQL(2) = " SELECT c.testnm, c.abbrnm5, '3' as testdiv, 'XM' as workarea, b.spccd, '' as storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        '혈액' as SpcNm, '혈액' as SpcNm5, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, a.deptcd,  '' as statflags, " & _
                "        1 as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_BBS001 & " c, " & _
                           T_LAB102 & " b, " & T_LAB101 & " a " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND   " & DBW("a.donefg =", enStsCd.StsCd_LIS_Order) & _
                " AND   " & DBW("a.orddiv = ", BBS_ORDDIV) & tmpStr1 & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND   " & DBW(" b.donefg =", enStsCd.StsCd_LIS_Order) & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd = b.ordcd " & _
                " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_BBS001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') "
 
    Select Case ORDDIV
        Case "B": SqlReadWardOrder = strSQL(2)
        Case "L": SqlReadWardOrder = strSQL(1)
        Case "W":
            If P_IncludeBBSSystem Then
                SqlReadWardOrder = strSQL(1) & " UNION ALL " & strSQL(2) '& " UNION ALL " & strSQL(1)
            Else
                SqlReadWardOrder = strSQL(1)
            End If
    End Select

    SqlReadWardOrder = SqlReadWardOrder & " Order By ColTm, orddt, ordno, ordcd "                          '<< D/C 처방 제외 >>
    
' 수정 : 채혈실 요구로 처방일자, 희망채혈일시,검체명 순으로
'    SqlReadWardOrder = SqlReadWardOrder & " Order By orddt, reqdt DESC, SpcNm ASC"                          '<< D/C 처방 제외 >>
    Debug.Print SqlReadWardOrder
End Function



'공통코드 마스터(s2lab032)의 'C242'에 ICU 항목이 정의되어 있슴
Public Function SqlReadICUOrder(ByVal PtId As String, ByVal ReqDt As String) As String
 

     '임상병리

    SqlReadICUOrder = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                        "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                        "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                                 FUNC_CONVERT("num", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                        "        '' as bedid, '' as fzfg " & _
                        " FROM " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
                                   T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f, " & T_LAB032 & " g " & _
                        " WHERE " & DBW("a.ptid = ", PtId) & _
                        " AND    a.donefg = '0' " & _
                        " AND  " & DBW("a.reqdt <= ", ReqDt) & _
                        " AND  " & DBW("a.orddiv = ", lis_orddiv) & _
                        " AND  " & DBW("a.bussdiv = ", enBussDiv.BussDiv_ICU) & _
                        " AND    b.ptid  = a.ptid  AND  b.orddt = a.orddt  AND  b.ordno = a.ordno " & _
                        " AND    b.donefg = '0' " & _
                        " AND   (b.dcfg = '' or b.dcfg is null) " & _
                        " AND  " & DBW("g.cdindex = ", LC3_ICUTestCd) & _
                        " AND    g.cdval1  = b.ordcd " & _
                        " AND    c.testcd  = b.ordcd " & _
                        " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                        " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                        " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                        " AND  " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                        " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                        " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND  spccd = f.spccd ) "
    
    SqlReadICUOrder = SqlReadICUOrder & " Order By ColTm, b.orddt, b.ordno, b.ordcd "                          '<< D/C 처방 제외 >>
      
End Function


Public Function SqlReadOrderForMornCol(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String) As String
 
    Dim tmpStr As String
    Dim strSQL(3) As String

    tmpStr = ""

     '임상병리 아침채혈
    SqlReadOrderForMornCol = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("num", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
                           T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND    a.donefg = '0' " & _
                " AND  " & DBW("a.reqdt < ", ReqDt) & _
                " AND  " & DBW("a.orddiv = ", lis_orddiv) & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND    b.donefg = '0' " & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND  " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) " & _
                " AND    f.rndfg = '1'"
    SqlReadOrderForMornCol = SqlReadOrderForMornCol & " UNION ALL " & _
                "SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("num", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
                           T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND    a.donefg = '0' " & _
                " AND  " & DBW("a.reqdt = ", ReqDt) & _
                " AND  " & DBW("a.reqtm <= ", ReqTm) & _
                " AND  " & DBW("a.orddiv = ", lis_orddiv) & _
                " AND    b.ptid  = a.ptid  AND    b.orddt = a.orddt  AND    b.ordno = a.ordno " & _
                " AND    b.donefg = '0' " & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND  " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) " & _
                " AND    f.rndfg = '1'"
                
                

    SqlReadOrderForMornCol = SqlReadOrderForMornCol & " Order By ColTm, orddt, ordno, ordcd "                         '<< D/C 처방 제외 >>
      
End Function
   

Public Function SqlWardBarReprint(ByVal intOption As Integer, ParamArray strKeys() As Variant) As String

    Dim tmpStr As String
    Dim tmpStr1 As String
    Dim tmpSql(3) As String
    Dim blnAllFg As Boolean
   
    tmpStr = ""
    blnAllFg = False
    If intOption = 1 Then  '환자ID-처방일 기준
        tmpStr = tmpStr & "b.ptid = " & strKeys(0) & _
                 " AND " & DBW("b.orddt = ", strKeys(1)) & _
                 " AND " & DBW("b.donefg > ", 0) & _
                 " AND (b.dcfg = '' or b.dcfg is null) "
        tmpStr1 = " AND e.ptid = b.ptid AND e.orddt = b.orddt " & _
                  " AND e.ordno = b.ordno AND e.ordseq = b.ordseq " & _
                  " AND a.workarea = e.workarea AND a.accdt = e.accdt " & _
                  " AND a.accseq = e.accseq "
        If UBound(strKeys) = 2 Then
            If strKeys(2) = "" Then blnAllFg = True
        End If
    Else        '접수번호 기준
        tmpStr = tmpStr & DBW("a.workarea = ", strKeys(0)) & _
                " AND " & DBW("a.accdt    = ", strKeys(1)) & _
                " AND " & DBW("a.accseq   = ", strKeys(2))
        tmpStr1 = " AND e.workarea = a.workarea AND e.accdt = a.accdt " & _
                  " AND e.accseq = a.accseq AND b.ptid = e.ptid " & _
                  " AND b.orddt = e.orddt AND b.ordno = e.ordno " & _
                  " AND b.ordseq = e.ordseq  " & _
                  " AND (b.dcfg = '' or b.dcfg is null) "
    End If
    
    '임상병리
    tmpSql(1) = " SELECT c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                "        a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                "        a.deptcd, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & _
                         FUNC_CONVERT("char", "a.accseq") & " as LabNo, d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
                "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "        e.field1 as BuildNm, f." & F_PTNM & " as PtNm, i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                " FROM " & T_LAB201 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i " & _
                " WHERE " & tmpStr & _
                " AND   (a.multifg = '' or a.multifg is null)" & _
                " AND    b.workarea = a.workarea " & _
                " AND    b.accdt    = a.accdt " & _
                " AND    b.accseq   = a.accseq " & _
                " AND   (b.dcfg = ''  or  b.dcfg is null)" & _
                " AND    c.testcd = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd " & _
                                    " AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND " & DBW("d.cdindex = ", LC3_Specimen) & _
                " AND    d.cdval1 = a.spccd " & _
                " AND " & DBJ(DBW("e.cdindex = ", LC3_Buildings)) & _
                " AND " & DBJ("e.cdval1 =* a.buildcd") & _
                " AND    f." & F_PTID & " = a.ptid " & _
                " AND    i.ptid = b.ptid " & _
                " AND    i.orddt = b.orddt " & _
                " AND    i.ordno = b.ordno "
     
    tmpSql(1) = tmpSql(1) & "UNION ALL " & _
                " SELECT c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                "        a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                "        a.deptcd, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & _
                         FUNC_CONVERT("char", "a.accseq") & " as LabNo, d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & _
                "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "       g.field1 as BuildNm, h." & F_PTNM & " as PtNm, i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                " FROM " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB203 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g, " & _
                           T_HIS001 & " h, " & T_LAB101 & " i " & _
                " WHERE " & tmpStr & tmpStr1 & _
                " AND    a.multifg <> ' ' AND  a.multifg is not null " & _
                " AND  " & DBW("f.cdindex = ", LC2_Panel) & _
                " AND    f.cdval1 = b.ordcd  AND  f.cdval2 = e.spcseq " & _
                " AND    c.testcd = f.field1 " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd " & _
                                    " AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBW("d.cdindex = ", LC3_Specimen) & "  AND   d.cdval1 = a.spccd " & _
                " AND  " & DBJ(DBW("g.cdindex = ", LC3_Buildings)) & " AND  " & DBJ("g.cdval1 =* a.buildcd") & _
                " AND    h." & F_PTID & " = a.ptid " & _
                " AND    i.ptid  = b.ptid " & _
                " AND    i.orddt = b.orddt " & _
                " AND    i.ordno = b.ordno " & _
                " AND  " & DBW("i.orddiv = ", lis_orddiv)

    '해부병리
    tmpSql(2) = " "
    '혈액은행
    tmpSql(3) = " SELECT c.testnm, c.abbrnm5, a.ptid, 'XM' as workarea, '' as accdt, 0 as accseq, " & _
                "        a.spcyy, a.spcno, '' spccd, '' as storecd, '1' as stscd, b.statfg, a.buildcd, " & _
                "        i.deptcd, i.hosilid, '' as LabNo, '혈액' as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
                "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "        e.field1 as BuildNm, f." & F_PTNM & " as PtNm, " & _
                "        i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                " FROM " & T_BBS201 & " a, " & T_LAB102 & " b, " & T_BBS001 & " c, " & _
                           T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i " & _
                " WHERE " & tmpStr & _
                " AND    i.ptid = b.ptid " & _
                " AND    i.orddt = b.orddt " & _
                " AND    i.ordno = b.ordno " & _
                " AND  " & DBW("i.orddiv = ", BBS_ORDDIV) & _
                " AND    a.ptid = i.ptid " & _
                " AND    a.spcyy = (SELECT max(spcyy) FROM " & T_BBS201 & " WHERE ptid = i.ptid) " & _
                " AND    a.spcno = (SELECT max(spcno) FROM " & T_BBS201 & _
                                  " WHERE ptid = a.ptid AND spcyy = a.spcyy) " & _
                " AND    c.testcd = b.ordcd " & _
                " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
                " AND  " & DBJ(DBW("e.cdindex = ", LC3_Buildings)) & _
                " AND  " & DBJ("e.cdval1 =* a.buildcd") & _
                " AND    f." & F_PTID & " = a.ptid "
    
    If blnAllFg Then
        If P_IncludeBBSSystem Then
            SqlWardBarReprint = tmpSql(1) & " UNION ALL " & tmpSql(3)
        Else
            SqlWardBarReprint = tmpSql(1)
        End If
    Else
        Select Case strKeys(2)
            Case BBS_ORDDIV: SqlWardBarReprint = tmpSql(3)
            Case lis_orddiv: SqlWardBarReprint = tmpSql(1)
        End Select
    End If
    SqlWardBarReprint = SqlWardBarReprint & " Order By LabNo, b.ordno, b.ordcd "

End Function

Public Function Get_SpcNm(ByVal SpcCd As String, ByVal ORDDIV As String) As String
    Dim rs   As Recordset
    Dim sSQL As String
                            
    Select Case ORDDIV
        Case lis_orddiv: sSQL = "SELECT field5 as spcname FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", LC3_Specimen) & " AND " & DBW("cdval1=", SpcCd)
        Case BBS_ORDDIV: Get_SpcNm = "": Exit Function
    End Select
    If sSQL = "" Then Exit Function
    
    
    Set rs = New Recordset
    rs.Open sSQL, DBConn
'    If RS.DBerror Then
'        Get_SpcNm = ""
'        dbconn.DisplayErrors
'        Exit Function
'    End If
    
    If Not rs.EOF Then
        Get_SpcNm = rs.Fields("spcname").Value & ""
    Else
         Get_SpcNm = ""
    End If
    
    Set rs = Nothing
End Function


Public Function SqlArchiveWardColList(ByVal pDeptcd As String, ByVal pYesterday As String) As String
    
    SqlArchiveWardColList = " Delete FROM " & T_LAB204 & _
                            " WHERE  " & DBW("workdt <  ", pYesterday) & _
                            " AND    " & DBW("wardid = ", pDeptcd)

End Function

'-- 추가 병동채혈 ToolTip에 해당환자에 대한 검사항목Display
' - By M.G.Choi 2002.08.06
Public Function WardMn_ORDCD(ByVal pPtid As String, ByVal pOrdDt As String, _
                             ByVal pWardId As String) As String
    
     
    WardMn_ORDCD = "SELECT c.abbrnm5, a.orddiv " _
                                & "  FROM " & T_LAB101 & " " & "a, " & T_LAB102 & " " & "b, " _
                                & T_BBS001 & " " & "c" _
                                & " WHERE " & DBW("a.ptid = ", pPtid) _
                                & "   AND a.reqdt <= '" & pOrdDt & "'" _
                                & "   AND a.wardid = '" & pWardId & "'" _
                                & "   AND " & DBW("a.donefg", enStsCd.StsCd_LIS_Order, 2) _
                                & "   AND a.orddt = b.orddt " _
                                & "   AND a.ordno = b.ordno " _
                                & "   AND b.ordcd = c.testcd " _
                                & "   AND a.ptid = b.ptid "
                                
    WardMn_ORDCD = WardMn_ORDCD & "Union " _
                                & "SELECT c.abbrnm5, a.orddiv " _
                                & "  FROM " & T_LAB101 & " " & "a, " & T_LAB102 & " " & "b, " _
                                & T_LAB001 & " " & "c" _
                                & " WHERE " & DBW("a.ptid = ", pPtid) _
                                & "   AND a.reqdt <= '" & pOrdDt & "'" _
                                & "   AND a.wardid = '" & pWardId & "'" _
                                & "   AND " & DBW("a.donefg", enStsCd.StsCd_LIS_Order, 2) _
                                & "   AND a.orddt = b.orddt " _
                                & "   AND a.ordno = b.ordno " _
                                & "   AND b.ordcd = c.testcd " _
                                & "   AND a.ptid = b.ptid " _
                                & " order by orddiv "
End Function
'-- 추가 외래일괄채혈 ToolTip에 해당환자에 대한 검사항목Display
' - By M.G.Choi 2002.11.14
Public Function DeptCd_ORDCD(ByVal pPtid As String, ByVal pOrdDt As String, _
                             ByVal pDeptcd As String) As String

    DeptCd_ORDCD = "SELECT c.abbrnm5, a.orddiv " _
                & "  FROM " & T_LAB101 & " " & "a, " & T_LAB102 & " " & "b, " _
                & T_BBS001 & " " & "c" _
                & " WHERE " & DBW("a.ptid = ", pPtid) _
                & "   AND a.reqdt <= '" & pOrdDt & "'" _
                & "   AND a.deptcd = '" & pDeptcd & "'" _
                & "   AND " & DBW("a.donefg", enStsCd.StsCd_LIS_Order, 2) _
                & "   AND a.orddt = b.orddt " _
                & "   AND a.ordno = b.ordno " _
                & "   AND b.ordcd = c.testcd " _
                & "   AND a.ptid = b.ptid "
                                
    DeptCd_ORDCD = DeptCd_ORDCD & "Union " _
                                & "SELECT c.abbrnm5, a.orddiv " _
                                & "  FROM " & T_LAB101 & " " & "a, " & T_LAB102 & " " & "b, " _
                                & T_LAB001 & " " & "c" _
                                & " WHERE " & DBW("a.ptid = ", pPtid) _
                                & "   AND a.reqdt <= '" & pOrdDt & "'" _
                                & "   AND a.deptcd = '" & pDeptcd & "'" _
                                & "   AND " & DBW("a.donefg", enStsCd.StsCd_LIS_Order, 2) _
                                & "   AND a.orddt = b.orddt " _
                                & "   AND a.ordno = b.ordno " _
                                & "   AND b.ordcd = c.testcd " _
                                & "   AND a.ptid = b.ptid " _
                                & " order by orddiv "
End Function

'===================================================
'산부인과 외래 일괄채혈시 Gram Stain 일괄채혈 조회문
'===================================================
Public Function OGYOutOrder(ByVal ReqDt As String, ByVal ReqTm As String, ByVal DeptCd As String) As String
   
    Dim sSQL      As String
    Dim rs        As Recordset
    Dim strtmp    As String
   
    sSQL = "SELECT cdval1,field1 FROM " & T_LAB031 & " WHERE " & DBW("cdindex=", LC2_OGYITEM)
   
    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        Do Until rs.EOF
            If Trim(rs.Fields("field1").Value & "") = "L" Then
                strtmp = strtmp & "'" & rs.Fields("cdval1").Value & "" & "'" & ","
            End If
            rs.MoveNext
        Loop
        strtmp = Mid(strtmp, 1, Len(strtmp) - 1)
        OGYOutOrder = " SELECT  b.deptcd , a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                                 F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                                 F_DOB2("a") & " as Dob, d.field5 spcnm, " & _
                        "        b.orddoct, b.majdoct, b.orddiv  " & _
                        " FROM   " & T_LAB032 & " d, " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                        " WHERE  " & DBW("b.donefg", "0", 2) & _
                        " AND    " & DBW("b.reqdt <= ", ReqDt) & _
                        " AND    " & DBW("b.orddiv=", "L") & _
                        " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                        " AND    " & DBW("b.deptcd", DeptCd, 2) & _
                        " AND    c.ordcd in (" & strtmp & ")" & _
                        " AND    a." & F_PTID & " = b.ptid " & _
                        " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                        " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                        " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
                        " AND    d.cdval1 = c.spccd "

    Else
        '검사항목 item 이 없을경우, 아무것도 없이 조회하기 위해서
        OGYOutOrder = "SELECT * FROM " & T_LAB031 & " WHERE " & DBW("cdindex=", "99999999")
        
    End If
    Set rs = Nothing
End Function

Public Function OGySQlOrderRead(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                Optional ByVal DeptCd As String = "") As String
    Dim sSQL      As String
    Dim rs        As Recordset
    Dim strtmp    As String
   
    sSQL = "SELECT cdval1 FROM " & T_LAB031 & " WHERE " & DBW("cdindex=", LC2_OGYITEM)
   
    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        Do Until rs.EOF
            strtmp = strtmp & "'" & rs.Fields("cdval1").Value & "" & "'" & ","
            rs.MoveNext
        Loop
        strtmp = Mid(strtmp, 1, Len(strtmp) - 1)
    End If
    Set rs = Nothing

    OGySQlOrderRead = "SELECT c.testcd, c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                    "        d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.paydt, " & _
                    "        b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                    "        e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, f.labelcnt, '' as ErrFg, '' wardid, '' roomid, '' bedid " & _
                    " FROM   " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_HIS005 & " e, " & _
                                 T_LAB004 & " f, " & T_LAB102 & " b, " & T_LAB101 & " a " & _
                    " WHERE  " & DBW("a.ptid   = ", PtId) & _
                    " AND    " & DBW("a.donefg = ", enStsCd.StsCd_LIS_Order) & _
                    " AND    " & DBW("a.orddt <= ", ReqDt) & _
                    " AND    " & DBW("a.orddiv = ", lis_orddiv) & _
                    " AND    b.ordcd in (" & strtmp & ")"
    If DeptCd <> "" Then
        OGySQlOrderRead = OGySQlOrderRead & " AND   " & DBW("a.deptcd = ", DeptCd)
    End If
    OGySQlOrderRead = OGySQlOrderRead & _
                    " AND    b.ptid  = a.ptid " & _
                    " AND    b.orddt = a.orddt " & _
                    " AND    b.ordno = a.ordno " & _
                    " AND    " & DBW("b.donefg = ", enStsCd.StsCd_LIS_Order) & _
                    " AND   ( b.dcfg = '' or b.dcfg is null ) " & _
                    " AND    c.testcd = b.ordcd " & _
                    " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                        " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    " AND    " & DBW("d.cdindex = ", LC3_Specimen) & _
                    " AND    d.cdval1  = b.spccd " & _
                    " AND    " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                    " AND    f.testcd  = b.ordcd AND f.spccd = b.spccd " & _
                    " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                                        " WHERE testcd = f.testcd   AND     spccd = f.spccd ) " & _
                    " Order By ColTm, orddt, ordno, ordseq "
                    
End Function

Public Function GetCollectionList(ByVal strCase As String, ByVal SearchKey As String, _
                                  Optional ByVal strWardId As String = "") As String
    Dim sSQL As String
    Dim sQuery As String
    Dim strReqDt As String
    Dim strReqTm As String

    strReqDt = ""
    
    Select Case ObjSysInfo.projectid
        Case "APS", "WMRA":
            sQuery = " AND b.orddiv='A'"
        Case Else
            sQuery = " AND b.orddiv<>'A'"
    End Select
            
    
    sSQL = "SELECT distinct a." & F_PTID & " as ptid, a." & F_PTNM & " as ptnm," & _
            F_SSN2("a") & " as SSN, " & _
            F_DOB2("a") & " as DOB " & _
            " " & _
            " FROM " & T_HIS001 & " a ," & T_LAB101 & " b," & T_LAB102 & " c " & _
            " WHERE  b.donefg = '0' " & _
            " AND b.bussdiv = '2'" & _
            " AND a." & F_PTID & " = b.ptid "
   sSQL = sSQL & strReqDt
   
   sSQL = sSQL & sQuery
   
   If Val(strCase) Mod 2 = 1 Then
        sSQL = sSQL & " AND a." & F_PTID & " >= " & DBV(F_PTID, SearchKey) & " "
    Else
        sSQL = sSQL & " AND a." & F_PTNM & " like '" & SearchKey & "%' "
    End If
    
    sSQL = sSQL & " AND b.wardid = '" & strWardId & "'" & _
                " AND b.ptid=c.ptid" & _
                " AND b.orddt=c.orddt" & _
                " AND b.ordno=c.ordno" & _
                " AND (c.dcfg='' or c.dcfg is null)" ' & _
                " AND c.ocsordno is not null"
    
    '성바오로병원인 경우 ocsordno 가 NUll인경우는 백업처방임
    
    If strCase = 1 Then
        sSQL = sSQL & " Order by a." & F_PTID
    Else
        sSQL = sSQL & " Order  by a." & F_PTNM & ", SSN"
    End If

    GetCollectionList = sSQL
End Function

'===================================================
'** 핵의학 외래채취 조회문
' - 핵의학 외래채취 여부 판단 (여:'Y', 부:'N")
'===================================================
Public Function RI_Collection_Flag(ByVal pPtid As String, ByVal pOrdDt As String) As String
    Dim rs          As New ADODB.Recordset
    Dim strtmp(3)   As String
    Dim strSQL      As String
    Dim strOrdDt    As String
    
    If pOrdDt <> "" Then
        strOrdDt = "TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')"
    Else
        strOrdDt = "''"
    End If
    
    '-- 종건
    strtmp(0) = " select patno,orddate, '3' as wrkdiv " & _
                "   from oras1.SU2EXAMT " & _
                "  where patno = " & DBS(pPtid) & _
                "    and orddate <= " & strOrdDt & _
                "    and slipcode like 'N'||'%' " & _
                "    and examstat is null "
                

    '-- 일건
    strtmp(1) = " select patno,orddate, '4' as wrkdiv " & _
                "   from oras1.SG2EXAMT " & _
                "  where patno = " & DBS(pPtid) & _
                "    and orddate <= " & strOrdDt & _
                "    and slipcode like 'N'||'%' " & _
                "    and examstat is null "
    
    '-- 일반처방
    strtmp(2) = " select patno,orddate, '' as wrkdiv " & _
                "   from oram1.MDEXMORT " & _
                "  where patno = " & DBS(pPtid) & _
                "    and ordgrp   = 'C1' " & _
                "    and orddate <= " & strOrdDt & _
                "    and slipcd like 'N'||'%' " & _
                "    and procstat is null " & _
                "    and nvl(DISCYN,  'N') = 'N' " & _
                "    and nvl(PRNYN,   'N') = 'N' " & _
                "    and nvl(PREACTYN,'N') = 'N' " & _
                "    and nvl(RCPSTAT, 'N') <> 'R' " & _
                "    and nvl(NURSTAT, '@') <> 'X' " & _
                "    and patsect <> 'I' "
                         
    '-- 병리검사
'    strtmp(3) = " select patno,orddate, '' as wrkdiv " & _
'                "   from oram1.MDEXMORT " & _
'                "  where patno = " & DBS(pPtid) & _
'                "    and ordgrp   = 'C1' " & _
'                "    and orddate <= " & strOrdDt & _
'                "    and procstat is null " & _
'                "    and ordcd IN ('C5930','C5931') " & _
'                "    and spccode1 IN ('S23','U10')"
                                                         
    strSQL = strtmp(0) & " UNION ALL " & strtmp(1) & " UNION ALL " & strtmp(2) '& " UNION ALL " & strtmp(3)
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        RI_Collection_Flag = "Y" & COL_DIV & rs.Fields("wrkdiv").Value & ""
    Else
        RI_Collection_Flag = "N" & COL_DIV & ""
    End If
                     
    rs.Close
    Set rs = Nothing
    
End Function


Public Function Gets2ORD101_Query(ByVal IDX As Integer, Optional ByVal pQuery As String) As String
Dim strSQL   As String
Dim strSQLT1 As String
    If IDX = 1 Then strSQLT1 = " and a.patno = '" & pQuery & "' "
    If IDX = 2 Then strSQLT1 = " and a.orddate  = to_date('" & pQuery & "', 'yyyymmdd') "
    
    strSQL = " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'hh24miss') as ordtm,  decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ,  to_char(a.hopedate, 'yyyymmdd') as REQDT,  DECODE(to_char(a.hopedate, 'hh24miss'), '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM,  a.meddept as DEPCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT,  to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from mdexmort a" & _
              " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P')  and to_char(a.meddate, 'yyyymmdd')  > '20090101'" & strSQLT1 & _
   " union all select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'hh24miss') as ordtm,  decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ,  to_char(a.orddate, 'yyyymmdd')  as REQDT,  to_char(a.orddate, 'hh24miss')                                                                                     as REQTM,  a.meddept as DEPCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT,  to_char(a.editdate, 'hh24miss') as ENTTM,  'B' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from mdbldort a " & _
              " where to_char(a.meddate, 'yyyymmdd')  > '20090101' " & strSQLT1 & _
   " union all select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.orddate, 'hh24miss') as ordtm,  '1'                              as bussdiv ,  to_char(a.admacpt, 'yyyymmdd') as BEDINDT ,  to_char(a.ordtime, 'yyyymmdd')  as REQDT,  to_char(a.ordtime, 'hh24miss')                                                                                     as REQTM,  a.meddept as DEPCD,  null      as ORDDOCT, null    as MAJDOCT,  a.entid  as ENTID,  to_char(a.entdate, 'yyyymmdd')  as ENTDT,  to_char(a.entdate, 'hh24miss')  as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  null     as ROOMID,  null     as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from su2examt a " & _
              " where a.ordgrp = 'C1'  and substr(a.slipcode, 1, 1) in ('L', 'N', 'P') " & strSQLT1 & _
   " union all select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.orddate, 'hh24miss') as ordtm,  '1'                              as bussdiv ,  to_char(a.admacpt, 'yyyymmdd') as BEDINDT ,  to_char(a.ordtime, 'yyyymmdd')  as REQDT,  to_char(a.ordtime, 'hh24miss')                                                                                     as REQTM,  a.meddept as DEPCD,  null      as ORDDOCT, null    as MAJDOCT,  a.entid  as ENTID,  to_char(a.entdate, 'yyyymmdd')  as ENTDT,  to_char(a.entdate, 'hh24miss')  as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  null     as ROOMID,  null     as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from sg2examt a " & _
              " where a.ordgrp = 'C1'  and substr(a.slipcode, 1, 1) in ('L', 'N', 'P') " & strSQLT1

    Gets2ORD101_Query = strSQL
    
End Function


