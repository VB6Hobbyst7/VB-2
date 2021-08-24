VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetReactionSQL(Optional ByVal cdval1 As String = "") As String
    '부작용사유 조회
    Dim sSql As String
    
    sSql = " SELECT cdval1,field1" & _
           " FROM " & T_COM003 & _
           " WHERE " & DBW("cdindex", BC2_REACTION_REASON, 2)
    If cdval1 <> "" Then
        sSql = sSql & " and " & DBW("cdval1=", cdval1)
    End If
    sSql = sSql & " ORDER BY cdval1"
    GetReactionSQL = sSql
End Function


Public Function PtListQuery(ByVal strFrDt As String, ByVal strToDt As String, ByVal sStscd As String) As String
'혈액출고/폐기 된 환자의 목록을 구한다.
    Dim sSql As String
    
    If sStscd = BBSBloodStatus.stsDELIVERY Then
        sSql = " SELECT distinct c.ptid,a.deliverydt as qdate FROM " & _
                    T_BBS202 & " c," & T_BBS401 & " b," & T_BBS402 & " a" & _
                " WHERE " & _
                          DBW("a.deliverydt>=", strFrDt) & _
                " AND " & DBW("a.deliverydt<=", strToDt) & _
                " AND  a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno" & _
                " AND " & DBW("b.stscd=", sStscd) & _
                " AND a.workarea=c.workarea AND a.accdt=c.accdt AND a.accseq=c.accseq" & _
                " order by qdate        "
                
    ElseIf sStscd = BBSBloodStatus.stsEXPIRE Then
        sSql = " SELECT distinct c.ptid,b.realexpdt as qdate FROM " & _
                        T_BBS202 & " c," & T_BBS401 & " b," & T_BBS402 & " a" & _
                " WHERE " & _
                          DBW("b.realexpdt>=", strFrDt) & _
                " AND " & DBW("b.realexpdt<=", strToDt) & _
                " AND  a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno" & _
                " AND " & DBW("b.stscd=", sStscd) & _
                " AND a.workarea=c.workarea AND a.accdt=c.accdt AND a.accseq=c.accseq" & _
                " order by qdate     "
    End If
    PtListQuery = sSql
End Function


Public Function GetOrderBodySQL(ByVal sPtid As String, ByVal sOrddt As String, ByVal sOrdno As String, _
                                ByVal sOrdseq As String) As String
    Dim sSql As String
    
    sSql = " SELECT a.mesg,b.testcd,b.testnm " & _
           " FROM " & T_BBS001 & " b," & T_LAB102 & " a" & _
           " WHERE " & _
                     DBW("a.ptid=", sPtid) & _
           " AND " & DBW("a.orddt=", sOrddt) & _
           " AND " & DBW("a.ordno=", sOrdno) & _
           " AND " & DBW("a.ordseq=", sOrdseq) & _
           " AND a.ordcd=b.testcd" & _
           " AND (b.expdt=' ' or b.expdt is null)" & _
           " AND " & DBW("b.applydt<=", Format(GetSystemDate, "YYYYMMDD"))
         
    GetOrderBodySQL = sSql
    
End Function

Public Function GetRealResultSQL(ByVal PtId As String, Optional ByVal sVfyDt As String) As String
    Dim Rs          As Recordset
    Dim sSql        As String
    Dim strTestCd   As String
    Dim strSpcCd    As String
    
    Set Rs = New Recordset
    
    sSql = " select cdval1,field2 from " & T_COM003 & _
           " where " & _
                  DBW("cdindex=", BC2_REACTION_TEST) & _
           " and (field5 is null or field5='')"
    Rs.Open sSql, DBConn

    If Not Rs.EOF Then
        Do Until Rs.EOF
            strTestCd = strTestCd & "'" & Rs.Fields("cdval1").Value & "" & "',"
            strSpcCd = strSpcCd & "'" & Rs.Fields("field2").Value & "" & "',"
            Rs.MoveNext
        Loop
        strTestCd = Mid(strTestCd, 1, Len(strTestCd) - 1)
        strSpcCd = Mid(strSpcCd, 1, Len(strSpcCd) - 1)
        
        GetRealResultSQL = " select a.abbrnm10,b.workarea,b.accdt,b.accseq,b.testcd,b.rstcd,b.rstunit,b.vfydt,b.vfytm " & _
               " from " & T_LAB001 & " a," & T_LAB302 & " b" & _
               " where " & _
                         DBW("b.ptid=", PtId) & _
               " and b.testcd in (" & strTestCd & ")" & _
               " and b.spccd in (" & strSpcCd & ")"
        If sVfyDt <> "" Then
            GetRealResultSQL = GetRealResultSQL & " and " & DBW("b.vfydt>=", sVfyDt)
        End If
       GetRealResultSQL = GetRealResultSQL & " and a.testcd=b.testcd order by b.workarea,b.accdt,b.accseq"
    End If
    Set Rs = Nothing
End Function

'----------------------------------
'공통마스터 테이블 lab031
'----------------------------------
Public Function Getlab031CdMST(ByVal CDINDEX As String, Optional ByVal cdval1 As String = "", _
                               Optional ByVal CdVal2 As String = "", Optional ByVal OrdField As String = "", _
                               Optional ByVal pField As String = "*") As String
    Getlab031CdMST = "select " & pField & " from " & T_LAB031 & _
                     " where " & DBW("cdindex", CDINDEX, 2)
                 
    If cdval1 <> "" Then
        Getlab031CdMST = Getlab031CdMST & " and " & DBW("cdval1", cdval1, 2)
        If CdVal2 <> "" Then Getlab031CdMST = Getlab031CdMST & " and " & DBW("cdval2", CdVal2, 2)
    End If
    If OrdField <> "" Then Getlab031CdMST = Getlab031CdMST & " order by " & OrdField
End Function

Public Function BloodDetailQuery(ByVal strFrDt As String, ByVal strToDt As String, _
                                 ByVal sPtid As String, ByVal sStscd As String) As String

    Dim sSql As String
    
    If sStscd = BBSBloodStatus.stsDELIVERY Then
        sSql = " SELECT f.step1,f.step2,f.step3,f.step4,f.rstv,d.ptid,d.orddt,d.ordno,d.ordseq,e.wardid," & _
               " e.hosilid,e.deptcd,f.vfydt,f.vfyid," & _
               " a.workarea,a.accdt,a.accseq,a.rstseq,a.deliverydt as qdate,a.bldsrc,a.bldyy,a.bldno,a.compocd,c.abbrnm as componm," & _
               " b.volumn , b.abo, b.rh, b.stscd, a.deliveryid, b.entdt, b.entid" & _
               " FROM " & T_BBS302 & " f," & T_LAB101 & " e," & T_BBS202 & " d," & T_BBS006 & " c," & T_BBS401 & " b," & T_BBS402 & " a" & _
               " WHERE " & DBW("a.deliverydt >=", strFrDt) & _
               " AND  " & DBW("a.deliverydt <=", strToDt) & _
               " AND " & DBW("b.stscd=", sStscd) & _
               " AND " & DBW("d.ptid =", sPtid) & _
               " AND a.workarea=d.workarea" & _
               " AND a.accdt=d.accdt" & _
               " AND a.accseq=d.accseq" & _
               " AND a.bldsrc=b.bldsrc" & _
               " AND a.bldyy =b.bldyy" & _
               " AND a.bldno=b.bldno" & _
               " AND a.compocd=b.compocd" & _
               " AND a.compocd=c.compocd" & _
               " AND (a.retfg=' ' or a.retfg is null)" & _
               " AND (a.expfg= ' ' or a.expfg is null)" & _
               " AND d.ptid=e.ptid and d.orddt=e.orddt and d.ordno=e.ordno" & _
               " AND a.workarea=f.workarea and a.accdt=f.accdt and a.accseq=f.accseq and a.rstseq=f.rstseq" & _
               " order by qdate"
    ElseIf sStscd = BBSBloodStatus.stsEXPIRE Then
        sSql = " SELECT f.step1,f.step2,f.step3,f.step4,f.rstv,d.ptid,d.orddt,d.ordno,d.ordseq,e.wardid," & _
               " e.hosilid,e.deptcd,f.vfydt,f.vfyid," & _
               " a.workarea,a.accdt,a.accseq,a.rstseq,b.realexpdt as qdate,a.bldsrc,a.bldyy,a.bldno,a.compocd,c.abbrnm as componm," & _
               " b.volumn , b.abo, b.rh, b.stscd, a.deliveryid, b.entdt, b.entid" & _
               " FROM " & T_BBS302 & " f," & T_LAB101 & " e," & T_BBS202 & " d," & T_BBS006 & " c," & T_BBS401 & " b," & T_BBS402 & " a" & _
               " WHERE " & DBW("b.realexpdt >=", strFrDt) & _
               " AND  " & DBW("b.realexpdt <=", strToDt) & _
               " AND " & DBW("b.stscd=", sStscd) & _
               " AND " & DBW("d.ptid =", sPtid) & _
               " AND a.workarea=d.workarea" & _
               " AND a.accdt=d.accdt" & _
               " AND a.accseq=d.accseq" & _
               " AND a.bldsrc=b.bldsrc" & _
               " AND a.bldyy =b.bldyy" & _
               " AND a.bldno=b.bldno" & _
               " AND a.compocd=b.compocd" & _
               " AND a.compocd=c.compocd" & _
               " AND (a.retfg=' ' or a.retfg is null)" & _
               " " & _
               " AND d.ptid=e.ptid and d.orddt=e.orddt and d.ordno=e.ordno" & _
               " AND a.workarea=f.workarea and a.accdt=f.accdt and a.accseq=f.accseq and a.rstseq=f.rstseq" & _
               " order by qdate"
               'AND (a.expfg= ' ' or a.expfg is null)
    End If
    BloodDetailQuery = sSql
End Function
                                
Public Function GetCrossmatchingStep() As String
'--------------------------------------------------
'X-Matching 검사Step 조회
'검사단계;검사명(각각의 검사는 ";"로 연결되어있다.)
'--------------------------------------------------
    Dim today As String
    Dim sSql  As String
    
    today = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    sSql = " select a.field1,a.text1 " & _
           " from " & T_COM003 & " a" & _
           " where" & _
                  "  " & DBW("cdindex", BC2_XM_STEP, 2) & _
                  " and cdval1=(select max(b.cdval1)" & _
                              " from " & T_COM003 & " b " & _
                              " where" & _
                                       "     " & DBW("b.cdval1<", today, 2) & _
                                       " and " & DBW("cdindex", BC2_XM_STEP, 2) & ")"
    
    
    GetCrossmatchingStep = sSql

End Function

Public Function GetReactionChkSQL(ByVal sPtid As String, Optional ByVal sBldsrc As String = "", _
                                  Optional ByVal sBldYY As String = "", Optional ByVal sBldno As String = "", _
                                  Optional ByVal sCompocd As String = "") As String
    Dim sSql As String
    
    If sBldsrc <> "" Then
        sSql = " SELECT * FROM " & T_ICS103 & " WHERE " & _
                         DBW("bldsrc=", sBldsrc) & _
               " and " & DBW("bldyy=", sBldYY) & _
               " and " & DBW("bldno=", sBldno) & _
               " and " & DBW("compocd=", sCompocd) & _
               " and " & DBW("ptid=", sPtid)
    Else
        sSql = " SELECT * FROM " & T_ICS103 & " WHERE " & _
                         DBW("ptid=", sPtid) & _
               " ORDER BY ptid,reactiondt"
    End If
    GetReactionChkSQL = sSql
End Function

Public Function SetReactionSaveSQL(ByVal sString As String, ParamArray NewVal() As Variant) As String
    Dim sSql    As String
    
    '1:bldsrc,      2:bldyy,        3:bldno,        4:compocd,      5:ptid,
    '6:workarea,    7:accdt,        8:accseq,       9:rstseq,       10:step1
    '11:step2,      12:step3,       13:step4,       14:rstv,        15:transdt,
    '16:transtm,    17:volumn,      18:reactiondt,  19:reactiontm,  20:reactionid
    '21:reactioncd, 22:reactionnm,  23:mesg         24:deliverydt
    
    If sString = "SAVE" Then
        sSql = " insert into " & T_ICS103 & _
               "(bldsrc,bldyy,bldno,compocd,ptid,workarea,accdt,accseq,rstseq,step1,step2,step3,step4,rstv," & _
               "transdt,transtm,volumn,reactiondt,reactiontm,reactionid,reactioncd,reactionnm,mesg,deliverydt) values(" & _
                DBV("bldsrc", NewVal(0), 1) & DBV("bldyy", NewVal(1), 1) & DBV("bldno", NewVal(2), 1) & DBV("compocd", NewVal(3), 1) & _
                DBV("ptid", NewVal(4), 1) & DBV("workarea", NewVal(5), 1) & DBV("accdt", NewVal(6), 1) & DBV("accseq", NewVal(7), 1) & _
                DBV("rstseq", NewVal(8), 1) & DBV("step1", NewVal(9), 1) & DBV("step2", NewVal(10), 1) & DBV("step3", NewVal(11), 1) & _
                DBV("step4", NewVal(12), 1) & DBV("rstv", NewVal(13), 1) & DBV("transdt", NewVal(14), 1) & DBV("transtm", NewVal(15), 1) & _
                DBV("volumn", NewVal(16), 1) & DBV("reactiondt", NewVal(17), 1) & DBV("reactiontm", NewVal(18), 1) & DBV("reactionid", NewVal(19), 1) & _
                DBV("reactioncd", NewVal(20), 1) & DBV("reactionnm", NewVal(21), 1) & DBV("mesg", NewVal(22), 1) & DBV("deliverydt", NewVal(23)) & _
                ")"
    ElseIf sString = "DELETE" Then
        sSql = " DELETE " & T_ICS103 & " WHERE " & _
                         DBW("bldsrc=", NewVal(0)) & _
               " AND " & DBW("bldyy=", NewVal(1)) & _
               " AND " & DBW("bldno=", NewVal(2)) & _
               " AND " & DBW("compocd=", NewVal(3)) & _
               " AND " & DBW("ptid=", NewVal(4))
    End If
    SetReactionSaveSQL = sSql
End Function

Public Function GetCompoentSQL(Optional ByVal sCompocd As String = "") As String
    Dim sSql As String
    
    If sCompocd <> "" Then
        sSql = " select * from " & T_BBS006 & _
               " where " & DBW("compocd=", sCompocd) & _
               " and (expdt=' '  or expdt is null)" & _
               " order by compocd"
    Else
        sSql = " select * from " & T_BBS006 & _
               " where  (expdt=' '  or expdt is null)" & _
               " order by compocd"
    End If
    GetCompoentSQL = sSql
End Function


Public Function GetReationStatics(ByVal sYear As String, ByVal StaticsDiv As String) As String
    Dim sFDate  As String
    Dim sTDate  As String
    Dim sSql    As String
    
    sFDate = sYear & "0101"
    sTDate = sYear & "1231"
    
    Select Case StaticsDiv
        Case "0": '혈액제제별 건수
            sSql = " select compocd as querydiv,substr(reactiondt,5,2) as reactiondate, count(*) as cnt " & _
                 " from " & T_ICS103 & _
                 " where " & DBW("ptid>", " ") & _
                 " and " & DBW("reactiondt>=", sFDate) & _
                 " and " & DBW("reactiondt<=", sTDate) & _
                 " group by compocd,reactiondt" & _
                 " order by compocd,reactiondt"
        Case "1": '수혈 부작용별 건수
            sSql = " select reactioncd as querydiv,substr(reactiondt,5,2) as reactiondate, count(*) as cnt " & _
                 " from " & T_ICS103 & _
                 " where " & DBW("ptid>", " ") & _
                 " and " & DBW("reactiondt>=", sFDate) & _
                 " and " & DBW("reactiondt<=", sTDate) & _
                 " group by reactioncd,reactiondt" & _
                 " order by reactioncd,reactiondt"

        Case "2" '수혈부작용별 제제건수
            sSql = " SELECT distinct a.reactioncd ,a.compocd ,b.abbrnm,count(*) as cnt " & _
                   " FROM " & T_BBS006 & " b," & T_ICS103 & " a" & _
                   " where " & DBW("ptid>", " ") & _
                   " and " & DBW("a.reactiondt>=", sFDate) & _
                   " AND " & DBW("a.reactiondt<=", sTDate) & _
                   " AND a.compocd=b.compocd" & _
                   " GROUP BY a.reactioncd,a.compocd,b.abbrnm" & _
                   " ORDER BY a.reactioncd,a.compocd"
    End Select
    
    GetReationStatics = sSql
    
End Function


