VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBloodQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public centercd As String
Public ABO      As String
Public RH       As String
Public EntdtF   As String
Public EntdtL   As String
Public CompoCd  As String
Public PtId     As String

Public Sub Clear()
    ABO = ""
    RH = ""
    EntdtF = ""
    EntdtL = ""
    CompoCd = ""
    PtId = ""
    centercd = ""
End Sub

Public Function GetBlood_UsedList() As Recordset
'--------------------------
'사용가능한 혈액만 조회한다
'--------------------------

    Dim sSql    As String

    sSql = " SELECT b.abbrnm as componm,a.compocd,a.bldsrc,a.bldyy,a.bldno, a.coldt,a.entdt," & _
           " a.available,a.entid,a.volumn,a.abo,a.rh,a.irrfg,a.stscd,a.hosfg, " & _
           " a.realexpdt, a.exprcvid, a.expid, a.exprsncd, a.larc, a.smlc, a.lare, a.smle " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & _
                   " a.entdt between '" & EntdtF & "' AND '" & EntdtL & "'"
                   
           
           
    
    If ABO <> "" Then sSql = sSql & " AND " & DBW("a.abo=", ABO)
    If RH <> "" Then sSql = sSql & " AND " & DBW("a.rh=", RH)
    If CompoCd <> "" Then sSql = sSql & " AND " & DBW("a.compocd=", CompoCd)
    If centercd <> "" Then sSql = sSql & " AND " & DBW("centercd=", centercd)
    
    
    sSql = sSql & "AND " & DBW("stscd<=", BBSBloodStatus.stsRETURN)
    
    sSql = sSql & " AND a.compocd=b.compocd " & _
           "ORDER BY a.abo,a.bldsrc,a.bldyy,a.bldno  "



    Set GetBlood_UsedList = New Recordset
    Call GetBlood_UsedList.Open(sSql, DBConn)
'    If GetBlood_UsedList.DBerror Then
'        dbconn.DisplayErrors
'        Set GetBlood_UsedList = Nothing
'    End If
End Function

Public Function GetBlood_AllList() As Recordset
    Dim sSql    As String

    sSql = " SELECT a.splitoutfg,b.abbrnm as componm,a.compocd,a.bldsrc,a.bldyy,a.bldno, a.coldt,a.entdt," & _
           " a.available,a.entid,a.volumn,a.abo,a.rh,a.stscd,a.irrfg,a.donorid,a.donoraccdt,a.hosfg, " & _
           " a.realexpdt, a.exprcvid, a.expid, a.exprsncd, a.larc, a.smlc, a.lare, a.smle " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE "
                   
    sSql = sSql & DBW("a.stscd<=", BBSBloodStatus.stsBAG) & " AND a.entdt between '" & EntdtF & "' AND '" & EntdtL & "'"
           
    
    If ABO <> "" Then sSql = sSql & " AND " & DBW("a.abo=", ABO)
    If RH <> "" Then sSql = sSql & " AND " & DBW("a.rh=", RH)
    If CompoCd <> "" Then sSql = sSql & " AND " & DBW("a.compocd=", CompoCd)
    If centercd <> "" Then sSql = sSql & " AND " & DBW("centercd=", centercd)
    
    sSql = sSql & " AND a.compocd=b.compocd " & _
           "ORDER BY a.abo,a.bldsrc,a.bldyy,a.bldno  "

    Debug.Print sSql
    Set GetBlood_AllList = New Recordset

    Call GetBlood_AllList.Open(sSql, DBConn)
'    If GetBlood_AllList.DBerror Then
'        dbconn.DisplayErrors
'        Set GetBlood_AllList = Nothing
'    End If
End Function

Public Function GetBlood_DeliveryList() As Recordset
    Dim sSql    As String
    
    '2005/05/30 modify by legends
    '유니크한 출고 건수를 구하기 위해서 distinct 구문 추가
    
    sSql = " SELECT distinct a.splitoutfg,b.abbrnm as componm,a.compocd,a.bldsrc,a.bldyy,a.bldno, a.coldt,a.entdt || a.enttm as entdt,d.ptid," & _
           " a.available,a.entid,a.volumn,a.abo,a.rh,a.stscd,a.irrfg,a.donorid,a.donoraccdt,a.hosfg," & _
           " a.realexpdt, a.exprcvid, a.expid, a.exprsncd, a.larc, a.smlc, a.lare, a.smle  " & _
           " FROM " & T_BBS202 & " d," & T_BBS402 & " c," & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & _
                   " c.deliverydt between '" & EntdtF & "' AND '" & EntdtL & "'"
    
    If ABO <> "" Then sSql = sSql & " AND " & DBW("a.abo=", ABO)
    If RH <> "" Then sSql = sSql & " AND " & DBW("a.rh=", RH)
    If CompoCd <> "" Then sSql = sSql & " AND " & DBW("a.compocd=", CompoCd)
    If centercd <> "" Then sSql = sSql & " AND " & DBW("a.centercd=", centercd)
    
    If PtId <> "" Then sSql = sSql & " AND " & DBW("d.ptid=", PtId)
    
    sSql = sSql & " AND " & DBW("a.stscd=", BBSBloodStatus.stsDELIVERY)
    
    sSql = sSql & " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd"
    
    sSql = sSql & " AND c.workarea=d.workarea AND c.accdt=d.accdt AND c.accseq=d.accseq"
    
    
    sSql = sSql & " AND a.compocd=b.compocd " & _
           "ORDER BY a.abo,a.bldsrc,a.bldyy,a.bldno  "

    Set GetBlood_DeliveryList = New Recordset
    Call GetBlood_DeliveryList.Open(sSql, DBConn)
End Function

Public Function DeLBloodRs(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Recordset
    Dim sSql As String
    Dim Rs   As Recordset
    Dim DelDt As String
    
    DelDt = MaxDelDt(Bldsrc, Bldyy, Bldno, CompoCd)
    
    '2001-11-29 수정 : 로컬출고된 혈액은 환자처방이 없다. (원내출고와 로컬출고를 UNION)
    '처방 뷰의 풀스캔으로 인하야 통계자료를 사용할 수 없도록 항상 Rule베이스를 사용하기 위해..
    '2005/05/31  modify by legends
    '처방 뷰의 풀스캔을 hash 조인으로 속도 향상을...
    '/*+ hash (b) */
    
'    sSql = " SELECT  b.ptid,c." & F_PTNM & " as ptnm, b.deptcd,b.wardid,to_char(c.BIRTDATE) as ssn," & _
'           " b.orddt,b.ordno,a.deliveryid,a.rcvid,a.deliverydt || a.DELIVERYTM AS deliverydt,'' localnm" & _
'           " FROM " & T_LAB102 & " d," & T_HIS001 & " c," & T_LAB101 & " b," & T_BBS402 & " a" & _
'           " WHERE " & _
'                     DBW("a.bldsrc=", Bldsrc) & _
'           " AND " & DBW("a.bldyy=", Bldyy) & _
'           " AND " & DBW("a.bldno=", Bldno) & _
'           " AND " & DBW("a.compocd=", CompoCd) & " AND " & DBW("b.orddiv=", C_WORKAREA) & _
'           " AND " & DBW("a.deliverydt=", DelDt) & _
'           " AND (a.localcd is null or a.localcd='')" & _
'           " AND d.workarea=a.workarea AND d.accdt =a.accdt AND d.accseq=a.accseq" & _
'           " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
'           " AND c." & F_PTID & "=b.ptid"
           
' 암호화
'    sSql = " SELECT b.ptid,c." & F_PTNM & " as ptnm, b.deptcd,b.wardid,CryptIT.decrypt(resno1,'pmc1898') || CryptIT.decrypt(resno2,'pmc1898') as ssn," & _
'           " b.orddt,b.ordno,a.deliveryid,a.rcvid,a.deliverydt || a.DELIVERYTM AS deliverydt,'' localnm" & _
'           " FROM " & T_LAB102 & " d," & T_HIS001 & " c," & T_LAB101 & " b," & T_BBS402 & " a" & _
'           " WHERE " & _
'                     DBW("a.bldsrc=", Bldsrc) & _
'           " AND " & DBW("a.bldyy=", Bldyy) & _
'           " AND " & DBW("a.bldno=", Bldno) & _
'           " AND " & DBW("a.compocd=", CompoCd) & " AND " & DBW("b.orddiv=", C_WORKAREA) & _
'           " AND " & DBW("a.deliverydt=", DelDt) & _
'           " AND (a.localcd is null or a.localcd='')" & _
'           " AND d.workarea=a.workarea AND d.accdt =a.accdt AND d.accseq=a.accseq" & _
'           " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
'           " AND c." & F_PTID & "=b.ptid"

' 암호화 해제
'    sSql = " SELECT b.ptid,c." & F_PTNM & " as ptnm, b.deptcd,b.wardid,resno1 || resno2 as ssn," & _
'           " b.orddt,b.ordno,a.deliveryid,a.rcvid,a.deliverydt || a.DELIVERYTM AS deliverydt,'' localnm" & _
'           " FROM " & T_LAB102 & " d," & T_HIS001 & " c," & T_LAB101 & " b," & T_BBS402 & " a" & _
'           " WHERE " & _
'                     DBW("a.bldsrc=", Bldsrc) & _
'           " AND " & DBW("a.bldyy=", Bldyy) & _
'           " AND " & DBW("a.bldno=", Bldno) & _
'           " AND " & DBW("a.compocd=", CompoCd) & " AND " & DBW("b.orddiv=", C_WORKAREA) & _
'           " AND " & DBW("a.deliverydt=", DelDt) & _
'           " AND (a.localcd is null or a.localcd='')" & _
'           " AND d.workarea=a.workarea AND d.accdt =a.accdt AND d.accseq=a.accseq" & _
'           " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
'           " AND c." & F_PTID & "=b.ptid"
           
' 신규암호화
    sSql = " SELECT b.ptid,c." & F_PTNM & " as ptnm, b.deptcd,b.wardid,crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno2) as ssn," & _
           " b.orddt,b.ordno,a.deliveryid,a.rcvid,a.deliverydt || a.DELIVERYTM AS deliverydt,'' localnm" & _
           " FROM " & T_LAB102 & " d," & T_HIS001 & " c," & T_LAB101 & " b," & T_BBS402 & " a" & _
           " WHERE " & _
                     DBW("a.bldsrc=", Bldsrc) & _
           " AND " & DBW("a.bldyy=", Bldyy) & _
           " AND " & DBW("a.bldno=", Bldno) & _
           " AND " & DBW("a.compocd=", CompoCd) & " AND " & DBW("b.orddiv=", C_WORKAREA) & _
           " AND " & DBW("a.deliverydt=", DelDt) & _
           " AND (a.localcd is null or a.localcd='')" & _
           " AND d.workarea=a.workarea AND d.accdt =a.accdt AND d.accseq=a.accseq" & _
           " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
           " AND c." & F_PTID & "=b.ptid"

    sSql = sSql & " AND a.deliveryseq=( SELECT max(z.deliveryseq) FROM " & T_BBS402 & " z" & _
                                      " WHERE " & _
                                                DBW("z.bldsrc=", Bldsrc) & _
                                      " AND " & DBW("z.bldyy=", Bldyy) & _
                                      " AND " & DBW("z.bldno=", Bldno) & _
                                      " AND " & DBW("z.compocd=", CompoCd) & _
                                      " AND " & DBW("z.deliverydt=", DelDt) & ")"
        
    sSql = sSql & " union " & _
           " SELECT null ptid, '' as ptnm, '' deptcd,'' wardid,'' as ssn," & _
           " '' orddt,0 ordno,a.deliveryid,a.rcvid,a.deliverydt || a.DELIVERYTM AS deliverydt,b.field1 localnm" & _
           " FROM " & T_COM003 & " b, " & T_BBS402 & " a" & _
           " WHERE " & _
                     DBW("a.bldsrc=", Bldsrc) & _
           " AND " & DBW("a.bldyy=", Bldyy) & _
           " AND " & DBW("a.bldno=", Bldno) & _
           " AND " & DBW("a.compocd=", CompoCd) & _
           " AND " & DBW("a.deliverydt=", DelDt) & _
           " AND a.localcd is not null AND  a.localcd<>' ' " & _
           " AND " & DBW("b.cdindex", BC2_LOCAL, 2) & _
           " AND b.cdval1 = a.localcd "

    sSql = sSql & " AND a.deliveryseq=( SELECT max(z.deliveryseq) FROM " & T_BBS402 & " z" & _
                                      " WHERE " & _
                                                DBW("z.bldsrc=", Bldsrc) & _
                                      " AND " & DBW("z.bldyy=", Bldyy) & _
                                      " AND " & DBW("z.bldno=", Bldno) & _
                                      " AND " & DBW("z.compocd=", CompoCd) & _
                                      " AND " & DBW("z.deliverydt=", DelDt) & ")"
        
    Set DeLBloodRs = New Recordset
    DeLBloodRs.Open sSql, DBConn
    
End Function

Private Function MaxDelDt(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As String
    Dim Rs   As Recordset
    Dim sSql As String
    
    MaxDelDt = ""
    
    sSql = " SELECT max(deliverydt) as vfydt FROM " & T_BBS402 & _
           " WHERE" & _
                     DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy=", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd)
          
    Set Rs = New Recordset
    Rs.Open sSql, DBConn

    MaxDelDt = Rs.Fields("vfydt").Value & ""

End Function

Public Function AssignBloodRs(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Recordset
    Dim sSql    As String
    Dim VfyDt   As String
    Dim StatChk As Boolean
    
    VfyDt = MaxVfyDt(Bldsrc, Bldyy, Bldno, CompoCd)
    StatChk = False
    If Len(VfyDt) > 8 Then StatChk = True
    
    VfyDt = Mid(VfyDt, 1, 8)
            
            
    sSql = " SELECT b.ptid,c." & F_PTNM & " as ptnm, b.deptcd,b.wardid," & F_SSN2("c") & " as ssn," & _
           " b.orddt,b.ordno,a.vfydt,a.vfyid,a.rstv,a.stat" & _
           " FROM " & T_LAB102 & " d," & T_HIS001 & " c," & T_LAB101 & " b," & T_BBS302 & " a" & _
           " WHERE " & _
                     DBW("a.bldsrc=", Bldsrc) & _
           " AND " & DBW("a.bldyy=", Bldyy) & _
           " AND " & DBW("a.bldno=", Bldno) & _
           " AND " & DBW("a.compocd=", CompoCd) & _
           " AND " & DBW("b.orddiv=", C_WORKAREA) & _
           " AND d.workarea=a.workarea AND d.accdt =a.accdt AND d.accseq=a.accseq" & _
           " AND b.ptid=d.ptid AND b.orddt=d.orddt AND b.ordno=d.ordno" & _
           " AND c." & F_PTID & "=b.ptid  AND (a.cancelfg = '' or a.cancelfg is null or a.cancelfg='0') "
                                       '--> 2001-11-12 추가

           
    If StatChk = False Then
        sSql = sSql & " AND " & DBW("a.vfydt=", VfyDt) & _
                      " AND a.vfytm=( SELECT max(z.vfytm) FROM " & T_BBS302 & " z" & _
                                    " WHERE " & _
                                            DBW("z.bldsrc=", Bldsrc) & _
                                    " AND " & DBW("z.bldyy=", Bldyy) & _
                                    " AND " & DBW("z.bldno=", Bldno) & _
                                    " AND " & DBW("z.compocd=", CompoCd) & _
                                    " AND " & DBW("z.vfydt=", VfyDt) & ")"
    Else
        sSql = sSql & " AND " & DBW("a.statdt=", VfyDt) & _
                      " AND a.stattm=( SELECT max(z.stattm) FROM " & T_BBS302 & " z" & _
                                     " WHERE " & _
                                               DBW("z.bldsrc=", Bldsrc) & _
                                     " AND " & DBW("z.bldyy=", Bldyy) & _
                                     " AND " & DBW("z.bldno=", Bldno) & _
                                     " AND " & DBW("z.compocd=", CompoCd) & _
                                     " AND " & DBW("z.stat=", VfyDt) & ")"
    End If
        
    Set AssignBloodRs = New Recordset
    AssignBloodRs.Open sSql, DBConn
        
End Function

Private Function MaxVfyDt(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As String
    Dim sSql As String
    Dim Rs   As Recordset
    Dim RsS  As Recordset
    
    MaxVfyDt = ""
    
    sSql = " SELECT max(vfydt) as vfydt FROM " & T_BBS302 & _
           " WHERE" & _
                     DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy=", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd)
          
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        MaxVfyDt = Rs.Fields("vfydt").Value & ""
    Else
        sSql = " SELECT max(statdt) as statdt FROM " & T_BBS302 & _
               " WHERE" & _
                         DBW("bldsrc=", Bldsrc) & _
               " AND " & DBW("bldyy=", Bldyy) & _
               " AND " & DBW("bldno=", Bldno) & _
               " AND " & DBW("compocd=", CompoCd) & _
               " AND " & DBW("stat=", "1")
        Set RsS = New Recordset
        RsS.Open sSql, DBConn
        If Not RsS.EOF Then
            MaxVfyDt = RsS.Fields("statdt").Value & "" & "S"
        End If
        Set RsS = Nothing
    End If
    
    Set Rs = Nothing
    
End Function

Public Function IpGoPossibleFg(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Boolean
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT stscd FROM " & T_BBS401 & _
           " WHERE " & _
                     DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy=", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        If Rs.Fields("stscd").Value & "" = "0" Or Rs.Fields("stscd").Value & "" = "1" Then
            IpGoPossibleFg = True
        End If
    End If
    
    Set Rs = Nothing
End Function

Public Function IpGoCancel(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As String
    IpGoCancel = " DELETE " & T_BBS401 & _
                 " WHERE " & _
                           DBW("bldsrc=", Bldsrc) & _
                 " AND " & DBW("bldyy=", Bldyy) & _
                 " AND " & DBW("bldno=", Bldno) & _
                 " AND " & DBW("compocd=", CompoCd)
End Function

'헌혈입고된혈액의 부적격여부를 확인한다.
'donorid,donoraccdt
Public Function GetBloodOkNotFg(ByVal Donorid As String, ByVal donoraccdt As String) As String
    GetBloodOkNotFg = " SELECT * FROM " & T_BBS603 & _
                      " WHERE " & DBW("donorid=", Donorid) & _
                              " AND " & DBW("donoraccdt=", donoraccdt)
End Function

Public Function GetBlood_SubList(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As String
    Dim strSQL      As String
    
'    strSQL = " SELECT c.stat, c.rstv, c.vfydt, c.vfyid, a.retdt, a.retid, a.retrcvid " & _
             "   FROM " & T_BBS402 & " a, " & T_BBS302 & " c " & _
             "  WHERE c.bldsrc = " & DBS(Bldsrc) & " AND c.bldyy = " & DBS(Bldyy) & " AND bldno = " & DBN(Bldno) & " AND c.compocd = " & DBN(CompoCd) & _
             "    AND a.bldsrc = c.bldsrc and a.bldyy = c.bldyy and a.bldno = c.bldno and a.compocd = c.compocd "
    
    strSQL = " select a.retdt, a.retid, a.retrcvid " & _
             "   from " & T_BBS402 & " a " & _
             "  WHERE a.bldsrc = " & DBS(Bldsrc) & " AND a.bldyy = " & DBS(Bldyy) & " AND a.bldno = " & DBN(Bldno) & " AND a.compocd = " & DBN(CompoCd)
    
    GetBlood_SubList = strSQL
    
End Function

Private Sub Class_Initialize()
    Call Clear
End Sub

Public Function UpdateVolumn401(ByVal pVolumn As String, ByVal pBldsrc As String, _
                                ByVal pBldyy As String, ByVal pBldno As String, ByVal pCompocd As String) As String
    UpdateVolumn401 = " UPDATE " & T_BBS401 & _
                      "    SET " & _
                                   DBW("volumn", pVolumn, 2) & _
                      "  WHERE " & DBW("bldsrc", pBldsrc, 2) & _
                      "    AND " & DBW("bldyy", pBldyy, 2) & _
                      "    AND " & DBW("bldno", pBldno, 2) & _
                      "    AND " & DBW("compocd", pCompocd, 2)
End Function

Public Function GetbIOInfo(ByVal bIOidx As Integer, ByVal bSTdt As String, ByVal bEDdt As String) As Recordset
    Dim sSql As String

    If bIOidx Then  ' 출고
sSql = "select /*+ ORDERED *//*+ INDEX(d s2bbs402_idx3) */ distinct a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bNo, c.field2 as jjCd, c.field3 as jjnm, " & _
       " decode(a.stscd,'3',1,'4',2) as boType," & _
       " to_char(to_date(d.deliverydt,'YYYY-MM-DD'),'YYYY-MM-DD') as biDT, " & _
       " to_char(to_date(substr(d.deliverytm,1,4),'HH24:MI'),'HH24:MI') as biTM, " & _
       " decode(a.abo||''||a.rh,'O+',1,'A+',2,'B+',3,'AB+',4,'O-',5,'A-',6,'B-',7,'AB-',8) as bType, " & _
       " a.abo||'('||a.rh||')' as bTnm, b.empnm as EMPnm  from s2bbs402 d, s2bbs401 a, s2com006 b, s2com003 c " & _
        " where d.deliverydt >= '" & bSTdt & "' and d.deliverydt <= '" & bEDdt & "' and a.bldsrc = d.bldsrc and a.bldyy = d.bldyy and a.bldno = d.bldno " & _
       " and d.deliveryid (+)= b.empid and " & _
       " a.ptid is null and d.deliveryid = b.empid and a.stscd in (3,4) and a.splitoutfg = '0' and " & _
       " a.splitinfg = '0' and  c.cdindex = 'B301' and c.field1 = a.compocd " 'order by  bidt, BITM "

    Else            ' 입고
sSql = "select /*+ ORDERED *//*+ INDEX(d s2bbs401_idx2) */ distinct a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bNo, c.field2 as jjCd, c.field3 as jjnm, " & _
        " to_char(to_date(a.entdt,'YYYY-MM-DD'),'YYYY-MM-DD') as biDT, " & _
        " to_char(to_date(substr(a.enttm,1,4),'HH24:MI'),'HH24:MI') as biTM, " & _
        " to_char(to_date(a.coldt,'YYYY-MM-DD'),'YYYY-MM-DD') as bCdDT, " & _
        " decode(a.abo||''||a.rh,'O+',1,'A+',2,'B+',3,'AB+',4,'O-',5,'A-',6,'B-',7,'AB-',8) as bType, " & _
        " a.abo||'('||a.rh||')' as bTnm, b.empnm as EMPnm  from s2bbs401 a, s2com006 b, s2com003 c " & _
        " where a.entdt >= '" & bSTdt & "' and a.entdt <= '" & bEDdt & "' and " & _
        " a.ptid is null and a.entid = b.empid and a.splitoutfg = '0' and " & _
        " a.splitinfg = '0' and c.cdindex = 'B301' and c.field1 = a.compocd " 'order by a.entdt, a.enttm "

    
    
    End If
    
    Set GetbIOInfo = New Recordset
    GetbIOInfo.Open sSql, DBConn
    
End Function


Public Function GetbIOCount(ByVal bIOidx As Integer, ByVal bSTdt As String, ByVal bEDdt As String) As Integer
    Dim Rs   As Recordset
    Dim sSql As String

    If bIOidx Then  ' 출고
sSql = "select count(bNo) as cnt from (select a.bldsrc||'-'||a.bldyy||'-'||to_char(a.bldno,'000000') as bNo from s2bbs401 a, s2com006 b, s2com003 c " & _
        " where a.entdt >= '" & bSTdt & "' and a.entdt <= '" & bEDdt & "' and " & _
       " a.entid = b.empid and " & _
       " a.ptid is null and a.entid = b.empid and a.stscd in (3,4) and a.splitoutfg = '0' and " & _
       " a.splitinfg = '0' and  c.cdindex = 'B301' and c.field1 = a.compocd order by a.entdt, a.enttm) "

    Else
sSql = "select count(bNo) as cnt from (select a.bldsrc||'-'||a.bldyy||'-'||to_char(a.bldno,'000000') as bNo from s2bbs401 a, s2com006 b, s2com003 c " & _
        " where a.entdt >= '" & bSTdt & "' and a.entdt <= '" & bEDdt & "' and " & _
        " a.ptid is null and a.entid = b.empid and a.splitoutfg = '0' and " & _
        " a.splitinfg = '0' and c.cdindex = 'B301' and c.field1 = a.compocd order by a.entdt, a.enttm) "

    End If
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        GetbIOCount = Rs.Fields("cnt").Value
    Else
        GetbIOCount = 0
    End If
    Set Rs = Nothing
End Function

Public Function GetBlood_BMS_LocalOutPut() As Recordset   'LOCAL 출고데이타 Reading.... 2014-09-17

    Dim sSql    As String

    sSql = "       SELECT a.deliverydt, a.bldsrc, a.bldyy, a.bldno, a.compocd, 'PRC' AS componm, a.deliverytm, b.abo, '' AS rh, b.coldt " & vbCrLf
    sSql = sSql & "     , b.entdt       " & vbCrLf      '유효일자
    sSql = sSql & "     , b.entid       " & vbCrLf      '입고자
    sSql = sSql & "     , b.STSCD       " & vbCrLf      '출고
    sSql = sSql & "     , '' AS ptid    " & vbCrLf      '차트번호
    sSql = sSql & "     , a.deliveryid  " & vbCrLf      '출고자"
    sSql = sSql & "     , a.rcvid       " & vbCrLf      '인수자"
    sSql = sSql & "     , b.volumn      " & vbCrLf
    sSql = sSql & "     , a.localcd     " & vbCrLf
    sSql = sSql & "     , b.available           " & vbCrLf
    sSql = sSql & "     , c.field1 AS localnm   " & vbCrLf
    sSql = sSql & "FROM   s2bbs402 a,  s2bbs401 b, S2COM003 c " & vbCrLf
    sSql = sSql & "WHERE  b.bldsrc = a.bldsrc AND   b.bldyy = a.bldyy  AND   b.bldno = a.bldno  AND   b.compocd = a.compocd " & vbCrLf
    sSql = sSql & "AND    c.cdindex  ='B006' AND c.cdval1 = a.localcd " & vbCrLf
    sSql = sSql & "AND    a.deliverydt between  '" & EntdtF & "' AND '" & EntdtL & "' " & vbCrLf
    sSql = sSql & "AND    a.compocd = '02' " & vbCrLf
    sSql = sSql & "AND    b.STSCD = '3'  and (a.localcd is not null or a.localcd<>' ') "
    sSql = sSql & "UNION  ALL " & vbCrLf
    sSql = sSql & "SELECT distinct c.deliverydt, a.bldsrc,a.bldyy,a.bldno,a.compocd,b.abbrnm as componm, c.deliverytm,a.abo,a.rh, a.coldt " & vbCrLf
    sSql = sSql & "     , a.entdt       " & vbCrLf      '유효일자
    sSql = sSql & "     , a.entid       " & vbCrLf      '입고자
    sSql = sSql & "     , a.STSCD       " & vbCrLf      '출고
    sSql = sSql & "     , d.ptid    " & vbCrLf      '차트번호
    sSql = sSql & "     , c.deliveryid  " & vbCrLf      '출고자"
    sSql = sSql & "     , c.rcvid       " & vbCrLf      '인수자"
    sSql = sSql & "     , a.volumn      " & vbCrLf
    sSql = sSql & "     , '' as localcd     " & vbCrLf
    sSql = sSql & "     , a.available           " & vbCrLf
    sSql = sSql & "     , '' as localnm   " & vbCrLf
    sSql = sSql & "FROM   s2bbs202 d,s2bbs402 c,s2bbs006 b,s2bbs401 a " & vbCrLf
    sSql = sSql & "WHERE  c.deliverydt between  '" & EntdtF & "' AND '" & EntdtL & "' " & vbCrLf
    sSql = sSql & "AND    a.stscd = '4' AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno " & vbCrLf
    sSql = sSql & "AND    a.compocd = '02' AND a.compocd=c.compocd AND c.workarea=d.workarea AND c.accdt=d.accdt AND c.accseq=d.accseq AND a.compocd=b.compocd " & vbCrLf
    If centercd <> "" Then sSql = sSql & " AND " & DBW("a.centercd=", centercd)
    
    Debug.Print sSql
    Set GetBlood_BMS_LocalOutPut = New Recordset
    GetBlood_BMS_LocalOutPut.Open sSql, DBConn
   
'    Call GetBlood_BMS_LocalOutPut.RsOpen(, sSql)
'    If GetBlood_BMS_LocalOutPut.Err Then
'        DBConn.DisplayErrors
'        Set GetBlood_BMS_LocalOutPut = Nothing
'    End If
End Function
'
'Public Function GetBlood_BMSList() As Recordset
'
'    Dim sSql    As String
'
'    '## 5.0.2: 이상대(2005-02-22)
'    '   - 교차시험 결과도 조회하기위해 쿼리수정
'    sSql = " SELECT distinct a.splitoutfg,b.abbrnm as componm,a.compocd,a.bldsrc,a.bldyy,a.bldno, a.coldt,a.entdt,d.ptid," & _
'           " a.available,a.entid,a.volumn,a.abo,a.rh,a.stscd,a.irrfg,a.donorid,a.donoraccdt,a.hosfg, " & _
'           " e.step1, e.step2, e.step3, e.step4, e.rmk, e.cancelfg" & _
'           " FROM " & T_BBS302 & " e, " & T_BBS202 & " d," & T_BBS402 & " c," & _
'             T_BBS006 & " b," & T_BBS401 & " a" & _
'           " WHERE c.deliverydt between '" & EntdtF & "' AND '" & EntdtL & "'"
'
'    If ABO <> "" Then sSql = sSql & " AND " & DBW("a.abo=", ABO)
'    If RH <> "" Then sSql = sSql & " AND " & DBW("a.rh=", RH)
'    If CompoCd <> "" Then sSql = sSql & " AND " & DBW("a.compocd=", CompoCd)
'    If centercd <> "" Then sSql = sSql & " AND " & DBW("a.centercd=", centercd)
'
'    If PtId <> "" Then sSql = sSql & " AND " & DBW("d.ptid=", PtId)
'
'    sSql = sSql & " AND A.STSCD IN ('3')"
'    sSql = sSql & " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd"
'    sSql = sSql & " AND c.workarea=d.workarea AND c.accdt=d.accdt AND c.accseq=d.accseq"
'    sSql = sSql & " AND a.compocd=b.compocd " & _
'           "AND a.bldsrc=e.bldsrc AND a.bldyy=e.bldyy AND a.bldno=e.bldno " & _
'           "AND a.compocd=e.compocd AND e.cancelfg<>'1' AND b.abbrnm = 'PRC' " & _
'           "ORDER BY a.abo,a.bldsrc,a.bldyy,a.bldno  "
'
'    Debug.Print sSql
'    Set GetBlood_BMSList = New DrRecordSet
'
'    GetBlood_BMSList.Open sSql, DBConn
'
'    If GetBlood_BMSList.DBerror Then
'        DBConn.DisplayErrors
'        Set GetBlood_BMSList = Nothing
'    End If
'End Function


