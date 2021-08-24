VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Private stsORDER   As String
Private stsCOLLECT As String
Private stsACCESS  As String
Private stsINPROCESS As String
Private stsCancelfg  As String

Private mvarStscd As String

Private mvarActFG As String
Public Property Get stscd() As String
    stscd = mvarStscd
End Property
Public Property Let stscd(ByVal vData As String)
    mvarStscd = vData
End Property
Public Property Let Cancelfg(ByVal vData As String)
    stsCancelfg = vData
End Property
Public Property Let ActFG(ByVal vData As String)
    mvarActFG = vData
End Property

Public Function QueryOrder(ByVal Frdt As String, ByVal Todt As String, ByVal DateDiv As Long, _
                           ByVal PtId As String, ByVal inout As String, _
                           ByVal AccessDiv As String, _
                           Optional wardid As String = "", _
                           Optional ByVal testdiv As String = "") As Recordset
    '-------------------------------------
    'DateDiv = 0 : 수혈예정일로 조회
    '          1 : 처방일로 조회
    'Ptid <> ""  : 특정환자 검색
    '      = ""  : 모든환자 검색
    'inout = 1   : 외래
    '      = 2   : 병동
    'AccessDiv=1 : 미접수제외
    'AccessDiv=0 : 미접수포함
    '-------------------------------------
    Dim sSql As String
    
    '-- 원본 ----------------------------------------------------------------------------------------------
    sSql = " SELECT distinct b.workarea,b.accdt,b.accseq,b.ptid,b.orddt," & _
                   "b.ordno,b.ordseq,b.ordcd,e.testnm,e.testdiv,b.unitqty,a.reqdt,a.reqtm,a.orddiv," & _
                   "a.orddoct,a.majdoct,a.deptcd,a.wardid,a.bedindt,a.bussdiv,b.statfg,b.dcfg,b.stscd,a.hosilid,b.mesg,b.irradfg,b.rcvdt,b.rcvtm " & _
           " FROM " & _
                    T_BBS006 & " c," & T_BBS001 & " e," & T_LAB101 & " a," & T_LAB102 & " b " & _
           " WHERE" & DBW("a.reqdt>=", Frdt) & " AND " & DBW("a.reqdt<=", Todt) & _
                    " AND a.ptid = b.ptid  " & _
                    " AND a.orddt = b.orddt" & _
                    " AND a.ordno = b.ordno" & _
                    " " & _
                    " AND b.ordcd=e.testcd " & _
                    " AND e.compocd=c.compocd" & _
                    " AND e.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                     " WHERE z.testcd=b.ordcd AND z.applydt<=a.orddt)" & _
                    " AND " & DBW("a.orddiv=", C_WORKAREA)
    '-------------------------------------------------------------------------------------------------------
    
'    sSql = " SELECT distinct b.workarea,b.accdt,b.accseq,b.ptid,b.orddt," & _
                   "b.ordno,b.ordseq,b.ordcd,e.testnm,e.testdiv,b.unitqty,a.reqdt,a.reqtm,a.orddiv," & _
                   "a.orddoct,a.majdoct,a.deptcd,a.wardid,a.bedindt,a.bussdiv,b.statfg,b.dcfg,b.stscd,a.hosilid,b.mesg,b.irradfg,b.rcvdt,b.rcvtm " & _
           " FROM " & _
                    T_BBS006 & " c," & T_BBS001 & " e," & T_LAB101 & " a," & T_LAB102 & " b, " & T_HIS002 & " d " & _
           " WHERE" & _
                    "     a.ptid = b.ptid  " & _
                    " AND a.orddt = b.orddt" & _
                    " AND a.ordno = b.ordno" & _
                    " " & _
                    " AND b.ordcd=e.testcd " & _
                    " AND e.compocd=c.compocd" & _
                    " AND e.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                     " WHERE z.testcd=b.ordcd AND z.applydt<=a.orddt)" & _
                    " AND " & DBW("a.orddiv=", C_WORKAREA) & _
                    " AND a.ptid = d.patno(+) AND a.orddt<= d.dschdate(+) " & _
                    " AND a.orddt<=nvl(d.dschdate,to_char(sysdate,'yyyymmdd'))(+) "

    '처방일별
'-->     sSql = sSql & " AND " & DBW("a.reqdt>=", Frdt) & " AND " & DBW("a.reqdt<=", Todt)
    'DC제외
    
    If AccessDiv <> "" Then
        sSql = sSql & " AND (" & DBW("b.dcfg<>", "1") & " or b.dcfg is null)"
    End If
    
    '응급만
    If DateDiv = 1 Then
        sSql = sSql & " AND " & DBW("b.statfg=", DateDiv)
    End If
    '병동/외래별
    If inout = "2" Then
        sSql = sSql & " AND " & DBW("a.bussdiv=", BBSBUSSDIV.stsBed)
        If wardid <> "" Then
            sSql = sSql & " AND " & DBW("a.wardid=", UCase(wardid))
        End If
    ElseIf inout = "1" Then
        sSql = sSql & " AND " & DBW("a.bussdiv=", BBSBUSSDIV.stsNotBed)
        If wardid <> "" Then
            sSql = sSql & " AND " & DBW("a.deptcd=", UCase(wardid))
        End If
    End If
    '병동별
    
'    If wardid <> "" Then
'        SSQL = SSQL & " AND " & DBW("a.wardid=", UCase(wardid))
'    End If

    '환자별
    
    If PtId <> "" Then
        sSql = sSql & " AND " & DBW("a.ptid=", PtId)
    End If
    
    
    '혈액제제별
    If testdiv <> "" Then
        sSql = sSql & " AND " & DBW("c.compocd=", testdiv)
    End If
    '상태별
    If mvarStscd <> "" Then
        sSql = sSql & " AND b.stscd in( " & mvarStscd & ")"
    Else
        sSql = sSql & " AND " & DBW("b.stscd<>", "9")
    End If
   
    If ObjSysInfo.dbtype = 2 Then ' = dbconn.ThisIsMs7 Then
        sSql = sSql & " ORDER BY b.ptid,b.orddt,a.orddiv,b.statfg,a.reqdt,a.reqtm desc "
    Else
        sSql = sSql & " ORDER BY ptid,orddt,orddiv,statfg,reqdt,reqtm desc "
    End If
    
    Set QueryOrder = New Recordset
    Call QueryOrder.Open(sSql, DBConn)
    
    If QueryOrder.EOF Then
'        dbconn.DisplayErrors
        Set QueryOrder = Nothing
    End If
    
End Function
Public Function RequestDelete(ByVal Reqdt As String) As String
    RequestDelete = "delete " & T_BBS204 & " WHERE " & DBW("reqdt<=", Reqdt)
End Function
Public Function RequsetCancel(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                              ByVal pReqdt As String, ByVal pReqTm As String, ByVal Cancelfg As String) As String
    
    RequsetCancel = " update " & T_BBS204 & " set " & _
                              DBW("cancelfg=", Cancelfg) & _
                    " WHERE " & _
                              DBW("workarea=", pWorkarea) & _
                    " AND " & DBW("accdt=", pAccDt) & _
                    " AND " & DBW("accseq=", pAccSeq) & _
                    " AND " & DBW("reqdt=", pReqdt) & _
                    " AND " & DBW("reqtm=", pReqTm)

End Function
Public Function CollectionList(ByVal Frdt As String, ByVal Todt As String) As String
    Dim sSql As String
    
    sSql = " SELECT distinct b.workarea,b.accdt,b.accseq,b.ptid,d." & F_PTNM & " as ptnm,b.orddt,a.ordtm,a.entdt,a.enttm," & _
                   "b.ordno,b.ordseq,b.ordcd,b.unitqty,e.testnm,e.testdiv,a.orddiv," & _
                   "a.orddoct,a.deptcd,a.wardid,a.bedindt,a.bussdiv,b.statfg,b.dcfg," & _
                   "b.stscd,a.hosilid,b.mesg,a.reqdt,a.reqtm,a.reqdt||a.reqtm as SortKey " & _
           " FROM " & _
                    T_BBS006 & " c," & T_HIS001 & " d," & T_BBS001 & " e," & _
                    T_LAB101 & " a," & T_LAB102 & " b " & _
           " WHERE" & _
                    DBW("a.orddiv=", C_WORKAREA) '
    If mvarActFG <> "" Then
        sSql = sSql & " AND " & DBW("a.entdt>=", Frdt) & _
                      " AND " & DBW("a.entdt<=", Todt)
    Else
        sSql = sSql & " AND " & DBW("a.orddt>=", Frdt) & _
                      " AND " & DBW("a.orddt<=", Todt)
    
    End If
                    
    sSql = sSql & _
                " AND " & DBW("b.stscd=", stsCOLLECT) & _
                " AND a.ptid = b.ptid  " & _
                " AND a.orddt = b.orddt" & _
                " AND a.ordno = b.ordno" & _
                " AND a.ptid=d." & F_PTID & _
                " AND b.ordcd=e.testcd " & _
                " AND " & DBJ("e.compocd*=c.compocd") & _
                " AND e.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                     " WHERE z.testcd=b.ordcd AND z.applydt<=a.orddt) AND (b.dcfg='' or b.dcfg is null)"
    sSql = sSql & " ORDER BY SortKey desc"
    CollectionList = sSql
End Function

Public Function QueryAccdt(ByVal Frdt As String, ByVal Todt As String) As String

    Dim sSql As String
    
    sSql = " SELECT b.workarea,b.accdt,b.accseq,b.ptid,d." & F_PTNM & " as ptnm,b.orddt,a.ordtm,a.entdt,a.enttm," & _
                   "b.ordno,b.ordseq,b.ordcd,b.unitqty,e.testnm,e.testdiv,a.orddiv," & _
                   "a.orddoct,a.deptcd,a.wardid,a.bedindt,a.bussdiv,b.statfg,b.dcfg," & _
                   "f.assigncnt,f.assigncancelcnt,f.retcnt,f.expcnt," & _
                   "b.stscd,a.hosilid,b.mesg,a.reqdt,a.reqtm,a.reqdt||a.reqtm as SortKey " & _
           " FROM " & _
                    T_BBS203 & " f," & T_BBS006 & " c," & T_HIS001 & " d," & T_BBS001 & " e," & _
                    T_LAB101 & " a," & T_LAB102 & " b " & _
           " WHERE" & _
                    DBW("a.orddiv=", C_WORKAREA)
                    
                    
    If mvarActFG <> "" Then
        sSql = sSql & " AND " & DBW("a.entdt>=", Frdt) & _
                      " AND " & DBW("a.entdt<=", Todt)
    Else
        sSql = sSql & " AND " & DBW("a.orddt>=", Frdt) & _
                      " AND " & DBW("a.orddt<=", Todt)
    
    End If
    sSql = sSql & _
                    " AND " & DBW("b.workarea=", C_WORKAREA) & _
                    " AND " & DBW("a.orddt>=", Frdt) & _
                    " AND " & DBW("a.orddt<=", Todt) & _
                    " AND (b.dcfg='' or b.dcfg is null)" & _
                    " AND a.ptid = b.ptid  " & _
                    " AND a.orddt = b.orddt" & _
                    " AND a.ordno = b.ordno" & _
                    " AND a.ptid=d." & F_PTID & _
                    " AND b.ordcd=e.testcd " & _
                    " AND b.workarea=f.workarea AND b.accdt=f.accdt AND b.accseq=f.accseq" & _
                    " AND " & DBJ("e.compocd*=c.compocd") & _
                    " AND e.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                     " WHERE z.testcd=b.ordcd AND z.applydt<=a.orddt)"
    sSql = sSql & " ORDER BY SortKey desc"
    QueryAccdt = sSql
End Function

Public Function QueryRequest(ByVal Frdt As String, ByVal Todt As String, ByVal DateDiv As Long, _
                           ByVal PtId As String, ByVal inout As String, _
                           ByVal AccessDiv As String, _
                           Optional wardid As String = "", _
                           Optional ByVal testdiv As String = "") As Recordset
    '-------------------------------------
    'DateDiv = 0 : 수혈예정일로 조회
    '          1 : 처방일로 조회
    'Ptid <> ""  : 특정환자 검색
    '      = ""  : 모든환자 검색
    'inout = 1   : 외래
    '      = 2   : 병동
    'AccessDiv=1 : 미접수제외
    'AccessDiv=0 : 미접수포함
    '-------------------------------------
    Dim sSql As String
    
    
    sSql = " SELECT b.workarea,b.accdt,b.accseq,b.ptid,d." & F_PTNM & " as ptnm,b.orddt," & _
                   "b.ordno,b.ordseq,b.ordcd,b.unitqty,e.testnm,e.testdiv,a.orddiv," & _
                   "a.orddoct,a.deptcd,a.wardid,a.bedindt,a.bussdiv,b.statfg,b.dcfg,b.stscd,a.hosilid,b.mesg,a.reqdt,a.reqtm ,a.reqdt||a.reqtm as SortKey " & _
           " FROM " & _
                    T_BBS006 & " c," & T_HIS001 & " d," & T_BBS001 & " e," & _
                    T_LAB101 & " a," & T_LAB102 & " b " & _
           " WHERE" & _
                    "     a.ptid = b.ptid  " & _
                    " AND a.orddt = b.orddt" & _
                    " AND a.ordno = b.ordno" & _
                    " AND a.ptid=d." & F_PTID & _
                    " AND b.ordcd=e.testcd " & _
                    " AND " & DBJ("e.compocd*=c.compocd") & _
                    " AND e.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                     " WHERE z.testcd=b.ordcd AND z.applydt<=a.orddt)" & _
                    " AND " & DBW("a.orddiv=", C_WORKAREA) & _
                    " AND (b.dcfg='' or b.dcfg is null)"
            
    '처방일별
    sSql = sSql & " AND " & DBW("a.reqdt>=", Frdt) & " AND " & DBW("a.reqdt<=", Todt)
        
    'DC제외
    
    If AccessDiv <> "" Then
        sSql = sSql & " AND (" & DBW("b.dcfg<>", "1") & " or b.dcfg is null)"
    End If
    
    '응급만
    If DateDiv = 1 Then
        sSql = sSql & " AND " & DBW("b.statfg=", DateDiv)
    End If
    '병동/외래별
    If inout = "2" Then
        sSql = sSql & " AND " & DBW("a.bussdiv=", BBSBUSSDIV.stsBed)
        
    ElseIf inout = "1" Then
        sSql = sSql & " AND " & DBW("a.bussdiv=", BBSBUSSDIV.stsNotBed)
    End If
    '병동별
    
    If wardid <> "" Then
        sSql = sSql & " AND " & DBW("a.wardid=", UCase(wardid))
    End If
    '환자별
    
    If PtId <> "" Then
        sSql = sSql & " AND " & DBW("a.ptid=", PtId)
    End If
    
    
    '혈액제제별
    If testdiv <> "" Then
        sSql = sSql & " AND " & DBW("c.compocd=", testdiv)
    End If
    '상태별
    If mvarStscd <> "" Then
        If TRANS_REQUIRE_USED Then
            sSql = sSql & " AND " & DBW("b.stscd >= ", BBSOrdStatus.stsREQUEST)
        Else
            sSql = sSql & " AND " & DBW("b.stscd =", mvarStscd)
        End If
    End If
    sSql = sSql & " ORDER BY SortKey desc"
    
    Set QueryRequest = New Recordset
    QueryRequest.Open sSql, DBConn
    
'    If QueryRequest.DBerror = True Then
'        dbconn.DisplayErrors
'        Set QueryRequest = Nothing
'    End If
End Function

'새로운 테이블의 수혈사유를 구한다.(S2BBS101)
Public Function GetTransReason(ByVal sPtid As String, ByVal sOrddt As String, ByVal sOrdno As String) As String

    Dim i As Long
    Dim sSql As String
    Dim DrRS As Recordset
    Dim Rs As Recordset
    Dim rsncd As String
    Dim rsnNm As String
    Dim OpCd As String
    Dim opNm As String

    If AttributeBBS = False Then
        sSql = " SELECT * " & _
               " FROM  " & T_LAB103 & _
               " WHERE " & _
                                   DBW("ptid=  ", sPtid) & _
                         " AND " & DBW("orddt=", sOrddt) & _
                         " AND " & DBW("ordno=", sOrdno) & _
                         " AND   (" & _
                                    DBW("attrcd=   ", BC2_ORDVAL) & " or " & _
                                    DBW("attrcd=", BC2_REASONVAL) & _
                                ") " & _
               " ORDER BY attrcd desc "
               
               
        Set DrRS = New Recordset
        Call DrRS.Open(sSql, DBConn)
        'Open error-----------------
'        If DrRS.DBerror = True Then
'            dbconn.DisplayErrors
'            Set DrRS = Nothing
'            GetTransReason = ""
'            Exit Function
'        End If
        
        '결과가 없음-----------------
        If DrRS.RecordCount < 1 Then
            Set DrRS = Nothing
            GetTransReason = ""
            Exit Function
        End If
        
        For i = 1 To DrRS.RecordCount
            If Trim(DrRS.Fields("attrcd").Value & "") = "B1" Then
                rsncd = DrRS.Fields("attrval").Value & ""
            ElseIf Trim(DrRS.Fields("attrcd").Value) = "A5" Then
                OpCd = DrRS.Fields("attrval").Value & ""
            End If
            DrRS.MoveNext
        Next i
        Set DrRS = Nothing
        
        
        '수혈사유명------------------
        
        Set DrRS = GetCom003RecordSet(BC2_Trans_Reason, rsncd)
        
        If DrRS Is Nothing Then
            GetTransReason = ""
            Exit Function
        End If
        
        If Trim(DrRS.Fields("field2").Value & "") <> "1" Then
            rsnNm = DrRS.Fields("field1").Value & ""
            Set DrRS = Nothing
            GetTransReason = rsnNm
            Exit Function
        End If
        Set DrRS = Nothing
        
        opNm = GetOPNm(OpCd)
        If opNm = "" Then
            GetTransReason = rsnNm
        Else
            GetTransReason = opNm
        End If
        
'   ====================================================
'   아래는 성모자애에서 수혈사유 테이블을 따로 사용할경우
'   ====================================================
    Else
    
        sSql = " SELECT ordseq,bloodrsn,bloodop FROM " & T_BBS101 & _
               " WHERE " & DBW("ptid=", sPtid) & _
               " AND " & DBW("orddt=", sOrddt) & _
               " AND " & DBW("ordno=", sOrdno) & _
               " ORDER BY ordseq"
        Set Rs = New Recordset
        Call Rs.Open(sSql, DBConn)
'        If Rs.DBerror = True Then
'            dbconn.DisplayErrors
'            Set Rs = Nothing
'            GetTransReason = ""
'            Exit Function
'        End If
        If Not Rs.EOF Then
            rsncd = Rs.Fields("bloodrsn").Value & ""
            OpCd = Rs.Fields("bloodop").Value & ""
            Set Rs = Nothing
    
            '수혈사유명------------------
    
            Set DrRS = GetCom003RecordSet(BC2_Trans_Reason, rsncd)
            
            If DrRS Is Nothing Then
                GetTransReason = ""
                Exit Function
            End If
    
            If Trim(DrRS.Fields("field2").Value & "") <> "1" Then
                rsnNm = DrRS.Fields("field1").Value & ""
                Set DrRS = Nothing
                GetTransReason = rsnNm
                Exit Function
            End If
            
            opNm = GetOPNm(OpCd)
            If opNm = "" Then
                GetTransReason = rsnNm
            Else
                GetTransReason = opNm
            End If
    
            Set DrRS = Nothing
        Else
            Set Rs = Nothing
            GetTransReason = ""
        End If
    
        Set Rs = Nothing
    
    End If
    
End Function

Public Sub GetSpcNoAndStore(ByVal PtId As String, _
                            spcno As String, _
                            storeleg As String, storerow As String, storecol As String, _
                            center As String)
    Dim i       As Long
    Dim sSql    As String
    Dim DrRS    As Recordset
    Dim coldttm As String
    


    Dim strYearChk As String
    Dim strYear    As String
    
    strYear = BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4, 1)
    strYear = "30"
    
    sSql = " SELECT a.spcno,a.spcyy,a.storeleg,a.storerno,a.storecno,a.coldt,a.coltm,b.cdval1 as buildcd, b.field1 as buildnm " & _
           " FROM     " & T_COM003 & " b," & T_BBS201 & " a" & _
           " WHERE    " & DBW("a.ptid", PtId, 2) & _
           " AND " & DBW("a.spcyy=", strYear) & _
           " AND a.spcno=(SELECT max(c.spcno) FROM " & T_BBS201 & " c WHERE " & DBW("c.ptid=", PtId) & _
           " AND " & DBW("c.spcyy=", strYear) & ")" & _
           " AND " & DBW("b.cdindex", BC2_CENTER, 2) & _
           " AND a.buildcd=b.cdval1" & _
           " ORDER BY spcyy desc ,spcno desc"
    
    Set DrRS = New Recordset
    DrRS.Open sSql, DBConn
    
    strYearChk = Mid(Format(GetSystemDate, "YYYYMMDD"), 5)
    If DrRS.EOF Then
        If strYearChk >= "0101" And strYearChk <= "0110" Then
            strYear = Val(strYear) + 9
            sSql = " SELECT a.spcno,a.spcyy,a.storeleg,a.storerno,a.storecno,a.coldt,a.coltm,b.cdval1 as buildcd, b.field1 as buildnm " & _
                   " FROM     " & T_COM003 & " b," & T_BBS201 & " a" & _
                   " WHERE    " & DBW("a.ptid", PtId, 2) & _
                   " AND spcno=(SELECT max(c.spcno) FROM " & T_BBS201 & " c WHERE " & DBW("c.ptid=", PtId) & _
                   " AND " & DBW("c.spcyy=", strYear) & ")" & _
                   " AND " & DBW("b.cdindex", BC2_CENTER, 2) & _
                   " AND a.buildcd=b.cdval1" & _
                   " ORDER BY spcno desc"
           Set DrRS = Nothing
           Set DrRS = New Recordset
           DrRS.Open sSql, DBConn
        End If
    End If
    
    If DrRS.EOF Then
        If strYearChk >= "0101" And strYearChk <= "0110" Then
            strYear = Val(strYear) - 1
            sSql = " SELECT a.spcno,a.spcyy,a.storeleg,a.storerno,a.storecno,a.coldt,a.coltm,b.cdval1 as buildcd, b.field1 as buildnm " & _
                   " FROM     " & T_COM003 & " b," & T_BBS201 & " a" & _
                   " WHERE    " & DBW("a.ptid", PtId, 2) & _
                   " AND spcno=(SELECT max(c.spcno) FROM " & T_BBS201 & " c WHERE " & DBW("c.ptid=", PtId) & _
                   " AND " & DBW("c.spcyy=", strYear) & ")" & _
                   " AND " & DBW("b.cdindex", BC2_CENTER, 2) & _
                   " AND a.buildcd=b.cdval1" & _
                   " ORDER BY spcno desc"
           Set DrRS = Nothing
           Set DrRS = New Recordset
           DrRS.Open sSql, DBConn
        End If
    End If
    
'    If DrRS.DBerror = True Then
'        spcno = ""
'        storeleg = ""
'        storerow = ""
'        storecol = ""
'        center = ""
'        Set DrRS = Nothing
'        Exit Sub
'    End If
    
    If DrRS.RecordCount < 1 Then
        spcno = ""
        storeleg = ""
        storerow = ""
        storecol = ""
        center = ""
    Else
        For i = 1 To DrRS.RecordCount
            If (DrRS.Fields("coldt").Value & "" & DrRS.Fields("coltm").Value & "") > coldttm Then
                coldttm = DrRS.Fields("coldt").Value & "" & DrRS.Fields("coltm").Value & ""
                
                spcno = DrRS.Fields("spcyy").Value & "" & "-" & DrRS.Fields("spcno").Value & ""
                storeleg = DrRS.Fields("storeleg").Value & ""
                storerow = DrRS.Fields("storerno").Value & ""
                storecol = DrRS.Fields("storecno").Value & ""
                center = DrRS.Fields("buildcd").Value & "" & vbTab & DrRS.Fields("buildnm").Value & ""
            End If
            DrRS.MoveNext
        Next i
    End If
    
    Set DrRS = Nothing
End Sub

'2001-12-06 추가 : 최근수혈일 찾기
Public Function GetLatestTrandDt(ByVal PtId As String) As String
    Dim i       As Long
    Dim sSql    As String
    Dim DrRS    As Recordset
    Dim coldttm As String
    
    sSql = " SELECT reqdt " & _
           " FROM   " & T_BBS202 & _
           " WHERE  " & DBW("ptid", PtId, 2) & _
           " AND    " & DBW("stscd >= ", stsINPROCESS) & _
           " AND    cancelfg = '0' " & _
           " ORDER BY reqdt desc "
    

    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
'    If DrRS.DBerror = True Then
'        GetLatestTrandDt = ""
'        Set DrRS = Nothing
'        Exit Function
'    End If
    
    If DrRS.RecordCount < 1 Then
        GetLatestTrandDt = ""
    Else
        GetLatestTrandDt = DrRS.Fields("reqdt").Value & ""
    End If
    
    Set DrRS = Nothing
End Function


'수혈처방 출력에서 미접수된 처방을 접수시키는 문장
'이하(lab101,lab102,bbs201,bbs206) 모두

Public Function Update_Lab101(PtId As String, orddt As String, OrdNo As Long)
    
'    Update_Lab101 = " update " & T_LAB101 & _
'                    " set    " & _
'                               DBW("donefg=", stsACCESS) & _
'                    " WHERE " & _
'                                        DBW("ptid=  ", PtId) & _
'                              " AND " & DBW("orddt=", orddt) & _
'                              " AND " & DBW("ordno=", OrdNo)
    '예수병원용
    Update_Lab101 = " update mdbldort " & _
                    " set    " & _
                               DBW("donefg=", stsACCESS) & _
                    " WHERE " & _
                                        DBW("patno=  ", PtId) & _
                              " AND orddate=to_date( " & orddt & ",'yyyymmdd')" & _
                              " AND " & DBW("ordseqno=", OrdNo)
              
End Function

Public Function Update_Lab102(PtId As String, orddt As String, _
                              OrdNo As Long, OrdSeq As Long, AccFormat As Long, accseq As Long)
    Dim strRcvdt As String
    Dim strrcvtm As String
    Dim strAccDt As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strAccDt = Mid(Format(GetSystemDate, PRESENTDATE_FORMAT), 1, AccFormat)
    
'    Update_Lab102 = " update " & T_LAB102 & _
'                    " set    " & _
'                               DBW("stscd", stsACCESS, 3) & _
'                               DBW("donefg", stsACCESS, 3) & _
'                               DBW("rcvdt", strRcvdt, 3) & _
'                               DBW("rcvtm", strrcvtm, 3) & _
'                               DBW("workarea", C_WORKAREA, 3) & _
'                               DBW("accdt", strAccDt, 3) & _
'                               DBW("accseq", accseq, 2) & _
'                    " WHERE " & _
'                                        DBW("ptid=    ", PtId) & _
'                              " AND " & DBW("orddt=  ", orddt) & _
'                              " AND " & DBW("ordno=  ", OrdNo) & _
'                              " AND " & DBW("ordseq=", OrdSeq)
    '예수병원 용
'상태전송 미포함
'    Update_Lab102 = " update mdbldort " & _
'                    " set    " & _
'                               DBW("stscd", stsACCESS, 3) & _
'                               DBW("donefg", stsACCESS, 3) & _
'                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss') " & _
'                               DBW("workarea", C_WORKAREA, 3) & _
'                               DBW("accdt", strAccDt, 3) & _
'                               DBW("accseq", accseq, 2) & _
'                    " WHERE " & _
'                                        DBW("patno=    ", PtId) & _
'                              " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
'                              " AND " & DBW("ordseqno=  ", OrdNo)
'상태전송 포함
'OCS에 상태를 E로 전송
    Update_Lab102 = " update mdbldort " & _
                    " set    " & _
                               DBW("stscd", stsACCESS, 3) & _
                               DBW("donefg", stsACCESS, 3) & _
                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss') " & _
                               DBW("workarea", C_WORKAREA, 3) & _
                               DBW("accdt", strAccDt, 3) & _
                               DBW("accseq", accseq, 3) & _
                               DBW("procstat", "E", 2) & _
                    " WHERE " & _
                                        DBW("patno=    ", PtId) & _
                              " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                              " AND " & DBW("ordseqno=  ", OrdNo)
End Function
Public Function Insert_BBS201(spcyy As String, spcno As Long, coldt As String, coltm As String, colid As String, _
                              rcvid As String, storeleg As String, storerno As Long, storecno As Long, _
                              PtId As String, busidiv As String, BedIndt As String, orddt As String, _
                              DeptCd As String, BuildCd As String, orgbuildcd As String)
    Dim strRcvdt As String
    Dim strrcvtm As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
    Insert_BBS201 = " insert into " & T_BBS201 & " (" & _
                    " spcyy,spcno,coldt,coltm,colid,rcvdt,rcvtm,rcvid,storeleg,storerno,storecno,ptid," & _
                    " busidiv,bedindt,orddt,deptcd,buildcd,orgbuildcd,localcd,expfg,addfg)" & _
                    " values(" & _
                               DBV("spcyy      ", spcyy, 1) & DBV("spcno", spcno, 1) & _
                               DBV("coldt      ", coldt, 1) & DBV("coltm", coltm, 1) & _
                               DBV("colid      ", colid, 1) & DBV("rcvdt", strRcvdt, 1) & _
                               DBV("rcvtm   ", strrcvtm, 1) & DBV("rcvid", rcvid, 1) & _
                               DBV("storeleg", storeleg, 1) & DBV("storerno", storerno, 1) & _
                               DBV("storecno", storecno, 1) & DBV("ptid", PtId, 1) & _
                               DBV("busidiv  ", busidiv, 1) & DBV("bedindt", BedIndt, 1) & _
                               DBV("orddt      ", orddt, 1) & DBV("deptcd", DeptCd, 1) & _
                               DBV("buildcd  ", BuildCd, 1) & DBV("orgbuildcd", orgbuildcd, 1) & _
                               DBV("localcd       ", "", 1) & DBV("expfg", "0", 1) & _
                               DBV("addfg", "0") & _
                           ")"
End Function
                              

Public Function Update_BBS201(spcyy As String, spcno As Long, rcvid As String, Leg As String, Row As Long, Col As Long)
    Dim strRcvdt As String
    Dim strrcvtm As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
     
    Update_BBS201 = " update " & T_BBS201 & _
                    " set   " & _
                              DBW("rcvdt", strRcvdt, 3) & _
                              DBW("rcvtm", strrcvtm, 3) & _
                              DBW("rcvid   ", rcvid, 3) & _
                              DBW("storeleg  ", Leg, 3) & _
                              DBW("storerno  ", Row, 3) & _
                              DBW("storecno  ", Col, 2) & _
                    " WHERE " & _
                                      DBW("spcyy=", spcyy) & _
                            " AND " & DBW("spcno=", spcno)
                    
End Function
Public Function Update_BBS206(centercd As String, legcd As String, Row As Long, Col As Long, spcyy As String, spcno As Long)

    Update_BBS206 = " update " & T_BBS206 & _
                    " set    " & _
                               DBW("spcyy", spcyy, 3) & _
                               DBW("spcno", spcno, 3) & _
                               DBW("stscd", BBSSaveStatue.stsNotUsed, 2) & _
                   " WHERE   " & _
                                         DBW("centercd=", centercd) & _
                               " AND " & DBW("legcd=      ", legcd) & _
                               " AND " & DBW("rowno=        ", Row) & _
                               " AND " & DBW("colno=        ", Col)
End Function

Public Function Get_spcTime(ByVal spcyy As String, ByVal spcno As String) As String
    Get_spcTime = "SELECT coldt,coltm FROM " & T_BBS201 & _
                  " WHERE " & DBW("spcyy=", spcyy) & _
                  " AND " & DBW("spcno=", spcno)
End Function
Public Function Get_ExistSPC(ByVal spcyy As String, ByVal spcno As String) As String
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT stscd FROM " & T_BBS206 & " WHERE " & DBW("spcyy=", spcyy) & " AND " & DBW("spcno=", spcno)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        Get_ExistSPC = Rs.Fields("stscd").Value & ""
    Else
        Get_ExistSPC = ""
    End If
    Set Rs = Nothing
End Function


Public Function GetRelationOrder(ByVal PtId As String, ByVal orddt As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT c.testnm,a.workarea,a.accdt, a.accseq ,a.ordno ,a.stscd,a.unitqty" & _
        " FROM " & T_BBS001 & " c," & T_LAB101 & " b," & T_LAB102 & " a" & _
        " WHERE " & _
                 DBW("a.ptid=", PtId) & _
         " AND " & DBW("a.orddt=", orddt) & _
         " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno" & _
         " AND c.testcd=a.ordcd" & _
         " AND " & DBW("b.orddiv=", C_WORKAREA) & _
         " AND c.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                         " WHERE z.testcd=a.ordcd AND z.applydt<=a.orddt)"
    
    Set GetRelationOrder = New Recordset
    GetRelationOrder.Open sSql, DBConn
    
End Function


Public Function GetKeepHour() As Long
    Dim DrRS As Recordset
    

    Set DrRS = OpenRecordSetDay(BC2_KEEP_HOUR)
    If DrRS Is Nothing Then
        GetKeepHour = 0
    End If
    With DrRS
        If .RecordCount < 1 Then
            GetKeepHour = 0
        Else
            GetKeepHour = .Fields("field1").Value & ""
        End If
    End With
    Set DrRS = Nothing
End Function

'공혈자를 선택하였을때 Donor의 id와 성명을 보여준다.
Public Function DonorInformation(ByVal PtId As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT a.tmpid,b.donornm,d.okdiv3 " & _
           " FROM " & T_BBS603 & " d," & T_BBS602 & " a," & T_BBS601 & " b" & _
           " WHERE " & _
                      DBW("a.reservedid=", PtId) & _
           " AND a.donorid=b.donorid" & _
           " AND a.donorid=d.donorid" & _
           " AND a.donoraccdt=d.donoraccdt"
           
    Set DonorInformation = New Recordset
    DonorInformation.Open sSql, DBConn
    
          
End Function

Public Function TransRequest_Contents(ByVal Frdt As String, ByVal Todt As String, _
                                      Optional ByVal statfg As String = "", _
                                      Optional ByVal BuildCd As String = "") As Recordset
    Dim sSql As String

    sSql = " SELECT b.workarea,b.accdt,b.accseq,b.ptid,b.orddt," & _
                   "b.ordno,b.ordseq,b.ordcd,b.unitqty,e.testnm,e.testdiv,f.reqcnt,f.reqdt,f.reqtm,a.orddiv," & _
                   "a.orddoct,a.deptcd,a.wardid,a.bedindt,a.bussdiv,f.cancelfg,b.statfg,b.dcfg,b.stscd,a.hosilid,b.mesg,f.statfg as stat," & _
                   "g.spcyy,g.spcno,g.buildcd,g.assigncnt,g.assigncancelcnt,g.retcnt,g.expcnt,g.storeleg,g.storerno,g.storecno,g.deliverycnt," & _
                   "f.reqdt+f.reqtm as SortKey " & _
           " FROM " & _
                    T_BBS203 & " g," & T_BBS006 & " c," & T_BBS001 & " e," & _
                    T_BBS204 & " f," & T_LAB101 & " a," & T_LAB102 & " b " & _
           " WHERE" & DBW("f.reqdt>=", Frdt) & _
                    " AND " & DBW("f.reqdt<=", Todt) & _
                    " AND a.ptid = b.ptid  " & _
                    " AND a.orddt = b.orddt" & _
                    " AND a.ordno = b.ordno" & _
                    " AND b.ordcd=e.testcd " & _
                    " AND " & DBJ("e.compocd*=c.compocd") & _
                    " AND e.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                     " WHERE z.testcd=b.ordcd AND z.applydt<=a.orddt)" & _
                    " AND " & DBW("a.orddiv=", C_WORKAREA) & _
                    " AND b.workarea = f.workarea " & _
                    " AND b.accdt = f.accdt " & _
                    " AND b.accseq = f.accseq " & _
                    " AND b.workarea=g.workarea" & _
                    " AND b.accdt=g.accdt" & _
                    " AND b.accseq=g.accseq"
            
    If statfg <> "" Then
        sSql = sSql & " AND f.cancelfg='1'"
    Else
        sSql = sSql & " AND (f.cancelfg='' or f.cancelfg is null)"
    End If
    '빌딩정보
    If BuildCd <> "" Then sSql = sSql & " AND " & DBW("g.buildcd=", BuildCd)
   
    
    sSql = sSql & " ORDER BY SortKey desc"
    
    Set TransRequest_Contents = New Recordset
    TransRequest_Contents.Open sSql, DBConn
    
'    If TransRequest_Contents.DBerror = True Then
'        dbconn.DisplayErrors
'        Set TransRequest_Contents = Nothing
'    End If
End Function

Public Function GetAccWorkLoad(ByVal AccNo As String) As String

    Dim Rs          As Recordset
    Dim strTmp      As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim sSql        As String
    
    strAccDt = medGetP(AccNo, 1, "-")
    strAccSeq = medGetP(AccNo, 2, "-")
    
    sSql = " SELECT deliverycnt,assigncnt,retcnt,expcnt " & _
           " FROM " & T_BBS203 & _
           " WHERE " & DBW("workarea=", C_WORKAREA) & _
           " AND   " & DBW("accdt=", strAccDt) & _
           " AND   " & DBW("accseq=", strAccSeq)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        strTmp = "  결과수량 : " & Rs.Fields("assigncnt").Value & "" & Space(10) & _
                 "출고수량 : " & Rs.Fields("deliverycnt").Value & "" & vbCrLf & _
                 "  반환수량 : " & Rs.Fields("retcnt").Value & "" & Space(10) & _
                 "폐기수량 : " & Rs.Fields("expcnt").Value & ""
    End If
    GetAccWorkLoad = strTmp
    Set Rs = Nothing
End Function


Private Sub Class_Initialize()
    If TRANS_REQUIRE_USED = True Then
        stsORDER = BBSOrdStatus.stsORDER
        stsCOLLECT = BBSOrdStatus.stsCOLLECT
        stsACCESS = BBSOrdStatus.stsACCESS
        stsINPROCESS = BBSOrdStatus.stsINPROCESS
    Else
        stsORDER = BBSOrderStatus.stsORDER
        stsCOLLECT = BBSOrderStatus.stsCOLLECT
        stsACCESS = BBSOrderStatus.stsACCESS
        stsINPROCESS = BBSOrderStatus.stsINPROCESS
    End If
End Sub



