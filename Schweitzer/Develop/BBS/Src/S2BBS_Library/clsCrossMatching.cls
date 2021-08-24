VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrossMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private stsORDER   As String
Private stsCOLLECT As String
Private stsACCESS  As String
Private stsINPROCESS As String

Private mvarStrTmp As String
Private mvarRstSeq As Long
Private mvarAssignCnt As Long   'Assign 된갯수
Private mvarCancelCnt As Long   'Assign 취소수량
Private mvarOutCnt As Long      '출고된갯수
Private mvarRetCnt As Long      '반환수량
Private mvarExpCnt As Long      '폐기수량
Private mvarTotCnt As Long      '총갯수
Private mvarRedimCnt As Long
'2001-11-29추가
Private mvarRequestCnt As Long   '요청된갯수
Private mvarReturnCompoCd As String '

Public Property Get RedimCnt() As Long
    RedimCnt = mvarRedimCnt
End Property

Public Property Get TotCnt() As Long
    TotCnt = mvarTotCnt
End Property
Public Property Get OutCnt() As Long
    OutCnt = mvarOutCnt
End Property
Public Property Get AssignCnt() As Long
    AssignCnt = mvarAssignCnt
End Property
Public Property Get RetCnt() As Long
    RetCnt = mvarRetCnt
End Property
Public Property Get ExpCnt() As Long
    ExpCnt = mvarExpCnt
End Property
Public Property Get CancelCnt() As Long
    CancelCnt = mvarCancelCnt
End Property
'2001-11-29추가
Public Property Get RequestCnt() As Long
    RequestCnt = mvarRequestCnt
End Property

Public Property Get RstSeq() As Long
    RstSeq = mvarRstSeq
End Property
Public Property Get strTmp() As String
    strTmp = mvarStrTmp
End Property

Public Property Get ReturnCompoCd() As String
    ReturnCompoCd = mvarReturnCompoCd
End Property

Public Property Let ReturnCompoCd(ByVal vData As String)
    mvarReturnCompoCd = vData
End Property

'Sql 문(이하)
'######
Public Function Assign_CancleBBS203(ByVal accdt As String, ByVal accseq As String, ByVal CancelCnt As Long)
'Assign 취소시 Assign취소갯수를 Update 해준다.
    Dim Rs As New Recordset
    Dim strSQL As String
'    Dim AssignCancelCnt As Long
    
    strSQL = " SELECT assigncancelcnt FROM " & T_BBS203 & _
             " WHERE" & _
             "     " & DBW("workarea", C_WORKAREA, 2) & _
             " AND " & DBW("accdt", accdt, 2) & _
             " AND " & DBW("accseq", accseq, 2)
             
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    CancelCnt = CancelCnt + Val(Rs.Fields("assigncancelcnt").Value & "")
    Set Rs = Nothing
    
    Assign_CancleBBS203 = " update " & T_BBS203 & " set " & DBW("assigncancelcnt", CancelCnt, 2) & _
                          " WHERE " & _
                          "      " & DBW("workarea", C_WORKAREA, 2) & _
                          "  AND " & DBW("accdt", accdt, 2) & _
                          "  AND " & DBW("accseq", accseq, 2)

End Function
Public Function Insert_BBS203(ByVal accdt As String, ByVal accseq As Long) As String

'결과등록시 등록하고자 하는 혈액의 갯수만큼 Insert or Update 해준다.
'테이블에 해당환자의 존재여부를 먼저 체크한다.
'이미 존재하면, Update 해준다.
    Dim Rs        As New Recordset
    Dim sSql      As String
    Dim AssignCnt As Long
    Dim cnt       As Long: cnt = 1
    
    sSql = " SELECT assigncnt FROM " & T_BBS203 & _
           " WHERE " & _
                   "      " & DBW("workarea", C_WORKAREA, 2) & _
                   "  AND " & DBW("accdt", accdt, 2) & _
                   "  AND " & DBW("accseq", accseq, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF = True Then
        Insert_BBS203 = " insert into " & T_BBS203 & "(workarea,accdt,accseq,assigncnt,deliverycnt," & _
                                                 " retcnt,expcnt,bagcnt,assigncancelcnt)" & _
                       " values(" & _
                                DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("assigncnt", cnt, 1) & _
                                DBV("deliverycnt", "", 1) & DBV("retcnt", "", 1) & DBV("expcnt", "", 1) & DBV("bagcnt", "", 1) & DBV("assigncancelcnt", "") & ")"
    Else
        AssignCnt = cnt + Val(Rs.Fields("assigncnt").Value & "")
        Insert_BBS203 = " update " & T_BBS203 & " set " & DBW("assigncnt", AssignCnt, 2) & _
                        " WHERE " & _
                                "      " & DBW("workarea", C_WORKAREA, 2) & _
                                "  AND " & DBW("accdt", accdt, 2) & _
                                "  AND " & DBW("accseq", accseq, 2)
 
    End If
    Set Rs = Nothing
End Function

Public Function Update_OrderStatus(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Long, _
                                   Optional ByVal OrdSeq As Long = -1)
'결과등록시 처방부분(LAB101,LAB102)의 status를 진행상태로 바꿔준다.(frmBBS102에서 사용함)
'donefg,stscd를 진행중(3)으로 Update 해준다.
    
'    If OrdSeq = -1 Then
'        Update_OrderStatus = " update " & T_LAB101 & " set " & DBW("donefg", stsINPROCESS, 2) & _
'                             " WHERE " & DBW("ptid", PtId, 2) & _
'                             " AND " & DBW("orddt", orddt, 2) & _
'                             " AND " & DBW("ordno", OrdNo, 2)
'    Else
'        Update_OrderStatus = " update " & T_LAB102 & _
'                             " set " & DBW("donefg", stsINPROCESS, 3) & _
'                                       DBW("stscd", stsINPROCESS, 2) & _
'                             " WHERE" & _
'                                    "     " & DBW("ptid", PtId, 2) & _
'                                    " AND " & DBW("orddt", orddt, 2) & _
'                                    " AND " & DBW("ordno", OrdNo, 2) & _
'                                    " AND " & DBW("ordseq", OrdSeq, 2)
'    End If
    
    '전주 예수병원용
    If OrdSeq = -1 Then
        Update_OrderStatus = " update mdbldort set " & DBW("donefg", stsINPROCESS, 2) & _
                             " WHERE " & DBW("patno", PtId, 2) & _
                             " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                             " AND " & DBW("ordseqno", OrdNo, 2)
    '전주예수병원에서만 사용됨..
    'OCS에 상태 전송
    'Assign 갯수와 처방갯수를 비교해서 Assign갯수가 처방갯수보다 크거나 같을때만 R(준비완료)로 업뎃한다.
    'Assign을 할경우에는 준비중(E)상태로 상태업뎃
    ElseIf OrdSeq = -99 Then
        Dim Rs As Recordset
        Dim strSQL As String
        
        strSQL = " select a.unitqty, b.assigncnt from s2ord102_v a, s2bbs203 b"
        strSQL = strSQL & " where " & DBW("a.ptid=", PtId)
        strSQL = strSQL & " and " & DBW("a.orddt=", orddt)
        strSQL = strSQL & " and " & DBW("a.ordno=", OrdNo)
        strSQL = strSQL & " and a.workarea=b.workarea"
        strSQL = strSQL & " and a.accdt=b.accdt"
        strSQL = strSQL & " and a.accseq=b.accseq"
                
        Set Rs = New Recordset
        
        Rs.Open strSQL, DBConn
        
        '2004/12/12 혈액이 준비완료되면 readydt에 현재시간을 박아준다.
        '2005/02/04 혈액이 Assign 될때는 readydt를 박아 줄필요 없다.(",readydt=to_date(" & Format(GetSystemDate, "yyyyMMddhhmmss") & ",'yyyymmdd hh24:mi:ss')" & _)
        If Rs.EOF = False Then
            If Val(Rs.Fields("unitqty").Value & "") <= Val(Rs.Fields("assigncnt").Value & "") Then
                Update_OrderStatus = " update mdbldort " & _
                                     " set " & DBW("procstat", "R", 2) & _
                                     " WHERE " & DBW("patno", PtId, 2) & _
                                     " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                                     " AND " & DBW("ordseqno", OrdNo, 2)
            Else
                Update_OrderStatus = " update mdbldort " & _
                                     " set " & DBW("procstat", "E", 2) & _
                                     " WHERE " & DBW("patno", PtId, 2) & _
                                     " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                                     " AND " & DBW("ordseqno", OrdNo, 2)
            End If
        End If
        
        Set Rs = Nothing
    Else
'상태전송 미포함
'        Update_OrderStatus = " update mdbldort " & _
'                             " set " & DBW("donefg", stsINPROCESS, 3) & _
'                                       DBW("stscd", stsINPROCESS, 2) & _
'                             " WHERE" & _
'                                    DBW("patno", PtId, 2) & _
'                                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
'                                    " AND " & DBW("ordseqno", OrdNo, 2)
'상태전송 포함
        Update_OrderStatus = " update mdbldort " & _
                             " set " & DBW("donefg", stsINPROCESS, 3) & _
                                       DBW("stscd", stsINPROCESS, 3) & _
                                       DBW("procstat", "E", 2) & _
                             " WHERE" & DBW("patno", PtId, 2) & _
                             " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                             " AND " & DBW("ordseqno", OrdNo, 2)
    End If

End Function

Public Function Update_BBS202(ByVal accdt As String, ByVal accseq As Long)
'처방 접수내역의 stscd를 진행중으로 바꾸자
    Update_BBS202 = " update " & T_BBS202 & " set " & _
                                              DBW("stscd", stsINPROCESS, 2) & _
                    " WHERE " & _
                            "     " & DBW("workarea", C_WORKAREA, 2) & _
                            " AND " & DBW("accdt", accdt, 2) & _
                            " AND " & DBW("accseq", accseq, 2)
                   
End Function

Public Function Insert_NotestBBS302(ByVal accdt As String, ByVal accseq As String, _
                              ByVal Bldsrc As String, ByVal Bldyy As String, _
                              ByVal Bldno As String, ByVal Compcd As String, _
                              ByVal norstdt As String, ByVal norsttm As String, _
                              ByVal norstid As String, ByVal rmk As String) As String
'결과없이 Assign 될수 있게 Xm결과테이블 작성
    Insert_NotestBBS302 = " insert into " & T_BBS302 & " (" & _
                    " workarea,accdt,accseq,rstseq,bldsrc,bldyy,bldno,compocd,xmethod, " & _
                    " du,step1,step2,step3,step4,rstv,vfydt,vfytm,vfyid,stat, " & _
                    " statdt,stattm,statid,cancelid,canceldt,cancelfg,norstfg,norstdt,norsttm,norstid,rmk )" & _
                    " values (" & _
                      DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("rstseq", "1", 1) & _
                      DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", Compcd, 1) & DBV("xmethod", "", 1) & _
                      DBV("du", "", 1) & DBV("step1", "", 1) & DBV("step2", "", 1) & DBV("step3", "", 1) & DBV("step4", "", 1) & DBV("rstv", "", 1) & _
                      DBV("vfydt", "", 1) & DBV("vfytm", "", 1) & DBV("vfyid", "", 1) & _
                      DBV("stat", "", 1) & DBV("statdt", "", 1) & DBV("stattm", "", 1) & DBV("statid", "", 1) & _
                      DBV("cancelid", "", 1) & DBV("canceldt", "", 1) & DBV("cancelfg", BBSCancelStatus.stsNotCancel, 1) & _
                      DBV("norstfg", "1", 1) & DBV("norstdt", norstdt, 1) & DBV("norsttm", norsttm, 1) & DBV("norstid", norstid, 1) & DBV("rmk", rmk) & ")"
                      
                        
End Function

Public Function Insert_BBS302(ByVal accdt As String, ByVal accseq As Long, ByVal RstSeq As Long, _
                              ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As Long, ByVal Compcd As String, _
                              ByVal xmethod As String, ByVal Du As String, ByVal step1 As String, ByVal step2 As String, ByVal step3 As String, ByVal step4 As String, _
                              ByVal rstval As String, ByVal spcyy As String, ByVal spcno As Long, ByVal Vftdt As String, ByVal Vfttm As String, _
                              ByVal VfyId As String, ByVal stat As String, ByVal statdt As String, ByVal stattm As String, ByVal statid As String, _
                              ByVal rmk As String) As String

    Dim strStatFg   As String
    Dim strFilter   As String
    
    strStatFg = medGetP(stat, 1, COL_DIV)
    strFilter = medGetP(stat, 2, COL_DIV)
    
'Cross-Matching 결과내역에 Insert 한다(Sql문장).
'cancelfg (취소여부:여=1,부=0),생성시는(default=0)
    Insert_BBS302 = " insert into " & T_BBS302 & " (" & _
                    " workarea,accdt,accseq,rstseq,bldsrc,bldyy,bldno,compocd,xmethod, " & _
                    " du,step1,step2,step3,step4,rstv,spcyy,spcno,vfydt,vfytm,vfyid,stat,filterfg, " & _
                    " statdt,stattm,statid,cancelid,canceldt,cancelfg,norstfg,norstdt,norsttm,norstid,rmk )" & _
                    " values (" & _
                      DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("rstseq", RstSeq, 1) & _
                      DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", Compcd, 1) & DBV("xmethod", xmethod, 1) & _
                      DBV("du", Du, 1) & DBV("step1", step1, 1) & DBV("step2", step2, 1) & DBV("step3", step3, 1) & DBV("step4", step4, 1) & DBV("rstv", rstval, 1) & _
                      DBV("spcyy", spcyy, 1) & DBV("spcno", spcno, 1) & DBV("vfydt", Vftdt, 1) & DBV("vfytm", Vfttm, 1) & DBV("vfyid", VfyId, 1) & _
                      DBV("stat", strStatFg, 1) & DBV("filterfg", strFilter, 1) & DBV("statdt", statdt, 1) & DBV("stattm", stattm, 1) & DBV("statid", statid, 1) & _
                      DBV("cancelid", "", 1) & DBV("canceldt", "", 1) & DBV("cancelfg", BBSCancelStatus.stsNotCancel, 1) & DBV("norstfg", "", 1) & DBV("norstdt", "", 1) & _
                      DBV("norsttm", "", 1) & DBV("norstid", "", 1) & DBV("rmk", rmk) & ")"
                      
                        
End Function
Public Function Insert_BBS302NotOk(ByVal accdt As String, ByVal accseq As Long, ByVal RstSeq As Long, _
                              ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As Long, ByVal Compcd As String, _
                              ByVal xmethod As String, ByVal Du As String, ByVal step1 As String, ByVal step2 As String, ByVal step3 As String, ByVal step4 As String, _
                              ByVal rstval As String, ByVal spcyy As String, ByVal spcno As Long, ByVal Vftdt As String, ByVal Vfttm As String, _
                              ByVal VfyId As String, ByVal stat As String, ByVal statdt As String, ByVal stattm As String, ByVal statid As String, _
                              ByVal rmk As String, ByVal canceldt As String, ByVal canceltm As String, ByVal cancelid As String) As String
    Dim strStatFg   As String
    Dim strFilter   As String
    
    strStatFg = medGetP(stat, 1, COL_DIV)
    strFilter = medGetP(stat, 2, COL_DIV)
    
'Cross-Matching 결과내역에 Insert 한다(Sql문장).
    Insert_BBS302NotOk = " insert into " & T_BBS302 & " (" & _
                    " workarea,accdt,accseq,rstseq,bldsrc,bldyy,bldno,compocd,xmethod, " & _
                    " du,step1,step2,step3,step4,rstv,spcyy,spcno,vfydt,vfytm,vfyid,stat,filterfg, " & _
                    " statdt,stattm,statid,cancelid,canceldt,cancelfg,norstfg,norstdt,norsttm,norstid,rmk, " & _
                    " canceltm)" & _
                    " values (" & _
                      DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("rstseq", RstSeq, 1) & _
                      DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", Compcd, 1) & DBV("xmethod", xmethod, 1) & _
                      DBV("du", Du, 1) & DBV("step1", step1, 1) & DBV("step2", step2, 1) & DBV("step3", step3, 1) & DBV("step4", step4, 1) & DBV("rstv", rstval, 1) & _
                      DBV("spcyy", spcyy, 1) & DBV("spcno", spcno, 1) & DBV("vfydt", Vftdt, 1) & DBV("vfytm", Vfttm, 1) & DBV("vfyid", VfyId, 1) & _
                      DBV("stat", strStatFg, 1) & DBV("filterfg", strFilter, 1) & DBV("statdt", statdt, 1) & DBV("stattm", stattm, 1) & DBV("statid", statid, 1) & _
                      DBV("cancelid", cancelid, 1) & DBV("canceldt", canceldt, 1) & DBV("cancelfg", BBSCancelStatus.stsCancel, 1) & DBV("norstfg", "", 1) & DBV("norstdt", "", 1) & _
                      DBV("norsttm", "", 1) & DBV("norstid", "", 1) & DBV("rmk", rmk, 1) & DBV("canceltm", canceltm) & _
                      ")"
                      
                        
End Function

Public Function Update_Stat_BBS302(ByVal accdt As String, ByVal accseq As Long, ByVal RstSeq As Long, _
                                   ByVal step1 As String, ByVal step2 As String, ByVal step3 As String, ByVal step4 As String, _
                                   ByVal VfyDt As String, ByVal rstval As String, ByVal VfyTm As String, ByVal VfyId As String, _
                                   ByVal rmk As String, ByVal Du As String) As String
'응급검체일경우 재 검사결과를 등록한다.

    Update_Stat_BBS302 = " update " & T_BBS302 & " set" & _
                                                   DBW("step1", step1, 3) & _
                                                   DBW("step2", step2, 3) & _
                                                   DBW("step3", step3, 3) & _
                                                   DBW("step4", step4, 3) & _
                                                   DBW("rstv", rstval, 3) & _
                                                   DBW("vfydt", VfyDt, 3) & _
                                                   DBW("vfytm", VfyTm, 3) & _
                                                   DBW("vfyid", VfyId, 3) & _
                                                   DBW("du", Du, 3) & _
                                                   DBW("rmk", rmk, 2) & _
                         " WHERE     " & DBW("workarea", C_WORKAREA, 2) & _
                         "       AND " & DBW("accdt", accdt, 2) & _
                         "       AND " & DBW("accseq", accseq, 2) & _
                         "       AND " & DBW("rstseq", RstSeq, 2)
                              
End Function

Public Function SetUpdateBBS302(ByVal accdt As String, ByVal accseq As Long, ByVal RstSeq As Long, _
                                   ByVal step1 As String, ByVal step2 As String, ByVal step3 As String, ByVal step4 As String, _
                                   ByVal VfyDt As String, ByVal rstval As String, ByVal VfyTm As String, ByVal VfyId As String, _
                                   ByVal rmk As String, ByVal stat As String, ByVal statdt As String, ByVal stattm As String, ByVal statid As String) As String
'이미 찜해놓은경우 모든 정보를 UPdate 해준다.

    Dim strStatFg   As String
    Dim strFilter   As String
    
    strStatFg = medGetP(stat, 1, COL_DIV)
    strFilter = medGetP(stat, 2, COL_DIV)

    SetUpdateBBS302 = " update " & T_BBS302 & " set" & _
                                        DBW("step1", step1, 3) & _
                                        DBW("step2", step2, 3) & _
                                        DBW("step3", step3, 3) & _
                                        DBW("step4", step4, 3) & _
                                        DBW("rstv", rstval, 3) & _
                                        DBW("vfydt", VfyDt, 3) & _
                                        DBW("vfytm", VfyTm, 3) & _
                                        DBW("vfyid", VfyId, 3) & _
                                        DBW("rmk", rmk, 3) & _
                                        DBW("stat", strStatFg, 3) & _
                                        DBW("filterfg", strFilter, 3) & _
                                        DBW("statdt", statdt, 3) & _
                                        DBW("stattm", stattm, 3) & _
                                        DBW("statid", statid, 2) & _
                         " WHERE     " & DBW("workarea", C_WORKAREA, 2) & _
                         "       AND " & DBW("accdt", accdt, 2) & _
                         "       AND " & DBW("accseq", accseq, 2) & _
                         "       AND " & DBW("rstseq", RstSeq, 2)
                              
End Function

Public Function Update_BBS401(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As Long, ByVal Compcd As String, ByVal stscd As String) As String
'혈액입고내역(BBS401)에 stscd를 Assign 상태로 바꾸어준다.
'조건 bldsrc,bldyy,bldno,compcd
    
    Update_BBS401 = " update " & T_BBS401 & " set " & DBW("stscd", stscd, 2) & _
                    " WHERE " & _
                                      DBW("bldsrc", Bldsrc, 2) & _
                            " AND " & DBW("bldyy", Bldyy, 2) & _
                            " AND " & DBW("bldno", Bldno, 2) & _
                            " AND " & DBW("compocd", Compcd, 2)
End Function


Public Function SetBBS401_IRRADD(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As Long, ByVal Compcd As String, _
                              ByVal stscd As String, ByVal irrfg As String, Irrdt As String, ByVal Irrid As String) As String
'혈액입고내역(BBS401)에 stscd를 Assign 상태로 바꾸어준다.
'조건 bldsrc,bldyy,bldno,compcd
    Dim strRTtm As String
    
    If irrfg = "1" Then
        strRTtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    End If
    SetBBS401_IRRADD = " update " & T_BBS401 & " set " & _
                                  DBW("irrfg=", irrfg, 1) & _
                                  DBW("irrdt=", Irrdt, 1) & _
                                  DBW("irrid=", Irrid, 1) & _
                                  DBW("irrtm=", strRTtm) & _
                       " WHERE " & _
                                         DBW("bldsrc", Bldsrc, 2) & _
                               " AND " & DBW("bldyy", Bldyy, 2) & _
                               " AND " & DBW("bldno", Bldno, 2) & _
                               " AND " & DBW("compocd", Compcd, 2)

End Function
Public Function DelAssignReadyBlood(ByVal WorkArea As String, ByVal accdt As String, ByVal accseq As String, ByVal RstSeq As String) As String
    DelAssignReadyBlood = "delete " & T_BBS302 & _
                          " WHERE " & _
                                    DBW("workarea=", WorkArea) & _
                          " AND " & DBW("accdt=", accdt) & _
                          " AND " & DBW("accseq=", accseq) & _
                          " AND " & DBW("rstseq=", RstSeq)
                            
End Function
Public Function GetAssignReadyBlood(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT workarea,accdt,accseq,rstseq " & _
           " FROM " & T_BBS302 & _
           " WHERE " & _
                     DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy=", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd) & _
           " AND stat is null" & _
           " AND rstv is null"
           
    Set GetAssignReadyBlood = New Recordset
    GetAssignReadyBlood.Open sSql, DBConn

End Function

Public Function update_BBS302(ByVal accdt As String, ByVal accseq As String, _
                              ByVal RstSeq As Long, ByVal stscd As String, _
                              ByVal cancelid As String) As String
'Assign 취소(frmBBS103)화면에서 사용
    Dim canceldt As String
    Dim canceltm As String  '2001-11-12 추가

    canceltm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    canceldt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    update_BBS302 = " update " & T_BBS302 & " set " & _
                                            DBW("cancelfg", stscd, 3) & _
                                            DBW("canceldt", canceldt, 3) & _
                                            DBW("canceltm", canceltm, 3) & _
                                            DBW("cancelid", cancelid, 2) & _
                    " WHERE " & _
                            "     " & DBW("workarea", C_WORKAREA, 2) & _
                            " AND " & DBW("accdt", accdt, 2) & _
                            " AND " & DBW("accseq", accseq, 2) & _
                            " AND " & DBW("rstseq", RstSeq, 2)
End Function

Public Function Get_ReTestList(ByVal accdt As String, ByVal accseq As String, ByVal BNum As String, ByVal CompoCd As String) As Recordset
'---------------------------------------------------
'혈액번호 입력시에
'응급검사이거나,Pheresis검사일경우 재등록 하기위해서
'---------------------------------------------------
    Dim sSql As String
    Dim Bldsrc As String
    Dim Bldyy As String
    Dim Bldno As String
    Dim Sql As String
    
    Bldsrc = Mid(BNum, 1, 2)
    Bldyy = Mid(BNum, 3, 2)
    Bldno = Mid(BNum, 5)
    
    sSql = " SELECT a.workarea,a.accdt,a.accseq,a.rstseq,a.bldsrc,a.bldyy,a.bldno,a.compocd,a.du," & _
          " a.step1,a.step2,a.step3,a.step4,a.rstv,a.stat,a.spcyy,a.spcno,b.componm,b.abbrnm,c.abo,c.rh,c.volumn,a.rmk" & _
          " FROM " & _
                   T_BBS302 & " a," & T_BBS006 & " b," & T_BBS401 & " c" & _
          " WHERE" & _
                   "     " & DBW("a.workarea", C_WORKAREA, 2) & _
                   " AND " & DBW("a.accdt", accdt, 2) & _
                   " AND " & DBW("a.accseq", accseq, 2) & _
                   " AND " & DBW("a.bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("a.bldyy", Bldyy, 2) & _
                   " AND " & DBW("a.bldno", Bldno, 2) & _
                   " AND " & DBW("a.compocd", CompoCd, 2) & _
                   " AND (a.rstv='' or a.rstv is null)" & _
                   " AND a.stat='1'" & _
                   " AND (a.norstfg<>'1' or a.norstfg is null)" & _
                   " AND " & DBW("c.stscd", BBSBloodStatus.stsASSIGN, 2) & _
                   " AND a.compocd=b.compocd" & _
                   " AND a.bldsrc=c.bldsrc" & _
                   " AND a.bldyy=c.bldyy" & _
                   " AND a.bldno=c.bldno" & _
                   " AND a.compocd=c.compocd"


    Set Get_ReTestList = New Recordset
    Get_ReTestList.Open sSql, DBConn

End Function

Public Function Get_Blood_Status(ByVal accdt As String, ByVal accseq As String, ByVal RstSeq As String) As Long
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = "SELECT * " & _
           "FROM " & T_BBS402 & " " & _
           "WHERE " & _
                             DBW("workarea=", C_WORKAREA) & " " & _
                    "AND " & DBW("accdt=", accdt) & " " & _
                    "AND " & DBW("accseq=", accseq) & " " & _
                    "AND " & DBW("rstseq=", RstSeq) & " "
    
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
        Get_Blood_Status = BBSBloodStatus.stsASSIGN
    Else
        With Rs
            If .RecordCount < 1 Then
                Get_Blood_Status = BBSBloodStatus.stsASSIGN
            Else
                If .Fields("retfg").Value & "" = "1" Then
                    Get_Blood_Status = BBSBloodStatus.stsRETURN
                ElseIf .Fields("expfg").Value & "" = "1" Then
                    Get_Blood_Status = BBSBloodStatus.stsEXPIRE
                ElseIf .Fields("bagfg").Value & "" = "1" Then
                    Get_Blood_Status = BBSBloodStatus.stsBAG
                Else
                    Get_Blood_Status = BBSBloodStatus.stsDELIVERY
                End If
            End If
        End With
    End If
    Set Rs = Nothing
End Function

Public Function Get_Collect_AssignList(ByVal accdt As String, ByVal accseq As String) As String
'---------------------------------
'이미 Assign 된 혈액의 정보를 조회
'조건:접수번호
'---------------------------------
    Get_Collect_AssignList = _
            " SELECT a.workarea,a.accdt,a.accseq,a.rstseq,a.bldsrc,a.bldyy,a.bldno," & _
                  "  a.compocd,a.du,a.step1,a.step2,a.step3,a.step4,a.rstv,a.stat,a.spcyy,a.filterfg," & _
                  "  a.spcno,a.vfydt,a.vfytm,a.vfyid,a.statdt,a.stattm,a.statid,a.cancelfg,a.norstfg,d.irrfg," & _
                  "  b.empnm,c.componm as field1,c.abbrnm,d.abo,d.rh,d.volumn,/*d.stscd,*/a.rmk" & _
            " FROM " & T_BBS302 & " a," & T_COM006 & " b," & T_BBS006 & " c," & T_BBS401 & " d" & _
            " WHERE" & _
                     "     " & DBW("a.workarea", C_WORKAREA, 2) & _
                     " AND " & DBW("a.accdt", accdt, 2) & _
                     " AND " & DBW("a.accseq", accseq, 2) & _
                     " AND " & DBJ("a.vfyid*=b.empid") & _
                     " AND a.compocd=c.compocd" & _
                     " AND a.bldsrc=d.bldsrc" & _
                     " AND a.bldyy=d.bldyy" & _
                     " AND a.bldno=d.bldno " & _
                     " AND a.compocd=d.compocd" & _
            " ORDER BY a.workarea,a.accdt,a.accseq,a.rstseq"
End Function
Public Function Get_Delivery(ByVal accdt As String, ByVal accseq As String, ByVal RstSeq As Long) As String
'----------------------------------
'출고혈액에 대한 반환,폐기,회수여부
'----------------------------------
    Get_Delivery = " SELECT deliverydt,retfg,expfg,retdt " & _
                    " FROM " & T_BBS402 & _
                    " WHERE" & _
                             "     " & DBW("workarea", C_WORKAREA, 2) & _
                             " AND " & DBW("accdt", accdt, 2) & _
                             " AND " & DBW("accseq", accseq, 2) & _
                             " AND " & DBW("rstseq", RstSeq, 2)
End Function
Public Function Assign_cnt(ByVal accdt As String, ByVal accseq As Long)
'-------------------------------------------
'해당 처방에대해서 처방 처리 여부를 조회한다
'Assign,Assign 취소,출고,반환,폐기,회수의 갯수
'-------------------------------------------
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = " SELECT * " & _
           " FROM " & T_BBS203 & " " & _
           " WHERE" & _
                    "     " & DBW("workarea", C_WORKAREA, 2) & _
                    " AND " & DBW("accdt", accdt, 2) & _
                    " AND " & DBW("accseq", accseq, 2)
             
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Assign_cnt = False
    Else
        With Rs
            If .RecordCount < 1 Then
                mvarAssignCnt = 0
                mvarCancelCnt = 0
                mvarOutCnt = 0
                mvarRetCnt = 0
                mvarExpCnt = 0
                mvarTotCnt = 0
            Else
                mvarAssignCnt = Val(.Fields("assigncnt").Value & "")
                mvarCancelCnt = Val(.Fields("assigncancelcnt").Value & "")
                mvarOutCnt = Val(.Fields("deliverycnt").Value & "")
                mvarRetCnt = Val(.Fields("retcnt").Value & "")
                mvarExpCnt = Val(.Fields("expcnt").Value & "")
                mvarTotCnt = 0
            End If
        End With
    End If
    Set Rs = Nothing
    
    Assign_cnt = True

End Function

'2001-11-29추가 : 요청수량 관리
Public Function Request_Cnt(ByVal accdt As String, ByVal accseq As Long)
'-------------------------------------------
'해당 처방에대해서 요청가능 여부를 조회한다
'요청수량의 갯수
'-------------------------------------------
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = " SELECT * " & _
           " FROM " & T_BBS204 & _
           " WHERE" & _
                    "     " & DBW("workarea", C_WORKAREA, 2) & _
                    " AND " & DBW("accdt", accdt, 2) & _
                    " AND " & DBW("accseq", accseq, 2)
             
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Request_Cnt = False
    Else
        With Rs
            While Not .EOF
                If .RecordCount < 1 Then
                    mvarRequestCnt = 0
                    mvarTotCnt = 0
                Else
                    mvarRequestCnt = mvarRequestCnt + Val("" & .Fields("reqcnt").Value)
                    mvarTotCnt = 0
                End If
                .MoveNext
            Wend
        End With
    End If
    Set Rs = Nothing
    
    Request_Cnt = True

End Function


Public Function Get_XM_Step() As Recordset
'사용함(확인)
'--------------------------------------------------
'X-Matching 검사Step 조회
'검사단계;검사명(각각의 검사는 ";"로 연결되어있다.)
'--------------------------------------------------
    Dim today As String
    Dim sSql  As String
    
    today = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    sSql = " SELECT a.field1,a.text1 " & _
           " FROM " & T_COM003 & " a" & _
           " WHERE" & _
                  "  " & DBW("cdindex", BC2_XM_STEP, 2) & _
                  " AND cdval1=(SELECT max(b.cdval1)" & _
                              " FROM " & T_COM003 & " b " & _
                              " WHERE" & _
                                       "     " & DBW("b.cdval1<", today, 2) & _
                                       " AND " & DBW("cdindex", BC2_XM_STEP, 2) & ")"
    
    Set Get_XM_Step = New Recordset
    Get_XM_Step.Open sSql, DBConn

End Function

Public Function Get_XM_Blood_List(ByVal accdt As String, ByVal accseq As Long) As Recordset
'------------------------------------------
'xm검사를 하기위해서 처방의 정보를 조회한다
'------------------------------------------
    Dim sSql As String
    
    sSql = " SELECT c.testnm,b.unitqty,b.orddt,a.reqdt,b.ordno,b.ordseq,a.ptid,b.dcfg,a.orddiv," & _
           " c.compocd, c.volumn, b.ordcd,b.statfg,c.xmethod,a.wardid,a.deptcd,b.mesg,b.accdt,b.accseq" & _
           " FROM " & T_BBS001 & " c," & T_LAB101 & " a," & T_LAB102 & " b," & T_LAB102 & " d" & _
           " WHERE " & _
                     DBW("d.workarea=", C_WORKAREA) & _
           " AND " & DBW("d.accdt=", accdt) & _
           " AND " & DBW("d.accseq=", accseq) & _
           " AND " & DBW("a.orddiv=", C_WORKAREA) & _
           " AND a.ptid = d.ptid" & _
           " AND a.orddt=d.orddt" & _
           " AND b.ptid=a.ptid" & _
           " AND b.orddt=a.orddt" & _
           " AND b.ordno=a.ordno" & _
           " AND c.testcd=b.ordcd" & _
           " AND c.applydt=(SELECT max(e.applydt) FROM " & T_BBS001 & " e " & _
                           "WHERE " & _
                           "c.testcd = e.testcd AND " & DBW("e.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & ")"
                           
'    If dbconn.Whatsthis = dbconn.ThisIsMs7 Then
    If ObjSysInfo.dbtype = 2 Then
        sSql = sSql & " ORDER BY a.orddiv "
    Else
        sSql = sSql & " ORDER BY orddiv "
    End If

    Set Get_XM_Blood_List = New Recordset
    Get_XM_Blood_List.Open sSql, DBConn
End Function
Public Function Get_BCNm(ByVal CompoCd As String) As String
'---------------------
'혈액제제명을 구해온다
'---------------------
    Dim Rs   As New Recordset
    Dim sSql As String
    
    sSql = " SELECT b.compocd, b.componm " & _
           " FROM " & _
                    T_BBS001 & " a," & T_BBS006 & " b " & _
           " WHERE" & _
                    " " & DBW("a.testcd", CompoCd, 2) & _
                    " AND a.compocd=b.compocd"
            
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF Then
        Get_BCNm = "" & COL_DIV & ""
    Else
        Get_BCNm = Rs.Fields("compocd").Value & "" & COL_DIV & Rs.Fields("componm").Value & ""
    End If
    
    Set Rs = Nothing
End Function
Public Function Get_PtInfo(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Long) As String
'환자정보를 불러온다.(환자명,성/나이,진료과/병동)
    Dim Rs        As Recordset
    Dim sSql      As String
    Dim strSexAge As String
    Dim strSEX    As String
    Dim strAge    As String
    Dim strDOB    As String
    
    Dim strTmp      As String
    Dim strYY       As String
    Dim strMM       As String
    Dim strDD       As String
    Dim strDeptNm   As String
    Dim strWardNm   As String
    
    
    sSql = " SELECT a." & F_PTNM & " as ptnm," & F_SSN2("a") & " as ssn,b.wardid,b.deptcd " & _
           " FROM " & _
                    T_HIS001 & " a," & T_LAB101 & " b" & _
           " WHERE" & _
                    "     " & DBW("b.ptid", PtId, 2) & _
                    " AND " & DBW("b.orddt", orddt, 2) & _
                    " AND " & DBW("b.ordno", OrdNo, 2) & _
                    " AND b.ptid=a." & F_PTID
    
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Rs.EOF = False Then
        strDeptNm = GetDeptNm(Rs.Fields("deptcd").Value & "")
        If Rs.Fields("wardid").Value & "" <> "" Then
            strWardNm = GetWardNm(Rs.Fields("wardid").Value & "")
        Else
            strWardNm = ""
        End If
        
        strYY = Mid(Rs.Fields("ssn").Value & "", 1, 2)
        strMM = Mid(Rs.Fields("ssn").Value & "", 3, 2)
        strDD = Mid(Rs.Fields("ssn").Value & "", 5, 2)
        
        If Val(strMM) < 1 Then strMM = "01"
        If Val(strMM) > 12 Then strMM = "12"
        If Val(strDD) < 1 Then strDD = "01"
        If Val(strDD) > 31 Then strDD = "31"
       
        If IsDate(strYY & "-" & strMM & "-" & strDD) = False Then
            strDD = "01"
        End If
        
        strTmp = Mid(Rs.Fields("ssn").Value & "", 7, 1)
        Select Case strTmp
            Case "0": strSEX = "여": strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1": strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2": strSEX = "여": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3": strSEX = "남": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4": strSEX = "여": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "5": strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "6": strSEX = "여": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "7": strSEX = "남": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "8": strSEX = "여": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        If Len(Rs.Fields("ssn").Value & "") = 13 Then
            strAge = medFindAge(Replace(strDOB, "-", ""), "Y")
        Else
            strAge = ""
        End If
            
        strSexAge = strAge & "/" & strSEX
        
        Get_PtInfo = Rs.Fields("ptnm").Value & "" & COL_DIV & strSexAge & COL_DIV & _
                    strDeptNm & COL_DIV & strWardNm & COL_DIV & Rs.Fields("ssn").Value & ""
    End If
    Set Rs = Nothing
End Function
Public Function Get_SpcInfo(ByVal PtId As String, ByVal orddt As String) As Recordset
'-------------------------------------------------------------------
'검체정보를 불러온다.(조건:환자id,처방일)
'폐기 않된 검체만 조회
'이 환자의 검체중 최종 검체만....
'-------------------------------------------------------------------
    Dim sSql As String
    
    sSql = " SELECT a.spcyy,a.spcno,a.coldt,a.coltm,a.rcvdt,a.rcvtm,a.buildcd,a.storeleg,a.storerno,a.storecno,a.expfg,a.addfg " & _
           " FROM " & T_BBS201 & " a" & _
           " WHERE" & _
                    "     " & DBW("a.ptid", PtId, 2) & _
                    " AND " & DBW("a.expfg<>", BBSExpFg.stsExpfg) & _
                    " AND a.coldt= ( SELECT max(z.coldt) FROM " & T_BBS201 & " z" & _
                                   " WHERE " & DBW("z.ptid=", PtId) & _
                                   " AND   " & DBW("z.expfg<>", BBSExpFg.stsExpfg) & _
                                   " AND   " & DBW("z.coldt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & ")" & _
                    " AND not storeleg is null " & _
           " ORDER BY rcvdt desc,rcvtm desc "
    
    Set Get_SpcInfo = New Recordset
    Get_SpcInfo.Open sSql, DBConn
    
End Function

Public Function Get_BloodINfo(ByVal BNum As String, ByVal Compcd As String, ByVal EmpID As String, ByVal centercd As String, ByVal PtId As String) As Boolean
'혈액번호를 입력받아서 혈액번호에 해당하는 Component 를 구한다. 검사자도 구한다.
'센터내에 있는 혈액만 가능하다.
'지정헌혈혈액이거나, 자가헌혈혈액은 지정된 환자만 사용가능하다.

    Dim Rs As Recordset
    
    Dim strbldsrc As String
    Dim strbldyy As String
    Dim strbldno As String
    Dim strSQL As String
    
    strbldsrc = Mid(BNum, 1, 2)
    strbldyy = Mid(BNum, 3, 2)
    strbldno = Mid(BNum, 5)
    
    If strbldsrc = "" Or strbldyy = "" Or strbldno = "" Then
        MsgBox "혈액번호가 완전하지 않습니다.", vbCritical, "오류"
        Get_BloodINfo = False
        Exit Function
    End If
    
    strSQL = " SELECT abo,rh,volumn,ptid,reserved,autofg,hosfg,pherefg,irrfg,donorid,donoraccdt,expdt " & _
             " FROM " & T_BBS401 & _
             " WHERE" & _
                      "     " & DBW("bldsrc", strbldsrc, 2) & _
                      " AND " & DBW("bldyy", strbldyy, 2) & _
                      " AND " & DBW("bldno", strbldno, 2) & _
                      " AND " & DBW("compocd", Compcd, 2) & _
                      " AND " & DBW("centercd", centercd, 2) & _
                      " AND " & DBW("stscd<", BBSBloodStatus.stsASSIGN)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF = False Then
        '성분헌혈 CHECK------------------------------------------------------------------------
        If Rs.Fields("pherefg").Value & "" = "1" Then
            If Rs.Fields("ptid").Value & "" <> PtId Then
                MsgBox "이 혈액은 " & Rs.Fields("ptid").Value & "" & "번 환자에게 지정되어 있는 혈액입니다.", vbCritical, "오류"
                Get_BloodINfo = False
                Set Rs = Nothing
                Exit Function
            End If
        End If
                
                
        '지정혈액여부 check--------------------------------------------------------------------
        If Rs.Fields("reserved").Value & "" = "1" Then
            If Rs.Fields("ptid").Value & "" <> PtId Then
                MsgBox "이 혈액은 " & Rs.Fields("ptid").Value & "" & "번 환자에게 지정되어 있는 혈액입니다.", vbCritical, "오류"
                Get_BloodINfo = False
                Set Rs = Nothing
                Exit Function
            End If
        End If
        '자가 여부 check-----------------------------------------------------------------------
        If Rs.Fields("autofg").Value & "" = "1" Then
        
            If Rs.Fields("ptid").Value & "" <> PtId Then
                MsgBox "이 혈액은 " & Rs.Fields("ptid").Value & "" & "번 환자의 자가수혈 혈액입니다.", vbCritical, "오류"
                Get_BloodINfo = False
                Set Rs = Nothing
                Exit Function
            End If
        End If
        
        '병원에서 헌혈받은 혈액이면, 적격여부가 결정 됬는지 검사한다.--------------------------
        If Rs.Fields("hosfg").Value & "" = "1" Then
            If isNotBlood(Rs.Fields("donorid").Value & "", Rs.Fields("donoraccdt").Value & "") = False Then
                MsgBox "부적격 판정된 혈액입니다.", vbCritical, "오류"
                Get_BloodINfo = False
                Set Rs = Nothing
                Exit Function
            End If
        End If
        
        '폐기일자가 임박한 혈액...(expdt와 오늘날짜가 같은 경우)
        If Rs.Fields("expdt").Value & "" <= Format(GetSystemDate, "yyyyMMdd") Then
            MsgBox "혈액폐기 대상입니다. 폐기 예정일자 : " & Format(Rs.Fields("expdt").Value & "", "####/##/##"), vbCritical, "오류"
            Get_BloodINfo = False
            Set Rs = Nothing
            Exit Function
        End If
        
        Dim strEmpNm As String
        
        strEmpNm = GetEmpNm(EmpID)
        '사용할 수 있네------------------------------------------------------------------------
        mvarStrTmp = Rs.Fields("abo").Value & "" & Rs.Fields("rh").Value & "" & vbTab & _
                   Rs.Fields("volumn").Value & "" & vbTab & strEmpNm & vbTab & Rs.Fields("irrfg").Value & ""
        Get_BloodINfo = True
    Else
'        '혈액제재가 다른 경우 강제등록할 수 있는지 여부 점검
'        'Assign할 혈액의 제재를 무시하고 있는지의 여부만 점검
'        Dim Rs1 As Recordset
'        Dim strSQL1 As String
'
'        strSQL1 = " SELECT abo,rh,volumn,ptid,reserved,autofg,hosfg,pherefg,irrfg,donorid,donoraccdt,compocd " & _
'                 " FROM " & T_BBS401 & _
'                 " WHERE" & _
'                          "     " & DBW("bldsrc", strbldsrc, 2) & _
'                          " AND " & DBW("bldyy", strbldyy, 2) & _
'                          " AND " & DBW("bldno", strbldno, 2) & _
'                          " AND " & DBW("centercd", centercd, 2) & _
'                          " AND " & DBW("stscd<", BBSBloodStatus.stsASSIGN)
'        Set Rs1 = New Recordset
'
'        Rs1.Open strSQL1, DBConn
'
'        If Rs1.EOF Then
            MsgBox "혈액번호에 해당되는 혈액(제제)이 없거나, 이미 Assign되었습니다.", vbCritical, "오류"
            Get_BloodINfo = False
'        Else
'            If MsgBox("수혈처방의 제재와 Assign할 혈액의 제재가 서로 다릅니다." & vbNewLine & _
'                      "이 혈액을 Assign하시겠습니까?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
'
'            '----
'                '성분헌혈 CHECK------------------------------------------------------------------------
'                If Rs1.Fields("pherefg").Value & "" = "1" Then
'                    If Rs1.Fields("ptid").Value & "" <> "0" Or IsNull(Rs1.Fields("ptid").Value) = True Then
'                        If Rs1.Fields("ptid").Value & "" <> PtId Then
'                            MsgBox "이 혈액은 " & Rs1.Fields("ptid").Value & "" & "번 환자에게 지정되어 있는 혈액입니다.", vbCritical, "오류"
'                            Get_BloodINfo = False
'                            Set Rs1 = Nothing
'                            Exit Function
'                        End If
'                    End If
'                End If
'
'
'                '지정혈액여부 check--------------------------------------------------------------------
'                If Rs1.Fields("reserved").Value & "" = "1" Then
'                    If Rs1.Fields("ptid").Value & "" <> PtId Then
'                        MsgBox "이 혈액은 " & Rs1.Fields("ptid").Value & "" & "번 환자에게 지정되어 있는 혈액입니다.", vbCritical, "오류"
'                        Get_BloodINfo = False
'                        Set Rs1 = Nothing
'                        Exit Function
'                    End If
'                End If
'                '자가 여부 check-----------------------------------------------------------------------
'                If Rs1.Fields("autofg").Value & "" = "1" Then
'
'                    If Rs1.Fields("ptid").Value <> PtId Then
'                        MsgBox "이 혈액은 " & Rs1.Fields("ptid").Value & "번 환자의 자가수혈 혈액입니다.", vbCritical, "오류"
'                        Get_BloodINfo = False
'                        Set Rs1 = Nothing
'                        Exit Function
'                    End If
'                End If
'
'                '병원에서 헌혈받은 혈액이면, 적격여부가 결정 됬는지 검사한다.--------------------------
'                If Rs1.Fields("hosfg").Value & "" = "1" Then
'                    If isNotBlood(Rs1.Fields("donorid").Value & "", Rs1.Fields("donoraccdt").Value & "") = False Then
'                        MsgBox "부적격 판정된 혈액입니다.", vbCritical, "오류"
'                        Get_BloodINfo = False
'                        Set Rs1 = Nothing
'                        Exit Function
'                    End If
'                End If
'
''                Dim strEmpNm As String
'
'                strEmpNm = GetEmpNm(EmpID)
'                '사용할 수 있네------------------------------------------------------------------------
'                mvarStrTmp = Rs1.Fields("abo").Value & "" & Rs1.Fields("rh").Value & "" & vbTab & _
'                           Rs1.Fields("volumn").Value & "" & vbTab & strEmpNm & vbTab & Rs1.Fields("irrfg").Value & ""
'
'                mvarReturnCompoCd = Rs1.Fields("compocd").Value & ""
'                Get_BloodINfo = True
'                '----
'            Else
'                Get_BloodINfo = False
'            End If
'        End If
        
'        Set Rs1 = Nothing
    End If
    
    Set Rs = Nothing
End Function

Private Function isNotBlood(ByVal Donorid As String, ByVal donoraccdt As String) As Boolean
    Dim sSql As String
    Dim Rs   As Recordset
    
    sSql = " SELECT * FROM " & T_BBS603 & _
           " WHERE " & DBW("donorid=", Donorid) & _
           " AND " & DBW("donoraccdt=", donoraccdt)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    isNotBlood = True
    If Not Rs.EOF Then
        isNotBlood = IIf(Rs.Fields("okdiv3").Value & "" = "0", False, True)
    End If
    Set Rs = Nothing

End Function
Private Function IsValidBlood(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String) As Boolean
    Dim sSql As String
    Dim DrRS As Recordset
    
    Dim Donorid As String
    Dim donordt As String
    
    
    '이 혈액의 헌혈자를 구한다.---------------------------------------------------
    sSql = "SELECT donorid,donoraccdt " & _
           "FROM " & T_BBS602 & " " & _
           "WHERE" & _
                   "     " & DBW("bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("bldyy", Bldyy, 2) & _
                   " AND " & DBW("bldno", Bldno, 2)
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        IsValidBlood = False
    Else
        With DrRS
            If .RecordCount < 1 Then
                MsgBox "적격인지 부적격인지 알 수 없습니다.", vbCritical, "오류"
                IsValidBlood = False
            Else
                Donorid = .Fields("donorid").Value & ""
                donordt = .Fields("donoraccdt").Value & ""
            End If
        End With
        Set DrRS = Nothing
    End If
    
    '판정이 어떻게 되었는지 검사한다.---------------------------------------------
    sSql = "SELECT okdiv3 " & _
           "FROM " & T_BBS603 & " " & _
           "WHERE" & _
                   "     donorid" & DBV("donorid", Donorid, 2) & _
                   " AND donoraccdt" & DBV("donoraccdt", donordt, 2)
                   
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        IsValidBlood = False
    Else
        With DrRS
            If .RecordCount < 1 Then
                MsgBox "적격인지 부적격인지 알 수 없습니다.", vbCritical, "오류"
                IsValidBlood = False
            Else
                Select Case .Fields("okdiv3").Value & ""
                    Case ""
                        MsgBox "아직 판정이 되지 않은 혈액입니다.", vbCritical, "오류"
                        IsValidBlood = False
                    Case "0"
                        MsgBox "부적격 판정된 혈액입니다.", vbCritical, "오류"
                        IsValidBlood = False
                    Case "1"
                        IsValidBlood = True
                End Select
            End If
        End With
        Set DrRS = Nothing
    End If
    
End Function

Public Function Get_RstSeq(ByVal accdt As String, ByVal accseq As Long) As Long

'bbs302에 저장시 결과 Seq를 구해온다.(rstseq)
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT max(rstseq) as maxrstseq " & _
           " FROM " & T_BBS302 & _
           " WHERE " & DBW("workarea", C_WORKAREA, 2) & _
           " AND   " & DBW("accdt", accdt, 2) & _
           " AND   " & DBW("accseq", accseq, 2)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF = False Then
        If IsNull(Rs.Fields("maxrstseq").Value) = True Then
            Get_RstSeq = 1
        Else
            Get_RstSeq = Val(Rs.Fields("maxrstseq").Value & "") + 1
        End If
    Else
        Get_RstSeq = 1
    End If
    
    Set Rs = Nothing
    
End Function

Public Function Get_AssignList(ByVal FromReqDt As String, ByVal ToReqdt As String, _
                               Optional ByVal PtId As String = "", Optional ByVal wardid As String = "") As String
'Assign 을 취소 하고자 할때 조회한다.(frmBBS103)
    
    '-- 수정 By M.G.Choi 2006.12.26 쿼리문 2단계로 분리 함 ... 테스트
    Get_AssignList = " SELECT /*+ rule +*/ c.workarea, c.accdt,c.accseq,c.bldsrc,c.bldyy,c.bldno,c.compocd,c.rstseq," & _
                             "f.componm as field1, " & _
                             "c.vfyid,c.vfydt,c.stat,c.statid," & _
                             "c.statdt,d.abo,d.rh,d.centercd" & _
                     " FROM " & T_BBS006 & " f," & T_BBS401 & " d," & T_BBS302 & " c" & _
                     " WHERE (c.cancelfg='0'  or c.cancelfg='' or c.cancelfg is null) AND (c.rstv='1' or c.stat='1')"
    
    'AND " & DBW("a.donefg", stsINPROCESS, 2)
    Get_AssignList = Get_AssignList & "  " & _
                                      " AND " & DBW("d.stscd", BBSBloodStatus.stsASSIGN, 2) & _
                                      " " & _
                                      " AND c.bldsrc=d.bldsrc" & _
                                      " AND c.bldyy=d.bldyy" & _
                                      " AND c.bldno=d.bldno" & _
                                      " AND c.compocd=d.compocd" & _
                                      " AND c.compocd=f.compocd"
    
    Get_AssignList = Get_AssignList & " ORDER BY c.vfydt,c.bldsrc,c.bldyy,c.bldno "
    
    '** 원본 ----------------------------------------------------------------------------------------------------------------
'    Get_AssignList = " SELECT c.accdt,c.accseq,c.bldsrc,c.bldyy,c.bldno,c.compocd,c.rstseq," & _
'                             "f.componm as field1,a.ptid, " & _
'                             "g.testnm,b.unitqty,a.reqdt,c.vfyid,c.vfydt,c.stat,c.statid," & _
'                             "c.statdt,b.orddt,b.ordno,b.dcfg," & _
'                             "b.ordseq,d.abo,d.rh,d.centercd" & _
'                     " FROM " & T_BBS006 & " f," & T_BBS001 & " g," & _
'                                T_BBS401 & " d," & T_LAB102 & " b," & T_LAB101 & " a," & _
'                                T_BBS302 & " c" & _
'                     " WHERE (c.cancelfg='0'  or c.cancelfg='' or c.cancelfg is null) AND (c.rstv='1' or c.stat='1')"
'
'
'    'AND " & DBW("a.donefg", stsINPROCESS, 2)
'    Get_AssignList = Get_AssignList & "  " & _
'                                      " AND " & DBW("d.stscd", BBSBloodStatus.stsASSIGN, 2) & _
'                                      " " & _
'                                      " AND a.ptid=b.ptid" & _
'                                      " AND a.orddt=b.orddt" & _
'                                      " AND a.ordno=b.ordno" & _
'                                      " AND b.workarea=c.workarea" & _
'                                      " AND b.accdt=c.accdt" & _
'                                      " AND b.accseq=c.accseq" & _
'                                      " AND b.ordcd=g.testcd" & _
'                                      " AND c.bldsrc=d.bldsrc" & _
'                                      " AND c.bldyy=d.bldyy" & _
'                                      " AND c.bldno=d.bldno" & _
'                                      " AND c.compocd=d.compocd" & _
'                                      " AND c.compocd=f.compocd" & _
'                                      " AND g.applydt=(SELECT max(i.applydt)" & _
'                                      " FROM " & T_BBS001 & " i" & _
'                                      " WHERE g.testcd = i.testcd" & _
'                                      " AND " & DBW("i.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & " )"
'
'    Get_AssignList = Get_AssignList & " ORDER BY c.vfydt,c.bldsrc,c.bldyy,c.bldno"
    '----------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Function Get_AssignListSub(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    
    '-- 수정 By M.G.Choi 2006.12.26 쿼리문 2단계로 분리 함 ... 테스트
    Get_AssignListSub = " SELECT a.ptid, g.testnm,b.unitqty,a.reqdt," & _
                             "b.orddt,b.ordno,b.dcfg,b.ordseq" & _
                     " FROM " & T_BBS001 & " g," & _
                                T_LAB102 & " b," & T_LAB101 & " a " & _
                     " WHERE b.workarea = " & DBS(pWorkarea) & " and b.accdt = " & DBS(pAccDt) & " and b.accseq = " & DBN(pAccSeq)
    
    Get_AssignListSub = Get_AssignListSub & "  " & _
                                      " AND a.ptid=b.ptid" & _
                                      " AND a.orddt=b.orddt" & _
                                      " AND a.ordno=b.ordno" & _
                                      " AND b.ordcd=g.testcd" & _
                                      " AND g.applydt=(SELECT max(i.applydt)" & _
                                      " FROM " & T_BBS001 & " i" & _
                                      " WHERE g.testcd = i.testcd" & _
                                      " AND " & DBW("i.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & " )"
    '----------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Function Get_OrdPt(ByVal qFromDt As String, ByVal qToDt As String, Optional ByVal qPtid As String = "") As String
'처방환자조회,orddiv='b'

    Get_OrdPt = " SELECT distinct a.ptid ,b." & F_PTNM & " as ptnm " & _
                " FROM  " & _
                          T_LAB101 & " a," & T_HIS001 & " b" & _
                 " WHERE" & _
                        DBW("a.reqdt>=", qFromDt) & _
                 " AND " & DBW("a.reqdt<=", qToDt) & _
                 " AND " & DBW("a.orddiv=", C_WORKAREA) & _
                 " AND      a.ptid=b." & F_PTID
    If qPtid <> "" Then
        Get_OrdPt = Get_OrdPt & " AND " & DBW("a.ptid=", qPtid)
    End If

End Function
Public Function Get_Leg(ByVal center As String) As String
'보관가능한 Leg을 구해온다.
    
    Get_Leg = " SELECT distinct legcd " & _
              " FROM " & T_BBS206 & _
              " WHERE" & _
                       "     " & DBW("stscd", BBSSaveStatue.stsUsed, 2) & _
                       " AND " & DBW("centercd", center, 2) & _
              " ORDER BY legcd"

End Function

Public Function Get_Row(ByVal Leg As String, ByVal center As String)
'해당leg의 Row를 구한다.
    Get_Row = " SELECT distinct rowno " & _
              " FROM " & T_BBS206 & _
              " WHERE" & _
                       "     " & DBW("stscd", BBSSaveStatue.stsUsed, 2) & _
                       " AND " & DBW("centercd", center, 2) & _
                       " AND " & DBW("legcd", Leg, 2) & _
              " ORDER BY rowno"
        
End Function
Public Function Get_Col(ByVal Leg As String, ByVal center As String)
'해당leg의 Col를 구한다.
    Get_Col = " SELECT distinct colno " & _
              " FROM " & T_BBS206 & _
              " WHERE" & _
                       "     " & DBW("stscd", BBSSaveStatue.stsUsed, 2) & _
                       " AND " & DBW("centercd", center, 2) & _
                       " AND " & DBW("legcd", Leg, 2) & _
              " ORDER BY colno"
End Function
Public Sub SetPtidRmk(ByVal PtId As String, ByVal rmk As String)
    Dim sSql As String
    
On Error GoTo Error
    DBConn.BeginTrans
    
    sSql = " delete " & T_BBS103 & " WHERE " & DBW("ptid =", PtId)
    
    DBConn.Execute sSql
    
    sSql = " insert into " & T_BBS103 & " (ptid,rmk) values(" & DBV("ptid", PtId, 1) & DBV("rmk", rmk) & ")"
    
    DBConn.Execute sSql
    
    DBConn.CommitTrans
    Exit Sub
Error:
    DBConn.RollbackTrans
End Sub
Public Function GetptidRmk(ByVal PtId As String) As String
    Dim Rs As Recordset
    Dim sSql As String
    
    sSql = " SELECT rmk FROM " & T_BBS103 & " WHERE " & DBW("ptid=", PtId)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetptidRmk = Rs.Fields("rmk").Value & ""
    Else
        GetptidRmk = ""
    End If
    
    Set Rs = Nothing
End Function

Public Function DeleteABO(ByVal PtId As String) As String
    DeleteABO = "delete " & T_BBS902 & " WHERE " & DBW("ptid=", PtId)
End Function
Public Function InsertABO(ByVal PtId As String, ByVal cABO As String, ByVal sABO As String, ByVal RH As String) As String
    InsertABO = " insert into " & T_BBS902 & " (ptid,abo1,abo2,rh)" & _
                " values(" & DBV("ptid", PtId, 1) & DBV("abo1", cABO, 1) & DBV("abo2", sABO, 1) & DBV("rh", RH) & ")"
End Function

Public Function GetTestcd() As Recordset
    Dim sSql As String
    
    sSql = " SELECT cdval1,field2 FROM " & T_COM003 & _
           " WHERE " & _
                  DBW("cdindex=", BC2_REACTION_TEST) & _
           " AND (field5 is null or field5='')"
    Set GetTestcd = New Recordset
    GetTestcd.Open sSql, DBConn
End Function

Public Function TestResultXM(ByVal PtId As String) As String
    
    Dim Rs          As Recordset
    Dim sSql        As String
    Dim strTestCd   As String
    Dim strSpcCd    As String
    Dim strVfyDt    As String
    
    strVfyDt = GetAvailDate 'Format(DateAdd("d", -3, GetSystemDate), "yyyyMMdd")
    Set Rs = GetTestcd
    If Not Rs.EOF Then
        Do Until Rs.EOF
            strTestCd = strTestCd & "'" & Rs.Fields("cdval1").Value & "" & "',"
            strSpcCd = strSpcCd & "'" & Rs.Fields("field2").Value & "" & "',"
            Rs.MoveNext
        Loop
        strTestCd = Mid(strTestCd, 1, Len(strTestCd) - 1)
        strSpcCd = Mid(strSpcCd, 1, Len(strSpcCd) - 1)
        
        TestResultXM = " SELECT a.abbrnm10,b.workarea,b.accdt,b.accseq,b.testcd,b.rstcd,b.rstunit,c.field1 as RstCdNm " & _
               " FROM " & T_LAB031 & " c," & T_LAB001 & " a," & T_LAB302 & " b" & _
               " WHERE " & _
                         DBW("b.ptid=", PtId) & _
               " AND b.testcd in (" & strTestCd & ")" & _
               " AND b.spccd in (" & strSpcCd & ")" & _
               " AND " & DBW("b.vfydt>=", strVfyDt) & _
               " AND a.testcd=b.testcd " & _
               " AND     " & DBJ(DBW("c.cdindex", LC2_ItemResult, 2)) & _
               " AND     " & DBJ("c.cdval1   =* b.testcd") & _
               " AND     " & DBJ("c.cdval2   =* b.rstcd") & _
               " ORDER BY b.workarea desc,b.accdt desc,b.accseq"
    End If
End Function

Private Function GetAvailDate() As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select field1 from " & T_COM003 & _
            " where " & DBW("cdindex=", BC2_REACTION_TEST2) & _
            " and cdval1='0'"
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetAvailDate = Format(DateAdd("d", -3, GetSystemDate), "yyyyMMdd")
    Else
        GetAvailDate = Format(DateAdd("d", Val(Rs.Fields("field1").Value & "") * -1, GetSystemDate), "yyyyMMdd")
    End If
    
    Set Rs = Nothing
End Function

Public Function LastTransInfo(ByVal PtId As String) As String
    Dim FDate As String
    Dim tDate As String
    
    FDate = Format(DateAdd("d", -30, GetSystemDate), "yyyymmdd")
    tDate = Format(GetSystemDate, "yyyymmdd")
    
    
    LastTransInfo = "SELECT a.bldsrc,a.bldyy,a.bldno,a.compocd,d.abo,d.rh,a.deliverydt,a.deliverytm,c.componm " & _
                    "FROM " & T_BBS401 & " d," & T_BBS006 & " c," & T_BBS202 & " b," & T_BBS402 & " a " & _
                    "WHERE " & _
                    DBW("a.deliverydt >=", FDate) & _
                    " AND " & DBW("a.deliverydt <=", tDate) & _
                    " AND " & DBW("b.ptid=", PtId) & _
                    " AND a.workarea=b.workarea" & _
                    " AND a.accdt=b.accdt" & _
                    " AND a.accseq=b.accseq" & _
                    " AND a.compocd=c.compocd" & _
                    " AND a.bldsrc=d.bldsrc AND a.bldyy=d.bldyy AND a.bldno=d.bldno AND a.compocd=d.compocd " & _
                    " ORDER BY deliverydt desc"

End Function
Public Function Get_BloodStsCD(ByVal BNum As String, ByVal Compcd As String, ByVal centercd As String) As String
'혈액번호를 입력받아서 혈액번호에 해당하는 Component 를 구한다. 검사자도 구한다.
'센터내에 있는 혈액만 가능하다.
'지정헌혈혈액이거나, 자가헌혈혈액은 지정된 환자만 사용가능하다.
    Dim Rs As New Recordset
    
    Dim strbldsrc As String
    Dim strbldyy As String
    Dim strbldno As String
    Dim strSQL As String
    
    strbldsrc = Mid(BNum, 1, 2)
    strbldyy = Mid(BNum, 3, 2)
    strbldno = Mid(BNum, 5)
    
    If strbldsrc = "" Or strbldyy = "" Or strbldno = "" Then
        MsgBox "혈액번호가 완전하지 않습니다.", vbCritical, "오류"
        Set Rs = Nothing
        Exit Function
    End If
    
    strSQL = " SELECT stscd " & _
             " FROM " & T_BBS401 & _
             " WHERE" & _
                      "     " & DBW("bldsrc", strbldsrc, 2) & _
                      " AND " & DBW("bldyy", strbldyy, 2) & _
                      " AND " & DBW("bldno", strbldno, 2) & _
                      " AND " & DBW("compocd", Compcd, 2) & _
                      " AND " & DBW("centercd", centercd, 2) & _
                      " AND " & DBW("stscd<", BBSBloodStatus.stsASSIGN)
    Dim strTmp As String
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        Get_BloodStsCD = Trim(Rs.Fields("stscd").Value & "")
    End If
    
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
