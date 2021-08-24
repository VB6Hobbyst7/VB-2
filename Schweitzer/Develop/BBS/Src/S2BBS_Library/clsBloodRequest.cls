VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBloodRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Set_BBS202(ByVal accdt As String, ByVal accseq As Long, _
                           ByVal reqid As String) As String
'-------------------------------------------
'처방 접수내역 Insert 한다
'수혈처방 접수내역에서만 사용한다(frmBBS102)
'-------------------------------------------
    Dim strReqTm As String
    Dim strReqDt As String
    
    strReqTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strReqDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    Set_BBS202 = " update " & T_BBS202 & " set " & _
                 DBW("reqdt = ", strReqDt, 1) & DBW("reqtm = ", strReqTm, 1) & DBW("reqid = ", reqid) & _
                 " WHERE " & _
                              DBW("workarea = ", C_WORKAREA) & _
                    " AND " & DBW("accdt = ", accdt) & _
                    " AND " & DBW("accseq = ", accseq)

End Function

Public Function Insert_BBS204(ByVal accdt As String, ByVal accseq As Long, _
                              ByVal reqid As String, ByVal reqcnt As Long) As String
'-------------------------------------------
'2001-11-29추가
'수혈요청내역을 insert한다.(frmBBS109)
'-------------------------------------------
    Dim strReqTm As String
    Dim strReqDt As String
    
    strReqTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strReqDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    Insert_BBS204 = " insert into " & T_BBS204 & " (workarea,accdt,accseq,reqdt,reqtm,reqid,reqcnt) " & _
                    " values (" & DBV("workarea", C_WORKAREA, 1) & _
                                 DBV("accdt", accdt, 1) & _
                                 DBV("accseq", accseq, 1) & _
                                 DBV("reqdt", strReqDt, 1) & _
                                 DBV("reqtm", strReqTm, 1) & _
                                 DBV("reqid", reqid, 1) & _
                                 DBV("reqcnt", reqcnt) & ")"


End Function
'
'
Public Function Set_UpdateL101(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String) As String
'----------------------
'접수시 처방헤더 update
'----------------------
   
'    Set_UpdateL101 = " update " & T_LAB101 & _
'                     " set    " & DBW("donefg", BBSOrdStatus.stsREQUEST, 2) & _
'                     " WHERE  " & _
'                                        DBW("ptid=  ", PtId) & _
'                              " AND " & DBW("orddt=", orddt) & _
'                              " AND " & DBW("ordno=", OrdNo)
    '예수병원용
    Set_UpdateL101 = " update mdbldort " & _
                     " set    " & DBW("donefg", BBSOrdStatus.stsREQUEST, 2) & _
                     " WHERE  " & _
                                        DBW("patno=  ", PtId) & _
                              " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                              " AND " & DBW("ordseqno=", OrdNo)
End Function

Public Function Set_UpdateL102(ByVal PtId As String, ByVal orddt As String, _
                               ByVal OrdNo As String, ByVal OrdSeq As String, _
                               ByVal accdt As String, ByVal accseq As String) As String
'----------------------
'접수시 처방바디 update
'----------------------
    Dim strRcvdt As String
    Dim strrcvtm As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
'    Set_UpdateL102 = " update " & T_LAB102 & " set " & _
'                                               DBW("rcvdt", strRcvdt, 3) & _
'                                               DBW("rcvtm", strrcvtm, 3) & _
'                                               DBW("donefg", BBSOrdStatus.stsREQUEST, 3) & _
'                                               DBW("stscd", BBSOrdStatus.stsREQUEST, 3) & _
'                                               DBW("workarea", C_WORKAREA, 3) & _
'                                               DBW("accdt", accdt, 3) & _
'                                               DBW("accseq", accseq, 2) & _
'                            " WHERE " & _
'                                               DBW("ptid=", PtId) & _
'                                    " AND  " & DBW("orddt=", orddt) & _
'                                    " AND  " & DBW("ordno=", OrdNo) & _
'                                    " AND  " & DBW("ordseq=", OrdSeq)
    
    '예수병원 용
    Set_UpdateL102 = " update mdbldort set " & _
                                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss')" & _
                                               DBW("donefg", BBSOrdStatus.stsREQUEST, 3) & _
                                               DBW("stscd", BBSOrdStatus.stsREQUEST, 3) & _
                                               DBW("workarea", C_WORKAREA, 3) & _
                                               DBW("accdt", accdt, 3) & _
                                               DBW("accseq", accseq, 2) & _
                            " WHERE " & _
                                               DBW("patno=", PtId) & _
                                    " AND  orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                                    " AND  " & DBW("ordseqno=", OrdNo)
End Function


Public Function Set_UpdateB206(ByVal centercd As String, ByVal legcd As String, _
                               ByVal rowno As Long, ByVal colno As Long, _
                               ByVal spcyy As String, ByVal spcno As String) As String

'------------------------------
'검체보관장소 update (검체정보)
'------------------------------
    
    
    
    Set_UpdateB206 = " update " & T_BBS206 & " set " & _
                                               DBW("spcyy", spcyy, 3) & _
                                               DBW("spcno", spcno, 3) & _
                                               DBW("stscd", BBSSaveStatue.stsNotUsed, 2) & _
                    " WHERE " & _
                                      DBW("centercd=", centercd) & _
                            " AND " & DBW("legcd=", legcd) & _
                            " AND " & DBW("rowno=", rowno) & _
                            " AND " & DBW("colno=", colno)
             
End Function

Public Function Set_UpdateB201(ByVal Fullspcno As String, _
                               ByVal rcvid As String, ByVal legcd As String, ByVal rowno As Long, _
                               ByVal colno As Long) As String
    
'----------------------------------------------
'채혈내역 update
'접수시 접수정보와 검체보관장소 정보를 저장한다
'----------------------------------------------
    Dim strRcvdt As String
    Dim strrcvtm As String
    Dim strSpcYY As String
    Dim strSpcNo As String
    
    strSpcYY = Mid(Fullspcno, 1, 2)
    strSpcNo = Format(Mid(Fullspcno, 3), "#########")
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
    Set_UpdateB201 = " update " & T_BBS201 & " set" & _
                                DBW("rcvdt", strRcvdt, 3) & _
                                DBW("rcvtm", strrcvtm, 3) & _
                                DBW("rcvid   ", rcvid, 3) & _
                                DBW("storeleg", legcd, 3) & _
                                DBW("storerno", rowno, 3) & _
                                DBW("storecno", colno, 2) & _
                     " WHERE " & _
                             "     " & DBW("spcyy=", strSpcYY) & _
                             " AND " & DBW("spcno=", strSpcNo)
                        
End Function
Public Function SavePositionRs(ByVal qCentercd As String, ByVal qSpcyy As String, ByVal qSpcno As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_BBS206 & _
           " WHERE " & _
                     DBW("centercd=", qCentercd) & _
           " AND " & DBW("spcyy=", qSpcyy) & _
           " AND " & DBW("spcno=", qSpcno) & _
           " AND " & DBW("stscd=", 1)
    
    Set SavePositionRs = New Recordset
    SavePositionRs.Open sSql, DBConn
    
End Function


Public Function Set_AccessCancel(ByVal PtId As String, ByVal AccNo As String, ByVal stscd As String, _
                                 ByVal cancelrsn As String, ByVal cancelid As String, ByVal cancelrmk As String, _
                                 ByVal cCnt As Long) As Boolean
'----------------------
'접수취소 Sql 문장
'1: 처방 헤더 update
'2: 처방 바디 update
'3: 처방 접수내역 update
'4: 처방 Status update
'5: 결과내역 update
'6: 혈액입고내역 update
'-----------------------
    Dim accdt       As String
    Dim accseq      As String
    Dim strCancelDt As String
    Dim strCancelTm As String
    Dim sSql        As String
    Dim DrRS        As Recordset
    Dim orddt       As String
    Dim OrdNo       As String
    Dim i           As Long
    
    accdt = medGetP(AccNo, 1, "-")
    accseq = medGetP(AccNo, 2, "-")
    strCancelDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strCancelTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    

    sSql = " SELECT ptid,orddt,ordno " & _
           " FROM " & T_LAB102 & _
           " WHERE " & DBW("ptid=", PtId) & _
           " AND   " & DBW("workarea=", C_WORKAREA) & _
           " AND   " & DBW("accdt=", accdt) & _
           " AND   " & DBW("accseq=", accseq)
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
        Set DrRS = Nothing
        Exit Function
    Else
        orddt = DrRS.Fields("orddt").Value & ""
        OrdNo = DrRS.Fields("ordno").Value & ""
        Set DrRS = Nothing
    End If
    
On Error GoTo Save_CanCel_Error
    DBConn.BeginTrans
    
   '---------------
   '처방헤더 update
   '---------------
'    sSql = " update " & T_LAB101 & _
'           " set    " & DBW("donefg", stscd, 2) & _
'           " WHERE  " & DBW("ptid=", PtId) & _
'           " AND    " & DBW("orddt=", orddt) & _
'           " AND    " & DBW("ordno=", OrdNo)
    '예수병원 용
    sSql = " update mdbldort " & _
           " set    " & DBW("donefg", stscd, 2) & _
           " WHERE  " & DBW("patno=", PtId) & _
           " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
           " AND    " & DBW("ordseqno=", OrdNo)
    DBConn.Execute sSql
    '---------------
    '처방바디 update
    '---------------
'    sSql = " update " & T_LAB102 & _
'           " set " & _
'                DBW("donefg", stscd, 3) & _
'                DBW("stscd", stscd, 3) & _
'                DBW("rcvdt", "", 3) & _
'                DBW("rcvtm", "", 3) & _
'                DBW("workarea", "", 3) & _
'                DBW("accdt", "", 3) & _
'                DBW("accseq", "", 2) & _
'            " WHERE " & _
'                    "     " & DBW("ptid=", PtId) & _
'                    " AND " & DBW("workarea=", C_WORKAREA) & _
'                    " AND " & DBW("accdt=", accdt) & _
'                    " AND " & DBW("accseq=", accseq)
    '예수병원 용
'상태 전송 미포함
'    sSql = " update mdbldort " & _
'           " set " & _
'                DBW("donefg", stscd, 3) & _
'                DBW("stscd", stscd, 3) & _
'                " rcpdate=null," & _
'                DBW("workarea", "", 3) & _
'                DBW("accdt", "", 3) & _
'                DBW("accseq", "", 2) & _
'            " WHERE " & _
'                    "     " & DBW("patno=", PtId) & _
'                    " AND " & DBW("workarea=", C_WORKAREA) & _
'                    " AND " & DBW("accdt=", accdt) & _
'                    " AND " & DBW("accseq=", accseq)
'상태 전송 포함
'OCS에 상태를 B로만 전송(채혈상태로만)
    sSql = " update mdbldort " & _
           " set " & _
                DBW("donefg", stscd, 3) & _
                DBW("stscd", stscd, 3) & _
                " rcpdate=null," & _
                DBW("workarea", "", 3) & _
                DBW("accdt", "", 3) & _
                DBW("accseq", "", 3) & _
                DBW("procstat", "B", 2) & _
            " WHERE " & _
                    "     " & DBW("patno=", PtId) & _
                    " AND " & DBW("workarea=", C_WORKAREA) & _
                    " AND " & DBW("accdt=", accdt) & _
                    " AND " & DBW("accseq=", accseq)
    DBConn.Execute sSql
    '-------------------
    '처방접수내역 update
    '-------------------
    sSql = " update " & T_BBS202 & _
           " set " & _
                DBW("stscd", stscd, 3) & _
                DBW("canceldt", strCancelDt, 3) & _
                DBW("canceltm", strCancelTm, 3) & _
                DBW("cancelid", cancelid, 3) & _
                DBW("cancelfg", BBSCancelStatus.stsCancel, 3) & _
                DBW("cancelrsn", cancelrsn, 3) & _
                DBW("cancelrmk", cancelrmk, 2) & _
           " WHERE" & _
                  "     " & DBW("workarea=", C_WORKAREA) & _
                  " AND " & DBW("accdt=", accdt) & _
                  " AND " & DBW("accseq=", accseq)
    
    DBConn.Execute sSql
    '----------------------
    'xm 결과등록내역 update
    '----------------------
    sSql = " update " & T_BBS302 & " set " & _
                                    DBW("cancelfg", BBSCancelStatus.stsCancel, 3) & _
                                    DBW("canceldt", strCancelDt, 3) & _
                                    DBW("cancelid", cancelid, 2) & _
            " WHERE " & _
                    "     " & DBW("workarea=", C_WORKAREA) & _
                    " AND " & DBW("accdt=        ", accdt) & _
                    " AND " & DBW("accseq=      ", accseq)
    
    DBConn.Execute sSql
    '-----------------------
    '처방 status 내역 update
    '----------------------
    sSql = " update " & T_BBS203 & " set " & DBW("assigncancelcnt", cCnt, 2) & _
            " WHERE " & _
                    "     " & DBW("workarea=", C_WORKAREA) & _
                    " AND " & DBW("accdt=        ", accdt) & _
                    " AND " & DBW("accseq=      ", accseq)
    DBConn.Execute sSql
    '-------------------
    '혈액입고내역 update
    '-------------------
    sSql = " SELECT bldsrc,bldyy,bldno,compocd " & _
           " FROM " & T_BBS302 & _
           " WHERE " & DBW("workarea=", C_WORKAREA) & _
           " AND   " & DBW("accdt=", accdt) & _
           " AND   " & DBW("accseq=", accseq)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then Set DrRS = Nothing: GoTo Save_CanCel_Error
    For i = 1 To DrRS.RecordCount
        sSql = " update " & T_BBS401 & _
               " set " & DBW("stscd", BBSBloodStatus.stsENTER, 2) & _
               " WHERE " & DBW("bldsrc=", DrRS.Fields("bldsrc").Value & "") & _
               " AND   " & DBW("bldyy=", DrRS.Fields("bldyy").Value & "") & _
               " AND   " & DBW("bldno=", DrRS.Fields("bldno").Value & "") & _
               " AND   " & DBW("compocd=", DrRS.Fields("compocd").Value & "")
        DBConn.Execute sSql
        
        DrRS.MoveNext
    Next i
    Set DrRS = Nothing
    '------------------------
    '검체별 처방내역 delete
    '------------------------
    If stscd = BBSOrdStatus.stsORDER Then
        sSql = " delete " & T_BBS102 & _
               " WHERE  " & DBW("ptid=", PtId) & _
               " AND    " & DBW("orddt=", orddt) & _
               " AND    " & DBW("ordno=", OrdNo)
        DBConn.Execute sSql
    End If
    
    
    DBConn.CommitTrans
    Set_AccessCancel = True
    Exit Function
    
Save_CanCel_Error:
    DBConn.RollbackTrans
    Set_AccessCancel = False
    MsgBox Err.Description, vbExclamation
End Function
                                 



Public Function B_Existence(ByVal PtId As String, ByVal coldt As String, ByVal coltm As String) As Boolean
'------------------------------------------------
'검체채취시간을 비교하여 검체 채취여부를 판단한다
'True: 채취한다, False:채취않한다
'------------------------------------------------
    Dim DrRS As New Recordset
    Dim strSql As String
    Dim strExistence As String
    Dim lngStoreHour As Long
    Dim strCompare As String
    Dim strCompare1 As String
    
    strCompare1 = coldt & Mid(coltm, 1, 4)
    strCompare1 = Format(strCompare1, "####-##-## ##:##")
    
    strSql = " SELECT coldt,coltm FROM " & T_BBS201 & " a" & _
             " WHERE a.coldt=(SELECT max(b.coldt) " & _
                            " FROM " & T_BBS201 & " b" & _
                            " WHERE  " & _
                                    "" & DBV("b.ptid", PtId, 2) & _
                                    " AND a.ptid=b.ptid " & _
                                    " AND " & DBW("b.coldt<", coldt, 2) & ")" & _
             " ORDER BY a.coltm desc "
    
    Set DrRS = New Recordset
    DrRS.Open strSql, DBConn
    If Not DrRS.EOF = True Then
        strCompare = DrRS.Fields("coldt").Value & "" & Mid(DrRS.Fields("coltm").Value & "", 1, 4)
        strCompare = Format(strCompare, "####-##-## ##:##")
        lngStoreHour = CLng(DateDiff("h", strCompare, strCompare1))
        If lngStoreHour <= 72 Then
             B_Existence = False
        Else
            B_Existence = True
        End If
    Else
        B_Existence = True
    End If
    Set DrRS = Nothing
End Function
Public Function Set_SpcAdder(ByVal PtId As String, ByVal spcyy As String, _
                             ByVal spcno As String, Reqdt As String) As String
'---------------------------------
'검체추가요청서를 update 해준다
'추가요청이 나서 검체채취를 한경우
'--------------------------------
    Set_SpcAdder = " update " & T_BBS207 & " set " & _
                                            DBW("spcyy", spcyy, 3) & _
                                            DBW("spcno", spcno, 3) & _
                                            DBW("donefg", BBSOrdStatus.stsCOLLECT, 2) & _
                   " WHERE " & _
                           "     " & DBW("ptid", PtId, 2) & _
                           " AND " & DBW("reqdt", Reqdt, 2) & _
                           " AND " & DBW("dongfg", BBSOrdStatus.stsORDER, 2)

End Function

Private Function Get_AddSpcSeq(ByVal PtId As String, ByVal strReqDt As String) As String
'---------------------------------------
'검체추가요청내역 작성시 순번을 체크한다
'---------------------------------------
    Dim sSql As String
    Dim DrRS As Recordset
    
    sSql = "SELECT MAX(seq) as maxseq FROM " & T_BBS207 & _
             " WHERE" & _
                    "     " & DBW("ptid", PtId, 2) & _
                    " AND " & DBW("reqdt", strReqDt, 2)
    
    Set DrRS = New Recordset
    DrRS.Open sSql, DBConn
    
    If IsNull(DrRS.Fields("maxseq").Value) = True Then
        Get_AddSpcSeq = 1
    Else
        Get_AddSpcSeq = DrRS.Fields("maxseq").Value & "" + 1
    End If
    
    Set DrRS = Nothing
End Function

Public Function Get_TargetChk(ByVal PtId As String) As Boolean
'------------------------------------
'검체 추가 요청내역의 유무를 확인한다
'------------------------------------
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = " SELECT * FROM " & T_BBS207 & _
            " WHERE " & _
                    "     " & DBW("ptid", PtId, 2) & _
                    " AND " & DBW("dongfg", BBSOrdStatus.stsORDER, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.RecordCount < 1 Then
        Get_TargetChk = False
    Else
        Get_TargetChk = True
    End If
    
    Set Rs = Nothing
End Function
Public Function Set_InsertSpcAdd(ByVal PtId As String, ByVal Reqdt As String, ByVal ReqTm As String, _
                                 ByVal accdt As String, ByVal AccNo As String, ByVal rsncd As String, _
                                 ByVal busidiv As String, ByVal BedIndt As String, ByVal orddt As String, _
                                 ByVal DeptCd As String, ByVal wardid As String, ByVal DoneFg As String, _
                                 ByVal reqid As String) As Boolean
                                 
'---------------------------------------------
'검체추가 요청내역 작성
'이미 검체 추가 요청이 있으면, 작성하지 않는다
'---------------------------------------------
    If Get_TargetChk(PtId) = True Then
        MsgBox "이미 검체추가 요청 상태입니다.", vbInformation + vbOKOnly, "검체추가요청"
        Set_InsertSpcAdd = True
        Exit Function
    End If
    
    Dim sSql   As String
    Dim addseq As String
    
    addseq = Get_AddSpcSeq(PtId, Format(GetSystemDate, PRESENTDATE_FORMAT))
    
    sSql = " insert into " & T_BBS207 & _
           " (ptid,reqdt,seq,reqtm,reqid,accdt,accno,rsncd,busidiv,bedindt," & _
           " orddt,deptcd,wardid,dongfg,spcyy,spcno )" & _
           " values (" & _
                        DBV("ptid", PtId, 1) & DBV("reqdt", Reqdt, 1) & _
                        DBV("seq", addseq, 1) & DBV("reqtm", ReqTm, 1) & _
                        DBV("reqid", reqid, 1) & DBV("accdt", accdt, 1) & _
                        DBV("accno", AccNo, 1) & DBV("rsncd", rsncd, 1) & _
                        DBV("busidiv", busidiv, 1) & DBV("bedindt", BedIndt, 1) & _
                        DBV("orddt", orddt, 1) & DBV("deptcd", DeptCd, 1) & _
                        DBV("wardid", wardid, 1) & DBV("donefg", DoneFg, 1) & _
                        DBV("spcyy", "", 1) & DBV("spcno", "") & _
                   ")"
             
On Error GoTo DBExecError

    DBConn.BeginTrans
    DBConn.Execute sSql
    DBConn.CommitTrans
    Set_InsertSpcAdd = True
    MsgBox "검체추가요청 내역이 작성되었습니다.", vbInformation + vbOKOnly, "검체추가요청"
    Exit Function

DBExecError:
    DBConn.RollbackTrans
    Set_InsertSpcAdd = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetSQL_CancelBloodResreved(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As String
'-----------------------
'지정혈액취소(frmBBS310)
'-----------------------

    GetSQL_CancelBloodResreved = " update " & T_BBS401 & " " & _
                                 " set " & DBW("reserved=", "0") & _
                                 " WHERE " & _
                                                   DBW("bldsrc=", Bldsrc) & _
                                         " AND " & DBW("bldyy=", Bldyy) & _
                                         " AND " & DBW("bldno=", Bldno) & _
                                         " AND " & DBW("compocd=", CompoCd)

End Function


Public Function SetBldStorageUpdateByCenterCd(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                              ByVal Bldno As Long, ByVal CompoCd As String, _
                                              ByVal centercd As String) As String
'------------------------------------------
'혈액 입고 내역에 CenterCd 를 업데이트 한다
'혈액 Transfer(frmBBS309)
'------------------------------------------
    SetBldStorageUpdateByCenterCd = " update " & T_BBS401 & _
                                    " set    " & DBW("centercd", centercd, 2) & _
                                    " WHERE" & _
                                             "     " & DBW("bldsrc", Bldsrc, 2) & _
                                             " AND " & DBW("bldyy", Bldyy, 2) & _
                                             " AND " & DBW("bldno", Bldno, 2) & _
                                             " AND " & DBW("compocd", CompoCd, 2)


End Function

Public Function SetBldDelivery(ByVal Bldsrc As String, ByVal Bldyy As String, _
                               ByVal Bldno As Long, ByVal CompoCd As String, _
                               ByVal deliverydt As String, ByVal Deliveryseq As Long, _
                               ByVal deliverytm As String, ByVal Deliveryid As String, _
                               ByVal rcvid As String, ByVal WorkArea As String, _
                               ByVal accdt As String, ByVal accseq As String, _
                               ByVal OrdCd As String, _
                               ByVal LocalCd As String, _
                               ByVal remark As String, _
                               Optional ByVal RstSeq As String = "0") As String
'---------------------------------------------
'To BBS402
'혈액 출고 내역(Local 병원 의뢰 혈액(frmBBS311)
'---------------------------------------------
    
    SetBldDelivery = " insert into " & T_BBS402 & _
                     " (bldsrc,bldyy,bldno,compocd,deliverydt," & _
                     " deliveryseq,deliverytm,deliveryid,rcvid,workarea," & _
                     " accdt,accseq,rstseq,ordcd,localcd,rmk) " & _
                     " values(" & _
                                DBV("bldsrc", Bldsrc) & "," & DBV("bldyy", Bldyy) & "," & _
                                DBV("bldno", Bldno) & "," & DBV("compocd", CompoCd) & "," & _
                                DBV("deliverydt", deliverydt) & "," & DBV("deliveryseq", Deliveryseq) & "," & _
                                DBV("deliverytm", deliverytm) & "," & DBV("deliveryid", Deliveryid) & "," & _
                                DBV("rcvid", rcvid) & "," & DBV("workarea", WorkArea) & "," & _
                                DBV("accdt", accdt) & "," & DBV("accseq", accseq) & "," & _
                                DBV("rstseq", RstSeq) & "," & DBV("ordcd", OrdCd) & "," & _
                                DBV("localcd", LocalCd) & "," & _
                                DBV("rmk", remark) & _
                            ")"
            
End Function
Public Function SetBldStorageUpdateByStsCd(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                              ByVal Bldno As Long, ByVal CompoCd As String, _
                                              ByVal stscd As String) As String
'---------------------------------------------
'To BBS401
'혈액 입고 내역에 StsCd 를 업데이트 한다.
'(Local 병원 의뢰 혈액(frmBBS311)
'---------------------------------------------
    SetBldStorageUpdateByStsCd = "update " & T_BBS401 & _
                                 " set   " & DBW(stscd, 2) & _
                                 " WHERE " & _
                                         "     " & DBW("bldsrc", Bldsrc, 2) & _
                                         " AND " & DBW("bldyy", Bldyy, 2) & _
                                         " AND " & DBW("bldno", Bldno, 2) & _
                                         " AND " & DBW("compocd", CompoCd, 2)


End Function
Public Function SetDonorMST(ByVal UpdateFg As Boolean, _
                            ByVal Donorid As String, ByVal donornm As String, _
                            ByVal ssn As String, ByVal dob As String, _
                            ByVal sex As String, ByVal ZipCd As String, _
                            ByVal Addr1 As String, ByVal Addr2 As String, _
                            ByVal TelNo As String, ByVal JobCd As String, _
                            ByVal ABO As String, ByVal RH As String, _
                            ByVal cnt As Long, ByVal totvol As Long) As String
'-----------------------------
'To BBS601
'헌혈자 마스터(frmBBS401)
'-----------------------------
    If UpdateFg Then
        SetDonorMST = "update " & T_BBS601 & _
                      " set  " & _
                                DBW("donornm", donornm, 3) & DBW("ssn", ssn, 3) & _
                                DBW("dob", dob, 3) & DBW("sex", sex, 3) & _
                                DBW("zipcd", ZipCd, 3) & DBW("addr1", Addr1, 3) & _
                                DBW("addr2", Addr2, 3) & DBW("telno", TelNo, 3) & _
                                DBW("jobcd", JobCd, 3) & DBW("abo", ABO, 3) & _
                                DBW("rh", RH, 3) & DBW("cnt", cnt, 3) & DBW("totvol", totvol, 2) & _
                      " WHERE " & DBW("donorid", Donorid, 2)
    Else
        SetDonorMST = "insert into " & T_BBS601 & "(" & _
                      "donorid,donornm,ssn,dob,sex,zipcd,addr1,addr2,telno,jobcd,abo,rh,cnt,totvol )" & _
                      "values ( " & _
                      DBV("donorid", Donorid, 1) & DBV("donornm", donornm, 1) & DBV("ssn", ssn, 1) & _
                      DBV("dob", dob, 1) & DBV("sex", sex, 1) & DBV("zipcd", ZipCd, 1) & DBV("Addr1", Addr1, 1) & DBV("addr2", Addr2, 1) & DBV("telno", TelNo, 1) & _
                      DBV("jobcd", JobCd, 1) & DBV("abo", ABO, 1) & DBV("rh", RH, 1) & DBV("cnt", cnt, 1) & DBV("totvol", totvol) & ")"
    End If
End Function

Public Function SetNoGiveInfo(ByVal UpdateFg As Boolean, ByVal CDINDEX As String, _
                              ByVal seq As Long) As String
'--------------------------------------------------------------------------------------
'To COM099
'NoIndex : B006 (BN_DONOR_ID, 헌혈자 ID), B001(BN_TMP_ID, 임상병리 검사용 임시 환자 ID)
'번호부여 정보
'--------------------------------------------------------------------------------------

    If UpdateFg Then
        SetNoGiveInfo = " update " & T_COM099 & _
                        " set " & DBW("seq=", seq) & _
                        " WHERE" & _
                                 "     " & DBW("noindex", CDINDEX, 2) & _
                                 " AND " & DBW("divcd1", C_WORKAREA, 2) & _
                                 " AND " & DBW("divcd2", "0", 2) & _
                                 " AND " & DBW("divcd3", "0", 2)
    Else
        SetNoGiveInfo = "insert into " & T_COM099 & "(noindex,divcd1,divcd2,divcd3,seq)" & _
                        " values ( " & DBV("cdindex", CDINDEX, 1) & DBV("divcd1", "B", 1) & _
                                     DBV("divcd2", "0", 1) & DBV("divcd3", "0", 1) & DBV("seq", seq, 1) & ")"
    End If
End Function


