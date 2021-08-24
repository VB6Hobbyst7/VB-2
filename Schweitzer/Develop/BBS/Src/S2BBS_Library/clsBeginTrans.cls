VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBeginTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private stsORDER   As String
Private stsCOLLECT As String
Private stsACCESS  As String
Private stsINPROCESS As String

Public Function SET_SplitInsertNew(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                                ByVal CompoCd As String, ByVal volumn As String, ByVal ABO As String, _
                                ByVal RH As String, ByVal PtId As String, ByVal reserved As String, _
                                ByVal autofg As String, ByVal coldt As String, ByVal coltm As String, _
                                ByVal colid As String, ByVal available As String, ByVal expdt As String, _
                                ByVal exptm As String, ByVal EntDt As String, ByVal EntTm As String, _
                                ByVal entid As String, ByVal centercd As String, ByVal hosfg As String, _
                                ByVal Donorid As String, ByVal donoraccdt As String) As String
'---------------------------------------------------------------------
'혈액분획 (BBS302)
',SET_SplitUpdate(분획출고상태로Update)
'---------------------------------------------------------------------
    SET_SplitInsertNew = " INSERT INTO " & T_BBS401 & "(" & _
                       " bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg," & _
                       " coldt,coltm,colid,available,expdt,exptm,entdt,enttm,entid,centercd," & _
                       " stscd,hosfg,splitoutfg,splitinfg,splitoutdt,donorid,donoraccdt)" & _
                       " values(" & _
                       DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", CompoCd, 1) & DBV("volumn", volumn, 1) & _
                       DBV("abo", ABO, 1) & DBV("rh", RH, 1) & DBV("ptid", PtId, 1) & DBV("reserved", reserved, 1) & DBV("autofg", autofg, 1) & _
                       DBV("coldt", coldt, 1) & DBV("coltm", coltm, 1) & DBV("colid", colid, 1) & DBV("available", available, 1) & DBV("expdt", expdt, 1) & _
                       DBV("exptm", exptm, 1) & DBV("entdt", EntDt, 1) & DBV("enttm", EntTm, 1) & DBV("entid", entid, 1) & DBV("centercd", centercd, 1) & _
                       DBV("stscd", BBSBloodStatus.stsENTER, 1) & DBV("hosfg", hosfg, 1) & DBV("splitoutfg", "0", 1) & DBV("splitinfg", "1", 1) & _
                       DBV("splitoutdt", "1", 1) & DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt) & ")"
                       
End Function

Public Function Set_SplitUpdate(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                ByVal Bldno As String, ByVal CompoCd As String) As String
'---------------------------------------------------------------------
'혈액분획 (BBS302)
'SET_SplitUpdate(분획출고상태로Update)
'---------------------------------------------------------------------
    Dim strSplitOutDt As String
    
    strSplitOutDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    Set_SplitUpdate = " UPDATE " & T_BBS401 & " SET " & _
                                                DBW("splitoutfg", "1", 3) & _
                                                DBW("splitinfg", "0", 3) & _
                                                DBW("stscd", BBSBloodStatus.stsDELIVERY, 3) & _
                                                DBW("splitoutdt", strSplitOutDt, 2) & _
                      " WHERE " & _
                                            DBW("bldsrc=", Bldsrc) & _
                                "   AND " & DBW("bldyy=", Bldyy) & _
                                "   AND " & DBW("bldno=", Bldno) & _
                                "   AND " & DBW("compocd=", CompoCd)
                        
End Function
Public Function Set_BBS202(ByVal accdt As String, ByVal accseq As Long, ByVal PtId As String, _
                           ByVal orddt As String, ByVal OrdNo As Long, ByVal OrdSeq As Long, _
                           ByVal rcvid As String, ByVal phere As String) As String
'-------------------------------------------
'처방 접수내역 INSERT 한다
'수혈처방 접수내역에서만 사용한다(frmBBS102)
'-------------------------------------------
    Dim strrcvtm As String
    Dim strRcvdt As String
    
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    Set_BBS202 = " INSERT INTO " & T_BBS202 & "(" & _
                 " workarea,accdt,accseq,ptid,orddt,ordno,ordseq,rcvdt,rcvtm,rcvid,stscd," & _
                 " canceldt,canceltm,cancelid,cancelfg,cancelrsn,pheresisfg)" & _
                 " values(" & _
                 DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & _
                 DBV("ordno", OrdNo, 1) & DBV("ordseq", OrdSeq, 1) & DBV("rcvdt", strRcvdt, 1) & DBV("rcvtm", strrcvtm, 1) & DBV("rcvid", rcvid, 1) & _
                 DBV("stscd", stsACCESS, 1) & DBV("canceldt", "", 1) & DBV("canceltm", "", 1) & DBV("cancelid", "", 1) & DBV("cancelfg", "0", 1) & _
                 DBV("cancelrsn", "", 1) & DBV("pheresisfg", phere) & ")"
                 


End Function

Public Function Set_BBS202_Insert(ByVal accdt As String, ByVal accseq As String, ByVal PtId As String, _
                            ByVal orddt As String, ByVal OrdNo As Long, ByVal OrdSeq As String, _
                            ByVal rcvid As String, ByVal phere As String) As String
'-------------------------------------------
'처방 접수내역 INSERT 한다
'수혈처방 접수내역에서만 사용한다(frmBBS102)
'-------------------------------------------
    Dim strrcvtm As String
    Dim strRcvdt As String
    
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    Set_BBS202_Insert = " INSERT INTO " & T_BBS202 & "(" & _
                 " workarea,accdt,accseq,ptid,orddt,ordno,ordseq,rcvdt,rcvtm,rcvid,stscd," & _
                 " canceldt,canceltm,cancelid,cancelfg,cancelrsn,pheresisfg)" & _
                 " values(" & _
                 DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & _
                 DBV("ordno", OrdNo, 1) & DBV("ordseq", OrdSeq, 1) & DBV("rcvdt", strRcvdt, 1) & DBV("rcvtm", strrcvtm, 1) & DBV("rcvid", rcvid, 1) & _
                 DBV("stscd", stsACCESS, 1) & DBV("canceldt", "", 1) & DBV("canceltm", "", 1) & DBV("cancelid", "", 1) & DBV("cancelfg", "0", 1) & _
                 DBV("cancelrsn", "", 1) & DBV("pheresisfg", phere) & ")"
                 


End Function

'
'
Public Function Set_UpdateL101(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String) As String
'----------------------
'접수시 처방헤더 UPDATE
'----------------------
    
'    Set_UpdateL101 = " UPDATE " & T_LAB101 & _
'                     " SET    " & DBW("donefg", stsACCESS, 2) & _
'                     " WHERE  " & _
'                                        DBW("ptid=  ", PtId) & _
'                              " AND " & DBW("orddt=", orddt) & _
'                              " AND " & DBW("ordno=", OrdNo)
    
    '전주 예수병원용 처방 업뎃을 위해 사용
    Set_UpdateL101 = " update mdbldort set " & DBW("donefg", stsACCESS, 2) & _
                    " where " & DBW("patno=", PtId) & _
                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                    " and " & DBW("ordseqno=", OrdNo)
End Function

Public Function Set_UpdateL102(ByVal PtId As String, ByVal orddt As String, _
                               ByVal OrdNo As String, ByVal OrdSeq As String, _
                               ByVal accdt As String, ByVal accseq As String) As String
'----------------------
'접수시 처방바디 UPDATE
'----------------------
    Dim strRcvdt As String
    Dim strrcvtm As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
'    Set_UpdateL102 = " UPDATE " & T_LAB102 & " SET " & _
'                                               DBW("rcvdt", strRcvdt, 3) & _
'                                               DBW("rcvtm", strrcvtm, 3) & _
'                                               DBW("donefg", stsACCESS, 3) & _
'                                               DBW("stscd", stsACCESS, 3) & _
'                                               DBW("workarea", C_WORKAREA, 3) & _
'                                               DBW("accdt", accdt, 3) & _
'                                               DBW("accseq", accseq, 2) & _
'                            " WHERE " & _
'                                               DBW("ptid=", PtId) & _
'                                    " AND  " & DBW("orddt=", orddt) & _
'                                    " AND  " & DBW("ordno=", OrdNo) & _
'                                    " AND  " & DBW("ordseq=", OrdSeq)
             
    '전주 예수병원용 처방 업뎃을 위해 사용
'상태전송 미포함
'    Set_UpdateL102 = " UPDATE mdbldort SET " & _
'                                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss') ," & _
'                                               DBW("donefg", stsACCESS, 3) & _
'                                               DBW("stscd", stsACCESS, 3) & _
'                                               DBW("workarea", C_WORKAREA, 3) & _
'                                               DBW("accdt", accdt, 3) & _
'                                               DBW("accseq", accseq, 2) & _
'                            " WHERE " & DBW("patno=", PtId) & _
'                                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
'                                    " AND  " & DBW("ordseqno=", OrdNo)
'상태전송 포함
'OCS에 상태를 C로 전송 임상병리 접수 Status E와 다름.
    Set_UpdateL102 = " UPDATE mdbldort SET " & _
                                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss') ," & _
                                               DBW("donefg", stsACCESS, 3) & _
                                               DBW("stscd", stsACCESS, 3) & _
                                               DBW("workarea", C_WORKAREA, 3) & _
                                               DBW("accdt", accdt, 3) & _
                                               DBW("accseq", accseq, 3) & _
                                               DBW("procstat", "C", 2) & _
                            " WHERE " & DBW("patno=", PtId) & _
                                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                                    " AND  " & DBW("ordseqno=", OrdNo)
End Function


Public Function Set_UpdateB206(ByVal centercd As String, ByVal legcd As String, _
                               ByVal rowno As Long, ByVal colno As Long, _
                               ByVal spcyy As String, ByVal spcno As String) As String

'------------------------------
'검체보관장소 UPDATE (검체정보)
'------------------------------
    
    
    
    Set_UpdateB206 = " UPDATE " & T_BBS206 & " SET " & _
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
'채혈내역 UPDATE
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
    
    Set_UpdateB201 = " UPDATE " & T_BBS201 & " SET" & _
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
'1: 처방 헤더 UPDATE
'2: 처방 바디 UPDATE
'3: 처방 접수내역 UPDATE
'4: 처방 Status UPDATE
'5: 결과내역 UPDATE
'6: 혈액입고내역 UPDATE
'-----------------------
    Dim DrRS        As Recordset
    Dim accdt       As String
    Dim accseq      As String
    Dim strCancelDt As String
    Dim strCancelTm As String
    Dim sSql        As String
    
    Dim orddt       As String
    Dim OrdNo       As String
    Dim OrdSeq      As String
    Dim Bussdiv     As String
    Dim OcsOrdNo    As String
    
    Dim i           As Long
    
    accdt = medGetP(AccNo, 1, "-")
    accseq = medGetP(AccNo, 2, "-")
    strCancelDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strCancelTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    

    sSql = " SELECT a.ptid,a.orddt,a.ordno,a.ordseq,a.ocsordno,b.bussdiv " & _
           " FROM " & T_LAB101 & " b," & T_LAB102 & " a " & _
           " WHERE " & DBW("a.ptid=", PtId) & _
           " AND   " & DBW("a.workarea=", C_WORKAREA) & _
           " AND   " & DBW("a.accdt=", accdt) & _
           " AND   " & DBW("a.accseq=", accseq) & _
           " and a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno"
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
        Set DrRS = Nothing
        Exit Function
    Else
        orddt = DrRS.Fields("orddt").Value & ""
        OrdNo = DrRS.Fields("ordno").Value & ""
        OrdSeq = DrRS.Fields("ordseq").Value & ""
        OcsOrdNo = Val(Trim(DrRS.Fields("ocsordno").Value & ""))
        Bussdiv = Trim(DrRS.Fields("bussdiv").Value & "")
                
        Set DrRS = Nothing
    End If
    
On Error GoTo Save_CanCel_Error
    DBConn.BeginTrans
    '-----------------
    'OCS Acting Update
    '-----------------
    

'    If OcsOrdNo <> "" Then
'        '병동은 ipd_order_dmc,ipd_order_update_dmc 업데이트
'        '외래는 opd_order_dmc 업데이트
'        If Bussdiv = enBussDiv.BussDiv_InPatient Then
'            sSql = " UPDATE med_ocs.ipd_order_dmc SET acting_check='0' where order_key=" & OcsOrdNo
'            DBConn.Execute sSql
'            sSql = " UPDATE med_ocs.ipd_order_update_dmc SET acting_check='0' where order_key=" & OcsOrdNo
'        Else
'            sSql = " UPDATE med_ocs.opd_order_dmc SET acting_check='0' where order_key=" & OcsOrdNo
'            DBConn.Execute sSql
'        End If
'    End If
    
'    MsgBox "처방내역 업데이트", vbExclamation
    
    '---------------
    '처방헤더 UPDATE
    '---------------
'     sSql = " UPDATE " & T_LAB101 & _
'            " SET    " & DBW("donefg", stscd, 2) & _
'            " WHERE  " & DBW("ptid=", PtId) & _
'            " AND    " & DBW("orddt=", orddt) & _
'            " AND    " & DBW("ordno=", OrdNo)
            
    '전주 예수병원용
    sSql = " update mdbldort set " & DBW("donefg", stscd, 2) & _
            " where " & DBW("patno=", PtId) & _
            " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
            " and " & DBW("ordseqno=", OrdNo)
    
    
     DBConn.Execute sSql
    '---------------
    '처방바디 UPDATE
    '---------------
'    sSql = " UPDATE " & T_LAB102 & _
'           " SET " & _
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
'예수병원용
'상태전송 미포함
'    sSql = " update mdbldort set " & _
'            DBW("donefg", stscd, 3) & _
'            DBW("stscd", stscd, 3) & _
'            " rcpdate=null, workarea='', accdt='', accseq =0 " & _
'            " where " & DBW("patno=", PtId) & _
'            " and " & DBW("workarea=", C_WORKAREA) & _
'            " and " & DBW("accdt=", accdt) & _
'            " and " & DBW("accseq=", accseq)
'상태전송 포함
'BBSOrdStatus.stsORDER 처방상태로 취소하는 경우에는 상태가 null이고
'BBSOrderStatus.stsCOLLECT 채혈상태로 취소하는 경우에는 상태가 B
    
    sSql = " update mdbldort set " & _
            DBW("donefg", stscd, 3) & _
            DBW("stscd", stscd, 3) & _
            DBW("procstat", IIf(stscd = BBSOrdStatus.stsORDER, "", "B"), 3) & _
            " rcpdate=null, workarea='', accdt='', accseq =0 " & _
            " where " & DBW("patno=", PtId) & _
            " and " & DBW("workarea=", C_WORKAREA) & _
            " and " & DBW("accdt=", accdt) & _
            " and " & DBW("accseq=", accseq)

    DBConn.Execute sSql
    '-------------------
    '처방접수내역 UPDATE
    '-------------------
    sSql = " UPDATE " & T_BBS202 & _
           " SET " & _
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
    'xm 결과등록내역 UPDATE
    '----------------------
    sSql = " UPDATE " & T_BBS302 & " SET " & _
                                    DBW("cancelfg", BBSCancelStatus.stsCancel, 3) & _
                                    DBW("canceldt", strCancelDt, 3) & _
                                    DBW("cancelid", cancelid, 2) & _
            " WHERE " & _
                    "     " & DBW("workarea=", C_WORKAREA) & _
                    " AND " & DBW("accdt=        ", accdt) & _
                    " AND " & DBW("accseq=      ", accseq)
    
    DBConn.Execute sSql
    '-----------------------
    '처방 status 내역 UPDATE
    '----------------------
    sSql = " UPDATE " & T_BBS203 & " SET " & DBW("assigncancelcnt", cCnt, 2) & _
            " WHERE " & _
                    "     " & DBW("workarea=", C_WORKAREA) & _
                    " AND " & DBW("accdt=        ", accdt) & _
                    " AND " & DBW("accseq=      ", accseq)
    DBConn.Execute sSql
    '-------------------
    '혈액입고내역 UPDATE
    '-------------------
    sSql = " SELECT bldsrc,bldyy,bldno,compocd " & _
           " FROM " & T_BBS302 & _
           " WHERE " & DBW("workarea=", C_WORKAREA) & _
           " AND   " & DBW("accdt=", accdt) & _
           " AND   " & DBW("accseq=", accseq)
           
    Set DrRS = New Recordset    'Assign이 안되었어도 접수취소 가능하도록 수정
    DrRS.Open sSql, DBConn
    
    If DrRS.EOF Then '일반 사용자는 취소할 수 없도록... 일반사용자는 Assign된 경우만 취소 가능
        If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor Or ObjMyUser.IsManager) Then
            Set DrRS = Nothing
            GoTo Save_CanCel_Error
        End If
    End If
    
    Do Until DrRS.EOF
        sSql = " UPDATE " & T_BBS401 & _
               " SET " & DBW("stscd", BBSBloodStatus.stsENTER, 2) & _
               " WHERE " & DBW("bldsrc=", DrRS.Fields("bldsrc").Value & "") & _
               " AND   " & DBW("bldyy=", DrRS.Fields("bldyy").Value & "") & _
               " AND   " & DBW("bldno=", DrRS.Fields("bldno").Value & "") & _
               " AND   " & DBW("compocd=", DrRS.Fields("compocd").Value & "")
        
        DBConn.Execute sSql
        
        DrRS.MoveNext
    Loop
    
    Set DrRS = Nothing
    
    '------------------------
    '검체별 처방내역 delete
    '------------------------
    If stscd = stsORDER Then
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
'    MsgBox Err.Description, vbExclamation
End Function
                                 



Public Function B_Existence(ByVal PtId As String, ByVal coldt As String, ByVal coltm As String) As Boolean
'------------------------------------------------
'검체채취시간을 비교하여 검체 채취여부를 판단한다
'True: 채취한다, False:채취않한다
'------------------------------------------------
    Dim DrRS As New Recordset
    Dim strSQL As String
    Dim strExistence As String
    Dim lngStoreHour As Long
    Dim strCompare As String
    Dim strCompare1 As String
    
    strCompare1 = coldt & Mid(coltm, 1, 4)
    strCompare1 = Format(strCompare1, "####-##-## ##:##")
    
    strSQL = " SELECT coldt,coltm FROM " & T_BBS201 & " a" & _
             " WHERE a.coldt=(SELECT max(b.coldt) " & _
                            " FROM " & T_BBS201 & " b" & _
                            " WHERE  " & _
                                    "" & DBV("b.ptid", PtId, 2) & _
                                    " AND a.ptid=b.ptid " & _
                                    " AND " & DBW("b.coldt<", coldt, 2) & ")" & _
             " ORDER BY a.coltm desc "
    
    Set DrRS = New Recordset
    DrRS.Open strSQL, DBConn
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
'검체추가요청서를 UPDATE 해준다
'추가요청이 나서 검체채취를 한경우
'--------------------------------
    Set_SpcAdder = " UPDATE " & T_BBS207 & " SET " & _
                                            DBW("spcyy", spcyy, 3) & _
                                            DBW("spcno", spcno, 3) & _
                                            DBW("donefg", stsCOLLECT, 2) & _
                   " WHERE " & _
                           "     " & DBW("ptid", PtId, 2) & _
                           " AND " & DBW("reqdt", Reqdt, 2) & _
                           " AND " & DBW("dongfg", stsORDER, 2)

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
                    " AND " & DBW("dongfg", stsORDER, 2)
    
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
    
    sSql = " INSERT INTO " & T_BBS207 & _
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

    GetSQL_CancelBloodResreved = " UPDATE " & T_BBS401 & " " & _
                                 " SET " & DBW("reserved=", "0") & _
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
    SetBldStorageUpdateByCenterCd = " UPDATE " & T_BBS401 & _
                                    " SET    " & DBW("centercd", centercd, 2) & _
                                    " WHERE" & _
                                             "     " & DBW("bldsrc", Bldsrc, 2) & _
                                             " AND " & DBW("bldyy", Bldyy, 2) & _
                                             " AND " & DBW("bldno", Bldno, 2) & _
                                             " AND " & DBW("compocd", CompoCd, 2)


End Function

Public Function SetTransInfoSSQL(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                 ByVal Bldno As Long, ByVal CompoCd As String, _
                                 ByVal BuildCd As String, _
                                 ByVal ABO As String, ByVal RH As String, ByVal Vol As String) As String
'------------------------------------------
'혈액이동내역을 삽입한다.
'------------------------------------------
    SetTransInfoSSQL = " INSERT INTO h7bbs409(bldsrc,bldyy,bldno,compocd,centercd,buildcd,abo,rh,volume,transdt,transtm,colid) " & _
                       " VALUES (" & _
                       DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & _
                       DBV("compocd", CompoCd, 1) & DBV("centercd", ObjSysInfo.buildingcd, 1) & DBV("buildcd", BuildCd, 1) & _
                       DBV("abo", ABO, 1) & DBV("rh", RH, 1) & DBV("volume", Vol, 1) & _
                       DBV("transdt", Format(GetSystemDate, "YYYYMMDD"), 1) & _
                       DBV("transtm", Format(GetSystemDate, "HHMMSS"), 1) & _
                       DBV("colid", ObjSysInfo.EmpID) & ")"


End Function

Public Function SetTransDelteSSQL(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                  ByVal Bldno As Long, ByVal CompoCd As String, _
                                  ByVal centercd As String) As String
'------------------------------------------
'혈액이동내역을 삭제한다.
'------------------------------------------
    SetTransDelteSSQL = " DELETE h7bbs409" & _
                        " WHERE" & _
                                 "     " & DBW("bldsrc", Bldsrc, 2) & _
                                 " AND " & DBW("bldyy", Bldyy, 2) & _
                                 " AND " & DBW("bldno", Bldno, 2) & _
                                 " AND " & DBW("compocd", CompoCd, 2) & _
                                 " AND " & DBW("centercd", centercd, 2)
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
    
    SetBldDelivery = " INSERT INTO " & T_BBS402 & _
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

Public Function SetLocalDelivery(ByVal Bldsrc As String, ByVal Bldyy As String, _
                               ByVal Bldno As Long, ByVal CompoCd As String, _
                               ByVal deliverydt As String, ByVal Deliveryseq As Long, _
                               ByVal deliverytm As String, ByVal Deliveryid As String, _
                               ByVal rcvid As String, ByVal WorkArea As String, _
                               ByVal accdt As String, ByVal accseq As String, _
                               ByVal OrdCd As String, _
                               ByVal LocalCd As String, _
                               ByVal remark As String, _
                               ByVal PtNm As String, _
                               ByVal ABO As String, _
                               Optional ByVal RstSeq As String = "0") As String
'---------------------------------------------
'To BBS402
'혈액 출고 내역(Local 병원 의뢰 혈액(frmBBS311)
'---------------------------------------------
    
    SetLocalDelivery = " INSERT INTO " & T_BBS402 & _
                     " (bldsrc,bldyy,bldno,compocd,deliverydt," & _
                     " deliveryseq,deliverytm,deliveryid,rcvid,workarea," & _
                     " accdt,accseq,rstseq,ordcd,localcd,rmk,ptnm,abo) " & _
                     " values(" & _
                                DBV("bldsrc", Bldsrc) & "," & DBV("bldyy", Bldyy) & "," & _
                                DBV("bldno", Bldno) & "," & DBV("compocd", CompoCd) & "," & _
                                DBV("deliverydt", deliverydt) & "," & DBV("deliveryseq", Deliveryseq) & "," & _
                                DBV("deliverytm", deliverytm) & "," & DBV("deliveryid", Deliveryid) & "," & _
                                DBV("rcvid", rcvid) & "," & DBV("workarea", WorkArea) & "," & _
                                DBV("accdt", accdt) & "," & DBV("accseq", accseq) & "," & _
                                DBV("rstseq", RstSeq) & "," & DBV("ordcd", OrdCd) & "," & _
                                DBV("localcd", LocalCd) & "," & _
                                DBV("rmk", remark) & "," & _
                                DBV("ptnm", PtNm) & "," & _
                                DBV("abo", ABO) & _
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
    SetBldStorageUpdateByStsCd = "UPDATE " & T_BBS401 & _
                                 " SET   " & DBW("stscd", stscd, 2) & _
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
        SetDonorMST = "UPDATE " & T_BBS601 & _
                      " SET  " & _
                                DBW("donornm", donornm, 3) & DBW("ssn", ssn, 3) & _
                                DBW("dob", dob, 3) & DBW("sex", sex, 3) & _
                                DBW("zipcd", ZipCd, 3) & DBW("addr1", Addr1, 3) & _
                                DBW("addr2", Addr2, 3) & DBW("telno", TelNo, 3) & _
                                DBW("jobcd", JobCd, 3) & DBW("abo", ABO, 3) & _
                                DBW("rh", RH, 3) & DBW("cnt", cnt, 3) & DBW("totvol", totvol, 2) & _
                      " WHERE " & DBW("donorid", Donorid, 2)
    Else
        SetDonorMST = "INSERT INTO " & T_BBS601 & "(" & _
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
        SetNoGiveInfo = " UPDATE " & T_COM099 & _
                        " SET " & DBW("seq=", seq) & _
                        " WHERE" & _
                                 "     " & DBW("noindex", CDINDEX, 2) & _
                                 " AND " & DBW("divcd1", C_WORKAREA, 2) & _
                                 " AND " & DBW("divcd2", "0", 2) & _
                                 " AND " & DBW("divcd3", "0", 2)
    Else
        SetNoGiveInfo = "INSERT INTO " & T_COM099 & "(noindex,divcd1,divcd2,divcd3,seq)" & _
                        " values ( " & DBV("cdindex", CDINDEX, 1) & DBV("divcd1", "B", 1) & _
                                     DBV("divcd2", "0", 1) & DBV("divcd3", "0", 1) & DBV("seq", seq, 1) & ")"
    End If
End Function

Public Function Change_Location(ByVal spcyy As String, ByVal spcno As String, ByVal buildingcd As String) As String
'----------------------------------------
'2001-11-29추가 : 환자의 검사장소 변경
'----------------------------------------
    Change_Location = " UPDATE " & T_BBS201 & _
                      " SET    " & DBW("buildcd", buildingcd, 2) & _
                      " WHERE  " & _
                                        DBW("spcyy=  ", spcyy) & _
                              " AND " & DBW("spcno=", spcno)


End Function

Public Function TransBloodSQL(ByVal centercd As String, ByVal FDate As String, ByVal tDate As String) As String
    TransBloodSQL = " SELECT a.transdt,a.bldsrc,a.bldyy,a.bldno,a.volume,a.rh,a.abo,a.centercd,a.buildcd,b.abbrnm " & _
                    " FROM " & T_BBS006 & " b," & T_BBS409 & " a " & _
                    " WHERE " & _
                              DBW("a.transdt>=", FDate) & _
                    " AND " & DBW("a.transdt<=", tDate) & _
                    " AND " & DBW("a.centercd=", centercd) & _
                    " AND  a.compocd=b.compocd" & _
                    " ORDER BY a.bldsrc,a.bldyy,a.bldno"
                    
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
