VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSSQLStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private stsORDER        As String
Private stsCOLLECT      As String
Private stsACCESS       As String
Private stsINPROCESS    As String


Public Function SetBldFileStorage(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                  ByVal Bldno As String, ByVal CompoCd As String, _
                                  ByVal volumn As String, ByVal ABO As String, _
                                  ByVal RH As String, ByVal PtId As String, _
                                  ByVal reserved As String, ByVal autofg As String, _
                                  ByVal coldt As String, ByVal coltm As String, _
                                  ByVal colid As String, ByVal available As String, _
                                  ByVal expdt As String, ByVal exptm As String, _
                                  ByVal EntDt As String, ByVal EntTm As String, _
                                  ByVal entid As String, ByVal centercd As String, _
                                  ByVal stscd As String) As String
    '-------------------------------------------------------------------------------
    'To BBS401:혈액 입고 내역(울산동강병원)
    '-------------------------------------------------------------------------------
    SetBldFileStorage = "insert into " & T_BBS401 & _
                       "(bldsrc,bldyy,bldno,compocd,volumn," & _
                       "abo,rh,ptid,reserved,autofg," & _
                       "coldt,coltm,colid,available,expdt," & _
                       "exptm,entdt,enttm,entid,centercd," & _
                       "stscd,hosfg,splitoutfg,splitinfg,splitoutdt) " & _
                       "values(" & _
                        DBV("bldsrc", Bldsrc) & "," & DBV("bldyy", Bldyy) & "," & DBV("bldno", Bldno) & "," & DBV("compocd", CompoCd) & "," & DBV("volumn", volumn) & "," & _
                        DBV("abo", ABO) & "," & DBV("rh", RH) & "," & DBV("ptid", PtId) & "," & DBV("reserved", reserved) & "," & DBV("autofg", autofg) & "," & _
                        DBV("coldt", coldt) & "," & DBV("coltm", coltm) & "," & DBV("colid", colid) & "," & DBV("available", available) & "," & DBV("expdt", expdt) & "," & _
                        DBV("exptm", exptm) & "," & DBV("entdt", EntDt) & "," & DBV("enttm", EntTm) & "," & DBV("entid", entid) & "," & DBV("centercd", centercd) & "," & _
                        DBV("stscd", stscd) & "," & DBV("hosfg", "0") & "," & DBV("splitoutfg", "0") & "," & DBV("splitinfg", "0") & "," & DBV("splitoutdt", "") & ")"
End Function

Public Function SetBldFileStorage_2014(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                  ByVal Bldno As String, ByVal CompoCd As String, _
                                  ByVal volumn As String, ByVal ABO As String, _
                                  ByVal RH As String, ByVal PtId As String, _
                                  ByVal reserved As String, ByVal autofg As String, _
                                  ByVal coldt As String, ByVal coltm As String, _
                                  ByVal colid As String, ByVal available As String, _
                                  ByVal expdt As String, ByVal exptm As String, _
                                  ByVal EntDt As String, ByVal EntTm As String, _
                                  ByVal entid As String, ByVal centercd As String, _
                                  ByVal stscd As String, ByVal larC As String, ByVal smlC As String, ByVal larE As String, ByVal smlE As String) As String
    '-------------------------------------------------------------------------------
    'To BBS401:혈액 입고 내역(울산동강병원)
    '-------------------------------------------------------------------------------
    SetBldFileStorage_2014 = "insert into " & T_BBS401 & _
                       "(bldsrc,bldyy,bldno,compocd,volumn," & _
                       "abo,rh,ptid,reserved,autofg," & _
                       "coldt,coltm,colid,available,expdt," & _
                       "exptm,entdt,enttm,entid,centercd," & _
                       "stscd,hosfg,splitoutfg,splitinfg,splitoutdt, larc, smlc, lare, smle ) " & _
                       "values(" & _
                        DBV("bldsrc", Bldsrc) & "," & DBV("bldyy", Bldyy) & "," & DBV("bldno", Bldno) & "," & DBV("compocd", CompoCd) & "," & DBV("volumn", volumn) & "," & _
                        DBV("abo", ABO) & "," & DBV("rh", RH) & "," & DBV("ptid", PtId) & "," & DBV("reserved", reserved) & "," & DBV("autofg", autofg) & "," & _
                        DBV("coldt", coldt) & "," & DBV("coltm", coltm) & "," & DBV("colid", colid) & "," & DBV("available", available) & "," & DBV("expdt", expdt) & "," & _
                        DBV("exptm", exptm) & "," & DBV("entdt", EntDt) & "," & DBV("enttm", EntTm) & "," & DBV("entid", entid) & "," & DBV("centercd", centercd) & "," & _
                        DBV("stscd", stscd) & "," & DBV("hosfg", "0") & "," & DBV("splitoutfg", "0") & "," & DBV("splitinfg", "0") & "," & DBV("splitoutdt", "") & ",'" & larC & "','" & smlC & "','" & larE & "','" & smlE & "' )"
End Function

Public Function SetBldStorage(ByVal UpdateFg As Boolean, _
                              ByVal Bldsrc As String, ByVal Bldyy As String, _
                              ByVal Bldno As Long, ByVal CompoCd As String, _
                              ByVal volumn As Long, ByVal ABO As String, _
                              ByVal RH As String, ByVal PtId As String, _
                              ByVal reserved As String, ByVal autofg As String, _
                              ByVal coldt As String, ByVal coltm As String, _
                              ByVal colid As String, ByVal available As String, _
                              ByVal expdt As String, ByVal exptm As String, _
                              ByVal EntDt As String, ByVal EntTm As String, _
                              ByVal entid As String, ByVal centercd As String, _
                              ByVal stscd As String) As String
    '-------------------------------------------------------------------------------
    'To BBS401:혈액 입고 내역
    '-------------------------------------------------------------------------------
    Dim sSql As String
    
    If UpdateFg Then
        'Update
    Else
        'Insert
        sSql = "insert into " & T_BBS401 & _
               "(bldsrc,bldyy,bldno,compocd,volumn," & _
                "abo,rh,ptid,reserved,autofg," & _
                "coldt,coltm,colid,available,expdt," & _
                "exptm,entdt,enttm,entid,centercd," & _
                "stscd,hosfg,splitoutfg,splitinfg,splitoutdt) " & _
               "values(" & _
                DBV("bldsrc", Bldsrc) & "," & DBV("bldyy", Bldyy) & "," & DBV("bldno", Bldno) & "," & DBV("compocd", CompoCd) & "," & DBV("volumn", volumn) & "," & _
                DBV("abo", ABO) & "," & DBV("rh", RH) & "," & DBV("ptid", PtId) & "," & DBV("reserved", reserved) & "," & DBV("autofg", autofg) & "," & _
                DBV("coldt", coldt) & "," & DBV("coltm", coltm) & "," & DBV("colid", colid) & "," & DBV("available", available) & "," & DBV("expdt", expdt) & "," & _
                DBV("exptm", exptm) & "," & DBV("entdt", EntDt) & "," & DBV("enttm", EntTm) & "," & DBV("entid", entid) & "," & DBV("centercd", centercd) & "," & _
                DBV("stscd", stscd) & "," & DBV("hosfg", "0") & "," & DBV("splitoutfg", "0") & "," & DBV("splitinfg", "0") & "," & DBV("splitoutdt", "") & ")"
        
        SetBldStorage = sSql
        
    End If

End Function

Public Function GetOrdCd(ByVal volumn As String, ByVal CompoCd As String, ByVal deliverydt As String, aryOrdCd() As String) As Long
    Dim sSql    As String
    Dim today   As String
    Dim Rs    As Recordset
    Dim i       As Long
    
    Erase aryOrdCd
    
    sSql = " SELECT testcd,testnm " & _
           " FROM " & T_BBS001 & " a " & _
           " WHERE  " & _
                    " " & DBW("a.compocd", CompoCd, 2) & _
                    "AND " & DBW("a.volumn", volumn, 2) & _
                    "AND applydt = (SELECT max(applydt) FROM " & T_BBS001 & " b" & _
                                  " WHERE a.testcd=b.testcd" & _
                                  " AND " & DBW("b.applydt<", deliverydt, 2) & ") " & _
           " ORDER BY testcd "

    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    With Rs
        If .RecordCount < 1 Then
            GetOrdCd = 0
        Else
            GetOrdCd = .RecordCount
            ReDim Preserve aryOrdCd(.RecordCount - 1)
            For i = 1 To .RecordCount
                aryOrdCd(i - 1) = .Fields("testcd").Value & "" & " " & .Fields("testnm").Value & ""
                .MoveNext
            Next i
        End If
        
    End With
    Set Rs = Nothing
End Function
Public Function GetMaxDeliverySeq(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                                  ByVal CompoCd As String) As Long
    Dim Rs As Recordset
    Dim sSql As String
    Dim dt   As String
    
    dt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    sSql = " SELECT max(deliveryseq) maxdeliveryseq " & _
           " FROM " & T_BBS402 & " " & _
           " WHERE" & _
                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("bldyy", Bldyy, 2) & _
                    " AND " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2) & _
                    " AND " & DBW("deliverydt", dt, 2)
                    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF = True Then
        GetMaxDeliverySeq = 0
    Else
        If IsNull(Rs.Fields("maxdeliveryseq").Value & "") = True Then
            GetMaxDeliverySeq = 0
        Else
            GetMaxDeliverySeq = Val(Rs.Fields("maxdeliveryseq").Value & "")
        End If
    End If
    Set Rs = Nothing
End Function
Public Function Get_XMethod(ByVal accdt As String, ByVal accseq As String) As Long
    Dim sSql As String
    Dim Rs   As New Recordset
    
    sSql = " SELECT b.xmethod" & _
           " FROM " & T_LAB102 & " a," & T_BBS001 & " b" & _
           " WHERE" & _
                    "     " & DBW("a.accdt", accdt, 2) & _
                    " AND " & DBW("a.accseq", accseq, 2) & _
                    " AND a.ordcd=b.testcd"
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        Get_XMethod = Rs.Fields("xmethod").Value & ""
    Else
        Get_XMethod = -1
    End If
    Set Rs = Nothing
End Function
Public Function Insert_BBS302(ByVal accdt As String, ByVal accseq As String, _
                              ByVal Bldsrc As String, ByVal Bldyy As String, _
                              ByVal Bldno As String, ByVal Compcd As String, _
                              ByVal norstdt As String, ByVal norsttm As String, _
                              ByVal norstid As String, ByVal method As String, _
                              Optional ByVal ResultTF As Boolean = False) As String
'결과없이 Assign 될수 있게 Xm결과테이블 작성
    Dim step1   As String
    Dim step2   As String
    Dim step3   As String
    Dim step4   As String
    Dim rstval  As String
    Dim xmethod As String
    
    If ResultTF = True Then
        step1 = "1": step2 = "1": step3 = "1": step4 = "1": rstval = "1": xmethod = method
    Else
        step1 = "": step2 = "": step3 = "": step4 = "": rstval = "": xmethod = ""
    End If
    
    Insert_BBS302 = " insert into " & T_BBS302 & " (" & _
                    " workarea,accdt,accseq,rstseq,bldsrc,bldyy,bldno,compocd,xmethod, " & _
                    " du,step1,step2,step3,step4,rstv,vfydt,vfytm,vfyid,stat, " & _
                    " statdt,stattm,statid,cancelid,canceldt,cancelfg,norstfg,norstdt,norsttm,norstid,rmk )" & _
                    " values (" & _
                      DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("rstseq", "1", 1) & _
                      DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", Compcd, 1) & DBV("xmethod", xmethod, 1) & _
                      DBV("du", "", 1) & DBV("step1", step1, 1) & DBV("step2", step2, 1) & DBV("step3", step3, 1) & DBV("step4", step4, 1) & DBV("rstv", rstval, 1) & _
                      DBV("vfydt", norstdt, 1) & DBV("vfytm", norsttm, 1) & DBV("vfyid", norstid, 1) & _
                      DBV("stat", "", 1) & DBV("statdt", "", 1) & DBV("stattm", "", 1) & DBV("statid", "", 1) & _
                      DBV("cancelid", "", 1) & DBV("canceldt", "", 1) & DBV("cancelfg", BBSCancelStatus.stsNotCancel, 1) & _
                      DBV("norstfg", "1", 1) & DBV("norstdt", norstdt, 1) & DBV("norsttm", norsttm, 1) & DBV("norstid", norstid, 1) & DBV("rmk", "") & ")"
                      
                        
End Function
Private Function Insert_BBS401(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                               ByVal CompoCd As String, ByVal volumn As String, ByVal ABO As String, _
                               ByVal RH As String, ByVal PtId As String, ByVal Rfg As String, _
                               ByVal Afg As String, ByVal Pfg As String, ByVal coldt As String, _
                               ByVal coltm As String, ByVal colid As String, ByVal available As String, _
                               ByVal expdt As String, ByVal exptm As String, ByVal EntDt As String, _
                               ByVal EntTm As String, ByVal entid As String, ByVal centercd As String, _
                               ByVal stscd As String) As String
'----------------
'혈액입고내역작성
'----------------
    Insert_BBS401 = " insert into " & T_BBS401 & "(" & _
                    " bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available," & _
                    " expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg," & _
                    " realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk)" & _
                    "values (" & _
                    DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", CompoCd, 1) & DBV("volumn", volumn, 1) & _
                    DBV("abo", ABO, 1) & DBV("rh", RH, 1) & DBV("ptid", PtId, 1) & DBV("reserved", Rfg, 1) & DBV("autofg", Afg, 1) & DBV("pherefg", Pfg, 1) & _
                    DBV("coldt", coldt, 1) & DBV("coltm", coltm, 1) & DBV("colid", colid, 1) & _
                    DBV("available", available, 1) & DBV("expdt", expdt, 1) & DBV("exptm", exptm, 1) & DBV("entdt", EntDt, 1) & _
                    DBV("enttm", EntTm, 1) & DBV("entid", entid, 1) & DBV("centercd", centercd, 1) & DBV("localcd", "", 1) & _
                    DBV("stscd", stscd, 1) & DBV("hosfg", "", 1) & DBV("splitoutfg", "", 1) & DBV("splitinfg", "0", 1) & DBV("realexpdt", "", 1) & DBV("realexptm", "", 1) & _
                    DBV("expid", 0, 1) & DBV("exprcvid", 0, 1) & DBV("expbilldiv", "", 1) & DBV("exprsnrmk", "") & ")"
End Function
'2001-11-27추가 : 혈액입고내역에 헌혈자ID와 헌혈일 추가 (donorid, donoraccdt)
Private Function Insert_BBS401_donor(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                               ByVal CompoCd As String, ByVal volumn As String, ByVal ABO As String, _
                               ByVal RH As String, ByVal PtId As String, ByVal Rfg As String, _
                               ByVal Afg As String, ByVal Pfg As String, ByVal coldt As String, _
                               ByVal coltm As String, ByVal colid As String, ByVal available As String, _
                               ByVal expdt As String, ByVal exptm As String, ByVal EntDt As String, _
                               ByVal EntTm As String, ByVal entid As String, ByVal centercd As String, _
                               ByVal stscd As String, ByVal Donorid As String, ByVal donordt As String) As String
'----------------
'혈액입고내역작성
'----------------
    Insert_BBS401_donor = " insert into " & T_BBS401 & "(" & _
                    " bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available," & _
                    " expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg," & _
                    " realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk,donorid,donoraccdt)" & _
                    "values (" & _
                    DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", CompoCd, 1) & DBV("volumn", volumn, 1) & _
                    DBV("abo", ABO, 1) & DBV("rh", RH, 1) & DBV("ptid", PtId, 1) & DBV("reserved", Rfg, 1) & DBV("autofg", Afg, 1) & DBV("pherefg", Pfg, 1) & _
                    DBV("coldt", coldt, 1) & DBV("coltm", coltm, 1) & DBV("colid", colid, 1) & _
                    DBV("available", available, 1) & DBV("expdt", expdt, 1) & DBV("exptm", exptm, 1) & DBV("entdt", EntDt, 1) & _
                    DBV("enttm", EntTm, 1) & DBV("entid", entid, 1) & DBV("centercd", centercd, 1) & DBV("localcd", "", 1) & _
                    DBV("stscd", stscd, 1) & DBV("hosfg", "1", 1) & DBV("splitoutfg", "", 1) & DBV("splitinfg", "0", 1) & DBV("realexpdt", "", 1) & DBV("realexptm", "", 1) & _
                    DBV("expid", 0, 1) & DBV("exprcvid", 0, 1) & DBV("expbilldiv", "", 1) & DBV("exprsnrmk", "", 1) & DBV("donorid", Donorid, 1) & DBV("donoraccdt", donordt) & ")"
End Function

Private Function Update_BBS401(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String, ByVal stscd As String) As String
'----------------
'혈액입고내역수정
'----------------
    Update_BBS401 = "update " & T_BBS401 & _
                    " set   " & DBW("stscd", stscd, 2) & _
                    " WHERE " & _
                              "     " & DBW("bldsrc", Bldsrc, 2) & _
                              " AND " & DBW("bldyy", Bldyy, 2) & _
                              " AND " & DBW("bldno", Bldno, 2) & _
                              " AND " & DBW("compocd", CompoCd, 2)
End Function
Private Function Update_BBS602(ByVal DonationDt As String, ByVal Bldsrc As String, ByVal Bldyy As String, _
                               ByVal Bldno As String, ByVal CompoCd As String, ByVal volumn As String, _
                               ByVal Donorid As String, ByVal DonationAccdt As String) As String
'----------------
'헌혈접수내역수정
'----------------
    Update_BBS602 = " update " & T_BBS602 & _
                    " set    " & DBW("donationdt", DonationDt, 3) & _
                                 DBW("bldsrc", Bldsrc, 3) & _
                                 DBW("bldyy", Bldyy, 3) & _
                                 DBW("bldno", Bldno, 3) & _
                                 DBW("compocd", CompoCd, 3) & _
                                 DBW("volumn", volumn, 3) & _
                                 DBW("entfg", "1", 3) & _
                                 DBW("cancelfg", "0", 2) & _
                    " WHERE  " & _
                                     DBW("donorid", Donorid, 2) & _
                           " AND " & DBW("donoraccdt", DonationAccdt, 2)

End Function
Private Function Update_BBS603(ByVal stscd As String, ByVal Donorid As String, ByVal DonationAccdt As String) As String
'----------------
'헌혈판정내역수정
'----------------
    Update_BBS603 = " update " & T_BBS603 & _
                    " set    " & DBW("stscd", stscd, 2) & _
                    " WHERE  " & _
                                  DBW("donorid", Donorid, 2) & _
                        " AND " & DBW("donoraccdt", DonationAccdt, 2)
End Function
Private Function Insert_BBS402(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                               ByVal CompoCd As String, ByVal EntDt As String, ByVal Deliveryseq As String, _
                               ByVal EntTm As String, ByVal entid As String, ByVal accdt As String, _
                               ByVal accseq As String, ByVal BordCd As String, ByVal IrradFg As String, _
                               ByVal FilterFg As String) As String
'-----------------
'혈액출고내역 작성
'-----------------
    Insert_BBS402 = " insert into " & T_BBS402 & "(" & _
                  " bldsrc,bldyy,bldno,compocd,deliverydt,deliveryseq,deliverytm,deliveryid,rcvid,workarea," & _
                  " accdt,accseq,rstseq,ordcd,localcd,irrafg,filter,retfg,expfg,bagfg,retdt,expdt,bagdt," & _
                  " rettm,exptm,bagtm,retid,retrcvid,bagid,bagrcvid,retrmk,rmk) values (" & _
                  DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", CompoCd, 1) & DBV("deliverydt", EntDt, 1) & DBV("deliveryseq", Deliveryseq, 1) & _
                  DBV("deliverytm", EntTm, 1) & DBV("deliveryid", entid, 1) & DBV("rcvid", "", 1) & DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & _
                  DBV("rstseq", "", 1) & DBV("ordcd", BordCd, 1) & DBV("localcd", "", 1) & DBV("irrafg", IrradFg, 1) & DBV("filter", FilterFg, 1) & DBV("retfg", "", 1) & _
                  DBV("expfg", "", 1) & DBV("bagfg", "", 1) & DBV("retdt", "", 1) & DBV("expdt", "", 1) & DBV("bagdt", "", 1) & DBV("rettm", "", 1) & DBV("exptm", "", 1) & _
                  DBV("bagtm", "", 1) & DBV("retid", "", 1) & DBV("retrcvid", "", 1) & DBV("bagid", "", 1) & DBV("bagrcvid", "", 1) & DBV("retrmk", "", 1) & DBV("rmk", "") & ")"
End Function
Private Function Update_Lab102(ByVal Orderstscd As String, ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String, ByVal OrdSeq As String) As String
'-----------------
'처방바디 수정
'-----------------
'    MsgBox "처방내역 업데이트", vbExclamation
    
'    Update_Lab102 = " update " & T_LAB102 & _
'                    " set    " & DBW("stscd", Orderstscd, 3) & _
'                                 DBW("donefg", Orderstscd, 2) & _
'                    " WHERE  " & _
'                           " " & DBW("ptid", PtId, 2) & _
'                       " AND " & DBW("orddt", orddt, 2) & _
'                       " AND " & DBW("ordno", OrdNo, 2) & _
'                       " AND " & DBW("ordseq", OrdSeq, 2)
    '예수병원용
    Update_Lab102 = " update mdemxort " & _
                    " set    " & DBW("stscd", Orderstscd, 3) & _
                                 DBW("donefg", Orderstscd, 2) & _
                    " WHERE  " & _
                           " " & DBW("patno", PtId, 2) & _
                       " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                       " AND " & DBW("ordseqno", OrdNo, 2)
End Function
Private Function Update_BBS202(ByVal Orderstscd As String, ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String, ByVal OrdSeq As String) As String
'-----------------
'처방접수내역 수정
'-----------------
    Update_BBS202 = " update " & T_BBS202 & " set" & _
                           " " & DBW("stscd", Orderstscd, 2) & _
                    " WHERE  " & _
                                 DBW("ptid", PtId, 2) & _
                       " AND " & DBW("orddt", orddt, 2) & _
                       " AND " & DBW("ordno", OrdNo, 2) & _
                       " AND " & DBW("ordseq", OrdSeq, 2)
End Function
Private Function Insert_BBS606(ByVal Donorid As String, ByVal DonationAccdt As String, ByVal OrdCd As String, ByVal UnitQty As String) As String
'-----------------
'추가재료내역 작성
'-----------------
    Insert_BBS606 = " insert into " & T_BBS606 & _
                    " (donorid, donoraccdt, ordcd,qty)" & _
                    " values(" & _
                            DBV("donoird", Donorid, 1) & DBV("donoraccdt", DonationAccdt, 1) & DBV("ordcd", OrdCd, 1) & _
                            DBV("qty", UnitQty) & ")"
End Function
Private Function Insert_INT002(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String, ByVal OrdCd As String, _
                               ByVal EntDt As String, ByVal EntTm As String) As String
'-----------------
'수각계산내역 작성
'-----------------
    Insert_INT002 = " insert into " & T_INT002 & _
                    " (ptid,orddt,ordno,ordcd,orddiv,ordfg,unitqty,sysdt,systm)" & _
                    " values(" & _
                            DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & _
                            DBV("ordno", OrdNo, 1) & DBV("ordcd", OrdCd, 1) & _
                            DBV("orddiv", C_WORKAREA, 1) & DBV("ordfg", stsINPROCESS, 1) & _
                            DBV("unitqty", "1", 1) & DBV("sysdt", EntDt, 1) & DBV("systm", EntTm) & ")"
End Function
Private Function Update_BBS411(ByVal EntDt As String, ByVal entid As String, ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String) As String
'----------------
'헌혈증서내역수정
'----------------
    Update_BBS411 = " update " & T_BBS411 & _
                    " set  " & DBW("usedt", EntDt, 3) & _
                               DBW("useid", entid, 2) & _
                    " WHERE" & _
                       "   " & DBW("bldsrc", Bldsrc, 2) & _
                       " AND " & DBW("bldyy", Bldyy, 2) & _
                       " AND " & DBW("bldno", Bldno, 2)
End Function
Private Function TestdivCheck(ByVal accdt As String, ByVal accseq As String) As Boolean
    Dim sSql As String
    Dim Rs   As New Recordset
    
    sSql = " SELECT newtestdiv" & _
           " FROM " & T_LAB102 & _
           " WHERE" & _
                    "     " & DBW("workarea", C_WORKAREA, 2) & _
                    " AND " & DBW("accdt", accdt, 2) & _
                    " AND " & DBW("accseq", accseq, 2)
                    
    Call Rs.Open(sSql, DBConn)
'    If Rs.DBerror = True Then
'        dbconn.DisplayErrors
'        Set Rs = Nothing
'        Exit Function
'    End If
    If Not Rs.EOF Then
        If Rs.Fields("newtestdiv").Value & "" = "Y" Then
            TestdivCheck = True
        End If
    End If
    Set Rs = Nothing
End Function
Private Function Save_Pheresis(ByVal Bldsrc As String, ByVal Bldyy As String, _
                             ByVal Bldno As String, ByVal CompoCd As String, _
                             ByVal volumn As String, ByVal ABO As String, _
                             ByVal RH As String, ByVal PtId As String, _
                             ByVal Rfg As String, ByVal Afg As String, _
                             ByVal Pfg As String, ByVal coldt As String, _
                             ByVal coltm As String, ByVal colid As String, _
                             ByVal available As Long, ByVal expdt As String, _
                             ByVal exptm As String, ByVal EntDt As String, _
                             ByVal EntTm As String, ByVal entid As String, _
                             ByVal centercd As String, ByVal stscd As String, _
                             ByVal Donorid As String, ByVal DonationDt As String, _
                             ByVal DonationAccdt As String, ByVal Orderstscd As String, _
                             ByVal IrradFg As String, ByVal FilterFg As String, _
                             ByVal BordCd As String, ByVal accdt As String, ByVal accseq As String, _
                             ByVal Deliverychk As String, ByVal xmethod As String, ByVal objdic As clsDictionary) As Boolean

 
    Dim sSql        As String
    Dim Deliveryseq As String
    Dim isSelect    As Boolean
    Dim isTestdiv   As Boolean
    
    
    On Error GoTo Errors

    If Deliverychk = "1" Then
        '-----------------------------------------------------------------------------
        '출고내역을 작성한다.(1번:헌혈등록화면의 출고등록체크박스가 선택되어있는 경우)
        '헌혈자 stscd=6는 (2번)
        'cross-matching테이블의 결과를 insert한다.(3번)
        '-----------------------------------------------------------------------------
        '1번
        Deliveryseq = GetMaxDeliverySeq(Bldsrc, Bldyy, Bldno, CompoCd) + 1
        sSql = Insert_BBS402(Bldsrc, Bldyy, Bldno, CompoCd, EntDt, Deliveryseq, EntTm, entid, _
                             accdt, accseq, BordCd, IrradFg, FilterFg)
        DBConn.Execute sSql
        
        '2번
        
        sSql = Update_BBS603(DonorStatus.stsFinish, Donorid, DonationAccdt)
        DBConn.Execute sSql
        
        '3번
        sSql = Insert_BBS302(accdt, accseq, Bldsrc, Bldyy, Bldno, CompoCd, _
                             EntDt, EntTm, entid, xmethod, True)
        DBConn.Execute sSql
        
    Else
    
        '--------------------------------------------------------------------------------------------
        '헌혈등록화면의 출고등록체크박스가 선택되지 않았을때....
        '1번:헌혈등록화면의 결과등록체크박스가 선택되어 있을때는 결과테이블을 만들며,그렇지 않을경우는
        '2번:결과없이 결과테이블을 만들어준다........
        '--------------------------------------------------------------------------------------------
        '1번
        If stscd = BBSBloodStatus.stsASSIGN Then
            sSql = Insert_BBS302(accdt, accseq, Bldsrc, Bldyy, Bldno, CompoCd, _
                                 EntDt, EntTm, entid, xmethod, True)
        Else
        '2번
        sSql = Insert_BBS302(accdt, accseq, Bldsrc, Bldyy, Bldno, CompoCd, _
                             EntDt, EntTm, entid, xmethod)
        End If
        DBConn.Execute sSql
        
    End If
    
    '
    '접수번호 Update
    '
    sSql = "update " & T_BBS602 & " " & _
           "set " & DBW("workarea=", C_WORKAREA) & "," & _
                    DBW("accdt=", accdt) & "," & _
                    DBW("accseq=", accseq) & " " & _
           "WHERE " & DBW("donorid=", Donorid) & " " & _
           "AND   " & DBW("donoraccdt=", DonationAccdt)
    DBConn.Execute sSql
    
    '------------------------------------------------------------------------------------------------
    'Pheresis 헌혈인경우...(objdic의 recordcount가 하나라도 존재하는 경우)...........
    '1번:혈액입고내역의 stscd를 update 해준다.
    '------------------------------------------------------------------------------------------------
    '1번
    sSql = Update_BBS401(Bldsrc, Bldyy, Bldno, CompoCd, stscd)
    DBConn.Execute sSql
    
    '-----------------------------------------------
    '처방바디의 newtestdiv가 'Y'일때 istestdiv=True
    '-----------------------------------------------
    isTestdiv = TestdivCheck(accdt, accseq)
    
    '------------------------------------------------
    '처방테이블을 update 해준다.(stscd=4:완료)
    '1번:처방바디 update
    '2번:수혈처방 update
    '------------------------------------------
    
    objdic.MoveFirst
    Do Until objdic.EOF
        If isSelect = False Then
            '1번
            sSql = Update_Lab102(Orderstscd, objdic.Fields("ptid"), objdic.Fields("orddt"), _
                                            objdic.Fields("ordno"), objdic.Fields("ordseq"))
            DBConn.Execute sSql
            '2번
            sSql = Update_BBS202(Orderstscd, objdic.Fields("ptid"), objdic.Fields("orddt"), _
                                            objdic.Fields("ordno"), objdic.Fields("ordseq"))
            DBConn.Execute sSql
            isSelect = True
        End If
        
        '---------------------------------
        '추가재료를 등록한다.(BBS606)
        'objdic.Fields("div") = "2"일때...
        '---------------------------------
        If objdic.Fields("div") = "2" Then
            sSql = Insert_BBS606(Donorid, DonationAccdt, objdic.Fields("ordcd"), objdic.Fields("unitqty"))
            DBConn.Execute sSql
        End If
        
        '-------------------------------------------------------------
        '수가내역을 생성한다..(int002),처방바디의 newtestdiv='Y'인거만
        'Deliverychk = "1"일때........
        '-------------------------------------------------------------
        
        If Deliverychk = "1" Then
            If isTestdiv = True Then
                sSql = Insert_INT002(objdic.Fields("ptid"), objdic.Fields("orddt"), _
                                     objdic.Fields("ordno"), objdic.Fields("ordcd"), EntDt, EntTm)
                DBConn.Execute sSql
            End If
        End If
        objdic.MoveNext
    Loop
    Save_Pheresis = True
    Exit Function
    
Errors:
    Save_Pheresis = False
    MsgBox Err.Description, vbExclamation
End Function
Public Function Set_BldEnter(ByVal Bldsrc As String, ByVal Bldyy As String, _
                             ByVal Bldno As String, ByVal CompoCd As String, _
                             ByVal volumn As String, ByVal ABO As String, _
                             ByVal RH As String, ByVal PtId As String, _
                             ByVal Rfg As String, ByVal Afg As String, _
                             ByVal Pfg As String, ByVal coldt As String, _
                             ByVal coltm As String, ByVal colid As String, _
                             ByVal available As Long, ByVal expdt As String, _
                             ByVal exptm As String, ByVal EntDt As String, _
                             ByVal EntTm As String, ByVal entid As String, _
                             ByVal centercd As String, ByVal stscd As String, _
                             ByVal Donorid As String, ByVal DonationDt As String, _
                             ByVal DonationAccdt As String, ByVal Orderstscd As String, _
                             ByVal IrradFg As String, ByVal FilterFg As String, _
                             ByVal BordCd As String, ByVal accdt As String, ByVal accseq As String, _
                             ByVal Deliverychk As String, ByVal xmethod As String, ByVal objdic As clsDictionary) As Boolean

'헌혈등록시에 혈액이 입고되며, 헌혈자 접수내역에 혈액정보가 들어간다.
'Pheresis헌혈일 경우(혈액 입고(출고),출고내역,수가내역(추가재료까지),처방내역,수혈처방내역,
'                   헌혈접수내역,추가재료내역)저장한다.
    Dim sSql      As String

On Error GoTo Errors
    '2001-11-27수정 : 헌혈자등록일 경우 혈액입고내역에 헌혈자ID도 같이 저장한다.
    sSql = Insert_BBS401_donor(Bldsrc, Bldyy, Bldno, CompoCd, volumn, ABO, RH, PtId, Rfg, Afg, Pfg, DonationDt, coltm, _
                                colid, available, expdt, exptm, EntDt, EntTm, entid, centercd, stscd, Donorid, DonationAccdt)
    
    DBConn.Execute sSql
    '----------------------------------------
    '헌혈자 접수내역에 혈액정보를 Update 한다
    '----------------------------------------
    sSql = Update_BBS602(DonationDt, Bldsrc, Bldyy, Bldno, CompoCd, volumn, Donorid, DonationAccdt)
    DBConn.Execute sSql
    '----------------------------------------------------------
    '헌혈자 stscd를 업데이트 한다(입고만 하면 stscd는 5번이다.)
    '----------------------------------------------------------
    sSql = Update_BBS603(DonorStatus.stsDonation, Donorid, DonationAccdt)
    DBConn.Execute sSql
    '--------------------------------------------------------------------------------------------------------
    'Pheresis 헌혈인경우...(objdic의 recordcount가 하나라도 존재하는 경우)...........
    '--------------------------------------------------------------------------------------------------------
    If objdic.RecordCount > 0 Then
        If Save_Pheresis(Bldsrc, Bldyy, Bldno, CompoCd, volumn, ABO, RH, PtId, _
                         Rfg, Afg, Pfg, coldt, coltm, colid, available, expdt, _
                         exptm, EntDt, EntTm, entid, centercd, stscd, Donorid, _
                         DonationDt, DonationAccdt, Orderstscd, IrradFg, _
                         FilterFg, BordCd, accdt, accseq, Deliverychk, xmethod, objdic) = False Then
            GoTo Errors
        End If
    End If
    '-------------------------------------------------------------------------------------
    '헌혈증서내역을 update 해준다.(헌혈증서 관리 하는 경우만)
    '-------------------------------------------------------------------------------------
    sSql = Update_BBS411(EntDt, entid, Bldsrc, Bldyy, Bldno)
    DBConn.Execute sSql
    Set_BldEnter = True
    Exit Function
    
Errors:
    Set_BldEnter = False
                                                                                                            
End Function
Public Function GetCompocdPheresis() As String
    Dim Rs   As Recordset
    Dim sSql As String


    sSql = " SELECT compocd FROM " & T_BBS006 & " WHERE " & DBW("pherefg=", "1")

    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetCompocdPheresis = Rs.Fields("compocd").Value & ""
    End If
    Set Rs = Nothing
    End Function
Public Function SetPheresis(ByVal DonationDt As String, ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal volumn As String, _
                            ByVal Donorid As String, ByVal DonationAccdt As String, ByVal EntDt As String, ByVal entid As String) As Boolean
    Dim sSql      As String
    Dim CompoCd As String
    
  
    CompoCd = GetCompocdPheresis
    
On Error GoTo Set_BldEnter_error
    
    '----------------------------------------
    '헌혈자 접수내역에 혈액정보를 Update 한다
    '----------------------------------------
    sSql = Update_BBS602(DonationDt, Bldsrc, Bldyy, Bldno, CompoCd, volumn, Donorid, DonationAccdt)
    DBConn.Execute sSql
    
    
    '----------------------------------------------------------
    '헌혈자 stscd를 업데이트 한다(입고만 하면 stscd는 5번이다.)
    '----------------------------------------------------------
                      
    sSql = Update_BBS603(DonorStatus.stsDonation, Donorid, DonationAccdt)
    DBConn.Execute sSql
    
   
    '-------------------------------------------------------------------------------------
    '헌혈증서내역을 update 해준다.(헌혈증서 관리 하는 경우만)
    '-------------------------------------------------------------------------------------
    sSql = Update_BBS411(EntDt, entid, Bldsrc, Bldyy, Bldno)
    DBConn.Execute sSql
    SetPheresis = True
    Exit Function
Set_BldEnter_error:
    SetPheresis = False
End Function

Public Function SetDonorScreenCancel(ByVal Donorid As String, ByVal donoraccdt As String, ByVal tmpPtId As String) As Boolean
    Dim strSQL As String
    
'    strSql = " update " & T_LAB102 & _
'             " set    " & DBW("dcfg", "1", 2) & _
'             " WHERE" & _
'                  " " & DBW("ptid", tmpPtId, 2) & _
'             " AND  orddt in (SELECT distinct orddt " & _
'                             "FROM " & T_BBS605 & " " & _
'                             "WHERE " & DBW("donorid=", Donorid) & " " & _
'                             "AND   " & DBW("donoraccdt=", donoraccdt)
    '예수병원요
    strSQL = " update mdbldort " & _
             " set    " & DBW("dcyn", "Y", 2) & _
             " WHERE" & _
                  " " & DBW("patno", tmpPtId, 2) & _
             " AND  orddate in (SELECT distinct to_date(orddt,'yyyymmdd') " & _
                             "FROM " & T_BBS605 & " " & _
                             "WHERE " & DBW("donorid=", Donorid) & " " & _
                             "AND   " & DBW("donoraccdt=", donoraccdt)
             
On Error GoTo SetDonorScreenCancel_error

    DBConn.BeginTrans
    DBConn.Execute strSQL
    DBConn.CommitTrans
    SetDonorScreenCancel = True
    Exit Function
    
SetDonorScreenCancel_error:
    DBConn.RollbackTrans
    SetDonorScreenCancel = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function SetPheresisCancel(ByVal Donorid As String, ByVal donoraccdt As String, ByVal TmpId As String) As Boolean
                             
    Dim strSQL(1) As String
    
    strSQL(0) = " update " & T_BBS602 & _
                " set" & _
                   " " & DBW("donationdt", "", 3) & _
                   " " & DBW("bldsrc", "", 3) & _
                   " " & DBW("bldyy", "", 3) & _
                   " " & DBW("bldno", 0, 3) & _
                   " " & DBW("compocd", "", 3) & _
                   " " & DBW("volumn", 0, 3) & _
                   " " & DBW("entfg", "0", 3) & _
                   " " & DBW("cancelfg", "1", 2) & _
                " WHERE" & _
                      "" & DBW("donaraccdt", donoraccdt, 2) & _
                " AND  " & DBW("donorid", Donorid, 2)


'    strSql(1) = " update " & T_LAB102 & _
'                " set    " & DBW("dcfg", "1", 2) & _
'                " WHERE" & _
'                     " " & DBW("ptid", TmpId, 2) & _
'                " AND  orddt in (SELECT distinct orddt " & _
'                                "FROM " & T_BBS605 & " " & _
'                                "WHERE " & DBW("donorid=", Donorid) & " " & _
'                                "AND   " & DBW("donoraccdt=", donoraccdt)
    '예수병원 용
    strSQL(1) = " update  mdbldort " & _
                " set    " & DBW("dcyn", "Y", 2) & _
                " WHERE" & _
                     " " & DBW("patno", TmpId, 2) & _
                " AND  orddate in (SELECT distinct to_date(orddt,'yyyymmdd') " & _
                                "FROM " & T_BBS605 & " " & _
                                "WHERE " & DBW("donorid=", Donorid) & " " & _
                                "AND   " & DBW("donoraccdt=", donoraccdt)

On Error GoTo Set_BldEnter_error
    DBConn.BeginTrans
    DBConn.Execute strSQL(0)
    DBConn.Execute strSQL(1)
    Call SetDonorStatus(Donorid, donoraccdt, DonorStatus.stsAskVerify, False)
    DBConn.CommitTrans
    
    SetPheresisCancel = True
    Exit Function
    
Set_BldEnter_error:
    DBConn.RollbackTrans
    SetPheresisCancel = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function SetBldCancel(ByVal Donorid As String, ByVal donoraccdt As String, ByVal TmpId As String, _
                             ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                             ByVal CompoCd As String) As Boolean
                             
    Dim strSQL    As String
    Dim Rs        As Recordset
    Dim sSql      As String
On Error GoTo Set_BldEnter_error
    DBConn.BeginTrans
   
    strSQL = " delete " & T_BBS401 & _
             " WHERE" & _
                  " " & DBW("bldsrc", Bldsrc, 2) & _
             " AND  " & DBW("bldyy", Bldyy, 2) & _
             " AND  " & DBW("bldno", Bldno, 2) & _
             " AND  " & DBW("compocd", CompoCd, 2)
    DBConn.Execute strSQL
    
    sSql = " SELECT * FROM " & T_BBS401 & _
           " WHERE " & _
                  "" & DBW("donoraccdt", donoraccdt, 2) & _
           " AND   " & DBW("donorid", Donorid, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    strSQL = ""
    If Rs.EOF Then
        strSQL = " update " & T_BBS602 & _
                    " set" & _
                       " " & DBW("donationdt", "", 3) & _
                       " " & DBW("bldsrc", "", 3) & _
                       " " & DBW("bldyy", "", 3) & _
                       " " & DBW("bldno", 0, 3) & _
                       " " & DBW("compocd", "", 3) & _
                       " " & DBW("volumn", 0, 3) & _
                       " " & DBW("entfg", "0", 3) & _
                       " " & DBW("cancelfg", "1", 2) & _
                    " WHERE" & _
                          "" & DBW("donoraccdt", donoraccdt, 2) & _
                    " AND  " & DBW("donorid", Donorid, 2)
        
    End If
    
    If strSQL <> "" Then DBConn.Execute strSQL

'    strSql = " update " & T_LAB102 & _
'                " set    " & DBW("dcfg", "1", 2) & _
'                " WHERE" & _
'                     " " & DBW("ptid", TmpId, 2) & _
'                " AND  orddt in (SELECT distinct orddt " & _
'                                "FROM " & T_BBS605 & " " & _
'                                "WHERE " & DBW("donorid=", Donorid) & " " & _
'                                "AND   " & DBW("donoraccdt=", donoraccdt) & ")"
    '예수병원용
    strSQL = " update mdbldort " & _
                " set    " & DBW("dcyn", "Y", 2) & _
                " WHERE" & _
                     " " & DBW("patno", TmpId, 2) & _
                " AND  orddate in (SELECT distinct to_date(orddt,'yyyymmdd') " & _
                                "FROM " & T_BBS605 & " " & _
                                "WHERE " & DBW("donorid=", Donorid) & " " & _
                                "AND   " & DBW("donoraccdt=", donoraccdt) & ")"
    DBConn.Execute strSQL
    

    Call SetDonorStatus(Donorid, donoraccdt, DonorStatus.stsAskVerify, False)
    
    DBConn.CommitTrans
    
    SetBldCancel = True
    Set Rs = Nothing
    Exit Function
    
Set_BldEnter_error:
    DBConn.RollbackTrans
    SetBldCancel = False
    Set Rs = Nothing
    MsgBox Err.Description, vbExclamation
End Function
Public Function Compolist(ByVal TF As Boolean) As Recordset
    Dim sSql As String
    
    sSql = " SELECT compocd,abbrnm,keepday FROM " & T_BBS006
    If TF = True Then
        sSql = sSql & " WHERE pherefg='1'"
    Else
        sSql = sSql & " WHERE pherefg='0'"
    End If
    
    Set Compolist = New Recordset
    Compolist.Open sSql, DBConn

End Function


Public Function Get_OrdInformation(ByVal AccNo As String) As Recordset
'접수번호를 가지고처방정보를 가지고 온다.(헌혈등록에서 사용)
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_LAB102 & _
           " WHERE " & _
                     "     " & DBW("workarea", C_WORKAREA, 2) & _
                     " AND " & DBW("accdt", medGetP(AccNo, 1, "-"), 2) & _
                     " AND " & DBW("accseq", medGetP(AccNo, 2, "-"), 2)
    
    Set Get_OrdInformation = New Recordset
    Get_OrdInformation.Open sSql, DBConn
End Function


Public Function GetLocalHp(Optional ByVal pLocalCd As String = "") As String
'FROM COM003
'로컬병원
    GetLocalHp = " SELECT cdval1,field1 FROM " & T_COM003 & _
                 " WHERE" & _
                          "     " & DBW("cdindex", BC2_LOCAL, 2)
                          
    If pLocalCd <> "" Then
        GetLocalHp = GetLocalHp & " AND " & DBW("cdval1", pLocalCd, 2)
    End If
End Function

Public Function GetCenterNm(Optional ByVal pCenterCd As String = "") As String
'FROM COM003
'Center Cd를 얻어온다
    GetCenterNm = " SELECT cdval1,field1 FROM " & T_COM003 & _
                  " WHERE" & _
                           "     " & DBW("cdindex", BC2_CENTER, 2)
                           
    If pCenterCd <> "" Then
        GetCenterNm = GetCenterNm & " AND " & DBW("cdval1", pCenterCd, 2)
    End If
End Function

Public Function GetTestReq4IDRange() As String
'FROM COM003 (Cdindex : BC2_TMP_ID)
'검사의뢰용 임시환자 ID의 범위를 읽어 온다.
   GetTestReq4IDRange = " SELECT field1, field2 FROM " & T_COM003 & _
                        " WHERE" & _
                                 "     " & DBW("cdindex", BC2_TMP_ID, 2) & _
                                 " AND " & DBW("cdval1", "0", 2)
End Function

Public Function GetBldDeliveryByLocalCd(ByVal LocalCd As String, ByVal DeliveryFrom As String, _
                                        ByVal deliveryTo As String) As String
'FROM BBS402
'혈액 출고 내역
'LocalCd값으로 Local에 해당되는 정보를 얻는다.
    GetBldDeliveryByLocalCd = " SELECT * FROM " & T_BBS402 & _
                              " WHERE" & _
                                       "     " & DBW("localcd", LocalCd, 2) & _
                                       " AND " & DBW("deliverydt >", DeliveryFrom, 2) & _
                                       " AND " & DBW("deliverydt <", deliveryTo, 2)

End Function

Public Function GetBldDeliveryMaxSeq(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                     ByVal Bldno As Long, ByVal CompoCd As String) As String
'FROM BBS402
'혈액 출고 내역
'출고 Seq를 얻는다.

    GetBldDeliveryMaxSeq = " SELECT max(deliveryseq) as maxseq" & _
                           " FROM " & T_BBS402 & _
                           " WHERE" & _
                                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                                    " AND " & DBW("bldyy", Bldyy, 2) & _
                                    " AND " & DBW("bldno", Bldno, 2) & _
                                    " AND " & DBW("compocd", CompoCd, 2)

End Function

Public Function GetBldStorageBldNo(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                    ByVal Bldno As Long) As String
'FROM BBS401
'혈액 입고 내역
'혈액번호 존재여부 체크에 사용
    
    GetBldStorageBldNo = " SELECT * FROM " & T_BBS401 & _
                         " WHERE" & _
                                  "     " & DBW("bldsrc", Bldsrc, 2) & _
                                  " AND " & DBW("bldyy", Bldyy, 2) & _
                                  " AND " & DBW("bldno", Bldno, 2) & _
                                  " AND (" & DBW("stscd", BBSBloodStatus.stsENTER, 2) & " or " & DBW("stscd", BBSBloodStatus.stsRETURN, 2) & ")"
                                  
                                  '2001-11-29수정 : 반환된것도 사용가능하니까...


End Function

Public Function GetBldStorageDeliveryChk(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                         ByVal Bldno As Long, ByVal CompoCd As String) As String
'FROM BBS401
'혈액 입고 내역
'혈액 출고여부 체크에 사용

    GetBldStorageDeliveryChk = " SELECT * FROM " & T_BBS401 & _
                               " WHERE" & _
                                        "     " & DBW("bldsrc", Bldsrc, 2) & _
                                        " AND " & DBW("bldyy", Bldyy, 2) & _
                                        " AND " & DBW("bldno", Bldno, 2) & _
                                        " AND " & DBW("compocd", CompoCd, 2) & _
                                        " AND " & DBW("stscd", BBSBloodStatus.stsASSIGN, 2)
End Function

                               
Public Function GetCompocd(ByVal Bldsrc As String, ByVal Bldyy As String, _
                           ByVal Bldno As Long) As String
'FROM BBS401
'혈액입고 내역
'제재코드와 제제명을 얻는다.

    GetCompocd = " SELECT a.compocd, b.abbrnm as field1" & _
                 " FROM " & T_BBS401 & " a, " & T_BBS006 & " b " & _
                 " WHERE" & _
                          "     " & DBW("a.bldsrc", Bldsrc, 2) & _
                          " AND " & DBW("a.bldyy", Bldyy, 2) & _
                          " AND " & DBW("a.bldno", Bldno, 2) & _
                          " AND (" & DBW("a.stscd", BBSBloodStatus.stsENTER, 2) & " or " & DBW("a.stscd", BBSBloodStatus.stsRETURN, 2) & ")" & _
                          " AND a.compocd = b.compocd "
                          
                          '2001-11-29 수정 : 반환된 혈액도 대상이므로...

End Function

Public Function GetABORh(ByVal Bldsrc As String, ByVal Bldyy As String, _
                           ByVal Bldno As Long, ByVal CompoCd As String) As String
'FROM BBS401
'혈액입고 내역
'제재코드와 제제명을 얻는다.

    GetABORh = " SELECT abo, rh FROM " & T_BBS401 & _
               " WHERE" & _
                        "     " & DBW("bldsrc", Bldsrc, 2) & _
                        " AND " & DBW("bldyy", Bldyy, 2) & _
                        " AND " & DBW("bldno", Bldno, 2) & _
                        " AND " & DBW("compocd", CompoCd, 2)
                 
                 
End Function

Public Function GetStorageHistory(ByVal Bldsrc As String, ByVal Bldyy As String, _
                                  ByVal Bldno As Long, ByVal CompoCd As String, _
                                  Optional ByVal centercd As String = "") As String
'FROM BBS401
'혈액입고 내역의 모든 정보를 얻는다.

    If centercd = "" Then
        GetStorageHistory = " SELECT * FROM " & T_BBS401 & _
                            " WHERE" & _
                                     "     " & DBW("bldsrc", Bldsrc, 2) & _
                                     " AND " & DBW("bldyy", Bldyy, 2) & _
                                     " AND " & DBW("bldno", Bldno, 2) & _
                                     " AND " & DBW("compocd", CompoCd, 2)
    Else
        GetStorageHistory = " SELECT * FROM " & T_BBS401 & _
                            " WHERE" & _
                                     "     " & DBW("bldsrc", Bldsrc, 2) & _
                                     " AND " & DBW("bldyy", Bldyy, 2) & _
                                     " AND " & DBW("bldno", Bldno, 2) & _
                                     " AND " & DBW("compocd", CompoCd, 2) & _
                                     " AND " & DBW("centercd", centercd, 2)
                     
    End If
End Function

Public Function GetCompCdForCboBox(Optional ByVal CompoCd As String = "") As String
'FROM COM003
'Component 코드, 코드명
    If CompoCd = "" Then
        GetCompCdForCboBox = "SELECT compocd as cdval1, abbrnm as field1 FROM " & T_BBS006
    Else
        GetCompCdForCboBox = " SELECT compocd as cdval1, abbrnm as field1 FROM " & T_BBS006 & _
                             " WHERE " & DBW("compocd", CompoCd, 2)
    End If
End Function

Public Function GetJobCd(Optional JobNm As String = "") As String
'FROM COM003
'CdIndex : B016 (BC2_JOB)
'직업코드를 얻어온다.
    
    GetJobCd = " SELECT cdval1, field1 FROM " & T_COM003 & _
               " WHERE  " & DBW("cdindex", BC2_JOB, 2)
               
    If JobNm <> "" Then
        GetJobCd = GetJobCd & _
                            " AND " & DBW("field1", JobNm, 2)
    End If
    
End Function

Public Function GetJobNm(ByVal JobCd As String) As String
'FROM COM003
'CdIndex : B016 (BC2_JOB)
'직업명를 얻어온다.
    
    GetJobNm = " SELECT cdval1, field1 FROM " & T_COM003 & _
               " WHERE" & _
                        "     " & DBW("cdindex", BC2_JOB, 2) & _
                        " AND " & DBW("cdval1", JobCd, 2)
End Function

Public Function GetDonorMst(Optional ByVal Donorid As String = "", Optional ByVal donornm As String = "") As String
'FROM BBS601
'헌혈자 마스터
    
    GetDonorMst = "SELECT * FROM " & T_BBS601
        
    If donornm = "" Then
        GetDonorMst = GetDonorMst & " WHERE  " & DBW("donorid", Donorid, 2)
    Else
'        GetDonorMst = GetDonorMst & " WHERE  " & DBW("donornm like ", donornm & "%")
        GetDonorMst = GetDonorMst & " WHERE  " & DBW("donornm=", donornm)
    End If
End Function

'2001-11-27추가
'헌혈자 혈액번호리스트를 검색한다.
Public Function GetDonorBldList(ByVal Donorid As String, ByVal donordt As String) As String
'FROM BBS401
'혈액입고내역
    
    GetDonorBldList = "SELECT a.bldsrc, a.bldyy, a.bldno, a.compocd,a.volumn,b.abbrnm as componm " & _
                      " FROM   " & T_BBS006 & " b, " & T_BBS401 & " a " & _
                      " WHERE  " & DBW("a.donorid", Donorid, 2) & _
                      " AND    " & DBW("a.donoraccdt", donordt, 2) & _
                      " AND    b.compocd = a.compocd ORDER BY bldsrc,bldyy,bldno "
End Function

Public Function GetDonorMstBySSN(ByVal donornm As String, ByVal ssn As String) As String
'FROM BBS601
'헌혈자 마스터에서 헌혈자명과 주민번호가 같은 헌혈자를 얻어온다.

    GetDonorMstBySSN = "SELECT * FROM " & T_BBS601 & _
                       " WHERE " & _
                               "     " & DBW("donornm", donornm, 2) & _
                               " AND " & DBW("ssn", ssn, 2)
End Function

Public Function GetDonorCnt(ByVal donornm As String) As String
'FROM BBS601
'헌혈자 마스터에서 동명의 헌혈자의 수를 구한다.

    GetDonorCnt = "SELECT count(donornm) as cnt FROM " & T_BBS601 & _
                  " WHERE  " & DBW("donornm", donornm, 2)
End Function

'Public Function SetPtInfomation(ByVal PtId As String, ByVal PtNm As String, ByVal ssn As String) As String
''환자마스터에 헌혈자 임시아이의 정보를 저장한다.
'    SetPtInfomation = " insert into " & T_HIS001 & " (" & F_PTID & "," & F_PTNM & "," & F_SSN & ")" & _
'                    " values(" & _
'                    DBV(F_PTID, PtId, 1) & DBV(F_PTNM, PtNm, 1) & DBV(F_SSN, ssn) & ")"
'End Function
Public Function SetNoGiveInfo(ByVal UpdateFg As Boolean, ByVal CDINDEX As String, _
                              ByVal seq As Long) As String
'To COM099
'NoIndex : B006 (BN_DONOR_ID, 헌혈자 ID), B001(BN_TMP_ID, 임상병리 검사용 임시 환자 ID)
'번호부여 정보
    
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
                                     DBV("divcd2", "0", 1) & DBV("divcd3", "0", 1) & DBV("seq", seq) & ")"
    End If
End Function

Public Function GetNoGiveMaxSeq(ByVal CDINDEX As String) As String
'FROM COM099
'번호부여 정보테이블에서 최고값을 얻어온다.

    GetNoGiveMaxSeq = "SELECT max(seq) as maxseq FROM " & T_COM099 & _
                      " WHERE " & DBW("noindex", CDINDEX, 2) & _
                      "   AND " & DBW("divcd1", C_WORKAREA, 2)
End Function

Public Function GetNoGiveInfo(ByVal NoIndex As String) As String
'FROM COM099
'NoIndex : B006 (BN_DONOR_ID, 헌혈자 ID), B001(BN_TMP_ID, 임상병리 검사용 임시 환자 ID)
'번호부여 정보
    GetNoGiveInfo = "SELECT * FROM " & T_COM099 & _
                    " WHERE " & _
                                        DBW("noindex", NoIndex, 2) & _
                              " AND " & DBW("divcd1", C_WORKAREA, 2) & _
                              " AND " & DBW("divcd2", "0", 2) & _
                              " AND " & DBW("divcd3", "0", 2)
End Function

Public Function GetEmpInfo(ByVal EmpID As String) As String
'FROM COM006
'EmpID, EmpNm을 얻어온다.
    
    GetEmpInfo = "SELECT empid,empnm FROM " & T_COM006 & _
                 " WHERE " & DBW("empid", EmpID, 2)
                 
End Function

Public Function GetDonnorTest(TmpId As String) As Recordset
'임시환자번호와 처방일을 가지고
'검사의뢰 정보를 가지고 온다.
'검사코드와 검체코드정보를 가지고 온다.
'희망 채혈일시는 처방헤더에서 가지고 오고, 검사코드와 검체코드는 처방 바디에서 가지고 온다.
    Dim sSql As String
    
    sSql = " SELECT a.reqdt,a.reqtm,b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.spccd,b.statfg" & _
            " FROM " & T_LAB101 & " a," & T_LAB102 & " b" & _
            " WHERE" & _
                     "       " & DBW("a.ptid", TmpId, 2) & _
                     "   AND (b.dcfg is null or " & DBW("b.dcfg<>", "1") & ")" & _
                     "   AND a.ptid=b.ptid" & _
                     "   AND a.orddt=b.orddt" & _
                     "   AND a.ordno=b.ordno"
            
    Set GetDonnorTest = New Recordset
    GetDonnorTest.Open sSql, DBConn
    
End Function

Public Function GetTestSpc() As Recordset
'검사항목 마스터에서 default 검사항목을 불러온다.
    Dim sSql As String
    sSql = " SELECT cdval1,field1 FROM " & T_COM003 & _
           " WHERE " & DBW("cdindex", BC2_DS_TEST, 2)
           
    Set GetTestSpc = New Recordset
    GetTestSpc.Open sSql, DBConn
End Function

Public Function GetTestSpc2(ByVal pGroup As String) As Recordset
'검사항목 마스터에서 default 검사항목을 불러온다.
    Dim sSql As String
    sSql = " SELECT cdval2,field1 FROM " & T_COM002 & _
           " WHERE " & DBW("cdindex=", BC2_DS_TEST) & _
           " AND   " & DBW("cdval1=", Trim(pGroup))
           
    Set GetTestSpc2 = New Recordset
    GetTestSpc2.Open sSql, DBConn
End Function

Public Function GetDonorTestDt(ByVal Donorid As String, ByVal accdt As String) As Recordset
'검사의뢰 날짜를 가지고온다
'검사의뢰 날짜가 결국은 처방의 처방일이다.
'검사접수내역(BBS602)에서 임시 환자ID도 불러온다.? 처방에서 검사항목등을 불러오기 위해서

'FROM BBS602 AND BBS605
    Dim sSql As String
    
    sSql = " SELECT distinct tmpid " & _
            " FROM " & T_BBS602 & _
            " WHERE " & DBW("donorid", Donorid, 2) & _
            " AND   " & DBW("donoraccdt", accdt, 2) & " AND cancelfg='0' "
    Set GetDonorTestDt = New Recordset
    GetDonorTestDt.Open sSql, DBConn
End Function
Public Function GetDefaultTestList(tmpTestcd As String) As String
'FROM LAB001(검사항목마스터), LAB004(지정 검체 마스터), COM003(검체 마스터)
'검사항목 마스터에 등록된 검사항목을 조회한다.
'AND (" & DBW("a.panelfg", "", 2) & " or a.panelfg is null)
GetDefaultTestList = " SELECT a.testnm, a.abbrnm5, a.testcd, c.cdval1 as spccd, b.statfg, a.workarea," & _
                     " b.storecd, b.rndfg,b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg," & _
                     " c.field2 as SpcGrp,c.field5 as spcnm,c.field2 as labdiv,c.field2 as labrange,'1' as insurfg" & _
                     " FROM " & _
                             T_LAB001 & " a, " & T_LAB004 & " b, " & T_LAB032 & " c " & _
                     " WHERE " & _
                             DBW("a.testcd", medGetP(tmpTestcd, 1, vbTab), 2) & _
                             " AND a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd )" & _
                             " AND (a.detailfg is null or a.detailfg='') " & _
                             " AND a.testcd = b.testcd" & _
                             " AND " & DBW("b.spccd", medGetP(tmpTestcd, 2, vbTab), 2) & _
                             " /* AND " & DBW("b.seq", "1", 2) & " */ " & _
                             " AND (" & DBW("b.expdt", "", 2) & " or b.expdt is null) " & _
                             " AND " & DBW("c.cdindex", CD2_SPECIMEN, 2) & _
                             " AND " & DBW("c.cdval1", medGetP(tmpTestcd, 2, vbTab), 2)

End Function


Public Function GetDonorAccdtHistoryDivPheresis(ByVal Donorid As String, Optional ByVal accdt As String = "", Optional ByVal phere As Boolean = False) As Recordset
    '성분헌혈과 그이외의 헌혈을 구별해서 조회하기 위해서
    '2001/10/04 울산 동강병원에서 수정
    Dim sSql As String
    sSql = "SELECT * FROM " & T_BBS602 & _
           " WHERE  " & DBW("donorid", Donorid, 2)
    If phere = True Then
        sSql = sSql & " AND donorcd='3'"
    Else
        sSql = sSql & " AND donorcd<>'3'"
    End If
    If accdt = "" Then
        sSql = sSql & " ORDER BY donoraccdt desc "
    Else
        sSql = sSql & " AND " & DBW("donoraccdt", accdt, 2)
    End If
    
    Set GetDonorAccdtHistoryDivPheresis = New Recordset
    GetDonorAccdtHistoryDivPheresis.Open sSql, DBConn
End Function

Public Function GetDonorAccHistory(ByVal Donorid As String, Optional ByVal accdt As String = "") As Recordset
'FROM BBS602
'헌혈자 접수 내역
    Dim sSql As String
    
    sSql = "SELECT * FROM " & T_BBS602 & _
                         " WHERE  " & DBW("donorid", Donorid, 2)
    If accdt = "" Then
        sSql = sSql & " ORDER BY donoraccdt desc "
    Else
        sSql = sSql & " AND " & DBW("donoraccdt", accdt, 2)
    End If
    Set GetDonorAccHistory = New Recordset
    GetDonorAccHistory.Open sSql, DBConn
    
End Function

Public Function SetDonorStatus(ByVal Donorid As String, ByVal accdt As String, ByVal stscd As String, _
                               Optional ByVal IsTransaction As Boolean = True) As Boolean
    Dim sSql As String
    
    sSql = " update " & T_BBS603 & " " & _
           " set    " & DBW("stscd", stscd, 2) & _
           " WHERE " & DBW("donorid", Donorid, 2) & _
           " AND   " & DBW("donoraccdt", accdt, 2)

On Error GoTo SetDonorStatus_error
    If IsTransaction Then DBConn.BeginTrans
    DBConn.Execute sSql
    If IsTransaction Then DBConn.CommitTrans
    
    SetDonorStatus = True
    Exit Function
    
SetDonorStatus_error:
    If IsTransaction Then DBConn.RollbackTrans
    
    SetDonorStatus = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetDonorResult(ByVal Donorid As String, ByVal accdt As String) As Recordset
    Dim sSql As String
    
    sSql = "SELECT * " & _
           "FROM " & T_BBS603 & " " & _
           "WHERE " & DBW("donorid", Donorid, 2) & _
           "AND " & DBW("donoraccdt", accdt, 2)
    
    Set GetDonorResult = New Recordset
    Call GetDonorResult.Open(sSql, DBConn)
'    If GetDonorResult.DBerror Then
'        dbconn.DisplayErrors
'        Set GetDonorResult = Nothing
'    End If
End Function

Public Function GetDonorStatus(ByVal Donorid As String, ByVal accdt As String, ByVal IsPhere As Boolean) As String
    Dim sSql As String
    Dim DrRS As Recordset
    
    Dim stscd       As String
    Dim stsnm       As String
    Dim okdiv1cd    As String
    Dim okdiv1nm    As String
    Dim okdiv2cd    As String
    Dim okdiv2nm    As String
    Dim okdiv3cd    As String
    Dim okdiv3nm    As String
    
    
    sSql = "SELECT stscd,okdiv1,okdiv2,okdiv3 " & _
           "FROM " & T_BBS603 & " " & _
           "WHERE " & DBW("donorid", Donorid, 2) & _
           "AND " & DBW("donoraccdt", accdt, 2)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        GetDonorStatus = ""
    Else
        With DrRS
            If .RecordCount < 1 Then
                GetDonorStatus = ""
            Else
                stscd = .Fields("stscd").Value & ""
                Select Case stscd
                    Case DonorStatus.stsAccessSave
                        stsnm = "접수등록"
                    Case DonorStatus.stsAccessVerify
                        stsnm = "접수확인"
                        okdiv1cd = .Fields("okdiv1").Value & ""
                        okdiv1nm = IIf(okdiv1cd = "1", "적격", IIf(okdiv1cd = "0", "부적격", ""))
                    Case DonorStatus.stsAskSave
                        stsnm = "문진등록"
                        okdiv1cd = .Fields("okdiv1").Value & ""
                        okdiv1nm = IIf(okdiv1cd = "1", "적격", IIf(okdiv1cd = "0", "부적격", ""))
                    Case DonorStatus.stsAskVerify
                        stsnm = "문진확인"
                        okdiv1cd = .Fields("okdiv1").Value & ""
                        okdiv1nm = IIf(okdiv1cd = "1", "적격", IIf(okdiv1cd = "0", "부적격", ""))
                        okdiv2cd = .Fields("okdiv2").Value & ""
                        okdiv2nm = IIf(okdiv2cd = "1", "적격", IIf(okdiv2cd = "0", "부적격", ""))
                    Case DonorStatus.stsDonation
                        stsnm = "검사중"
                        okdiv1cd = .Fields("okdiv1").Value & ""
                        okdiv1nm = IIf(okdiv1cd = "1", "적격", IIf(okdiv1cd = "0", "부적격", ""))
                        okdiv2cd = .Fields("okdiv2").Value & ""
                        okdiv2nm = IIf(okdiv2cd = "1", "적격", IIf(okdiv2cd = "0", "부적격", ""))
                    Case DonorStatus.stsFinish
                        stsnm = "판정완료"
                        okdiv1cd = .Fields("okdiv1").Value & ""
                        okdiv1nm = IIf(okdiv1cd = "1", "적격", IIf(okdiv1cd = "0", "부적격", ""))
                        okdiv2cd = .Fields("okdiv2").Value & ""
                        okdiv2nm = IIf(okdiv2cd = "1", "적격", IIf(okdiv2cd = "0", "부적격", ""))
                        okdiv3cd = .Fields("okdiv3").Value & ""
                        okdiv3nm = IIf(okdiv3cd = "1", "적격", IIf(okdiv3cd = "0", "부적격", ""))
                    Case DonorStatus.stsPrint
                        stsnm = "DM 발송"
                        okdiv1cd = .Fields("okdiv1").Value & ""
                        okdiv1nm = IIf(okdiv1cd = "1", "적격", IIf(okdiv1cd = "0", "부적격", ""))
                        okdiv2cd = .Fields("okdiv2").Value & ""
                        okdiv2nm = IIf(okdiv2cd = "1", "적격", IIf(okdiv2cd = "0", "부적격", ""))
                        okdiv3cd = .Fields("okdiv3").Value & ""
                        okdiv3nm = IIf(okdiv3cd = "1", "적격", IIf(okdiv3cd = "0", "부적격", ""))
                End Select
                GetDonorStatus = stsnm & vbTab & stscd & vbTab & _
                                 okdiv1nm & vbTab & okdiv1cd & vbTab & _
                                 okdiv2nm & vbTab & okdiv2cd & vbTab & _
                                 okdiv3nm & vbTab & okdiv3cd
            End If
        End With
        Set DrRS = Nothing
    End If
End Function

Public Function Get_TestHistory(ByVal Donorid As String, Optional ByVal accdt As String = "") As Recordset
'FROM BBS605
'헌혈자 검사의뢰가 이루어져있는지를 체크한다.
    Dim sSql As String
    
    sSql = " SELECT distinct a.tmpid " & _
            " FROM " & T_BBS602 & " a," & T_BBS605 & " b" & _
            " WHERE" & _
                     "     " & DBW("a.donorid", Donorid, 2) & _
                     " AND " & DBW("a.donoraccdt", accdt, 2) & _
                     " AND (a.cancelfg='0')" & _
                     " AND a.donorid=b.donorid" & _
                     " AND a.donoraccdt=b.donoraccdt"
    Set Get_TestHistory = New Recordset
    Get_TestHistory.Open sSql, DBConn
    
End Function
Public Function Get_Pheresis(ByVal Donorid As String, ByVal accdt As String) As Recordset
'성분헌혈인 경우....
    Dim sSql As String
    
    sSql = " SELECT * " & _
            " FROM " & T_BBS602 & _
            " WHERE" & _
                     "     " & DBW("donorid", Donorid, 2) & _
                     " AND " & DBW("donoraccdt", accdt, 2) & _
                     " AND " & DBW("cancelfg", "0", 2)
    
    Set Get_Pheresis = New Recordset
    Get_Pheresis.Open sSql, DBConn
End Function
Public Function SetDonorAccHistory(ByVal UpdateFg As Boolean, _
                                   ByVal Donorid As String, _
                                   ByVal donoraccdt As String, _
                                   ByVal DonorCd As String, _
                                   ByVal reservedid As String, _
                                   ByVal TmpId As String, _
                                   ByVal Weight As String, _
                                   ByVal Height As String, _
                                   ByVal Pulse As String, _
                                   ByVal BodyTemp As String, _
                                   ByVal BldPres1 As String, _
                                   ByVal BldPres2 As String, _
                                   ByVal DonationDt As String, _
                                   ByVal Bldsrc As String, _
                                   ByVal Bldyy As String, _
                                   ByVal Bldno As Long, _
                                   ByVal CompoCd As String, _
                                   ByVal EntFg As String, _
                                   ByVal okdiv As String, _
                                   ByVal rmk As String, _
                                   ByVal accdt As String, _
                                   ByVal accseq As String, _
                                   ByVal WorkArea As String, _
                                   ByVal IsHold As Boolean) As String
'To BBS602
'헌혈자 접수 내역
    Dim OkDt1 As String
    Dim stscd As String
    
    If IsHold Then
        stscd = DonorStatus.stsAccessSave
    Else
        stscd = DonorStatus.stsAccessVerify
    End If
    
    OkDt1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    If UpdateFg Then
        
    
        SetDonorAccHistory = "update " & T_BBS602 & _
                             " set  " & _
                                      DBW("donorcd", DonorCd, 3) & _
                                      DBW("reservedid ", reservedid, 3) & _
                                      DBW("tmpid", TmpId, 3) & _
                                      DBW("workarea", WorkArea, 3) & _
                                      DBW("accdt   ", accdt, 3) & _
                                      DBW("accseq  ", accseq, 3) & _
                                      DBW("weight  ", Weight, 3) & _
                                      DBW("height  ", Height, 3) & _
                                      DBW("pulse   ", Pulse, 3) & _
                                      DBW("bodytemp", BodyTemp, 3) & _
                                      DBW("bldpres1", BldPres1, 3) & _
                                      DBW("bldpres2", BldPres2, 2) & _
                             " WHERE  " & _
                                        DBW("donorid", Donorid, 2) & _
                             " AND  " & DBW("donoraccdt", donoraccdt, 2)
       
        SetDonorAccHistory = SetDonorAccHistory & ";" & _
                             "update " & T_BBS603 & " " & _
                             "set " & DBW("stscd", stscd, 3) & _
                                      DBW("okdiv1", okdiv, 3) & _
                                      DBW("rmk1", rmk, 3) & _
                                      DBW("okdt1", OkDt1, 2) & _
                             " WHERE " & DBW("donorid", Donorid, 2) & _
                               " AND " & DBW("donoraccdt", donoraccdt, 2)

    Else
        SetDonorAccHistory = "insert into " & T_BBS602 & _
                             " (donorid,donoraccdt,donorcd,workarea,accdt,accseq,reservedid,tmpid,weight,height," & _
                             " pulse,bodytemp,bldpres1,bldpres2,donationdt,bldsrc,bldyy,bldno,compocd,volumn,entfg,cancelfg) values (" & _
                             DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & _
                             DBV("donorcd", DonorCd, 1) & DBV("workarea", WorkArea, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("reservedid", reservedid, 1) & _
                             DBV("tmpid", TmpId, 1) & DBV("weight", Weight, 1) & _
                             DBV("height", Height, 1) & DBV("pulse", Pulse, 1) & DBV("bodytemp", BodyTemp, 1) & DBV("bldpres1", BldPres1, 1) & DBV("bldpres2", BldPres2, 1) & _
                             DBV("donationdt", DonationDt, 1) & DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", CompoCd, 1) & DBV("volumn", 0, 1) & _
                             DBV("entfg", EntFg, 1) & DBV("cancelfg", "0") & " )"
   
                                 
        SetDonorAccHistory = SetDonorAccHistory & ";" & _
                             "insert into " & T_BBS603 & "(donorid,donoraccdt,stscd,okdiv1,okdiv2,okdiv3,okdt1,okdt2,okdt3,rmk1,rmk2,rmk3) " & _
                             "values(" & _
                             DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & DBV("stscd", stscd, 1) & DBV("okdiv1", okdiv, 1) & _
                             DBV("okdiv2", "", 1) & DBV("okdiv3", "", 1) & DBV("okdt1", OkDt1, 1) & _
                             DBV("okdt2", "", 1) & DBV("okdt3", "", 1) & DBV("rmk1", rmk, 1) & DBV("rmk2", "", 1) & DBV("rmk3", "") & ")"
                             
                            
    End If
End Function
Public Function GetTestFindList(tmpTestcd As String) As Recordset
'FROM LAB001(검사항목마스터), LAB004(지정 검체 마스터), COM003(검체 마스터)
    
    Dim sSql As String
    
    sSql = " SELECT a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea,  b.storecd," & _
            " b.rndfg,b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg, c.field2 as SpcGrp" & _
            " FROM " & _
                     T_LAB001 & " a, " & T_LAB004 & " b, " & T_COM003 & " c " & _
            " WHERE" & _
                     "     a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd )" & _
                     " AND " & DBW("c.cdindex", CD2_SPECIMEN, 2) & _
                     " AND " & DBW("c.cdval1", medGetP(tmpTestcd, 2, vbTab), 2) & _
                     " AND " & DBW("b.testcd", medGetP(tmpTestcd, 1, vbTab), 2) & _
                     " AND c.cdval1 = b.spccd AND a.testcd = b.testcd"
   Set GetTestFindList = New Recordset
   GetTestFindList.Open sSql, DBConn

End Function

Public Function GetLastOrdNo(ByVal pPtID As String, ByVal pDate As String) As String
' 가장 마지막 처방번호

      GetLastOrdNo = " SELECT ordno " & _
                     " FROM   " & T_LAB101 & _
                     " WHERE" & _
                              "     " & DBW("ptid =", pPtID) & _
                              " AND " & DBW("orddt=", pDate) & _
                     " ORDER BY ordno desc"
End Function

Public Function GetSpecList(ByVal pTestCd As String) As String
'%   Get Specimen List
'%      - Calling FROM [frm101Order]
'%      - 검체코드,검체명,응급여부,보관구분,복수검체여부,검체군,우선순위

   GetSpecList = " SELECT a.spccd, b.field3 as SpcNm, a.statfg, a.storecd, a.labelcnt, a.statflags, " & _
                 "         b.field1 as MultiFg, b.field2 as SpcGrp, a.seq  " & _
                 " FROM " & _
                          T_LAB004 & " a, " & T_COM003 & " b " & _
                 " WHERE" & _
                          " " & DBW("a.testcd", pTestCd, 2) & _
                          "AND     a.applydt = ( SELECT max(applydt) FROM " & T_LAB004 & " " & _
                          "                             WHERE testcd = a.testcd  " & _
                          "                             AND     spccd = a.spccd ) " & _
                          "AND    (a.expdt is null " & _
                                   "or             " & DBW("a.expdt", "", 2) & ") " & _
                          "AND     " & DBW("b.cdindex", CD2_SPECIMEN, 2) & _
                          "AND     b.cdval1 = a.spccd " & _
                 "ORDER BY a.seq "
End Function

'%   복수검체
'%       - Calling FROM [ clsDonorBusiCollection ] :
Public Function GetMultiSpc(ByVal SpcCd As String) As String
   
   GetMultiSpc = "SELECT cdval2 as SpcSeq, field1 as SpcCd  " & _
                            "FROM " & T_COM002 & " " & _
                            "WHERE " & DBW("cdindex", CD1_MULTISPC, 2) & _
                            "AND   " & DBW("cdval1", SpcCd, 2)

End Function


'% 이미 생성된 복수검체 갯수
'%       - Calling FROM [ clsDonorBusiCollection ] :
Public Function MultiSpcCnt2(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Integer, ByVal OrdSeq As Integer) As String
'FROM 연속 검사내역

   MultiSpcCnt2 = "SELECT * " & _
                            " FROM " & T_LAB203 & " " & _
                            " WHERE " & DBW("ptid", PtId, 2) & _
                            " AND   " & DBW("orddt", orddt, 2) & _
                            " AND   " & DBW("ordno", OrdNo, 2) & _
                            " AND   " & DBW("ordseq", OrdSeq, 2)
End Function

'%  0. 공통코드 마스터 (COM003)
Public Function SqlCommonCode(ByVal TblName As String, ByVal CDINDEX As String, _
                                                Optional ByVal cdval1 As Variant, Optional ByVal CdVal2 As Variant) As String
   Dim tmpStr As String
   
   tmpStr = ""
   If Not IsMissing(cdval1) Then tmpStr = tmpStr & " AND " & DBW("cdval1", cdval1, 2)
   If Not IsMissing(CdVal2) Then tmpStr = tmpStr & " AND " & DBW("cdval2", CdVal2, 2)
   
   SqlCommonCode = "SELECT * " & _
                   "FROM   " & TblName & " " & _
                   "WHERE " & DBW("cdindex", CDINDEX, 2) & tmpStr
End Function

Public Function SetTestRequest(ByVal Donorid As String, ByVal donoraccdt As String, _
                               ByVal orddt As String, ByVal seq As String, _
                               ByVal WorkArea As String, ByVal accdt As String, _
                               ByVal accseq As Long) As String
'To BBS605
'헌혈자 검사 의뢰 내역
    
    SetTestRequest = "insert into " & T_BBS605 & " (" & _
                    "donorid,donoraccdt,orddt,seq,workarea,accdt,accseq) " & _
                    " values ( " & _
                     DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & DBV("orddt", orddt, 1) & _
                     DBV("seq", seq, 1) & DBV("workarea", WorkArea, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq) & " ) "

End Function

'Public Function SetDonorAccHistoryUpdateByTmpID(ByVal Donorid As String, ByVal donoraccdt As String, _
'                                                ByVal TmpId As Long) As String
'
'    MsgBox "더 이상 사용되지 않습니다." & vbNewLine & _
'           "대신 SetDonorAccHistoryUpdateByTmpID2를 사용하십시요", vbInformation, "SetDonorAccHistoryUpdateByTmpID"
''To BBS602
''헌혈자 접수 내역의 임시 환자 ID 업데이트
'
'    SetDonorAccHistoryUpdateByTmpID = "update " & T_BBS602 & " set  " & DBW("tmpid", TmpId, 2) & _
'                                      " WHERE " & DBW("donorid ", Donorid, 2) & _
'                                      " AND   " & DBW("donoraccdt ", donoraccdt, 2)
'End Function
                                                
Public Function SetDonorAccHistoryUpdateByTmpID(ByVal Donorid As String, ByVal donoraccdt As String, _
                                                 ByVal TmpId As String) As String
'To BBS602
'헌혈자 접수 내역의 임시 환자 ID 업데이트

    SetDonorAccHistoryUpdateByTmpID = "update " & T_BBS602 & " set  " & DBW("tmpid", TmpId, 2) & _
                                      " WHERE " & DBW("donorid ", Donorid, 2) & _
                                      " AND   " & DBW("donoraccdt ", donoraccdt, 2)
End Function
                                                
Public Function GetTestRequestMaxSeq(ByVal Donorid As String, ByVal donoraccdt As String, _
                                     ByVal orddt As String) As String
'FROM BBS605
'헌혈자 검사의뢰 내역으로부터
'검사의뢰번호의 최고값을 얻어온다.
        
    GetTestRequestMaxSeq = "SELECT max(seq) as maxseq FROM " & T_BBS605 & _
                           " WHERE " & DBW("donorid", Donorid, 2) & _
                           " AND   " & DBW("donoraccdt ", donoraccdt, 2) & _
                           " AND   " & DBW("orddt", orddt, 2)
End Function
Public Function GetBloodCheck(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Boolean
'혈액번호가 입고내역에 존재하는 지 여부 체크
    Dim sSql As String
    Dim Rs   As Recordset
    
    sSql = " SELECT * FROM " & T_BBS401 & _
         " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
         " AND " & DBW("bldyy", Bldyy, 2) & _
         " AND " & DBW("bldno", Bldno, 2) & _
         " AND " & DBW("compocd", CompoCd, 2) ' & _
         " AND stscd in('0','1')"
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Rs.EOF Then
        GetBloodCheck = True
    Else
        GetBloodCheck = False
    End If
    Set Rs = Nothing
End Function

Public Function GetBloodInfo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT * " & _
           " FROM " & T_BBS401 & " " & _
           " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
           " AND " & DBW("bldyy", Bldyy, 2) & _
           " AND " & DBW("bldno", Bldno, 2)
           
    Set GetBloodInfo = New Recordset
    Call GetBloodInfo.Open(sSql, DBConn)
End Function

Public Function Get_PheresisInfo(ByVal accdt As String, ByVal accseq As String) As Recordset
'성분헌혈일경우 접수번호에 해당하는 환자의 정보를 가지고 온다.
    Dim sSql As String
    
    sSql = " SELECT a.deptcd,a.wardid," & _
             " b.ordcd,b.ptid," & _
             " c.testnm,d." & F_PTNM & " as ptnm," & F_SSN2("d") & " as ssn,b.mesg" & _
             " FROM " & T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c," & T_HIS001 & " d" & _
             " WHERE " & DBW("b.accdt", accdt, 2) & _
             " AND " & DBW("b.accseq", accseq, 2) & _
             " AND a.orddt=b.orddt" & _
             " AND a.ordno=b.ordno " & _
             " AND a.ptid=b.ptid" & _
             " AND a.ptid=d." & F_PTID & _
             " AND b.ordcd=c.testcd" & _
             " AND c.applydt=(SELECT max(d.applydt)" & _
             " FROM " & T_BBS001 & " d" & _
             " WHERE c.testcd = d.testcd" & _
             " AND   " & DBW("d.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & ")"
              
    
    Set Get_PheresisInfo = New Recordset
    Get_PheresisInfo.Open sSql, DBConn
End Function

Public Function GetPtntNm(ByVal pPtID As String) As String
    Dim Rs As New Recordset
    Dim strSQL As String
    
    strSQL = " SELECT " & F_PTID & " as ptid," & F_PTNM & " as ptnm " & _
             " FROM " & T_HIS001 & _
             " WHERE " & DBW(F_PTID, pPtID, 2)
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF Then
        GetPtntNm = ""
    Else
        GetPtntNm = Rs.Fields("ptnm").Value & ""
    End If
    Set Rs = Nothing
End Function

Public Function GetSpecmenInfomation(ByVal PtId As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT * " & _
           " FROM " & T_BBS201 & _
           " WHERE " & DBW("ptid", PtId, 2) & _
           " ORDER BY coldt desc,coltm desc "
    Set GetSpecmenInfomation = New Recordset
    Call GetSpecmenInfomation.Open(sSql, DBConn)
'    If GetSpecmenInfomation.DBerror Then
'        dbconn.DisplayErrors
'        Set GetSpecmenInfomation = Nothing
'    End If
End Function
Public Function SetDonorAcc(ByVal Donorid As String, ByVal donoraccdta As String) As String
    SetDonorAcc = "update " & T_BBS602 & " set " & DBW("cancelfg", "0", 2) & _
                " WHERE " & DBW("donorid=", Donorid) & " AND " & DBW("donoraccdt=", donoraccdta)
End Function

'병원별 환자마스터를 작성해준다..
'병원별로 바뀔수 있다...
Public Function GetPtMasterInsertSQL(ByVal qPtid As String, ByVal qPtnm As String, ByVal qSex As String, ByVal qDob As String, _
                                     ByVal qSSN As String) As String
    
    Dim ConvToDate As String
    
'    ConvToDate = "To_Date('" & qDob & "', 'YYYY-MM-DD') "

'    GetPtMasterInsertSQL = " insert into " & T_HIS001 & " (bunho,suname,sex,birth,sujumin1,sujumin2) " & _
'                           " values('" & qPtid & "','" & qPtnm & "','" & qSex & "', '197602','" & _
'                                        Mid(qSSN, 1, 6) & "','" & Mid(qSSN, 7) & "'" & _
'                                  ")"
'    GetPtMasterInsertSQL = " insert into " & T_HIS001 & " (patno,patname,sex,birtdate,resno1,resno2) " & _
'                           " values('" & qPtid & "','" & qPtnm & "','" & qSex & "', '" & ConvToDate & "','" & _
'                                        Mid(qSSN, 1, 6) & "','" & Mid(qSSN, 7) & "'" & _
'                                  ")"
    GetPtMasterInsertSQL = " insert into " & T_BBS009 & " (ptid,ptnm,sex,dob,ssn) " & _
                           " values('" & qPtid & "','" & qPtnm & "','" & qSex & "', '" & qDob & "','" & qSSN & "'" & ")"
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
