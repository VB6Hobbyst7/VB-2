VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetSqlStatement"
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

Private mvarBldDic As New clsDictionary

Public Property Set BldDic(ByVal vData As clsDictionary)
    Set mvarBldDic = vData
End Property


Public Property Get BldDic() As clsDictionary
    Set BldDic = mvarBldDic
End Property


Public Function Get_TransOrderList() As Recordset
'-------------------
'수혈처방을 불러온다
'-------------------
    Dim sSql As String
    Dim dt   As String
    
    dt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    sSql = " SELECT testcd,testnm,testdiv " & _
           " FROM  " & T_BBS001 & " a" & _
           " WHERE a.applydt = ( SELECT max(b.applydt) " & _
                               " FROM " & T_BBS001 & " b " & _
                               " WHERE a.testcd = b.testcd " & _
                               " AND   " & DBW("b.applydt<", dt, 2) & ")" & _
           " AND (expdt is null or expdt=' ' or expdt > '" & dt & "' )" & _
           " ORDER BY testdiv,testcd "
    
    Set Get_TransOrderList = New Recordset
    Get_TransOrderList.Open sSql, DBConn
    
    
End Function

Public Function GEtReasonList() As Recordset
   Set GEtReasonList = GetCom003RecordSet(BC2_Trans_Reason)
End Function


Public Function Compolist(ByRef Compo As Object)
'----------------------------
'혈액제제를 ComboBox에 담는다
'----------------------------
    Dim Rs  As Recordset
    
    Set Rs = Get_CompoRecordSet
    
    With Rs
        Compo.Clear
        While .EOF = False
            Compo.AddItem .Fields("compocd").Value & "" & " " & _
                          .Fields("abbrnm").Value & "" & Space(80) & COL_DIV & .Fields("keepday").Value & ""
            .MoveNext
        Wend
    End With
    Set Rs = Nothing
End Function

Public Function GetOPCodeList() As Recordset
'    Dim objSql As clsHosComSQLStmt
    Dim sSql As String
    
'    Set objSql = New clsHosComSQLStmt
'    sSql = objSql.LoadOCD
'    Set objSql = Nothing
    
    Set GetOPCodeList = New Recordset
    GetOPCodeList.Open GetSQLOPList, DBConn
End Function

Public Sub OcdList(ByRef pSpcList As Object)
'-------------------
'수술코드를 불러온다
'-------------------
    Dim Rs As New Recordset
    Dim sSql As String
    Dim tmpStr As String
    Dim strOcd As String
    Dim strONm As String
   
   
'    sSql = " SELECT " & F_OCD & " as ocd," & F_ONM & " as onm" & _
'           " FROM  " & T_HIS007 & _
'           " WHERE " & _
'                     DBW(F_ODIV, "21", 2)
    sSql = GetSQLOPList
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    With Rs
        pSpcList.Clear
        pSpcList.AddItem "(없음)"
        While .EOF = False
            tmpStr = ""
            strONm = tmpStr & Trim(.Fields("onm").Value & "") & Chr(9)       '수술명
            strONm = strONm & Space(40 - Len(strONm))
            strOcd = tmpStr & Trim(.Fields("ocd").Value & "") & Chr(9)       '수술코드
            strOcd = strOcd & Space(8 - Len(strOcd))
            
            pSpcList.AddItem strONm & strOcd & tmpStr                   '
            .MoveNext
        Wend
    End With
   
    Set Rs = Nothing

End Sub

Public Function Get_AssignPt(ByVal PtId As String) As String
'------------------------------------------------------
'Assign취소(frmBBS202)에서 취소 대상이 되는 환자의 조회
'조건: BBS302의 Cancelfg='0'
'------------------------------------------------------
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT distinct a.ptid,c." & F_PTNM & " as ptnm" & _
           " FROM " & _
                    T_BBS202 & " a," & T_BBS302 & " b," & T_HIS001 & " c" & _
           " WHERE" & _
                    "     " & DBW("a.ptid", PtId, 2) & _
                    " AND a.ptid=c." & F_PTID & _
                    " AND a.workarea=b.workarea" & _
                    " AND a.accdt=b.accdt" & _
                    " AND a.accseq=b.accseq" & _
                    " AND " & DBW("b.cancelfg", BBSCancelStatus.stsNotCancel, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        Get_AssignPt = Rs.Fields("ptid").Value & "" & COL_DIV & Rs.Fields("ptnm").Value & ""
    Else
        Get_AssignPt = ""
    End If
    
    Set Rs = Nothing
End Function

Public Function Get_OrderStatusPt(ByVal PtId As String, ByVal qFromDt As String, ByVal qToDt As String) As String
'-------------------------------------------------------
'처방상태조회(frmBBS206)에서 조회대상이 되는 환자의 조회
'LAB101에 있는 환자..WORKAREA=B
'-------------------------------------------------------
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT distinct a.ptid ,b." & F_PTNM & " as ptnm " & _
           " FROM " & _
                    T_LAB101 & " a," & T_HIS001 & " b" & _
           " WHERE" & _
                    "     " & DBW("a.ptid", PtId, 2) & _
                    " AND " & DBW("a.reqdt>=", qFromDt) & _
                    " AND " & DBW("a.reqdt<=", qToDt) & _
                    " AND a.ptid=b." & F_PTID & _
                    " AND " & DBW("a.orddiv", C_WORKAREA, 2)
             
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        Get_OrderStatusPt = Rs.Fields("ptid").Value & "" & COL_DIV & Rs.Fields("ptnm").Value & ""
    Else
        Get_OrderStatusPt = ""
    End If
    Set Rs = Nothing
End Function

'-----------------------
'# 혈액조회    (BBS307)
'-----------------------
Public Function Get_CenterRecordSet() As Recordset
'-------------------------
'센터를 RecordSet으로 반환
'-------------------------
    Dim sSql As String
    
    sSql = " SELECT cdval1,field1" & _
           " FROM " & T_COM003 & _
           " WHERE" & _
                    " " & DBW("cdindex", BC2_CENTER, 2)
    
    Set Get_CenterRecordSet = New Recordset
    Get_CenterRecordSet.Open sSql, DBConn
End Function
Public Function Get_BldEnter(ByVal centercd As String) As Recordset
'-----------------------------------------------------
'입고내역에서 현재 남아있는 혈액의 수량을 가지고 온다.
'stscd='0','1','2'
'-----------------------------------------------------
Dim sSql As String
    
    sSql = "SELECT  compocd,abo,rh,count(rh) as cnt" & _
             " FROM " & T_BBS401 & _
             " WHERE " & _
             "  stscd in(" & _
                           DBV("stscd", BBSBloodStatus.stsENTER, 1) & _
                           DBV("stscd", BBSBloodStatus.stsRETURN, 1) & _
                           DBV("stscd", BBSBloodStatus.stsASSIGN) & ")"
                 
    If centercd <> "" Then
        sSql = sSql & " AND " & DBW("centercd", centercd, 2)
    End If
    sSql = sSql & " GROUP BY compocd,abo,rh"
    
    Set Get_BldEnter = New Recordset
    Get_BldEnter.Open sSql, DBConn
End Function
Public Function Get_AssignCnt(ByVal centercd As String) As Recordset
'-----------------------------------------
'입고내역에서 Assign 된 수량을 가지고 온다
'stscd='2'
'-----------------------------------------
    Dim sSql As String
    
    
    sSql = "SELECT compocd,abo,rh,count(rh) as cnt " & _
          " FROM " & T_BBS401 & _
          " WHERE" & _
                   " " & DBW("stscd", BBSBloodStatus.stsASSIGN, 2)
             
    If centercd <> "" Then
        sSql = sSql & _
                    " AND " & DBW("centercd", centercd, 2)
    End If
    
    sSql = sSql & " GROUP BY compocd,abo,rh ORDER BY compocd"
         
    Set Get_AssignCnt = New Recordset
    Get_AssignCnt.Open sSql, DBConn
End Function

Public Function Get_CompoRecordSet() As Recordset
'------------------------------
'bbs006에서 혈액제제를 조회한다
'------------------------------
    Dim sSql As String
    
    sSql = " SELECT compocd,abbrnm,keepday " & _
           " FROM " & T_BBS006 & _
           " WHERE" & _
                    " expdt='' or expdt is null ORDER BY compocd"
    
    Set Get_CompoRecordSet = New Recordset
    Get_CompoRecordSet.Open sSql, DBConn

End Function
Public Function Get_Split_TarGet(ByVal BloodNum As String, ByVal centercd As String, ByVal TF As Boolean) As Recordset
'-------------------------------------------------------------
'Get_Split_TarGet:  분획대상여부를 판단해서 반환
'TF:true 일때 해당 센터의 혈액을 조회한다.(제제)
'TF=false 일때 해당센터의 혈액을 제외한 혈액을 조회한다.(제제)
'-------------------------------------------------------------
    Dim sSql   As String
    Dim Bldsrc As String
    Dim Bldyy  As String
    Dim Bldno  As Long
    
    Bldsrc = medGetP(BloodNum, 1, "-")
    Bldyy = medGetP(BloodNum, 2, "-")
    Bldno = Val(medGetP(BloodNum, 3, "-"))

    sSql = " SELECT a.compocd,b.abbrnm,a.centercd,c.field1 as buildnm " & _
           " FROM " & _
                    T_BBS401 & " a," & T_BBS006 & " b," & T_COM003 & " c" & _
           " WHERE" & _
                    "      " & DBW("a.bldsrc", Bldsrc, 2) & _
                    " AND  " & DBW("a.bldyy", Bldyy, 2) & _
                    " AND  " & DBW("a.bldno ", Bldno, 2) & _
                    " AND  " & DBW("c.cdindex", BC2_CENTER, 2) & _
                    " AND a.stscd in (" & _
                                        DBV("stscd", BBSBloodStatus.stsENTER) & "," & _
                                        DBV("stscd", BBSBloodStatus.stsRETURN) & ")" & _
                    " AND a.compocd=b.compocd " & _
                    " AND a.centercd=c.cdval1"
    
    If TF = True Then
        sSql = sSql & _
                    " AND " & DBW("a.centercd", centercd, 2) & _
                    " AND " & DBW("a.splitinfg", "0", 2)
    Else
        sSql = sSql & _
                    " AND (" & DBW("a.centercd<>", centercd) & _
                         " or " & DBW("a.splitoutfg", "1", 2) & _
                         " or " & DBW("a.splitinfg", "1", 2) & ")"
    End If
   
   Set Get_Split_TarGet = New Recordset
   Get_Split_TarGet.Open sSql, DBConn
End Function

Public Function Get_BLOOD_INFORMATION(BloodNum As String, CompoCd As String) As Recordset
'------------------------------------------------------------------------------------
'혈액번호와 component 코드를 입력받아서 BBS401에서 해당혈액의 기본정보를 가지고 온다.
'혈액분획과 혈액History에서 같이 사용한다.
'------------------------------------------------------------------------------------
    Dim sSql   As String
    Dim Bldsrc As String
    Dim Bldyy  As String
    Dim Bldno  As Long
    
    Bldsrc = medGetP(BloodNum, 1, "-")
    Bldyy = medGetP(BloodNum, 2, "-")
    Bldno = Val(medGetP(BloodNum, 3, "-"))

    sSql = " SELECT bldsrc,bldyy,bldno,compocd,volumn,abo,rh,reserved,ptid,autofg," & _
           "         coldt,coltm,colid,available,expdt,exptm,entdt,enttm,entid,centercd,hosfg,donorid,donoraccdt" & _
           " FROM " & T_BBS401 & _
           " WHERE" & _
                    "      " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND  " & DBW("bldyy", Bldyy, 2) & _
                    " AND  " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2)
   Set Get_BLOOD_INFORMATION = New Recordset
   Get_BLOOD_INFORMATION.Open sSql, DBConn
End Function
Public Function Get_BLOOD_COMPONENT(ByVal BloodNum As String) As Recordset
'---------------------------------------------------
'혈액번호를 입력받아서 Component코드와 명을 가져온다
'혈액Hitstory에서 사용(frmBBS308)
'---------------------------------------------------
    Dim sSql   As String
    Dim Bldsrc As String
    Dim Bldyy  As String
    Dim Bldno  As Long
    
    Bldsrc = medGetP(BloodNum, 1, "-")
    Bldyy = medGetP(BloodNum, 2, "-")
    Bldno = Val(medGetP(BloodNum, 3, "-"))

    sSql = " SELECT a.compocd,b.abbrnm as field1,a.centercd,c.field1 as buildnm " & _
           " FROM " & _
                    T_BBS401 & " a," & T_BBS006 & " b," & T_COM003 & " c" & _
           " WHERE" & _
                    "      " & DBW("a.bldsrc", Bldsrc, 2) & _
                    " AND  " & DBW("a.bldyy", Bldyy, 2) & _
                    " AND  " & DBW(" a.bldno", Bldno, 2) & _
                    " AND a.compocd=b.compocd " & _
                    " AND a.centercd=c.cdval1" & _
                    " AND " & DBW("c.cdindex", BC2_CENTER, 2)
            
   Set Get_BLOOD_COMPONENT = New Recordset
   Get_BLOOD_COMPONENT.Open sSql, DBConn
End Function


'#########################
'# 혈액 History (BBS209)##
'#########################
Public Function Get_Ent_InFo(ByVal BloodNum As String, ByVal CompoCd As String)
'혈액입고내역에서 입고정보를 가지고 오며, Dictionary 필드를 만들어준다.
     Dim Rs As New Recordset
     Dim strSQL As String
     Dim Bldsrc As String
     Dim Bldyy As String
     Dim Bldno As Long
     Dim cnt As Long: cnt = 1
     
     Bldsrc = medGetP(BloodNum, 1, "-")
     Bldyy = medGetP(BloodNum, 2, "-")
     Bldno = Val(medGetP(BloodNum, 3, "-"))
    '입고일,입고자ID,센터코드,지정혈액여부,자가헌혈여부,환자Id(지정,자가),병원헌열입고,분획입고
     strSQL = " SELECT  entdt,enttm,entid,centercd,reserved,autofg,ptid,hosfg,splitoutfg,splitinfg,splitoutdt," & _
                        "realexpdt,realexptm,expid,exprcvid,exprsncd,expbilldiv,irrfg,irrdt,irrid,irrtm" & _
              " FROM " & T_BBS401 & _
              " WHERE" & _
                       "     " & DBW("bldsrc", Bldsrc, 2) & _
                       " AND " & DBW("bldyy", Bldyy, 2) & _
                       " AND " & DBW("bldno", Bldno, 2) & _
                       " AND " & DBW("compocd", CompoCd, 2)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    mvarBldDic.Clear
    
    '처리일자,번호,  내역,담당자,검사방법,검체번호,간호사,환자,비고
    mvarBldDic.FieldInialize "indexdt,indextm,seq", "contents,chargenm,xmethod,spcno,nurse,ptid,etc"
    
    If Rs.EOF Then
    Else
        With mvarBldDic
            .Sort = False
            .AddNew Join(Array(Rs.Fields("entdt").Value & "", Rs.Fields("enttm").Value & "", cnt), COL_DIV), _
                    Join(Array("", "", "", "", "", "", ""), COL_DIV)
            If .Exists(Rs.Fields("entdt").Value & "" & COL_DIV & Rs.Fields("enttm").Value & "" & COL_DIV & cnt) Then
                .Fields("contents") = "입고"
                .Fields("chargenm") = GetEmpNm(Rs.Fields("entid").Value & "")
                
                If Rs.Fields("reserved").Value & "" = "0" And Rs.Fields("autofg").Value & "" = "0" Then
                    .Fields("etc") = "일반혈액,"
                End If
                    
                If Rs.Fields("reserved").Value & "" = "1" Then
                    .Fields("etc") = .Fields("etc") & "지정혈액,"
                End If
                
                If Rs.Fields("autofg").Value & "" = "1" Then
                    .Fields("etc") = .Fields("etc") & "자가헌혈,"
                End If
                
                If Rs.Fields("hosfg").Value & "" = "1" Then
                    .Fields("etc") = .Fields("etc") & "병원헌혈입고,"
                End If
                .Fields("etc") = Mid(.Fields("etc"), 1, Len(.Fields("etc")) - 1)
                If Rs.Fields("splitoutfg").Value & "" = "1" Then
                    '분획출고
                    cnt = cnt + 1
                    .AddNew Join(Array(Rs.Fields("splitoutdt").Value & "", "", cnt), COL_DIV), _
                            Join(Array("", "", "", "", "", "", ""), COL_DIV)
                    If .Exists(Rs.Fields("splitoutdt").Value & "" & COL_DIV & "" & COL_DIV & cnt) Then
                        .Fields("contents") = "분획출고"
                        .Fields("chargenm") = GetEmpNm(Rs.Fields("entid").Value & "")
                    End If
                End If
                
                If Rs.Fields("splitinfg").Value & "" = "1" Then
                    '분획입고
                    cnt = cnt + 1
                    .AddNew Join(Array(Rs.Fields("entdt").Value & "", Rs.Fields("enttm").Value & "", cnt), COL_DIV), Join(Array("", "", "", "", "", "", ""), COL_DIV)
                    If .Exists(Rs.Fields("entdt").Value & "" & COL_DIV & Rs.Fields("enttm").Value & "" & COL_DIV & cnt) Then
                        .Fields("contents") = "분획입고"
                        .Fields("chargenm") = GetEmpNm(Rs.Fields("entid").Value & "")
                    End If
                End If
                If Rs.Fields("irrfg").Value & "" = "1" Then
                    cnt = cnt + 1
                    .AddNew Join(Array(Rs.Fields("irrdt").Value & "", Rs.Fields("irrtm").Value & "", cnt), COL_DIV), Join(Array("", "", "", "", "", "", ""), COL_DIV)
                    If .Exists(Rs.Fields("irrdt").Value & "" & COL_DIV & Rs.Fields("irrtm").Value & "" & COL_DIV & cnt) Then
                        .Fields("contents") = "RT처리"
                        .Fields("chargenm") = GetEmpNm(Rs.Fields("irrid").Value & "")
                    End If
                End If
                
                If Rs.Fields("realexpdt").Value & "" <> "" Then
                    '혈액폐기
                    cnt = cnt + 1
                    .AddNew Join(Array(Rs.Fields("realexpdt").Value & "", Rs.Fields("realexptm").Value & "", cnt), COL_DIV), _
                            Join(Array("", "", "", "", "", "", ""), COL_DIV)
                    If .Exists(Rs.Fields("realexpdt").Value & "" & COL_DIV & Rs.Fields("realexptm").Value & "" & COL_DIV & cnt) Then
                        .Fields("contents") = "혈액폐기"
                        .Fields("chargenm") = GetEmpNm(Rs.Fields("expid").Value & "")
                        .Fields("nurse") = GetEmpNm(Rs.Fields("exprcvid").Value & "")
                        If Rs.Fields("expbilldiv").Value & "" = "1" Then
                            .Fields("etc") = "환자부담(Yes)" & ","
                        Else
                            .Fields("etc") = "환자부담(No)" & ","
                        End If
                        .Fields("etc") = .Fields("etc") & Get_ExpRsn(Rs.Fields("exprsncd").Value & "")
                    End If
                End If
            End If
        End With
    End If
    Get_XM_InFo Bldsrc, Bldyy, Bldno, CompoCd, cnt
    Set Rs = Nothing
End Function

Public Function Get_Delivery_InFo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As Long, ByVal CompoCd As String, ByVal cnt As Long)
'혈액출고내역에서 출고정보,반환정보,폐기정보를 가지고 온다.
    Dim Rs As New Recordset
    Dim strSQL As String
    
    Dim objGetSql   As New clsBBSSQLStatement
    Dim Rs1         As New Recordset
    Dim strLocalNm  As String
    
    '출고일자,출고자,출고수령자,로컬병원,irr여부,Filter 여부,반환여부,폐기여부,Bag회수여부,반환일자,폐기일자,bag회수일자
    '폐기자,폐기수령자,폐기사유,환자수납여부,폐기사유Remark
    'realexpdt realexptm expid       exprcvid    exprsncd expbilldiv
    
    '---------------------------------------------------------------------------------
    '2001-11-29 수정 : OUTER-JOIN 사용 (로컬출고인 경우 처방접수내역이 없으므로...)
    '---------------------------------------------------------------------------------
    strSQL = " SELECT a.deliverydt,a.deliveryid,a.deliverytm,a.rcvid,a.localcd,a.irrafg,a.filter," & _
                    " a.retfg,a.expfg,a.bagfg,a.retdt,a.rettm,a.bagdt,a.bagtm," & _
                    " b.splitoutfg,a.retid,a.retrcvid," & _
                    " a.bagid,a.bagrcvid,a.retrmk," & _
                    " b.realexpdt,b.realexptm,b.expid,b.exprcvid,b.exprsncd,b.expbilldiv,b.exprsnrmk," & _
                    " c.ptid" & _
             " FROM " & _
                      T_BBS402 & " a," & T_BBS401 & " b," & T_BBS202 & " c" & _
             " WHERE" & _
                      "      " & DBW("a.bldsrc", Bldsrc, 2) & _
                      " AND  " & DBW("a.bldyy ", Bldyy, 2) & _
                      " AND  " & DBW("a.bldno ", Bldno, 2) & _
                      " AND  " & DBW("a.compocd", CompoCd, 2) & _
                      " AND  " & DBJ("c.workarea=*a.workarea") & _
                      " AND  " & DBJ("c.accdt=*a.accdt") & _
                      " AND  " & DBJ("c.accseq=*a.accseq") & _
                      " AND  " & DBJ("b.bldsrc=*a.bldsrc") & _
                      " AND  " & DBJ("b.bldyy=*a.bldyy") & _
                      " AND  " & DBJ("b.bldno=*a.bldno") & _
                      " AND  " & DBJ("b.compocd=*a.compocd")
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF Then
    Else
        With mvarBldDic
            Do Until Rs.EOF
                cnt = cnt + 1
                .AddNew Join(Array(Rs.Fields("deliverydt").Value, Rs.Fields("deliverytm").Value & "", cnt), COL_DIV), _
                        Join(Array("", "", "", "", "", "", ""), COL_DIV)
                If .Exists(Rs.Fields("deliverydt").Value & "" & COL_DIV & Rs.Fields("deliverytm").Value & "" & COL_DIV & cnt) Then
                    '---------------------------
                    '혈액출고
                    '---------------------------
                    .Fields("contents") = "출고"
                    .Fields("chargenm") = GetEmpNm(Rs.Fields("deliveryid").Value & "")
                    .Fields("nurse") = GetEmpNm(Rs.Fields("rcvid").Value & "")
                    .Fields("ptid") = Rs.Fields("ptid").Value & ""
                    
                    '2001-11-29 추가 : 로컬병원명 가져오기
                    strLocalNm = ""
                    If Trim(Rs.Fields("localcd").Value & "") <> "" Then
                        strLocalNm = GetLocalNm(Rs.Fields("localcd").Value & "")
                    End If
                    
                    .Fields("etc") = IIf(Rs.Fields("localcd").Value & "" <> "", "Local출고:" & strLocalNm & "" & ",", "")
                    
                    If Rs.Fields("irrafg").Value & "" = "1" Then .Fields("etc") = .Fields("etc") & "RT: Yes,"
                    
                    If Rs.Fields("filter").Value & "" = "1" Then .Fields("etc") = .Fields("etc") & "Filter 출고,"
                    
                    If Len(.Fields("etc")) > 1 Then .Fields("etc") = Mid(.Fields("etc"), 1, Len(.Fields("etc")) - 1)
                    
                    '---------------------------
                    '혈액반환
                    '---------------------------
                    If Rs.Fields("retfg").Value & "" = "1" Then
                        cnt = cnt + 1
                        .AddNew Join(Array(Rs.Fields("retdt").Value & "", Rs.Fields("rettm").Value & "", cnt), COL_DIV), _
                                Join(Array("", "", "", "", "", "", ""), COL_DIV)
                        If .Exists(Rs.Fields("retdt").Value & "" & COL_DIV & Rs.Fields("rettm").Value & "" & COL_DIV & cnt) Then
                            .Fields("contents") = "반환"
                            .Fields("chargenm") = GetEmpNm(Rs.Fields("retid").Value & "")
                            .Fields("nurse") = GetEmpNm(Rs.Fields("retrcvid").Value & "")
                            .Fields("etc") = Rs.Fields("retrmk").Value & ""
                        End If
                    End If
                    
                    '---------------------
                    '혈액폐기
                    '---------------------
                    If Rs.Fields("expfg").Value = "1" Then
                        cnt = cnt + 1
                       .AddNew Join(Array(Rs.Fields("realexpdt").Value & "", Rs.Fields("realexptm").Value & "", cnt), COL_DIV), _
                               Join(Array("", "", "", "", "", "", ""), COL_DIV)
                        If .Exists(Rs.Fields("realexpdt").Value & "" & COL_DIV & Rs.Fields("realexptm").Value & "" & COL_DIV & cnt) Then
                            
                            .Fields("contents") = "폐기"
                            .Fields("chargenm") = GetEmpNm(Rs.Fields("expid").Value & "")
                            .Fields("nurse") = GetEmpNm(Rs.Fields("exprcvid").Value & "")
                            .Fields("etc") = "폐기사유:" & Rs.Fields("exprsncd").Value & "" & ","
                            If Rs.Fields("expbilldiv").Value & "" = "1" Then .Fields("etc") = .Fields("etc") & "환자부담,"
                            If Rs.Fields("exprsnrmk").Value & "" <> "" Then .Fields("etc") = .Fields("etc") & Rs.Fields("exprsnrmk").Value & "" & ","
                            
                            .Fields("etc") = Mid(.Fields("etc"), 1, Len(.Fields("etc")) - 1)
                        End If
                    End If
                    
                    '----------------------------
                    '혈액 Bag회수
                    '----------------------------
                    If Rs.Fields("bagfg").Value = "1" Then
                        cnt = cnt + 1
                       .AddNew Join(Array(Rs.Fields("bagdt").Value & "", Rs.Fields("bagtm").Value & "", cnt), COL_DIV), _
                               Join(Array("", "", "", "", "", "", ""), COL_DIV)
                        If .Exists(Rs.Fields("bagdt").Value & "" & COL_DIV & Rs.Fields("bagtm").Value & "" & COL_DIV & cnt) Then
                            .Fields("contents") = "회수"
                            .Fields("chargenm") = GetEmpNm(Rs.Fields("bagid").Value & "")
                            .Fields("nurse") = GetEmpNm(Rs.Fields("bagrcvid").Value & "")
                        End If
                    End If
                End If
                Rs.MoveNext
            Loop
            
        End With
    End If
    Set Rs = Nothing
    Set Rs1 = Nothing
    Get_Trans_InFo Bldsrc, Bldyy, Bldno, CompoCd, cnt
End Function
Public Function Get_XM_InFo(Bldsrc As String, Bldyy As String, Bldno As Long, CompoCd As String, seq As Long)
'bbs302에서 혈액의 검사정보를 조회한다.
    Dim Rs   As New Recordset
    Dim strSQL As String
    
    
    strSQL = " SELECT a.xmethod,a.step1,a.step2,a.step3,a.step4,a.rstv,a.spcyy,a.spcno,a.vfydt,a.vfyid,a.vfytm,a.cancelfg,a.stat," & _
                     "a.canceldt,a.canceltm,a.cancelid,a.rmk,b.ptid,a.filterfg" & _
            " FROM " & _
                     T_BBS302 & " a," & T_BBS202 & " b" & _
            " WHERE" & _
                     "     " & DBW("bldsrc", Bldsrc, 2) & _
                     " AND " & DBW("bldyy", Bldyy, 2) & _
                     " AND " & DBW("bldno", Bldno, 2) & _
                     " AND " & DBW("compocd", CompoCd, 2) & _
                     " AND a.workarea=b.workarea" & _
                     " AND a.accdt=b.accdt" & _
                     " AND a.accseq=b.accseq"
        
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        With mvarBldDic
            Do Until Rs.EOF
            
                seq = seq + 1
                .AddNew Join(Array(Rs.Fields("vfydt").Value & "", Rs.Fields("vfytm").Value & "", seq), COL_DIV), _
                        Join(Array("", "", "", "", "", "", ""), COL_DIV)
                If .Exists(Rs.Fields("vfydt").Value & "" & COL_DIV & Rs.Fields("vfytm").Value & "" & COL_DIV & seq) Then
                    '-----------------------------
                    'XM검사
                    '-----------------------------
                    .Fields("contents") = "XM검사"
                    Select Case Rs.Fields("xmethod").Value
                        Case "0": .Fields("xmethod") = "Major"
                        Case "1": .Fields("xmethod") = "Minor"
                        Case "2": .Fields("xmethod") = "Both"
                    End Select
                    
                    Select Case Rs.Fields("rstv").Value
                        Case "1": .Fields("etc") = "검사결과:" & "OK,"
                        Case "0": .Fields("etc") = "검사결과:" & "NOT,"
                        Case Else: .Fields("etc") = "미검사,"
                    End Select
                    
                    Dim objXM As clsCrossMatching
                    Dim RsXm As Recordset
                    Dim i As Long
                    Dim strStep As String
                    Dim strDetail As String
                    
                    strStep = ""
                    strDetail = ""
                    
                    Set objXM = New clsCrossMatching
                    Set RsXm = New Recordset
                    
                    Set RsXm = objXM.Get_XM_Step
                    
                    strStep = RsXm.Fields("text1").Value & ""
                    
                    Set RsXm = Nothing
                    Set objXM = Nothing
                    
                    If strStep <> "" Then '상세결과가 설정이 된 경우에만 표시하도록...
                        For i = 1 To 4
                            If Rs.Fields("step" & i).Value & "" = "1" Then '결과가 하나라도 들어 있으면 표시
                                strDetail = strDetail & medGetP(strStep, i, ";") & ";"
                            End If
                        Next
                        If strDetail <> "" Then
                            strDetail = Mid(strDetail, 1, Len(strDetail) - 1)
                            
                            .Fields("etc") = .Fields("etc") & "상세결과:" & strDetail & ","
                        Else
                            .Fields("etc") = .Fields("etc") & "상세결과:없음,"
                        End If
                    End If
                    
                    If Rs.Fields("stat").Value & "" = "1" Then .Fields("etc") = .Fields("etc") & "응급,"
                    
                    If Rs.Fields("rmk").Value & "" <> "" Then .Fields("etc") = .Fields("etc") & Rs.Fields("rmk").Value & "" & ","
                    
                    If Rs.Fields("filterfg").Value & "" = "1" Then .Fields("etc") = .Fields("etc") & "Filter 처리,"
                    
                    If Len(.Fields("etc")) > 1 Then .Fields("etc") = Mid(.Fields("etc"), 1, Len(.Fields("etc")) - 1)
                    
                    .Fields("spcno") = Rs.Fields("spcyy").Value & "" & "-" & Rs.Fields("spcno").Value & ""
                    .Fields("chargenm") = GetEmpNm(Rs.Fields("vfyid").Value & "")
                    .Fields("ptid") = Rs.Fields("ptid").Value & ""
                    
                    '-------------------------------
                    'Assign 취소
                    '-------------------------------
                    If Rs.Fields("cancelfg").Value & "" = "1" Then
                        seq = seq + 1
                        .AddNew Join(Array(Rs.Fields("canceldt").Value & "", Rs.Fields("canceltm").Value & "", seq), COL_DIV), Join(Array("", "", "", "", "", "", ""), COL_DIV)
                        If .Exists(Rs.Fields("canceldt").Value & "" & COL_DIV & Rs.Fields("canceltm").Value & "" & COL_DIV & seq) Then
                            .Fields("contents") = "Assign취소"
                            .Fields("chargenm") = GetEmpNm(Rs.Fields("cancelid").Value & "")
                        End If
                    End If
                End If
                Rs.MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
    Get_Delivery_InFo Bldsrc, Bldyy, Bldno, CompoCd, seq
End Function
Public Function Get_Trans_InFo(Bldsrc As String, Bldyy As String, Bldno As Long, CompoCd As String, seq As Long)
'----------------------------
'수혈정보 조회
'----------------------------
    Dim Rs As New Recordset
    Dim sSql As String
    
    sSql = " SELECT a.transdt,a.transtm,a.transvol,a.reactdiv,a.reactrmk,b.ptid" & _
           " FROM " & _
                    T_BBS501 & " a," & T_BBS202 & " b " & _
           " WHERE" & _
                    "     a.workarea=b.workarea" & _
                    " AND a.accdt=b.accdt" & _
                    " AND a.accseq=b.accseq " & _
                    " AND " & DBW("a.bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("a.bldyy", Bldyy, 2) & _
                    " AND " & DBW("a.bldno", Bldno, 2) & _
                    " AND " & DBW("a.compocd", CompoCd, 2) & _
                    " AND  a.transdt <> null AND a.transdt<>' ' "
        
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        With mvarBldDic
            Do Until Rs.EOF
                '---------------
                '수혈정보 조회
                '---------------
                seq = seq + 1
                .AddNew Join(Array(Rs.Fields("transdt").Value & "", Rs.Fields("transtm").Value & "", seq), COL_DIV), _
                        Join(Array("", "", "", "", "", "", ""), COL_DIV)
                If .Exists(Rs.Fields("transdt").Value & "" & COL_DIV & Rs.Fields("transtm").Value & "" & COL_DIV & seq) Then
                    .Fields("contents") = "수혈"
                    .Fields("transpt") = Rs.Fields("ptid").Value & ""
                    .Fields("etc") = Rs.Fields("transvol").Value & "" & "cc,"
                    If .Fields("etc") = "cc," Then .Fields("etc") = ""
                    If Rs.Fields("reactdiv").Value & "" = "1" Then .Fields("etc") = "부작용발생:" & Rs.Fields("reactrmk").Value & "" & ","
                    .Fields("etc") = Mid(.Fields("etc"), 1, Len(.Fields("etc")) - 1)
                End If
                Rs.MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
End Function

'############################
'# 병동채혈      (BBS104,5)##
'############################

Public Function Get_ORDER_103(ByVal orddt As String, Optional ByVal wardid As String = "") As Recordset
    '병동입력후 조회시 화면에 보여주는 정보(901)
    '처방(donefg=0),업무구분(bussdiv in(2,3))=병동,응급
    Dim strSQL As String
    
    Set Get_ORDER_103 = New Recordset
        strSQL = " SELECT  distinct a.ptid ,b." & F_PTNM & " as ptnm,a.bussdiv,a.bedindt,a.reqdt " & _
                 " FROM " & _
                          T_LAB101 & " a," & T_HIS001 & " b," & T_HIS003 & " d" & _
                 " WHERE" & _
                          " " & DBW("a.orddiv", C_WORKAREA, 2) & _
                          " AND " & DBW("a.orddt", orddt, 2) & _
                          " AND " & DBW("a.donefg", stsORDER, 2) & _
                          " AND a.bussdiv in('2','3')" & _
                          " AND " & DBW("a.wardid", wardid, 2) & _
                          " AND a.ptid=b." & F_PTID & _
                          " AND a.wardid=d." & F_DEPTCD
    
    Set Get_ORDER_103 = New Recordset
    Get_ORDER_103.Open strSQL, DBConn
    
End Function


Public Function Get_PtInfo(PtId As String, ByVal Bussdiv As String, Optional orddt As String = "", _
                           Optional accdt As String = "", Optional accseq As String = "") As Recordset
'frmBBS104(간호사채혈)에서 조회된 환자를 선택하였을때 환자의 정보를 조회한다.
'frmBBS105(외래채혈)에서 조회된 환자를 선택하였을때 환자의 정보를 조회한다.
'accdt,accseq가 입력되었다면, 추가처방이다.

    Dim strSQL As String
        
    Set Get_PtInfo = New Recordset
    strSQL = " SELECT distinct a.ptid, a.deptcd,a.wardid,a.hosilid,a.bedid,a.bedindt," & _
                     " b." & F_PTNM & " as ptnm, " & F_SSN2("b") & " as ssn," & _
                     " c." & F_DEPTNM & " as deptnm" & _
            " FROM " & _
                     T_LAB101 & " a," & T_HIS001 & " b," & T_HIS003 & " c," & T_LAB102 & " d" & _
            " WHERE" & _
                     "     " & DBW("a.ptid=", PtId) & _
                     " AND a.ptid=d.ptid" & _
                     " AND a.orddt=d.orddt" & _
                     " AND a.ordno=d.ordno" & _
                     " AND " & DBW("a.bussdiv=", Bussdiv) & _
                     " AND a.ptid=b." & F_PTID & _
                     " AND " & DBJ("a.deptcd *=c." & F_DEPTCD)
    
    If orddt <> "" Then
        strSQL = strSQL & _
                        " AND " & DBW("a.orddt=", orddt)
    End If
    If accdt <> "" And accseq <> "" Then
        strSQL = strSQL & _
                        " AND a.ptid=d.ptid" & _
                        " AND a.orddt=d.orddt" & _
                        " AND a.ordno=d.ordno" & _
                        " AND " & DBW("d.workarea=", C_WORKAREA) & _
                        " AND " & DBW("d.accdt=", accdt) & _
                        " AND " & DBW("d.accseq=", accseq)
    Else
        strSQL = strSQL & _
                        " AND a.donefg in( '" & stsORDER & "','" & stsCOLLECT & "')"
    End If
    
    Set Get_PtInfo = New Recordset
    Get_PtInfo.Open strSQL, DBConn
    
End Function
Public Function Get_Order_104(ByVal PtId As String, ByVal orddt As String) As Recordset
    'frm104(간호사채혈시 환자id를 넘기면 처방 정보를 보여준다.
    '처방(donefg=0),업무구분(bussdiv in(2,3))=병동,응급

    Dim strSQL As String
    
    strSQL = " SELECT b.orddt,b.ordno,b.ordseq,b.ordcd,b.statfg,a.reqdt," & _
                      " a.reqtm,a.deptcd,a.orddoct,a.bussdiv,a.wardid,b.mesg " & _
             " FROM " & _
                      T_LAB101 & " a," & T_LAB102 & " b" & _
             " WHERE" & _
                      "     " & DBW("a.ptid=", PtId) & _
                      " AND " & DBW("b.orddt=", orddt) & _
                      " AND " & DBW("a.donefg=", stsORDER) & _
                      " AND a.bussdiv in('2','3') " & _
                      " AND a.ptid =b.ptid " & _
                      " AND a.orddt=b.orddt " & _
                      " AND a.ordno=b.ordno " & _
                      " AND a.donefg=b.donefg " & _
             " ORDER BY a.ptid,b.ordno "
    
    Set Get_Order_104 = New Recordset
    Get_Order_104.Open strSQL, DBConn
    
End Function

Public Function Get_BagGround(AccNo As String, lenAccNo As Long) As Recordset
'접수번호를 가지고 추가검체를 요청할 자료를 조회한다.
'조건 =접수일,접수Seq
    Dim strSQL As String
    Dim NowDate As String
    Dim strAccDt As String
    Dim strAccNo As String
    
    strAccDt = Mid(AccNo, 1, lenAccNo)
    strAccNo = Val(Mid(AccNo, lenAccNo + 2))
    NowDate = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    strSQL = " SELECT a.ptid,a.reqdt,a.bedindt,a.orddoct,a.deptcd,a.wardid,a.hosilid,a.bedid,a.bussdiv," & _
                      " b.ordcd,b.orddt,b.unitqty,b.ordno,b.stscd,b.dcfg," & _
                      " c.testnm,d." & F_PTNM & " as ptnm," & F_SSN2("d") & " as ssn,b.mesg" & _
             " FROM " & _
                      T_HIS001 & " d," & T_BBS001 & " c," & T_LAB101 & " a," & T_LAB102 & " b" & _
             " WHERE" & _
                      "     " & DBW("b.workarea=", "B") & " AND " & DBW("b.accdt", strAccDt, 2) & _
                      " AND " & DBW("b.accseq", strAccNo, 2) & _
                      " AND a.ordno=b.ordno " & _
                      " AND a.ptid=b.ptid" & _
                      " AND a.ptid=d." & F_PTID & _
                      " AND a.orddt=b.orddt" & _
                      " AND b.ordcd=c.testcd" & _
                      " AND c.applydt=(SELECT max(d.applydt)" & _
                                      " FROM " & T_BBS001 & " d" & _
                                      " WHERE c.testcd = d.testcd" & _
                                      " AND   " & DBW("d.applydt<", NowDate, 2) & ")"
             
             
    Set Get_BagGround = New Recordset
    Get_BagGround.Open strSQL, DBConn
    
End Function

'#################################
'# 처방상태 조회(frmBBS206)  #####
'#################################

Public Function Order_Status_LIst(FDt As String, TDt As String, TF As Boolean, _
                                  strque As String, _
                                  Optional ByVal PtId As String = "", _
                                  Optional ByVal wardid As String = "", _
                                  Optional ByVal DeptCd As String = "", _
                                  Optional ByVal DoctCd As String = "")
'처방상태 조회
'TABLE:LAB101,LAB102,BBS302,HIS003(부서마스터),HIS005(의사마스터)
'      BBS303(ABO결과내역),HIS001(환자마스터),BBS001(처방명)
'예정일로 찾기(TF=true);처방일로찾기(TF=False)
    
    Order_Status_LIst = " SELECT a.ptid,b.stscd,b.ordcd,h.testnm," & _
                        "        b.unitqty,a.reqdt,a.deptcd,a.wardid," & _
                        "        a.orddoct,a.orddt,a.ordno,b.dcfg,b.mesg," & _
                        "        c.assigncancelcnt,c.deliverycnt,c.assigncnt,c.retcnt,c.expcnt" & _
                        " FROM " & T_LAB101 & " a," & T_LAB102 & " b," & _
                                   T_BBS203 & " c," & T_BBS001 & " h"
    If TF = True Then
        Order_Status_LIst = Order_Status_LIst & _
                                              " WHERE a.reqdt between " & DBV("reqdt", FDt) & " AND " & DBV("reqdt", TDt)
    Else
        Order_Status_LIst = Order_Status_LIst & _
                                              " WHERE a.orddt between " & DBV("orddt", FDt) & " AND " & DBV("orddt", TDt)
    End If
    
    
    If PtId <> "" Then
        Order_Status_LIst = Order_Status_LIst & _
                                              " AND " & DBW("a.ptid", PtId, 2)
    End If
    If wardid <> "" Then
        Order_Status_LIst = Order_Status_LIst & _
                                              " AND " & DBW("a.wardid", wardid, 2)
    End If
    If DeptCd <> "" Then
        Order_Status_LIst = Order_Status_LIst & _
                                              " AND " & DBW("a.deptcd", DeptCd, 2)
    End If
    If DoctCd <> "" Then
        Order_Status_LIst = Order_Status_LIst & _
                                              " AND " & DBW("a.orddoct", DoctCd, 2)
    End If
'상태에 따른 조건
    Dim blnQue As Boolean
    
    If medGetP(strque, 1, COL_DIV) <> "" Or medGetP(strque, 2, COL_DIV) <> "" Or _
                                            medGetP(strque, 3, COL_DIV) <> "" Or _
                                            medGetP(strque, 4, COL_DIV) <> "" Or _
                                                                     medGetP(strque, 5, COL_DIV) <> "" Then
         
         Order_Status_LIst = Order_Status_LIst & " AND b.donefg in("
        If medGetP(strque, 1, COL_DIV) = "0" Then
            Order_Status_LIst = Order_Status_LIst & "'" & stsORDER & "',"
            blnQue = True
        End If
        If medGetP(strque, 2, COL_DIV) = "1" Then
            Order_Status_LIst = Order_Status_LIst & "'" & stsCOLLECT & "',"
            blnQue = True
        End If
        If medGetP(strque, 3, COL_DIV) = "2" Then
            Order_Status_LIst = Order_Status_LIst & " '" & stsACCESS & "',"
            blnQue = True
        End If
        If medGetP(strque, 4, COL_DIV) = "3" Then
            Order_Status_LIst = Order_Status_LIst & "'" & stsINPROCESS & "',"
            blnQue = True
        End If
        If medGetP(strque, 5, COL_DIV) = "4" Then
            Order_Status_LIst = Order_Status_LIst & "'" & stsINPROCESS & "',"
            blnQue = True
        End If
        If blnQue = True Then
            Order_Status_LIst = Mid(Order_Status_LIst, 1, Len(Order_Status_LIst) - 1) & ")"
        End If
    End If
    
'lab101과 lab102의 조건
    Order_Status_LIst = Order_Status_LIst & _
                                          " AND a.ptid=b.ptid" & _
                                          " AND a.orddt=b.orddt" & _
                                          " AND a.ordno=b.ordno "
'lab102과 bbs405의 조건(Assign개수와 출고개수를 구하기위해)
    Order_Status_LIst = Order_Status_LIst & _
                                          " AND " & DBJ("b.workarea*=c.workarea") & _
                                          " AND " & DBJ("b.accdt*=c.accdt") & _
                                          " AND " & DBJ("b.accseq*=c.accseq")
'LAB102,BBS001의 조건(처방명)
    Order_Status_LIst = Order_Status_LIst & _
                         " AND b.ordcd=h.testcd " & _
                         " AND h.applydt=(SELECT max(i.applydt)" & _
                         "                FROM " & T_BBS001 & " i" & _
                         "                WHERE b.ordcd = i.testcd" & _
                         "                AND " & DBW("i.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & ")"


    Order_Status_LIst = Order_Status_LIst & _
                                          "ORDER BY a.ptid"
    
End Function

'####################################
'# 헌혈접수상태조회(frmBBS405)  #####
'####################################
Public Function Get_DonorQuery(ByVal Frdt As String, ByVal Todt As String) As Recordset
    Dim sSql As String
    Dim Rs   As New Recordset
    
    sSql = " SELECT a.donornm,a.dob,a.sex,a.abo,a.rh," & _
                    " b.donoraccdt,b.tmpid,b.donorcd,b.cancelfg,b.reservedid,b.donorid,b.bldsrc,b.bldyy,b.bldno,b.volumn," & _
                    " c.okdiv1,c.okdiv2,c.okdiv3,b.entfg" & _
           " FROM " & _
                     T_BBS601 & " a," & T_BBS602 & " b," & T_BBS603 & " c" & _
           " WHERE" & _
                    " b.donoraccdt between" & DBV("donoraccdt", Frdt) & " AND " & DBV("donoraccdt", Todt) & _
                    " AND b.donorid=a.donorid " & _
                    " AND b.donorid=c.donorid" & _
                    " AND b.donoraccdt=c.donoraccdt" & _
           " ORDER BY a.donornm,b.donoraccdt"
    
     
'     SSQL = " SELECT a.donornm,a.dob,a.sex,a.abo,a.rh," & _
'                    " b.donoraccdt,b.donorid,b.bldsrc,b.bldyy,b.bldno,b.volumn," & _
'                    " d.reservedid,d.tmpid,d.donorcd,d.cancelfg,d.entfg, " & _
'                    " c.okdiv1,c.okdiv2,c.okdiv3" & _
'           " FROM " & _
'                     T_BBS601 & " a," & T_BBS401 & " b," & T_BBS603 & " c," & T_BBS602 & " d " & _
'           " WHERE" & _
'                    " b.donoraccdt between" & DBV("donoraccdt", Frdt) & " AND " & DBV("donoraccdt", Todt) & _
'                    " AND b.donorid=a.donorid " & _
'                    " AND b.donorid=c.donorid " & _
'                    " AND b.donoraccdt=c.donoraccdt " & _
'                    " AND b.donorid = d.donorid AND b.donoraccdt = d.donoraccdt " & _
'           " ORDER BY a.donornm,b.donoraccdt"
'
    
    
    
  '  SSQL = " SELECT distinct a.donornm,a.dob,a.sex,a.abo,a.rh," & _
                    " b.donoraccdt,b.tmpid,b.donorcd,b.cancelfg,b.reservedid,b.donorid,b.bldsrc,b.bldyy,b.bldno,b.volumn," & _
                    " c.okdiv1,c.okdiv2,c.okdiv3,d.reserved" & _
           " FROM " & _
                    T_BBS401 & " d," & T_BBS601 & " a," & T_BBS602 & " b," & T_BBS603 & " c" & _
           " WHERE" & _
                              DBW("b.donoraccdt>=", Frdt) & _
                    " AND " & DBW("b.donoraccdt<=", Todt) & _
                    " AND b.donorid=a.donorid " & _
                    " AND b.donorid=c.donorid" & _
                    " AND b.donoraccdt=c.donoraccdt" & _
                    " AND b.bldsrc=d.bldsrc" & _
                    " AND b.bldyy=d.bldyy" & _
                    " AND b.bldno=d.bldno" & _
           " ORDER BY a.donornm,b.donoraccdt"
    Set Get_DonorQuery = New Recordset
    Get_DonorQuery.Open sSql, DBConn

End Function

Public Function Get_DonorBlood(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT b.abbrnm,a.* FROM " & T_BBS401 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                     DBW("donorid=", Donorid) & _
           " AND " & DBW("donoraccdt=", donoraccdt) & _
           " AND a.compocd=b.compocd"
            
    Set Get_DonorBlood = New Recordset
    Get_DonorBlood.Open sSql, DBConn
End Function

'####################################
'# 접수취소관련문장(frmBBS108)  #####
'####################################

Public Function Get_Ptinformation(ByVal accdtseq As String, ByVal AccFormat As Long) As Recordset
'접수번호를 가지고 환자정보를 구한다....
    Dim Rs     As New Recordset
    Dim sSql   As String
    Dim accdt  As String
    Dim accseq As String
    
    accdt = Trim(Mid(accdtseq, 1, AccFormat))
    accseq = Val(Trim(Mid(accdtseq, AccFormat + 2)))
    
    sSql = " SELECT a.ptid,b.donefg,b.orddt,c.testnm,b.ordcd,b.unitqty,b.dcfg,a.reqdt," & _
                      "b.ordno,b.ordseq," & _
                      "b.stscd,b.mesg,a.orddoct,a.deptcd,a.wardid,a.hosilid,a.bedid," & _
                      "d." & F_PTNM & " as ptnm, " & F_SSN2("d") & " as ssn" & _
             " FROM " & _
                      T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c," & T_HIS001 & " d" & _
             " WHERE" & _
                      "     " & DBW("b.workarea", C_WORKAREA, 2) & _
                      " AND " & DBW("b.accdt", accdt, 2) & _
                      " AND " & DBW("b.accseq", accseq, 2) & _
                      " AND a.ptid=b.ptid" & _
                      " AND a.orddt=b.orddt" & _
                      " AND a.ordno=b.ordno " & _
                      " AND a.ptid=d." & F_PTID & _
                      " AND b.ordcd=c.testcd" & _
                      " AND c.applydt=(SELECT max(d.applydt)" & _
                                     " FROM " & T_BBS001 & " d" & _
                                     " WHERE c.testcd = d.testcd" & _
                                     " AND   " & DBW("d.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & ")" & _
             " ORDER BY b.ordno,b.ordseq "
    
    Set Get_Ptinformation = New Recordset
    Get_Ptinformation.Open sSql, DBConn
    
End Function
Public Function Get_PheresisChk(ByVal accdtseq As String, ByVal AccFormat As Long) As Boolean
'성분헌혈여부를 판단한다....
    Dim Rs     As New Recordset
    Dim sSql   As String
    Dim accdt  As String
    Dim accseq As String
    
    accdt = Trim(Mid(accdtseq, 1, AccFormat))
    accseq = Val(Trim(Mid(accdtseq, AccFormat + 2)))
    
    sSql = " SELECT * FROM " & T_BBS202 & _
           " WHERE" & _
                    "     " & DBW("workarea", C_WORKAREA, 2) & _
                    " AND " & DBW("accdt", accdt, 2) & _
                    " AND " & DBW("accseq", accseq, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        If Rs.Fields("pheresisfg").Value & "" = "1" Then
            Get_PheresisChk = True
        Else
            Get_PheresisChk = False
        End If
    Else
        Get_PheresisChk = False
    End If
    Set Rs = Nothing
    
End Function
Public Function Get_CollectChk(ByVal accdtseq As String, ByVal AccFormat As Long) As Boolean
'접수번호를 가지고 접수여부를 판단한다.....
    Dim Rs     As New Recordset
    Dim sSql   As String
    Dim accdt  As String
    Dim accseq As String
    
    accdt = Trim(Mid(accdtseq, 1, AccFormat))
    accseq = Val(Trim(Mid(accdtseq, AccFormat + 2)))
    
    sSql = " SELECT * " & _
           " FROM " & T_LAB102 & _
           " WHERE" & _
                    "     " & DBW("accdt", accdt, 2) & _
                    " AND " & DBW("accseq", accseq, 2) & _
                    " AND " & DBW("workarea", C_WORKAREA, 2)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        Get_CollectChk = True
    Else
        MsgBox "해당 접수번호에 대한 정보가 없습니다.", vbInformation + vbOKOnly, "접수취소"
    End If
    Set Rs = Nothing
End Function
Public Function Get_SpcInformation(ByVal PtId As String, ByVal orddt As String) As Recordset
'검체정보를 가지고 온다....
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT spcyy,spcno FROM " & T_BBS201 & _
           " WHERE" & _
                    " " & DBW("ptid", PtId, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF = False Then
        sSql = " SELECT a.spcyy,a.spcno,a.coldt,a.coltm,c.empnm,a.storeleg,a.storerno,a.storecno, " & _
                        "a.ptid,a.rcvdt,a.rcvtm,a.rcvid,a.colid," & _
                        "b." & F_PTNM & " as ptnm," & F_SSN2("b") & " as ssn" & _
               " FROM " & _
                       T_BBS201 & " a," & T_HIS001 & " b," & T_COM006 & " c" & _
               " WHERE" & _
                        " a.ptid=(SELECT a.ptid " & _
                                 " FROM " & T_HIS001 & " d," & T_BBS201 & " a" & _
                                 " WHERE  " & DBW("a.spcyy", Rs.Fields("spcyy"), 2) & _
                                 " AND " & DBW("a.spcno", Rs.Fields("spcno"), 2) & _
                                 " AND a.ptid=d." & F_PTID & ") " & _
                        " AND a.ptid=b." & F_PTID & _
                        " AND " & DBJ("a.colid*=c.empid")
        
        Set Get_SpcInformation = New Recordset
        Get_SpcInformation.Open sSql, DBConn
    End If
    
    Set Rs = Nothing
End Function

Public Function Spc_TimeChk(ByVal PtId As String) As Long
'검체의 채취 경과시간을 구한다.
    Dim Rs           As New Recordset
    Dim sSql         As String
    Dim strExistence As String
    Dim strCompare   As String
    Dim strCompare1  As String
    Dim coldt        As String
    Dim coltm        As String
    
    Dim lngStoreHour As Long
    
    coldt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    coltm = Format(GetSystemDate, "hhmm")
    strCompare1 = coldt & Mid(coltm, 1, 4)
    strCompare1 = Format(strCompare1, "####-##-## ##:##")
    
    sSql = " SELECT coldt,coltm FROM " & T_BBS201 & " a" & _
           " WHERE" & _
           "      " & DBW("a.ptid=", PtId) & " AND  a.coldt=(SELECT max(b.coldt) " & _
                            " FROM " & T_BBS201 & " b" & _
                            " WHERE " & DBW("b.ptid", PtId, 2) & "" & _
                            " AND a.ptid=b.ptid " & _
                            " AND " & DBW("b.coldt<", coldt, 2) & ") " & _
           " ORDER BY a.coltm desc "
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
        
    If Not Rs.EOF = True Then
        strCompare = Rs.Fields("coldt").Value & "" & Mid(Rs.Fields("coltm").Value & "", 1, 4)
        strCompare = Format(strCompare, "####-##-## ##:##")
        lngStoreHour = CLng(DateDiff("h", strCompare, strCompare1))
        Spc_TimeChk = lngStoreHour
    Else
        Spc_TimeChk = 0
    End If
    
    Set Rs = Nothing
End Function


Public Function Get_ExpRsn(ByVal Exprsncd As String) As String
    Dim Rs   As New Recordset
    Dim sSql As String
    
    sSql = " SELECT cdval1,field1" & _
           " FROM " & T_COM003 & _
           " WHERE" & _
                    " " & DBW("cdindex", BC2_EXP_RESON, 2) & _
                    " AND " & DBW("cdval1", Exprsncd, 2)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF Then
        Get_ExpRsn = ""
    Else
        Get_ExpRsn = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
    
End Function
Private Sub Class_Terminate()
    Set mvarBldDic = Nothing
End Sub

Public Function TransPtidHistory(ByVal PtId As String, ByVal FDate As String, ByVal LDate As String) As String
'------------------------------------------
'환자별 수혈내역에서 해당기간의 환자를 조회
'------------------------------------------
    Dim Rs   As Recordset
    Dim sSql As String
    Dim strSEX As String
    
    
    sSql = " SELECT distinct a." & F_PTNM & " as ptnm ," & F_SEX2("a") & " as sex ," & F_SSN2("a") & " as ssn" & _
           " FROM  " & T_HIS001 & " a," & T_LAB102 & " b" & _
           " WHERE " & DBW("b.ptid=", PtId) & _
           " AND b.orddt between '" & FDate & "' AND '" & LDate & "'" & _
           " AND " & DBW("b.workarea=", C_WORKAREA) & _
           " AND b.ptid=a." & F_PTID
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        If Rs.Fields("sex").Value & "" = "1" Or Rs.Fields("sex").Value & "" = "3" Then
            strSEX = "M"
        ElseIf Rs.Fields("sex").Value & "" = "2" Or Rs.Fields("sex").Value & "" = "4" Then
            strSEX = "F"
        End If
        
        TransPtidHistory = Rs.Fields("ptnm").Value & "" & COL_DIV & _
                           strSEX & "/" & SexCheck(Rs.Fields("ssn").Value & "")
    Else
        TransPtidHistory = ""
    End If
        
    Set Rs = Nothing

End Function
Public Function PtTrasnHistory(ByVal PtId As String, ByVal FDt As String, ByVal TDt As String) As Recordset
'--------------------------------------------------
'환자별 수혈내역에서 환자에 대한 처방을 가지고 온다
'--------------------------------------------------
    Dim sSql As String
    
    sSql = " SELECT c.abbrnm10,c.testnm,c.compocd," & _
         " a.unitqty,a.statfg,a.dcfg,a.donefg,a.stscd,a.workarea,a.accdt,a.accseq,a.orddt,a.ordno," & _
         " b.orddiv,b.reqdt,b.reqtm,b.wardid,b.hosilid,b.deptcd,a.rcvdt,a.rcvtm" & _
         " FROM " & T_BBS001 & " c," & T_LAB101 & " b," & T_LAB102 & " a" & _
         " WHERE " & _
                   DBW("a.ptid=", PtId) & _
         " AND a.orddt between '" & FDt & "' AND '" & TDt & "'" & _
         " AND a.ordcd=c.testcd" & _
         " AND a.ptid=b.ptid" & _
         " AND a.orddt=b.orddt" & _
         " AND a.ordno=b.ordno" & _
         " AND c.applydt = (SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                         " WHERE z.testcd=a.ordcd AND z.applydt<='" & Format(GetSystemDate, PRESENTDATE_FORMAT) & "')" & _
         " ORDER BY a.ptid ,a.orddt,b.orddiv"
         
    Set PtTrasnHistory = New Recordset
    PtTrasnHistory.Open sSql, DBConn

End Function

Public Function PtXmResultHistory(ByVal WorkArea As String, ByVal accdt As String, ByVal accseq As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT a.bldsrc,a.bldyy,a.bldno,a.rstv,a.stat,a.step1,a.step2,a.step3,a.step4,a.rstseq," & _
         " a.vfydt,a.vfyid,a.spcyy,a.spcno,a.cancelfg,a.cancelid,canceldt,a.statid,a.statdt," & _
         " b.abo,b.rh,b.volumn,c.componm,c.abbrnm,b.irrfg,b.irrdt,b.realexpdt,expid,exprsncd,expbilldiv,b.compocd" & _
         " FROM " & T_BBS006 & " c," & T_BBS401 & " b," & T_BBS302 & " a" & _
         " WHERE " & _
                   DBW("a.workarea=", C_WORKAREA) & _
         " AND " & DBW("a.accdt=", accdt) & _
         " AND " & DBW("a.accseq=", accseq) & _
         " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd" & _
         " AND a.compocd=c.compocd"
         
    Set PtXmResultHistory = New Recordset
    PtXmResultHistory.Open sSql, DBConn

End Function
Public Function GetBloodStatus(ByVal WorkArea As String, ByVal accdt As String, ByVal accseq As String, ByVal RstSeq As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_BBS402 & _
           " WHERE " & DBW("workarea=", WorkArea) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq) & _
           " AND " & DBW("rstseq=", RstSeq)
    
    Set GetBloodStatus = New Recordset
    GetBloodStatus.Open sSql, DBConn
End Function




Public Function PtXMSpcNum(ByVal WorkArea As String, ByVal accdt As String, ByVal accseq As String)
    Dim Rs  As Recordset
    Dim sSql As String
    
    sSql = " SELECT distinct spcyy,spcno FROM " & T_BBS302 & _
           " WHERE " & _
                     DBW("workarea=", WorkArea) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        PtXMSpcNum = Rs.Fields("spcyy").Value & "" & COL_DIV & Rs.Fields("accseq").Value & ""
    Else
        PtXMSpcNum = ""
    End If
    Set Rs = Nothing
End Function

Public Function ExpRsnName(ByVal rsncd As String) As String
    Dim Rs   As Recordset
    Dim sSql As String
    'BC2_EXP_RESON
    sSql = " SELECT field1 FROM " & T_COM003 & " WHERE " & DBW("cdindex=", BC2_EXP_RESON) & " AND " & DBW("cdval1=", rsncd)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        ExpRsnName = Rs.Fields("field1").Value & ""
    Else
        ExpRsnName = ""
    End If
    Set Rs = Nothing
    
End Function

'혈액의 존재여부체크(입고)
Public Function BloodExistChk(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Boolean
    Dim Rs As Recordset
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_BBS401 & _
           " WHERE " & _
                     DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy =", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd)
         
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.RecordCount > 0 Then
        BloodExistChk = True
    End If
    Set Rs = Nothing
End Function


Private Function SexCheck(ByVal ssn As String) As String
    Dim strTmp  As String
    Dim strDOB  As String
    Dim strYY   As String
    Dim strMM   As String
    Dim strDD   As String
    
    strYY = Mid(ssn, 1, 2)
    strMM = Mid(ssn, 3, 2)
    strDD = Mid(ssn, 5, 2)
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    If ssn <> "" Then
        strTmp = Mid(ssn, 7, 1)
        Select Case strTmp
            Case "0":  strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1":  strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2":  strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3":  strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4":  strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        If Len(ssn) = 13 Then
            If strDOB <> "" Then
                SexCheck = medFindAge(Replace(strDOB, "-", ""), "Y")
            End If
        End If
    End If
End Function
Public Function GetBarString(ByVal strBar As String, ByVal strDiv As String) As String
    Dim sSql As String
    Dim Rs   As Recordset
    
    sSql = " SELECT field2 FROM " & T_COM003 & " WHERE " & _
           DBW("cdindex=", BC2_BLOOD_BAR) & " AND " & DBW("cdval1=", strBar) & " AND " & DBW("field1=", strDiv)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    GetBarString = ""
    If Not Rs.EOF Then
        GetBarString = Rs.Fields("field2").Value & ""
    End If
    Set Rs = Nothing
End Function

Public Function GetAccDate_TransDate(ByVal accdt As String, ByVal accseq As String, ByVal AccTrans_Div As String) As String
    Dim sSql As String
    Dim Rs   As Recordset
    
    'AccTrans_Div:""(요청일시)
    If AccTrans_Div <> "" Then
        sSql = " SELECT rcvdt,rcvtm FROM " & T_LAB102 & _
               " WHERE " & DBW("workarea=", C_WORKAREA) & _
               " AND " & DBW("accdt=", accdt) & _
               " AND " & DBW("accseq=", accseq)
    Else
        sSql = " SELECT reqdt as rcvdt,reqtm as rcvtm FROM " & T_BBS204 & _
               " WHERE " & DBW("workarea=", C_WORKAREA) & _
               " AND " & DBW("accdt=", accdt) & _
               " AND " & DBW("accseq=", accseq)
    End If
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetAccDate_TransDate = Format(Rs.Fields("rcvdt").Value & "", "0###-##-##") & "   " & _
                               Format(Rs.Fields("rcvtm").Value & "", "0#:##:##")
    
    End If
End Function

Public Function GetBloodTransfusionDate(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As String
    Dim sSql As String
    Dim Rs   As Recordset
    
    sSql = " SELECT deliverydt,deliverytm FROM " & T_BBS402 & _
           " WHERE " & DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy=", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd) & _
           " ORDER BY deliverydt,deliverytm"
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetBloodTransfusionDate = Format(Rs.Fields("deliverydt").Value & "", "0###-##-##") & "   " & _
                                  Format(Rs.Fields("deliverytm").Value & "", "0#:##:##")
    
    End If
End Function

Public Function GetBldSrcList(Optional fg As Integer = 0) As String
    GetBldSrcList = " SELECT cdval1 ,field1 " & _
                    " FROM " & T_COM003 & " " & _
                    " WHERE " & DBW("cdindex=", BC2_BLD_SRC)
End Function


Public Function GetActiveBldSrc(code As String, name As String, Optional ByVal dt As String = "") As Boolean
    Dim sSql As String
    Dim Rs As Recordset
    
    If dt = "" Then dt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    sSql = "SELECT * FROM " & T_COM003 & " " & _
           " WHERE " & DBW("cdindex=", BC2_ACTIVE_BLD_SRC) & " " & _
           " AND   cdval1 = (SELECT max(cdval1) FROM " & T_COM003 & _
                           " WHERE " & DBW("cdindex=", BC2_ACTIVE_BLD_SRC) & _
                           " AND " & DBW("cdval1<", dt) & ")"
           
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF Then
'        dbconn.DisplayErrors
        code = ""
        name = ""
        Exit Function
    End If
    
    ' 혈액원코드 ------------------------------------------------
    code = Rs.Fields("field1").Value & ""
    
    sSql = "SELECT * FROM " & T_COM003 & " " & _
           "WHERE " & DBW("cdindex=", BC2_BLD_SRC) & " " & _
           "AND   " & DBW("cdval1=", code)
    
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Rs.EOF Then
'        dbconn.DisplayErrors
        name = ""
        Exit Function
    End If
    
    name = Rs.Fields("field1").Value & ""
    
    Set Rs = Nothing
End Function

Public Function GetBldSrcNm(ByVal code As String) As String
    Dim Rs      As Recordset
    Dim sSql    As String
    
    sSql = "SELECT * FROM " & T_COM003 & " " & _
           "WHERE " & DBW("cdindex=", BC2_BLD_SRC) & " " & _
           "AND   " & DBW("cdval1=", code)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Rs.EOF Then
        GetBldSrcNm = ""
    Else
        If Rs.RecordCount > 0 Then
            GetBldSrcNm = Rs.Fields("field1").Value & ""
        Else
            GetBldSrcNm = ""
        End If
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

