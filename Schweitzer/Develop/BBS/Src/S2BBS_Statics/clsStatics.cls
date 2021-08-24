VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCentercd As String
Private mvarDept As String
Private mvarWard As String
Private mvarDoct As String
Public Property Let Doct(ByVal vData As String)
    mvarDoct = vData
End Property
Public Property Let Dept(ByVal vData As String)
    mvarDept = vData
End Property
Public Property Let Ward(ByVal vData As String)
    mvarWard = vData
End Property
Public Property Let Centercd(ByVal vData As String)
    mvarCentercd = vData
End Property

Public Function Get_GroupCompo() As Recordset
'---------------------
'GROUP 코드를 불러온다
'---------------------
    Dim SSQL As String
    
    
    SSQL = " SELECT cdval1,field1 FROM " & T_COM003 & _
           " WHERE " & DBW("cdindex", BC2_COMPO_GROUP, 2) & _
           " order by field2"
    Set Get_GroupCompo = New Recordset
    Get_GroupCompo.Open SSQL, DBConn
    

End Function


Public Function Get_DonorHistory(ByVal FromDt As String, ByVal ToDt As String) As Recordset
'--------------------
'헌혈자 경력서 조회서(헌혈한사람에 한해서 조회)
'--------------------

    Dim SSQL As String
    

        
    SSQL = " SELECT a.donorid,a.donornm,a.ssn,a.dob,a.sex,a.jobcd,a.abo,a.rh,a.addr1,a.addr2,a.cnt," & _
           " b.bldsrc,b.bldyy,b.bldno,b.donationdt,b.volumn,c.okdiv3,c.okdt3,d.realexpdt,d.exprsncd,e.field1 " & _
           " FROM " & T_BBS601 & " a," & T_BBS602 & " b," & T_BBS603 & " c," & T_BBS401 & " d," & T_COM003 & " e" & _
           " WHERE " & _
           " b.donationdt between" & DBV("donationdt", FromDt) & " AND " & DBV("donationdt", ToDt) & _
           " AND a.donorid=b.donorid" & _
           " AND " & DBJ("b.donorid*=c.donorid") & _
           " AND " & DBJ("b.donoraccdt*=c.donoraccdt") & _
           " AND b.bldsrc=d.bldsrc AND b.bldyy=d.bldyy AND b.bldno=d.bldno AND b.compocd=d.compocd" & _
           " AND " & DBJ("a.jobcd*=e.cdval1") & _
           " AND " & DBW("e.cdindex", BC2_JOB, 2)
    Set Get_DonorHistory = New Recordset
    Get_DonorHistory.Open SSQL, DBConn
End Function


Public Function Get_DonorPaper(ByVal Fdt As String, ByVal Tdt As String) As Recordset
'----------------------------------------------------------
'헌혈자대장
'헌혈자마스터,헌혈자내역,적격여부 판정내역,혈액입고내역
'----------------------------------------------------------

    Dim SSQL As String

    SSQL = " SELECT a.donorid,a.donornm,a.ssn,a.dob,a.sex,a.jobcd,d.abo,d.rh,a.addr1,a.addr2,a.cnt," & _
           " d.bldsrc,d.bldyy,d.bldno,b.donationdt,b.volumn,c.okdt3,d.realexpdt,d.exprsncd,e.field1 " & _
           " FROM " & T_BBS601 & " a," & T_BBS602 & " b," & T_BBS603 & " c," & T_BBS401 & " d," & T_COM003 & " e" & _
           " WHERE " & _
           " b.donationdt between" & DBV("donationdt", Fdt) & " AND " & DBV("donationdt", Tdt) & _
           " AND a.donorid=b.donorid" & _
           " AND " & DBJ("b.donorid*=c.donorid") & _
           " AND " & DBJ("b.donoraccdt*=c.donoraccdt") & _
           " AND b.donorid=d.donorid AND b.donoraccdt=d.donoraccdt" & _
           " AND " & DBJ("a.jobcd*=e.cdval1") & _
           " AND " & DBW("e.cdindex", BC2_JOB, 2) & " order by d.bldsrc,b.bldyy,b.bldno "
    
    Set Get_DonorPaper = New Recordset
    Get_DonorPaper.Open SSQL, DBConn
    
End Function


Public Function Get_RsnNm(ByVal Rsncd As String) As String
'---------------------
'헌혈자대장의 폐기사유
'---------------------
    Dim SSQL As String
    Dim RS   As New Recordset
    
    SSQL = " SELECT field1 FROM " & T_COM003 & _
           " WHERE " & DBW("cdindex", BC2_EXP_RESON, 2) & _
           " AND " & DBW("cdval1", Rsncd, 2)
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Get_RsnNm = RS.Fields("field1")
    Else
        Get_RsnNm = ""
    End If
    Set RS = Nothing
End Function


Public Function Get_DonorRecord(ByVal Fdt As String, ByVal Tdt As String) As Recordset
'----------------------------------
'헌혈실적통지서의 헌혈지원자 구하기
'----------------------------------
    Dim SSQL       As String
    Dim RecordDate As Date
    Dim YearDt     As String
    
    RecordDate = DateAdd("M", -1, Format(Fdt, "####-##-##"))
    
    YearDt = Format(RecordDate, "yyyymm") & "31"
    '남/녀 합계
    SSQL = " SELECT a.sex as sex,count(a.sex) as cnt, 'sexdiv' as div FROM " & T_BBS601 & " a," & T_BBS602 & " b" & _
           " WHERE b.donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & _
           " AND a.donorid=b.donorid GROUP BY a.sex"
    '월계
    SSQL = SSQL & " UNION ALL" & _
                 " SELECT 'sex' as sex,count(*) as cnt, 'M' as div FROM " & T_BBS602 & _
                 " WHERE donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt)
    '누계=월계+이전 월계
    SSQL = SSQL & " UNION ALL" & _
                " SELECT 'sex' as sex,count(*) as cnt, 'Y' as div FROM " & T_BBS602 & _
                " WHERE " & DBW("donoraccdt<", YearDt, 2)
    Set Get_DonorRecord = New Recordset
    Get_DonorRecord.Open SSQL, DBConn
End Function


Public Function WholeBloodCnt(ByVal Fdt As String, ByVal Tdt As String) As Recordset
'------------------------------
'헌혈실적통지서의 전혈자 구하기
'------------------------------
    Dim SSQL As String
    Dim RecordDate As Date
    Dim YearDt     As String
    
    RecordDate = DateAdd("M", -1, Format(Fdt, "####-##-##"))
    YearDt = Format(RecordDate, "yyyymm") & "31"
    
    
    '성분헌혈(남/여)
    SSQL = " SELECT a.sex,b.volumn,count(b.volumn) as cnt,'sexdiv' as div FROM " & T_BBS601 & " a," & T_BBS602 & " b" & _
           " WHERE b.donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & _
           " AND b.volumn in(320,400)" & _
           " AND " & DBW("b.cancelfg=", "0") & _
           " AND a.donorid=b.donorid GROUP BY a.sex,b.volumn"
    '성분헌혈 (월계)
    SSQL = SSQL & " UNION ALL" & _
                      " SELECT 'sex' as sex ,volumn,count(volumn) as cnt ,'M' as div FROM " & T_BBS602 & _
                      " WHERE donoraccdt  between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & _
                      " AND volumn in (320,400)" & _
                      " AND " & DBW("cancelfg=", "0") & _
                      " GROUP BY volumn"
    '성분헌혈(누계)
    SSQL = SSQL & " UNION ALL" & _
                      " SELECT 'sex' as sex, volumn,count(volumn) as cnt ,'Y' as div FROM " & T_BBS602 & _
                      " WHERE " & DBW("donoraccdt<=", YearDt) & " " & _
                      " AND volumn in (320,400)" & _
                      " AND " & DBW("cancelfg=", "0") & _
                      " GROUP BY volumn"
        
    Set WholeBloodCnt = New Recordset
    WholeBloodCnt.Open SSQL, DBConn
End Function


Public Function PheresisCnt(ByVal Fdt As String, ByVal Tdt As String) As Recordset
'----------------------
'Pheresis 헌혈자 구하기
'----------------------
    Dim SSQL As String
    Dim RecordDate As Date
    Dim YearDt     As String
    
    RecordDate = DateAdd("M", -1, Format(Fdt, "####-##-##"))
    YearDt = Format(RecordDate, "yyyymm") & "31"
    '성분헌혈(남/여)
    SSQL = " SELECT a.sex as sex,b.volumn,b.compocd,count(b.volumn) as cnt,'sexdiv' as div " & _
           " FROM " & _
                    T_BBS601 & " a," & T_BBS602 & " b," & T_BBS006 & " c" & _
           " WHERE " & _
                     "     b.donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & _
                     " AND b.compocd=c.compocd" & _
                     " AND " & DBW("c.pherefg=", "1") & _
                     " AND not b.volumn in(320,400)" & _
                     " AND " & DBW("b.cancelfg", "0", 2) & _
                     " AND a.donorid=b.donorid" & _
                     " GROUP BY a.sex,b.volumn,b.compocd"
    '성분헌혈 (월계)
    SSQL = SSQL & " UNION ALL" & _
                      " SELECT 'sex' as sex,a.volumn,a.compocd,count(a.volumn) as cnt ,'M' as div " & _
                      " FROM " & _
                               T_BBS602 & " a," & T_BBS006 & " b" & _
                      " WHERE " & _
                                "     a.donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & _
                                " AND a.compocd=b.compocd" & _
                                " AND " & DBW("b.pherefg=", "1") & _
                                " AND not a.volumn in (320,400)" & _
                                " AND " & DBW("a.cancelfg", "0", 2) & _
                                " GROUP BY a.volumn,a.compocd"
    '성분헌혈(누계)
    SSQL = SSQL & " UNION ALL" & _
                      " SELECT 'sex' as sex,a.volumn,a.compocd,count(a.volumn) as cnt ,'Y' as div " & _
                      " FROM " & T_BBS602 & " a," & T_BBS006 & " b" & _
                      " WHERE " & _
                                          DBW("a.donoraccdt<=", YearDt) & _
                                " AND a.compocd=b.compocd" & _
                                " AND " & DBW("b.pherefg=", "1") & _
                                " AND not a.volumn in (320,400)" & _
                                " AND " & DBW("a.cancelfg", "0", 2) & _
                                " GROUP BY a.volumn,a.compocd"
    Set PheresisCnt = New Recordset
    PheresisCnt.Open SSQL, DBConn
End Function


Public Function PheresisList() As Recordset
'--------------------
'성분헌혈 목록 구하기
'--------------------
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_COM003 & _
          " WHERE " & DBW("cdindex", BC2_PHERESIS_COMPO, 2)
    
    Set PheresisList = New Recordset
    PheresisList.Open SSQL, DBConn
End Function


Public Function DonorPaperCnt() As Long
'---------------------
'헌혈증서 재고량구하기
'---------------------
    Dim RS   As New Recordset
    Dim SSQL As String
    
    SSQL = " SELECT count(*)as cnt FROM " & T_BBS411 & _
          " WHERE (usedt='' or usedt is null ) " & _
          " and (returndt='' or returndt is null) "
    
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        DonorPaperCnt = RS.Fields("cnt").Value & ""
    Else
        DonorPaperCnt = 0
    End If
    Set RS = Nothing
End Function


Public Function Get_NotDonorList(ByVal Fdt As String, ByVal Tdt As String, ByVal Compocd As String, ByVal Vol As String) As Recordset
'-----------------------
'부적격 혈액 리스트 조회
'-----------------------
    Dim SSQL As String
    
    SSQL = " SELECT a.donorid,a.donoraccdt,d.bldsrc,d.bldyy,d.bldno,d.compocd,c.abbrnm,d.volumn" & _
           " FROM " & T_BBS401 & " d," & T_BBS006 & " c," & T_BBS603 & " b," & T_BBS602 & " a" & _
           " WHERE " & _
           " a.donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & "" & _
           " AND " & DBW("b.stscd>=", DonorStatus.stsDonation) & _
           " AND " & DBW("b.okdiv3=", "0")
           
    If Compocd <> "" Then
        SSQL = SSQL & " AND " & DBW("d.compocd=", Compocd)
    End If
    
    Select Case Vol
        Case "320", "400"
            SSQL = SSQL & " AND " & DBW("d.volumn=", Vol)
        Case "Etc"
            SSQL = SSQL & " AND d.volumn not in ('320','400')"
    End Select
    
    SSQL = SSQL & _
           " AND a.donorid=b.donorid" & _
           " AND a.donoraccdt=b.donoraccdt" & _
           " AND a.compocd=c.compocd" & _
           " AND a.donoraccdt=d.donoraccdt " & _
           " AND a.donorid=d.donorid" & _
           " order by a.bldsrc,a.bldyy,a.bldno"
         
         
    'ssql = " SELECT distinct a.donorid,a.donoraccdt, a.bldsrc,a.bldyy,a.bldno," & _
          "  a.compocd,b.abbrnm " & _
          " FROM " & T_BBS602 & " a," & T_BBS006 & " b," & T_BBS607 & " c" & _
          " WHERE a.donoraccdt between " & DBV("donoraccdt", Fdt) & " AND " & DBV("donoraccdt", Tdt) & "" & _
          " AND a.compocd=b.compocd" & _
          " AND a.donorid=c.donorid" & _
          " AND a.donoraccdt=c.donoraccdt  order by a.bldsrc,a.bldyy,a.bldno"
    
    Set Get_NotDonorList = New Recordset
    Get_NotDonorList.Open SSQL, DBConn
    
End Function


Public Function Get_NotDonorRsncd(ByVal donorid As String, ByVal donoraccdt As String) As String
'------------------------
'부적격 사유코드 가져오기
'------------------------
    Dim RS   As New Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_BBS607 & _
          " WHERE " & DBW("donorid", donorid, 2) & _
          " AND " & DBW("donoraccdt", donoraccdt, 2)
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            Get_NotDonorRsncd = Get_NotDonorRsncd & RS.Fields("rsncd").Value & "" & COL_DIV
            RS.MoveNext
        Loop
    End If
    
    Get_NotDonorRsncd = Mid(Get_NotDonorRsncd, 1, Len(Get_NotDonorRsncd) - 1)
    
    Set RS = Nothing
End Function


Public Function NotAcceptRsncd() As Recordset
'-------------------------------
'부적격 사유코드 마스터에 있는거
'-------------------------------
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_COM003 & _
          " WHERE " & DBW("cdindex", BC2_UNACCEPT_REASON, 2)
    
    Set NotAcceptRsncd = New Recordset
    NotAcceptRsncd.Open SSQL, DBConn
End Function


Public Function Get_NotdonorVdt(ByVal donorid As String, ByVal donoraccdt As String) As String
'------------------------------
'부적격 리스트의 검사일자(최종)
'------------------------------
    Dim RS   As New Recordset
    Dim SSQL As String
    
    SSQL = " SELECT distinct max(orddt) as orddt FROM " & T_BBS605 & _
          " WHERE " & DBW("donorid", donorid, 2) & _
          " AND " & DBW("donoraccdt", donoraccdt, 2)
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Get_NotdonorVdt = Format(RS.Fields("orddt").Value & "", "####-##-##")
    Else
        Get_NotdonorVdt = ""
    End If
    
    Set RS = Nothing
End Function


Public Function bloodcnt(ByVal Fdt As String, ByVal Tdt As String, ByVal mode As Long) As Boolean
 '---------------------------------------------
'혈액현황별로 갯수 구하기(입고,출고,반환,폐기)
'---------------------------------------------
   Dim RS     As New Recordset
    Dim SSQL   As String
    Dim status As String
    
    Select Case mode
        Case 1: status = "01"
        Case 2: status = "02"
        Case 3: status = "03"
        Case 4: status = "04"
    End Select
    SSQL = "SELECT * FROM " & T_BBS901 & _
           " WHERE closedt between " & DBV("closedt", Fdt) & " AND " & DBV("closedt", Tdt) & _
           " AND " & DBW("bloodtype=", status)
           
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        bloodcnt = True
    End If
    Set RS = Nothing
End Function


Public Function BloodCondition(ByVal Fdt As String, ByVal Tdt As String, ByVal center As String, _
                               ByVal ABO As String, ByVal Rh As String, ByVal Volume As String, _
                               ByVal Phere As String, ByVal mode As Long) As Recordset
'-----------------------------------------------
'입고형태별로 혈액현황구하기
'pheresis 인경우의 카운트는 모든 센터이거나,중앙
'pheresis 인경우의 카운트는 용량이 320이거나 ALL
'-----------------------------------------------
    Dim SSQL   As String
    Dim status As String
    
    Select Case mode
        Case 1: status = "01"
        Case 2: status = "02"
        Case 3: status = "03"
        Case 4: status = "04"
    End Select
    '--------------------------------------------
    '혈액원입고
    '--------------------------------------------
    SSQL = " SELECT a.closedt,a.bloodcnt as cnt,b.groupcd,'org' as div FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
           " WHERE a.closedt between '" & Fdt & "' AND '" & Tdt & "' "
    Select Case mode
        Case 1: SSQL = SSQL & " AND a.bloodtype in('01','06')"
        Case 2: SSQL = SSQL & " AND a.bloodtype in('02','05')"
        Case Else: SSQL = SSQL & " AND " & DBW("a.bloodtype", status, 2)
    End Select
        
    If center <> "ALL" Then
        SSQL = SSQL & " AND " & DBW("a.centercd", center, 2)
    End If
    
    If ABO <> "ALL" Then
        SSQL = SSQL & " AND " & DBW("a.abo", ABO, 2)
    End If
    If Rh <> "ALL" Then
        SSQL = SSQL & " AND " & DBW("a.rh", Rh, 2)
    End If
    
    
    Select Case Volume
        Case "320":  SSQL = SSQL & " AND " & DBW("a.volumn", Volume, 2)
        Case "400":  SSQL = SSQL & " AND " & DBW("a.volumn", Volume, 2)
        Case "기타": SSQL = SSQL & " AND not a.volumn in(320,400)"
    End Select
    
    If center = "ALL" Or Volume = "ALL" Or Volume = "320" Then
        SSQL = SSQL & " AND b.pherefg in ('0','1') AND a.compocd=b.compocd"
    Else
        SSQL = SSQL & " AND " & DBW("b.pherefg", "0", 2) & " AND a.compocd=b.compocd"
    End If
    '--------------------------------------------
    '헌혈입고
    '--------------------------------------------
    If DonorUserFg = True Then
        SSQL = SSQL & " UNION ALL" & _
                      " SELECT a.closedt,a.donationcnt as cnt,b.groupcd,'donor' as div FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
                      " WHERE a.closedt between " & DBV("closedt", Fdt) & " AND " & DBV("closedt", Tdt)
                      
        Select Case mode
            Case 1: SSQL = SSQL & " AND a.bloodtype in('01','06')"
            Case 2: SSQL = SSQL & " AND a.bloodtype in('02','05')"
            Case Else: SSQL = SSQL & " AND " & DBW("a.bloodtype", status, 2)
        End Select
        
        If center <> "ALL" Then
            SSQL = SSQL & " AND " & DBW("a.centercd", center, 2)
        End If
        
        If ABO <> "ALL" Then
            SSQL = SSQL & " AND " & DBW("a.abo", ABO, 2)
        End If
        If Rh <> "ALL" Then
            SSQL = SSQL & " AND " & DBW("a.rh", Rh, 2)
        End If
        
        Select Case Volume
            Case "320":  SSQL = SSQL & " AND " & DBW("a.volumn", Volume, 2)
            Case "400":  SSQL = SSQL & " AND " & DBW("a.volumn", Volume, 2)
            Case "기타": SSQL = SSQL & " AND not a.volumn in(320,400)"
        End Select
        
        If center = "ALL" Or Volume = "ALL" Or Volume = "320" Then
            SSQL = SSQL & " AND b.pherefg in ('0','1') AND a.compocd=b.compocd"
        Else
            SSQL = SSQL & " AND " & DBW("b.pherefg", "0", 2) & " AND a.compocd=b.compocd"
        End If
    End If
    '--------------------------------------------
    'Local 병원입고
    '--------------------------------------------
    SSQL = SSQL & " UNION ALL" & _
                  " SELECT a.closedt,a.localcnt as cnt,b.groupcd,'local' as div FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
                  " WHERE a.closedt between '" & Fdt & "' AND '" & Tdt & "' "
    
    Select Case mode
        Case 1: SSQL = SSQL & " AND a.bloodtype in('01','06')"
        Case 2: SSQL = SSQL & " AND a.bloodtype in('02','05')"
        Case Else: SSQL = SSQL & " AND " & DBW("a.bloodtype", status, 2)
    End Select
    
    If center <> "ALL" Then
        SSQL = SSQL & " AND " & DBW("a.centercd", center, 2)
    End If
    
    If ABO <> "ALL" Then
        SSQL = SSQL & " AND " & DBW("a.abo", ABO, 2)
    End If
    If Rh <> "ALL" Then
        SSQL = SSQL & " AND " & DBW("a.rh", Rh, 2)
    End If
    
    Select Case Volume
        Case "320":  SSQL = SSQL & " AND " & DBW("a.volumn", Volume, 2)
        Case "400":  SSQL = SSQL & " AND " & DBW("a.volumn", Volume, 2)
        Case "기타": SSQL = SSQL & " AND not a.volumn in(320,400)"
    End Select
    
    If center = "ALL" Or Volume = "ALL" Or Volume = "320" Then
        SSQL = SSQL & " AND b.pherefg in ('0','1') AND a.compocd=b.compocd"
    Else
        SSQL = SSQL & " AND " & DBW("b.pherefg", "0", 2) & " AND a.compocd=b.compocd"
    End If
    
    Set BloodCondition = New Recordset
    BloodCondition.Open SSQL, DBConn
    
End Function


Public Function Bld_Dic(ByVal RS As Recordset) As clsDictionary
'-------------------------------------------------------------
'입고형태별로 혈액현황을 날짜별로 Dictionary에 담아서 넘겨준다
'-------------------------------------------------------------
    Dim GRs  As New Recordset
    Dim SSQL As String
    Dim seq  As Long
    Dim ii   As Long
    Dim cnt  As Long
    
    Set Bld_Dic = New clsDictionary
    Bld_Dic.Clear
    Bld_Dic.FieldInialize "seq,gorupcd", "day1,day2,day3,day4,day5,day6,day7,day8,day9,day10,day11," & _
                                        "day12,day13,day14,day15,day16,day17,day18,day19,day20,day21," & _
                                        "day22,day23,day24,day25,day26,day27,day28,day29,day30,day31,tot"
    Set GRs = Get_GroupCompo
    
    For ii = 1 To 3
        Do Until GRs.EOF
            Bld_Dic.AddNew Join(Array(ii, GRs.Fields("cdval1").Value & ""), COL_DIV), Join(Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
            GRs.MoveNext
        Loop
        GRs.MoveFirst
    Next
    
    
    
    With RS
        Do Until .EOF
            Select Case .Fields("div").Value & ""
                Case "org":   seq = 1 ': cnt = .Fields("bloodcnt")
                Case "donor": seq = 2 ': cnt = .Fields("donationcnt")
                Case "local": seq = 3 ': cnt = .Fields("localcnt")
            End Select
            cnt = .Fields("cnt").Value & ""
            If Bld_Dic.Exists(Join(Array(seq, .Fields("groupcd").Value & ""), COL_DIV)) Then
                Bld_Dic.KeyChange Join(Array(seq, .Fields("groupcd").Value & ""), COL_DIV)
                Select Case Mid(.Fields("closedt").Value & "", 7)
                    Case "01"
                        Bld_Dic.Fields("day1") = Bld_Dic.Fields("day1") + cnt
                    Case "02"
                        Bld_Dic.Fields("day2") = Bld_Dic.Fields("day2") + cnt
                    Case "03"
                        Bld_Dic.Fields("day3") = Bld_Dic.Fields("day3") + cnt
                    Case "04"
                        Bld_Dic.Fields("day4") = Bld_Dic.Fields("day4") + cnt
                    Case "05"
                        Bld_Dic.Fields("day5") = Bld_Dic.Fields("day5") + cnt
                    Case "06"
                        Bld_Dic.Fields("day6") = Bld_Dic.Fields("day6") + cnt
                    Case "07"
                        Bld_Dic.Fields("day7") = Bld_Dic.Fields("day7") + cnt
                    Case "08"
                        Bld_Dic.Fields("day8") = Bld_Dic.Fields("day8") + cnt
                    Case "09"
                        Bld_Dic.Fields("day9") = Bld_Dic.Fields("day9") + cnt
                    Case "10"
                        Bld_Dic.Fields("day10") = Bld_Dic.Fields("day10") + cnt
                    Case "11"
                        Bld_Dic.Fields("day11") = Bld_Dic.Fields("day11") + cnt
                    Case "12"
                        Bld_Dic.Fields("day12") = Bld_Dic.Fields("day12") + cnt
                    Case "13"
                        Bld_Dic.Fields("day13") = Bld_Dic.Fields("day13") + cnt
                    Case "14"
                        Bld_Dic.Fields("day14") = Bld_Dic.Fields("day14") + cnt
                    Case "15"
                        Bld_Dic.Fields("day15") = Bld_Dic.Fields("day15") + cnt
                    Case "16"
                        Bld_Dic.Fields("day16") = Bld_Dic.Fields("day16") + cnt
                    Case "17"
                        Bld_Dic.Fields("day17") = Bld_Dic.Fields("day17") + cnt
                    Case "18"
                        Bld_Dic.Fields("day18") = Bld_Dic.Fields("day18") + cnt
                    Case "19"
                        Bld_Dic.Fields("day19") = Bld_Dic.Fields("day19") + cnt
                    Case "20"
                        Bld_Dic.Fields("day20") = Bld_Dic.Fields("day20") + cnt
                    Case "21"
                        Bld_Dic.Fields("day21") = Bld_Dic.Fields("day21") + cnt
                    Case "22"
                        Bld_Dic.Fields("day22") = Bld_Dic.Fields("day22") + cnt
                    Case "23"
                        Bld_Dic.Fields("day23") = Bld_Dic.Fields("day23") + cnt
                    Case "24"
                        Bld_Dic.Fields("day24") = Bld_Dic.Fields("day24") + cnt
                    Case "25"
                        Bld_Dic.Fields("day25") = Bld_Dic.Fields("day25") + cnt
                    Case "26"
                        Bld_Dic.Fields("day26") = Bld_Dic.Fields("day26") + cnt
                    Case "27"
                        Bld_Dic.Fields("day27") = Bld_Dic.Fields("day27") + cnt
                    Case "28"
                        Bld_Dic.Fields("day28") = Bld_Dic.Fields("day28") + cnt
                    Case "29"
                        Bld_Dic.Fields("day29") = Bld_Dic.Fields("day29") + cnt
                    Case "30"
                        Bld_Dic.Fields("day30") = Bld_Dic.Fields("day30") + cnt
                    Case "31"
                        Bld_Dic.Fields("day31") = Bld_Dic.Fields("day31") + cnt
                End Select
                Bld_Dic.Fields("tot") = Bld_Dic.Fields("tot") + cnt
            End If
            .MoveNext
        Loop
    End With
End Function

'----------------------
'환자별 수혈 부작용내역
'----------------------
Public Function ReactionPtRecord(Optional ByVal PtId As String = "") As Recordset
'-------------------------------
'부작용이 발생한 환자를 불러온다
'-------------------------------
    Dim SSQL As String
    
    SSQL = " SELECT distinct b.workarea,b.accdt,b.accseq,b.ptid,c." & F_PTNM & " as ptnm" & _
          " FROM " & T_BBS501 & " a," & T_BBS202 & " b," & T_HIS001 & " c" & _
          " WHERE " & DBW("a.workarea", C_WORKAREA, 2) & _
          "   AND a.workarea=b.workarea" & _
          "   AND a.accdt=b.accdt" & _
          "   AND a.accseq=b.accseq" & _
          "   AND b.ptid=c." & F_PTID & _
          "   AND a.reactdiv='1'"
          
    If PtId <> "" Then
        SSQL = SSQL & " AND " & DBW("b.ptid", PtId, 2)
    End If
    
    Set ReactionPtRecord = New Recordset
    ReactionPtRecord.Open SSQL, DBConn
    
End Function

Public Function ReactionRecordset(ByVal Fdt As String, ByVal Tdt As String) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT distinct b.ptid,b.orddt,b.ordno,b.ordseq,c." & F_PTNM & " as ptnm,a.reactrmk" & _
          " FROM " & _
                    T_BBS501 & " a," & T_BBS202 & " b," & T_HIS001 & " c" & _
          " WHERE" & _
                    "     a.transdt between " & DBV("transdt", Fdt) & " AND " & DBV("transdt", Tdt) & " " & _
                    " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
                    " AND b.ptid=c." & F_PTID & " AND " & DBW("a.reactdiv=", "1")
    
'    If ptid <> "" Then
'        ssql = ssql & " AND b.ptid" & DBN(ptid, 2)
'    End If
    

    Set ReactionRecordset = New Recordset
    ReactionRecordset.Open SSQL, DBConn
End Function

Public Function ReactionDetail(ByVal PtId As String, ByVal orddt As String, ByVal ordno As String, ByVal ordseq As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " SELECT c.testnm,b.unitqty,a.reqdt,a.wardid " & _
         " FROM " & _
                T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c" & _
         " WHERE" & _
                "     " & DBW("b.ptid", PtId, 2) & _
                " AND " & DBW("b.orddt", orddt, 2) & _
                " AND " & DBW("b.ordno", ordno, 2) & _
                " AND " & DBW("b.ordseq", ordseq, 2) & _
                " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno" & _
                " AND b.ordcd=c.testcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        ReactionDetail = RS.Fields("testnm").Value & "" & COL_DIV & RS.Fields("unitqty").Value & "" & COL_DIV & _
                       RS.Fields("reqdt").Value & "" & COL_DIV & RS.Fields("wardid").Value & ""
    Else
        ReactionDetail = ""
    End If
    
    Set RS = Nothing
End Function

'---------------------------
'환자별 수혈내역 (frmBBS912)
'---------------------------
Public Function Pt_TransBloodRecord(ByVal PtId As String) As Recordset
    
    Dim SSQL As String
    
    SSQL = " SELECT distinct a.workarea,a.accdt,a.accseq,a.bldsrc,a.bldyy,a.bldno,a.compocd " & _
         " FROM " & _
                T_BBS402 & " a," & T_BBS202 & " b" & _
         " WHERE " & _
                 "     " & DBW("b.ptid", PtId, 2) & _
                 " AND a.workarea=b.workarea" & _
                 " AND a.accdt=b.accdt" & _
                 " AND a.accseq=b.accseq" & _
                 " AND (a.retfg='' or a.retfg is null)" & _
                 " AND (a.expfg='' or a.expfg is null)"
    
    Set Pt_TransBloodRecord = New Recordset
    Pt_TransBloodRecord.Open SSQL, DBConn

End Function
Public Function Pt_OrderRecordset(ByVal accdt As String, ByVal accseq As String) As String
    
    Dim RS   As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT a.abbrnm5,b.unitqty,b.orddt,b.ordno,b.ordseq,c.reqdt " & _
          " FROM " & _
                 T_BBS001 & " a," & T_LAB102 & " b," & T_LAB101 & " c" & _
          " WHERE " & _
                  "     " & DBW("b.workarea", C_WORKAREA, 2) & _
                  " AND " & DBW("b.accdt", accdt, 2) & _
                  " AND " & DBW("b.accseq", accseq, 2) & _
                  " AND b.ptid=c.ptid" & _
                  " AND b.orddt=c.orddt" & _
                  " AND b.ordno=c.ordno" & _
                  " AND b.ordcd=a.testcd"
                  
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Pt_OrderRecordset = RS.Fields("abbrnm5").Value & "" & COL_DIV & RS.Fields("unitqty").Value & "" & COL_DIV & _
                          RS.Fields("reqdt").Value & "" & COL_DIV & RS.Fields("orddt").Value & "" & COL_DIV & _
                          RS.Fields("ordno").Value & "" & COL_DIV & RS.Fields("ordseq").Value & ""
    Else
        Pt_OrderRecordset = ""
    End If
    
    Set RS = Nothing
End Function
Public Function Blood_DeliveryRecord(ByVal accdt As String, ByVal accseq As String) As Recordset
'-------------------------------------
'환자에대한 처방별 수혈내역을 조회한다
'-------------------------------------
    Dim SSQL As String
    
    SSQL = " SELECT distinct a.deliverydt,a.bldsrc,a.bldyy,a.bldno,b.componm,c.abo,c.rh,c.volumn," & _
                " d.transdt,d.transtm,d.transvol,d.reactdiv,d.reactrmk" & _
         " FROM " & _
                T_BBS402 & " a," & T_BBS006 & " b," & T_BBS401 & " c," & T_BBS501 & " d" & _
         " WHERE" & _
                "     " & DBW("a.workarea", C_WORKAREA, 2) & _
                " AND " & DBW("a.accdt", accdt, 2) & _
                " AND " & DBW("a.accseq", accseq, 2) & _
                " AND a.compocd=b.compocd AND (a.retfg='' or a.retfg is null) AND (a.expfg='' or a.expfg is null)" & _
                " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd" & _
                " AND a.bldsrc=d.bldsrc AND a.bldyy=d.bldyy AND a.bldno=d.bldno AND a.compocd=d.compocd"
    
    Set Blood_DeliveryRecord = New Recordset
    Blood_DeliveryRecord.Open SSQL, DBConn

End Function
Public Function Blood_TotalDeliveryRecord(ByVal PtId As String) As Recordset
'------------------------------------
'환자에 대한 전체 수혈내역을 조회한다
'------------------------------------
    Dim SSQL As String
    
    SSQL = " SELECT distinct a.deliverydt,a.bldsrc,a.bldyy,a.bldno,b.componm,c.abo,c.rh,c.volumn," & _
                " d.transdt,d.transvol,d.reactdiv,d.reactrmk" & _
         " FROM " & _
                T_BBS402 & " a," & T_BBS006 & " b," & T_BBS401 & " c," & T_BBS501 & " d," & T_BBS202 & " e" & _
         " WHERE" & _
                "     " & DBW("e.ptid", PtId, 2) & _
                " AND a.workarea=e.workarea AND a.accdt=e.accdt AND a.accseq=e.accseq" & _
                " AND a.compocd=b.compocd AND (a.retfg='' or a.retfg is null) AND (a.expfg='' or a.expfg is null)" & _
                " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd" & _
                " AND a.bldsrc=d.bldsrc AND a.bldyy=d.bldyy AND a.bldno=d.bldno AND a.compocd=d.compocd"
    
    Set Blood_TotalDeliveryRecord = New Recordset
    Blood_TotalDeliveryRecord.Open SSQL, DBConn

End Function
Public Function Trans_PtNm(ByVal PtId As String) As String
    Dim RS   As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT a.ptid,b." & F_PTNM & " as ptnm, " & F_SSN2("b") & " as ssn,d.wardid,d.deptcd" & _
           " FROM " & _
                  T_BBS402 & " c," & T_BBS202 & " a," & T_HIS001 & " b," & T_LAB101 & " d," & T_LAB102 & " e" & _
           " WHERE" & _
                  "    " & DBW("a.ptid", PtId, 2) & _
                  " AND a.ptid=b." & F_PTID & _
                  " AND c.workarea=a.workarea" & _
                  " AND c.accdt=a.accdt" & _
                  " AND c.accseq=a.accseq AND (c.retfg='' or c.retfg is null) AND (c.expfg='' or c.expfg is null)" & _
                  " AND a.workarea=e.workarea AND a.accdt=e.accdt AND a.accseq=e.accseq" & _
                  " AND d.ptid=e.ptid AND d.orddt=e.orddt AND d.ordno=e.ordno"
                  
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Trans_PtNm = RS.Fields("ptnm").Value & "" & COL_DIV & RS.Fields("ssn").Value & "" & COL_DIV & RS.Fields("wardid").Value & "" & COL_DIV & RS.Fields("deptcd").Value & ""
    End If
    Set RS = Nothing
    
End Function
Public Function Trans_PtRecordSet() As String
    
    Trans_PtRecordSet = " SELECT distinct a.ptid,b." & F_PTNM & " as ptnm" & _
                        " FROM " & _
                               T_BBS402 & " c," & T_BBS202 & " a," & T_HIS001 & " b" & _
                        " WHERE" & _
                               "  a.ptid=b." & F_PTID & _
                               " AND c.workarea=a.workarea" & _
                               " AND c.accdt=a.accdt" & _
                               " AND c.accseq=a.accseq" & _
                               " AND (c.retfg='' or c.retfg is null) AND (c.expfg='' or c.expfg is null)"

End Function


'-------------------
'혈액일보(frmBBS913)
'-------------------
'혈액타입1:입고,2:출고,3:반환,4:폐기,5:분획출고,6:분획입고,7:Xm건수,8:Assign
Public Function Get_BloodDayCount(ByVal Fdt As String, ByVal Tdt As String) As clsDictionary
    Dim objDic As clsDictionary
    Dim RS   As Recordset
    Dim SSQL As String
    Dim cnt  As Long
    
    Set Get_BloodDayCount = New clsDictionary
    Set objDic = New clsDictionary
    
    objDic.Clear
    objDic.FieldInialize "closedt", "enter,splitin,xm,assign,deliver,splitout,return,expire,reaction"
    
    '입고
    SSQL = " SELECT entdt,count(*) as cnt,'1' as div FROM " & T_BBS401 & _
           " WHERE " & DBW("entdt>=", Fdt) & " AND " & DBW("entdt<=", Tdt) & " AND (splitinfg<>'1') " & _
           " GROUP BY entdt"
    '분획입고
    
    SSQL = SSQL & " UNION ALL " & " SELECT entdt,count(*) as cnt,'2' as div  FROM " & T_BBS401 & _
           " WHERE " & DBW("entdt>=", Fdt) & " AND " & DBW("entdt<=", Tdt) & " AND (splitinfg='1') " & _
           " GROUP BY entdt"
    '검사
    SSQL = SSQL & " UNION ALL " & " SELECT vfydt as entdt ,count(*) as cnt,'3' as div  FROM " & T_BBS302 & _
           " WHERE " & DBW("vfydt>=", Fdt) & " AND " & DBW("vfydt<=", Tdt) & _
           " GROUP BY vfydt"
    '검사 ok
    SSQL = SSQL & " UNION ALL " & " SELECT vfydt as entdt ,count(*) as cnt ,'4' as div FROM " & T_BBS302 & _
           " WHERE " & DBW("vfydt>=", Fdt) & " AND " & DBW("vfydt<=", Tdt) & " AND " & DBW("rstv=", "1") & _
           " GROUP BY vfydt"
    '출고
    SSQL = SSQL & " UNION ALL " & " SELECT deliverydt as entdt,count(*) as cnt,'5' as div  FROM " & T_BBS402 & _
           " WHERE " & DBW("deliverydt>=", Fdt) & " AND " & DBW("deliverydt<=", Tdt) & _
           " GROUP BY deliverydt"
    '분획출고
    SSQL = SSQL & " UNION ALL " & " SELECT a.splitoutdt as entdt,count(*) as cnt,'6' as div FROM " & T_BBS401 & " a" & _
           " WHERE " & DBW("a.splitoutdt>=", Fdt) & " AND " & DBW("a.splitoutdt<=", Tdt) & " AND " & DBW("a.splitoutfg=", "1") & _
           " " & _
           " GROUP BY a.splitoutdt"
    '반환
    SSQL = SSQL & " UNION ALL " & " SELECT retdt as entdt,count(*) as cnt,'7' as div  FROM " & T_BBS402 & _
           " WHERE " & DBW("retdt>=", Fdt) & " AND " & DBW("retdt<=", Tdt) & _
           " GROUP BY retdt"
    '폐기
    SSQL = SSQL & " UNION ALL " & " SELECT realexpdt as entdt,count(*) as cnt,'8' as div  FROM " & T_BBS401 & _
           " WHERE " & DBW("realexpdt>=", Fdt) & " AND " & DBW("realexpdt<=", Tdt) & _
           " GROUP BY realexpdt"
    
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    
    If Not RS.EOF Then
        Do Until RS.EOF
            cnt = Val(RS.Fields("cnt").Value & "")
            If objDic.Exists(RS.Fields("entdt").Value & "") Then
                
                objDic.KeyChange RS.Fields("entdt").Value & ""
                Select Case RS.Fields("div").Value & ""
                    Case "1": objDic.Fields("enter") = objDic.Fields("enter") + cnt       '1
                    Case "2": objDic.Fields("splitin") = objDic.Fields("splitin") + cnt   '2
                    Case "3": objDic.Fields("xm") = objDic.Fields("xm") + cnt              '3
                    Case "4": objDic.Fields("assign") = objDic.Fields("assign") + cnt      '4
                    Case "5": objDic.Fields("deliver") = objDic.Fields("deliver") + cnt    '5
                    Case "6": objDic.Fields("splitout") = objDic.Fields("splitout") + cnt  '6
                    Case "7": objDic.Fields("return") = objDic.Fields("return") + cnt      '8
                    Case "8": objDic.Fields("expire") = objDic.Fields("expire") + cnt      '9
                End Select
            Else
                Select Case RS.Fields("div")
                    Case "1": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(cnt, 0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
                    Case "2": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, cnt, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
                    Case "3": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, 0, cnt, 0, 0, 0, 0, 0, 0), COL_DIV)
                    Case "4": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, 0, 0, cnt, 0, 0, 0, 0, 0), COL_DIV)
                    Case "5": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, 0, 0, 0, cnt, 0, 0, 0, 0), COL_DIV)
                    Case "6": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, 0, 0, 0, 0, cnt, 0, 0, 0), COL_DIV)
                    Case "7": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, 0, 0, 0, 0, 0, cnt, 0, 0), COL_DIV)
                    Case "8": objDic.AddNew RS.Fields("entdt").Value & "", Join(Array(0, 0, 0, 0, 0, 0, 0, cnt, 0), COL_DIV)
                End Select
            End If
            RS.MoveNext
        Loop
    End If
    
    
    Set Get_BloodDayCount = objDic

    Set RS = Nothing
    Set objDic = Nothing
End Function
Private Function Get_ReactionCount(ByVal objDic As clsDictionary, ByVal Fdt As String, ByVal Tdt As String) As clsDictionary
    Dim RS   As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT transdt FROM " & T_BBS501 & _
         " WHERE transdt between " & DBV("transdt", Fdt) & " AND " & DBV("transdt", Tdt) & _
         " AND " & DBW("reactdiv=", "1")
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        objDic.MoveFirst
        Do Until RS.EOF
            If objDic.Exists(RS.Fields("transdt").Value & "") Then
                objDic.KeyChange RS.Fields("transdt").Value & ""
                objDic.Fields("reaction") = objDic.Fields("reaction") + 1
            Else
                objDic.AddNew RS.Fields("transdt").Value & "", Join(Array(0, 0, 0, 0, 0, 0, 1, 0, 0), COL_DIV)
            End If
            RS.MoveNext
        Loop
    End If
    
    Set Get_ReactionCount = New clsDictionary
    Set Get_ReactionCount = objDic
    
    Set RS = Nothing
    
End Function
'-------------------------
'수혈부작용건수(frmBBs916)
'-------------------------
Public Function Get_TransReactionCnt(ByVal Fdt As String, ByVal Tdt As String) As clsDictionary
    Dim objDic As clsDictionary
    Dim RS     As Recordset
    Dim SSQL   As String
    
    Set Get_TransReactionCnt = New clsDictionary
    Set objDic = New clsDictionary
    
    With objDic
        .Clear
        .FieldInialize "seq", "y10,y1to2,y2to3,y3to4,y4to5,y5to6,y60,yetc"
        .AddNew 1, Join(Array(0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
        .AddNew 2, Join(Array(0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
        .AddNew 3, Join(Array(0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
    End With
    
    SSQL = " SELECT " & F_SSN2("c") & " as ssn " & _
           " FROM " & _
                  T_BBS501 & " a," & T_BBS202 & " b," & T_HIS001 & " c" & _
           " WHERE" & _
                  "     a.transdt between " & DBV("transdt", Fdt) & " AND " & DBV("transdt", Tdt) & _
                  " AND " & DBW("a.reactdiv=", "1") & _
                  " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
                  " AND b.ptid=c." & F_PTID
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Dim strTmp As String
        Dim age    As String
        Dim sex    As String
        
        Do Until RS.EOF
            strTmp = SexCheck(RS.Fields("ssn").Value & "")
            age = medGetP(strTmp, 2, COL_DIV)
            sex = medGetP(strTmp, 1, COL_DIV)
            With objDic
                If .Exists(sex) Then
                    .KeyChange sex
                    If age = "etc" Then
                        .Fields("yetc") = .Fields("yetc") + 1
                    Else
                        Select Case Val(age)
                            Case 0 To 10:  .Fields("y10") = .Fields("y10") + 1
                            Case 10 To 20: .Fields("y1to2") = .Fields("y1to2") + 1
                            Case 20 To 30: .Fields("y2to3") = .Fields("y2to3") + 1
                            Case 30 To 40: .Fields("y3to4") = .Fields("y3to4") + 1
                            Case 40 To 50: .Fields("y4to5") = .Fields("y4to5") + 1
                            Case 50 To 60: .Fields("y5to6") = .Fields("y5to6") + 1
                            Case Else:     .Fields("y60") = .Fields("y60") + 1
                        End Select
                    End If
               End If
            End With
            RS.MoveNext
        Loop
    End If
    
    Set Get_TransReactionCnt = objDic
    Set objDic = Nothing
End Function

Private Function SexCheck(ByVal SSN As String) As String
    Dim strTmp As String
    Dim lngsex As Long
    
    strTmp = Mid(SSN, 3, 6) & "-" & Mid(SSN, 9)
    
    If strTmp <> "" Then
        lngsex = Val(Mid(medGetP(strTmp, 2, "-"), 1, 1))
        If lngsex = 1 Or lngsex = 3 Then
            SexCheck = "1"
        ElseIf lngsex = 2 Or lngsex = 4 Then
            SexCheck = "2"
        Else
            SexCheck = "3"
        End If
    Else
        SexCheck = "3"
    End If
    
    If Len(SSN) = 15 Then
        SexCheck = SexCheck & COL_DIV & medFindAge(Mid(SSN, 1, 8), "Y")
    Else
        SexCheck = SexCheck & COL_DIV & "etc"
    End If
End Function
Public Function Get_ExpireRecord(ByVal Year As String) As clsDictionary
    Dim objDic As New clsDictionary
    Dim RS     As Recordset
    Dim SSQL   As String
    Dim Fyear  As String
    Dim Tyear  As String
    
    
    Fyear = Year & "0101"
    Tyear = Year & "1231"
    
    objDic.Clear
    objDic.FieldInialize "rsncd", "rsnnm,mon1,mon2,mon3,mon4,mon5,mon6,mon7,mon8,mon9,mon10,mon11,mon12"
    Set Get_ExpireRecord = New clsDictionary
    Set Get_ExpireRecord = Get_ExpireRsnRecord(objDic)
    
    SSQL = " SELECT realexpdt,exprsncd FROM " & T_BBS401 & _
           " WHERE" & _
                  " " & DBW("stscd", BBSBloodStatus.stsEXPIRE, 2) & _
                  " AND realexpdt between " & DBV("realexpdt", Fyear) & " AND " & DBV("realexpdt", Tyear)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    With Get_ExpireRecord
        If Not RS.EOF Then
            Do Until RS.EOF
                If .Exists(RS.Fields("exprsncd").Value & "") Then
                    .KeyChange RS.Fields("exprsncd").Value & ""
                    Select Case Mid(RS.Fields("realexpdt").Value & "", 5, 2)
                        Case "01": .Fields("mon1") = .Fields("mon1") + 1
                        Case "02": .Fields("mon2") = .Fields("mon2") + 1
                        Case "03": .Fields("mon3") = .Fields("mon3") + 1
                        Case "04": .Fields("mon4") = .Fields("mon4") + 1
                        Case "05": .Fields("mon5") = .Fields("mon5") + 1
                        Case "06": .Fields("mon6") = .Fields("mon6") + 1
                        Case "07": .Fields("mon7") = .Fields("mon7") + 1
                        Case "08": .Fields("mon8") = .Fields("mon8") + 1
                        Case "09": .Fields("mon9") = .Fields("mon9") + 1
                        Case "10": .Fields("mon10") = .Fields("mon10") + 1
                        Case "11": .Fields("mon11") = .Fields("mon11") + 1
                        Case "12": .Fields("mon12") = .Fields("mon12") + 1
                    End Select
                End If
                RS.MoveNext
            Loop
        End If
    End With
    
    Set objDic = Nothing
End Function
Private Function Get_ExpireRsnRecord(ByVal objDic As clsDictionary) As clsDictionary
    Dim RS   As Recordset
    Dim SSQL As String
    
    Set Get_ExpireRsnRecord = New clsDictionary
    SSQL = " SELECT * FROM " & T_COM003 & " WHERE " & DBW("cdindex", BC2_EXP_RESON, 2)
    
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            objDic.AddNew RS.Fields("cdval1").Value & "", Join(Array(RS.Fields("field1").Value & "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
            RS.MoveNext
        Loop
    End If
    Set Get_ExpireRsnRecord = objDic
    
    Set RS = Nothing
End Function
Public Function Get_MSBOS(ByVal Fdt As String, ByVal Tdt As String, Optional ByVal ocd As String = "") As clsDictionary
    Dim objDic As New clsDictionary
    Dim objABO As New clsQueryOrder
    Dim RS     As Recordset
    Dim SSQL   As String
    Dim PtId   As String
    Dim orddt  As String
    Dim ordno  As String
    Dim ABO    As String
    
    Set Get_MSBOS = New clsDictionary
    
    objDic.Clear
    objDic.FieldInialize "orddt,accdt,accseq", "ordcd,ordnm,abo,componm,volume,unit,rmk"
    
    SSQL = " SELECT a.ptid,a.orddt,a.ordno,a.bloodop,b." & F_ONM & " as onm,c.accdt,c.accseq,e.componm,d.volumn" & _
           " FROM " & _
                    T_BBS101 & " a," & T_HIS007 & " b," & T_LAB102 & " c," & T_BBS001 & " d," & T_BBS006 & " e" & _
           " WHERE" & _
                  "" & _
                  "" & _
                  "  a.bloodop=b." & F_OCD & _
                  " AND a.ptid=c.ptid AND a.orddt=c.orddt AND a.ordno=c.ordno AND a.ordseq=c.ordseq" & _
                  " AND c.ordcd=d.testcd" & _
                  " AND d.applydt=(SELECT max(z.applydt)" & _
                                 " FROM " & T_BBS001 & " z" & _
                                 " WHERE z.testcd=c.ordcd" & _
                                 "   AND " & DBW("z.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & ")" & _
                  " AND d.compocd=e.compocd"
    
    If ocd <> "" Then SSQL = SSQL & " AND " & DBW("a.bloodop", ocd, 2)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Set Get_MSBOS = Get_DetailMsBos(RS, objDic)
    End If
    
    
    
    Set RS = Nothing
    Set objABO = Nothing
    Set objDic = Nothing
End Function
Private Function Get_DetailMsBos(ByVal RS As Recordset, ByVal objDic As clsDictionary) As clsDictionary
    Dim objABO As New clsABO
    Dim RS1     As Recordset
    
    Dim PtId   As String
    Dim orddt  As String
    Dim accdt  As String
    Dim accseq As String
    Dim ABO    As String
    Dim strTmp As String
    
    Dim SSQL   As String
    
    Set Get_DetailMsBos = New clsDictionary
    
    Do Until RS.EOF
        PtId = RS.Fields("ptid").Value & ""
        orddt = RS.Fields("orddt").Value & ""
        accdt = RS.Fields("accdt").Value & ""
        accseq = RS.Fields("accseq").Value & ""
        With objABO
            .PtId = PtId
            .GetABO
            ABO = .ABO
        End With
        objDic.AddNew Join(Array(orddt, accdt, accseq), COL_DIV), _
                      Join(Array(RS.Fields("attrval").Value & "", RS.Fields("onm").Value & "", _
                      ABO, RS.Fields("componm").Value & "", RS.Fields("volumn").Value & "", 0, ""), COL_DIV)
        SSQL = " SELECT * FROM " & T_BBS402 & _
               " WHERE " & DBW("workarea", C_WORKAREA, 2) & _
               " AND " & DBW("accdt", accdt, 2) & _
               " AND " & DBW("accseq", accseq, 2) & _
               " AND (retfg='' or retfg is null) AND (expfg='' or expfg is null)"
               
        Set RS1 = New Recordset
        RS1.Open SSQL, DBConn
        If Not RS1.EOF Then
            Do Until RS1.EOF
                If objDic.Exists(Join(Array(orddt, RS1.Fields("accdt").Value & "", RS1.Fields("accseq").Value & ""), COL_DIV)) Then
                    objDic.KeyChange Join(Array(RS1.Fields("accseq").Value & "", RS1.Fields("accseq").Value & ""), COL_DIV)
                    objDic.Fields("unit") = objDic.Fields("unit") + 1
                End If
                RS1.MoveNext
            Loop
        End If
        strTmp = PtId
        RS.MoveNext
    Loop
    Set Get_DetailMsBos = objDic
    
    Set RS = Nothing
    Set objABO = Nothing
End Function

Public Function Get_OcdList(Optional ByVal ocd As String = "") As String
    Get_OcdList = " SELECT " & F_OCD & " as ocd," & F_ONM & " as onm " & _
                  " FROM " & T_HIS007 & _
                  " WHERE " & DBW(F_ODIV, "21", 2)
                  
    If ocd <> "" Then
        Get_OcdList = Get_OcdList & " AND " & DBW(F_OCD, ocd, 2)
    End If
End Function

Public Function Get_Ct_Ratio(ByVal Year As String, Optional ByVal Centercd As String = "") As String
'---------------------------------------------------
'C-T Ratio 를 구하기위해 XM건수와 출고 건수를 구한다
'조건이 ALL 일때(출고건수 &Col_Div & XM건수)
'---------------------------------------------------
    Dim SSQL        As String
    Dim RS          As Recordset
    Dim YearF       As String
    Dim YearT       As String
    Dim XMcnt       As Long
    Dim DeliveryCnt As Long
    
    YearF = medGetP(Year, 1, COL_DIV) 'Year & "0101"
    YearT = medGetP(Year, 2, COL_DIV) 'Year & "1231"

    SSQL = " SELECT sum(assigncnt) as assign,sum(deliverycnt) as delivery,sum(expcnt) as expcnt ,sum(retcnt) as retcnt " & _
           " FROM " & T_LAB101 & " a," & T_LAB102 & " b," & T_BBS203 & " c" & _
           " WHERE " & _
                     DBW("a.orddiv=", C_WORKAREA) & _
           " AND " & DBW("b.rcvdt>=", YearF) & _
           " AND " & DBW("b.rcvdt<=", YearT) & _
           " AND a.ptid=b.ptid " & _
           " AND a.orddt=b.orddt" & _
           " AND a.ordno=b.ordno" & _
           " AND b.workarea=c.workarea" & _
           " AND b.accdt=c.accdt" & _
           " AND b.accseq=c.accseq"
    
    'If centercd <> "" Then ssql = ssql & " AND centercd" & DBS(centercd, 2)
    If mvarDept <> "" Then SSQL = SSQL & " AND " & DBW("a.deptcd=", mvarDept)
    If mvarWard <> "" Then SSQL = SSQL & " AND " & DBW("a.wardid=", mvarWard)
    If mvarDoct <> "" Then SSQL = SSQL & " AND " & DBW("a.majdoct=", mvarDoct)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Get_Ct_Ratio = RS.Fields("assign").Value & "" & COL_DIV & (Val(RS.Fields("delivery").Value & "") - Val(RS.Fields("expcnt").Value & "")) - Val(RS.Fields("retcnt").Value & "")
'        Do Until Rs.EOF
'            SELECT Case Rs.Fields("bloodtype").Value
'                Case "02": DeliveryCnt = DeliveryCnt + Rs.Fields("bloodcnt").Value + Rs.Fields("donationcnt").Value + Rs.Fields("localcnt").Value
'                Case "07": XMcnt = XMcnt + Rs.Fields("bloodcnt").Value + Rs.Fields("donationcnt").Value + Rs.Fields("localcnt").Value
'            End SELECT
'            Rs.MoveNext
'        Loop
'        Get_Ct_Ratio = DeliveryCnt & COL_DIV & XMcnt
    End If
    
    Set RS = Nothing
    
End Function
Public Function Statics962() As Recordset
    Dim SSQL As String
    
    
    SSQL = "SELECT cdval1,field1 FROM " & T_COM003 & " WHERE  " & DBW("cdindex=", "B024") & " order by field2 "
    
    
    Set Statics962 = New Recordset
    Statics962.Open SSQL, DBConn

End Function
Public Function MonthStatics(ByVal CloseMonth As String, ByVal ABO As String) As Recordset
    Dim FMonth As String
    Dim TMonth As String
    Dim SSQL   As String
    
    FMonth = CloseMonth & "01"
    TMonth = CloseMonth & "31"
    SSQL = " SELECT  b.groupcd, sum(a.trans1) as trans1, sum(a.trans2) as trans2" & _
           " FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                    DBW("a.closedt<=", TMonth) & _
           " AND " & DBW("a.closedt>=", FMonth) & _
           " AND " & DBW("a.abo=", ABO) & _
           " AND a.bloodtype ='09'" & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY  groupcd"
    
    Set MonthStatics = New Recordset
    MonthStatics.Open SSQL, DBConn
    
End Function
Public Function MonthBloodType(ByVal CloseMonth As String, ByVal ABO As String) As Recordset
    Dim FMonth As String
    Dim TMonth As String
    Dim SSQL   As String
    
    FMonth = CloseMonth & "01"
    TMonth = CloseMonth & "31"
    
    '현월에 입고,출고 ,반환,폐기 된 내역을 조회한다.
    
    SSQL = " SELECT  a.bloodtype,b.groupcd, sum(bloodcnt) as bloodcnt ,sum(donationcnt) as donationcnt,sum(localcnt) as localcnt" & _
           " FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                    DBW("a.closedt<=", TMonth) & _
           " AND " & DBW("a.closedt>=", FMonth) & _
           " AND " & DBW("a.abo=", ABO) & _
           " AND a.bloodtype in('01','02','03','04')" & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY bloodtype, groupcd"
    
    Set MonthBloodType = New Recordset
    MonthBloodType.Open SSQL, DBConn

End Function
Public Function MonthBloodType_전월재고(ByVal CloseMonth As String, ByVal ABO As String) As Recordset
    Dim FMonth As String
    Dim TMonth As String
    Dim SSQL   As String
    
    FMonth = CloseMonth & "01"
    TMonth = CloseMonth & "31"
    
    '전월재고 조회시 이전달의 혈액에대해서만 조회를 하였는데
    '전월재고는 선택된 현월보다 작은 모든 혈액 수급현황을 조회한다.
    
    SSQL = " SELECT  a.bloodtype,b.groupcd, sum(bloodcnt) as bloodcnt ,sum(donationcnt) as donationcnt,sum(localcnt) as localcnt" & _
           " FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                    DBW("a.closedt<=", TMonth) & _
           " AND " & DBW("a.abo=", ABO) & _
           " AND a.bloodtype in('01','02','03','04')" & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY bloodtype, groupcd"
    
    Set MonthBloodType_전월재고 = New Recordset
    MonthBloodType_전월재고.Open SSQL, DBConn

End Function

Public Function MonthTotal(ByVal FClose As String, ByVal TClose As String) As Recordset
    Dim SSQL As String
    
'    ssql = " SELECT  a.bloodtype,b.groupcd, sum(bloodcnt) as bloodcnt ,sum(donationcnt) as donationcnt,sum(localcnt) as localcnt,sum(trans1) as trans1, sum(trans2) as trans2" & _
           " FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                    DBW("a.closedt<=", TClose) & _
           " AND " & DBW("a.closedt>=", FClose) & _
           " AND a.bloodtype in('01','02','03','04','09')" & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY bloodtype, groupcd"
    SSQL = " SELECT  a.bloodtype,b.groupcd, sum(bloodcnt) as bloodcnt ,sum(donationcnt) as donationcnt,sum(localcnt) as localcnt,sum(trans1) as trans1, sum(trans2) as trans2" & _
           " FROM " & T_BBS901 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                    DBW("a.closedt<=", TClose) & _
           " AND a.bloodtype in('01','02','03','04','09')" & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY bloodtype, groupcd"
    
    Set MonthTotal = New Recordset
    MonthTotal.Open SSQL, DBConn

End Function

Public Function IRR_Count(ByVal PreMonth As String, ByVal CurMonth As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    Dim strCMonthFrom As String
    Dim strCMonthTo   As String
    Dim strPMonTh     As String
    
    strCMonthFrom = CurMonth & "01"
    strCMonthTo = CurMonth & "31"
    strPMonTh = PreMonth & "31"
    '월계
    SSQL = " SELECT sum(irrcnt)  as cnt FROM " & T_BBS901 & _
           " WHERE " & DBW("bloodtype=", "11") & _
           " AND " & DBW("closedt<=", strCMonthTo) & _
           " AND " & DBW("closedt>=", strCMonthFrom) & _
           " GROUP BY bloodtype"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then IRR_Count = RS.Fields("cnt").Value & ""
    Set RS = Nothing
    
    SSQL = " SELECT sum(irrcnt) as cnt FROM " & T_BBS901 & _
           " WHERE " & DBW("bloodtype=", "11") & _
           " AND " & DBW("closedt<=", strPMonTh) & _
           " GROUP BY bloodtype"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then IRR_Count = IRR_Count & COL_DIV & RS.Fields("cnt").Value & ""
    Set RS = Nothing
         
         
End Function

Public Function GetFilterSQL(ByVal strFrdate As String, ByVal strToDate As String, ByVal QueryDiv As String) As String
    Dim SSQL    As String
    
    
    SSQL = " SELECT a.ordcd,a.ptid,a.orddt,a.entdt,a.enttm,a.entid,b." & F_PTNM & " as ptnm,c.testnm" & _
           " FROM " & T_BBS001 & " C," & T_HIS001 & " B," & T_BBS304 & " A" & _
           " WHERE " & DBW("a.entdt>=", strFrdate) & _
           " AND " & DBW("a.entdt<=", strToDate) & _
           " AND a.ordcd=c.testcd" & _
           " AND (c.expdt='' or c.expdt is null)" & _
           " AND a.ptid =b." & F_PTID
           
    Select Case QueryDiv
        Case "0": SSQL = SSQL & " AND A.EXPFG IS NULL AND A.RETFG IS NULL"
        Case "1": SSQL = SSQL & " AND " & DBW("A.RETFG=", "1")
        Case "2": SSQL = SSQL & " AND " & DBW("A.EXPFG=", "1")
    End Select
    
    SSQL = SSQL & " ORDER BY ordcd,ptid,orddt,entdt"
    GetFilterSQL = SSQL
    
End Function

Public Function GetBloodStatic(ByVal strFrdate As String, ByVal strToDate As String, _
                               ByVal Vol As String, ByVal Div As String, ByVal Hosfg As String) As String
'제재별/용량별 입고현황
    Dim SSQL As String
    
    SSQL = " SELECT a.compocd,b.componm,a.abo,a.volumn,count(*) as cnt " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.entdt>=", strFrdate) & _
           " AND " & DBW("a.entdt<=", strToDate) & _
           " AND " & DBW("a.volumn=", Vol)
    
    If Hosfg = "1" Then
        SSQL = SSQL & " and " & DBW("a.hosfg<>", "1")
    End If
    
    SSQL = SSQL & _
           " AND a.compocd=b.compocd" & _
           " AND (b.expdt='' or b.expdt is null)" & _
           " GROUP BY a.compocd,b.componm,a.abo,a.volumn" & _
           " ORDER BY a.compocd,a.abo,a.volumn"
    GetBloodStatic = SSQL
    
End Function

Public Function GetBloodStaticForDelivery(ByVal strFrdate As String, ByVal strToDate As String, _
                               ByVal Vol As String, ByVal Div As String, ByVal Hosfg As String) As String
'제재별/용량별 출고현황(입고현황에서 stscd='3' 인 조건만 추가됨)
    Dim SSQL As String
    
    SSQL = " SELECT a.compocd,b.componm,a.abo,a.volumn,count(*) as cnt " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a, " & T_BBS402 & " c " & _
           " WHERE a.bldsrc = c.bldsrc " & _
           " AND a.bldyy = c.bldyy " & _
           " AND a.bldno = c.bldno " & _
           " AND a.compocd = c.compocd " & _
           " AND " & DBW("c.deliverydt>=", strFrdate) & _
           " AND " & DBW("c.deliverydt<=", strToDate) & _
           " AND " & DBW("a.volumn=", Vol) & _
           " and a.stscd='3' "
    
    If Hosfg = "1" Then
        SSQL = SSQL & " and " & DBW("a.hosfg<>", "1")
    End If
    
'    SSQL = SSQL & _
           " AND a.compocd=b.compocd " & _
           " AND (b.realexpdt='' or b.realexpdt is null) " & _
           " GROUP BY a.compocd,b.componm,a.abo,a.volumn " & _
           " ORDER BY a.compocd,a.abo,a.volumn "
    
    SSQL = SSQL & _
           " AND a.compocd=b.compocd " & _
           " GROUP BY a.compocd,b.componm,a.abo,a.volumn " & _
           " ORDER BY a.compocd,a.abo,a.volumn "
           
    GetBloodStaticForDelivery = SSQL
    
    '-- 원본 -----------------------------------------------------------------
'    SSQL = " SELECT a.compocd,b.componm,a.abo,a.volumn,count(*) as cnt " & _
'           " FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
'           " WHERE " & DBW("a.entdt>=", strFrdate) & _
'           " AND " & DBW("a.entdt<=", strToDate) & _
'           " AND " & DBW("a.volumn=", Vol) & _
'           " and a.stscd='3' "
'
'    If Hosfg = "1" Then
'        SSQL = SSQL & " and " & DBW("a.hosfg<>", "1")
'    End If
'
'    SSQL = SSQL & _
'           " AND a.compocd=b.compocd" & _
'           " AND (b.expdt='' or b.expdt is null)" & _
'           " GROUP BY a.compocd,b.componm,a.abo,a.volumn" & _
'           " ORDER BY a.compocd,a.abo,a.volumn"
'
'    GetBloodStaticForDelivery = SSQL
    '--------------------------------------------------------------------------
    
End Function

