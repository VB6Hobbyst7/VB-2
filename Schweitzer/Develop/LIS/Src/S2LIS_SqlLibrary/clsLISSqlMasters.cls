VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'+--------------------------------------------------------------------------------+"
'| clsItem에서 사용하는 멤버 변수                                                 |
'+--------------------------------------------------------------------------------+"

Public Function SqlRstType_EtcTest() As String
    SqlRstType_EtcTest = "SELECT stcd FROM " & T_LAB350 & " group by stcd"
End Function

Public Function SqlWorkArea() As String
   SqlWorkArea = " SELECT cdval1 as WorkArea, field1 as WorkNm " & _
                 " FROM " & T_LAB032 & _
                 " WHERE " & DBW("cdindex = ", lc3_workarea) & _
                 " order by WorkArea"
End Function

Public Function SqlGroupCd() As String
   SqlGroupCd = " SELECT cdval1 as GroupCd, field1 as GroupNm " & _
                " FROM " & T_LAB032 & _
                " WHERE " & DBW("cdindex = ", LC3_StaticGroup) & _
                " order by GroupCd"
End Function

Public Function SqlOutLabList() As String
   SqlOutLabList = " SELECT cdval1 as OutLabCd, field1 as OutLabNm " & _
                   " FROM " & T_LAB032 & _
                   " WHERE " & DBW("cdindex = ", LC3_OutLab) & _
                   " order by OutLabCd"
End Function

Public Function SqlItemList() As String
   SqlItemList = " SELECT testcd, testnm FROM " & T_LAB001 & " group by testcd, testnm order by testcd"
End Function

Public Function SqlItemQuery(ByVal strTestCd As String, _
                                      Optional ByVal strApplyDt As Variant) As String
   SqlItemQuery = " SELECT a.* " & _
                  " FROM " & T_LAB001 & " a " & _
                  " WHERE " & DBW("a.testcd = ", strTestCd)
   If IsMissing(strApplyDt) Then
      SqlItemQuery = SqlItemQuery & _
                  " AND a.applydt = (SELECT max(applydt) " & _
                                   " FROM " & T_LAB001 & _
                                   " WHERE testcd = a.testcd) "
   Else
      SqlItemQuery = SqlItemQuery & _
                  " AND " & DBW("applydt = ", strApplyDt)
   End If
End Function

Public Function SqlItemInsert(ByVal objTLab001 As clsTLab001) As String
    With objTLab001
        SqlItemInsert = " Insert into " & _
                          T_LAB001 & " (testcd, applydt, testnm, abbrnm5, abbrnm10, " & _
                                     "  workarea, rsttype, testdiv, rptseq, panelfg, " & _
                                     "  rstdiv, outlabcd, grpfg, workunit, matchcd, " & _
                                     "  txttype, attrcd, expdt, detailfg, itemseq, groupcd) " & _
                        " Values  (" & DBV("testcd", .TestCd, 1) & DBV("applydt", .ApplyDt, 1) & _
                                       DBV("testnm", .TestNm, 1) & DBV("abbrnm5", .AbbrNm5, 1) & _
                                       DBV("abbrnm10", .AbbrNm10, 1) & DBV("workarea", .WorkArea, 1) & _
                                       DBV("rsttype", .RstType, 1) & DBV("testdiv", .TestDiv, 1) & _
                                       DBV("rptseq", .RptSeq, 1) & DBV("panelfg", .PanelFg, 1) & _
                                       DBV("rstdiv", .RstDiv, 1) & DBV("outlabcd", .OutLabCd, 1) & _
                                       DBV("grpfg", .GrpFg, 1) & DBV("workunit", .WorkUnit, 1) & _
                                       DBV("matchcd", .MatchCd, 1) & DBV("txttype", .TxtType, 1) & _
                                       DBV("attrcd", .AttrCd, 1) & DBV("expdt", .ExpDt, 1) & _
                                       DBV("detailfg", .DetailFg, 1) & DBV("itemseq", .ItemSeq, 1) & _
                                       DBV("groupcd", .GroupCd) & _
                                   " ) "
    End With
    
End Function

Public Function SqlItemUpdate(ByVal objTLab001 As clsTLab001) As String
    With objTLab001
        SqlItemUpdate = " Update " & T_LAB001 & _
                        " Set " & DBW("testnm   = ", .TestNm) & ", " & _
                                  DBW("abbrnm5  = ", .AbbrNm5) & ", " & _
                                  DBW("abbrnm10 = ", .AbbrNm10) & ", " & _
                                  DBW("workarea = ", .WorkArea) & ", " & _
                                  DBW("rsttype  = ", .RstType) & ", " & _
                                  DBW("testdiv  = ", .TestDiv) & ", " & _
                                  DBW("rptseq   = ", .RptSeq) & ", " & _
                                  DBW("panelfg  = ", .PanelFg) & ", " & _
                                  DBW("detailfg = ", .DetailFg) & ", " & _
                                  DBW("rstdiv   = ", .RstDiv) & ", " & _
                                  DBW("outlabcd = ", .OutLabCd) & ", " & _
                                  DBW("grpfg    = ", .GrpFg) & ", " & _
                                  DBW("workunit = ", .WorkUnit) & ", " & _
                                  DBW("attrcd   = ", .AttrCd) & ", " & _
                                  DBW("expdt    = ", .ExpDt) & ", " & _
                                  DBW("txttype  = ", .TxtType) & ", " & _
                                  DBW("matchcd  = ", .MatchCd) & ", " & _
                                  DBW("itemseq  = ", .ItemSeq) & ", " & _
                                  DBW("groupcd  = ", .GroupCd) & _
                        " WHERE " & DBW("testcd  = ", .TestCd) & _
                        " AND   " & DBW("applydt = ", .ApplyDt)
    End With
End Function

Public Function SqlItemDelete(ByVal objTLab001 As clsTLab001) As String
    With objTLab001
        SqlItemDelete = " Delete FROM " & T_LAB001 & _
                        " WHERE " & DBW("testcd  = ", .TestCd) & _
                        " AND   " & DBW("applydt = ", .ApplyDt)
    End With
End Function

'-----------------------------------------------------------------
'기준치 마스터(LAB005) 로부터 해당 검사아이템의 기준치정보를 검색
'-----------------------------------------------------------------
Public Function SqlRefQuery(ByVal strTestCd As String, _
                                      ByVal strSpcCd As String, _
                                      ByVal strApplyDt As String, _
                                      ByVal strApplySex As String, _
                                      ByVal intAgeFrom As Integer, _
                                      ByVal intAgeTo As Integer) As String
   SqlRefQuery = " SELECT * " & _
                 " FROM  " & T_LAB005 & _
                 " WHERE " & DBW("testcd   = ", strTestCd) & _
                 " AND   " & DBW("spccd    = ", strSpcCd) & _
                 " AND   " & DBW("applydt  = ", strApplyDt) & _
                 " AND   " & DBW("applyxex = ", strApplySex) & _
                 " AND   " & DBW("agefrom  = ", intAgeFrom) & _
                 " AND   " & DBW("ageto    = ", intAgeTo)
End Function

Private Function Lab005Chk(ByVal objTlab005 As clsTLab005) As Boolean
    Dim rs      As Recordset
    Dim sSQL    As String
    
    With objTlab005
    
        sSQL = " SELECT * FROM " & T_LAB005 & _
               " WHERE " & DBW("testcd     = ", .TestCd) & _
               " AND   " & DBW("spccd      = ", .SpcCd) & _
               " AND   " & DBW("applydt    = ", .ApplyDt) & _
               " AND   " & DBW("applysex   = ", .ApplySex) & _
               " AND   " & DBW("agefrom    = ", .AgeFrom) & _
               " AND   " & DBW("ageto      = ", .AgeTo)
        Set rs = New Recordset
        rs.Open sSQL, DBConn
        
        If Not rs.EOF Then
            Lab005Chk = True
        End If
    End With
    Set rs = Nothing
End Function

Public Function SqlRefInsert(ByVal objTlab005 As clsTLab005) As String
    
   Dim blnLab005Chk As Boolean
    
    
    
    blnLab005Chk = Lab005Chk(objTlab005)
    
    With objTlab005
        If .RefDiv = "1" Then
        '일반참고치
            If blnLab005Chk = False Then
                SqlRefInsert = " Insert into " & _
                                 T_LAB005 & " (testcd, spccd, applydt, applysex, agefrom, ageto, " & _
                                             " refvalfrom, refvalto, refcd, reftext, expdt,arefvalfrom,arefvalto )" & _
                                " Values  (" & _
                                              DBV("testcd", .TestCd, 1) & _
                                              DBV("spccd", .SpcCd, 1) & _
                                              DBV("applydt", .ApplyDt, 1) & _
                                              DBV("applysex", .ApplySex, 1) & _
                                              DBV("agefrom", .AgeFrom, 1) & _
                                              DBV("ageto", .AgeTo, 1) & _
                                              DBV("refvalfrom", .RefValFrom, 1) & _
                                              DBV("refvalto", .RefValTo, 1) & _
                                              DBV("refcd", .RefCd, 1) & _
                                              DBV("reftext", .RefText, 1) & _
                                              DBV("expdt", .ExpDt, 1) & _
                                              DBV("arefvalfrom", .ARefValFrom, 1) & _
                                              DBV("arefvalto", .ARefValTo) & _
                                        ") "
            Else
                SqlRefInsert = " Update " & T_LAB005 & _
                               " Set " & DBW("refvalfrom = ", .RefValFrom) & ", " & _
                                         DBW("refvalto   = ", .RefValTo) & ", " & _
                                         DBW("arefvalfrom = ", .ARefValFrom) & ", " & _
                                         DBW("arefvalto   = ", .ARefValTo) & ", " & _
                                         DBW("refcd      = ", .RefCd) & ", " & _
                                         DBW("reftext    = ", .RefText) & ", " & _
                                         DBW("expdt      = ", .ExpDt) & _
                               " WHERE " & DBW("testcd     = ", .TestCd) & _
                               " AND   " & DBW("spccd      = ", .SpcCd) & _
                               " AND   " & DBW("applydt    = ", .ApplyDt) & _
                               " AND   " & DBW("applysex   = ", .ApplySex) & _
                               " AND   " & DBW("agefrom    = ", .AgeFrom) & _
                               " AND   " & DBW("ageto      = ", .AgeTo)
            
            End If
            
        Else
        'Panic 참고치
            If blnLab005Chk = False Then
                SqlRefInsert = " Insert into " & _
                                 T_LAB005 & " (testcd, spccd, applydt, applysex, agefrom, ageto, " & _
                                             " panicfrval, panictoval, amrfrval, amrtoval, refcd, reftext, expdt,arletfrval,arlettoval )" & _
                                " Values  (" & _
                                              DBV("testcd", .TestCd, 1) & _
                                              DBV("spccd", .SpcCd, 1) & _
                                              DBV("applydt", .ApplyDt, 1) & _
                                              DBV("applysex", .ApplySex, 1) & _
                                              DBV("agefrom", .AgeFrom, 1) & _
                                              DBV("ageto", .AgeTo, 1) & _
                                              DBV("panicfrval", .PanicFrVal, 1) & _
                                              DBV("panictoval", .PanicToVal, 1) & _
                                              DBV("amrfrval", .AMRFrVal, 1) & _
                                              DBV("amrtoval", .AMRToVal, 1) & _
                                              DBV("refcd", .RefCd, 1) & _
                                              DBV("reftext", .RefText, 1) & _
                                              DBV("expdt", .ExpDt, 1) & _
                                              DBV("arletfrval", .ArletFrVal, 1) & _
                                              DBV("arlettoval", .ArletToVal) & _
                                        ") "
            Else
                SqlRefInsert = " Update " & T_LAB005 & _
                               " Set " & DBW("panicfrval = ", .PanicFrVal) & ", " & _
                                         DBW("panictoval   = ", .PanicToVal) & ", " & _
                                         DBW("amrfrval = ", .AMRFrVal) & ", " & _
                                         DBW("amrtoval   = ", .AMRToVal) & ", " & _
                                         DBW("arletfrval = ", .ArletFrVal) & ", " & _
                                         DBW("arlettoval   = ", .ArletToVal) & ", " & _
                                         DBW("refcd      = ", .RefCd) & ", " & _
                                         DBW("reftext    = ", .RefText) & ", " & _
                                         DBW("expdt      = ", .ExpDt) & _
                               " WHERE " & DBW("testcd     = ", .TestCd) & _
                               " AND   " & DBW("spccd      = ", .SpcCd) & _
                               " AND   " & DBW("applydt    = ", .ApplyDt) & _
                               " AND   " & DBW("applysex   = ", .ApplySex) & _
                               " AND   " & DBW("agefrom    = ", .AgeFrom) & _
                               " AND   " & DBW("ageto      = ", .AgeTo)
            
            End If
        End If
    End With
End Function

Public Function SqlRefUpdate(ByVal objTlab005 As clsTLab005) As String
    With objTlab005
        If .RefDiv = "1" Then
        '일반참고치
            SqlRefUpdate = " Update " & T_LAB005 & _
                           " Set " & DBW("refvalfrom = ", .RefValFrom) & ", " & _
                                     DBW("refvalto   = ", .RefValTo) & ", " & _
                                     DBW("arefvalfrom = ", .ARefValFrom) & ", " & _
                                     DBW("arefvalto   = ", .ARefValTo) & ", " & _
                                     DBW("refcd      = ", .RefCd) & ", " & _
                                     DBW("reftext    = ", .RefText) & ", " & _
                                     DBW("expdt      = ", .ExpDt) & _
                           " WHERE " & DBW("testcd     = ", .TestCd) & _
                           " AND   " & DBW("spccd      = ", .SpcCd) & _
                           " AND   " & DBW("applydt    = ", .ApplyDt) & _
                           " AND   " & DBW("applysex   = ", .ApplySex) & _
                           " AND   " & DBW("agefrom    = ", .AgeFrom) & _
                           " AND   " & DBW("ageto      = ", .AgeTo)
        ElseIf .RefDiv = "2" Then
        'Panic 참고치
            SqlRefUpdate = " Update " & T_LAB005 & _
                           " Set " & DBW("panicfrval = ", .PanicFrVal) & ", " & _
                                     DBW("panictoval = ", .PanicToVal) & ", " & _
                                     DBW("arletfrval = ", .ArletFrVal) & ", " & _
                                     DBW("arlettoval = ", .ArletToVal) & ", " & _
                                     DBW("refcd      = ", .RefCd) & ", " & _
                                     DBW("reftext    = ", .RefText) & ", " & _
                                     DBW("expdt      = ", .ExpDt) & _
                           " WHERE " & DBW("testcd     = ", .TestCd) & _
                           " AND   " & DBW("spccd      = ", .SpcCd) & _
                           " AND   " & DBW("applydt    = ", .ApplyDt) & _
                           " AND   " & DBW("applysex   = ", .ApplySex) & _
                           " AND   " & DBW("agefrom    = ", .AgeFrom) & _
                           " AND   " & DBW("ageto      = ", .AgeTo)
        Else
            SqlRefUpdate = " Update " & T_LAB005 & _
                           " Set " & DBW("amrfrval = ", .AMRFrVal) & ", " & _
                                     DBW("amrtoval   = ", .AMRToVal) & " " & _
                           " WHERE " & DBW("testcd    = ", .TestCd) & _
                           " AND   " & DBW("spccd     = ", .SpcCd) & _
                           " AND   " & DBW("applydt   = ", .ApplyDt) & _
                           " AND   " & DBW("applysex  = ", .ApplySex) & _
                           " AND   " & DBW("agefrom   = ", .AgeFrom) & _
                           " AND   " & DBW("ageto     = ", .AgeTo)
        End If
    End With
End Function

Public Function SqlRefDelete(ByVal objTlab005 As clsTLab005, _
                             Optional ByVal blnDelAll As Boolean = False) As String
    With objTlab005
        SqlRefDelete = " Delete FROM " & T_LAB005 & _
                       " WHERE " & DBW("testcd   = ", .TestCd) & _
                       " AND   " & DBW("spccd    = ", .SpcCd) & _
                       " AND   " & DBW("applydt  = ", .ApplyDt)
        If Not blnDelAll Then
            SqlRefDelete = SqlRefDelete & " AND   " & DBW("applysex = ", .ApplySex) & _
                            " AND   " & DBW("agefrom  = ", .AgeFrom) & _
                            " AND   " & DBW("ageto    = ", .AgeTo)
        End If
    End With
    
End Function

'-----------------------------------------------------------------
'지정검체 마스터(LAB004) 로부터 해당 검사아이템의 검체정보를 검색
'-----------------------------------------------------------------
Public Function SqlStoreCd() As String
   SqlStoreCd = " SELECT cdval1, field1 " & _
                " FROM  " & T_LAB032 & _
                " WHERE " & DBW("cdindex = ", LC3_StoreCd) & _
                " order by cdval1 "
End Function

Public Function SqlBuildings() As String
   SqlBuildings = " SELECT cdval1, field1 " & _
                  " FROM  " & T_LAB032 & _
                  " WHERE " & DBW("cdindex = ", LC3_Buildings) & _
                  " order by cdval1 "
End Function

Public Function SqlSpcGrp() As String
   SqlSpcGrp = " SELECT cdval1, field1 " & _
               " FROM  " & T_LAB032 & _
               " WHERE " & DBW("cdindex = ", LC3_SGroup) & _
               " order by cdval1 "
End Function

Public Function SqlSpcQuery(ByVal strTestCd As String, _
                                      ByVal strSpcCd As String, _
                                      Optional ByVal strApplyDt As Variant) As String
   SqlSpcQuery = " SELECT * " & _
                 " FROM  " & T_LAB004 & " a " & _
                 " WHERE " & DBW("a.testcd = ", strTestCd) & _
                 " AND   " & DBW("a.spccd  = ", strSpcCd)
                 
   If IsMissing(strApplyDt) Then
      SqlSpcQuery = SqlSpcQuery & _
                 " AND   a.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                 "                    WHERE testcd = a.testcd " & _
                 "                    AND    spccd = a.spccd) "
   Else
      SqlSpcQuery = SqlSpcQuery & _
                 " AND " & DBW("applydt = ", strApplyDt)
   End If

End Function

Public Function SqlSpcInsert(ByVal objTLab004 As clsTLab004) As String
    With objTLab004
        SqlSpcInsert = " Insert into " & _
                         T_LAB004 & "(testcd, spccd, applydt, seq, labelcnt, " & _
                                    "rstunit, rndfg, statfg, statflags, avalval, panicfg, " & _
                                    "panicfrval, panictoval, deltafg, deltaval, " & _
                                    "testcost, storecd, tatavg, spcqty, spcunit, expdt, deltaval2, tats,arletfg,arletfrval,arlettoval) " & _
                        " Values (" & _
                                    DBV("testcd", .TestCd, 1) & DBV("spccd", .SpcCd, 1) & _
                                    DBV("applydt", .ApplyDt, 1) & DBV("seq", .Seq, 1) & _
                                    DBV("labelcnt", .LabelCnt, 1) & DBV("rstunit", .RstUnit, 1) & _
                                    DBV("rndfg", .RndFg, 1) & DBV("statfg", .StatFg, 1) & _
                                    DBV("statflags", .StatFlags, 1) & DBV("avalval", .AvalVal, 1) & _
                                    DBV("panicfg", .PanicFg, 1) & DBV("panicfrval", .PanicFrVal, 1) & _
                                    DBV("panictoval", .PanicToVal, 1) & DBV("deltafg", .DeltaFg, 1) & _
                                    DBV("deltaval", .DeltaVal1, 1) & DBV("testcost", .TestCost, 1) & _
                                    DBV("storecd", .StoreCd, 1) & DBV("tatavg", .TatAvg, 1) & _
                                    DBV("spcqty", .SpcQty, 1) & DBV("spcunit", .SpcUnit, 1) & _
                                    DBV("expdt", .ExpDt, 1) & DBV("deltaval2", .DeltaVal2, 1) & _
                                    DBV("tats", .TATS, 1) & DBV("arletfg", .ArletFg, 1) & DBV("arletfrval", .ArletFrVal, 1) & _
                                    DBV("arlettoval", .ArletToVal) & _
                                ")"
    End With
End Function

Public Function SqlSpcUpdate(ByVal objTLab004 As clsTLab004) As String
    With objTLab004
        SqlSpcUpdate = " Update " & T_LAB004 & _
                       " Set " & DBW("labelcnt   = ", .LabelCnt) & ", " & DBW("rstunit    = ", .RstUnit) & ", " & _
                                 DBW("seq        = ", .Seq) & ", " & DBW("rndfg      = ", .RndFg) & ", " & _
                                 DBW("statfg     = ", .StatFg) & ", " & DBW("statflags  = ", .StatFlags) & ", " & _
                                 DBW("avalval    = ", .AvalVal) & ", " & DBW("panicfg    = ", .PanicFg) & ", " & _
                                 DBW("panicfrval = ", .PanicFrVal) & ", " & DBW("panictoval = ", .PanicToVal) & ", " & _
                                 DBW("deltafg    = ", .DeltaFg) & ", " & DBW("deltaval   = ", .DeltaVal1) & ", " & _
                                 DBW("testcost   = ", .TestCost) & ", " & DBW("storecd    = ", .StoreCd) & ", " & _
                                 DBW("tatavg     = ", .TatAvg) & ", " & DBW("spcqty     = ", .SpcQty) & ", " & _
                                 DBW("spcunit    = ", .SpcUnit) & ", " & DBW("expdt      = ", .ExpDt) & ", " & _
                                 DBW("deltaval2  = ", .DeltaVal2) & ", " & DBW("tats       = ", .TATS) & ", " & _
                                 DBW("arletfg       =", .ArletFg) & ", " & DBW("arletfrval    =", .ArletFrVal) & ", " & _
                                 DBW("arlettoval    =", .ArletToVal) & " " & _
                       " WHERE " & DBW("testcd  = ", .TestCd) & _
                       " AND   " & DBW("spccd   = ", .SpcCd) & _
                       " AND   " & DBW("applydt = ", .ApplyDt)
    End With
End Function

Public Function SqlSpcDelete(ByVal objTLab004 As clsTLab004) As String
    With objTLab004
        SqlSpcDelete = " Delete FROM " & T_LAB004 & _
                       " WHERE " & DBW("testcd  = ", .TestCd) & _
                       " AND   " & DBW("spccd   = ", .SpcCd) & _
                       " AND   " & DBW("applydt = ", .ApplyDt)
    End With
End Function

Public Function SqlCheckCostCd(ByVal pCostCd As String) As String
    SqlCheckCostCd = " SELECT * " & _
                     " FROM  " & T_HIS008 & _
                     " WHERE " & DBW(F_AMTCD, pCostCd, 2)
End Function

'-----------------------------------------------------------------
'일반검사 Worksheet 마스터
'-----------------------------------------------------------------

Public Function SqlDeleteWorkSheetMaster(ByVal pWsCd As String) As String
    
    SqlDeleteWorkSheetMaster = " Delete " & T_LAB008 & _
                               " WHERE  " & DBW("workcd", pWsCd, 2)

End Function

Public Function SqlInsertWorkSheetMaster(ByVal pWsCd As String, ByVal pTestCd As String, _
                                         ByVal pSpcCd As String, ByVal pWaCd As String, _
                                         ByVal pWorkInfo As String, ByVal pEmpNm As String) As String
    SqlInsertWorkSheetMaster = " Insert into " & T_LAB008 & _
                               " (workcd, testcd, spccd, workarea, workinfo, empnm) " & _
                               " values( " & _
                                           DBV("workcd", pWsCd, 1) & _
                                           DBV("testcd", pTestCd, 1) & _
                                           DBV("spccd", pSpcCd, 1) & _
                                           DBV("workarea", pWaCd, 1) & _
                                           DBV("workinfo", pWorkInfo, 1) & _
                                           DBV("empnm", pEmpNm) & _
                                      ")"
End Function

Public Function SqlGetWorkSheetInfo(ByVal pWsCd As String) As String

    SqlGetWorkSheetInfo = " SELECT a.workcd as workcd, " & _
                          "        a.empnm as empnm , " & _
                          "        a.testcd as testcd, " & _
                          "        a.spccd as spccd, " & _
                          "        a.workinfo as workinfo, " & _
                          "        b.field1 as wsnm , " & _
                          "        c.cdval1 as workareacd , " & _
                          "        c.field1 as workareanm , " & _
                          "        d.cdval1 as buildcd, " & _
                          "        d.field1 as buildnm " & _
                          " FROM " & T_LAB008 & " a, " & _
                                     T_LAB032 & " b, " & _
                                     T_LAB032 & " c, " & _
                                     T_LAB032 & " d " & _
                          " WHERE " & _
                                    DBW("a.workcd=", pWsCd) & _
                          " AND   a.workcd = b.cdval1  " & _
                          " AND   c.cdval1 = a.workarea  " & _
                          " AND " & DBW("c.cdindex =", lc3_workarea) & _
                          " AND " & DBW("b.cdindex =", LC3_WorkSheetName) & _
                          " AND " & DBJ(DBW("d.cdindex =", LC3_Buildings)) & _
                          " AND " & DBJ("d.cdval1  =* b.field2")
End Function
 


Public Function SqlGetSpcList(ByVal sTestCd As String) As String
     
    SqlGetSpcList = " SELECT a.spccd as spcCd , b.field3 as spcNm " & _
                    " FROM  " & T_LAB004 & " a, " & T_LAB032 & " b " & _
                    " WHERE " & DBW("a.testcd  =", sTestCd) & _
                    " AND   " & DBW("b.cdindex =", LC3_Specimen) & _
                    " AND   a.spccd = b.cdval1 "
End Function

'-----------------------------------------------------------------
'미생물 검사 Worksheet 마스터
'-----------------------------------------------------------------

Public Function SqlGetDetailTest(ByVal sTestNm As String, ByVal sWorkArea As String) As String
    
    SqlGetDetailTest = " SELECT a.testcd, a.testnm " & _
                       " FROM " & T_LAB001 & " a, " & T_LAB031 & " b " & _
                       " WHERE " & DBW("b.cdindex=", LC2_Detail) & _
                       " AND   " & DBW("b.cdval1 =", sTestNm) & _
                       " AND   a.testcd = b.cdval2" & _
                       " AND   a.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                          " WHERE testcd = a.testcd) " & _
                       " AND " & DBW("a.workarea=", sWorkArea) & _
                       " AND " & DBW("a.rsttype =", MRT_Stain) & _
                       " AND " & DBW("a.testdiv =", TST_MicTest) & _
                       " AND " & DBW("a.detailfg=", "*")
End Function
                      

Public Function SqlGetGSTest(ByVal sWorkArea As String) As String
    
    SqlGetGSTest = " SELECT testcd, testnm " & _
                   " FROM " & T_LAB001 & _
                   " WHERE " & _
                             DBW("workarea=", sWorkArea) & _
                   " AND " & DBW("rsttype =", MRT_Stain) & _
                   " AND " & DBW("testdiv =", TST_MicTest) & _
                   " AND (detailfg = '' or detailfg is null) "
'                   " AND " & DBW("panelfg =", PN_Detail)
End Function
                      
Public Function SqlGetCheckedMedia(ByVal sSpcGroupCd As String) As String
    
    SqlGetCheckedMedia = " SELECT a.text1 as MediaNm , b.cdval2 as MediaCd" & _
                         " FROM " & T_LAB032 & " a, " & T_LAB031 & " b " & _
                         " WHERE " & DBW("a.cdindex=", LC3_Media) & _
                         " AND   " & DBW("b.cdindex=", LC2_SpcMedia) & _
                         " AND     a.cdval1 = b.cdval2" & _
                         " AND   " & DBW("b.cdval1=", sSpcGroupCd)
End Function

Public Function SqlGetWsGroup() As String

    SqlGetWsGroup = " SELECT b.cdval1 wscd, b.field1 wsnm, b.field2 rsttp, a.field2 wa, a.field3 seq " & _
                    " FROM  " & T_LAB032 & " b, " & T_LAB032 & " a " & _
                    " WHERE " & DBW("a.cdindex = ", LC3_SGroup) & _
                    " AND   " & DBW("b.cdindex = ", LC3_MWSKinds) & _
                    " AND    b.cdval1 = a.field4" & _
                    " Order  by seq asc"

End Function

'-----------------------------------------------------------------
'기타 검사 Template 마스터
'-----------------------------------------------------------------

Public Function SqlLoadSpecialTest(Optional ByVal pRstType As String = "") As String
    
    SqlLoadSpecialTest = " SELECT testcd, testnm FROM " & T_LAB001 & _
                         " WHERE " & DBW("testdiv = ", TST_SpeTest)
    If pRstType = "" Then
        SqlLoadSpecialTest = SqlLoadSpecialTest & " AND  (rsttype is null  or  rsttype = '') "
    Else
        SqlLoadSpecialTest = SqlLoadSpecialTest & " AND " & DBW("rsttype = ", pRstType)
    End If
    
End Function

Public Function SqlLoadSpecialTestAll(Optional ByVal pRstType As String = "") As String
    
    SqlLoadSpecialTestAll = " SELECT testcd, testnm FROM " & T_LAB001 & _
                            " WHERE " & DBW("testdiv = ", TST_SpeTest)
    
End Function

Public Function SqlGetStCd(ByVal pFlag As String, Optional ByVal StSeq As String = "", _
                           Optional ByVal StCd As String = "", _
                           Optional ByVal StItem As String = "") As String
    If pFlag = "1" Then
        SqlGetStCd = " SELECT distinct stcd FROM " & T_LAB350
    Else
        SqlGetStCd = " SELECT stcd FROM " & T_LAB350 & _
                     " WHERE  " & DBW("stseq  =", "0") & _
                     " AND    " & DBW("stcd   =", StCd) & _
                     ""
        If StItem <> "" Then
            SqlGetStCd = SqlGetStCd & "  AND    " & DBW("stitem =", StItem)
        End If
    End If
    
End Function

Public Function SqlUpdateRstType(ByVal sCondition As String, ByVal sRstType As String, ByVal pFlag As String) As String
    
    SqlUpdateRstType = " UPDATE " & T_LAB001 & _
                       " SET    " & DBW("rsttype = ", sRstType) & _
                       " WHERE  " & DBW("testdiv = ", TST_SpeTest)
    If pFlag = "1" Then
        SqlUpdateRstType = SqlUpdateRstType & " AND  " & DBW("testcd  = ", sCondition)
    Else
        SqlUpdateRstType = SqlUpdateRstType & " AND  " & DBW("rsttype = ", sCondition)
    End If
    
End Function


Public Function SqlGetLAB036(Optional ByVal sStCd As String = "", Optional ByVal sTpCd As String = "") As String

    SqlGetLAB036 = " SELECT * FROM " & T_LAB036    ' DB 중복체크
    If sStCd <> "" Then
        SqlGetLAB036 = SqlGetLAB036 & " WHERE  " & DBW("stcd = ", sStCd)
        If sTpCd <> "" Then
            SqlGetLAB036 = SqlGetLAB036 & " AND  " & DBW("tpcd = ", sTpCd)
        End If
    Else
        If sTpCd <> "" Then
            SqlGetLAB036 = SqlGetLAB036 & " WHERE  " & DBW("tpcd = ", sTpCd)
        End If
    End If
    
    SqlGetLAB036 = SqlGetLAB036 & " order by tpcd "
    
End Function

Public Function SqlGetLAB350(Optional ByVal sStCd As String = "", Optional ByVal sWhere As String = "") As String

    SqlGetLAB350 = " SELECT * FROM " & T_LAB350
    If sStCd <> "" Then
        SqlGetLAB350 = SqlGetLAB350 & " WHERE  " & DBW("stcd = ", sStCd)
    End If
    If sWhere <> "" Then
        SqlGetLAB350 = SqlGetLAB350 & " " & sWhere
    End If
    
End Function

Public Function SqlInsertLAB350(ByVal sStCd As String, ByVal sStSeq As String, _
                                ByVal sStItem As String) As String

    SqlInsertLAB350 = " insert into " & T_LAB350 & _
                      " (stcd, stseq, stitem) " & _
                      " values (" & DBV("stcd", sStCd) & " ," & _
                                    DBV("stseq", sStSeq) & ", " & _
                                    DBV("stitem", sStItem) & ")"

End Function
                    
Public Function SqlDeleteLAB350(ByVal sStCd As String) As String

    SqlDeleteLAB350 = " delete FROM " & T_LAB350 & _
                      " WHERE  " & DBW("stcd =", sStCd)
End Function
                    
Public Function SqlDeleteLAB036(ByVal sStCd As String) As String
    
    SqlDeleteLAB036 = " delete FROM " & T_LAB036 & _
                      " WHERE  " & DBW("stcd =", sStCd)

End Function

Public Function SqlInsertLAB036(ByVal sStCd As String, ByVal sTpCd As String, ByVal sWsFg As String, _
                                ByVal sTpNm As String, ByVal sTpData As String) As String
            
    SqlInsertLAB036 = " insert into " & T_LAB036 & "(stcd, tpcd,wsfg, tpnm, tpdata) " & _
                      " values (" & DBV("stcd", sStCd) & " , " & _
                                    DBV("tpcd", sTpCd) & " , " & _
                                    DBV("wsfg", sWsFg) & " , " & _
                                    DBV("tpnm", sTpNm) & " , " & _
                                    DBV("tpdata", sTpData) & ")"


End Function
