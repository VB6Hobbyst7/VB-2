VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlQc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------------------
'clsControlItem
'----------------------------------------------------------------------------------------------
Public Function SqlItemSave(objTLab022 As clsTLab022) As String
    With objTLab022
        SqlItemSave = "insert into " & T_LAB022 & "(ctrlcd, levelcd, testcd, eqpcd) " & _
                      "values (" & DBV("ctrlcd", .CtrlCd, 1) & DBV("levelcd", .LevelCd, 1) & _
                                   DBV("testcd", .TestCd, 1) & DBV("eqpcd", .EqpCd) & ")"
    End With
End Function


'----------------------------------------------------------------------------------------------
'clsControlMaster
'----------------------------------------------------------------------------------------------
Public Function SqlControlCd(ByVal pCtrlCd As String, Optional ByVal pLevelCd As String = "A") As String
    If pLevelCd = "A" Then
        SqlControlCd = " SELECT a.*, b.eqpnm, c.field1 as SectNm " & _
                       " FROM  " & T_LAB021 & " a, " & T_LAB006 & " b, " & T_LAB032 & " c " & _
                       " WHERE " & DBW("a.ctrlcd = ", pCtrlCd) & " " & _
                       " AND   " & DBJ("b.eqpcd =* a.eqpcd") & _
                       " AND   " & DBW("c.cdindex = ", LC3_Section) & " " & _
                       " AND   c.cdval1 = a.sectcd "
    Else
        SqlControlCd = " SELECT a.*, b.eqpnm, c.field1 as SectNm " & _
                       " FROM  " & T_LAB021 & " a, " & T_LAB006 & " b, " & T_LAB032 & " c " & _
                       " WHERE " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
                       " AND   " & DBW("a.levelcd = ", pLevelCd) & " " & _
                       " AND   " & DBJ("b.eqpcd =* a.eqpcd") & _
                       " AND   " & DBW("c.cdindex = ", LC3_Section) & " " & _
                       " AND   c.cdval1 = a.sectcd "
    End If
End Function

Public Function SqlQCItems(ByVal pCtrlCd As String, ByVal pLevelCd As String) As String
    '## 5.0.23: 이상대(2005-04-13)
    '   - 지정검체 마스터에 폐기된 코드와 폐기되지 않은 코드가 seq가 동일할 경우
    '     중복된 검사항목을 조회하여 쿼리수정
    '## 5.0.24: 이상대(2005-05-14)
    '   - 지정검체 마스터에서 폐기된 코드가 조회되어 쿼리수정
    SqlQCItems = "SELECT a.*, b.testnm, b.panelfg, c.avalval, c.rstunit " & _
                 "FROM  " & T_LAB022 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                 "WHERE " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
                 "AND   " & DBW("a.levelcd = ", pLevelCd) & " " & _
                 "AND   b.testcd = a.testcd " & _
                 "AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                   " WHERE testcd = b.testcd) " & _
                 "AND   c.testcd = a.testcd " & _
                 "AND   c.seq = (SELECT min(seq) FROM " & T_LAB004 & " WHERE testcd = c.testcd AND (expdt='' OR expdt IS NULL)) " & _
                 "AND   (c.expdt='' OR c.expdt IS NULL)"
End Function

Public Function SqlDetails(ByVal pTestCd As String) As String
    '------- 맞을까??????????? ----------------------------------------------------------------------------------
    SqlDetails = "SELECT a.cdval2 as testcd, b.testnm, b.rstdiv, b.rsttype, b.panelfg, c.avalval, c.rstunit " & _
                 "FROM  " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                 "WHERE " & DBW("a.cdindex = ", LC2_Detail) & " " & _
                 "AND   " & DBW("a.cdval1  = ", pTestCd) & " " & _
                 "AND     b.testcd = a.cdval2 " & _
                 "AND     b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd) " & _
                 "AND     c.testcd = a.cdval2 " & _
                 "AND     c.seq = (SELECT min(seq) FROM " & T_LAB004 & " " & _
                                 " WHERE testcd = c.testcd AND (expdt = '' or expdt is null))"
End Function

Public Function SqlSaveData(ByVal NewFg As Boolean, ByVal objTLab021 As clsTLab021) As String
    With objTLab021
        If NewFg Then
            SqlSaveData = "insert into " & T_LAB021 & _
                                   "(ctrlcd, levelcd, ctrlnm, eqpcd, sectcd, ctrldiv, workarea, buildcd) " & _
                          "values (" & DBV("ctrlcd", .CtrlCd, 1) & DBV("levelcd", .LevelCd, 1) & _
                                       DBV("ctrlnm", .CtrlNm, 1) & DBV("eqpcd", .EqpCd, 1) & _
                                       DBV("sectcd", .SectCd, 1) & DBV("ctrldiv", .CtrlDiv, 1) & _
                                       DBV("workarea", .WorkArea, 1) & DBV("buildcd", .BuildCd) & ")"
        Else
            SqlSaveData = " update " & T_LAB021 & " set "
            SqlSaveData = SqlSaveData & DBW("ctrlnm   = ", .CtrlNm, 1)
            SqlSaveData = SqlSaveData & DBW("eqpcd    = ", .EqpCd, 1)
            SqlSaveData = SqlSaveData & DBW("sectcd   = ", .SectCd, 1)
            SqlSaveData = SqlSaveData & DBW("ctrldiv  = ", .CtrlDiv, 1)
            SqlSaveData = SqlSaveData & DBW("workarea = ", .WorkArea, 1)
            SqlSaveData = SqlSaveData & DBW("buildcd  = ", .BuildCd)
            SqlSaveData = SqlSaveData & " WHERE " & DBW("ctrlcd = ", .CtrlCd)
            SqlSaveData = SqlSaveData & " AND " & DBW("levelcd  = ", .LevelCd)
        End If
    End With
End Function

Public Function SqlDeleteData(ByVal pCtrlCd As String, ByVal pLevelCd As String) As String
    SqlDeleteData = "delete FROM " & T_LAB022 & " " & _
                    "WHERE " & DBW("ctrlcd  = ", pCtrlCd) & " " & _
                    "AND   " & DBW("levelcd = ", pLevelCd)

End Function

Public Function SqlDeleteAllData(ByVal pCtrlCd As String, ByVal pLevelCd As String, ByVal pCase As String) As String
    
    Dim strTblNm As String
    
    Select Case pCase
    Case 1:
        strTblNm = T_LAB021
    Case 2:
        strTblNm = T_LAB022
    Case 3:
        strTblNm = T_LAB023
    Case 4:
        strTblNm = T_LAB024
    End Select
    SqlDeleteAllData = " delete FROM " & strTblNm & " " & _
                       " WHERE " & DBW("ctrlcd  = ", pCtrlCd) & " " & _
                       " AND   " & DBW("levelcd = ", pLevelCd)
    
End Function

Public Function SqlCreateSchedule(ByVal pFromDt As String, ByVal pToDt As String, _
                          ByVal pCtrlCd As String, Optional ByVal pLevelCd As String = "A") As String
    If pLevelCd = "A" Then
        SqlCreateSchedule = " delete FROM " & T_LAB025 & " " & _
                            " WHERE " & DBW("dodt  >= ", pFromDt) & _
                            " AND   " & DBW("dodt  <= ", pToDt) & _
                            " AND   " & DBW("ctrlcd = ", pCtrlCd)
    Else
        SqlCreateSchedule = " delete FROM " & T_LAB025 & " " & _
                            " WHERE " & DBW("dodt   >= ", pFromDt) & _
                            " AND   " & DBW("dodt   <= ", pToDt) & _
                            " AND   " & DBW("ctrlcd  = ", pCtrlCd) & _
                            " AND   " & DBW("levelcd = ", pLevelCd)
    End If
End Function

Public Function SqlSchedule(ByVal pCtrlCd As String, ByVal pFromDt As String, _
                       ByVal pToDt As String, Optional ByVal pLevelCd As String = "A") As String
                       
    SqlSchedule = " SELECT a.*, b.ctrlnm " & _
                  " FROM " & T_LAB025 & " a, " & T_LAB021 & " b " & _
                  " WHERE " & DBW("a.dodt >= ", pFromDt) & _
                  " AND   " & DBW("a.dodt <= ", pToDt) & _
                  " AND   " & DBW("a.ctrlcd = ", pCtrlCd) & _
                  " AND   b.ctrlcd = a.ctrlcd  AND b.levelcd = a.levelcd "
    If pLevelCd <> "A" Then _
        SqlSchedule = SqlSchedule & " AND " & DBW("a.levelcd = ", pLevelCd)
End Function


'----------------------------------------------------------------------------------------------
'clsQcItem
'----------------------------------------------------------------------------------------------
Public Function SqlQcItemSave(ByVal objTLab024 As clsTLab024) As String
    With objTLab024
        SqlQcItemSave = "insert into " & T_LAB024 & _
                               "(ctrlcd, levelcd, lotno, testcd, meanval, sdval, avalval, " & _
                               "refcd, rstunit, cvval, minval, maxval, schedfg, wmset, calfg, " & _
                               "entdt, entid) " & _
                        "values ( " & DBV("ctrlcd", .CtrlCd, 1) & DBV("levelcd", .LevelCd, 1) & _
                                DBV("lotno", .LotNo, 1) & DBV("testcd", .TestCd, 1) & _
                                DBV("meanval", .MeanVal, 1) & DBV("sdval", .SdVal, 1) & _
                                DBV("avalval", .AvalVal, 1) & DBV("refcd", .RefCd, 1) & _
                                DBV("rstunit", .RstUnit, 1) & DBV("cvval", .CvVal, 1) & _
                                DBV("minval", .MinVal, 1) & DBV("maxval", .MaxVal, 1) & _
                                DBV("schedfg", .SchedFg, 1) & DBV("wmset", .WmSet, 1) & _
                                DBV("calfg", .CalFg, 1) & _
                                DBV("entdt", .EntDt, 1) & DBV("entid", .EntId) & ")"
    End With
End Function

'----------------------------------------------------------------------------------------------
'clsQcMicRst
'----------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------
'clsQcMst
'----------------------------------------------------------------------------------------------
Public Function SqlMstLotnoList(ByVal pCtrlCd As String, ByVal pLevelCd As String) As String
    SqlMstLotnoList = " SELECT a.lotno, a.opendt, a.expdt " & _
                      " FROM " & T_LAB023 & " a " & _
                      " WHERE " & DBW("a.ctrlcd = ", pCtrlCd) & _
                      " AND   " & DBW("a.levelcd = ", pLevelCd) & _
                      " order by a.opendt desc "
End Function

Public Function SqlMstQcData(ByVal pCtrlCd As String, ByVal pLevelCd As String, ByVal pLotNo As String) As String
    SqlMstQcData = " SELECT a.*, b.ctrlnm " & _
                   " FROM " & T_LAB023 & " a, " & T_LAB021 & " b " & _
                   " WHERE " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
                   " AND   " & DBW("a.levelcd = ", pLevelCd) & " " & _
                   " AND   " & DBW("a.lotno   = ", pLotNo) & " " & _
                   " AND   b.ctrlcd = a.ctrlcd"
End Function

Public Function SqlMstQCItems(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                      ByVal pLotNo As String, Optional ByVal pDiv As String = "D") As String
    SqlMstQCItems = " SELECT a.*, b.abbrnm5, b.testnm, b.testdiv, b.rstdiv, b.rsttype, b.panelfg, b.detailfg " & _
                    " FROM  " & T_LAB001 & " b, " & T_LAB024 & " a " & _
                    " WHERE " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
                    " AND   " & DBW("a.levelcd = ", pLevelCd) & " " & _
                    " AND   " & DBW("a.lotno   = ", pLotNo) & " " & _
                    " AND   b.testcd  = a.testcd " & _
                    " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                      " WHERE testcd = b.testcd) "
    If pDiv <> "D" Then SqlMstQCItems = SqlMstQCItems & " AND (b.detailfg = ''  or  b.detailfg is null) "
    SqlMstQCItems = SqlMstQCItems & " order by b.rptseq, b.testcd"

End Function

Public Function SqlOrderQCItems(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                                ByVal pLotNo As String, ByVal pWorkArea As String, _
                                ByVal pAccDt As String, ByVal paccseq As String) As String
    
    SqlOrderQCItems = " SELECT a.*, b.abbrnm5, b.testnm, b.testdiv, b.rstdiv, b.rsttype, b.panelfg, b.detailfg, c.testcd ordfg " & _
                      " FROM  " & T_LAB001 & " b, " & T_LAB026 & " c, " & T_LAB024 & " a " & _
                      " WHERE " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
                      " AND   " & DBW("a.levelcd = ", pLevelCd) & " " & _
                      " AND   " & DBW("a.lotno   = ", pLotNo) & " " & _
                      " AND   b.testcd  = a.testcd " & _
                      " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                         " WHERE testcd = b.testcd) " & _
                      " AND   " & DBJ(DBW("c.workarea = ", pWorkArea)) & _
                      " AND   " & DBJ(DBW("c.accdt = ", pAccDt)) & _
                      " AND   " & DBJ(DBW("c.accseq = ", paccseq)) & _
                      " AND   " & DBJ("c.testcd =* a.testcd") & _
                      "" & _
                      " order by b.rptseq, b.testcd"
                      
'2001-12-10 수정:
'Detail Item만 처방을 낼수 있으므로 다음조건 삭제
'" AND  (b.detailfg = ''  or  b.detailfg is null) "

End Function

Public Function SqlMstSaveData(ByVal NewFg As Boolean, ByVal objTLab023 As clsTLab023) As String
    With objTLab023
        If NewFg Then
            SqlMstSaveData = "insert into " & T_LAB023 & _
                                        "(ctrlcd, levelcd, lotno, opendt, expdt, makecd, remark) " & _
                             "values (" & DBV("ctrlcd", .CtrlCd, 1) & DBV("levelcd", .LevelCd, 1) & _
                                          DBV("lotno", .LotNo, 1) & DBV("opendt", .OpenDt, 1) & _
                                          DBV("expdt", .ExpDt, 1) & DBV("makecd", .MakeCd, 1) & _
                                          DBV("remark", .Remark) & ")"
        Else
            SqlMstSaveData = " update " & T_LAB023 & " set "
            SqlMstSaveData = SqlMstSaveData & " " & DBW("opendt = ", .OpenDt, 1)
            SqlMstSaveData = SqlMstSaveData & " " & DBW("expdt  = ", .ExpDt, 1)
            SqlMstSaveData = SqlMstSaveData & " " & DBW("makecd = ", .MakeCd, 1)
            SqlMstSaveData = SqlMstSaveData & " " & DBW("remark = ", .Remark)
            SqlMstSaveData = SqlMstSaveData & " WHERE " & DBW("ctrlcd  = ", .CtrlCd)
            SqlMstSaveData = SqlMstSaveData & " AND   " & DBW("levelcd = ", .LevelCd)
            SqlMstSaveData = SqlMstSaveData & " AND   " & DBW("lotno   = ", .LotNo)
        End If
    End With
End Function

Public Function SqlMstDeleteData(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                                 ByVal pLotNo As String) As String
    SqlMstDeleteData = " delete FROM " & T_LAB024 & " " & _
                       " WHERE " & DBW("ctrlcd  = ", pCtrlCd) & _
                       " AND   " & DBW("levelcd = ", pLevelCd) & _
                       " AND   " & DBW("lotno   = ", pLotNo)
End Function

Public Function SqlMstDeleteAll(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                                ByVal pLotNo As String, ByVal pCase As String) As String
    Select Case pCase
    Case "1":
        SqlMstDeleteAll = " delete FROM " & T_LAB023 & " " & _
                           " WHERE " & DBW("ctrlcd  = ", pCtrlCd) & _
                           " AND   " & DBW("levelcd = ", pLevelCd) & _
                           " AND   " & DBW("lotno   = ", pLotNo)
    Case "2":
        SqlMstDeleteAll = " delete FROM " & T_LAB024 & " " & _
                           " WHERE " & DBW("ctrlcd  = ", pCtrlCd) & _
                           " AND   " & DBW("levelcd = ", pLevelCd) & _
                           " AND   " & DBW("lotno   = ", pLotNo)
    End Select
    
End Function


'----------------------------------------------------------------------------------------------
'clsQcOrder
'----------------------------------------------------------------------------------------------
Public Function SqlTestNames(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal paccseq As String) As String
                             
    SqlTestNames = "SELECT distinct a.abbrnm5 " & _
                   "FROM " & T_LAB001 & " a, " & T_LAB026 & " b " & _
                   "WHERE " & DBW("b.workarea = ", pWorkArea) & " " & _
                   "AND   " & DBW("b.accdt    = ", pAccDt) & " " & _
                   "AND   " & DBW("b.accseq   = ", paccseq) & " " & _
                   "AND   a.testcd = b.testcd "
                   '"order by a.rptseq, a.testcd"
End Function

Public Function SqlLastLotNo(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                             ByVal pFrDt As String, ByVal pToDt As String) As String
'distinct 삭제함
    SqlLastLotNo = "SELECT a.lotno " & _
                   "FROM  " & T_LAB023 & " a " & _
                   "WHERE " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
                   "AND   " & DBW("a.levelcd = ", pLevelCd) & " " & _
                   "AND   " & DBW("a.opendt <= ", pFrDt) & " " & _
                   "AND   " & DBW("a.expdt  >= ", pToDt) & " " & _
                   "order by a.opendt desc "
End Function

Public Function SqlLabNumbers(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                              ByVal pLotNo As String, ByVal pWorkArea As String, _
                              ByVal pFrDt As String, ByVal pToDt As String, _
                              ByVal lngLength As Long) As String
                              
    SqlLabNumbers = "SELECT distinct a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.accdt" & _
                            FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("char", "a.accseq") & " as labno, " & _
                    "       a.lotno, b.entdt " & _
                    "FROM  " & T_LAB201 & " b, " & T_LAB026 & " a " & _
                    "WHERE " & DBW("a.ctrlcd   = ", pCtrlCd) & " " & _
                    "AND   " & DBW("a.levelcd  = ", pLevelCd) & " "
    If Trim(pLotNo) <> "" Then
        SqlLabNumbers = SqlLabNumbers & "AND   " & DBW("a.lotno    = ", pLotNo) & " "
    End If
    SqlLabNumbers = SqlLabNumbers & _
                    "AND   " & DBW("a.workarea = ", pWorkArea) & " " & _
                    "AND   " & DBW("a.accdt   >= ", Mid(pFrDt, 1, lngLength)) & " " & _
                    "AND   " & DBW("a.accdt   <= ", Mid(pToDt, 1, lngLength)) & " " & _
                    "AND   b.workarea = a.workarea " & _
                    "AND   b.accdt = a.accdt " & _
                    "AND   b.accseq = a.accseq " & _
                    "order by entdt desc"
End Function

Public Function SqlCollectQuery(ByVal pWorkArea As String, _
                                ByVal pAccDt As String, ByVal paccseq As Long) As String

    SqlCollectQuery = " SELECT a.* " & _
                      " FROM  " & T_LAB201 & " a " & _
                      " WHERE " & DBW("a.workarea = ", pWorkArea) & _
                      " AND   " & DBW("a.accdt    = ", pAccDt) & _
                      " AND   " & DBW("a.accseq   = ", paccseq)
End Function


Public Function SqlGetControlList(ByVal pTestCd As String) As String

    SqlGetControlList = " SELECT distinct a.ctrlcd, b.ctrlnm, b.workarea " & _
                        " FROM  " & T_LAB024 & " a, " & T_LAB021 & " b " & _
                        " WHERE " & DBW("a.testcd = ", pTestCd) & _
                        " AND   b.ctrlcd  = a.ctrlcd" & _
                        " AND   b.levelcd = a.levelcd"

End Function

Public Function SqlGetLotNoList(ByVal pCtrlCd As String) As String
           
    SqlGetLotNoList = " SELECT distinct a.levelcd, a.lotno, a.opendt, a.expdt, b.eqpcd, c.eqpnm " & _
                      " FROM " & T_LAB023 & " a, " & T_LAB021 & " b, " & T_LAB006 & " c " & _
                      " WHERE" & DBW("a.ctrlcd =", pCtrlCd) & _
                      " AND  " & DBW("a.opendt <= ", Format(Now, CS_DateDbFormat)) & _
                      " AND ( a.expdt = ''  or  a.expdt is null or " & _
                                 DBW("a.expdt > ", Format(Now, CS_DateDbFormat)) & " )" & _
                      " AND   b.ctrlcd  = a.ctrlcd" & _
                      " AND   b.levelcd = a.levelcd " & _
                      " AND   " & DBJ("c.eqpcd   =* b.eqpcd ")

End Function

Public Function SqlGetOutQcWorkArea() As String
           
    SqlGetOutQcWorkArea = " SELECT * FROM " & T_LAB032 & _
                          " WHERE " & DBW("cdindex =", lc3_workarea) & _
                          " AND   " & DBW("field4  =", "Y")

End Function

Public Function SqlGetOutQcLastSeq(ByVal pWorkArea As Integer, ByVal pOutDiv As Integer) As String
    
    SqlGetOutQcLastSeq = " SELECT max(" & FUNC_CONVERT("num", "ptid") & ") as seq FROM " & T_LAB201 & _
                         " WHERE " & DBW("workarea = ", pWorkArea) & _
                         " AND   " & DBW("accdt    = ", Format(Now, "YYYY")) & _
                         " AND   " & DBW("buildcd  = ", pOutDiv)
End Function

'----------------------------------------------------------------------------------
'5. 미생물 QC Master 관련 Sql
'----------------------------------------------------------------------------------

Public Function SqlMicQcRstCode() As String

    SqlMicQcRstCode = " SELECT * FROM " & T_LAB031 & _
                      " WHERE " & DBW("cdindex = ", LC2_MicQcRst) & _
                      " order by cdval1, cdval2"

End Function

Public Function SqlMicQcRstGroup(Optional ByVal sCdval2 As String = "*") As String
    
    SqlMicQcRstGroup = " SELECT cdval1, cdval2, field1 FROM " & T_LAB031 & _
                       " WHERE " & DBW("cdindex = ", LC2_MicTemp) & _
                       " AND   " & DBW("cdval2  = ", sCdval2) & _
                       " order by cdval1"

End Function


Public Function SqlDeleteMicQcRstGroup(ByVal sCdval1 As String) As String
        
    SqlDeleteMicQcRstGroup = " delete " & T_LAB031 & _
                             " WHERE  " & DBW("cdindex = ", LC2_MicTemp) & _
                             " AND    " & DBW("cdval1  = ", sCdval1)
End Function


Public Function SqlInsertMicQcRstGroup(ByVal sCdval1 As String, ByVal sCdval2 As String, _
                                       ByVal sField1 As String, ByVal sField2 As String, _
                                       ByVal sField3 As String, ByVal sField4 As String, _
                                       ByVal sText1 As String) As String
        
    SqlInsertMicQcRstGroup = " insert into " & T_LAB031 & _
                             "   (cdindex, cdval1, cdval2, field1, field2, field3, field4, text1) " & _
                             " values " & _
                             "(" & _
                                  DBV("cdindex", LC2_MicTemp, 1) & DBV("cdval1", sCdval1, 1) & _
                                  DBV("cdval2", sCdval2, 1) & DBV("field1", sField1, 1) & _
                                  DBV("field2", sField2, 1) & DBV("field3", sField3, 1) & _
                                  DBV("field4", sField4, 1) & DBV("text1", sText1) & _
                             ")"

End Function


'----------------------------------------------------------------------------------
'6. 미생물 QC 결과등록 관련 Sql
'----------------------------------------------------------------------------------


Public Function SqlFindMicQcDate(ByVal pGroupCd As String) As String
    
    SqlFindMicQcDate = " SELECT rstdt FROM " & T_LAB029 & _
                       " WHERE " & DBW("grpcd=", pGroupCd) & _
                       " group by rstdt order by rstdt desc"

End Function

Public Function SqlLastDataExist(ByVal pGroupCd As String, ByVal pRstDt As String, _
                                 ByVal pQcKind As String) As String
    
    If pQcKind = "1" Then '미생물QC
        SqlLastDataExist = " SELECT max(rstdt) rstdate FROM " & T_LAB029 & _
                           " WHERE " & DBW("grpcd = ", pGroupCd) & _
                           " AND   " & DBW("rstdt < ", pRstDt)
    Else    '혈액은행QC
        SqlLastDataExist = " SELECT max(rstdt) rstdate FROM " & T_LAB030 & _
                           " WHERE " & DBW("rstdt < ", pRstDt)
    End If
    
End Function

Public Function SqlLastDataShow(ByVal pTempKey As String, ByVal pRstDate As String) As String
    
    SqlLastDataShow = " SELECT a.*, b.text1 as micnm, c.text1 as rsttemp, d.empnm " & _
                      " FROM  " & T_LAB029 & " a, " & T_LAB032 & " b, " & _
                                  T_LAB031 & " c, " & T_LAB015 & " d  " & _
                      " WHERE " & DBW("a.grpcd  = ", pTempKey) & _
                      " AND   " & DBW("a.rstdt  = ", pRstDate) & _
                      " AND   " & DBJ(DBW("b.cdindex= ", LC3_Microbe)) & _
                      " AND   " & DBJ("b.cdval1 =* a.miccd") & _
                      " AND   " & DBJ(DBW("c.cdindex= ", LC2_MicTemp)) & _
                      " AND   " & DBJ("c.cdval1 =* a.grpcd") & _
                      " AND   " & DBJ("c.cdval2 =* a.miccd") & _
                      " AND   " & DBJ("d.empid  =* a.vfyid")

End Function


Public Function SqlTemplateShow(ByVal pTempKey As String) As String
        
    SqlTemplateShow = " SELECT a.*, b.text1 as micnm " & _
                      " FROM  " & T_LAB031 & " a, " & T_LAB032 & " b " & _
                      " WHERE " & DBW("a.cdindex = ", LC2_MicTemp) & _
                      " AND   " & DBW("a.cdval1  = ", pTempKey) & _
                      " AND   " & DBJ(DBW("b.cdindex = ", LC3_Microbe)) & _
                      " AND   " & DBJ("b.cdval1  =* a.cdval2")

End Function

Public Function SqlDeleteMicQcResult(ByVal pGroupKey As String, ByVal pRstDate As String) As String
        
    SqlDeleteMicQcResult = " delete " & T_LAB029 & _
                           " WHERE  " & DBW("grpcd = ", pGroupKey) & _
                           " AND    " & DBW("rstdt = ", pRstDate)

End Function

Public Function SqlInsertMicQcResult(ByVal sGrpCd As String, ByVal sRstDt As String, _
                                     ByVal sSeq As String, ByVal sMicCd As String, _
                                     ByVal sVfyId As String, ByVal sRstText As String) As String
        
    SqlInsertMicQcResult = " insert into " & T_LAB029 & _
                           "   (grpcd, rstdt, seq, miccd, vfyid, rsttext) " & _
                           " values " & _
                           " ( " & _
                                DBV("grpcd", sGrpCd, 1) & DBV("rstdt", sRstDt, 1) & _
                                DBV("seq", sSeq, 1) & DBV("miccd", sMicCd, 1) & _
                                DBV("vfyid", sVfyId, 1) & DBV("rsttext", sRstText) & _
                           " ) "

End Function


'----------------------------------------------------------------------------------
'7. 혈액은행 QC 결과등록 관련 Sql
'----------------------------------------------------------------------------------

Public Function SqlFindBldQcDate() As String
    
    SqlFindBldQcDate = " SELECT rstdt FROM " & T_LAB030 & _
                       " group by rstdt order by rstdt desc"
End Function

Public Function SqlGetBldQcResult(ByVal pRstDate As String) As String
    SqlGetBldQcResult = " SELECT a.*, b.empnm " & _
                        " FROM  " & T_LAB030 & " a, " & T_LAB015 & " b " & _
                        " WHERE " & DBW("a.rstdt =", pRstDate) & _
                        " AND   " & DBJ("b.empid =* a.vfyid")
End Function

Public Function SqlDeleteBldQcResult(ByVal pRstDate As String) As String
        
    SqlDeleteBldQcResult = " delete " & T_LAB030 & _
                           " WHERE    " & DBW("rstdt = ", pRstDate)

End Function

Public Function SqlInsertBldQcResult(ByVal sRstDt As String, ByVal sSeq As String, _
                                     ByVal sVfyId As String, ByVal sRstCd As String, _
                                     ByVal sTitle As String, ByVal sRemark As String) As String
        
    SqlInsertBldQcResult = " insert into " & T_LAB030 & _
                           " (rstdt, seq, vfyid, rstcd, title, remark) " & _
                           " values " & _
                           " ( " & _
                                DBV("rstdt", sRstDt, 1) & DBV("seq", sSeq, 1) & _
                                DBV("vfyid", sVfyId, 1) & DBV("rstcd", sRstCd, 1) & _
                                DBV("title", sTitle, 1) & DBV("remark", sRemark) & _
                           " ) "

End Function


Public Function SqlGetBldTemplate() As String
    
    SqlGetBldTemplate = " SELECT cdval1, field1 FROM " & T_LAB032 & _
                        " WHERE " & DBW("cdindex = ", LC3_BldQcRst) & _
                        " order by cdval1"

End Function

'------------------------------------------------------------------------------------------------
'Coding By Legends


Public Function GetCtrlMst() As String
'FROM LAB021
'QC Cotrol Master

    GetCtrlMst = " SELECT * FROM " & T_LAB021 & _
                 " order by ctrlcd "
End Function

Public Function GetEqpMst(Optional ByVal pEqpCd As String = "") As String
'FROM LAB006
'장비 마스터
    GetEqpMst = " SELECT eqpcd, eqpnm FROM " & T_LAB006
    If pEqpCd <> "" Then
        GetEqpMst = GetEqpMst & " WHERE " & DBW("eqpcd=", pEqpCd)
    End If
    
    GetEqpMst = GetEqpMst & " order by eqpcd "
End Function

Public Function GetSection() As String
'FROM LAB032
'Section

    GetSection = " SELECT cdval1 as SectCd , field1 as SectNm FROM " & T_LAB032 & _
                 " WHERE " & DBW("cdindex =", LC3_Section)
End Function

Public Function GetBuilding() As String
'FROM LAB032
'건물정보

    GetBuilding = " SELECT * FROM " & T_LAB032 & _
                  " WHERE " & DBW("cdindex=", LC3_Buildings) & _
                  " order by field2 asc"
End Function

Public Function GetWorkArea() As String
'FROM LAB032
'WorkArea

    GetWorkArea = " SELECT cdval1, field1 " & _
                  " FROM " & T_LAB032 & _
                  " WHERE " & DBW("cdindex = ", lc3_workarea)
End Function

Public Function GetTestItem(ByVal pWorkArea As String, ByVal pSpcFG As Boolean) As String
'FROM LAB001, LAB004

    If Not pSpcFG Then
        GetTestItem = " SELECT distinct a.testcd, a.testnm FROM " & T_LAB001 & " a " & _
                      " WHERE (a.panelfg is null or a.panelfg=' ') "
    Else
        GetTestItem = " SELECT distinct a.testcd, a.testnm, a.rsttype, a.rstdiv, b.rstunit, b.avalval" & _
                      " FROM " & T_LAB001 & " a, " & T_LAB004 & " b " & _
                      " WHERE (a.detailfg = '' or a.detailfg is null) " & _
                      " AND   (a.panelfg is null or " & DBW("a.panelfg = ", PN_Detail) & ")" & _
                      " AND   b.testcd = a.testcd  " & _
                      " AND  (b.expdt = '' or b.expdt is null) " & _
                      " AND   b.seq = (SELECT max(seq) FROM  " & T_LAB004 & _
                                      " WHERE testcd = b.testcd " & _
                                      " AND (b.expdt = '' or b.expdt is null)) "
    End If
    
    If pWorkArea <> "" Then
        GetTestItem = GetTestItem & " AND " & DBW("a.workarea =", pWorkArea)
    End If
End Function


