VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHosComSQLStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'Private objConstants    As New clsHosComConstants
'Private objTables       As New clsTables
'Private objFields       As New clsFields
'
'Public Function LoadCOM004ResultTemp()
'    'com004 진단병리 결과 템플릿 Load
'    LoadCOM004ResultTemp = " SELECT * FROM " & objTables.T_COM004 & _
'                           "  WHERE cdindex like 'A%' "
'End Function
'
'Public Function LoadPtid() As String
'    '환자정보 Load
'    LoadPtid = " SELECT " & objFields.F_PTID & " as ptid , " & objFields.F_PTNM & " as ptnm  " & _
'               "   FROM " & objTables.T_HIS001
'End Function
'
'Public Function LoadBuilding() As String
'    '건물정보 Load
'    LoadBuilding = " SELECT cdval1 as buildcd, field1 as buildnm, field2 as buildno " & _
'                   "   FROM " & objTables.T_LAB032 & _
'                   "  WHERE " & DBW("cdindex", "C230", 2)
'End Function
'
'Public Function LoadPthDoctId() As String
'    '판독의정보 Load
'    LoadPthDoctId = " SELECT cdval1 as pthdoct, field1 as pthdoctnm " & _
'                    "   FROM " & objTables.T_COM003 & _
'                    "  WHERE " & DBW("cdindex", "A351", 2)
'End Function
'
'Public Function LoadScreenId() As String
'    '스크린 의사 정보 Load
'    LoadScreenId = " SELECT cdval1 as screenid, field1 as screennm " & _
'                   "   FROM " & objTables.T_COM003 & _
'                   "  WHERE " & DBW("cdindex", "A350", 2)
'End Function
'
'Public Function LoadResId() As String
'    '레지던트 의사 정보 Load
'    LoadResId = " SELECT cdval1 as resid, field1 as resnm " & _
'                "   FROM " & objTables.T_COM003 & _
'                "  WHERE " & DBW("cdindex", "A353", 2)
'End Function
'
'Public Function LoadEmpId(Optional pDeptCd As String) As String
'    '직원정보 Load
'    LoadEmpId = " SELECT * FROM " & objTables.T_COM006
'
'End Function
'
'Public Function SqlTcdMcdOcd(ByVal pTable As String, Optional ByVal pFleld As String) As String
''조직코드, 진단코드, 채취방법코드를 얻어온다.
'    SqlTcdMcdOcd = " SELECT * FROM " & pTable & _
'                   " WHERE (expdt='' or expdt is null) "
'End Function
'
'Public Function LoadICD() As String
''상병 코드
'    LoadICD = " select   " & objFields.F_ICD & " as icd, " & _
'                             objFields.F_IENM & " as ienm " & _
'              "   from   " & objTables.T_HIS006 & _
'              " order by " & objFields.F_ICD
'End Function
'
'Public Function LoadOCD() As String
''수술 코드
'
'    LoadOCD = " select   " & objFields.F_OCD & " as ocd, " & _
'                             objFields.F_ONM & " as onm " & _
'              "   from   " & objTables.T_HIS007 & _
'              " order by " & objFields.F_OCD
'End Function
'
'Public Function SqlPtntInfo(ByVal PtntId As String) As String
''Get the Patient's Infromation : 환자정보
'   SqlPtntInfo = " SELECT * " & _
'                 "   FROM   " & objTables.T_HIS001 & _
'                 "   WHERE  " & DBW(objFields.F_PTID & " = ", PtntId)
'End Function
'
'
'
'Public Function GetCollectionList(ByVal strCase As String, ByVal SearchKey As String, _
'                                  Optional ByVal strWardId As String = "") As String
'
'    Dim sSql As String
'
'    sSql = "SELECT distinct a." & objFields.F_PTID & " as ptid, a." & objFields.F_PTNM & " as ptnm," & _
'                                  objFields.F_SSN2("a") & " as SSN, " & _
'                                  objFields.F_DOB2("a") & " as DOB " & _
'            " FROM " & objTables.T_HIS001 & " a ," & objTables.T_LAB101 & " b," & objTables.T_LAB102 & " c " & _
'            " WHERE " & DBW("b.donefg =", "0") & _
'            " AND   " & DBW("b.bussdiv =", "2") & _
'            " AND    a." & objFields.F_PTID & " = b.ptid"
'
'   If Val(strCase) Mod 2 = 1 Then
'        sSql = sSql & " AND a." & objFields.F_PTID & " >= " & DBV(objFields.F_PTID, SearchKey) & " "
'    Else
'        sSql = sSql & " AND a." & objFields.F_PTNM & " like '" & SearchKey & "%' "
'    End If
'
'    sSql = sSql & " AND b.wardid = '" & strWardId & "'" & _
'                " AND b.ptid=c.ptid" & _
'                " AND b.orddt=c.orddt" & _
'                " AND b.ordno=c.ordno" & _
'                " AND (c.dcfg='' or c.dcfg is null)"
'
'
'    If strCase = 1 Then
'        sSql = sSql & " ORDER BY a." & objFields.F_PTID
'    Else
'        sSql = sSql & " ORDER BY a." & objFields.F_PTNM & ", SSN"
'    End If
'
'    GetCollectionList = sSql
'
'End Function
'
'
'Public Function SqlPtntSearch(ByVal strCase As String, ByVal SearchKey As String, _
'                            Optional ByVal strWardId As String = "", _
'                            Optional ByVal strVerifyDt As String = "") As String
'' 환자Id 또는 환자명을 입력받아 검색한다.
'
'    Dim strPtNoFormat As String
'    Dim QField        As String
'
'    strPtNoFormat = "#"
'    QField = ",'' as telno "
'
'' strCase : 1-환자ID, 2-환자명, 3-환자ID(외래채혈대상), 4-환자명(외래채혈대상), 5-환자ID(병동채혈대상), 6-환자명(병동채혈대상)
'
'    SqlPtntSearch = "Select distinct a." & objFields.F_PTID & " as ptid, a." & objFields.F_PTNM & " as ptnm, " & _
'                            objFields.F_SSN2("a") & " as SSN, " & _
'                            objFields.F_DOB2("a") & " as DOB, " & _
'                    "       a." & objFields.F_ADDRESS & " as address " & QField & _
'                    " from " & objTables.T_HIS001 & " a "
'
'
'    If strVerifyDt <> "" Then   '결과보고대기내역
'
'        If Val(strCase) Mod 2 = 1 Then
'            'SqlPtntSearch = SqlPtntSearch & _
'                           "Where  a." & objFields.F_PTID & " between " & DBV(objFields.F_PTID, Format(SearchKey, strPtNoFormat)) & _
'                                                            " and " & DBV(objFields.F_PTID, Format((SearchKey + 1000), strPtNoFormat)) & " "
'            SqlPtntSearch = SqlPtntSearch & _
'                           "Where  a." & objFields.F_PTID & " >= " & DBV(objFields.F_PTID, SearchKey) & " "
'        Else
'            SqlPtntSearch = SqlPtntSearch & _
'                           "Where  a." & objFields.F_PTNM & " like '" & SearchKey & "%' "
'        End If
'
'        If strWardId <> "" Then
'            SqlPtntSearch = SqlPtntSearch & " and exists (select * from " & objTables.T_LAB202 & " b " & _
'                          " where b.deptcd = '" & strWardId & "'  and  b.vfydt = '" & strVerifyDt & "' " & _
'                          " and b.ptid = a." & objFields.F_PTID & " ) "
'        End If
'
'    Else    '환자검색(채혈대상)
'
'        If Val(strCase) > 2 Then    '채혈대상
'            SqlPtntSearch = SqlPtntSearch & ", " & objTables.T_LAB101 & " b "
'        End If
'
'        If strCase = "3" Or strCase = "4" Then  '외래
'            SqlPtntSearch = SqlPtntSearch & _
'                            " WHERE  " & DBW("b.donefg =", "0") & _
'                            " and " & DBW("b.bussdiv =", "1") & _
'                            " and " & DBW("b.reqdt>=", Format(DateAdd("m", -6, Now), "YYYYMMDD")) & _
'                            " and a." & objFields.F_PTID & " = b.ptid  and "
'
'        ElseIf strCase = "5" Or strCase = "6" Then  '병동
'            SqlPtntSearch = SqlPtntSearch & _
'                            " WHERE  " & DBW("b.donefg =", "0") & _
'                            " and " & DBW("b.bussdiv =", "2") & _
'                            " and " & DBW("b.reqdt>=", Format(DateAdd("m", -6, Now), "YYYYMMDD")) & _
'                            " and a." & objFields.F_PTID & " = b.ptid  and "
'
'            'SqlPtntSearch = SqlPtntSearch & " WHERE  b.donefg = '0' and b.bussdiv = '2' and a." & objFields.F_PTID & " = b.ptid  and "
'        Else
'            SqlPtntSearch = SqlPtntSearch & " WHERE  "
'        End If
'
'        If Val(strCase) Mod 2 = 1 Then
'            If strCase = "99" Then
'                SqlPtntSearch = SqlPtntSearch & _
'                               " a." & objFields.F_PTID & " = " & DBV(objFields.F_PTID, SearchKey) & " "
'            Else
'                SqlPtntSearch = SqlPtntSearch & _
'                               " a." & objFields.F_PTID & " >= " & DBV(objFields.F_PTID, SearchKey) & " "
'            End If
'
'            'SqlPtntSearch = SqlPtntSearch & _
'                           " a." & objFields.F_PTID & " between " & DBV(objFields.F_PTID, Format(SearchKey, strPtNoFormat)) & _
'                                                            " and " & DBV(objFields.F_PTID, Format((SearchKey + 1000), strPtNoFormat)) & " "
'        Else
'            SqlPtntSearch = SqlPtntSearch & _
'                           " a." & objFields.F_PTNM & " like '" & SearchKey & "%' "
'        End If
'
'        If strWardId <> "" Then
'            SqlPtntSearch = SqlPtntSearch & " and b.wardid = '" & strWardId & "' "
'        End If
'
'    End If
'
'
'    If strCase = 1 Then
'        SqlPtntSearch = SqlPtntSearch & " Order  by a." & objFields.F_PTID & ""
'    Else
'        SqlPtntSearch = SqlPtntSearch & " Order  by a." & objFields.F_PTNM & ", SSN"
'    End If
'End Function
'
'Public Function LoadWardId() As String
'    LoadWardId = " select distinct " & objFields.F_WARDID & " as wardid, " & _
'                          objFields.F_WARDID & " as wardnm, '10' as bldgb, " & _
'                 "        '본원' as bldnm, 1 as bldno " & _
'                 "   from " & objTables.T_HIS004
'End Function
'
'Public Function LoadDeptCd() As String
'    LoadDeptCd = " select distinct " & objFields.F_DEPTCD & " as deptcd, " & _
'                          objFields.F_DEPTNM & " as deptnm " & _
'                 "   from " & objTables.T_HIS003 & _
'                 "  where " & objFields.F_DEPTDIV & " ='2' "
'End Function
'
'Public Function LoadDoctId() As String
''처방의정보 Load
'    LoadDoctId = " select distinct " & objFields.F_DOCTID & " as doctid, " & _
'                          objFields.F_DOCTNM & " as doctnm " & _
'                 "   from " & objTables.T_HIS005
'End Function
'
'Public Function SqlDeptList(Optional ByVal DEPTCD As Variant, Optional ByVal DeptClass As String = "B") As String
''Get Department List : 부서리스트
'
'    DeptClass = "2"
'
'    If IsMissing(DEPTCD) Then
'       SqlDeptList = " Select a." & objFields.F_DEPTCD & " as DeptCd, a." & objFields.F_DEPTNM & " as DeptNm " & _
'                     " From " & objTables.T_HIS003 & " a " & _
'                     " Where  a." & objFields.F_DEPTDIV & " = '" & DeptClass & "' " & _
'                     " Order by a." & objFields.F_DEPTNM
'    Else
'       SqlDeptList = " Select a." & objFields.F_DEPTCD & " as DeptCd, a." & objFields.F_DEPTNM & " as DeptNm, a.center_gb as BldGb " & _
'                     " From " & objTables.T_HIS003 & " a " & _
'                     " where  a." & objFields.F_DEPTCD & " = '" & DEPTCD & "' "
'    End If
'
'End Function
'
'Public Function SqlWardList(Optional ByVal WARDID As Variant) As String
''Get Ward Id List : 병동 리스트
'    If IsMissing(WARDID) Then
'        SqlWardList = " Select distinct a." & objFields.F_WARDID & " as WardId, " & _
'                      "                 a." & objFields.F_WARDID & " as WardNm " & _
'                      " From " & objTables.T_HIS004 & " a " & _
'                      " Order by WardId "
'   Else
'        SqlWardList = " Select a." & objFields.F_WARDID & " as WardId, a." & objFields.F_WARDID & " as WardNm " & _
'                      " From " & objTables.T_HIS004 & " a " & _
'                      " Where  a." & objFields.F_WARDID & " = '" & WARDID & "' "
'   End If
'End Function
'
'Public Function SqlDoctList(Optional ByVal EMPID As Variant) As String
''Get Doctor List : 의사Id 리스트
'
'    If IsMissing(EMPID) Then
'        SqlDoctList = " Select a." & objFields.F_DOCTID & " as EmpId, a." & objFields.F_DOCTNM & " as EmpNm " & _
'                      " From " & objTables.T_HIS005 & " a " & _
'                      " Order by a." & objFields.F_DOCTNM
'    Else
'        SqlDoctList = " Select a." & objFields.F_DOCTID & " as EmpId, a." & objFields.F_DOCTNM & " as EmpNm " & _
'                      " From " & objTables.T_HIS005 & " a " & _
'                      " where  " & DBW("a." & objFields.F_DOCTID & " = ", EMPID)
'    End If
'End Function
'
'Public Function SqlEmpList(Optional ByVal pEmpId As Variant, Optional ByVal pDeptCd As Variant) As String
''직원
'
'    If IsMissing(pEmpId) Then
'        SqlEmpList = " Select empid , empnm " & _
'                     " From " & objTables.T_COM006 & _
'                     " Order by empnm "
'    Else
'        SqlEmpList = " Select empid , empnm " & _
'                     " From " & objTables.T_COM006 & _
'                     " where  " & DBW("empid = ", pEmpId) & _
'                     " and    deptcd ='" & pDeptCd & "'"
'    End If
'End Function
'
'Public Function SqlItemList() As String
''임상병리 처방입력시 사용될 검사항목 리스트
'    SqlItemList = " SELECT a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
'                  "        b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg, c.field2 as SpcGrp, c.field5 as SpcNm, " & _
'                  "        d.field2 as LabDiv, e.field2 as LabRange, '1' InsurFg " & _
'                  " FROM " & objTables.T_LAB032 & " c, " & objTables.T_LAB032 & " d, " & objTables.T_LAB032 & " e, " & _
'                             objTables.T_LAB004 & " b, " & objTables.T_LAB001 & " a " & _
'                  " WHERE  a.applydt = ( SELECT max(applydt) FROM " & objTables.T_LAB001 & _
'                  "                     WHERE testcd = a.testcd ) " & _
'                  " AND   (a.detailfg = '' or a.detailfg is null) " & _
'                  " AND    a.testcd = b.testcd " & _
'                  " AND    b.seq = ( SELECT min(seq) FROM " & objTables.T_LAB004 & _
'                  "                  WHERE testcd = b.testcd ) " & _
'                  " AND   (b.expdt = '' or b.expdt is null)" & _
'                  " AND    b.applydt = ( SELECT max(applydt) FROM " & objTables.T_LAB004 & _
'                  "                      WHERE testcd = b.testcd AND spccd = b.spccd AND seq=b.seq) " & _
'                  " AND    c.cdindex = 'C215' " & _
'                  " AND    c.cdval1 = b.spccd  " & _
'                  " AND    d.cdindex = 'C213' " & _
'                  " AND    d.cdval1 = a.workarea " & _
'                  " AND    " & DBJ("e.cdindex = 'C217'") & _
'                  " AND    " & DBJ("e.cdval1 =* c.field2")
'
'End Function
'
'Public Function SqlSpcList() As String
''임상병리 검체정보 리스트
'    SqlSpcList = "SELECT a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm, " & _
'                 "       a.field1 multifg, a.field2 spcgrp, b.field2 labrange  " & _
'                  "FROM " & objTables.T_LAB032 & " b, " & objTables.T_LAB032 & " a " & _
'                  "WHERE  a.cdindex = 'C215' " & _
'                  "AND    " & DBJ("b.cdindex = 'C217'") & _
'                  "AND    " & DBJ("b.cdval1  =* a.field2")
'
'End Function
