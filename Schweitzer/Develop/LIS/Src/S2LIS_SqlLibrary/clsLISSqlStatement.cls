VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'+------------------------------------------------------+
'|    Sql Query for Orders, Collections, Accession      |
'|                                                      |
'|    작성자        : 김미경                            |
'|    작성일        : 1999.04.18                        |
'|                                                      |
'|  CopyRight(C) 1999 POMIS                      |
'+------------------------------------------------------+

'#Const APS_FLAG = True
'#Const BBS_FLAG = True
Private mvarInOut       As String
Private mvarOrderDate   As String
Private mvarPayDt       As String
Private mvarWorkarea As String

Public Property Let PayDt(ByVal vData As String)
    mvarPayDt = vData
End Property

Public Property Let InOut(ByVal vData As String)
    mvarInOut = vData
End Property

Public Property Let OrderDate(ByVal vData As String)
    mvarOrderDate = vData
End Property

'%  0. 공통코드 마스터 (LAB031)
Public Function SqlCommonCode(ByVal TblName As String, ByVal cdindex As String, _
                            Optional ByVal CdVal1 As Variant, Optional ByVal CdVal2 As Variant) As String
    Dim tmpStr As String
    
    tmpStr = ""
    If Not IsMissing(CdVal1) Then tmpStr = tmpStr & "AND " & DBW("cdval1", CdVal1, 2)
    If Not IsMissing(CdVal2) Then tmpStr = tmpStr & "AND " & DBW("cdval2", CdVal2, 2)
    
    SqlCommonCode = " SELECT * " & _
                    " FROM  " & TblName & _
                    " WHERE " & DBW("cdindex", cdindex, 2) & tmpStr
End Function

'%  1. Get CodeList
'%      - Calling FROM [frm351ItemMaster]
Public Function SqlLAB001CodeList() As String
    SqlLAB001CodeList = " SELECT a.testcd, a.testnm, a.ABBRNM5, a.ABBRNM10 " & _
                        " FROM  " & T_LAB001 & " a " & _
                        " WHERE a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & "  " & _
                        "                              WHERE testcd = a.testcd ) " & _
                        " Order by a.testcd "
End Function

'%  2. SELECT FROM LAB001
'%      - Calling FROM [frm351ItemMaster]
Public Function SqlLAB001Read(ByVal TestCd As String) As String
    SqlLAB001Read = " SELECT * " & _
                    " FROM  " & T_LAB001 & _
                    " WHERE " & DBW("testcd", TestCd, 2) & _
                    " Order by applydt"
End Function

'%  3. SELECT FROM LAB004
'%      - Calling FROM [frm351ItemMaster]
Public Function SqlSpecimenRead(ByVal TestCd As String) As String

    SqlSpecimenRead = " SELECT a.spccd as spccd, b.field3 as spcnm, b.field2 as spcgrpcd, c.field1 as spcgrp, a.seq " & _
                      " FROM  " & T_LAB004 & " a, " & T_LAB032 & " b, " & T_LAB032 & " c " & _
                      " WHERE " & DBW("a.testcd", TestCd, 2) & _
                      " AND   " & DBW("b.cdindex", LC3_Specimen, 2) & _
                      " AND     b.cdval1 = a.spccd " & _
                      " AND   " & DBJ(DBW("c.cdindex", LC3_SGroup, 2)) & _
                      " AND   " & DBJ("c.cdval1 =* b.field2") & _
                      " Group by a.spccd, b.field3, b.field2, c.field1, a.seq " & _
                      " Order by seq"
End Function

'%  4. SELECT FROM LAB004
'%      - Calling FROM [frm351ItemMaster]
Public Function SqlLAB004Read(ByVal TestCd As String, ByVal SpcCd As String, _
                              ByVal Seq As String) As String
    SqlLAB004Read = " SELECT a.*, b.field2 as sgroup " & _
                    " FROM  " & T_LAB004 & " a, " & T_LAB032 & " b " & _
                    " WHERE " & DBW("a.testcd", TestCd, 2) & _
                    " AND   " & DBW("a.spccd", SpcCd, 2) & _
                    " AND   " & DBW("a.seq", Seq, 2) & _
                    " AND   " & DBJ(DBW("b.cdindex", LC3_Specimen, 2)) & _
                    " AND   " & DBJ("b.cdval1 =* a.spccd") & _
                    " Order by applydt"
End Function

'%  5. Get Common Code List
'%      - Calling FROM [frm352ItemMaster]
Public Function SqlLAB031CodeList(ByVal cdindex As String, ByVal Fields As String, _
                                  Optional ByVal CdVal1 As Variant, _
                                  Optional ByVal CdVal2 As Variant, _
                                  Optional ByVal Orderby As Variant) As String
    SqlLAB031CodeList = " SELECT " & Fields & _
                        " FROM  " & T_LAB031 & _
                        " WHERE " & DBW("cdindex", cdindex, 2)
    If IsMissing(CdVal1) = False Then
       SqlLAB031CodeList = SqlLAB031CodeList & _
                        " AND    " & DBW("cdval1", CdVal1, 2)
    End If
    If IsMissing(CdVal2) = False Then
       SqlLAB031CodeList = SqlLAB031CodeList & _
                        " AND    " & DBW("cdval2", CdVal2, 2)
    End If
    
    '** 특수검사 시 정렬 변경 By M.G.Choi 2008.04.17
    If cdindex = LC2_SpeTemp Then
        SqlLAB031CodeList = SqlLAB031CodeList & " order by cdval1, cdval2, field1 "
    Else
        If IsMissing(Orderby) = False Then
           SqlLAB031CodeList = SqlLAB031CodeList & Orderby
        End If
    End If
    
End Function

Public Function SqlLAB032CodeList(ByVal cdindex As String, ByVal Fields As String, _
                                  Optional ByVal CdVal1 As Variant, _
                                  Optional ByVal Orderby As Variant, _
                                  Optional ByVal strWhere As Variant) As String
    SqlLAB032CodeList = " SELECT " & Fields & _
                        " FROM  " & T_LAB032 & _
                        " WHERE " & DBW("cdindex", cdindex, 2)
    If IsMissing(CdVal1) = False Then
       SqlLAB032CodeList = SqlLAB032CodeList & _
                        " AND    " & DBW("cdval1", CdVal1, 2)
    End If
    If IsMissing(strWhere) = False Then
       SqlLAB032CodeList = SqlLAB032CodeList & " " & strWhere & " "
    End If
    If IsMissing(Orderby) = False Then
       SqlLAB032CodeList = SqlLAB032CodeList & "ORDER BY " & Orderby
    End If

End Function

'%  5. Get Text Template List
'%      - Calling FROM [frm352ItemMaster]
Public Function SqlLAB034CodeList(ByVal cdindex As String, ByVal Fields As String, _
                                  Optional ByVal CdVal1 As Variant, _
                                  Optional ByVal Orderby As Variant) As String
    Select Case cdindex
        Case LC4_TextResult, LC4_ClinicalNotice, LC4_TestItemComment
            SqlLAB034CodeList = " SELECT  a.cdindex,a.cdval1,a.field1,a.text1,a.text2 ,b.testnm,b.abbrnm10" & _
                                " FROM  " & T_LAB034 & " a," & T_LAB001 & " b" & _
                                " WHERE " & DBW("a.cdindex", cdindex, 2) & _
                                " AND a.cdval1=b.testcd"
            If IsMissing(CdVal1) = False Then
               SqlLAB034CodeList = SqlLAB034CodeList & _
                                " AND    " & DBW("a.cdval1", CdVal1, 2)
            End If
            If IsMissing(Orderby) = False Then
               SqlLAB034CodeList = SqlLAB034CodeList & Orderby
            End If
        Case Else
            SqlLAB034CodeList = " SELECT  cdindex,cdval1,field1,text1,text2 " & _
                                " FROM  " & T_LAB034 & _
                                " WHERE " & DBW("cdindex", cdindex, 2)
            If IsMissing(CdVal1) = False Then
               SqlLAB034CodeList = SqlLAB034CodeList & _
                                " AND    " & DBW("cdval1", CdVal1, 2)
            End If
            If IsMissing(Orderby) = False Then
               SqlLAB034CodeList = SqlLAB034CodeList & Orderby
            End If
            SqlLAB034CodeList = SqlLAB034CodeList & " order by cdval1 "
    End Select
    
End Function

Public Function SqlDeleteLAB034(ByVal cdindex As String, _
                                Optional ByVal CdVal1 As Variant) As String
    SqlDeleteLAB034 = " Delete " & T_LAB034 & _
                      " WHERE  " & DBW("cdindex", cdindex, 2)
    If IsMissing(CdVal1) = False Then
       SqlDeleteLAB034 = SqlDeleteLAB034 & _
                        " AND    " & DBW("cdval1", CdVal1, 2)
    End If

End Function

Public Function SqlSaveLAB034(ByVal pCdindex As String, _
                              ByVal pCdVal1 As Variant, ByVal pField1 As String, _
                              ByVal pText1 As String, ByVal pText2 As String, _
                              ByVal pFlag As Long) As String
    
    If pFlag = 1 Then 'insert
        SqlSaveLAB034 = " INSERT INTO " & T_LAB034 & "(cdindex, cdval1, field1, text1, text2)" & _
                        " VALUES (" & DBV("cdindex", pCdindex, 1) & DBV("cdval1", pCdVal1, 1) & _
                                      DBV("field1", pField1, 1) & DBV("text1", pText1, 1) & _
                                      DBV("text2", pText2) & ")"
    Else    'update
        SqlSaveLAB034 = " UPDATE " & T_LAB034 & _
                        " SET    " & DBW("field1  ", pField1, 3) & _
                                     DBW("text1   ", pText1, 3) & _
                                     DBW("text2   ", pText2, 2) & _
                        " WHERE  " & DBW("cdindex =", pCdindex) & _
                        " AND    " & DBW("cdval1  =", pCdVal1)
    End If

End Function

Public Function SqlDeleteLAB031(ByVal cdindex As String, _
                                Optional ByVal CdVal1 As Variant, _
                                Optional ByVal CdVal2 As Variant) As String
    SqlDeleteLAB031 = " Delete " & T_LAB031 & _
                      " WHERE  " & DBW("cdindex", cdindex, 2)
    If IsMissing(CdVal1) = False Then
       SqlDeleteLAB031 = SqlDeleteLAB031 & _
                        " AND    " & DBW("cdval1", CdVal1, 2)
    End If
    If IsMissing(CdVal2) = False Then
       SqlDeleteLAB031 = SqlDeleteLAB031 & _
                        " AND    " & DBW("cdval2", CdVal2, 2)
    End If

End Function

Public Function SqlSaveLAB031(ByVal pCdindex As String, ByVal pCdVal1 As Variant, _
                              ByVal pCdVal2 As Variant, ByVal pField1 As String, _
                              ByVal pField2 As String, ByVal pField3 As String, _
                              ByVal pField4 As String, ByVal pField5 As String, _
                              ByVal pText1 As String, ByVal pText2 As String, _
                              ByVal pFlag As Long) As String
    
    If pFlag = 1 Then 'insert
        SqlSaveLAB031 = " INSERT INTO " & T_LAB031 & "(cdindex, cdval1, cdval2, field1, field2, field3, field4, field5, text1, text2)" & _
                        " VALUES (" & DBV("cdindex", pCdindex, 1) & DBV("cdval1", pCdVal1, 1) & _
                                      DBV("cdval2", pCdVal2, 1) & DBV("field1", pField1, 1) & DBV("field2", pField2, 1) & _
                                      DBV("field3", pField3, 1) & DBV("field4", pField4, 1) & _
                                      DBV("field5", pField5, 1) & DBV("text1", pText1, 1) & _
                                      DBV("text2", pText2) & ")"
    Else    'update
        SqlSaveLAB031 = " UPDATE " & T_LAB031 & _
                        " SET    " & DBW("field1  ", pField1, 3) & _
                                     DBW("field2  ", pField2, 3) & _
                                     DBW("field3  ", pField3, 3) & _
                                     DBW("field4  ", pField4, 3) & _
                                     DBW("field5  ", pField5, 3) & _
                                     DBW("text1   ", pText1, 3) & _
                                     DBW("text2   ", pText2, 2) & _
                        " WHERE  " & DBW("cdindex =", pCdindex) & _
                        " AND    " & DBW("cdval1  =", pCdVal1) & _
                        " AND    " & DBW("cdval2  =", pCdVal2)
    End If

End Function

Public Function SqlDeleteLAB032(ByVal cdindex As String, _
                                Optional ByVal CdVal1 As Variant) As String
    SqlDeleteLAB032 = " Delete " & T_LAB032 & _
                      " WHERE  " & DBW("cdindex", cdindex, 2)
    If IsMissing(CdVal1) = False Then
       SqlDeleteLAB032 = SqlDeleteLAB032 & _
                        " AND    " & DBW("cdval1", CdVal1, 2)
    End If
End Function

Public Function SqlSaveLAB032(ByVal pCdindex As String, _
                              ByVal pCdVal1 As Variant, ByVal pField1 As String, _
                              ByVal pField2 As String, ByVal pField3 As String, _
                              ByVal pField4 As String, ByVal pField5 As String, _
                              ByVal pText1 As String, ByVal pText2 As String, _
                              ByVal pFlag As Long) As String
    
    If pFlag = 1 Then 'insert
        SqlSaveLAB032 = " INSERT INTO " & T_LAB032 & "(cdindex, cdval1, field1, field2, field3, field4, field5, text1, text2)" & _
                        " VALUES (" & DBV("cdindex", pCdindex, 1) & DBV("cdval1", pCdVal1, 1) & _
                                      DBV("field1", pField1, 1) & DBV("field2", pField2, 1) & _
                                      DBV("field3", pField3, 1) & DBV("field4", pField4, 1) & _
                                      DBV("field5", pField5, 1) & DBV("text1", pText1, 1) & _
                                      DBV("text2", pText2) & ")"
    Else    'update
        SqlSaveLAB032 = " UPDATE " & T_LAB032 & _
                        " SET    " & DBW("field1  ", pField1, 3) & _
                                     DBW("field2  ", pField2, 3) & _
                                     DBW("field3  ", pField3, 3) & _
                                     DBW("field4  ", pField4, 3) & _
                                     DBW("field5  ", pField5, 3) & _
                                     DBW("text1   ", pText1, 3) & _
                                     DBW("text2   ", pText2, 2) & _
                         " WHERE  " & DBW("cdindex =", pCdindex) & _
                         " AND    " & DBW("cdval1  =", pCdVal1)
    End If
End Function

Public Function SqlSaveSpcGrp(ByVal pSpcCd As String, ByVal pSpcGroupCd As String) As String

    SqlSaveSpcGrp = " UPDATE " & T_LAB032 & _
                    " SET    " & DBW("field2  =", pSpcGroupCd) & _
                    " WHERE  " & DBW("cdindex =", LC3_Specimen) & _
                    " AND    " & DBW("cdval1  =", pSpcCd)

End Function


'%  6. Get ApplyDt FROM LAB005
'%      - Calling FROM [frm353ItemMaster]
Public Function SqlLAB005AppDt(ByVal TestCd As String, ByVal SpcCd As String) As String
    SqlLAB005AppDt = " SELECT applydt " & _
                     " FROM  " & T_LAB005 & _
                     " WHERE " & DBW("testcd", TestCd, 2) & _
                     " AND   " & DBW("spccd", SpcCd, 2) & _
                     " Group by applydt " & _
                     " Order by applydt"
End Function

'%  7. SELECT FROM LAB005
'%      - Calling FROM [frm353ItemMaster]
Public Function SqlLAB005Read(ByVal TestCd As String, ByVal SpcCd As String, _
                                              ByVal ApplyDt As String) As String
    SqlLAB005Read = " SELECT * " & _
                    " FROM  " & T_LAB005 & _
                    " WHERE " & DBW("testcd", TestCd, 2) & _
                    " AND   " & DBW("spccd", SpcCd, 2) & _
                    " AND   " & DBW("applydt", ApplyDt, 2) & _
                    " Order by applysex, agefrom"
End Function


'% 9. Get the Patient's Infromation : 환자정보
'%     - PtntId : 환자 ID
Public Function SqlHIS001Read(ByVal PtntId As String) As String
    SqlHIS001Read = " SELECT * " & _
                    " FROM  " & T_HIS001 & _
                    " WHERE " & DBW(F_PTID, PtntId, 2)
End Function


'%  10. Get Doctor List : 의사Id 리스트
'%      - Calling FROM [frm101Order]
'Public Function SqlHIS007CodeList(Optional ByVal EmpId As Variant) As String
'
'        If IsMissing(EmpId) Then
'           SqlHIS007CodeList = " SELECT a." & F_DOCTID & " as EmpId, a." & F_DOCTNM & " as EmpNm, a.a.doctor_gwa as deptcd " & _
'                               " FROM  " & T_HIS005 & " a " & _
'                               " Order by a." & F_DOCTNM & " "
'        Else
'           SqlHIS007CodeList = " SELECT a." & F_DOCTID & " as EmpId, a." & F_DOCTNM & " as EmpNm,  a.doctor_gwa as deptcd " & _
'                               " FROM  " & T_HIS005 & " a " & _
'                               " WHERE " & DBW("a." & F_DOCTID & "", EmpId, 2)
'        End If
'End Function

'%  11. Get Department List : 부서리스트
'%      - Calling FROM [frm101Order]
'Public Function SqlHIS003CodeList(Optional ByVal DeptCd As Variant, Optional ByVal DeptClass As String = "B") As String
'
'    If IsMissing(DeptCd) Then
'       SqlHIS003CodeList = " Select a." & F_DEPTCD & " as DeptCd, a." & F_DEPTNM & " as DeptNm " & _
'                           " From  " & T_HIS003 & " a " & _
'                           " WHERE " & F_DEPTDIV & " ='2' " & _
'                           " Order by a." & F_DEPTNM
'    Else
'       SqlHIS003CodeList = " Select a." & F_DEPTCD & " as DeptCd, a." & F_DEPTNM & " as DeptNm, a." & F_DEPTDIV & " as BldGb " & _
'                           " From  " & T_HIS003 & " a " & _
'                           " where   " & DBW("a." & F_DEPTCD & "", DeptCd, 2)
'    End If
'End Function

'%  12. Get Ward Id List : 병동 리스트
'%      - Calling FROM [frm101Order]
'Public Function SqlHIS004CodeList(Optional ByVal WardId As Variant) As String
'    If IsMissing(WardId) Then
'        SqlHIS004CodeList = " SELECT distinct " & F_WARDID & " as wardid, " & _
'                                                  F_WARDID & " as wardnm" & _
'                            " FROM " & T_HIS004 & _
'                            " ORDER BY rank"
'   Else
'        SqlHIS004CodeList = " SELECT distinct " & F_WARDID & " as wardid, " & _
'                                            F_WARDID & " as wardnm" & _
'                            " FROM " & T_HIS004 & _
'                            " WHERE " & DBW(F_WARDID, WardId, 2) & _
'                            " ORDER BY rank"
'   End If
'End Function

'%  27-1. 처방Body내역 Update
'%      - Calling FROM [frm156Referral]
Public Function SqlStatusUpdate1(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As Long, _
                                 ByVal TestCd As String, ByVal StsCd As String) As String
    '** 원본 ==========================================================================
'    SqlStatusUpdate1 = " Update " & T_LAB102 & _
'                       " Set    " & DBW("stscd", StsCd, 2) & _
'                       " WHERE  " & DBW("workarea", WorkArea, 2) & _
'                       " AND    " & DBW("accdt", AccDt, 2) & _
'                       " AND    " & DBW("accseq", AccSeq, 2) & _
'                       " AND    " & DBW("ordcd", TestCd, 2)
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim strSQL      As String
    Dim rs          As New ADODB.Recordset
    
    strSQL = " select wrkdiv from " & T_LAB102 & _
             "  where workarea = " & DBS(WorkArea) & _
             "    and accdt = " & DBS(AccDt) & _
             "    and accseq = " & DBS(accseq)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strWrkDiv = "" & rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    rs.Close: Set rs = Nothing
        
    If strWrkDiv = "3" Then         '종건
        SqlStatusUpdate1 = " Update su2examt " & _
                           " Set    " & DBW("stscd", StsCd, 2) & _
                           " WHERE  workarea = " & DBS(WorkArea) & _
                           " AND    accdt = " & DBS(AccDt) & _
                           " AND    accseq = " & DBN(accseq) & _
                           " AND    examcode= " & DBS(TestCd)
    ElseIf strWrkDiv = "4" Then     '일건
            SqlStatusUpdate1 = " Update sg2examt " & _
                               " Set    " & DBW("stscd", StsCd, 2) & _
                               " WHERE  workarea = " & DBS(WorkArea) & _
                               " AND    accdt = " & DBS(AccDt) & _
                               " AND    accseq = " & DBN(accseq) & _
                               " AND    examcode= " & DBS(TestCd)
    Else
        SqlStatusUpdate1 = " Update mdexmort " & _
                           " Set    " & DBW("stscd", StsCd, 2) & _
                           " WHERE  workarea = " & DBS(WorkArea) & _
                           " AND    accdt = " & DBS(AccDt) & _
                           " AND    accseq = " & DBN(accseq) & _
                           " AND    " & DBW("ordcd", TestCd, 2)
    End If
    '===================================================================================
    
End Function


'%  31. 환자Id 또는 환자명을 입력받아 검색한다.
Public Function SqlPtntSearch(ByVal SearchKey As String, ByVal SorkKey As String) As String

    If IsNumeric(SearchKey) Then
        SqlPtntSearch = " SELECT " & F_PTID & " as ptid, " & F_PTNM & " as ptnm, " & _
                                     F_SSN2 & " as SSN FROM " & T_HIS001 & _
                        " WHERE  " & F_PTID & " between " & DBV(F_PTID, SearchKey) & _
                        " AND    " & DBV(F_PTID, SearchKey + 1000) & _
                        " Order by " & SorkKey
    Else
        SqlPtntSearch = " SELECT " & F_PTID & " as ptid, " & F_PTNM & " as ptnm, " & _
                                     F_SSN2 & " as SSN FROM " & T_HIS001 & _
                        " WHERE " & F_PTNM & " like " & DBV(F_PTNM, SearchKey & "%") & _
                        " Order by " & SorkKey & ", SSN"
    End If
End Function


'%  36. 외부의뢰내역 (LAB405)
'%       - Calling FROM [ frm156Referral ] : 외부의뢰
Public Function SqlOutLabData(ByVal OutLabCd As String, ByVal FromDt As String, ByVal toDt As String, Optional ByVal StsCd As Variant) As String

    Dim tmpStr As String
    
    tmpStr = ""
    If Not IsMissing(StsCd) Then tmpStr = " AND " & DBW("a.stscd", StsCd, 2)
    
    SqlOutLabData = " SELECT a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.accseq") & " as LabNo, a.rcvdt, a.ptid, " & _
                    "        c.testnm,d.field3 as SpcNm, e.spccd, " & _
                    "         a.stscd, a.workarea, a.accdt, a.accseq, a.testcd , a.senddt, a.chargedt, " & _
                    "        '' as Gubun , e.wardid,e.deptcd, e.rcvtm " & _
                    " FROM   " & T_LAB205 & " a, " & T_LAB001 & " c, " & _
                                 T_LAB032 & " d, " & T_LAB201 & " e " & _
                    " WHERE  " & DBW("a.outlabcd", OutLabCd, 2) & _
                    " AND    " & DBW("a.rcvdt >= ", FromDt) & _
                    " AND    " & DBW("a.rcvdt <= ", toDt) & tmpStr & _
                    " AND    c.testcd   =  a.testcd " & _
                    " AND    c.applydt  = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd ) " & _
                    " AND    " & DBW("cdindex", LC3_Specimen, 2) & _
                    " AND    d.cdval1   =  a.spccd " & _
                    " AND    a.workarea =  e.workarea " & _
                    " AND    a.accdt    =  e.accdt " & _
                    " AND    a.accseq   =  e.accseq " & _
                    " Order by a.workarea, a.accdt, a.accseq "     '& _

End Function


'%  44. 의뢰내역 Update
'%       - Calling FROM [ frm156Referral ] :
Public Function SqlUpdateLab205(ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As String, _
                                ByVal TestCd As String, ByVal SendDt As String, ByVal SendId As String) As String

    SqlUpdateLab205 = " Update " & T_LAB205 & _
                      " set    " & DBW("senddt", SendDt, 3) & _
                      "        " & DBW("sendid", SendId, 3) & _
                      "        " & DBW("stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                      " WHERE  " & DBW("workarea", WorkArea, 2) & _
                      " AND    " & DBW("accdt", AccDt, 2) & _
                      " AND    " & DBW("accseq", accseq, 2) & _
                      " AND    " & DBW("testcd", TestCd, 2)

End Function


'%  45. 바코드라벨 재출력을 위해 해당 환자의 처방내역을 조회한다.
'%       - Calling FROM [frm157BarReprint] : 라벨재발행
'%       - 검사구분, WorkArea, 검체코드, 보관구분, 응급여부, 희망채혈일시, 복수검체여부, 검체군, 처방일, 처방번호, 처방Seq, 처방코드
'%       - Lab201 : 채혈접수내역
'%       - Lab102 : 처방Body내역
'%       - Lab001 : 검사항목 내역
'%       - Lab032 : 검체명, 건물명
Public Function SqlBarReprint(ByVal intOption As Integer, ParamArray strKeys() As Variant) As String

    Dim tmpStr As String
    Dim tmpStr1 As String
    
    tmpStr = ""
    If intOption = 1 Then  '환자ID-처방일 기준
       tmpStr = tmpStr & DBW("b.ptid", strKeys(0), 2) & _
               " AND " & DBW("b.orddt", strKeys(1), 2) & _
               " AND " & DBW("b.donefg > ", "0")
       tmpStr1 = " AND e.ptid = b.ptid AND e.orddt = b.orddt AND e.ordno = b.ordno AND e.ordseq = b.ordseq " & _
                 " AND a.workarea = e.workarea AND a.accdt = e.accdt AND a.accseq = e.accseq AND ( b.dcfg ='' or b.dcfg is null ) "
    Else   '접수번호 기준
       tmpStr = tmpStr & DBW("a.workarea", strKeys(0), 2) & _
               " AND " & DBW("a.accdt", strKeys(1), 2) & _
               " AND " & DBW("a.accseq", strKeys(2), 2)
       tmpStr1 = " AND e.workarea = a.workarea AND e.accdt = a.accdt AND e.accseq = a.accseq " & _
                 " AND b.ptid = e.ptid AND b.orddt = e.orddt AND b.ordno = e.ordno AND b.ordseq = e.ordseq AND ( b.dcfg ='' or b.dcfg is null ) "
    End If
    SqlBarReprint = " SELECT c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                    "        a.deptcd, a.wardid, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.accseq") & " as LabNo, " & _
                    "        d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                    "        b.orddt, b.ordno, b.ordseq, b.ordcd, e.field1 as BuildNm, a.coltm, " & _
                    "        g.reqdt, g.reqtm,a.coldt as new_coldt,g.orddiv,'' fzfg, c.testdiv " & _
                    " FROM   " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & _
                                 "" & T_LAB101 & " g,  " & T_LAB201 & " a, " & T_LAB102 & " b " & _
                    " WHERE  " & tmpStr & _
                    " AND   (a.multifg ='' or a.multifg is null) " & _
                    " AND    b.workarea = a.workarea " & _
                    " AND    b.accdt = a.accdt " & _
                    " AND    b.accseq = a.accseq " & _
                    " AND   (b.dcfg='' or b.dcfg is null ) " & _
                    " AND    c.testcd = b.ordcd " & _
                    " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                    " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & _
                    " AND    d.cdval1 = a.spccd " & _
                    " AND    " & DBJ(DBW("e.cdindex", LC3_Buildings, 2)) & _
                    " AND    " & DBJ("e.cdval1 =* a.buildcd") & _
                    "  " & _
                    " AND    g.ptid = b.ptid " & _
                    " AND    g.orddt = b.orddt " & _
                    " AND    g.ordno = b.ordno "
    'Union
    SqlBarReprint = SqlBarReprint & _
                    " UNION ALL " & _
                    " SELECT c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                    "        a.deptcd, a.wardid, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.accseq") & " as LabNo, " & _
                    "        d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                    "        b.orddt, b.ordno, b.ordseq, b.ordcd, g.field1 as BuildNm, a.coltm, " & _
                    "        i.reqdt, i.reqtm,a.coldt as new_coldt ,i.orddiv,'' fzfg, c.testdiv " & _
                    " FROM   " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_LAB203 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g, " & _
                                 " " & T_LAB101 & " i, " & T_LAB201 & " a, " & T_LAB102 & " b  " & _
                    " WHERE  " & tmpStr & tmpStr1 & _
                    " AND    a.multifg <> ' ' " & _
                    " AND    " & DBW("f.cdindex", LC2_Panel, 2) & _
                    " AND    f.cdval1 = b.ordcd " & _
                    " AND    f.cdval2 = e.spcseq " & _
                    " AND    c.testcd = f.field1 " & _
                    " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                    " AND    " & DBW("d.cdindex", LC3_Specimen, 2) & _
                    " AND    d.cdval1 = a.spccd " & _
                    " AND    " & DBJ(DBW("g.cdindex", LC3_Buildings, 2)) & _
                    " AND    " & DBJ("g.cdval1 =* a.buildcd") & _
                    " " & _
                    " AND    i.ptid  = b.ptid " & _
                    " AND    i.orddt = b.orddt " & _
                    " AND    i.ordno = b.ordno " & _
                    " Order  By LabNo, ordcd "

End Function

'frm157BarReprint
Public Function SqlGetOrdDateForBarprint(ByVal pPtid As String, ByVal pOrdDiv As String, _
                                         Optional ByVal pMonthFg As Boolean = True) As String

    SqlGetOrdDateForBarprint = " SELECT  a.orddt " & _
                               " FROM  " & T_LAB101 & " b, " & T_LAB102 & " a " & _
                               " WHERE " & DBW("a.ptid   = ", pPtid) & _
                               " AND   " & DBW("a.donefg > ", enStsCd.StsCd_LIS_Order) & _
                               " AND    b.ptid  = a.ptid " & _
                               " AND    b.orddt = a.orddt " & _
                               " AND    b.ordno = a.ordno " & _
                               " AND    " & DBW("a.stscd <=", enStsCd.StsCd_LIS_MidRst) & _
                               " AND    " & DBW("b.orddt <=", mvarOrderDate) & _
                               " AND    (a.dcfg = '' or a.dcfg is null)"
    If pOrdDiv <> "" Then
        SqlGetOrdDateForBarprint = SqlGetOrdDateForBarprint & " AND " & DBW("b.orddiv = ", pOrdDiv)
    End If
    If pMonthFg Then
        SqlGetOrdDateForBarprint = SqlGetOrdDateForBarprint & _
                                    " AND " & DBW("reqdt >= ", Format(DateAdd("m", -1, GetSystemDate), CS_DateDbFormat))
    End If
    SqlGetOrdDateForBarprint = SqlGetOrdDateForBarprint & " group by a.orddt order by a.orddt desc "

End Function

'frm158AccList
Public Function SqlGetPtListForAccess() As String

    SqlGetPtListForAccess = " SELECT a.ptid, b." & F_PTNM & " as ptnm, a.orddt " & _
                            " FROM   " & T_LAB101 & " a, " & T_HIS001 & " b " & _
                            " WHERE  " & DBW("a.donefg  = ", enStsCd.StsCd_LIS_Order) & _
                            " AND    " & DBW("a.reqdt  = ", Format(GetSystemDate, CS_DateDbFormat)) & _
                            " AND    " & DBW("a.bussdiv = ", enBussDiv.BussDiv_OutPatient) & " " & _
                            " AND    b." & F_PTID & " = a.ptid " & _
                            " AND    exists (SELECT * FROM " & T_LAB102 & " c, " & T_LAB001 & " d, " & T_LAB004 & " e " & _
                                             " WHERE c.ptid   = a.ptid   AND   c.orddt = a.orddt " & _
                                             " AND   c.ordno  = a.ordno  AND ( c.dcfg  = '' or c.dcfg is null ) " & _
                                             " AND   " & DBW("c.donefg  = ", enStsCd.StsCd_LIS_Order) & _
                                             " AND   d.testcd = c.ordcd  AND c.ordcd=e.testcd AND c.spccd=e.spccd ) " & _
                            " group by ptid, " & F_PTNM & ", orddt " '  order by orddt desc,ptid"

    SqlGetPtListForAccess = SqlGetPtListForAccess & "union" & _
                            " SELECT a.ptid, b." & F_PTNM & " as ptnm, a.orddt " & _
                            " FROM   " & T_LAB101 & " a, " & T_HIS001 & " b " & _
                            " WHERE  " & DBW("a.donefg  = ", enStsCd.StsCd_LIS_Order) & _
                            " AND    " & DBW("a.reqdt  >= ", Format(GetSystemDate, CS_DateDbFormat)) & _
                            " AND    " & DBW("a.bussdiv = ", enBussDiv.BussDiv_OutPatient) & " " & _
                            " AND    b." & F_PTID & " = a.ptid " & _
                            " AND    exists (SELECT * FROM " & T_LAB102 & " c, " & T_BBS001 & " d " & _
                                             " WHERE c.ptid   = a.ptid   AND   c.orddt = a.orddt " & _
                                             " AND   c.ordno  = a.ordno  AND ( c.dcfg  = '' or c.dcfg is null ) " & _
                                             " AND   " & DBW("c.donefg  = ", enStsCd.StsCd_LIS_Order) & _
                                             " AND   d.testcd = c.ordcd ) " & _
                            " group by ptid, " & F_PTNM & ", orddt  order by orddt desc,ptid"

 End Function
 
  

Public Function SqlLoadSpcItem(ByVal pSpcCd As String)
    SqlLoadSpcItem = " SELECT distinct a.testcd, a.testnm " & _
                     " FROM " & T_LAB001 & " a, " & T_LAB004 & " b " & _
                     " WHERE (a.detailfg = ''  or  a.detailfg is null)" & _
                     " AND   b.testcd   = a.testcd " & _
                     " AND   " & DBW("b.spccd  = ", pSpcCd)
End Function

Public Function GetSqlLab004(Optional ByVal pTestCd As String = "")

    GetSqlLab004 = " SELECT * FROM " & T_LAB004
    
    If pTestCd <> "" Then GetSqlLab004 = GetSqlLab004 & " WHERE " & DBW("testcd", pTestCd, 2)
    
End Function

Public Function SqlGetCumList(ByVal pDeptcd As String) As String
    
'    SqlGetCumList = " SELECT distinct cdval1 as CumCd, field1 as CumNm, field3 as PWD " & _
'                    " FROM  " & T_LAB031 & _
'                    " WHERE " & DBW("cdindex", LC2_CumItem, 2)
'
'    If pDeptCd <> "ALL" Then SqlGetCumList = SqlGetCumList & " AND " & DBW("field2", pDeptCd, 2)
'    SqlGetCumList = SqlGetCumList & " order by CumCd"
    
    SqlGetCumList = " SELECT distinct field2,cdval1 as CumCd, field1 as CumNm, field3 as PWD " & _
                    " FROM  " & T_LAB031 & _
                    " WHERE " & DBW("cdindex", LC2_CumItem, 2)
                    
    If pDeptcd = "" Then
        SqlGetCumList = SqlGetCumList & " AND " & DBW("field2", "0", 2)
        
    
    ElseIf pDeptcd = "dept" Then
        SqlGetCumList = SqlGetCumList & " AND " & DBW("field2<>", "0")
    End If
    
End Function



Public Function SqlBatchBarReprint(ByVal pWorkDt As String, ByVal pWorkTm As String, ByVal pWardId As String) As String
    

     SqlBatchBarReprint = " SELECT x.seq, c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                           "    a.deptcd, a.wardid, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("char", "a.accseq") & " as LabNo, " & _
                           "    d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                           "    b.orddt, b.ordno, b.ordseq, b.ordcd, e.field1 as BuildNm, f." & F_PTNM & " as PtNm, " & _
                           "    g.reqdt, g.reqtm, '" & lis_orddiv & "' as orddiv, '' fzfg " & _
                           " FROM  " & T_LAB204 & " x, " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & _
                                       T_LAB032 & " d, " & T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " g " & _
                           " WHERE " & DBW("x.workdt", pWorkDt, 2) & _
                           " AND   " & DBW("x.wardid", pWardId, 2) & " AND " & DBW("x.worktm", pWorkTm, 2) & " AND " & DBW("x.orddiv", lis_orddiv, 2) & _
                           " AND   a.workarea = x.workarea AND a.accdt = x.accdt AND a.accseq = x.accseq " & _
                           " AND  (a.multifg = ''  or  a.multifg is null)" & _
                           " AND   b.workarea = a.workarea " & _
                           " AND   b.accdt = a.accdt " & _
                           " AND   b.accseq = a.accseq " & _
                           " AND  (b.dcfg='' or b.dcfg is null ) " & _
                           " AND   c.testcd = b.ordcd " & _
                           " AND   c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                           " AND   " & DBW("d.cdindex", LC3_Specimen, 2) & _
                           " AND   d.cdval1 = a.spccd " & _
                           " AND   " & DBJ(DBW("e.cdindex", LC3_Buildings, 2)) & _
                           " AND   " & DBJ("e.cdval1 =* a.buildcd") & _
                           " AND   g.ptid = b.ptid " & _
                           " AND   g.orddt = b.orddt " & _
                           " AND   g.ordno = b.ordno " & _
                           " AND   f." & F_PTID & " = a.ptid "
    'Union
     SqlBatchBarReprint = SqlBatchBarReprint & _
                           " UNION ALL " & _
                           " SELECT x.seq, c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                           "           a.deptcd, a.wardid, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.accseq") & " as LabNo, " & _
                           "           d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                           "           b.orddt, b.ordno, b.ordseq, b.ordcd, g.field1 as BuildNm, h." & F_PTNM & " as PtNm, " & _
                           "           i.reqdt, i.reqtm, '" & lis_orddiv & "' as orddiv, '' fzfg " & _
                           " FROM  " & T_LAB204 & " x, " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                                       T_LAB203 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g, " & T_HIS001 & " h, " & T_LAB101 & " i " & _
                           " WHERE " & DBW("x.workdt", pWorkDt, 2) & " AND " & DBW("x.wardid", pWardId, 2) & " AND " & DBW("x.worktm", pWorkTm, 2) & " AND " & DBW("x.orddiv", lis_orddiv, 2) & _
                           " AND   a.workarea = x.workarea AND a.accdt = x.accdt AND a.accseq = x.accseq " & _
                           " AND   e.workarea = a.workarea AND e.accdt = a.accdt AND e.accseq = a.accseq " & _
                           " AND   b.ptid = e.ptid AND b.orddt = e.orddt AND b.ordno = e.ordno AND b.ordseq = e.ordseq AND ( b.dcfg='' or b.dcfg is null ) " & _
                           " AND   a.multifg <> ' ' " & _
                           " AND   " & DBW("f.cdindex", LC2_Panel, 2) & _
                           " AND   f.cdval1 = b.ordcd  AND  f.cdval2 = e.spcseq " & _
                           " AND   c.testcd = f.field1 " & _
                           " AND   c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                           " AND   " & DBW("d.cdindex", LC3_Specimen, 2) & " AND  d.cdval1 = a.spccd " & _
                           " AND   " & DBJ(DBW("g.cdindex", LC3_Buildings, 2)) & " AND  " & DBJ("g.cdval1 =* a.buildcd") & _
                           " AND   i.ptid = b.ptid " & _
                           " AND   i.orddt = b.orddt " & _
                           " AND   i.ordno = b.ordno " & _
                           " AND   h." & F_PTID & " = a.ptid "
                               

    If P_IncludeBBSSystem Then
    '혈액
        SqlBatchBarReprint = SqlBatchBarReprint & _
                            " UNION ALL " & _
                            " SELECT x.seq, c.testnm, c.abbrnm5, a.ptid, 'XM' as workarea, '' as accdt, 0 as accseq, " & _
                            "        a.spcyy, a.spcno, '' spccd, '' as storecd, '1' as stscd, b.statfg, j.buildcd, " & _
                            "        i.deptcd, i.wardid,i.hosilid, x.spcyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("char", "x.spcno") & " as LabNo, " & _
                            "        '혈액' as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
                            "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                            "        e.field1 as BuildNm, f." & F_PTNM & "  as PtNm, " & _
                            "        i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                            " FROM " & T_LAB204 & " x, " & T_BBS102 & " a, " & T_LAB102 & " b, " & T_BBS001 & " c, " & _
                                       T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i, " & T_BBS201 & " j " & _
                            " WHERE " & DBW("x.workdt", pWorkDt, 2) & " AND " & DBW("x.wardid", pWardId, 2) & " AND " & DBW("x.worktm", pWorkTm, 2) & _
                            " AND    a.spcyy = x.spcyy  AND    a.spcno = x.spcno " & _
                            " AND    j.spcyy = x.spcyy  AND    j.spcno = x.spcno " & _
                            " AND    i.ptid = a.ptid   AND    i.orddt = a.orddt  AND    i.ordno = a.ordno " & _
                            " AND    b.ptid = i.ptid   AND    b.orddt = i.orddt  AND    b.ordno = i.ordno " & _
                            " AND  " & DBW("i.orddiv = ", BBS_ORDDIV) & _
                            " AND    c.testcd = b.ordcd " & _
                            " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
                            " AND  " & DBJ(DBW("e.cdindex = ", LC3_Buildings)) & _
                            " AND  " & DBJ("e.cdval1 =* j.buildcd") & _
                            " AND    f." & F_PTID & " = a.ptid "
    End If
    
    SqlBatchBarReprint = SqlBatchBarReprint & _
                        " Order By seq, LabNo, ordcd "


End Function

Public Function SqlBatchBarReprint_New(ByVal pWorkDt As String, ByVal pWorkTm As String, _
                                    ByVal pWardId As String, Optional ByVal pCorpID As String = "") As String
    

    SqlBatchBarReprint_New = " SELECT x.seq, c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                             "    a.deptcd, a.wardid, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("char", "a.accseq") & " as LabNo, " & _
                             "    d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                             "    b.orddt, b.ordno, b.ordseq, b.ordcd, e.field1 as BuildNm, f." & F_PTNM & " as PtNm, " & _
                             "    g.reqdt, g.reqtm, '" & lis_orddiv & "' as orddiv, '' fzfg " & _
                             " FROM  " & T_LAB204 & " x, " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & _
                                         T_LAB032 & " d, " & T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " g " & _
                             " WHERE " & DBW("x.workdt", pWorkDt, 2) & _
                             " AND   " & DBW("x.wardid", pWardId, 2) & " AND " & DBW("x.worktm", pWorkTm, 2) & " AND " & DBW("x.orddiv", lis_orddiv, 2)
    
    If pCorpID <> "" Then
        SqlBatchBarReprint_New = SqlBatchBarReprint_New & " AND b.inout = " & DBS(pCorpID)
    End If
                           
    SqlBatchBarReprint_New = SqlBatchBarReprint_New & _
                             " AND   a.workarea = x.workarea AND a.accdt = x.accdt AND a.accseq = x.accseq " & _
                             " AND  (a.multifg = ''  or  a.multifg is null)" & _
                             " AND   b.workarea = a.workarea " & _
                             " AND   b.accdt = a.accdt " & _
                             " AND   b.accseq = a.accseq " & _
                             " AND  (b.dcfg='' or b.dcfg is null ) " & _
                             " AND   c.testcd = b.ordcd " & _
                             " AND   c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                             " AND   " & DBW("d.cdindex", LC3_Specimen, 2) & _
                             " AND   d.cdval1 = a.spccd " & _
                             " AND   " & DBJ(DBW("e.cdindex", LC3_Buildings, 2)) & _
                             " AND   " & DBJ("e.cdval1 =* a.buildcd") & _
                             " AND   g.ptid = b.ptid " & _
                             " AND   g.orddt = b.orddt " & _
                             " AND   g.ordno = b.ordno " & _
                             " AND   f." & F_PTID & " = a.ptid "
    'Union
    SqlBatchBarReprint_New = SqlBatchBarReprint_New & _
                             " UNION ALL " & _
                             " SELECT x.seq, c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                             "           a.deptcd, a.wardid, a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.accseq") & " as LabNo, " & _
                             "           d.field5 as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                             "           b.orddt, b.ordno, b.ordseq, b.ordcd, g.field1 as BuildNm, h." & F_PTNM & " as PtNm, " & _
                             "           i.reqdt, i.reqtm, '" & lis_orddiv & "' as orddiv, '' fzfg " & _
                             " FROM  " & T_LAB204 & " x, " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                                         T_LAB203 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g, " & T_HIS001 & " h, " & T_LAB101 & " i " & _
                             " WHERE " & DBW("x.workdt", pWorkDt, 2) & " AND " & DBW("x.wardid", pWardId, 2) & " AND " & DBW("x.worktm", pWorkTm, 2) & " AND " & DBW("x.orddiv", lis_orddiv, 2)
                           
    If pCorpID <> "" Then
        SqlBatchBarReprint_New = SqlBatchBarReprint_New & " AND b.inout = " & DBS(pCorpID)
    End If
    
    SqlBatchBarReprint_New = SqlBatchBarReprint_New & _
                             " AND   a.workarea = x.workarea AND a.accdt = x.accdt AND a.accseq = x.accseq " & _
                             " AND   e.workarea = a.workarea AND e.accdt = a.accdt AND e.accseq = a.accseq " & _
                             " AND   b.ptid = e.ptid AND b.orddt = e.orddt AND b.ordno = e.ordno AND b.ordseq = e.ordseq AND ( b.dcfg='' or b.dcfg is null ) " & _
                             " AND   a.multifg <> ' ' " & _
                             " AND   " & DBW("f.cdindex", LC2_Panel, 2) & _
                             " AND   f.cdval1 = b.ordcd  AND  f.cdval2 = e.spcseq " & _
                             " AND   c.testcd = f.field1 " & _
                             " AND   c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND " & DBW("applydt <= ", Format(Now, CS_DateDbFormat)) & ") " & _
                             " AND   " & DBW("d.cdindex", LC3_Specimen, 2) & " AND  d.cdval1 = a.spccd " & _
                             " AND   " & DBJ(DBW("g.cdindex", LC3_Buildings, 2)) & " AND  " & DBJ("g.cdval1 =* a.buildcd") & _
                             " AND   i.ptid = b.ptid " & _
                             " AND   i.orddt = b.orddt " & _
                             " AND   i.ordno = b.ordno " & _
                             " AND   h." & F_PTID & " = a.ptid "
                               
'    If P_IncludeBBSSystem Then
'    '혈액
'        SqlBatchBarReprint_New = SqlBatchBarReprint_New & _
'                             " UNION ALL " & _
'                             " SELECT x.seq, c.testnm, c.abbrnm5, a.ptid, 'XM' as workarea, '' as accdt, 0 as accseq, " & _
'                             "        a.spcyy, a.spcno, '' spccd, '' as storecd, '1' as stscd, b.statfg, j.buildcd, " & _
'                             "        i.deptcd, i.wardid,i.hosilid, x.spcyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("char", "x.spcno") & " as LabNo, " & _
'                             "        '혈액' as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
'                             "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
'                             "        e.field1 as BuildNm, f." & F_PTNM & "  as PtNm, " & _
'                             "        i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
'                             " FROM " & T_LAB204 & " x, " & T_BBS102 & " a, " & T_LAB102 & " b, " & T_BBS001 & " c, " & _
'                                        T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i, " & T_BBS201 & " j " & _
'                             " WHERE " & DBW("x.workdt", pWorkDt, 2) & " AND " & DBW("x.wardid", pWardId, 2) & " AND " & DBW("x.worktm", pWorkTm, 2) & _
'                             " AND    a.spcyy = x.spcyy  AND    a.spcno = x.spcno " & _
'                             " AND    j.spcyy = x.spcyy  AND    j.spcno = x.spcno " & _
'                             " AND    i.ptid = a.ptid   AND    i.orddt = a.orddt  AND    i.ordno = a.ordno " & _
'                             " AND    b.ptid = i.ptid   AND    b.orddt = i.orddt  AND    b.ordno = i.ordno " & _
'                             " AND  " & DBW("i.orddiv = ", BBS_ORDDIV) & _
'                             " AND    c.testcd = b.ordcd " & _
'                             " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
'                             " AND  " & DBJ(DBW("e.cdindex = ", LC3_Buildings)) & _
'                             " AND  " & DBJ("e.cdval1 =* j.buildcd") & _
'                             " AND    f." & F_PTID & " = a.ptid "
'    End If

    SqlBatchBarReprint_New = SqlBatchBarReprint_New & _
                        " Order By seq, LabNo, ordcd "


End Function

Public Function SqlWardBarReprint(ByVal intOption As Integer, ParamArray strKeys() As Variant) As String

    Dim tmpStr      As String
    Dim tmpStr1     As String
    Dim tmpSql(3)   As String
    Dim blnAllFg    As Boolean
   
    tmpStr = ""
    blnAllFg = False
    If intOption = 1 Then  '환자ID-처방일 기준
        '## 5.0.34: 이상대(2005-07-26)
        '   - 재발행 화면에서 처방조회 속도가 느려 쿼리수정
        tmpStr = tmpStr & DBW("a.ptid = ", strKeys(0)) & _
                " AND " & DBW("b.orddt = ", strKeys(1)) & _
                " AND  b.donefg > '0'  AND b.donefg is not null " & _
                " AND  a.ptid = b.ptid " & _
                " AND (b.dcfg = '' or b.dcfg is null) "
        tmpStr1 = " AND e.ptid = b.ptid AND a.ptid = b.ptid AND e.orddt = b.orddt " & _
                  " AND e.ordno = b.ordno AND e.ordseq = b.ordseq " & _
                  " AND a.workarea = e.workarea AND a.accdt = e.accdt " & _
                  " AND a.accseq = e.accseq " & _
                  " AND a.stscd < '4' "
        If UBound(strKeys) = 2 Then
            If strKeys(2) = "" Then blnAllFg = True
        End If
    Else        '접수번호 기준
        tmpStr = tmpStr & DBW("a.workarea = ", strKeys(0)) & _
                " AND " & DBW("a.accdt  = ", strKeys(1)) & _
                " AND " & DBW("a.accseq = ", strKeys(2))
        tmpStr1 = " AND e.workarea = a.workarea AND e.accdt = a.accdt " & _
                  " AND e.accseq = a.accseq AND b.ptid = e.ptid AND a.ptid = b.ptid " & _
                  " AND b.orddt = e.orddt AND b.ordno = e.ordno " & _
                  " AND b.ordseq = e.ordseq " & _
                  " AND (b.dcfg = '' or b.dcfg is null) "
    End If
    
    '임상병리
    tmpSql(1) = " SELECT c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                "        a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, b.statfg, a.buildcd, " & _
                "        a.deptcd, a.wardid,a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & _
                         FUNC_CONVERT("char", "a.accseq") & " as LabNo, d.field5 as SpcNm, a.coltm, " & _
                         FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, " & _
                "        b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "        e.field1 as BuildNm,  i.reqdt, i.reqtm, i.orddiv, '' fzfg, c.testdiv, a.coldt as new_coldt " & _
                " FROM " & T_LAB201 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB032 & " e, " & T_LAB101 & " i " & _
                " WHERE " & tmpStr '
               
    tmpSql(1) = tmpSql(1) & " AND   (a.multifg = '' or a.multifg is null) "
    
    tmpSql(1) = tmpSql(1) & _
                " AND    b.workarea = a.workarea  AND    b.accdt = a.accdt  AND    b.accseq = a.accseq " & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd " & _
                                    " AND applydt <= a.coldt) " & _
                " AND  " & DBW("d.cdindex = ", LC3_Specimen) & _
                " AND    d.cdval1 = a.spccd " & _
                " AND  " & DBJ(DBW("e.cdindex = ", LC3_Buildings)) & _
                " AND  " & DBJ("e.cdval1 =* a.buildcd") & _
                " " & _
                " AND    i.ptid  = b.ptid " & _
                " AND    i.orddt = b.orddt " & _
                " AND    i.ordno = b.ordno " & _
                " AND    a.stscd <= '4' "
    
    'Union
    tmpSql(1) = tmpSql(1) & " UNION ALL " & _
                " SELECT c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                "        a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                "        a.deptcd, a.wardid,a.hosilid, a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.accdt,3,6)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & _
                         FUNC_CONVERT("char", "a.accseq") & " as LabNo, d.field5 as SpcNm, a.coltm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
                "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "        g.field1 as BuildNm,  i.reqdt, i.reqtm, i.orddiv, '' fzfg, c.testdiv, a.coldt as new_coldt " & _
                " FROM " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB203 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g, " & _
                           " " & T_LAB101 & " i " & _
                " WHERE " & tmpStr & tmpStr1 & _
                " AND    a.multifg <> ' '  AND a.multifg is not null" & _
                " AND    f.cdindex = '" & LC2_Panel & "' " & _
                " AND    f.cdval1 = b.ordcd  AND  f.cdval2 = e.spcseq " & _
                " AND    c.testcd = f.field1 " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd " & _
                                    " AND applydt <= a.coldt) " & _
                " AND  " & DBW("d.cdindex = ", LC3_Specimen) & "  AND  d.cdval1 = a.spccd" & _
                " AND  " & DBJ(DBW("g.cdindex = ", LC3_Buildings)) & " AND " & DBJ("g.cdval1 =* a.buildcd") & _
                "  " & _
                " AND    i.ptid  = b.ptid " & _
                " AND    i.orddt = b.orddt " & _
                " AND    i.ordno = b.ordno " & _
                " AND    a.stscd <='4' "
    
    If P_IncludeBBSSystem Then
    '혈액은행
        tmpSql(3) = " SELECT c.testnm, c.abbrnm5, a.ptid, 'XM' as workarea, '' as accdt, 0 as accseq, " & _
                    "        a.spcyy, a.spcno, '' spccd, '' as storecd, b.stscd, b.statfg, a.buildcd, " & _
                    "        i.deptcd, i.wardid,i.hosilid, '' as LabNo, '혈액' as SpcNm, a.coltm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
                    "        " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                    "        e.field1 as BuildNm,  " & _
                    "        i.reqdt, i.reqtm, i.orddiv, '' fzfg, c.testdiv, a.coldt as new_coldt " & _
                    " FROM " & T_BBS201 & " a, " & T_LAB102 & " b, " & T_BBS001 & " c, " & _
                               T_LAB032 & " e, " & T_LAB101 & " i " & _
                    " WHERE " & tmpStr & _
                    " AND    i.ptid = b.ptid " & _
                    " AND    i.orddt = b.orddt " & _
                    " AND    i.ordno = b.ordno " & _
                    " AND  " & DBW("i.orddiv = ", BBS_ORDDIV) & _
                    " AND    a.ptid = i.ptid " & _
                    " AND    a.spcyy = (SELECT max(spcyy) FROM " & T_BBS201 & " WHERE ptid = i.ptid) " & _
                    " AND    a.spcno = (SELECT max(spcno) FROM " & T_BBS201 & _
                                      " WHERE ptid = a.ptid AND spcyy = a.spcyy) " & _
                    " AND    c.testcd = b.ordcd " & _
                    " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
                    " AND  " & DBJ(DBW("e.cdindex = ", LC3_Buildings)) & _
                    " AND  " & DBJ("e.cdval1 =* a.buildcd") & _
                    " " & _
                    " AND  b.stscd < ='4'"
    End If
    
    If blnAllFg Then
        If P_IncludeBBSSystem Then
            SqlWardBarReprint = tmpSql(1) & " UNION ALL " & tmpSql(3)
        Else
            SqlWardBarReprint = tmpSql(1)
        End If
    Else
        If UBound(strKeys) = 2 Then
            Select Case strKeys(2)
                Case "B": SqlWardBarReprint = tmpSql(3)
                Case "L": SqlWardBarReprint = tmpSql(1)
            End Select
        Else
        
            SqlWardBarReprint = tmpSql(1)
        End If
    End If
    SqlWardBarReprint = SqlWardBarReprint & " Order By LabNo, ordno, ordcd "

End Function
   

Public Function SqlGsOrder(ByVal ReqDt As String, ByVal ReqTm As String, ByVal DeptCd As String) As String
   
   SqlGsOrder = " SELECT b.deptcd as DeptCd, a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                         F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                         F_DOB2("a") & " as Dob, b.bedindt, b.orddoct, b.majdoct, " & _
                "        b.roomid, '' as bedid, b.hosilid " & _
                " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                " WHERE  " & DBW("b.donefg", "0", 2) & _
                " AND    " & DBW("b.reqdt", ReqDt, 2) & _
                " AND    b.deptcd in (" & DeptCd & ") " & _
                " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_OutPatient, 2) & _
                " AND    a." & F_PTID & " = b.ptid " & _
                " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                " AND    " & DBW("c.stscd", "0", 2) & " AND ( c.dcfg='' or b.dcfg is null ) " & _
                " AND    " & DBW("c.ordcd", "B400", 2) & " AND " & DBW("c.spccd", "75", 2) & _
                " Order By DeptCd, PtId, c.statfg, c.spccd "
End Function


Public Function SqlInstrument(Optional ByVal pEqpCd As String = "")

    SqlInstrument = " SELECT * " & _
                    " FROM  " & T_LAB006 & " a " & _
                    " WHERE " & DBW("eqpdiv", "E", 2)
                    
    If pEqpCd <> "" Then
        SqlInstrument = SqlInstrument & " AND " & DBW("a.eqpcd", pEqpCd, 2)
    End If
    SqlInstrument = SqlInstrument & " ORDER BY eqpnm"

End Function

Public Function SqlInstrument_New(Optional ByVal pEqpCd As String = "")

    SqlInstrument_New = " SELECT * " & _
                    " FROM  " & T_LAB006 & " a "
                    
    If pEqpCd <> "" Then
        SqlInstrument_New = SqlInstrument_New & " WHERE " & DBW("a.eqpcd", pEqpCd, 2)
    End If
    SqlInstrument_New = SqlInstrument_New & " ORDER BY eqpcd, eqpnm"

End Function

Public Function SqlDeleteInstrument(ByVal pEqpCd As String)
    SqlDeleteInstrument = " delete " & T_LAB006 & _
                          " WHERE  " & DBW("eqpcd = ", pEqpCd)
End Function

Public Function SqlInsertInstrument(ByVal pEqpCd As String, ParamArray pData())
    
    SqlInsertInstrument = " Insert into " & T_LAB006 & _
                         "                 (eqpcd, eqpnm, eqpdiv, modelnm, purchdt, " & _
                         "                  vandcd, temphigh, templow, remark,location,inusefg )" & _
                         " values( " & _
                                    DBV("eqpcd    ", pEqpCd) & " ," & _
                                    DBV("eqpnm    ", pData(0)) & " ," & _
                                    DBV("eqpdiv   ", pData(1)) & " , " & _
                                    DBV("modelnm  ", pData(2)) & " , " & _
                                    DBV("purchdt  ", pData(3)) & " ," & _
                                    DBV("vandcd   ", pData(4)) & " ," & _
                                    DBV("temphigh ", pData(5)) & " ," & _
                                    DBV("templow  ", pData(6)) & " ," & _
                                    DBV("remark   ", pData(7)) & " ," & _
                                    DBV("location ", pData(8)) & " ," & _
                                    DBV("inusefg  ", pData(9)) & _
                        ")"
End Function

Public Function SqlInsertInstrument_New(ByVal pEqpCd As String, ParamArray pData())
    
    SqlInsertInstrument_New = " Insert into " & T_LAB006 & _
                              "                 (eqpcd, eqpnm, eqpdiv, modelnm, purchdt, " & _
                              "                  vandcd, temphigh, templow, remark,location,inusefg, tempscale)" & _
                              " values( " & _
                                DBV("eqpcd    ", pEqpCd) & " ," & _
                                DBV("eqpnm    ", pData(0)) & " ," & _
                                DBV("eqpdiv   ", pData(1)) & " , " & _
                                DBV("modelnm  ", pData(2)) & " , " & _
                                DBV("purchdt  ", pData(3)) & " ," & _
                                DBV("vandcd   ", pData(4)) & " ," & _
                                DBV("temphigh ", pData(5)) & " ," & _
                                DBV("templow  ", pData(6)) & " ," & _
                                DBV("remark   ", pData(7)) & " ," & _
                                DBV("location ", pData(8)) & " ," & _
                                DBV("inusefg  ", pData(9)) & " ," & _
                                DBV("tempscale", pData(10)) & _
                                ")"
End Function

'%  14. 직원정보 검색
'%      - Calling FROM [frm101Order]
Public Function SqlLAB015Read(ByVal paraEmpId As String, ByVal Index As Integer) As String
    If Index = 1 Then
        SqlLAB015Read = " SELECT * " & _
                        " FROM   " & T_LAB015 & " a " & _
                        " WHERE " & DBW("a.logonid", paraEmpId, 2)
    Else
        SqlLAB015Read = " SELECT * " & _
                        " FROM  " & T_LAB015 & " a " & _
                        " WHERE " & DBW("a.empid", paraEmpId, 2)
    End If
End Function


Public Function SqlGetColTimes(ByVal strDate As String, ByVal strWardId As String) As String

    SqlGetColTimes = " SELECT worktm FROM " & T_LAB204 & _
                     " WHERE " & DBW("workdt = ", strDate) & _
                     " AND   " & DBW("wardid = ", strWardId) & _
                     " group by worktm order by worktm"

End Function

Public Function GetIpwonInfo(ByVal pPtid As String) As String
'입원환자의 병동병실을 구해온다.
    Dim sSQL As String
    Dim rs   As Recordset
    
    sSQL = " SELECT " & F_PTWARDID & " as wardid," & F_PTROOMID & " as hosilid FROM " & T_HIS002 & _
           " WHERE " & DBW(F_INPTID & "=", pPtid) & _
           " AND (out_date is null)"
    
    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        GetIpwonInfo = rs.Fields("wardid").Value & "-" & rs.Fields("hosilid").Value
    End If
    Set rs = Nothing
    
End Function

' --------------------------------------------
'|이하는 접수건수통계를 위해서 추가로 작성된  |
'|부분입니다.                                 |
'|where:성바오로병원                          |
'|when :2002/05/27                            |
'|who  :김정규                                |
' --------------------------------------------
Public Function GetAccCntSQL(ByVal pStart As String, ByVal pEnd As String, _
                             Optional ByVal DeptCd As String = "", _
                             Optional ByVal TestCd As String = "", _
                             Optional ByVal EqpCd As String = "", _
                             Optional ByVal WorkArea As String = "", _
                             Optional ByVal OrdDoct As String = "", _
                             Optional ByVal MajDoct As String = "") As String
'FROM LAB201, LAB302, LAB001, HIS003

    Dim SqlStmt   As String
    Dim strTable  As String
    Dim strTable1 As String
    Dim strTable2 As String
    Dim strQuery  As String
    Dim strTmpTable As String
    

    strTmpTable = T_HIS003
    
    
    If mvarInOut <> "" Then
        strTable = T_LAB101 & " e," & strTmpTable & " c, " & T_LAB001 & " d, " & T_LAB302 & " b, " & T_LAB201 & " a "
        strTable1 = T_LAB101 & " e," & strTmpTable & " c, " & T_LAB001 & " d, " & T_LAB351 & " b, " & T_LAB201 & " a "
        strTable2 = T_LAB101 & " e," & strTmpTable & " c, " & T_LAB001 & " d, " & T_LAB404 & " b, " & T_LAB201 & " a "
        strQuery = "AND b.ptid=e.ptid AND b.orddt=e.orddt AND b.ordno=e.ordno AND " & DBW("e.bussdiv=", mvarInOut)
    Else
        strTable = strTmpTable & " c, " & T_LAB001 & " d, " & T_LAB302 & " b, " & T_LAB201 & " a "
        strTable1 = strTmpTable & " c, " & T_LAB001 & " d, " & T_LAB351 & " b, " & T_LAB201 & " a "
        strTable2 = strTmpTable & " c, " & T_LAB001 & " d, " & T_LAB404 & " b, " & T_LAB201 & " a "
        strQuery = ""
    End If
    
            
    '일반검사
    SqlStmt = " SELECT a.rcvdt,a.rcvtm,a.orddoct,a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt " & _
              " FROM  " & strTable & _
              " WHERE " & DBW("a.rcvdt>=", Format(pStart, "########")) & " " & _
              " AND   " & DBW("a.rcvdt<=", Format(pEnd, "########")) & " " & _
              " AND b.workarea = a.workarea " & _
              " AND b.accdt    = a.accdt " & _
              " AND b.accseq   = a.accseq "
              
    SqlStmt = SqlStmt & strQuery
              
              
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' " & _
                        " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & _
                        " AND d.testcd = b.testcd " & _
                        " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd) " '& _
                        " AND (d.detailfg = ''  or  d.detailfg is null)"
    

    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    
    If TestCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("b.testcd=", TestCd)
    End If
    If EqpCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("b.eqpcd=", EqpCd)
    End If
    If DeptCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.deptcd=", DeptCd)
    End If
    If WorkArea <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.workarea=", WorkArea)
    End If
    If OrdDoct <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.orddoct=", OrdDoct)
    End If

    SqlStmt = SqlStmt & " GROUP BY a.rcvdt,a.rcvtm,a.orddoct,a.buildcd, a.workarea, b.eqpcd, c." & F_DEPTNM & ", b.testcd, d.testnm "
    
    
    
    '기타검사
    SqlStmt = SqlStmt & "UNION ALL SELECT a.rcvdt,a.rcvtm,a.orddoct,a.buildcd, a.workarea, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & strTable1
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, "########")) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, "########")) & " "
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    SqlStmt = SqlStmt & strQuery
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd) "
    
'    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    

    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    
    

    If TestCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("b.testcd=", TestCd)
    End If
    If DeptCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.deptcd=", DeptCd)
    End If
    If WorkArea <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.workarea=", WorkArea)
    End If
    If OrdDoct <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.orddoct=", OrdDoct)
    End If

    SqlStmt = SqlStmt & " GROUP BY a.rcvdt,a.rcvtm,a.orddoct,a.buildcd, a.workarea, c." & F_DEPTNM & ", b.testcd, d.testnm "
    
    '미생물검사
    SqlStmt = SqlStmt & "UNION ALL SELECT a.rcvdt,a.rcvtm,a.orddoct,a.buildcd, a.workarea, '' eqpcd, c." & F_DEPTNM & " as DeptNm, b.testcd, d.testnm, count(*) as Cnt "
    SqlStmt = SqlStmt & "FROM  " & strTable2
    SqlStmt = SqlStmt & "WHERE " & DBW("a.rcvdt>=", Format(pStart, "########")) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.rcvdt<=", Format(pEnd, "########")) & " "
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    SqlStmt = SqlStmt & strQuery
    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND " & DBJ("c." & F_DEPTCD & " =* a.deptcd")
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = d.testcd) "
'    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    

    SqlStmt = SqlStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    
    If TestCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("b.testcd=", TestCd)
    End If
    If DeptCd <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.deptcd=", DeptCd)
    End If
    If WorkArea <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.workarea=", WorkArea)
    End If
    If OrdDoct <> "" Then
        SqlStmt = SqlStmt & " AND " & DBW("a.orddoct=", OrdDoct)
    End If

    SqlStmt = SqlStmt & " GROUP BY a.rcvdt,a.rcvtm,a.orddoct,a.buildcd, a.workarea, c." & F_DEPTNM & ", b.testcd, d.testnm "
    
    SqlStmt = SqlStmt & " ORDER BY buildcd, workarea, eqpcd, DeptNm, testcd, testnm "
    
    GetAccCntSQL = SqlStmt
    
End Function

Public Function WorkTime() As String
    WorkTime = "SELECT cdval1,field1,field2 FROM " & T_LAB031 & " WHERE " & DBW("cdindex=", "C123")
End Function

Public Function ConditionNm(ByVal strMode As String, ByVal NameCd As String) As String
    Dim rs     As Recordset
    Dim sSQL   As String
    
    Select Case strMode
        Case "C2"
'                    sSQL = " SELECT " & F_DEPTNM & " AS Nm FROM " & T_HIS004 & " WHERE " & DBW(F_DEPTCD, NameCd, 2)
'                    Set Rs = OpenRecordSet(sSQL)
'                    If Not Rs.EOF Then
'                        ConditionNm = Rs.Fields("Nm").Value & ""
'                    End If
        Case "C3"
            sSQL = " SELECT  field1 as Nm " & _
              " FROM   " & T_LAB032 & _
              " WHERE  " & DBW("cdindex =", lc3_workarea) & _
              " AND    " & DBW("cdval1=", NameCd)
        Case "C4"
            sSQL = " SELECT eqpnm as nm FROM " & T_LAB006 & " WHERE " & DBW("eqpcd=", NameCd)
        Case "C5"
            
           sSQL = " SELECT  a." & F_DOCTNM & " as Nm " & _
                  " FROM  " & T_HIS005 & " a " & _
                  " WHERE " & DBW("a." & F_DOCTID & "", NameCd, 2)
        Case Else
            Exit Function
    End Select
    
    If strMode <> "C2" Then
        Set rs = New Recordset
        rs.Open sSQL, DBConn
        
        If Not rs.EOF Then
            ConditionNm = rs.Fields("Nm").Value & ""
        End If
    End If
    Set rs = Nothing

End Function
Public Function TestCharge(ByVal cdindex As String, ByVal cdVal As String, ByVal Field1 As String, _
                           ByVal Text1 As String, ByVal Text2 As String, Optional ByVal TF As Boolean = False) As String
                           
    If TF = False Then
        TestCharge = " delete " & T_LAB032 & " WHERE " & _
                      DBW("cdindex=", cdindex) & " AND " & DBW("cdval1=", cdVal)
    Else
        TestCharge = " insert into " & T_LAB032 & " (cdindex,cdval1,field1,text1,text2) values(" & _
                     DBV("cdindex", cdindex, 1) & DBV("cdval1", cdVal, 1) & DBV("field1", Field1, 1) & _
                     DBV("text1", Text1, 1) & DBV("text2", Text2) & ")"
                   
    End If
End Function

Public Function TestChangeRs(Optional ByVal TF As Boolean = False) As String
    If TF = False Then
        TestChangeRs = "SELECT * FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", "C249")
    Else
        TestChangeRs = "SELECT * FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", "C250")
    End If
End Function

'====================================================
'bone marrow report sql문
'====================================================

Public Function GetBMTestCd() As String
    
    GetBMTestCd = " SELECT * FROM  " & T_LAB032 & _
                  "  WHERE " & DBW("cdindex", LC3_DiffCount, 2)
    
End Function


Public Function GetPBFindings(Optional ByVal pCdVal1 As String = "") As String
    
    GetPBFindings = " SELECT * FROM  " & T_LAB034 & _
                    "  WHERE " & DBW("cdindex", LC4_Peripheral, 2)
                        
    If pCdVal1 <> "" Then GetPBFindings = GetPBFindings & " AND " & DBW("cdval1", pCdVal1, 2)
  
End Function


Public Function GetBMComment(Optional ByVal pCdVal1 As String = "") As String
    
    GetBMComment = " SELECT * FROM  " & T_LAB034 & _
                   "  WHERE " & DBW("cdindex", LC4_BMComment, 2)
                        
    If pCdVal1 <> "" Then GetBMComment = GetBMComment & " AND " & DBW("cdval1", pCdVal1, 2)
  
End Function

Public Function GetHDiagnosis(Optional ByVal pCdVal1 As String = "") As String
    
    GetHDiagnosis = " SELECT * FROM  " & T_LAB034 & _
                    "  WHERE " & DBW("cdindex", LC4_Hematologic, 2)
                        
    If pCdVal1 <> "" Then GetHDiagnosis = GetHDiagnosis & " AND " & DBW("cdval1", pCdVal1, 2)
  
End Function

Public Function GetCumulative(ByVal pTestCd As String, ByVal pSpcCd As String) As String
    
    GetCumulative = " SELECT b.testcd,b.panelfg,b.testdiv,c.cdval1 as spccd,b.workarea,b.abbrnm10 as TestNm, c.field3 as SpcNm " & _
                    " FROM   " & T_LAB001 & " b, " & T_LAB032 & " c " & _
                    " WHERE " & DBW("b.testcd", pTestCd, 2) & _
                    " AND   " & DBW("c.cdval1", pSpcCd, 2) & _
                    " AND   b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & _
                                        " WHERE testcd = b.testcd) " & _
                    " AND  (" & DBW("b.panelfg != ", PN_Group) & " or b.panelfg is null) " & _
                    " AND   " & DBW("c.cdindex", LC3_Specimen, 2)
  
End Function


Public Function SqlGetPtForAccess(ByVal FrDT As String, ByVal toDt As String) As String
    Dim strtmp  As String
    Dim SQL     As String
    
'    If mvarPayDt <> "" Then
'        strTmp = " And (c.paydt<>'' or c.paydt is not null)"
'    End If
       
    '원본 ========================================================================================
'    SqlGetPtForAccess = " SELECT a.ptid, a.orddt " & _
'                            " FROM   " & T_LAB101 & " a " & _
'                            " WHERE  " & DBW("a.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                            " AND    " & DBW("a.reqdt  >= ", FrDT) & _
'                            " AND    " & DBW("a.reqdt  <= ", toDt) & _
'                            " AND    " & DBW("a.bussdiv = ", enBussDiv.BussDiv_OutPatient) & " " & _
'                            " AND    exists (SELECT * FROM " & T_LAB102 & " c, " & T_LAB001 & " d, " & T_LAB004 & " e " & _
'                                             " WHERE c.ptid   = a.ptid   AND   c.orddt = a.orddt " & _
'                                             " AND   c.ordno  = a.ordno  AND ( c.dcfg  = '' or c.dcfg is null ) " & _
'                                             " AND   " & DBW("c.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                                             " AND   d.testcd = c.ordcd  AND c.ordcd=e.testcd AND c.spccd=e.spccd" & strtmp & " ) " & _
'                            " group by ptid, orddt " '  order by orddt desc,ptid"
'
'    SqlGetPtForAccess = SqlGetPtForAccess & "union" & _
'                            " SELECT a.ptid, a.orddt " & _
'                            " FROM   " & T_LAB101 & " a " & _
'                            " WHERE  " & DBW("a.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                            " AND    " & DBW("a.reqdt  >= ", FrDT) & _
'                            " AND    " & DBW("a.reqdt  <= ", toDt) & _
'                            " AND    " & DBW("a.bussdiv = ", enBussDiv.BussDiv_OutPatient) & " " & _
'                            " AND    exists (SELECT * FROM " & T_LAB102 & " c, " & T_BBS001 & " d " & _
'                                             " WHERE c.ptid   = a.ptid   AND   c.orddt = a.orddt " & _
'                                             " AND   c.ordno  = a.ordno  AND ( c.dcfg  = '' or c.dcfg is null ) " & _
'                                             " AND   " & DBW("c.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                                             " AND   d.testcd = c.ordcd " & strtmp & "  ) " & _
'                            " group by ptid,  orddt  order by orddt desc,ptid"
    '=============================================================================================
            
    '-- 전주 예수병원 ============================================================================
'    SqlGetPtForAccess = " SELECT a.ptid, a.orddt " & _
'                            " FROM   " & T_LAB101 & " a " & _
'                            " WHERE  " & DBW("a.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                            " AND    " & DBW("a.reqdt  >= ", FrDT) & _
'                            " AND    " & DBW("a.reqdt  <= ", toDt) & _
'                            " AND    " & DBW("a.bussdiv = ", enBussDiv.BussDiv_OutPatient) & " " & _
'                            " AND    exists (SELECT * FROM " & T_LAB102 & " c, " & T_LAB001 & " d, " & T_LAB004 & " e " & _
'                                             " WHERE c.ptid   = a.ptid   AND   c.orddt = a.orddt " & _
'                                             " AND   c.ordno  = a.ordno  AND ( c.dcfg  = '' or c.dcfg is null ) " & _
'                                             " AND   " & DBW("c.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                                             " AND   d.testcd = c.ordcd  AND c.ordcd=e.testcd AND c.spccd=e.spccd" & strtmp & " ) "
'
'    SqlGetPtForAccess = SqlGetPtForAccess & "union all" & _
'                            " SELECT a.ptid, a.orddt " & _
'                            " FROM   " & T_LAB101 & " a " & _
'                            " WHERE  " & DBW("a.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                            " AND    " & DBW("a.reqdt  >= ", FrDT) & _
'                            " AND    " & DBW("a.reqdt  <= ", toDt) & _
'                            " AND    " & DBW("a.bussdiv = ", enBussDiv.BussDiv_OutPatient) & " " & _
'                            " AND    exists (SELECT * FROM " & T_LAB102 & " c, " & T_BBS001 & " d " & _
'                                             " WHERE c.ptid   = a.ptid   AND   c.orddt = a.orddt " & _
'                                             " AND   c.ordno  = a.ordno  AND ( c.dcfg  = '' or c.dcfg is null ) " & _
'                                             " AND   " & DBW("c.donefg  = ", enStsCd.StsCd_LIS_Order) & _
'                                             " AND   d.testcd = c.ordcd " & strtmp & "  ) " & _
'                            " order by orddt desc,ptid"
    '=============================================================================================

    '## 5.0.25: 이상대(2005-05-31)
    '   - 뷰를 통한 조회가 느려서 직접테이블을 통해 조회하도록 수정
    '   - 일건, 종건 처방은 조회하지 않음
    If mvarPayDt <> "" Then
        strtmp = " AND a.rcpdate IS NOT NULL"
    End If
    
    SQL = " SELECT /*+ RULE S2LAB004 (S2LAB004_PK) */ a.patno as ptid, to_char(a.orddate, 'yyyymmdd') as orddt" & _
          " FROM mdexmort a, " & T_LAB001 & " b, " & T_LAB004 & " c" & _
          " WHERE " & DBW("a.donefg=", enStsCd.StsCd_LIS_Order) & _
          "     AND a.patsect='O'" & _
          "     AND a.hopedate BETWEEN '" & FrDT & "' AND '" & toDt & "'" & _
          "     AND discyn IS NULL" & strtmp & _
          "     AND b.testcd=a.ordcd" & _
          "     AND c.testcd=a.ordcd AND c.spccd=a.spccode1"
          
    SQL = SQL & " UNION ALL" & _
          " SELECT /*+ RULE S2BBS001 (S2BBS001_PK) */ a.patno as ptid, to_char(a.orddate, 'yyyymmdd') as orddt" & _
          " FROM mdbldort a, " & T_BBS001 & " b" & _
          " WHERE " & DBW("a.donefg=", enStsCd.StsCd_LIS_Order) & _
          "     AND a.patsect='O'" & _
          "     AND a.orddate BETWEEN '" & FrDT & "' AND '" & toDt & "'" & _
          "     AND discyn IS NULL" & strtmp & _
          "     AND b.testcd=a.ordcd" & _
          " ORDER BY orddt desc, ptid"
          
    SqlGetPtForAccess = SQL
End Function


'------------------------
'이미지 시스템
'------------------------
Public Function SqlGetTestCdGraphList() As String
    SqlGetTestCdGraphList = " SELECT distinct testcd, testnm, rsttype " & _
                            "   FROM " & T_LAB001 & _
                            "  WHERE " & DBW("grpfg", "1", 2) & _
                            "    AND (expdt is null or expdt <> ' ') "

End Function

Public Function SqlGetTestNmImageList(Optional ByVal pTestCd As String = "") As String
    SqlGetTestNmImageList = " SELECT distinct testcd, testnm " & _
                            "   FROM " & T_LAB001 & _
                            "  WHERE " & DBW("grpfg", "1", 2)
    
    If pTestCd <> "" Then SqlGetTestNmImageList = SqlGetTestNmImageList & " AND " & DBW("testcd", pTestCd, 2)
End Function


Public Function SqlGetGraphAccList(ByVal pTestCd As String) As String
    SqlGetGraphAccList = " SELECT a.workarea, a.accdt, a.accseq, a.ptid " & _
                         "   FROM " & T_LAB302 & " a, " & T_LAB001 & " b " & _
                         "  WHERE " & DBW("a.testcd", pTestCd, 2) & _
                         "    AND b.workarea = a.workarea AND a.testcd = b.testcd " & _
                         "    AND (a.vfydt is  null or a.vfydt = ' ')"
    
    SqlGetGraphAccList = SqlGetGraphAccList & "UNION ALL "
    
    SqlGetGraphAccList = SqlGetGraphAccList & " SELECT a.workarea, a.accdt, a.accseq, a.ptid " & _
                         "   FROM " & T_LAB351 & " a, " & T_LAB001 & " b " & _
                         "  WHERE " & DBW("a.testcd", pTestCd, 2) & _
                         "    AND b.workarea = a.workarea AND a.testcd = b.testcd " & _
                         "    AND (a.vfydt is  null or a.vfydt = ' ')"
                         
    SqlGetGraphAccList = SqlGetGraphAccList & "UNION ALL "
    
    SqlGetGraphAccList = SqlGetGraphAccList & " SELECT a.workarea, a.accdt, a.accseq, a.ptid " & _
                         "   FROM " & T_LAB404 & " a, " & T_LAB001 & " b " & _
                         "  WHERE " & DBW("a.testcd", pTestCd, 2) & _
                         "    AND b.workarea = a.workarea AND a.testcd = b.testcd " & _
                         "    AND (a.vfydt is  null or a.vfydt = ' ')"
End Function

Public Function SqlGetGraphResultData(ByVal sWorkArea As String, ByVal sAccDt As String, _
                                 ByVal sAccSeq As String, Optional pDisplay As Boolean = True) As String
    
    SqlGetGraphResultData = " SELECT a.testcd, a.testnm, a.rsttype " & _
                       " FROM " & T_LAB001 & " a, " & T_LAB302 & " b " & _
                       " WHERE a.testcd = b.testcd " & _
                       " AND " & DBW("b.workarea", sWorkArea, 2) & _
                       " AND " & DBW("b.accdt", sAccDt, 2) & _
                       " AND " & DBW("b.accseq", sAccSeq, 2) & _
                       " AND   a.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                       "                WHERE testcd = a.testcd AND grpfg ='1' " & _
                       "                AND applydt <= '" & Format(GetSystemDate, CS_DateDbFormat) & "')"
    
    
    If pDisplay = True Then SqlGetGraphResultData = SqlGetGraphResultData & _
                       " AND (b.vfydt is  null or b.vfydt = ' ') "
    
    SqlGetGraphResultData = SqlGetGraphResultData & " UNION ALL " & _
                        " SELECT a.testcd, a.testnm, a.rsttype " & _
                       " FROM " & T_LAB001 & " a, " & T_LAB351 & " b " & _
                       " WHERE a.testcd = b.testcd " & _
                       " AND " & DBW("b.workarea", sWorkArea, 2) & _
                       " AND " & DBW("b.accdt", sAccDt, 2) & _
                       " AND " & DBW("b.accseq", sAccSeq, 2) & _
                       " AND   a.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                       "                WHERE testcd = a.testcd AND grpfg ='1' " & _
                       "                AND applydt <= '" & Format(GetSystemDate, CS_DateDbFormat) & "')"
                       
    If pDisplay = True Then SqlGetGraphResultData = SqlGetGraphResultData & _
                       " AND (b.vfydt is  null or b.vfydt = ' ') "
    
    SqlGetGraphResultData = SqlGetGraphResultData & " UNION ALL " & _
                        " SELECT a.testcd, a.testnm, a.rsttype " & _
                       " FROM " & T_LAB001 & " a, " & T_LAB404 & " b " & _
                       " WHERE a.testcd = b.testcd " & _
                       " AND " & DBW("b.workarea", sWorkArea, 2) & _
                       " AND " & DBW("b.accdt", sAccDt, 2) & _
                       " AND " & DBW("b.accseq", sAccSeq, 2) & _
                       " AND   a.applydt=(SELECT max(applydt) FROM " & T_LAB001 & " " & _
                       "                WHERE testcd = a.testcd AND grpfg ='1' " & _
                       "                AND applydt <= '" & Format(GetSystemDate, CS_DateDbFormat) & "')"
                       
     If pDisplay = True Then SqlGetGraphResultData = SqlGetGraphResultData & _
                       " AND (b.vfydt is  null or b.vfydt = ' ') "
End Function

Public Function SqlGetGraphDataByLabNo(ByVal sTestCd As String, ByVal sWorkArea As String, _
                                  ByVal sAccDt As String, ByVal sAccSeq As String, Optional pDisplay As Boolean = True) As String
    
    SqlGetGraphDataByLabNo = " SELECT a.ptid,f." & F_PTNM & " as ptnm,a.sex,a.ageday,a.deptcd,a.wardid,a.roomid,a.bedid,a.spccd," & _
                        "        a.stscd,a.rmkcd,a.footnotefg,a.rcvdt,c.field4 as spcnm, a.testdiv " & _
                        " FROM   " & T_LAB201 & " a, " & T_HIS001 & " f," & T_LAB032 & " c, " & T_LAB302 & " b " & _
                        " WHERE  " & DBW("b.workarea", sWorkArea, 2) & _
                        " AND    " & DBW("b.accdt", sAccDt, 2) & _
                        " AND    " & DBW("b.accseq", sAccSeq, 2) & _
                        " AND    " & DBW("b.testcd", sTestCd, 2) & _
                        " AND    b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                        " AND    " & DBW("c.cdindex", LC3_Specimen, 2) & _
                        " AND    c.cdval1 = a.spccd" & _
                        " AND    a.ptid   = f." & F_PTID
    
    If pDisplay = True Then SqlGetGraphDataByLabNo = SqlGetGraphDataByLabNo & _
                        " AND    (b.vfydt is  null or b.vfydt =' ') "
    
    SqlGetGraphDataByLabNo = SqlGetGraphDataByLabNo & " UNION ALL " & _
                        " SELECT a.ptid,f." & F_PTNM & " as ptnm,a.sex,a.ageday,a.deptcd,a.wardid,a.roomid,a.bedid,a.spccd," & _
                        "        a.stscd,a.rmkcd,a.footnotefg,a.rcvdt,c.field4 as spcnm, a.testdiv " & _
                        " FROM   " & T_LAB201 & " a, " & T_HIS001 & " f," & T_LAB032 & " c, " & T_LAB351 & " b " & _
                        " WHERE  " & DBW("b.workarea", sWorkArea, 2) & _
                        " AND    " & DBW("b.accdt", sAccDt, 2) & _
                        " AND    " & DBW("b.accseq", sAccSeq, 2) & _
                        " AND    " & DBW("b.testcd", sTestCd, 2) & _
                        " AND    b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                        " AND    " & DBW("c.cdindex", LC3_Specimen, 2) & _
                        " AND    c.cdval1 = a.spccd" & _
                        " AND    a.ptid   = f." & F_PTID
    
    If pDisplay = True Then SqlGetGraphDataByLabNo = SqlGetGraphDataByLabNo & _
                        " AND    (b.vfydt is  null or b.vfydt =' ') "
    
    SqlGetGraphDataByLabNo = SqlGetGraphDataByLabNo & " UNION ALL " & _
                        " SELECT a.ptid,f." & F_PTNM & " as ptnm,a.sex,a.ageday,a.deptcd,a.wardid,a.roomid,a.bedid,a.spccd," & _
                        "        a.stscd,a.rmkcd,a.footnotefg,a.rcvdt,c.field4 as spcnm, a.testdiv " & _
                        " FROM   " & T_LAB201 & " a, " & T_HIS001 & " f," & T_LAB032 & " c, " & T_LAB404 & " b " & _
                        " WHERE  " & DBW("b.workarea", sWorkArea, 2) & _
                        " AND    " & DBW("b.accdt", sAccDt, 2) & _
                        " AND    " & DBW("b.accseq", sAccSeq, 2) & _
                        " AND    " & DBW("b.testcd", sTestCd, 2) & _
                        " AND    b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                        " AND    " & DBW("c.cdindex", LC3_Specimen, 2) & _
                        " AND    c.cdval1 = a.spccd" & _
                        " AND    a.ptid   = f." & F_PTID
    
    If pDisplay = True Then SqlGetGraphDataByLabNo = SqlGetGraphDataByLabNo & _
                        " AND    (b.vfydt is  null or b.vfydt =' ') "

End Function

Public Function InsertGraphImages(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pSeq As Long, _
                                  ByVal pRcvDt As String, ByVal pPtid As String, _
                                  ByVal pImgDiv As String, ByVal pImgDir As String, ByVal pImgsize As String, _
                                  ByVal pPrtFg As String, ByVal pRmk As String, ByVal pIcd As String, Optional pDesc As String = "") As String
    
    'lab310 slide 이미지 내역 등록
    InsertGraphImages = " insert into " & T_LAB310 & "(workarea,accdt,accseq,testcd,seq,rcvdt,ptid,imgdiv,imgdir,imgsize,prtfg,rmk,icd) values( " & _
                        DBV("workarea", pWorkArea, 1) & DBV("accdt", pAccDt, 1) & DBV("accseq", pAccSeq, 1) & DBV("testcd", pTestCd, 1) & DBV("seq", pSeq, 1) & _
                        DBV("rcvdt", pRcvDt, 1) & DBV("ptid", pPtid, 1) & DBV("imgdiv", pImgDiv, 1) & DBV("imgdir", pImgDir, 1) & DBV("imgsize", pImgsize, 1) & _
                        DBV("prtfg", pPrtFg, 1) & DBV("rmk", pRmk, 1) & DBV("icd", pIcd) & ")"



End Function
 
Public Function DeleteGraphImages(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String, Optional pSeq As String = "") As String
    'lab310 slide 이미지 내역 삭제
    DeleteGraphImages = " delete " & _
                   "   FROM " & T_LAB310 & _
                   "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                   "    AND " & DBW("accdt", pAccDt, 2) & _
                   "    AND " & DBW("accseq", pAccSeq, 2) & _
                   "    AND " & DBW("testcd", pTestCd, 2)
    
    If pSeq <> "" Then DeleteGraphImages = DeleteGraphImages & " AND " & DBW("seq", pSeq, 2)
End Function


Public Function SqlGetImagesList(ByVal pFrDt As String, ByVal pToDt As String, _
                                 Optional pWorkArea As String = "", _
                                 Optional pAccDt As String = "", _
                                 Optional pAccSeq As String = "", _
                                 Optional pTestCd As String = "", _
                                 Optional pPtid As String = "", _
                                 Optional pIcd As String = "", _
                                 Optional pChkField As String = "") As String
                                 
    'lab310 slide 이미지 내역 리스트
    SqlGetImagesList = " SELECT distinct a.workarea,a.accdt,a.accseq,a.ptid,d." & F_PTNM & " as ptnm,a.testcd,stscd,b.spccd,c.field3 as spcnm, b.vfydt, b.vfytm, b.testdiv " & _
                       "   FROM " & T_HIS001 & " d, " & T_LAB032 & " c, " & T_LAB201 & " b, " & T_LAB310 & " a " & _
                       "  WHERE " & DBW("a.rcvdt>=", pFrDt) & _
                       "    AND " & DBW("a.rcvdt<=", pToDt)
                       
    Select Case pChkField
        Case "0":
            If pWorkArea <> "" Then SqlGetImagesList = SqlGetImagesList & _
                       "    AND " & DBW("a.workarea", pWorkArea, 2)
        Case "1":
            SqlGetImagesList = SqlGetImagesList & _
                       "    AND " & DBW("a.ptid", pPtid, 2)
        Case "2":
            SqlGetImagesList = SqlGetImagesList & _
                       "    AND " & DBW("a.workarea", pWorkArea, 2) & _
                       "    AND " & DBW("a.accdt", pAccDt, 2) & _
                       "    AND " & DBW("a.accseq", pAccSeq, 2)
        Case "3":
            SqlGetImagesList = SqlGetImagesList & _
                       "    AND " & DBW("a.icd", pIcd, 2)
    End Select
    
    
    SqlGetImagesList = SqlGetImagesList & _
                       "    AND a.workarea = b.workarea AND a.accdt = b.accdt AND a.accseq = b.accseq " & _
                       "    AND " & DBW("c.cdindex", LC3_Specimen, 2) & _
                       "    AND b.spccd = c.cdval1 " & _
                       "    AND a.ptid = d." & F_PTID & _
                       "  order by a.workarea,a.accdt,a.accseq desc "
                       
    
End Function

'Public Function GetICDCd(Optional ByVal ChkLoad As Boolean = True, Optional pCd As String = "") As String
''FROM HIS006(상병 마스터)
''" WHERE (expdt=' ' or expdt is null) "
'    If ChkLoad = True Then
'        GetICDCd = "SELECT " & F_ICD & " as icd, " & F_IENM & " as ienm FROM " & T_HIS006 & _
'                   " WHERE " & DBW(F_ICD, pCd, 2)
'    Else
'        GetICDCd = "SELECT " & F_ICD & " as icd, " & F_IENM & " as ienm FROM " & T_HIS006
'    End If
'End Function


Public Function SqlGetImageReportList(ByVal pFrDt As String, ByVal pToDt As String, _
                                      Optional pPtid As String, Optional pWorkArea As String = "", _
                                      Optional pAccDt As String = "", Optional pAccSeq As String = "", _
                                      Optional chkString As String = "0", Optional pPrtFg As String = "1") As String
    Select Case chkString
        Case "0":
            SqlGetImageReportList = " SELECT a.workarea, a.accdt, a.accseq, a.stscd, a.examdt, a.examtm, a.examdoct, a.ptid, a.ordcd, " & _
                                    "        d." & F_PTNM & " as ptnm, d." & F_SEX & " as sex, " & F_DOB2("d") & " as dob, " & _
                                    "        c.deptcd, c.wardid, c.hosilid, c.majdoct, c.orddt, c.bussdiv, " & _
                                    "        b.imgdir, b.imgsize, e.testdiv " & _
                                    "   FROM " & T_LAB201 & " e," & T_HIS001 & " d, " & T_LAB101 & " c, " & T_LAB310 & " b, " & T_LAB102 & " a " & _
                                    "  WHERE " & DBW("a.examdt >=", pFrDt) & _
                                    "    AND " & DBW("a.examdt <=", pToDt)
            
            
            If Trim(pPtid) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.ptid", pPtid, 2)
            If Trim(pWorkArea) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.workarea", pWorkArea, 2)
            If Trim(pAccDt) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.accdt", pAccDt, 2)
            If Trim(pAccSeq) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.accseq", pAccSeq, 2)
            
            SqlGetImageReportList = SqlGetImageReportList & _
                                    "    AND " & DBW("a.stscd > ", StsCd_LIS_MidRst) & _
                                    "    AND a.workarea =b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq AND a.ordcd = b.testcd " & _
                                    "    AND " & DBW("b.prtfg", pPrtFg, 2) & _
                                    "    AND (b.rptdt is null or b.rptdt =' ') " & _
                                    "    AND a.ptid=c.ptid AND a.orddt=c.orddt AND a.ordno=c.ordno AND a.ptid = d." & F_PTID & _
                                    "    AND a.workarea = e.workarea AND a.accdt = e.accdt AND a.accseq = e.accseq "
                    
        Case "1":
            SqlGetImageReportList = " SELECT a.workarea, a.accdt, a.accseq, a.stscd, a.examdt, a.examtm, a.examdoct, a.ptid, a.ordcd, " & _
                                    "        d." & F_PTNM & " as ptnm, d." & F_SEX & " as sex, " & F_DOB2("d") & " as dob, " & _
                                    "        c.deptcd, c.wardid, c.hosilid, c.majdoct, c.orddt, c.bussdiv, " & _
                                    "        b.imgdir, b.imgsize, e.testdiv " & _
                                    "   FROM " & T_LAB201 & " e," & T_HIS001 & " d, " & T_LAB101 & " c, " & T_LAB310 & " b, " & T_LAB102 & " a " & _
                                    "  WHERE " & DBW("a.examdt >=", pFrDt) & _
                                    "    AND " & DBW("a.examdt <=", pToDt)
            
            
            If Trim(pPtid) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.ptid", pPtid, 2)
            If Trim(pWorkArea) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.workarea", pWorkArea, 2)
            If Trim(pAccDt) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.accdt", pAccDt, 2)
            If Trim(pAccSeq) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.accseq", pAccSeq, 2)
            
            SqlGetImageReportList = SqlGetImageReportList & _
                                    "    AND " & DBW("a.stscd > ", StsCd_LIS_MidRst) & _
                                    "    AND a.workarea =b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq AND a.ordcd = b.testcd " & _
                                    "    AND " & DBW("b.prtfg", pPrtFg, 2) & _
                                    "    AND b.rptdt is not null " & _
                                    "    AND a.ptid=c.ptid AND a.orddt=c.orddt AND a.ordno=c.ordno AND a.ptid = d." & F_PTID & _
                                    "    AND a.workarea = e.workarea AND a.accdt = e.accdt AND a.accseq = e.accseq "
        Case "2":
            SqlGetImageReportList = " SELECT a.workarea, a.accdt, a.accseq, a.stscd, a.examdt, a.examtm, a.examdoct, a.ptid, a.ordcd, " & _
                                    "        d." & F_PTNM & " as ptnm, d." & F_SEX & " as sex, " & F_DOB2("d") & " as dob, " & _
                                    "        c.deptcd, c.wardid, c.hosilid, c.majdoct, c.orddt, c.bussdiv, " & _
                                    "        b.imgdir, b.imgsize, e.testdiv " & _
                                    "   FROM " & T_LAB201 & " e," & T_HIS001 & " d, " & T_LAB101 & " c, " & T_LAB310 & " b, " & T_LAB102 & " a " & _
                                    "  WHERE " & DBW("a.examdt >=", pFrDt) & _
                                    "    AND " & DBW("a.examdt <=", pToDt)
            
            
            If Trim(pPtid) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.ptid", pPtid, 2)
            If Trim(pWorkArea) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.workarea", pWorkArea, 2)
            If Trim(pAccDt) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.accdt", pAccDt, 2)
            If Trim(pAccSeq) <> "" Then SqlGetImageReportList = SqlGetImageReportList & " AND " & DBW("a.accseq", pAccSeq, 2)
            
            SqlGetImageReportList = SqlGetImageReportList & _
                                    "    AND " & DBW("a.stscd > ", StsCd_LIS_MidRst) & _
                                    "    AND a.workarea =b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq AND a.ordcd = b.testcd " & _
                                    "    AND " & DBW("b.prtfg", pPrtFg, 2) & _
                                    "    AND a.ptid=c.ptid AND a.orddt=c.orddt AND a.ordno=c.ordno AND a.ptid = d." & F_PTID & _
                                    "    AND a.workarea = e.workarea AND a.accdt = e.accdt AND a.accseq = e.accseq "
    End Select
End Function

Public Function SqlGetImageModfy(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pTestCd As String) As String
    
    SqlGetImageModfy = " SELECT mfydt " & _
                       "   FROM " & T_LAB308 & _
                       "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                       "    AND " & DBW("accdt", pAccDt, 2) & _
                       "    AND " & DBW("accseq", pAccSeq, 2) & _
                       "    AND " & DBW("testcd", pTestCd, 2)
                       
                       
    SqlGetImageModfy = SqlGetImageModfy & " UNION ALL"
    
    SqlGetImageModfy = SqlGetImageModfy & _
                       " SELECT vfydt as mfydt " & _
                       "   FROM " & T_LAB351 & _
                       "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                       "    AND " & DBW("accdt", pAccDt, 2) & _
                       "    AND " & DBW("accseq", pAccSeq, 2) & _
                       "    AND " & DBW("testcd", pTestCd, 2)
                       
    SqlGetImageModfy = SqlGetImageModfy & " UNION ALL"
    
    SqlGetImageModfy = SqlGetImageModfy & _
                       " SELECT vfydt as mfydt " & _
                       "   FROM " & T_LAB404 & _
                       "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                       "    AND " & DBW("accdt", pAccDt, 2) & _
                       "    AND " & DBW("accseq", pAccSeq, 2) & _
                       "    AND " & DBW("testcd", pTestCd, 2)
    
End Function

Public Function SqlUpDateImageRpt(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pTestCd As String, _
                                  ByVal pRptDt As String, ByVal pRptTm As String, _
                                  ByVal pRptId As String) As String

    SqlUpDateImageRpt = " update " & T_LAB310 & _
                        "    set " & _
                                   DBW("rptdt", pRptDt, 3) & _
                                   DBW("rpttm", pRptTm, 3) & _
                                   DBW("rptid", pRptId, 2) & _
                        "  WHERE " & DBW("workarea", pWorkArea, 2) & _
                        "    AND " & DBW("accdt", pAccDt, 2) & _
                        "    AND " & DBW("accseq", pAccSeq, 2) & _
                        "    AND " & DBW("testcd", pTestCd, 2) & _
                        "    AND prtfg ='1' "
End Function


Public Function SqlImageStatistic(ByVal pFrDt As String, ByVal pToDt As String, _
                                  ByVal pWorkArea As String, ByVal pPtid As String, _
                                  ByVal pBussDiv As String, ByVal pSex As String, _
                                  ByVal pFrAge As String, ByVal pToAge As String, _
                                  ByVal pWardId As String, ByVal pDoct As String, _
                                  ByVal pSpcCd As String, ByVal pDeptcd As String, _
                                  ByVal pTestCd As String, ByVal pIcd As String, _
                                  ByVal pPrtFg As String, Optional ByVal pTcd As String = "") As String

    SqlImageStatistic = " SELECT a.ptid, a.sex, a.ageday, a.spccd, a.rcvdt, a.rcvtm, a.rcvid, " & _
                        "        a.vfydt, a.vfytm, a.vfyid, a.coldt, a.coltm, a.colid, " & _
                        "        a.deptcd, a.wardid, a.hosilid, a.majdoct, a.orddoct, " & _
                        "        b.workarea, b.accdt, b.accseq, b.testcd, b.seq, b.imgdir, b.rmk, b.icd, " & _
                        "        c.abbrnm10 as testnm, d.field3 as spcnm, f." & F_PTNM & " as ptnm " & _
                        "   FROM " & T_HIS001 & " f, " & T_LAB032 & " d, " & T_LAB001 & " c, " & T_LAB310 & " b, " & T_LAB201 & " a " & _
                        "  WHERE " & DBW("a.vfydt>=", pFrDt) & _
                        "    AND " & DBW("a.vfydt<=", pToDt)
                        
                     
    If pWorkArea <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.workarea", pWorkArea, 2)
    If pPtid <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.ptid", pPtid, 2)
    If pBussDiv <> "" Then
        If pBussDiv = "1" Then
            SqlImageStatistic = SqlImageStatistic & " AND (a.wardid is null or a.wardid = ' ') "
        Else
            SqlImageStatistic = SqlImageStatistic & " AND (a.wardid is not null or a.wardid <>' ') "
        End If
    End If
    
    If pSex <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.sex", pSex, 2)
    If pFrAge <> "" And pToAge <> "" Then
        SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.ageday >=", pFrAge) & _
                                                " AND " & DBW("a.ageday <=", pToAge)
    End If
    
'    If pWardId <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.wardid", pWardId, 2)
'    If pDoct <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.majdoct", pDoct, 2)
'    If pSpcCd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.spccd", pSpcCd, 2)
'    If pDeptcd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("a.deptcd", pDeptcd, 2)
'    If pTestCd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("b.testcd", pTestCd, 2)
'    If pIcd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("b.icd", pIcd, 2)
    
    If pWardId <> "" Then SqlImageStatistic = SqlImageStatistic & " AND a.wardid in (" & pWardId & ") "
    If pDoct <> "" Then SqlImageStatistic = SqlImageStatistic & " AND a.majdoct in (" & pDoct & ") "
    If pSpcCd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND a.spccd in (" & pSpcCd & ") "
    If pDeptcd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND a.deptcd in (" & pDeptcd & ") "
    If pTestCd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND b.testcd in (" & pTestCd & ") "
    If pIcd <> "" Then SqlImageStatistic = SqlImageStatistic & " AND b.icd in (" & pIcd & ") "
    
    If pPrtFg <> "" Then SqlImageStatistic = SqlImageStatistic & " AND " & DBW("b.prtfg", pPrtFg, 2)
    
    SqlImageStatistic = SqlImageStatistic & " AND a.workarea = b.workarea AND a.accdt= b.accdt AND a.accseq = b.accseq " & _
                                            " AND a.ptid = f." & F_PTID & " AND b.testcd = c.testcd " & _
                                            " AND " & DBW("d.cdindex", LC3_Specimen, 2) & " AND a.spccd = d.cdval1 "
                                            
                                            
    SqlImageStatistic = SqlImageStatistic & " order by a.vfydt,b.workarea, b.accdt, b.accseq, b.seq "
                                 
End Function

Public Function GetLabNo(ByVal qSpcyy As String, ByVal qSpcno As String, _
                         ByRef tmpWA As String, ByRef tmpAccdt As String, _
                         ByRef tmpAccSeq As String)
    Dim rs   As Recordset
    Dim sSQL As String
    
    sSQL = " SELECT workarea,accdt,accseq FROM " & T_LAB201 & _
           " WHERE " & DBW("spcyy=", qSpcyy) & _
           " AND " & DBW("spcno=", qSpcno)
    Set rs = New Recordset
    rs.Open sSQL, DBConn
    
    If Not rs.EOF Then
        tmpWA = rs.Fields("workarea").Value & ""
        tmpAccdt = rs.Fields("accdt").Value & ""
        tmpAccSeq = rs.Fields("accseq").Value & ""
    End If

    Set rs = Nothing

End Function

Public Function GetItemInfo(ByVal vTestCd As String) As String
'검사코드를 인수로 검사항목의 정보를 구하는 sql
    Dim strSQL As String
    
' 2018.05.28 majestic seq 막음. 검체별 폐기일 문제
'    strSQL = " select a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
'            "        b.labelcnt, b.statflags, a.testdiv, c.field1 as multifg, c.field2 as spcgrp, c.field5 as spcnm, " & _
'            "        d.field2 as labdiv, e.field2 as labrange, '1' insurfg " & _
'            " from " & T_LAB032 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & _
'                       T_LAB004 & " b, " & T_LAB001 & " a " & _
'            " where  " & DBW("a.testcd=", vTestCd) & _
'            " and a.applydt = ( select max(applydt) from " & T_LAB001 & _
'            "                     where testcd = a.testcd ) " & _
'            " and   (a.detailfg = '' or a.detailfg is null) " & _
'            " and    a.testcd = b.testcd " & _
'            " and    b.seq = ( select min(seq) from " & T_LAB004 & _
'            "                  where testcd = b.testcd ) " & _
'            " and   (b.expdt = '' or b.expdt is null)" & _
'            " and    b.applydt = ( select max(applydt) from " & T_LAB004 & _
'            "                      where testcd = b.testcd and spccd = b.spccd and seq=b.seq) " & _
'            " and    c.cdindex = 'C215' " & _
'            " and    c.cdval1 = b.spccd  " & _
'            " and    d.cdindex = 'C213' " & _
'            " and    d.cdval1 = a.workarea " & _
'            " and    " & DBJ("e.cdindex = 'C217'") & _
'            " and    " & DBJ("e.cdval1 =* c.field2")
            
    strSQL = " select a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
            "        b.labelcnt, b.statflags, a.testdiv, c.field1 as multifg, c.field2 as spcgrp, c.field5 as spcnm, " & _
            "        d.field2 as labdiv, e.field2 as labrange, '1' insurfg " & _
            " from " & T_LAB032 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & _
                       T_LAB004 & " b, " & T_LAB001 & " a " & _
            " where  " & DBW("a.testcd=", vTestCd) & _
            " and a.applydt = ( select max(applydt) from " & T_LAB001 & _
            "                     where testcd = a.testcd ) " & _
            " and   (a.detailfg = '' or a.detailfg is null) " & _
            " and    a.testcd = b.testcd " & _
            " and   (b.expdt = '' or b.expdt is null)" & _
            " and    b.applydt = ( select max(applydt) from " & T_LAB004 & _
            "                      where testcd = b.testcd and spccd = b.spccd and seq=b.seq) " & _
            " and    c.cdindex = 'C215' " & _
            " and    c.cdval1 = b.spccd  " & _
            " and    d.cdindex = 'C213' " & _
            " and    d.cdval1 = a.workarea " & _
            " and    " & DBJ("e.cdindex = 'C217'") & _
            " and    " & DBJ("e.cdval1 =* c.field2")
    GetItemInfo = strSQL
End Function



