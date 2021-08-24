VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+------------------------------------------------------+
'|    Sql Query for Orders, Collections, Accession      |
'|                                                      |
'|    작성자        : 김미경                            |
'|    작성일        : 1999.04.18                        |
'|                                                      |
'|  CopyRight(C) 1999 대련엠티에스                      |
'+------------------------------------------------------+

'%  0. 공통코드 마스터 (LAB031)
Public Function SqlCommonCode(ByVal TblName As String, ByVal CdIndex As String, _
                                                Optional ByVal CdVal1 As Variant, Optional ByVal CdVal2 As Variant) As String
   Dim tmpStr As String
   
   tmpStr = ""
   If Not IsMissing(CdVal1) Then tmpStr = tmpStr & "and cdval1 = '" & CdVal1 & "' "
   If Not IsMissing(CdVal2) Then tmpStr = tmpStr & "and cdval2 = '" & CdVal2 & "' "
   
   SqlCommonCode = "Select * " & _
                   "From   " & TblName & " " & _
                   "Where cdindex = '" & CdIndex & "' " & tmpStr
End Function

'%  1. Get CodeList
'%      - Calling from [frm351ItemMaster]
Public Function SqlLAB001CodeList() As String
   SqlLAB001CodeList = "Select a.testcd, a.testnm " & _
                       "From   " & T_LAB001 & " a " & _
                       "Where a.applydt = ( select max(applydt) from " & T_LAB001 & "  " & _
                       "                             where testcd = a.testcd ) " & _
                       "Order by a.testcd "
End Function

'%  2. Select from LAB001
'%      - Calling from [frm351ItemMaster]
Public Function SqlLAB001Read(ByVal TestCd As String) As String
   SqlLAB001Read = "Select * " & _
                   "From " & T_LAB001 & "  " & _
                   "Where testcd = '" & TestCd & "' " & _
                   "Order by applydt"
End Function

'%  3. Select from LAB004
'%      - Calling from [frm351ItemMaster]
Public Function SqlSpecimenRead(ByVal TestCd As String) As String

   SqlSpecimenRead = "Select a.spccd as spccd, b.field3 as spcnm, b.field2 as spcgrpcd, c.field1 as spcgrp, a.seq " & _
                     "From " & T_LAB004 & " a, " & T_LAB032 & " b, " & T_LAB032 & " c " & _
                     "Where a.testcd = '" & TestCd & "' " & _
                     "and     b.cdindex = '" & CD2_Specimen & "' " & _
                     "and     b.cdval1 = a.spccd " & _
                     "and     c.cdindex = '" & CD2_SGroup & "' " & _
                     "and     c.cdval1 =* b.field2 " & _
                     "Group by a.spccd, b.field3, b.field2, c.field1, a.seq " & _
                     "Order by seq"
End Function

'%  4. Select from LAB004
'%      - Calling from [frm351ItemMaster]
Public Function SqlLAB004Read(ByVal TestCd As String, ByVal SpcCd As String, _
                                              ByVal Seq As String) As String
   SqlLAB004Read = "Select a.*, b.field2 as sgroup " & _
                   "From " & T_LAB004 & " a, " & T_LAB032 & " b " & _
                   "Where a.testcd = '" & TestCd & "' " & _
                   "and     a.spccd = '" & SpcCd & "' " & _
                   "and     a.seq = " & Seq & " " & _
                   "and     b.cdindex = '" & CD2_Specimen & "' " & _
                   "and     b.cdval1 =* a.spccd " & _
                   "Order by applydt"
End Function

'%  5. Get Common Code List
'%      - Calling from [frm352ItemMaster]
Public Function SqlLAB032CodeList(ByVal CdIndex As String, ByVal Fields As String, _
                                                   Optional ByVal CdVal1 As Variant, _
                                                   Optional ByVal Orderby As Variant) As String
   SqlLAB032CodeList = "Select " & Fields & " " & _
                       "From " & T_LAB032 & " " & _
                       "Where cdindex = '" & CdIndex & "' "
   If IsMissing(CdVal1) = False Then
      SqlLAB032CodeList = SqlLAB032CodeList & _
                        "and     cdval1 = '" & CdVal1 & "' "
   End If
   If IsMissing(Orderby) = False Then
      SqlLAB032CodeList = SqlLAB032CodeList & Orderby
   End If

End Function

'%  6. Get ApplyDt from LAB005
'%      - Calling from [frm353ItemMaster]
Public Function SqlLAB005AppDt(ByVal TestCd As String, ByVal SpcCd As String) As String
   SqlLAB005AppDt = "Select applydt " & _
                    "From " & T_LAB005 & " " & _
                    "Where testcd = '" & TestCd & "' " & _
                    "and     spccd = '" & SpcCd & "' " & _
                    "Group by applydt " & _
                    "Order by applydt"
End Function

'%  7. Select from LAB005
'%      - Calling from [frm353ItemMaster]
Public Function SqlLAB005Read(ByVal TestCd As String, ByVal SpcCd As String, _
                                              ByVal ApplyDt As String) As String
   SqlLAB005Read = "Select * " & _
                   "From " & T_LAB005 & " " & _
                   "Where testcd = '" & TestCd & "' " & _
                   "and     spccd = '" & SpcCd & "' " & _
                   "and     applydt = '" & ApplyDt & "' " & _
                   "Order by applysex, agefrom"
End Function


'%  7-1. Get Reference Range for a Patient
'%      - Calling from [frm401ResultReview]
Public Function SqlGetReference(ByVal strTestCd As String, _
                                              ByVal strSpcCd As String, _
                                              ByVal strApplyDt As String, _
                                              ByVal strApplySex As String, _
                                              Optional ByVal intAge As Variant, _
                                              Optional ByVal intAgeFrom As Variant, _
                                              Optional ByVal intAgeTo As Variant) As String

   SqlGetReference = "Select * " & _
                     "From " & T_LAB005 & " " & _
                     "Where testcd = '" & strTestCd & "' " & _
                     "and    spccd = '" & strSpcCd & "' " & _
                     "and    applydt <= '" & strApplyDt & "' " & _
                     "and    applysex = '" & strApplySex & "' "
                 
   If Not (IsMissing(intAgeFrom) Or IsMissing(intAgeTo)) Then
      SqlGetReference = SqlGetReference & _
                 "and    agefrom = " & intAgeFrom & " " & _
                 "and    ageto = " & intAgeTo & " "
   ElseIf Not IsMissing(intAge) Then
      SqlGetReference = SqlGetReference & _
                 "and    agefrom <= " & intAge & " " & _
                 "and    ageto >= " & intAge
   End If
   
   SqlGetReference = SqlGetReference & " order by applydt desc"
   
End Function
   

'%  8. Get High Frequency Test List
'%      - intDiv : 1 -  (다빈도처방 리스트), 2 - (전체 처방 리스트)
'%      - Calling from [modOrder]
'%      - 검사명,검사코드,검체코드,응급여부,WorkArea,보관구분,아침채혈여부,
'%         검사구분,복수검체여부,접수번호부여기준,검체군

Public Function SqlItemList(ByVal intDiv As Integer) As String
   If intDiv = 1 Then
      SqlItemList = "Select a.testnm, a.abbrnm5, b.cdval1 as testcd, b.field1 as spccd, c.statfg, a.workarea, " & _
                    "          c.storecd, c.rndfg, c.labelcnt, c.statflags, a.testdiv, d.field1 as MultiFg, d.field2 as SpcGrp " & _
                    "From " & T_LAB001 & " a, " & T_LAB032 & " b, " & T_LAB004 & " c, " & T_LAB032 & " d " & _
                    "Where b.cdindex = '" & CD2_HighItem & "' " & _
                    "and     b.cdval1 = a.testcd " & _
                    "and     a.applydt = ( select max(applydt) from " & T_LAB001 & "  " & _
                    "                             where testcd = a.testcd ) " & _
                    "and     (a.detailfg is null " & _
                    "or         a.detailfg = '') " & _
                    "and     c.testcd = b.cdval1 " & _
                    "and     c.spccd = b.field1 " & _
                    "and     c.applydt = ( select max(applydt) from " & T_LAB004 & " " & _
                    "                             where testcd = c.testcd  " & _
                    "                             and     spccd = c.spccd ) " & _
                    "and     d.cdindex = '" & CD2_Specimen & "' " & _
                    "and     d.cdval1 = c.spccd  "
   Else
      SqlItemList = "Select a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
                    "         b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg, c.field2 as SpcGrp " & _
                    "From " & T_LAB001 & " a, " & T_LAB004 & " b, " & T_LAB032 & " c " & _
                    "Where a.applydt = ( select max(applydt) from " & T_LAB001 & " " & _
                    "                             where testcd = a.testcd ) " & _
                    "and     a.detailfg = ''  " & _
                    "and     a.testcd = b.testcd " & _
                    "and     b.seq = 1 " & _
                    "and     b.expdt = '' " & _
                    "and     c.cdindex = '" & CD2_Specimen & "' " & _
                    "and     c.cdval1 = b.spccd  "
   End If
End Function


'% 9. Get the Patient's Infromation : 환자정보
'%     - PtntId : 환자 ID
Public Function SqlHIS001Read(ByVal PtntId As String) As String
   SqlHIS001Read = "Select * " & _
                            "From " & T_HIS001 & " " & _
                            "Where ptnt_no = " & PtntId & " "
End Function


'%  10. Get Doctor List : 의사Id 리스트
'%      - Calling from [frm101Order]
Public Function SqlHIS007CodeList(Optional ByVal EmpId As Variant) As String
   If IsMissing(EmpId) Then
      SqlHIS007CodeList = "Select a.empl_no as EmpId, a.empl_nm as EmpNm " & _
                          "From   " & T_HIS005 & " a " & _
                          "Where    a.job_type_gb = '05' " & _
                          "Order by a.empl_nm "
   Else
      SqlHIS007CodeList = "Select a.empl_no as EmpId, a.empl_nm as EmpNm " & _
                          "From   " & T_HIS005 & " a " & _
                          "where a.empl_no = " & EmpId
   End If
End Function

'%  11. Get Department List : 부서리스트
'%      - Calling from [frm101Order]
Public Function SqlHIS003CodeList(Optional ByVal DeptCd As Variant, Optional ByVal DeptClass As String = "B") As String
   If IsMissing(DeptCd) Then
      SqlHIS003CodeList = "Select a.dept_cd as DeptCd, a.dept_nm as DeptNm " & _
                          "From   " & T_HIS003 & " a " & _
                          "Where    a.hosp_gb = '" & HosptGb & "' " & _
                          "and    a.dept_class_cd2 = '" & DeptClass & "' " & _
                          "Order by a.dept_nm "
   Else
      SqlHIS003CodeList = "Select a.dept_cd as DeptCd, a.dept_nm as DeptNm, a.center_gb as BldGb " & _
                          "From   " & T_HIS003 & " a " & _
                          "where a.hosp_gb =  '" & HosptGb & "' " & _
                          "and    a.dept_cd = '" & DeptCd & "' "
   End If
End Function


'%  12. Get Ward Id List : 병동 리스트
'%      - Calling from [frm101Order]
Public Function SqlHIS004CodeList(Optional ByVal WardId As Variant) As String
   If IsMissing(WardId) Then
        SqlHIS004CodeList = "Select a.dept_cd as WardId, a.dept_nm as WardNm " & _
                            "From   " & T_HIS004 & " a " & _
                            "Where    a.hosp_gb = '" & HosptGb & "' " & _
                            "and    a.dept_class_cd2 = 'C' " & _
                            "Order by WardId "
   Else
        SqlHIS004CodeList = "Select a.dept_cd as WardId, a.dept_nm as WardNm " & _
                            "From   " & T_HIS004 & " a " & _
                            "Where    a.hosp_gb = '" & HosptGb & "' " & _
                            "and    a.dept_cd = '" & WardId & "' "
   End If
   'SqlHIS004CodeList = "Select a.wardid, a.wardnm " & _
                                  "From   " & T_HIS004 & " a " & _
                                  "Order by a.wardid "
End Function


'%  13. Get Specimen List
'%      - Calling from [frm101Order]
'%      - 검체코드,검체명,응급여부,보관구분,복수검체여부,검체군,우선순위
Public Function SqlSpecList(ByVal ParaTestCd As String) As String
   SqlSpecList = "Select a.spccd, b.field3 as SpcNm, a.statfg, a.storecd, a.labelcnt, a.statflags, " & _
                        "         b.field1 as MultiFg, b.field2 as SpcGrp, a.seq  " & _
                        "From   " & T_LAB004 & " a, " & T_LAB032 & " b " & _
                        "Where a.testcd = '" & ParaTestCd & "' " & _
                        "and     a.applydt = ( select max(applydt) from " & T_LAB004 & " " & _
                        "                             where testcd = a.testcd  " & _
                        "                             and     spccd = a.spccd ) " & _
                        "and    (a.expdt is null " & _
                        "or        a.expdt = '') " & _
                        "and     b.cdindex = '" & CD2_Specimen & "' " & _
                        "and     b.cdval1 = a.spccd " & _
                        "Order by a.seq "
End Function


'%  14. 직원정보 검색
'%      - Calling from [frm101Order]
Public Function SqlLAB015Read(ByVal paraEmpId As String, ByVal Index As Integer) As String
   If Index = 1 Then
      SqlLAB015Read = "Select * " & _
                               "From   " & T_LAB015 & " a " & _
                               "Where a.logonid = '" & paraEmpId & "'"
   Else
      SqlLAB015Read = "Select * " & _
                               "From   " & T_LAB015 & " a " & _
                               "Where a.empid = " & paraEmpId
   End If
End Function


'%  15. 가장 마지막 처방번호
'%      - Calling from [frm101Order]
Public Function SqlLastOrdNo(ByVal paraPtId As String, ByVal paraDate As String) As String
      SqlLastOrdNo = "Select ordno " & _
                               "From   " & T_LAB101 & " " & _
                               "Where ptid = " & paraPtId & " " & _
                               "and    orddt = '" & paraDate & "' " & _
                               "Order by ordno desc"
End Function


'%  16. 검체번호
'%      - Calling from [frm101Order]
Public Function SqlSpcNo(ByVal SpcYY As String, ByVal flag As Integer, Optional ByVal Seq As Variant) As String
   Select Case flag
   Case 1:
      SqlSpcNo = "Select seq " & _
                        "From   " & T_LAB099 & " " & _
                        "Where cdindex = '" & NO_Specimen & "' " & _
                        "and     divcd1 = 'SPCNO' " & _
                        "and     divcd2 = '" & SpcYY & "' " & _
                        "and     divcd3 = '0' "
   Case 2:
      SqlSpcNo = "Insert Into " & T_LAB099 & " (cdindex, divcd1, divcd2, divcd3, seq) " & _
                       "Values ('" & NO_Specimen & "', 'SPCNO', '" & SpcYY & "', '0', " & Seq & ")"
   Case 3:
      SqlSpcNo = "Update " & T_LAB099 & " Set seq = " & Seq & " " & _
                        "Where cdindex = '" & NO_Specimen & "' " & _
                        "and     divcd1 = 'SPCNO' " & _
                        "and     divcd2 = '" & SpcYY & "' " & _
                        "and     divcd3 = '0' "
   End Select
End Function

'%  17. 접수번호
'%      - Calling from [frm101Order]
Public Function SqlLabNo(ByVal WorkArea As String, ByVal AccDt As String, ByVal SpcGrp As String, _
                                      ByVal flag As Integer, Optional ByVal Seq As Variant) As String
   Select Case flag
   Case 1:
      SqlLabNo = "Select seq " & _
                        "From   " & T_LAB099 & " " & _
                        "Where cdindex = '" & NO_LabNo & "' " & _
                        "and    divcd1 = '" & WorkArea & "' " & _
                        "and    divcd2 = '" & AccDt & "' " & _
                        "and    divcd3 = '" & SpcGrp & "' "
   Case 2:
      SqlLabNo = "Insert Into " & T_LAB099 & " (cdindex, divcd1, divcd2, divcd3, seq) " & _
                       "Values ('" & NO_LabNo & "', '" & WorkArea & "', '" & AccDt & "', '" & SpcGrp & "', " & Seq & ")"
   Case 3:
      SqlLabNo = "Update " & T_LAB099 & " Set seq = " & Seq & " " & _
                        "Where cdindex = '" & NO_LabNo & "' " & _
                        "and     divcd1 = '" & WorkArea & "' " & _
                        "and     divcd2 = '" & AccDt & "' " & _
                        "and     divcd3 = '" & SpcGrp & "' "
   End Select
   
End Function


'%  18. 처방 Header(LAB101) / Body(LAB102) Insert
'%      - Calling from [clsOrder]
Public Function SqlSaveOrder(ByVal flag As Integer) As String
                                      
   If flag = 1 Then
      SqlSaveOrder = "Insert into " & _
                     "" & T_LAB101 & " (PtId, OrdDt, OrdTm, OrdNo, BussDiv, BedInDt, ReqDt, " & _
                     "             ReqTm, DeptCd, OrdDoct, MajDoct, EntId, EntDt, " & _
                     "             EntTm, OrdDiv, OrgAccNo, SpOrdDiv, DoneFg) " & _
                     "Values  (:PtId, :OrdDt, :OrdTm, :OrdNo, :BussDiv, :BedInDt, :ReqDt, " & _
                     "             :ReqTm, :DeptCd, :OrdDoct, :MajDoct, :EntId, :EntDt, " & _
                     "             :EntTm, :OrdDiv, :OrgAccNo, :SpOrdDiv, '0') "
   Else
      SqlSaveOrder = "Insert into " & _
                     "" & T_LAB102 & " (PtId, OrdDt, OrdNo, OrdSeq, OrdCd, spccd, StoreCd, DcFg, DcDt, DcNo, " & _
                     "             AttrCd, ExamDt, ExamTm, ExamDoct, StsCd, StatFg, InsDiv) " & _
                     "Values  (:PtId, :OrdDt, :OrdNo, :OrdSeq, :OrdCd, :SpcCd, :StoreCd, :DcFg, :DcDt, :DcNo, " & _
                     "             :AttrCd, :ExamDt, :ExamTm, :ExamDoct, :StsCd, :StatFg, :InsDiv) "
   End If
  
  End Function

'%  19. 채혈접수내역 LAB201 Insert
'%      - Calling from [clsOrder]
Public Function SqlSaveCollection() As String

   SqlSaveCollection = "Insert into " & _
                                "" & T_LAB201 & " (SpcYy, SpcNo, PtId, Sex, AgeDay, BedInDt, DeptCd, OrdDoct, MajDoct, " & _
                                "            WorkArea, AccDt, AccSeq, StsCd, ReqTotCnt, ReqInputCnt, AltTotCnt, " & _
                                "            AltInputCnt, VfyDt, VfyTm, VfyId, ColDt, ColTm, ColId, RcvDt, RcvTm, " & _
                                "            RcvId, EntDt, EntTm, EntId, spccd, MultiFg, OrgAccNo, WardId, RoomId, " & _
                                "            BedId, FootNoteFg, StoreCd, RptFg, TestDiv, QcFg, RmkCd, StatFg) " & _
                                "Values  (:SpcYy, :SpcNo, :PtId, :Sex, :AgeDay, :BedInDt, :DeptCd, :OrdDoct, :MajDoct, " & _
                                "            :WorkArea, :AccDt, :AccSeq, :StsCd, :ReqTotCnt, :ReqInputCnt, :AltTotCnt, " & _
                                "            :AltInputCnt, :VfyDt, :VfyTm, :VfyId, :ColDt, :ColTm, :ColId, :RcvDt, :RcvTm, " & _
                                "            :RcvId, :EntDt, :EntTm, :EntId, :SpcCd, :MultiFg, :OrgAccNo, :WardId, :RoomId, " & _
                                "            :BedId, :FootNoteFg, :StoreCd, :RptFg, :TestDiv, :QcFg, :RmkCd, :StatFg) "

End Function

'%  20. 처방Body(LAB102)에 접수번호 Update
'%      - Calling from [clsCollection]
Public Function SqlUpdateLabNo(ByVal PtId As String, ByVal OrdDt As String, _
                                                ByVal OrdNo As Integer, ByVal OrdSeq As String) As String
   SqlUpdateLabNo = "Update " & T_LAB102 & " " & _
                              "Set     WorkArea = :WorkArea , " & _
                              "          AccDt = :AccDt , " & _
                              "          AccSeq = :AccSeq , " & _
                              "          StsCd = '" & STS_HaveSpc & "' " & _
                              "Where PtId = '" & PtId & "' " & _
                              "and     OrdDt = '" & OrdDt & "' " & _
                              "and     OrdNo = " & OrdNo & " " & _
                              "and     OrdSeq = " & OrdSeq
End Function


'%  21. 연속검사내역(LAB203) Insert
'%      - Calling from [clsCollection]
Public Function SqlSaveMultiSpc() As String
   SqlSaveMultiSpc = "Insert into " & T_LAB203 & "( PtId, OrdDt, OrdNo, OrdSeq, SpcSeq, spccd, WorkArea, AccDt, AccSeq ) " & _
                              "Values ( :PtId, :OrdDt, :OrdNo, :OrdSeq, :SpcSeq, :SpcCd, :WorkArea, :AccDt, :AccSeq )"
End Function


'%  22. 접수번호로 처방Body 검색
'%      - Calling from [clsAccession]
Public Function SqlOrderItems(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, Optional ByVal pMultiFg As String) As String
   'If pMultiFg <> "" Then
   SqlOrderItems = "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd, a.spccd, a.attrcd, a.statfg, " & _
                          "          b.testdiv, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                          "          e.testdiv as L201TestDiv, e.multifg, e.spccd as ColSpcCd, d.spcseq, " & _
                          "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, 0 as LastVfyId " & _
                          "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                        T_LAB203 & " d, " & T_LAB201 & " e " & _
                          "Where d.workarea = '" & pWorkArea & "' " & _
                          "and    d.accdt = '" & pAccDt & "' " & _
                          "and    d.accseq = " & pAccSeq & " " & _
                          "and    e.workarea = d.workarea " & _
                          "and    e.accdt = d.accdt " & _
                          "and    e.accseq = d.accseq " & _
                          "and    e.multifg <> '' " & _
                          "and    a.ptid = d.ptid " & _
                          "and    a.orddt = d.orddt " & _
                          "and    a.ordno = d.ordno " & _
                          "and    a.ordseq = d.ordseq " & _
                          "and    b.testcd = a.ordcd " & _
                          "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                          "and    c.testcd = a.ordcd " & _
                          "and    c.spccd = a.spccd " & _
                          "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) "
   'Else
   SqlOrderItems = SqlOrderItems & " Union All " & _
                          "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd, a.spccd, a.attrcd, a.statfg, " & _
                          "          b.testdiv, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                          "          e.testdiv as L201TestDiv, e.multifg, e.spccd as ColSpcCd, '' as SpcSeq, " & _
                          "          d.rstcd as LastRst, d.vfydt as LastVfyDt, d.vfytm as LastVfyTm, d.vfyid as LastVfyId " & _
                          "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB302 & " d, " & T_LAB201 & " e " & _
                          "Where a.workarea = '" & pWorkArea & "' " & _
                          "and    a.accdt = '" & pAccDt & "' " & _
                          "and    a.accseq = " & pAccSeq & " " & _
                          "and    e.workarea = a.workarea " & _
                          "and    e.accdt = a.accdt " & _
                          "and    e.accseq = a.accseq " & _
                          "and    e.multifg = '' " & _
                          "and    b.testcd = a.ordcd " & _
                          "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                          "and    c.testcd = a.ordcd " & _
                          "and    c.spccd = a.spccd " & _
                          "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) " & _
                          "and    d.ptid = a.ptid   and    d.testcd = a.ordcd " & _
                          "and    d.vfydt+d.vfytm = (select max(vfydt+vfytm) from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt <= convert(char(8),getdate(),112)) " & _
                          "and    exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt <= convert(char(8),getdate(),112)) "
   
   SqlOrderItems = SqlOrderItems & " Union All " & _
                          "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd, a.spccd, a.attrcd, a.statfg, " & _
                          "          b.testdiv, b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                          "          e.testdiv as L201TestDiv, e.multifg, e.spccd as ColSpcCd, '' as SpcSeq, " & _
                          "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, 0 as LastVfyId " & _
                          "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB201 & " e " & _
                          "Where a.workarea = '" & pWorkArea & "' " & _
                          "and    a.accdt = '" & pAccDt & "' " & _
                          "and    a.accseq = " & pAccSeq & " " & _
                          "and    e.workarea = a.workarea " & _
                          "and    e.accdt = a.accdt " & _
                          "and    e.accseq = a.accseq " & _
                          "and    e.multifg = '' " & _
                          "and    b.testcd = a.ordcd " & _
                          "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                          "and    c.testcd = a.ordcd " & _
                          "and    c.spccd = a.spccd " & _
                          "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) " & _
                          "and    not exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt <= convert(char(8),getdate(),112)) "
   'End If
End Function

'%  22. 접수번호로 처방Body 검색
'%      - Calling from [clsAccession]
Public Function SqlOrderItems1(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, Optional ByVal ItemCase As String) As String
   
    Dim sSelectClause(3) As String
    Dim sFromClause(3) As String
    Dim sWhereClause(3) As String
    
    Select Case ItemCase
    Case "Normal":
        sSelectClause(1) = "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, a.ordcd as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                                    "          b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                                    "          d.spccd, d.multifg, d.testdiv, 'N' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "          e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
        
        sFromClause(1) = "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB201 & " d "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
        
        sWhereClause(1) = "Where a.workarea = '" & pWorkArea & "' " & _
                                    "and    a.accdt = '" & pAccDt & "' " & _
                                    "and    a.accseq = " & pAccSeq & " " & _
                                    "and    d.workarea = a.workarea   and    d.accdt = a.accdt    and    d.accseq = a.accseq   and    d.multifg = ''  " & _
                                    "and    b.testcd = a.ordcd   and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                                    "and    b.panelfg = '" & PN_Normal & "' " & _
                                    "and    c.testcd = a.ordcd   and    c.spccd = a.spccd " & _
                                    "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                                    "and    e.ptid = a.ptid   and    e.testcd = a.ordcd " & _
                                    "and    e.vfydt = (select max(vfydt) from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt < d.rcvdt) " & _
                                    "and    exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "and    not exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " Union All " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Detail":
        sSelectClause(1) = "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, f.cdval2 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                                    "          b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                                    "          d.spccd, d.multifg, d.testdiv, 'Y' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "          e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
                                    
        sFromClause(1) = "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                                T_LAB201 & " d, " & T_LAB031 & " f, " & T_LAB001 & " g "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
                                    
        sWhereClause(1) = "Where a.workarea = '" & pWorkArea & "' " & _
                                    "and    a.accdt = '" & pAccDt & "' " & _
                                    "and    a.accseq = " & pAccSeq & " " & _
                                    "and    b.testcd = a.ordcd " & _
                                    "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                                    "and    b.panelfg = '" & PN_Detail & "' " & _
                                    "and    d.workarea = a.workarea   and    d.accdt = a.accdt   and    d.accseq = a.accseq   and    d.multifg = ''  " & _
                                    "and    f.cdindex = '" & CD1_Detail & "' " & _
                                    "and    f.cdval1 = a.ordcd " & _
                                    "and    g.testcd = f.cdval2 " & _
                                    "and    g.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = g.testcd) " & _
                                    "and    c.testcd = f.cdval2   and    c.spccd = a.spccd " & _
                                    "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                                    "and    e.ptid = a.ptid   and    e.testcd = f.cdval2 " & _
                                    "and    e.vfydt = (select max(vfydt) from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.cdval2 and vfydt <> '' and vfydt < d.rcvdt) " & _
                                    "and    exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.cdval2 and vfydt <> '' and vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "and    not exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = a.ordcd and vfydt <> '' and vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " Union All " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Group1":   '일반 그룹처방
        sSelectClause(1) = "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, f.field1 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                                    "          b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                                    "          d.spccd, d.multifg, d.testdiv, 'N' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "          e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
                                    
        sFromClause(1) = "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                                T_LAB201 & " d, " & T_LAB031 & " f, " & T_LAB001 & " h "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
                                    
        sWhereClause(1) = "Where a.workarea = '" & pWorkArea & "' " & _
                                    "and    a.accdt = '" & pAccDt & "' " & _
                                    "and    a.accseq = " & pAccSeq & " " & _
                                    "and    d.workarea = a.workarea   and    d.accdt = a.accdt   and    d.accseq = a.accseq   and    d.multifg = ''  " & _
                                    "and    h.testcd = a.ordcd " & _
                                    "and    h.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = h.testcd) " & _
                                    "and    h.panelfg = '" & PN_Group & "' " & _
                                    "and    f.cdindex = '" & CD1_Panel & "' " & _
                                    "and    f.cdval1 = h.testcd " & _
                                    "and    b.testcd = f.field1 " & _
                                    "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                                    "and    b.panelfg = '" & PN_Normal & "' " & _
                                    "and    c.testcd = b.testcd   and    c.spccd = d.spccd " & _
                                    "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                                    "and    e.ptid = a.ptid   and    e.testcd = f.field1 " & _
                                    "and    e.vfydt = (select max(vfydt) from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.field1 and vfydt <> '' and vfydt < d.rcvdt) " & _
                                    "and    exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.field1 and vfydt <> '' and vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "and    not exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.field1 and vfydt <> '' and vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " Union All " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Group2":   '그룹처방이면서 상세항목
        sSelectClause(1) = "Select a.ptid, a.orddt, a.ordno, a.ordseq, f.field1 as Testcd1, g.cdval2 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                                    "          b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                                    "          d.spccd, d.multifg, d.testdiv, 'Y' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "          e.rstcd as LastRst, e.vfydt as LastVfyDt, e.vfytm as LastVfyTm, e.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "
                                    
        sFromClause(1) = "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB031 & " g, " & _
                                                T_LAB201 & " d, " & T_LAB031 & " f, " & T_LAB001 & " h, " & T_LAB001 & " i "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " e "
        sFromClause(3) = sFromClause(1)
                                    
        sWhereClause(1) = "Where a.workarea = '" & pWorkArea & "' " & _
                                    "and    a.accdt = '" & pAccDt & "' " & _
                                    "and    a.accseq = " & pAccSeq & " " & _
                                    "and    d.workarea = a.workarea   and    d.accdt = a.accdt   and    d.accseq = a.accseq   and    d.multifg = ''  " & _
                                    "and    i.testcd = a.ordcd " & _
                                    "and    i.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = i.testcd) " & _
                                    "and    i.panelfg = '" & PN_Group & "' " & _
                                    "and    f.cdindex = '" & CD1_Panel & "'   and    f.cdval1 = i.testcd " & _
                                    "and    h.testcd = f.field1 " & _
                                    "and    h.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = h.testcd) " & _
                                    "and    h.panelfg = '" & PN_Detail & "' " & _
                                    "and    g.cdindex = '" & CD1_Detail & "'  and    g.cdval1 = h.testcd " & _
                                    "and    b.testcd = g.cdval2 " & _
                                    "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                                    "and    c.testcd = g.cdval2   and    d.spccd = a.spccd " & _
                                    "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                                    "and    e.ptid = a.ptid   and    e.testcd = g.cdval2 " & _
                                    "and    e.vfydt = (select max(vfydt) from " & T_LAB302 & " where ptid = a.ptid  and  testcd = g.cdval2 and vfydt <> '' and vfydt < d.rcvdt) " & _
                                    "and    exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = g.cdval2 and vfydt <> '' and vfydt < d.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "and    not exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = g.cdval2 and vfydt <> '' and vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " Union All " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    Case "Group3":    '그룹이면서 연속검사(복수검체)
        sSelectClause(1) = "Select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as Testcd1, f.field1 as TestCd2, a.spccd as OrdSpc, a.attrcd, a.statfg, " & _
                                    "          b.panelfg, b.rsttype, b.rstdiv, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                                    "          e.spccd, e.multifg, e.testdiv, 'N' as DetailFg, "
        sSelectClause(2) = sSelectClause(1) & "          g.rstcd as LastRst, g.vfydt as LastVfyDt, g.vfytm as LastVfyTm, g.vfyid as LastVfyId "
        sSelectClause(3) = sSelectClause(1) & "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, null as LastVfyId "

        sFromClause(1) = "From " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & _
                                                T_LAB203 & " d, " & T_LAB201 & " e, " & T_LAB031 & " f "
        sFromClause(2) = sFromClause(1) & ", " & T_LAB302 & " g "
        sFromClause(3) = sFromClause(1)
        
        sWhereClause(1) = "Where d.workarea = '" & pWorkArea & "' " & _
                                    "and    d.accdt = '" & pAccDt & "' " & _
                                    "and    d.accseq = " & pAccSeq & " " & _
                                    "and    e.workarea = a.workarea " & _
                                    "and    e.accdt = a.accdt " & _
                                    "and    e.accseq = a.accseq  " & _
                                    "and    e.multifg <> ''  " & _
                                    "and    a.ptid = d.ptid " & _
                                    "and    a.orddt = d.orddt " & _
                                    "and    a.ordno = d.ordno " & _
                                    "and    a.ordseq = d.ordseq " & _
                                    "and    f.cdindex = '" & CD1_Panel & "' " & _
                                    "and    f.cdval1 = a.ordcd    and    f.cdval2 = e.multifg " & _
                                    "and    b.testcd = f.field1 " & _
                                    "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                                    "and    c.testcd = f.field1   and    c.spccd = e.spccd " & _
                                    "and    c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) "
        sWhereClause(2) = sWhereClause(1) & _
                                    "and    g.ptid = a.ptid   and    g.testcd = f.field1 " & _
                                    "and    g.vfydt = (select max(vfydt) from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.field1 and vfydt <> '' and vfydt < e.rcvdt) " & _
                                    "and    exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.field1 and vfydt <> '' and vfydt < e.rcvdt) "
        sWhereClause(3) = sWhereClause(1) '& _
                                    "and    not exists (select * from " & T_LAB302 & " where ptid = a.ptid  and  testcd = f.field1 and vfydt <> '' and vfydt < d.rcvdt) "
                                    
        'SqlOrderItems = sSelectClause(2) & sFromClause(2) & sWhereClause(2) & " Union All " & sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        SqlOrderItems1 = sSelectClause(3) & sFromClause(3) & sWhereClause(3)
        
    End Select
        
    SqlOrderItems1 = SqlOrderItems1 & " Order by Testcd1, TestCd2 "
                               
 End Function

'%  23. 최근결과 검색 (환자ID - 검사항목)
'%      - Calling from [clsAccession]
Public Function SqlLastResult(ByVal pPID As String, ByVal pTCd As String) As String

   SqlLastResult = "Select rstcd, vfydt, vfytm, vfyid From " & T_LAB302 & " " & _
                          "Where ptid = " & pPID & " " & _
                          "and    testcd = '" & pTCd & "' " & _
                          "and    vfydt is not null " & _
                          "and    vfydt <> '' " & _
                          "Order By vfydt desc, vfytm desc"

End Function

'%  24. 결과내역(LAB302) Insert
'%      - Calling from [clsAccession]
Public Function SqlCreateResult(ByVal TestDiv As String) As String
   Select Case TestDiv
   Case "0":   '일반검사
   SqlCreateResult = "Insert into " & T_LAB302 & "(workarea, accdt, accseq, testcd, rstunit, " & _
                             "                           spccd, statfg, lastrst, lastvfydt, lastvfytm, lastvfyid, " & _
                             "                           attrcd, mfyfg, grpfg, txtfg, rsttype, rstdiv, detailfg, ptid, orddt, ordno, ordseq) " & _
                             "Values               (:WorkArea,:AccDt,:AccSeq,:testcd,:RstUnit, " & _
                             "                           :SpcCd,:StatFg,:LastRst,:LastVfyDt,:LastVfyTm,:LastVfyId, " & _
                             "                           :AttrCd,:MfyFg,:GrpFg,:TxtFg,:RstType,:RstDiv,:DetailFg,:PtId,:OrdDt,:OrdNo,:OrdSeq)"
   Case "1":   '기타검사
   SqlCreateResult = "Insert into " & T_LAB351 & "(WORKAREA,ACCDT,ACCSEQ,TESTCD,StsCd,RstType,ValFg,TxtFg, " & _
                             "                          MfyFg,VfyDt,VfyTm,VfyId,PtId,OrdDt,OrdNo,OrdSeq) " & _
                             "Values               (:WORKAREA,:ACCDT,:ACCSEQ,:TESTCD,:StsCd,:RstType,:ValFg,:TxtFg, " & _
                             "                          :MfyFg,:VfyDt,:VfyTm,:VfyId,:PtId,:OrdDt,:OrdNo,:OrdSeq)"
   Case "2":   '미생물검사
   SqlCreateResult = "Insert into " & T_LAB404 & "(WORKAREA,ACCDT,ACCSEQ,TESTCD,SenFg,StsCd,RstType," & _
                             "                          MfyFg,RstDiv,DetailFg,LastRst,LastVfyDt,LastVfyTm, " & _
                             "                          LastVfyId,PtId,OrdDt,OrdNo,OrdSeq) " & _
                             "Values               (:WORKAREA,:ACCDT,:ACCSEQ,:TESTCD,:SenFg,:StsCd,:RstType," & _
                             "                          :MfyFg,:RstDiv,:DetailFg,:LastRst,:LastVfyDt,:LastVfyTm, " & _
                             "                          :LastVfyId,:PtId,:OrdDt,:OrdNo,:OrdSeq)"
   End Select
End Function


'%  25. 외부의뢰내역(LAB205) Insert
'%      - Calling from [clsAccession]
Public Function SqlCreateOutLab() As String
   SqlCreateOutLab = "Insert into " & T_LAB205 & "(WorkArea,AccDt,AccSeq,testcd,OutLabCd," & _
                              "                           StsCd,RcvDt,SendDt,ChargeDt,SendId) " & _
                              "Values ( :WorkArea,:AccDt,:AccSeq,:TestCd,:OutLabCd," & _
                              "                           :StsCd,:RcvDt,:SendDt,:ChargeDt,:SendId)"
End Function

'%  26. 채혈접수내역 Update
'%      - Calling from [clsAccession]
Public Function SqlColUpdate(ByVal StsCd As String, ByVal ReqTotCnt As Integer, ByVal AltTotCnt As Integer, _
                                           ByVal RcvDt As String, ByVal RcvTm As String, ByVal RcvId As String) As String
   SqlColUpdate = "Update " & T_LAB201 & " " & _
                         "Set     StsCd = '" & StsCd & "', ReqTotCnt = " & ReqTotCnt & ", AltTotCnt = " & AltTotCnt & ", " & _
                         "          RcvDt = '" & RcvDt & "', RcvTm = '" & RcvTm & "', RcvId = '" & RcvId & "' " & _
                         "Where WorkArea = :WorkArea " & _
                         "and     AccDt = :AccDt " & _
                         "and     AccSeq = :AccSeq "
                         
End Function


'%  27. 처방Body내역 Update
'%      - Calling from [clsAccession]
Public Function SqlStatusUpdate(ByVal PtId As String, ByVal OrdDt As String, ByVal OrdNo As Integer, _
                                               ByVal OrdSeq As Integer, ByVal StsCd As String, Optional ByVal DoneFg As Variant) As String
   Dim tmpStr As String
   Dim tmpStr1 As String
   
   If StsCd = STS_Access Then  '접수
      tmpStr = "rcvdt = convert(char(8),getdate(),112), " & _
                  "rcvtm = substring(convert(char(8),getdate(),108),1,2)+" & _
                  "              substring(convert(char(8),getdate(),108),4,2)+" & _
                  "              substring(convert(char(8),getdate(),108),7,2)"
   Else
      tmpStr = ""
   End If
   
   If IsMissing(DoneFg) Then
      tmpStr1 = ""
   Else
      tmpStr1 = "donefg = '" & DoneFg & "', "
   End If
   SqlStatusUpdate = "Update " & T_LAB102 & " " & _
                             "Set       stscd = '" & StsCd & "', " & tmpStr1 & tmpStr & _
                             "Where   ptid = " & PtId & " " & _
                             "and      orddt = '" & OrdDt & "' " & _
                             "and      ordno = " & OrdNo & " " & _
                             "and      ordseq = " & OrdSeq

End Function

'%  27-1. 처방Body내역 Update
'%      - Calling from [frm156Referral]
Public Function SqlStatusUpdate1(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Integer, _
                                               ByVal TestCd As String, ByVal StsCd As String) As String
   
   SqlStatusUpdate1 = "Update " & T_LAB102 & " " & _
                             "Set       stscd = '" & StsCd & "' " & _
                             "Where   workarea = '" & WorkArea & "' " & _
                             "and      accdt = '" & AccDt & "' " & _
                             "and      accseq = " & AccSeq & " " & _
                             "and      ordcd = '" & TestCd & "' "

End Function

'%  27-2. 처방Header내역 Update (DoneFg)
'%      - Calling from [clsAccession]
Public Function SqlHeaderUpdate(ByVal PtId As String, ByVal OrdDt As String, ByVal OrdNo As Integer, ByVal DoneFg As String, ByVal flag As Integer) As String
   If flag = 1 Then
   SqlHeaderUpdate = "Select * from " & T_LAB102 & " " & _
                               "Where   ptid = " & PtId & " " & _
                               "and      orddt = '" & OrdDt & "' " & _
                               "and      ordno = " & OrdNo & " " & _
                               "and      donefg < '" & DoneFg & "' "
   Else
   SqlHeaderUpdate = "Update " & T_LAB101 & " " & _
                               "Set       donefg = '" & DoneFg & "' " & _
                               "Where   ptid = " & PtId & " " & _
                               "and      orddt = '" & OrdDt & "' " & _
                               "and      ordno = " & OrdNo
   End If

End Function


'%  28. 연속검사인 경우 채혈접수내역의 검체Seq를 Key로 해당항목 검색, 또는 전체 검색
'%       - Calling from [clsAccession]
'%       -CdIndex = 'C101' : 그룹처방(Field1), 'C103' : 상세항목(CdVal2)
Public Function SqlPanelTest(ByVal PtId As String, ByVal CdIndex As String, ByVal TestCd As String, _
                                           ByVal SpcCd As String, Optional ByVal MultiFg As Variant) As String
   Dim tmpStr As String
   Dim TestCdField As String
   
   If Not IsMissing(MultiFg) Then tmpStr = "and     a.cdval2 = '" & MultiFg & "' "
   If CdIndex = CD1_Panel Then
      TestCdField = "a.field1"
   ElseIf CdIndex = CD1_Detail Then
      TestCdField = "a.cdval2"
   End If

   SqlPanelTest = "Select " & TestCdField & " as testcd, b.rstdiv, b.rsttype, b.panelfg, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                         "          d.rstcd as LastRst, d.vfydt as LastVfyDt, d.vfytm as LastVfyTm, d.vfyid as LastVfyId " & _
                         "From   " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c, " & T_LAB302 & " d " & _
                         "where a.cdindex = '" & CdIndex & "' " & _
                         "and     a.cdval1 = '" & TestCd & "' " & tmpStr & _
                         "and     b.testcd = " & TestCdField & " " & _
                         "and     b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                         "and     c.testcd = " & TestCdField & "  " & _
                         "and     c.spccd = '" & SpcCd & "' " & _
                         "and     c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) " & _
                         "and     d.ptid = " & PtId & "   and    d.testcd = " & TestCdField & " " & _
                         "and     d.vfydt+d.vfytm = (select max(vfydt+vfytm) from " & T_LAB302 & " where ptid = d.ptid  and  testcd = d.testcd and vfydt <> '' ) " & _
                         "and     exists (select * from " & T_LAB302 & " where ptid = " & PtId & "  and  testcd = " & TestCdField & " and vfydt <> '' ) "
   SqlPanelTest = SqlPanelTest & " Union All " & _
                         "Select " & TestCdField & " as testcd, b.rstdiv, b.rsttype, b.panelfg, b.outlabcd, c.labelcnt,  c.statflags, c.rstunit, " & _
                         "          '' as LastRst, '' as LastVfyDt, '' as LastVfyTm, 0 as LastVfyId " & _
                         "From   " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB004 & " c " & _
                         "where a.cdindex = '" & CdIndex & "' " & _
                         "and     a.cdval1 = '" & TestCd & "' " & tmpStr & _
                         "and     b.testcd = " & TestCdField & " " & _
                         "and     b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                         "and     c.testcd = " & TestCdField & "  " & _
                         "and     c.spccd = '" & SpcCd & "' " & _
                         "and     c.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = c.testcd and spccd = c.spccd) " & _
                         "and     not exists (select * from " & T_LAB302 & " where ptid = " & PtId & "  and  testcd = " & TestCdField & " and vfydt <> '' ) "
                         
End Function

'%  29. 병동 일괄 채혈을 위해 병동별로 처방내역을 검색한다. (날짜만 비교, 12.29)
'%       - Calling from [frm151Round]
Public Function SqlWardOrder(ByVal pReqDt As String, ByVal pReqTm As String, ByVal pWardId As String) As String
   
    SqlWardOrder = "Select b.wardid as wardid, a." & F_PTID & " as ptid, a." & F_PTNM & " as ptnm, " _
                 & "       " & F_SEX2("a") & " as Sex, c.statfg, c.spccd, " _
                 & "       " & F_DOB2("a") & " as Dob, b.bedindt, b.deptcd, " _
                 & "       b.orddoct, b.majdoct, b.roomid, '' as bedid, b.hosilid, b.orddiv " _
                 & "From " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " _
                 & "Where  b.wardid = " & DBS(pWardId) & " " _
                 & "and    b.donefg = '0' " _
                 & "and    " & DBW("b.reqdt <= ", pReqDt) & " " _
                 & "and    b.bussdiv = " & DBS(CS_BussIn) & " " _
                 & "and    a." & F_PTID & " = b.ptid " _
                 & "and    c.ptid = b.ptid    and   c.orddt = b.orddt  and   c.ordno = b.ordno  " _
                 & "and    c.stscd = '0'  and   (c.dcfg = '' or c.dcfg is null) " _
                 & "Order By wardid, ptid, c.statfg, c.spccd "

    '나중에 Union 걸자...(해부병리 Database가 분리되면...)
'    SqlWardOrder = SqlWardOrder & " Union " _
'                 & "Select b.wardid as wardid, a." & F_PTID & " as ptid, a." & F_PTNM & " as ptnm, " _
'                 & "       substring(a." & F_SSN & ",9,1) as Sex, c.statfg, c.spccd, " _
'                 & "       substring(a." & F_SSN & ",1,8) as Dob, b.bedindt, b.deptcd, " _
'                 & "       b.orddoct, b.majdoct, b.roomid, '' as bedid, b.hosilid, b.orddiv " _
'                 & "From " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " _
'                 & "Where  b.wardid = " & dbs(pWardId) & "' " _
'                 & "and    b.donefg = '0' " _
'                 & "and    b.reqdt <= " & dbs(ReqDt) & " " _
'                 & "and    b.bussdiv = " & dbs(CS_BussIn) & " " _
'                 & "and    a.ptnt_no = b.ptid " _
'                 & "and    c.ptid = b.ptid    and   c.orddt = b.orddt  and   c.ordno = b.ordno  " _
'                 & "and    c.stscd = '0'  and   c.dcfg = '' " _
'                 & "Order By wardid, ptid, c.statfg, c.spccd "


End Function

'%  30. 해당 환자의 처방내역을 조회한다.
'%       - Calling from [frm151Round] : 병동 일괄채혈, [frm153SendPt] : 외래채혈접수, [frm154NulCol] : 간호사채혈
'%       - 검사구분, WorkArea, 검체코드, 보관구분, 응급여부, 희망채혈일시, 복수검체여부, 검체군, 처방일, 처방번호, 처방Seq, 처방코드
'%       - Lab101 : 처방Header내역
'%       - Lab102 : 처방Body내역
'%       - Lab001 : 검사항목 내역
'%       - Lab004 : 지정검체내역(라벨출력장수)
'%       - Lab032 : 검체정보(C215 - 복수검체여부(Field1),검체군(Field2))
Public Function SqlReadOrder(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                            Optional ByVal StatFg As String = "", Optional ByVal Bussdiv As String = "", _
                            Optional ByVal ReceptNo As String = "", Optional ByVal OrdDiv As String = "") As String

   Dim tmpStr As String
   Dim strSql(3) As String

   tmpStr = ""
   If StatFg <> "" Then tmpStr = tmpStr & "and    b.statfg = '" & StatFg & "' "         '응급여부
   If Bussdiv <> "" Then tmpStr = tmpStr & "and   a.bussdiv = '" & Bussdiv & "' "      '외래병동구분
   If ReceptNo <> "" Then tmpStr = tmpStr & "and  a.receptno = " & ReceptNo & " "     '영수증번호(외래환자)
   
   If Bussdiv = "1" Then '외래
      SqlReadOrder = "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                    "       d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, " & _
                    "       b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                    "       e.empl_nm as DoctNm, a.deptcd,  f.statflags, f.labelcnt, '' as ErrFg " & _
                    "From   " & T_LAB101 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & _
                                T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                    "Where a.ptid = " & PtId & " " & _
                    "and    a.donefg = '0' " & _
                    "and    a.orddt = '" & ReqDt & "' " & _
                    "and    b.ptid = a.ptid " & _
                    "and    b.orddt = a.orddt " & _
                    "and    b.ordno = a.ordno " & _
                    "and    b.donefg = '0' " & _
                    "and    b.dcfg = '' " & _
                    tmpStr & _
                    "and    c.testcd = b.ordcd " & _
                    "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    "and    d.cdindex = '" & CD2_Specimen & "' " & _
                    "and    d.cdval1 = b.spccd " & _
                    "and    e.empl_no =* a.orddoct " & _
                    "and    f.testcd = b.ordcd and f.spccd = b.spccd " & _
                    "and    f.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = f.testcd   and     spccd = f.spccd ) "
      SqlReadOrder = SqlReadOrder & " UNION ALL " & _
                    "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, '' as spccd, b.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                    "       '' as SpcNm, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, " & _
                    "       b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                    "       e.empl_nm as DoctNm, a.deptcd,  '' as statflags, 0 as labelcnt, 'Y' as ErrFg " & _
                    "From   " & T_LAB101 & " a, " & T_LAB102 & " b,  " & _
                                T_LAB001 & " c, " & T_HIS005 & " e " & _
                    "Where a.ptid = " & PtId & " " & _
                    "and    a.donefg = '0' " & _
                    "and    a.orddt = '" & ReqDt & "' " & _
                    "and    b.ptid = a.ptid " & _
                    "and    b.orddt = a.orddt " & _
                    "and    b.ordno = a.ordno " & _
                    "and    b.donefg = '0' " & _
                    "and    b.dcfg = '' " & _
                    tmpStr & _
                    "and    c.testcd = b.ordcd " & _
                    "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    "and    e.empl_no =* a.orddoct " & _
                    "and    NOT EXISTS ( select * from " & T_LAB004 & " where testcd = b.ordcd and spccd = b.spccd " & _
                    "                    and applydt <= '" & Format(Now, CS_DateDbFormat) & "' ) " & _
                    "Order By ColTm, orddt, ordno, ordseq "
   Else
                             ' ## 나중에 HIS002가 사용가능하면 Query문 수정 ##
                             '"          g.adm_ymd as BedInDt, g.ward_cd as WardId, g.room_cd as RoomId "
                             '               T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f, " & T_HIS002 & " g "
                             '"and    g.hosp_gb = '" & HosptGb & "' " & _
                             '"and    g.recept_no =* a.receptno "
      '임상병리
      strSql(1) = "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                    "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                    "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                    "        convert(int,f.labelcnt) as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                    "        '' as fzfg " & _
                    " From " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
                               T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                    " Where  a.ptid = " & PtId & " " & _
                    " and    a.donefg = '0' " & _
                    " and    a.reqdt <= '" & ReqDt & "' " & _
                    " and    a.orddiv = 'L' " & _
                    " and    b.ptid = a.ptid " & _
                    " and    b.orddt = a.orddt " & _
                    " and    b.ordno = a.ordno " & _
                    " and    b.donefg = '0' " & tmpStr & _
                    " and   (b.dcfg = '' or b.dcfg is null) " & _
                    " and    c.testcd = b.ordcd " & _
                    " and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    " and    d.cdindex = '" & CD2_Specimen & "' " & _
                    " and    d.cdval1 =* b.spccd " & _
                    " and    e." & F_DOCTID & " =* a.orddoct " & _
                    " and    f.testcd = b.ordcd and f.spccd = b.spccd " & _
                    " and    f.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = f.testcd  and     spccd = f.spccd ) "
      '혈액은행 : Phersis검사는 제외...
      strSql(2) = "Select c.testnm, c.abbrnm5, '3' as testdiv, 'XM' as workarea, b.spccd, '' as storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                    "        '혈액' as SpcNm, '혈액' as SpcNm5, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                    "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  '' as statflags, " & _
                    "        1 as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                    "        '' as fzfg " & _
                    " From " & T_LAB101 & " a, " & T_LAB102 & " b, " & _
                               T_BBS001 & " c, " & T_HIS005 & " e " & _
                    " Where  a.ptid = " & PtId & " " & _
                    " and    a.donefg = '0' " & _
                    "  " & _
                    " and    a.orddiv = 'B' " & _
                    " and    b.ptid = a.ptid " & _
                    " and    b.orddt = a.orddt " & _
                    " and    b.ordno = a.ordno " & _
                    " and    b.donefg = '0' " & tmpStr & _
                    " and   (b.dcfg = '' or b.dcfg is null) " & _
                    " and    c.testcd = b.ordcd " & _
                    " and    c.testdiv <> '3'  " & _
                    " and    c.applydt = (select max(applydt) from " & T_BBS001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    " and    e." & F_DOCTID & " =* a.orddoct "
                    'and    a.reqdt <= '" & ReqDt & "'  2001.2.13 김미경 막음 : 처방상태면 무조건 읽어옴.
    '해부병리
    strSql(3) = "select c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, '' as storecd, b.statfg, a.reqdt " & func_concat & " ' ' " & func_concat & " a.reqtm as ColTm, " _
                   & "       d.spcnm, d.spcnm5, '' as multifg, '' as spcgrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " _
                   & "       a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd, '' as statflags, " _
                   & "       d.baccdcnt as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid, " _
                   & "       c.fzfg " _
                   & "from " & T_LAB101 & " a," & T_LAB102 & " b," & T_APS001 & " c," _
                             & T_APS002 & " d," & T_HIS005 & " e " _
                   & " where a.ptid = " & DBV("ptid", PtId) & strCondition _
                   & "   and a.donefg = " & DBS(DoneFg.DoneFg_Order) _
                   & "   and a.bussdiv = " & DBS(Bussdiv) _
                   & "   and a.orddiv = " & DBS(APS_ORDDIV) _
                   & "   and b.ptid = a.ptid " _
                   & "   and b.orddt = a.orddt " _
                   & "   and b.ordno = a.ordno " _
                   & "   and b.donefg = " & DBS(DoneFg.DoneFg_Order) _
                   & "   and (b.dcfg = '' or b.dcfg is null) " _
                   & "   and c.testcd = b.ordcd " _
                   & "   and d.spccd = b.spccd " _
                   & "   and e." & F_DOCTID & " =* a.orddoct  "
                   
    Select Case OrdDiv
    Case "A": SqlReadOrder = strSql(3)
    Case "B": SqlReadOrder = strSql(2)
    Case "L": SqlReadOrder = strSql(1)
    Case "": SqlReadOrder = strSql(3) & " union all " & strSql(2) & " union all " & strSql(1)
    End Select
    
    SqlReadOrder = SqlReadOrder & " Order By ColTm, orddt, ordno, ordcd "                          '<< D/C 처방 제외 >>
      
      'Union
'                             "and    a.reqtm <= '" & ReqTm & "' " & _   '시간 제외
'      SqlReadOrder = SqlReadOrder & "Union All " & _
'                             "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, b.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
'                             "          d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
'                             "          a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, e.empl_nm as DoctNm, a.deptcd,  f.statflags, f.labelcnt, " & _
'                             "          a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid " & _
'                             "From  " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
'                                        T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
'                             "Where  a.ptid = " & PtId & " " & _
'                             "and    a.donefg = '0' " & _
'                             "and    a.reqdt < '" & ReqDt & "' " & _
'                             "and    b.ptid = a.ptid " & _
'                             "and    b.orddt = a.orddt " & _
'                             "and    b.ordno = a.ordno " & _
'                             "and    b.donefg = '0' " & _
'                             "and    b.dcfg = '' " & tmpStr & _
'                             "and    c.testcd = b.ordcd " & _
'                             "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
'                             "and    d.cdindex = '" & CD2_Specimen & "' " & _
'                             "and    d.cdval1 = b.spccd " & _
'                             "and    e.empl_no =* a.orddoct " & _
'                             "and    f.testcd = b.ordcd and f.spccd = b.spccd " & _
'                             "and    f.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = f.testcd  and     spccd = f.spccd ) " & _
                             "Order By ColTm, orddt, ordno, ordseq "
   End If

End Function
   
Public Function SqlReadOrderER(ByVal PtId As String, ByVal ReqDt As String, ByVal ReceptNo As String) As String

   Dim tmpStr As String
   
  SqlReadOrderER = "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, b.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                           "          d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                           "          a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, e.empl_nm as DoctNm, a.deptcd,  f.statflags, f.labelcnt " & _
                           "From   " & T_LAB101 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                           "Where a.ptid = " & PtId & " " & _
                           "and    a.donefg = '0' " & _
                           "and    a.orddt = '" & ReqDt & "' " & _
                           "and    a.bussdiv = '" & CS_BussOut & "' " & _
                           "and    a.deptcd = 'ER' " & _
                           "and    a.receptno = " & ReceptNo & " " & _
                           "and    b.ptid = a.ptid " & _
                           "and    b.orddt = a.orddt " & _
                           "and    b.ordno = a.ordno " & _
                           "and    b.donefg = '0' " & _
                           "and    b.dcfg = '' " & _
                           "and    c.testcd = b.ordcd " & _
                           "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                           "and    d.cdindex = '" & CD2_Specimen & "' " & _
                           "and    d.cdval1 = b.spccd " & _
                           "and    e.empl_no =* a.orddoct " & _
                           "and    f.testcd = b.ordcd and f.spccd = b.spccd " & _
                           "and    f.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = f.testcd   and     spccd = f.spccd ) " & _
                           "Order By ColTm, orddt, ordno, ordseq "
   
End Function

Public Function SqlGsBatchOrder(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                           Optional ByVal StatFg As String = "", Optional ByVal Bussdiv As String = "", Optional ByVal ReceptNo As String = "") As String

   Dim tmpStr As String
   
   '여성클리닉 환자들의 Gram Stain 일괄검사처리를 위한 Query 문
   '--> 검사항목-'B400', 검체-'75', 진료과-'OB'
   tmpStr = ""
   If StatFg <> "" Then tmpStr = tmpStr & "and    b.statfg = '" & StatFg & "' "   '응급여부
   
      SqlGsBatchOrder = "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, b.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                    "       d.field3 as SpcNm, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, " & _
                    "       b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                    "       e.empl_nm as DoctNm, a.deptcd,  f.statflags, f.labelcnt, '' as ErrFg " & _
                    "From   " & T_LAB101 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & _
                                T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                    "Where a.ptid = " & PtId & " " & _
                    "and    a.donefg = '0'    and  a.bussdiv = '" & CS_BussOut & "' " & _
                    "and    a.orddt = '" & ReqDt & "'  and  a.deptcd in ('OB','CO') " & _
                    "and    b.ptid = a.ptid   and  b.orddt = a.orddt   and    b.ordno = a.ordno " & _
                    "and    b.ordcd = 'B400'  and  b.donefg = '0'  and  b.spccd = '75' " & _
                    "and    b.dcfg = ''       " & _
                    tmpStr & _
                    "and    c.testcd = b.ordcd " & _
                    "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    "and    d.cdindex = '" & CD2_Specimen & "' " & _
                    "and    d.cdval1 = b.spccd " & _
                    "and    e.empl_no =* a.orddoct " & _
                    "and    f.testcd = b.ordcd and f.spccd = b.spccd " & _
                    "and    f.applydt = (select max(applydt) from " & T_LAB004 & " where testcd = f.testcd   and     spccd = f.spccd ) "
      SqlGsBatchOrder = SqlGsBatchOrder & " UNION ALL " & _
                    "Select c.testnm, c.abbrnm5, c.testdiv, c.workarea, '' as spccd, b.storecd, b.statfg, a.reqdt+' '+a.reqtm as ColTm, " & _
                    "       '' as SpcNm, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, " & _
                    "       b.ordcd, b.mesg, a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, " & _
                    "       e.empl_nm as DoctNm, a.deptcd,  '' as statflags, 0 as labelcnt, 'Y' as ErrFg " & _
                    "From   " & T_LAB101 & " a, " & T_LAB102 & " b,  " & _
                                T_LAB001 & " c, " & T_HIS005 & " e " & _
                    "Where a.ptid = " & PtId & " " & _
                    "and    a.donefg = '0'    and  a.bussdiv = '" & CS_BussOut & "'" & _
                    "and    a.orddt = '" & ReqDt & "'  and  a.deptcd in ('OB','CO') " & _
                    "and    b.ptid = a.ptid   and  b.orddt = a.orddt   and    b.ordno = a.ordno " & _
                    "and    b.ordcd = 'B400'  and  b.donefg = '0'  and  b.spccd = '75' " & _
                    "and    b.dcfg = ''        " & _
                    tmpStr & _
                    "and    c.testcd = b.ordcd " & _
                    "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                    "and    e.empl_no =* a.orddoct " & _
                    "and    NOT EXISTS ( select * from " & T_LAB004 & " where testcd = b.ordcd and spccd = b.spccd " & _
                    "                    and applydt <= '" & Format(Now, CS_DateDbFormat) & "' ) " & _
                    "Order By ColTm, orddt, ordno, ordseq "
End Function

'%  31. 환자Id 또는 환자명을 입력받아 검색한다.
Public Function SqlPtntSearch(ByVal SearchKey As String, ByVal SorkKey As String) As String
   'SqlPtntSearch = "Select ptid, ptnm from " & T_HIS001 & " " & _
                          "Where ptid like '" & SearchKey & "%' or ptnm like '" & SearchKey & "%' " & _
                          "Order by " & SorkKey
   If IsNumeric(SearchKey) Then
      SqlPtntSearch = "Select ptnt_no as ptid, ptnt_nm as ptnm, " & _
                             "          substring(ptnt_prsn_no,3,6)+'-'+substring(ptnt_prsn_no,9,7) as SSN from " & T_HIS001 & " " & _
                             "Where ptnt_no between " & SearchKey & " and " & (SearchKey + 1000) & " " & _
                             "Order by " & SorkKey
   Else
      SqlPtntSearch = "Select ptnt_no as ptid, ptnt_nm as ptnm, " & _
                             "          substring(ptnt_prsn_no,3,6)+'-'+substring(ptnt_prsn_no,9,7) as SSN from " & T_HIS001 & " " & _
                             "Where ptnt_nm like '" & SearchKey & "%' " & _
                             "Order by " & SorkKey & ", SSN"
   End If
End Function


'%  32. 환자Id, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
'%       - Lab101 : 처방Header내역
'%       - Lab102 : 처방Body내역
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlQueryOrders(ByVal PtId As String, ByVal QueryKey As String, ByVal FromDate As String, ByVal ToDate As String) As String
   SqlQueryOrders = SqlQueryOrders & _
                    "Select substring(a.orddt,3,2)+'-'+substring(a.orddt,5,2)+'-'+substring(a.orddt,7,2) as OrdDate, " & _
                    "       a.ordno, a.ordseq, a.ordcd, c.testnm, a.spccd, a.statfg, " & _
                    "       f.empl_nm as DoctNm, a.rcvdt, substring(a.rcvtm,1,2)+':'+substring(a.rcvtm,3,2) as RcvTm, " & _
                    "       a.stscd, a.dcfg, a.mesg, d.field3 as SpcNm, d.field1 as MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
                    "       substring(b.ordtm,1,2)+':'+substring(b.ordtm,3,2) as OrdTm, " & _
                    "       substring(a.examdt,1,4)+'-'+substring(a.examdt,5,2)+'-'+substring(a.examdt,7,2) as ExamDt, " & _
                    "       substring(a.examtm,1,2)+':'+substring(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, " & _
                    "       a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
                    "       '' spcyy, 0 spcno, 0 unitqty, '' reqdt, '' reqtm " & _
                    "From " & T_LAB102 & " a, " & T_LAB101 & " b, " & T_LAB001 & " c, " & _
                              T_LAB032 & " d, " & T_HIS005 & " f, " & T_LAB015 & " g " & _
                    "Where  a.ptid = " & PtId & " " & _
                    "and    a." & QueryKey & " >= '" & FromDate & "' " & _
                    "and    a." & QueryKey & " <= '" & ToDate & "' " & _
                    "and    a.ptid    =  b.ptid " & _
                    "and    a.orddt   =  b.orddt " & _
                    "and    a.ordno   =  b.ordno " & _
                    "and    b.orddiv  =  'L' " & _
                    "and    c.testcd  =  a.ordcd " & _
                    "and    c.applydt =  (select max(applydt) from " & T_LAB001 & _
                                       " where testcd = c.testcd and applydt <= a.orddt) " & _
                    "and    d.cdindex = '" & CD2_Specimen & "' and    d.cdval1 = a.spccd " & _
                    "and    f.empl_no =* b.orddoct   " & _
                    "and    g.empid   =* a.examdoct   "
                       
   SqlQueryOrders = SqlQueryOrders & "union all " & _
                    "Select substring(a.orddt,3,2)+'-'+substring(a.orddt,5,2)+'-'+substring(a.orddt,7,2) as OrdDate, " & _
                    "       a.ordno, a.ordseq, a.ordcd, c.testnm, a.spccd, a.statfg, " & _
                    "       f.empl_nm as DoctNm, a.rcvdt, substring(a.rcvtm,1,2)+':'+substring(a.rcvtm,3,2) as RcvTm, " & _
                    "       a.stscd, a.dcfg, a.mesg, d.spcnm, '' MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
                    "       substring(b.ordtm,1,2)+':'+substring(b.ordtm,3,2) as OrdTm, " & _
                    "       substring(a.examdt,1,4)+'-'+substring(a.examdt,5,2)+'-'+substring(a.examdt,7,2) as ExamDt, " & _
                    "       substring(a.examtm,1,2)+':'+substring(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, " & _
                    "       a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '*' as Verified, " & _
                    "       h.spcyy, h.spcno, 0 unitqty, '' reqdt, '' reqtm " & _
                    "From " & T_LAB102 & " a, " & T_LAB101 & " b, " & T_APS001 & " c, " & _
                              T_APS002 & " d, " & T_HIS005 & " f, " & T_LAB015 & " g, " & T_APS103 & " h " & _
                    "Where  a.ptid = " & PtId & " " & _
                    "and    a." & QueryKey & " >= '" & FromDate & "' " & _
                    "and    a." & QueryKey & " <= '" & ToDate & "' " & _
                    "and    a.ptid = b.ptid   and   a.orddt = b.orddt   and   a.ordno = b.ordno  and b.orddiv = 'A' " & _
                    "and    c.testcd = a.ordcd " & _
                    "and    d.spccd  = a.spccd " & _
                    "and    f.empl_no =* b.orddoct   " & _
                    "and    g.empid =* a.examdoct   " & _
                    "and    h.ptid=*a.ptid   and h.orddt=*a.orddt  and  h.ordno=*a.ordno   and   h.ordseq=*a.ordseq "
                    
   SqlQueryOrders = SqlQueryOrders & "union all " & _
                    "Select substring(a.orddt,3,2)+'-'+substring(a.orddt,5,2)+'-'+substring(a.orddt,7,2) as OrdDate, " & _
                    "       a.ordno, a.ordseq, a.ordcd, c.testnm, a.spccd, a.statfg, " & _
                    "       f.empl_nm as DoctNm, a.rcvdt, substring(a.rcvtm,1,2)+':'+substring(a.rcvtm,3,2) as RcvTm, " & _
                    "       a.stscd, a.dcfg, a.mesg, '혈액' SpcNm, '' MultiFg, b.orddt, b.orddiv, b.wardid, b.hosilid, " & _
                    "       substring(b.ordtm,1,2)+':'+substring(b.ordtm,3,2) as OrdTm, " & _
                    "       substring(a.examdt,1,4)+'-'+substring(a.examdt,5,2)+'-'+substring(a.examdt,7,2) as ExamDt, " & _
                    "       substring(a.examtm,1,2)+':'+substring(a.examtm,3,2) as ExamTm, g.empnm as ExamNm, " & _
                    "       a.workarea, a.accdt, a.accseq, a.stscd as RstIn, '' testdiv, '*' as Verified, " & _
                    "       '' spcyy, 0 spcno, a.unitqty, b.reqdt, b.reqtm " & _
                    "From " & T_LAB102 & " a, " & T_LAB101 & " b, " & T_BBS001 & " c, " & _
                              T_HIS005 & " f, " & T_LAB015 & " g " & _
                    "Where  a.ptid = " & PtId & " " & _
                    "and    a." & QueryKey & " >= '" & FromDate & "' " & _
                    "and    a." & QueryKey & " <= '" & ToDate & "' " & _
                    "and    a.ptid = b.ptid   and   a.orddt = b.orddt   and   a.ordno = b.ordno  and b.orddiv = 'B' " & _
                    "and    c.testcd = a.ordcd " & _
                    "and    f.empl_no =* b.orddoct  " & _
                    "and    g.empid =* a.examdoct   " & _
                    "Order  by orddt desc , ordno, ordseq "
                    
                             '"and     exists ( select * from " & T_LAB302 & " where ptid = a.ptid and orddt = a.orddt and ordno = a.ordno and ordseq = a.ordseq and vfydt <> '' ) "
   'SqlQueryOrders = SqlQueryOrders & _
                             "Union all " & _
                             "Select  substring(a.orddt,3,2)+'-'+substring(a.orddt,5,2)+'-'+substring(a.orddt,7,2) as OrdDate, a.ordno, a.ordseq, c.testnm, a.spccd, a.statfg, " & _
                             "          f.empl_nm as DoctNm, a.rcvdt, substring(a.rcvtm,1,2)+':'+substring(a.rcvtm,3,2) as RcvTm, a.stscd, " & _
                             "          b.orddt, substring(b.ordtm,1,2)+':'+substring(b.ordtm,3,2) as OrdTm, d.field3 as SpcNm, " & _
                             "          d.field1 as MultiFg, a.examdt, substring(a.examtm,1,2)+':'+substring(a.examtm,3,2) as ExamTm, " & _
                             "          a.workarea, a.accdt, a.accseq, a.stscd as RstIn, c.testdiv, '' as Verified " & _
                             "From " & T_LAB102 & " a, " & T_LAB101 & " b, " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_HIS005 & " f " & _
                             "Where a.ptid = " & PtId & " " & _
                             "and     a." & QueryKey & " >= '" & FromDate & "' " & _
                             "and     a." & QueryKey & "<= '" & ToDate & "' " & _
                             "and     a.ptid = b.ptid " & _
                             "and     a.orddt = b.orddt " & _
                             "and     a.ordno = b.ordno " & _
                             "and     c.testcd = a.ordcd " & _
                             "and     c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= a.orddt) " & _
                             "and     d.cdindex  = '" & CD2_Specimen & "' " & _
                             "and     d.cdval1 = a.spccd " & _
                             "and     f.empl_no =* b.orddoct " & _
                             "and     not exists ( select * from " & T_LAB302 & " where ptid = a.ptid and orddt = a.orddt and ordno = a.ordno and ordseq = a.ordseq and vfydt <> '' ) " & _
                             "Order by orddt desc , ordno, ordseq "
                             
End Function
   

'%  32-1. 연속검사인 경우 환자Id, 처방일(채혈일)을 기준으로 연속검사내역을 검색한다.
'%       - Lab203 : 연속검사내역
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlMultiTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Integer) As String
   SqlMultiTest = "Select  a.workarea, a.accdt, a.accseq, a.spcseq " & _
                        "From    " & T_LAB203 & " a, " & T_LAB203 & " b " & _
                        "Where  b.workarea = '" & WorkArea & "' " & _
                        "and      b.accdt = '" & AccDt & "' " & _
                        "and      b.accseq = " & AccSeq & " " & _
                        "and      a.ptid = b.ptid " & _
                        "and      a.orddt = b.orddt " & _
                        "and      a.ordno = b.ordno " & _
                        "and      a.ordseq = b.ordseq " & _
                        "Order by spcseq "

End Function


'%  33. Lab No를 기준으로 결과내역을 검색한다.
'%       - Lab302 : 결과내역
'%       - Lab001 : 검사항목 마스터
'%       - Lab303 : Text 결과 내역
'%       - Lab031 : 결과코드명 (C110-Field1)
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlQueryResults(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Integer, ByVal TestDiv As String, _
                                                Optional ByVal Sex As Variant, Optional ByVal AgeDay As Variant) As String
    Select Case TestDiv
   Case "0":
      'SqlQueryResults = "Select  a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.statfg, a.rstdiv, " & _
                                "           a.lastrst, substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                                "           a.lastvfyid, e.empnm as LastVfyNm, a.vfydt, a.vfytm, a.vfyid, f.empnm as VfyNm, " & _
                                "           a.attrcd, a.mfyfg, a.grpfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, " & _
                                "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                                "           d.field1 as RstCdNm, b.rptseq, g.refvalfrom, g.refvalto, g.refcd " & _
                                "From    " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB015 & " e, " & T_LAB015 & " f, " & T_LAB005 & " g " & _
                                "Where  a.workarea = '" & WorkArea & "' " & _
                                "and      a.accdt = '" & AccDt & "' " & _
                                "and      a.accseq = " & AccSeq & " " & _
                                "and      b.testcd = a.testcd " & _
                                "and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                                "and      d.cdindex = '" & CD1_ItemResult & "' "
      'SqlQueryResults = SqlQueryResults & _
                                "and      d.cdval1 =* a.rstcd " & _
                                "and      e.empid =* a.lastvfyid " & _
                                "and      f.empid =* a.vfyid " & _
                                "and      g.testcd = a.testcd " & _
                                "and      g.spccd = a.spccd " & _
                                "and      g.applydt  = (select max(applydt) from " & T_LAB005 & " where testcd = a.testcd and spccd = a.spccd and applydt <= a.vfydt " & _
                                "                              and applysex = 'B' and agefrom <= " & AgeDay & " and ageto >= " & AgeDay & ") " & _
                                "and      g.applysex = 'B' " & _
                                "and      g.agefrom <= " & AgeDay & " " & _
                                "and      g.ageto >= " & AgeDay & " " & _
                                "and      exists (select * from " & T_LAB005 & " where testcd = a.testcd and spccd = a.spccd and applydt <= a.vfydt " & _
                                "                      and applysex = 'B' and agefrom <= " & AgeDay & " and ageto >= " & AgeDay & ") "
      'SqlQueryResults = SqlQueryResults & _
                                "Union all " & _
                                "Select  a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.statfg, a.rstdiv, " & _
                                "           a.lastrst, substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                                "           a.lastvfyid, e.empnm as LastVfyNm, a.vfydt, a.vfytm, a.vfyid, f.empnm as VfyNm, " & _
                                "           a.attrcd, a.mfyfg, a.grpfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, " & _
                                "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                                "           d.field1 as RstCdNm, b.rptseq, 0 as refvalfrom, 0 as refvalto, '' as refcd " & _
                                "From    " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB015 & " e, " & T_LAB015 & " f " & _
                                "Where  a.workarea = '" & WorkArea & "' " & _
                                "and      a.accdt = '" & AccDt & "' " & _
                                "and      a.accseq = " & AccSeq & " " & _
                                "and      b.testcd = a.testcd " & _
                                "and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                                "and      d.cdindex = '" & CD1_ItemResult & "' "
      'SqlQueryResults = SqlQueryResults & _
                                "and      d.cdval1 =* a.rstcd " & _
                                "and      e.empid =* a.lastvfyid " & _
                                "and      f.empid =* a.vfyid " & _
                                "and      not exists (select * from " & T_LAB005 & " where testcd = a.testcd and spccd = a.spccd and applydt <= a.vfydt " & _
                                "                          and applysex = 'B' and agefrom <= " & AgeDay & " and ageto >= " & AgeDay & ") " & _
                                "Order by rptseq "
      SqlQueryResults = "Select  a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.statfg, a.rstdiv, " & _
                                "           a.lastrst, substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                                "           a.lastvfyid, e.empnm as LastVfyNm, a.vfydt, a.vfytm, a.vfyid, f.empnm as VfyNm, " & _
                                "           a.attrcd, a.mfyfg, a.grpfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, " & _
                                "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                                "           c.rsttxt as TextResult, d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq " & _
                                "From    " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB303 & " c, " & T_LAB031 & " d, " & T_LAB015 & " e, " & T_LAB015 & " f " & _
                                "Where  a.workarea = '" & WorkArea & "' " & _
                                "and      a.accdt = '" & AccDt & "' " & _
                                "and      a.accseq = " & AccSeq & " " & _
                                "and      b.testcd = a.testcd " & _
                                "and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                                "and      c.workarea =* a.workarea " & _
                                "and      c.accdt =* a.accdt " & _
                                "and      c.accseq =* a.accseq " & _
                                "and      c.testcd =* a.testcd " & _
                                "and      d.cdindex = '" & CD1_ItemResult & "' " & _
                                "and      d.cdval1 =* a.testcd " & _
                                "and      d.cdval2 =* a.rstcd " & _
                                "and      e.empid =* a.lastvfyid " & _
                                "and      f.empid =* a.vfyid " & _
                                "Order by RptSeq, a.testcd "
   Case "1":
      SqlQueryResults = "Select  a.testcd, a.rsttype, a.valfg, a.txtfg, a.mfyseq, a.stscd, " & _
                                "           a.vfydt, a.vfytm, a.vfyid, e.empnm as VfyNm, " & _
                                "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                                "           d.txtrst as TextResult " & _
                                "From    " & T_LAB351 & " a, " & T_LAB001 & " b, " & T_LAB353 & " d, " & T_LAB015 & " e " & _
                                "Where  a.workarea = '" & WorkArea & "' " & _
                                "and      a.accdt = '" & AccDt & "' " & _
                                "and      a.accseq = " & AccSeq & " " & _
                                "and      b.testcd = a.testcd " & _
                                "and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                                "and      d.workarea =* a.workarea " & _
                                "and      d.accdt =* a.accdt " & _
                                "and      d.accseq =* a.accseq " & _
                                "and      d.testcd =* a.testcd " & _
                                "and      d.mfyseq = 0 " & _
                                "and      e.empid =* a.vfyid " & _
                                "Order by testcd "
   Case "2":
      SqlQueryResults = "Select  a.testcd, a.rstcd, a.lastrst, substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                                "           a.rstdiv, a.senfg, a.lastvfyid, e.empnm as LastVfyNm, a.vfydt, a.vfytm, a.vfyid, f.empnm as VfyNm, " & _
                                "           a.mfyseq, a.rsttype, a.detailfg, a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
                                "           c.rsttxt as TextResult, d.field1 as RstCdNm1, convert(int,b.rptseq) as RptSeq, '' as RstCdNm2 "   'g.field1 as RstCdNm2 "
      SqlQueryResults = SqlQueryResults & "From    " & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB303 & " c, " & T_LAB031 & " d, " & _
                                                 T_LAB015 & " e, " & T_LAB015 & " f "     '& T_LAB031 & " g "
      SqlQueryResults = SqlQueryResults & "Where  a.workarea = '" & WorkArea & "' " & _
                                "and      a.accdt = '" & AccDt & "' " & _
                                "and      a.accseq = " & AccSeq & " " & _
                                "and      b.testcd = a.testcd " & _
                                "and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                                "and      c.workarea =* a.workarea " & _
                                "and      c.accdt =* a.accdt " & _
                                "and      c.accseq =* a.accseq " & _
                                "and      c.testcd =* a.testcd " & _
                                "and      d.cdindex = '" & CD1_ItemResult & "' " & _
                                "and      d.cdval1 =* a.testcd " & _
                                "and      d.cdval2 =* a.rstcd " & _
                                "and      e.empid =* a.lastvfyid " & _
                                "and      f.empid =* a.vfyid " & _
                                "Order by RptSeq, a.testcd "
                                '**미생물 Batch결과코드는 일반결과코드에도 똑같이 등록이 된다.
                                '"and      g.cdindex = '" & CD1_MBatchRst & "' " & _
                                "and      g.cdval2 =* a.rstcd "

   End Select
End Function

'%  34. Text Result
'%       - Lab303  : Text결과 내역 검색
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlGetRstText(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Integer, ByVal TestCd As String) As String
   SqlGetRstText = "Select a.rsttxt " & _
                             "From " & T_LAB303 & " a " & _
                             "Where  a.workarea = '" & WorkArea & "' " & _
                             "and      a.accdt = '" & AccDt & "' " & _
                             "and      a.accseq = " & AccSeq & " " & _
                             "and      a.testcd = '" & TestCd & "' "
End Function

'%  34-1. Supplemental Report
'%       - 305 : Text결과 수정내역 검색
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlGetSuppText(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Integer, ByVal TestCd As String) As String
   SqlGetSuppText = "Select a.rsttxt, b.empnm, a.seq, substring(a.mfydt,3,2)+'-'+substring(a.mfydt,5,2)+'-'+substring(a.mfydt,7,2)+' '+substring(a.mfytm,1,2)+':'+substring(a.mfytm,3,2) as MfyDtTm " & _
                             "From " & T_LAB305 & " a, " & T_LAB015 & " b " & _
                             "Where  a.workarea = '" & WorkArea & "' " & _
                             "and      a.accdt = '" & AccDt & "' " & _
                             "and      a.accseq = " & AccSeq & " " & _
                             "and      a.testcd = '" & TestCd & "' " & _
                             "and      b.empid =* a.mfyid " & _
                             "Order by seq "
End Function

'%  34-2. Supplemental Report
'%       - Lab353 : 기타검사 Text결과 수정내역 검색
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlETextRst(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String) As String
   SqlETextRst = "Select txtrst as TextResult From " & T_LAB353 & " " & _
                       "Where workarea = '" & pWorkArea & "' " & _
                       "and    accdt = '" & pAccDt & "' " & _
                       "and    accseq = " & pAccSeq & " " & _
                       "and    testcd = '" & pTestCd & "' " & _
                       "and    mfyseq > 0 "
End Function

'%  35. 미생물 감수성결과 (LAB405)
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlSenResult(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Integer, ByVal TestCd As String) As String
   SqlSenResult = "Select a.seq, b.text1 as MicroNm, c.field2 as GrowthQty, a.scnt, a.srst1, a.srst2, a.srst3, a.srst4, a.srst5, a.srst6, a.srst7, " & _
                         "          a.srst8, a.srst9, a.srst10, a.srst11, a.srst12, a.srst13, a.srst14, a.srst15, a.srst16, a.srst17, a.srst18, a.srst19, a.srst20, " & _
                         "          a.srst21, a.srst22, a.srst23, a.srst24, a.srst25 " & _
                         "From   " & T_LAB404 & " d, " & T_LAB405 & " a, " & T_LAB032 & " b, " & T_LAB032 & " c " & _
                         "Where d.workarea = '" & WorkArea & "' " & _
                         "and     d.accdt = '" & AccDt & "' " & _
                         "and     d.accseq = " & AccSeq & " " & _
                         "and     d.testcd = '" & TestCd & "' " & _
                         "and     a.workarea = d.workarea " & _
                         "and     a.accdt = d.accdt " & _
                         "and     a.accseq = d.accseq " & _
                         "and     a.testcd = d.testcd  " & _
                         "and     a.mfyseq = d.mfyseq  " & _
                         "and     b.cdindex = '" & CD2_Microbe & "' " & _
                         "and     b.cdval1 =* a.mnmcd " & _
                         "and     c.cdindex = '" & CD2_Volume & "' " & _
                         "and     c.cdval1 =* a.mqtcd " & _
                         "Order by seq "
End Function


'%  36. 외부의뢰내역 (LAB405)
'%       - Calling from [ frm156Referral ] : 외부의뢰
Public Function SqlOutLabData(ByVal OutLabCd As String, ByVal FromDt As String, ByVal ToDt As String, Optional ByVal StsCd As Variant) As String

   Dim tmpStr As String
   
   If Not IsMissing(StsCd) Then
      tmpStr = " and a.stscd = '" & StsCd & "' "
   Else: tmpStr = ""
   End If
   
   SqlOutLabData = "Select a.workarea+'-'+a.accdt+'-'+convert(varchar,a.accseq)  as LabNo, a.rcvdt, a.ptid, b.ptnt_nm as PtNm, substring(b.ptnt_prsn_no,3,6)+'-'+substring(b.ptnt_prsn_no,9,7) as ssn, substring(b.ptnt_prsn_no,9,1) as sex, c.testnm, d.field3 as SpcNm, " & _
                           "          substring(b.ptnt_prsn_no,1,8) as dob, a.stscd, a.workarea, a.accdt, a.accseq, a.testcd , f.dept_nm as DeptNm, g.ptnt_charge_gb as Gubun " & _
                           "From   " & T_LAB205 & " a, " & T_HIS001 & " b, " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_LAB201 & " e, " & T_HIS003 & " f, " & T_HIS010 & " g " & _
                           "Where a.outlabcd = '" & OutLabCd & "' " & _
                           "and    a.rcvdt >= '" & FromDt & "' " & _
                           "and    a.rcvdt <= '" & ToDt & "' " & tmpStr & _
                           "and    b.ptnt_no = a.ptid " & _
                           "and    c.testcd = a.testcd " & _
                           "and    c.applydt = ( select max(applydt) from " & T_LAB001 & " where testcd = c.testcd ) " & _
                           "and    d.cdindex = '" & CD2_Specimen & "' " & _
                           "and    d.cdval1 = a.spccd " & _
                           "and    a.workarea = e.workarea " & _
                           "and    a.accdt = e.accdt " & _
                           "and    a.accseq = e.accseq " & _
                           "and    e.deptcd = f.dept_cd " & _
                           "and    a.testcd = g.medfee_cd " & _
                           "and    g.aply_ymd = ( select max(aply_ymd) from " & T_HIS010 & " where medfee_cd = g.medfee_cd ) " & _
                           "Order by a.rcvdt, a.workarea, a.accdt, a.accseq "
                           

End Function

'%  37. 같은 Lab No 중 확인된 결과존재 여부 (LAB302)
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlVerifiedResult1(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer) As String

   SqlVerifiedResult1 = "Select * from " & T_LAB302 & " where workarea = '" & pWorkArea & "' and accdt = '" & pAccDt & "' " & _
                               "and accseq = " & pAccSeq & " and vfydt is not null and vfydt <> '' "

End Function

'%  38. 같은 처방번호,처방Seq로 발생한 항목중 결과확인 상태인 데이타 존재여부 (LAB302)
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlVerifiedResult2(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, ByVal pOrdSeq As Integer) As String
            
   SqlVerifiedResult2 = "Select * from " & T_LAB302 & " b where ptid = " & pPtId & " and orddt = '" & pOrdDt & "' " & _
                               "and ordno = " & pOrdNo & " and ordseq = " & pOrdSeq & " and vfydt is not null and vfydt <> '' "

End Function


'%  39. Remark
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlGetRemark(ByVal pRmkCd As String) As String

   SqlGetRemark = "Select text1 as Remark from " & T_LAB034 & " where cdindex = '" & CD4_Remark & "' and cdval1 = '" & pRmkCd & "'"

End Function


'%  40. Foot Note
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlGetFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer) As String
   
   SqlGetFootNote = "Select rsttxt as FootNote " & _
                            "from " & T_LAB304 & " " & _
                            "where workarea = '" & pWorkArea & "' " & _
                            "and     accdt = '" & pAccDt & "' " & _
                            "and     accseq = " & pAccSeq

End Function


'%  41. 기타검사 Value Result
'%       - Calling from [ frm401ResultView ] : 결과조회
Public Function SqlEValueRst(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String, ByVal pMfyFg As String) As String

   SqlEValueRst = "Select seq, head, valrst, unit " & _
                         "From " & T_LAB352 & " " & _
                         "Where workarea = '" & pWorkArea & "' " & _
                         "and    accdt = '" & pAccDt & "' " & _
                         "and    accseq = " & pAccSeq & " " & _
                         "and    testcd = '" & pTestCd & "' " & _
                         "and    mfyseq = " & pMfyFg & " " & _
                         "Order By seq "

End Function
   

'%  42. 복수검체
'%       - Calling from [ clsCollection ] :
Public Function SqlGetMultiSpc(ByVal SpcCd As String) As String
   
   SqlGetMultiSpc = "Select cdval2 as SpcSeq, field1 as SpcCd  " & _
                            "From " & T_LAB031 & " " & _
                            "Where cdindex = '" & CD1_MultiSpc & "' " & _
                            "and    cdval1 = '" & SpcCd & "' "

End Function

'%  42-1. 복수검체 갯수
'%       - Calling from [ clsCollection ] :
Public Function SqlMultiSpcCnt1(ByVal SpcCd As String) As String

   SqlMultiSpcCnt1 = "Select Count(*) as Count " & _
                              "From " & T_LAB031 & " " & _
                              "Where cdindex = '" & CD1_MultiSpc & "' " & _
                              "and    cdval1 = '" & SpcCd & "' "

End Function

'%  42-2. 이미 생성된 복수검체 갯수
'%       - Calling from [ clsCollection ] :
Public Function SqlMultiSpcCnt2(ByVal PtId As String, ByVal OrdDt As String, ByVal OrdNo As Integer, ByVal OrdSeq As Integer) As String

   SqlMultiSpcCnt2 = "Select * " & _
                            "From " & T_LAB203 & " " & _
                            "Where ptid = " & PtId & " " & _
                            "and     orddt = '" & OrdDt & "' " & _
                            "and     ordno = " & OrdNo & " " & _
                            "and     ordseq = " & OrdSeq & " "
   'SqlMultiSpcCnt2 = "Select Count(*) as SpcCnt " & _
                            "From " & T_LAB203 & " " & _
                            "Where ptid = " & PtId & " " & _
                            "and     orddt = '" & OrdDt & "' " & _
                            "and     ordno = " & OrdNo & " " & _
                            "and     ordseq = " & OrdSeq & " "

End Function
   


'%  43. 일반접수시 처방 Display
'%       - Calling from [ frm155Accession ] :
Public Function SqlOrdersForAccess(ByVal QueryOption As Integer, Optional ByVal Key1 As Variant, _
                                                      Optional ByVal Key2 As Variant, Optional ByVal Key3 As Variant) As String

      Dim tmpStr As String

      If QueryOption = 1 Then
         tmpStr = "Where a.workarea = '" & Key1 & "' " & _
                     "and    a.accdt = '" & Key2 & "' " & _
                     "and    a.accseq = " & Key3 & " "
      Else
         tmpStr = "Where a.spcyy = '" & Key1 & "' " & _
                     "and     a.spcno = " & Key2 & " "
      End If
                   
      SqlOrdersForAccess = "Select a.workarea, a.accdt, a.accseq, a.ptid, b.ptnt_nm as PtNm, c.dept_nm as DeptNm, a.wardid+'-'+a.hosilid as Location, " & _
                                       "          a.stscd, a.statfg, a.storecd, d.field3 as SpcNm, e.orddt, e.ordcd, f.testnm " & _
                                       "From   " & T_LAB201 & " a, " & T_HIS001 & " b, " & T_HIS003 & " c, " & T_LAB032 & " d, " & T_LAB102 & " e, " & T_LAB001 & " f " & _
                                       tmpStr & _
                                       "and    a.multifg = '' " & _
                                       "and    b.ptnt_no = a.ptid " & _
                                       "and    c.dept_cd =* a.deptcd " & _
                                       "and    d.cdindex = '" & CD2_Specimen & "' " & _
                                       "and    d.cdval1 = a.spccd " & _
                                       "and    a.workarea = e.workarea " & _
                                       "and    a.accdt = e.accdt " & _
                                       "and    a.accseq = e.accseq " & _
                                       "and    f.testcd = e.ordcd " & _
                                       "and    f.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = f.testcd) "
      
      SqlOrdersForAccess = SqlOrdersForAccess & "Union All " & _
                                       "Select a.workarea, a.accdt, a.accseq, a.ptid, b.ptnt_nm as ptnm, c.dept_nm as DeptNm, a.wardid+'-'+a.hosilid as Location, " & _
                                       "          a.stscd, a.statfg, a.storecd, d.field3 as SpcNm, e.orddt, e.ordcd, f.testnm " & _
                                       "From   " & T_LAB201 & " a, " & T_HIS001 & " b, " & T_HIS003 & " c, " & T_LAB032 & " d, " & T_LAB102 & " e, " & T_LAB001 & " f, " & T_LAB203 & " g " & _
                                       tmpStr & _
                                       "and    a.multifg <> '' " & _
                                       "and    b.ptnt_no = a.ptid " & _
                                       "and    c.dept_cd =* a.deptcd " & _
                                       "and    d.cdindex = '" & CD2_Specimen & "' " & _
                                       "and    d.cdval1 = a.spccd " & _
                                       "and    a.workarea = g.workarea " & _
                                       "and    a.accdt = g.accdt " & _
                                       "and    a.accseq = g.accseq " & _
                                       "and    e.ptid = g.ptid " & _
                                       "and    e.orddt = g.orddt " & _
                                       "and    e.ordno = g.ordno " & _
                                       "and    e.ordseq = g.ordseq " & _
                                       "and    f.testcd = e.ordcd " & _
                                       "and    f.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = f.testcd) " & _
                                       "Order by orddt "

End Function


'%  44. 의뢰내역 Update
'%       - Calling from [ frm156Referral ] :
Public Function SqlUpdateLab205(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, _
                                                ByVal TestCd As String, ByVal SendId As String) As String

   SqlUpdateLab205 = "Update " & T_LAB205 & " " & _
                              "set senddt = " & CS_SybaseDate & ", " & _
                              "     sendid = " & DBNum(SendId) & ", " & _
                              "     stscd = " & DBStr(STS_Worksheet) & " " & _
                              "where  workarea = " & DBStr(WorkArea) & " " & _
                              "and      accdt = " & DBStr(AccDt) & " " & _
                              "and      accseq = " & DBNum(AccSeq) & " " & _
                              "and      testcd = " & DBStr(TestCd)

End Function


'%  45. 바코드라벨 재출력을 위해 해당 환자의 처방내역을 조회한다.
'%       - Calling from [frm157BarReprint] : 라벨재발행
'%       - 검사구분, WorkArea, 검체코드, 보관구분, 응급여부, 희망채혈일시,
'%         복수검체여부, 검체군, 처방일, 처방번호, 처방Seq, 처방코드
'%       - Lab201 : 채혈접수내역
'%       - Lab102 : 처방Body내역
'%       - Lab001 : 검사항목 내역
'%       - Lab032 : 검체명, 건물명
Public Function SqlBarReprint(ByVal intOption As Integer, ParamArray strKeys() As Variant) As String

    Dim tmpStr As String
    Dim tmpStr1 As String
    Dim tmpSql(3) As String
    Dim blnAllFg As Boolean
   
    tmpStr = ""
    blnAllFg = False
    If intOption = 1 Then  '환자ID-처방일 기준
        tmpStr = tmpStr & "b.ptid = " & strKeys(0) & " and b.orddt = '" & _
                 strKeys(1) & "' and b.donefg > '0'  and    b.dcfg = '' "
        tmpStr1 = "and e.ptid = b.ptid and e.orddt = b.orddt " & _
                  "and e.ordno = b.ordno and e.ordseq = b.ordseq " & _
                  "and a.workarea = e.workarea and a.accdt = e.accdt " & _
                  "and a.accseq = e.accseq "
        If UBound(strKeys) = 2 Then
            If strKeys(2) = "" Then blnAllFg = True
        End If
    Else        '접수번호 기준
        tmpStr = tmpStr & "a.workarea = '" & strKeys(0) & "' and a.accdt = '" & _
                 strKeys(1) & "' and a.accseq = " & strKeys(2) & " "
        tmpStr1 = "and e.workarea = a.workarea and e.accdt = a.accdt " & _
                  "and e.accseq = a.accseq and b.ptid = e.ptid " & _
                  "and b.orddt = e.orddt and b.ordno = e.ordno " & _
                  "and b.ordseq = e.ordseq   and    b.dcfg = '' "
    End If
    
    '임상병리
    tmpSql(1) = "Select c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                "       a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                "       a.deptcd, a.hosilid, a.workarea+'-'+substring(a.accdt,3,6)+'-'+" & _
                "       convert(char,a.accseq) as LabNo, d.field5 as SpcNm, substring(a.coldt,5,2)+'/'+" & _
                "       substring(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "       e.field1 as BuildNm, f.ptnt_nm as PtNm, i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                "From " & T_LAB201 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                          T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i " & _
                "Where " & tmpStr & _
                "and    a.multifg = '' " & _
                "and    b.workarea = a.workarea " & _
                "and    b.accdt = a.accdt " & _
                "and    b.accseq = a.accseq " & _
                "and    b.dcfg = '' " & _
                "and    c.testcd = b.ordcd " & _
                "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd " & _
                                    "and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                "and    d.cdindex = '" & CD2_Specimen & "' " & _
                "and    d.cdval1 = a.spccd " & _
                "and    e.cdindex = '" & CD2_Buildings & "' " & _
                "and    e.cdval1 =* a.buildcd " & _
                "and    f.ptnt_no = a.ptid " & _
                "and    i.ptid = b.ptid " & _
                "and    i.orddt = b.orddt " & _
                "and    i.ordno = b.ordno "
    
    'Union
    tmpSql(1) = tmpSql(1) & "Union All " & _
                "Select c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                "       a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                "       a.deptcd, a.hosilid, a.workarea+'-'+substring(a.accdt,3,6)+'-'+" & _
                "       convert(char,a.accseq) as LabNo, d.field5 as SpcNm, substring(a.coldt,5,2)+'/'+" & _
                "       substring(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "       g.field1 as BuildNm, h.ptnt_nm as PtNm, i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                "From " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                          T_LAB203 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g, " & _
                          T_HIS001 & " h, " & T_LAB101 & " i " & _
                "Where " & tmpStr & tmpStr1 & _
                "and    a.multifg <> '' " & _
                "and    f.cdindex = '" & CD1_Panel & "' " & _
                "and    f.cdval1 = b.ordcd  and  f.cdval2 = e.spcseq " & _
                "and    c.testcd = f.field1 " & _
                "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd " & _
                                    "and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                "and    d.cdindex = '" & CD2_Specimen & "'  and   d.cdval1 = a.spccd " & _
                "and    g.cdindex = '" & CD2_Buildings & "' and   g.cdval1 =* a.buildcd " & _
                "and    h.ptnt_no = a.ptid " & _
                "and    i.ptid = b.ptid " & _
                "and    i.orddt = b.orddt " & _
                "and    i.ordno = b.ordno "

    '해부병리
    tmpSql(2) = "Select c.testnm, c.abbrnm5, b.ptid, b.workarea, b.accdt, b.accseq, " & _
                "       i.spcyy, i.spcno, i.spccd, '' storecd, b.stscd, b.statfg, '" & APS_BLDCD & "' as buildcd, " & _
                "       a.deptcd, a.hosilid, i.spcyy+'-'+convert(char,i.spcno) as LabNo, " & _
                "       d.spcnm5 as SpcNm, substring(i.coldt,5,2)+'/'+" & _
                "       substring(i.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                " '" & APS_BARLOC & "' as BuildNm, f.ptnt_nm as PtNm, " & _
                "       i.coldt as reqdt, i.coltm as reqtm, '" & APS_ORDDIV & "' as orddiv, c.fzfg " & _
                "From " & T_APS101 & " a, " & T_LAB102 & " b, " & T_APS001 & " c, " & _
                          T_APS002 & " d, " & T_HIS001 & " f, " & T_APS103 & " i " & _
                "Where " & tmpStr & _
                "and    i.ptid = b.ptid " & _
                "and    i.orddt = b.orddt " & _
                "and    i.ordno = b.ordno   and   i.ordseq = b.ordseq " & _
                "and    a.ptid = i.ptid " & _
                "and    a.orddt = i.orddt " & _
                "and    a.ordno = i.ordno " & _
                "and    c.testcd = b.ordcd " & _
                "and    d.spccd = b.spccd " & _
                "and    f.ptnt_no = a.ptid "
    '혈액은행
    tmpSql(3) = "Select c.testnm, c.abbrnm5, a.ptid, 'XM' as workarea, '' as accdt, 0 as accseq, " & _
                "       a.spcyy, a.spcno, '' spccd, '' as storecd, '1' as stscd, b.statfg, a.buildcd, " & _
                "       i.deptcd, i.hosilid, '' as LabNo, '혈액' as SpcNm, substring(a.coldt,5,2)+'/'+" & _
                "       substring(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "       e.field1 as BuildNm, f.ptnt_nm as PtNm, " & _
                "       i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                "From " & T_BBS201 & " a, " & T_LAB102 & " b, " & T_BBS001 & " c, " & _
                          T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i " & _
                "Where " & tmpStr & _
                "and    i.ptid = b.ptid " & _
                "and    i.orddt = b.orddt " & _
                "and    i.ordno = b.ordno " & _
                "and    i.orddiv = '" & C_WORKAREA & "' " & _
                "and    a.ptid = i.ptid " & _
                "and    a.spcyy = (select max(spcyy) from " & T_BBS201 & " where ptid = i.ptid) " & _
                "and    a.spcno = (select max(spcno) from " & T_BBS201 & _
                                 " where ptid = a.ptid and spcyy = a.spcyy) " & _
                "and    c.testcd = b.ordcd " & _
                "and    c.testdiv <> '3' " & _
                "and    e.cdindex = '" & CD2_Buildings & "' " & _
                "and    e.cdval1 =* a.buildcd " & _
                "and    f.ptnt_no = a.ptid "
    
    If blnAllFg Then
        SqlBarReprint = tmpSql(1) & " union all " & tmpSql(2) & " union all " & tmpSql(3)
    Else
        Select Case strKeys(2)
        Case "A": SqlBarReprint = tmpSql(2)
        Case "B": SqlBarReprint = tmpSql(3)
        Case "L": SqlBarReprint = tmpSql(1)
        End Select
    End If
    SqlBarReprint = SqlBarReprint & " Order By LabNo, b.ordno, b.ordcd "

End Function
   
   
'%  46. 해당 기간동안 발생한 모든 검사결과를 조회한다.
'%       - Lab201 : 채혈접수내역
'%       - Lab302 : 일반검사 결과내역
'%       - Lab351,353 : 특수검사 결과내역
'%       - Lab404 : 미생물검사 결과내역
'%       - Lab001 : 검사항목 마스터
'%       - Lab303 : Text 결과 내역
'%       - Lab031 : 결과코드명 (C110-Field1)
'%       - Calling from [ frm404AllResult ] : 전체결과조회
Public Function SqlQueryAllResults_Back(ByVal PtId As String, ByVal QueryKey As String, ByVal FromDate As String, ByVal ToDate As String, _
                                                   ByVal TestDiv As String, Optional ByVal Sex As Variant, Optional ByVal AgeDay As Variant) As String

      Dim Result1 As String
      Dim Result2 As String
      Dim Result3 As String
      Dim tmpStr As String
      Dim tmpOrderby As String
      
      tmpStr = "  substring(x." & QueryKey & ",3,2)+'-'+substring(x." & QueryKey & ",5,2)+'-'+substring(x." & QueryKey & ",7,2) as KeyDate "
      tmpOrderby = "Order by KeyDate desc, workarea, accdt, accseq, RptSeq, a.testcd "

         Result1 = "Select  a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, " & _
                       "           a.lastvfyid, e.empnm as LastVfyNm, a.vfydt as VfyDt, a.vfytm as VfyTm, a.vfyid as VfyId, f.empnm as VfyNm, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
                       "           a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
                       "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
                       "           c.deptcd, c.sex, c.ageday, k.orddoct, g.empl_nm as DoctNm, c.colid, h.empnm as ColNm, c.rcvid, i.empnm as RcvNm, " & _
                       "           substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
                       "           substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       "           substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
                       "           substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
                       "           substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
                       "           d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
                       "From    " & T_LAB102 & " x, " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                                        T_LAB031 & " d, " & T_LAB015 & " e, " & T_LAB015 & " f, " & T_HIS005 & " g, " & _
                                        T_LAB015 & " h, " & T_LAB015 & " i, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
                       "Where  x.ptid = " & PtId & " " & _
                       "and      x." & QueryKey & " >= '" & FromDate & "' and      x." & QueryKey & "<= '" & ToDate & "' " & _
                       "and      a.ptid = x.ptid    and    a.orddt = x.orddt    and    a.ordno = x.ordno    and    a.ordseq = x.ordseq " & _
                       "and      k.ptid = x.ptid    and    k.orddt = x.orddt    and    k.ordno = x.ordno " & _
                       "and      b.testcd = a.testcd    and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                       "and      c.workarea = a.workarea    and    c.accdt = a.accdt    and    c.accseq = a.accseq " & _
                       "and      d.cdindex = '" & CD1_ItemResult & "'   and    d.cdval1 =* a.testcd    and    d.cdval2 =* a.rstcd " & _
                       "and      j.cdindex = '" & CD2_Specimen & "'   and    j.cdval1 =* a.spccd  " & _
                       "and      e.empid =* a.lastvfyid   and   f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
         Result2 = "Select  a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       "           0 as lastvfyid, '' as LastVfyNm, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, f.empnm as VfyNm, '1' as testdiv, c.footnotefg, c.rmkcd, " & _
                       "           convert(char,a.mfyseq) as mfyfg, a.txtfg, a.rsttype, '' as detailfg, '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
                       "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
                       "           c.deptcd, c.sex, c.ageday, k.orddoct, g.empl_nm as DoctNm, c.colid, h.empnm as ColNm, c.rcvid, i.empnm as RcvNm, " & _
                       "           substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
                       "           '' as LstVfyDtTm, " & _
                       "           substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
                       "           substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
                       "           substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
                       "           '' as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
                       "From    " & T_LAB102 & " x, " & T_LAB351 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                                        T_LAB015 & " f , " & T_HIS005 & " g, " & _
                                        T_LAB015 & " h, " & T_LAB015 & " i, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
                       "Where  x.ptid = " & PtId & " " & _
                       "and      x." & QueryKey & " >= '" & FromDate & "' and      x." & QueryKey & "<= '" & ToDate & "' " & _
                       "and      a.ptid = x.ptid    and    a.orddt = x.orddt    and    a.ordno = x.ordno    and    a.ordseq = x.ordseq " & _
                       "and      k.ptid = x.ptid    and    k.orddt = x.orddt    and    k.ordno = x.ordno " & _
                       "and      b.testcd = a.testcd    and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                       "and      c.workarea = a.workarea    and    c.accdt = a.accdt    and    c.accseq = a.accseq " & _
                       "and      j.cdindex = '" & CD2_Specimen & "'   and    j.cdval1 =* c.spccd  " & _
                       "and      f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
         Result3 = "Select  a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
                       "           a.lastvfyid, e.empnm as LastVfyNm, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, f.empnm as VfyNm, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
                       "           convert(char,a.mfyseq) as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
                       "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
                       "           c.deptcd, c.sex, c.ageday, k.orddoct, g.empl_nm as DoctNm, c.colid, h.empnm as ColNm, c.rcvid, i.empnm as RcvNm, " & _
                       "           substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
                       "           substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       "           substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
                       "           substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
                       "           substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
                       "           d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
                       "From    " & T_LAB102 & " x, " & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                                        T_LAB031 & " d, " & T_LAB015 & " e , " & T_LAB015 & " f , " & T_HIS005 & " g, " & _
                                        T_LAB015 & " h, " & T_LAB015 & " i, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
                       "Where  x.ptid = " & PtId & " " & _
                       "and      x." & QueryKey & " >= '" & FromDate & "' and      x." & QueryKey & "<= '" & ToDate & "' " & _
                       "and      a.ptid = x.ptid    and    a.orddt = x.orddt    and    a.ordno = x.ordno    and    a.ordseq = x.ordseq   and  a.rsttype = '" & MRT_Stain & "' " & _
                       "and      k.ptid = x.ptid    and    k.orddt = x.orddt    and    k.ordno = x.ordno " & _
                       "and      b.testcd = a.testcd    and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                       "and      c.workarea = a.workarea    and    c.accdt = a.accdt    and    c.accseq = a.accseq " & _
                       "and      d.cdindex = '" & CD1_ItemResult & "'   and    d.cdval1 =* a.testcd    and    d.cdval2 =* a.rstcd " & _
                       "and      j.cdindex = '" & CD2_Specimen & "'   and    j.cdval1 =* c.spccd  " & _
                       "and      e.empid =* a.lastvfyid   and    f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
         Result3 = Result3 & "Union all " & _
                       "Select  a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
                       "           a.lastvfyid, e.empnm as LastVfyNm, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, f.empnm as VfyNm, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
                       "           convert(char,a.mfyseq) as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
                       "           a.ptid, a.orddt, a.ordno, a.ordseq, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
                       "           c.deptcd, c.sex, c.ageday, k.orddoct, g.empl_nm as DoctNm, c.colid, h.empnm as ColNm, c.rcvid, i.empnm as RcvNm, " & _
                       "           substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
                       "           substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       "           substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
                       "           substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
                       "           substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
                       "           d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
                       "From    " & T_LAB102 & " x, " & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                                        T_LAB031 & " d, " & T_LAB015 & " e , " & T_LAB015 & " f , " & T_HIS005 & " g, " & _
                                        T_LAB015 & " h, " & T_LAB015 & " i, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
                       "Where  x.ptid = " & PtId & " " & _
                       "and      x." & QueryKey & " >= '" & FromDate & "' and      x." & QueryKey & "<= '" & ToDate & "' " & _
                       "and      a.ptid = x.ptid    and    a.orddt = x.orddt    and    a.ordno = x.ordno    and    a.ordseq = x.ordseq   and  a.rsttype <> '" & MRT_Stain & "' " & _
                       "and      k.ptid = x.ptid    and    k.orddt = x.orddt    and    k.ordno = x.ordno " & _
                       "and      b.testcd = a.testcd    and      b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
                       "and      c.workarea = a.workarea    and    c.accdt = a.accdt    and    c.accseq = a.accseq " & _
                       "and      d.cdindex = '" & CD1_MBatchRst & "'   and    d.cdval2 =* a.rstcd " & _
                       "and      j.cdindex = '" & CD2_Specimen & "'   and    j.cdval1 =* c.spccd  " & _
                       "and      e.empid =* a.lastvfyid   and    f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
      
      Select Case TestDiv
         Case "0":  SqlQueryAllResults_Back = Result1 & tmpOrderby   '일반결과만..
         Case "1":  SqlQueryAllResults_Back = Result2 & tmpOrderby   '특수검사결과만..
         Case "2":  SqlQueryAllResults_Back = Result3 & tmpOrderby   '미생물결과만..
         Case "3":  SqlQueryAllResults_Back = Result1 & "Union all " & Result2 & "Union all " & Result3 & tmpOrderby      '모두...
      End Select
                           
End Function

Public Function SqlQueryAllResults(ByVal PtId As String, ByVal QueryKey As String, ByVal FromDate As String, ByVal ToDate As String, _
                                                   ByVal TestDiv As String, Optional ByVal Sex As Variant, Optional ByVal AgeDay As Variant) As String

    Dim Result1 As String
    Dim Result2 As String
    Dim Result3 As String
    Dim Result4 As String   '해부병리
    Dim Result5 As String   '혈액은행
    Dim tmpStr As String
    Dim tmpOrderby As String
      
    tmpStr = "  substring(x." & QueryKey & ",3,2)+'-'+substring(x." & QueryKey & ",5,2)+'-'+substring(x." & QueryKey & ",7,2) as KeyDate "
    'tmpStr = "  x." & QueryKey & "  as KeyDate "
    tmpOrderby = "Order by KeyDate desc, workarea, accdt, accseq, RptSeq, x.ordcd "

    Result1 = "Select a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "       a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "       a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "       a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "       x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, " & _
              "       b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "       c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
              "       substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
              "       substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
              "       substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
              "       substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
              "       substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
              "       0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
              "From " & T_LAB102 & " x, " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                        T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
              "Where  x.ptid = " & PtId & " " & _
              "and    x." & QueryKey & " >= '" & FromDate & "' and  x." & QueryKey & "<= '" & ToDate & "' " & _
              "and    a.ptid = x.ptid  and  a.orddt = x.orddt  and  a.ordno = x.ordno  and  a.ordseq = x.ordseq " & _
              "and    k.ptid = x.ptid  and  k.orddt = x.orddt  and  k.ordno = x.ordno " & _
              "and    b.testcd = a.testcd  and  b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
              "and    c.workarea = a.workarea  and  c.accdt = a.accdt  and  c.accseq = a.accseq " & _
              "and    d.cdindex = '" & CD1_ItemResult & "' and d.cdval1 =* a.testcd  and  d.cdval2 =* a.rstcd " & _
              "and    j.cdindex = '" & CD2_Specimen & "'   and j.cdval1 =* a.spccd  "
                  '"and      e.empid =* a.lastvfyid   and   f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
    Result2 = "Select a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "       '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
              "       0 as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "       c.footnotefg, c.rmkcd, convert(char,a.mfyseq) as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "       '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "       a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "       b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
              "       substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
              "       '' as LstVfyDtTm, " & _
              "       substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
              "       substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
              "       substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
              "       0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
              "From " & T_LAB102 & " x, " & T_LAB351 & " a, " & T_LAB001 & " b, " & _
                        T_LAB201 & " c, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
              "Where  x.ptid = " & PtId & " " & _
              "and    x." & QueryKey & " >= '" & FromDate & "' and  x." & QueryKey & "<= '" & ToDate & "' " & _
              "and    a.ptid = x.ptid  and  a.orddt = x.orddt  and  a.ordno = x.ordno  and  a.ordseq = x.ordseq " & _
              "and    k.ptid = x.ptid  and  k.orddt = x.orddt  and  k.ordno = x.ordno " & _
              "and    b.testcd = a.testcd  and  b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
              "and    c.workarea = a.workarea  and  c.accdt = a.accdt  and  c.accseq = a.accseq " & _
              "and    j.cdindex = '" & CD2_Specimen & "'  and  j.cdval1 =* c.spccd  "
                  '"and      f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
    Result3 = "Select a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "       a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
              "       convert(char,a.mfyseq) as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "       a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "       c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
              "       substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
              "       substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
              "       substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
              "       substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
              "       substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
              "       0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
              "From " & T_LAB102 & " x, " & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                        T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
              "Where  x.ptid = " & PtId & " " & _
              "and    x." & QueryKey & " >= '" & FromDate & "' and  x." & QueryKey & "<= '" & ToDate & "' " & _
              "and    a.ptid = x.ptid  and  a.orddt = x.orddt  and  a.ordno = x.ordno  and  a.ordseq = x.ordseq  and  a.rsttype = '" & MRT_Stain & "' " & _
              "and    k.ptid = x.ptid  and  k.orddt = x.orddt  and  k.ordno = x.ordno " & _
              "and    b.testcd = a.testcd  and  b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
              "and    c.workarea = a.workarea  and  c.accdt = a.accdt  and  c.accseq = a.accseq " & _
              "and    d.cdindex = '" & CD1_ItemResult & "'  and  d.cdval1 =* a.testcd  and  d.cdval2 =* a.rstcd " & _
              "and    j.cdindex = '" & CD2_Specimen & "'  and  j.cdval1 =* c.spccd  "
                  '"and      e.empid =* a.lastvfyid   and    f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
    Result3 = Result3 & "Union all " & _
              "Select a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "       '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "       a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, convert(char,a.mfyseq) as mfyfg, " & _
              "       '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "       a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "       b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & _
              "       substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
              "       substring(a.lastvfydt,3,2)+'-'+substring(a.lastvfydt,5,2)+'-'+substring(a.lastvfydt,7,2)+' '+substring(a.lastvfytm,1,2)+':'+substring(a.lastvfytm,3,2) as LstVfyDtTm, " & _
              "       substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
              "       substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
              "       substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
              "       0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq, " & tmpStr & _
              "From " & T_LAB102 & " x, " & T_LAB404 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & _
                        T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB101 & " k " & _
              "Where  x.ptid = " & PtId & " " & _
              "and    x." & QueryKey & " >= '" & FromDate & "' and  x." & QueryKey & "<= '" & ToDate & "' " & _
              "and    a.ptid = x.ptid  and  a.orddt = x.orddt  and  a.ordno = x.ordno  and  a.ordseq = x.ordseq  and  a.rsttype <> '" & MRT_Stain & "' " & _
              "and    k.ptid = x.ptid  and  k.orddt = x.orddt  and  k.ordno = x.ordno " & _
              "and    b.testcd = a.testcd  and  b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = a.testcd and applydt <= a.orddt ) " & _
              "and    c.workarea = a.workarea  and  c.accdt = a.accdt    and    c.accseq = a.accseq " & _
              "and    d.cdindex = '" & CD1_ItemResult & "' and  d.cdval1 =* a.testcd  and  d.cdval2 =* a.rstcd " & _
              "and    j.cdindex = '" & CD2_Specimen & "'   and  j.cdval1 =* c.spccd  "
                  '"and      e.empid =* a.lastvfyid   and    f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
    Result4 = "Select x.workarea as WorkArea, x.accdt as AccDt, x.accseq as AccSeq, x.ordcd testcd, '' rstcd, '' rstunit, " & _
              "       '' hldiv, '' dpdiv, x.spccd, '' rstdiv, x.statfg, '' lastrst, 0 lastvfyid, x.examdt VfyDt, " & _
              "       x.examtm VfyTm, x.examdoct VfyId, '4' as testdiv, '' footnotefg, '' rmkcd, " & _
              "       '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, j.spcnm SpcNm, '' as SenFg, " & _
              "       x.dcfg, x.mesg, x.ptid, x.orddt, x.ordno, x.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, " & _
              "       b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, '' DetailItem, " & _
              "       c.deptcd, c.sex, 0 ageday, k.orddoct, c.colid, c.rcvid, " & _
              "       substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
              "       '' as LstVfyDtTm, " & _
              "       substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
              "       substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
              "       substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
              "       0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, 0 as RptSeq, " & tmpStr & _
              "From " & T_LAB102 & " x, " & T_APS001 & " b, " & T_APS201 & " c, " & T_LAB101 & " k, " & T_APS002 & " j " & _
              "Where  x.ptid = " & PtId & " " & _
              "and    x." & QueryKey & " >= '" & FromDate & "' and  x." & QueryKey & "<= '" & ToDate & "' " & _
              "and    k.ptid = x.ptid  and  k.orddt = x.orddt  and  k.ordno = x.ordno " & _
              "and    b.testcd = x.ordcd  " & _
              "and    c.workarea = x.workarea  and  c.accdt = x.accdt    and    c.accseq = x.accseq " & _
              "and    j.spccd =* x.spccd  "
                  '"and      e.empid =* a.lastvfyid   and   f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
     Result5 = "Select x.workarea as WorkArea, x.accdt as AccDt, x.accseq as AccSeq, x.ordcd testcd, '' rstcd, '' rstunit, " & _
              "       '' hldiv, '' dpdiv, '' spccd, '' rstdiv, x.statfg, '' lastrst, 0 lastvfyid, x.examdt VfyDt, " & _
              "       x.examtm VfyTm, x.examdoct VfyId, '5' as testdiv, '' footnotefg, '' rmkcd, " & _
              "       '' mfyfg, '' txtfg, '' rsttype, '' detailfg, '' eqpcd, '혈액' SpcNm, '' as SenFg, " & _
              "       x.dcfg, x.mesg, x.ptid, x.orddt, x.ordno, x.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, " & _
              "       b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, '' DetailItem, " & _
              "       k.deptcd, '' sex, 0 ageday, k.orddoct, c.colid, c.rcvid, " & _
              "       substring(k.orddt ,3,2)+'-'+substring(k.orddt,5,2)+'-'+substring(k.orddt,7,2)+' '+substring(k.ordtm,1,2)+':'+substring(k.ordtm,3,2) as OrdDate, " & _
              "       '' as LstVfyDtTm, " & _
              "       substring(c.coldt,3,2)+'-'+substring(c.coldt,5,2)+'-'+substring(c.coldt,7,2)+' '+substring(c.coltm,1,2)+':'+substring(c.coltm,3,2) as ColDtTm, " & _
              "       substring(c.rcvdt,3,2)+'-'+substring(c.rcvdt,5,2)+'-'+substring(c.rcvdt,7,2)+' '+substring(c.rcvtm,1,2)+':'+substring(c.rcvtm,3,2) as RcvDtTm, " & _
              "       substring(x.examdt,3,2)+'-'+substring(x.examdt,5,2)+'-'+substring(x.examdt,7,2)+' '+substring(x.examtm,1,2)+':'+substring(x.examtm,3,2) as VfyDtTm, " & _
              "       x.unitqty, k.reqdt, k.reqtm, '' as RstCdNm, 0 as RptSeq, " & tmpStr & _
              "From " & T_LAB102 & " x, " & T_BBS001 & " b, " & T_BBS201 & " c, " & T_LAB101 & " k " & _
              "Where  x.ptid = " & PtId & " " & _
              "and    x." & QueryKey & " >= '" & FromDate & "' and  x." & QueryKey & "<= '" & ToDate & "' " & _
              "and    c.ptid = x.ptid " & _
              "and    c.spcyy = (select max(spcyy) from " & T_BBS201 & " where ptid = x.ptid) " & _
              "and    c.spcno = (select max(spcno) from " & T_BBS201 & _
                               " where ptid = c.ptid and spcyy = c.spcyy) " & _
              "and    k.ptid = x.ptid  and  k.orddt = x.orddt  and  k.ordno = x.ordno " & _
              "and    b.testcd = x.ordcd  and  b.applydt = (select max(applydt) from " & T_BBS001 & " where testcd = x.ordcd and applydt <= x.orddt ) "
                  '"and      e.empid =* a.lastvfyid   and   f.empid =* a.vfyid   and    g.empl_no =* k.orddoct   and    h.empid =* c.colid   and    i.empid =* c.rcvid "
      
    Select Case TestDiv
       Case "0":  SqlQueryAllResults = Result1 & tmpOrderby   '일반결과만..
       Case "1":  SqlQueryAllResults = Result2 & tmpOrderby   '특수검사결과만..
       Case "2":  SqlQueryAllResults = Result3 & tmpOrderby   '미생물결과만..
       Case "3":  SqlQueryAllResults = Result1 & "Union all " & Result2 & "Union all " & Result3 & "Union all " & _
                                       Result4 & "Union all " & Result5 & tmpOrderby   '모두...
    End Select
                           
End Function

Public Function SqlResultsForCmt(ByVal PtId As String, ByVal FromDate As String, ByVal ToDate As String) As String

    SqlResultsForCmt = "Select  a.workarea, a.accdt, a.accseq, a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, " & _
                        "        a.orddt, a.vfydt, a.rstdiv, a.detailfg, j.field3 as SpcNm, b.testnm as TestLongNm, " & _
                        "        substring(a.vfydt,3,2)+'-'+substring(a.vfydt,5,2)+'-'+substring(a.vfydt,7,2)+' '+substring(a.vfytm,1,2)+':'+substring(a.vfytm,3,2) as VfyDtTm, " & _
                        "        d.field1 as RstCdNm, convert(int,b.rptseq) as RptSeq " & _
                        "From    " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB031 & " d, " & T_LAB032 & " j " & _
                        "Where   a.ptid = " & PtId & " " & _
                        "and     a.orddt >= '" & FromDate & "' and  a.orddt<= '" & ToDate & "' " & _
                        "and    ( a.hldiv <> '' or a.dpdiv <> '' ) " & _
                        "and     a.vfydt <> '' " & _
                        "and     b.testcd = a.testcd  " & _
                        "and     b.applydt = (select max(applydt) from " & T_LAB001 & _
                        "                     where testcd = a.testcd and applydt <= a.orddt ) " & _
                        "and     d.cdindex = '" & CD1_ItemResult & "'   and    d.cdval1 =* a.testcd    and    d.cdval2 =* a.rstcd " & _
                        "and     j.cdindex = '" & CD2_Specimen & "'   and    j.cdval1 =* a.spccd  " & _
                        "order by a.workarea, a.vfydt, RptSeq, a.testcd "
      
End Function


'%  47. LD, CPK 등 관련검사의 결과 조회
'%       - Calling from [ frm202AccDataEntry, frm204WSDataEntry ] :
Public Function SqlGetRelTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String) As String
      
      SqlGetRelTest = "select a.testcd, b.abbrnm5, a.rstcd, a.rstunit, c.empnm as VfyNm, " & _
                             "          substring(a.vfydt,3,2)+'-'+substring(a.vfydt,5,2)+'-'+substring(a.vfydt,7,2)+' '+" & _
                             "          substring(a.vfytm,1,2)+':'+substring(a.vfytm,3,2) as VfyDtTm " & _
                             "from " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB015 & " c, " & T_LAB302 & " d, " & T_LAB031 & " e " & _
                             "where d.workarea = '" & WorkArea & "' " & _
                             "and    d.accdt = '" & AccDt & "' " & _
                             "and    d.accseq = " & AccSeq & " " & _
                             "and    e.cdindex = '" & CD1_RelTest & "' " & _
                             "and    e.cdval1 = d.testcd " & _
                             "and    a.ptid = d.ptid " & _
                             "and    a.testcd = e.cdval2 " & _
                             "and    a.vfydt <> '' " & _
                             "and    a.vfydt+a.vfytm = (select max(vfydt+vfytm) from " & T_LAB302 & " where ptid = a.ptid and testcd = a.testcd) " & _
                             "and    c.empid =* a.vfyid " & _
                             "and    b.testcd = a.testcd " & _
                             "and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) "
               
End Function


'%  48. 접수번호로 Status Check
'%       - Calling from [ frm108AccCancel ] :
Public Function SqlCheckStatus(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String) As String
      
      SqlCheckStatus = "select a.stscd, a.ptid, a.orddoct, a.deptcd, a.wardid, a.roomid, a.bedid, a.hosilid, a.spccd, " & _
                             "          a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, a.multifg, b.field3 as SpcNm " & _
                             "from " & T_LAB201 & " a, " & T_LAB032 & " b " & _
                             "where a.workarea = '" & WorkArea & "' " & _
                             "and    a.accdt = '" & AccDt & "' " & _
                             "and    a.accseq = " & AccSeq & " " & _
                             "and    b.cdindex = '" & CD2_Specimen & "' " & _
                             "and    b.cdval1 = a.spccd "
               
End Function


'%  49. 접수번호로 결과내역 검색
'%       - Calling from [ frm108AccCancel ] :
Public Function SqlAccOrder(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal StsCd As String, Optional ByVal MultiFg As String = "") As String
      
      If StsCd > "1" Then
            SqlAccOrder = "select a.ptid, a.orddt, a.ordno, a.ordseq, a.testcd, a.vfydt, b.testnm " & _
                                   "from " & T_LAB302 & " a, " & T_LAB001 & " b " & _
                                   "where a.workarea = '" & WorkArea & "' " & _
                                   "and    a.accdt = '" & AccDt & "' " & _
                                   "and    a.accseq = " & AccSeq & " " & _
                                   "and    (a.detailfg = ''    or    a.rstdiv = '*') " & _
                                   "and    b.testcd = a.testcd "
            SqlAccOrder = SqlAccOrder & "UNION ALL " & _
                                   "select a.ptid, a.orddt, a.ordno, a.ordseq, a.testcd, a.vfydt, b.testnm " & _
                                   "from " & T_LAB351 & " a, " & T_LAB001 & " b " & _
                                   "where a.workarea = '" & WorkArea & "' " & _
                                   "and    a.accdt = '" & AccDt & "' " & _
                                   "and    a.accseq = " & AccSeq & " " & _
                                   "and    b.testcd = a.testcd "
            SqlAccOrder = SqlAccOrder & "UNION ALL " & _
                                   "select a.ptid, a.orddt, a.ordno, a.ordseq, a.testcd, a.vfydt, b.testnm " & _
                                   "from " & T_LAB404 & " a, " & T_LAB001 & " b " & _
                                   "where a.workarea = '" & WorkArea & "' " & _
                                   "and    a.accdt = '" & AccDt & "' " & _
                                   "and    a.accseq = " & AccSeq & " " & _
                                   "and    b.testcd = a.testcd "
            SqlAccOrder = SqlAccOrder & "ORDER BY orddt, ordno, ordseq"
    Else
        If MultiFg = "" Then
            SqlAccOrder = "select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, '' as vfydt, b.testnm " & _
                                   "from " & T_LAB102 & " a, " & T_LAB001 & " b " & _
                                   "where a.workarea = '" & WorkArea & "' " & _
                                   "and    a.accdt = '" & AccDt & "' " & _
                                   "and    a.accseq = " & AccSeq & " " & _
                                   "and    b.testcd = a.ordcd " & _
                                   "ORDER BY orddt, ordno, ordseq"
        Else
            SqlAccOrder = "select a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, '' as vfydt, b.testnm " & _
                                   "from " & T_LAB102 & " a, " & T_LAB001 & " b, " & T_LAB203 & " c " & _
                                   "where  c.workarea = '" & WorkArea & "' " & _
                                   "and    c.accdt = '" & AccDt & "' " & _
                                   "and    c.accseq = " & AccSeq & " " & _
                                   "and    a.ptid = c.ptid  and a.orddt = c.orddt and a.ordno = c.ordno and a.ordseq = c.ordseq " & _
                                   "and    b.testcd = a.ordcd " & _
                                   "ORDER BY orddt, ordno, ordseq"
        End If
    End If
               
End Function


Public Function SqlDelRstTable(ByVal pTblNm As String, ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String

    Dim strTblNm As String
    
    SqlDelRstTable = "delete from " & pTblNm & " " & _
                             "where workarea = '" & pWorkArea & "' " & _
                             "and    accdt = '" & pAccDt & "' " & _
                             "and    accseq = " & pAccSeq
    
End Function



Public Function SqlStaticItem() As String
    
    SqlStaticItem = "SELECT a.cdval1 as testcd, convert(int,a.field1) as seq, b.testnm, b.workarea, a.field2 as wanm " & _
                    "FROM " & T_LAB032 & " a, " & T_LAB001 & " b " & _
                    "WHERE a.cdindex = '" & CD2_StaticItem & "' " & _
                    "AND   b.testcd = a.cdval1 " & _
                    "AND   b.applydt = (select max(applydt) from " & T_LAB001 & _
                    "                   where  testcd=b.testcd) " & _
                    "ORDER BY seq"

End Function



Public Function SqlTestCount(ByVal intTable As Integer, ByVal strColDt As String, _
                             ByVal strFrDt As String, ByVal strToDt As String, ByVal BldCd As String) As String
    
    Dim strTblNm As String
    
    '"and   c.stscd > '" & STS_Access & "' "    '검사수행된것만...
    '** 임상병리과(CL01)에서 낸 처방은 제외..

    strTblNm = Choose(intTable, T_LAB302, T_LAB351, T_LAB404)
    SqlTestCount = "select  b.groupcd, a.field1 as WAnm, b.itemseq,b.testcd, b.testnm, c.rcvdt, count(d.testcd) as Cnt " & _
                    "from " & T_LAB032 & " a, " & T_LAB001 & " b, " & T_LAB201 & " c, " & strTblNm & " d " & _
                    "where c.coldt>='" & strColDt & "' " & _
                    "and   c.coldt<='" & strToDt & "' " & _
                    "and   c.rcvdt>='" & strFrDt & "' " & _
                    "and   c.rcvdt<='" & strToDt & "' " & _
                    "and   c.buildcd = '" & BldCd & "' " & _
                    "and   c.deptcd <> 'CL01' " & _
                    "and   d.workarea=c.workarea " & _
                    "and   d.accdt=c.accdt " & _
                    "and   d.accseq=c.accseq " & _
                    "and   b.testcd=d.testcd " & _
                    "and   b.detailfg='' " & _
                    "and   b.applydt=(select max(applydt) from " & T_LAB001 & " " & _
                                    "where testcd=b.testcd) " & _
                    "and   a.cdindex='" & CD2_StaticGroup & "' " & _
                    "and   a.cdval1=*b.workarea " & _
                    "group by b.groupcd, a.field1, b.itemseq, b.testcd, b.testnm, c.rcvdt " & _
                    "having count(*) > 0 " & _
                    "order by b.groupcd, b.itemseq, b.testcd, c.rcvdt"

End Function

Public Function SqlGetCumList(ByVal pDeptCd As String) As String
    SqlGetCumList = "select distinct cdval1 as CumCd, field1 as CumNm, field3 as PWD " & _
                    "from " & T_LAB031 & " where cdindex = '" & CD1_CumItem & "' "
    If pDeptCd <> "ALL" Then SqlGetCumList = SqlGetCumList & "and  field2 = '" & pDeptCd & "' "
    SqlGetCumList = SqlGetCumList & " order by CumCd"
End Function

Public Function SqlGetCumItem(ByVal pCumCd As String, Optional ByVal iCase As Integer = 0) As String
    If iCase = 0 Then
        SqlGetCumItem = "select a.cdval2 as TestCd, a.field2 as DeptCd, a.field3 as PWD, convert(int,a.field4) as RptSeq, a.field5, b.testnm, b.panelfg, b.testdiv " & _
                        "from " & T_LAB031 & " a, " & T_LAB001 & " b " & _
                        "where a.cdindex = '" & CD1_CumItem & "' " & _
                        "and   a.cdval1 = '" & pCumCd & "' " & _
                        "and   b.testcd = a.cdval2 " & _
                        "and   b.applydt = (select max(applydt) from " & T_LAB001 & _
                                          " where testcd = b.testcd) " & _
                        "order by RptSeq"
    Else
        SqlGetCumItem = "select a.cdval2 as TestCd, convert(int, a.field4) as Seq, 1 as Seq1, " & _
                               "a.field5 as SpcCd, b.testnm, b.panelfg, b.testdiv, b.workarea, c.field3 as SpcNm " & _
                        "from " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB032 & " c " & _
                        "where a.cdindex = '" & CD1_CumItem & "' " & _
                        "and   a.cdval1 = '" & pCumCd & "' " & _
                        "and   b.testcd = a.cdval2 " & _
                        "and   b.applydt = (select max(applydt) from " & T_LAB001 & _
                                          " where testcd = b.testcd) " & _
                        "and   b.panelfg != '" & PN_Group & "' " & _
                        "and   c.cdindex = '" & CD2_Specimen & "' " & _
                        "and   c.cdval1 = a.field5 "
        SqlGetCumItem = SqlGetCumItem & "Union All " & _
                        "select c.field1 as TestCd, convert(int, a.field4) as Seq, d.rptseq as Seq1, " & _
                               "a.field5 as SpcCd, d.testnm, d.panelfg, d.testdiv, d.workarea, e.field3 as SpcNm " & _
                        "from " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB031 & " c, " & _
                                  T_LAB001 & " d, " & T_LAB032 & " e " & _
                        "where a.cdindex = '" & CD1_CumItem & "' " & _
                        "and   a.cdval1 = '" & pCumCd & "' " & _
                        "and   b.testcd = a.cdval2 " & _
                        "and   b.applydt = (select max(applydt) from " & T_LAB001 & _
                                          " where testcd = b.testcd) " & _
                        "and   b.panelfg = '" & PN_Group & "' " & _
                        "and   c.cdindex = '" & CD1_Panel & "' " & _
                        "and   c.cdval1 = b.testcd " & _
                        "and   d.testcd = c.field1 " & _
                        "and   d.applydt = (select max(applydt) from " & T_LAB001 & _
                                          " where testcd = d.testcd) " & _
                        "and   e.cdindex = '" & CD2_Specimen & "' " & _
                        "and   e.cdval1 = a.field5 " & _
                        "and   e.field1 = '' "
        SqlGetCumItem = SqlGetCumItem & "Union All " & _
                        "select c.field1 as TestCd, convert(int, a.field4)  as Seq, d.rptseq as Seq1, " & _
                               "f.field1 as SpcCd, d.testnm, d.panelfg, d.testdiv, d.workarea, g.field3 as SpcNm " & _
                        "from " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB031 & " c, " & _
                                  T_LAB001 & " d, " & T_LAB032 & " e, " & T_LAB031 & " f, " & T_LAB032 & " g " & _
                        "where a.cdindex = '" & CD1_CumItem & "' " & _
                        "and   a.cdval1 = '" & pCumCd & "' " & _
                        "and   b.testcd = a.cdval2 " & _
                        "and   b.applydt = (select max(applydt) from " & T_LAB001 & _
                                          " where testcd = b.testcd) " & _
                        "and   b.panelfg = '" & PN_Group & "' " & _
                        "and   c.cdindex = '" & CD1_Panel & "' " & _
                        "and   c.cdval1 = b.testcd " & _
                        "and   d.testcd = c.field1 " & _
                        "and   d.applydt = (select max(applydt) from " & T_LAB001 & _
                                          " where testcd = d.testcd) " & _
                        "and   e.cdindex = '" & CD2_Specimen & "' " & _
                        "and   e.cdval1 = a.field5 " & _
                        "and   e.field1 = 'Y' " & _
                        "and   f.cdindex = '" & CD1_MultiSpc & "' " & _
                        "and   f.cdval1 = a.field5 " & _
                        "and   f.cdval2 = c.cdval2 " & _
                        "and   g.cdindex = '" & CD2_Specimen & "' " & _
                        "and   g.cdval1 = f.field1 " & _
                        "order by Seq, Seq1, TestCd"
                        
    End If
    
End Function

Public Function SqlGetCumDetail(ByVal pTestCd As String) As String

    SqlGetCumDetail = "select a.cdval2 as TestCd, b.testnm, b.workarea, b.panelfg, b.testdiv, convert(int, b.rptseq) as RptSeq " & _
                      "from " & T_LAB031 & " a, " & T_LAB001 & " b " & _
                      "where a.cdindex = '" & CD1_Detail & "' " & _
                      "and   a.cdval1 = '" & pTestCd & "' " & _
                      "and   b.testcd = a.cdval2 " & _
                      "and   b.applydt = (select max(applydt) from " & T_LAB001 & _
                                        " where testcd = b.testcd) " & _
                      "order by RptSeq, TestCd"
End Function

Public Function SqlCumResult(ByVal pPtId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                             ByVal pTestCd As String, ByVal pPanelFg As String, ByVal pTestDiv As String, _
                             ByVal pSpcCd As String, ByVal pWorkArea As String)
    Dim strTblNm As String
    Dim strRcvDt As String
    
    strRcvDt = Format(DateAdd("d", -2, Format(pToDt, CS_DateLMask)), CS_DateDbFormat)
    strTblNm = Choose(Val(pTestDiv) + 1, T_LAB302, T_LAB351, T_LAB405)
    
    Select Case pPanelFg
    Case PN_Detail:
        SqlCumResult = " "
    Case "":
        SqlCumResult = "Select  b.coldt, substring(b.coltm, 1, 4) as ColTm, " & _
                               "a.testcd, a.rstcd, a.rstunit, a.hldiv, a.dpdiv, a.spccd, a.statfg, " & _
                               "a.workarea, a.accdt, a.accseq, b.footnotefg, b.rmkcd, " & _
                               "b.deptcd, b.wardid, b.hosilid, " & _
                               "a.rstdiv, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, " & _
                               "c.rsttxt as TextResult, d.field1 as RstCdNm " & _
                       "From " & T_LAB302 & " a, " & T_LAB201 & " b, " & _
                                 T_LAB303 & " c, " & T_LAB031 & " d "
        SqlCumResult = SqlCumResult & _
                       "Where   a.ptid = " & pPtId & " " & _
                       "and     a.testcd = '" & pTestCd & "' " & _
                       "and     a.spccd = '" & pSpcCd & "' " & _
                       "and     b.workarea = a.workarea " & _
                       "and     b.accdt = a.accdt " & _
                       "and     b.accseq = a.accseq " & _
                       "and     b.rcvdt >= '" & pFromDt & "' " & _
                       "and     c.workarea =* a.workarea " & _
                       "and     c.accdt =* a.accdt " & _
                       "and     c.accseq =* a.accseq " & _
                       "and     c.testcd =* a.testcd " & _
                       "and     d.cdindex = '" & CD1_ItemResult & "' " & _
                       "and     d.cdval1 =* a.testcd " & _
                       "and     d.cdval2 =* a.rstcd "
                       '"and     b.rcvdt <= '" & strRcvDt & "' " & _
                       "and     b.coldt <= '" & pToDt & "' "
    End Select
    
End Function

Public Function SqlCumOldResult(ByVal pPtId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                             ByVal pTestCd As String, ByVal pPanelFg As String, ByVal pTestDiv As String, _
                             ByVal pSpcCd As String, ByVal pWorkArea As String)
    Dim strTblNm As String
    Dim strRcvDt As String
    
    strRcvDt = Format(DateAdd("d", -2, Format(pToDt, CS_DateLMask)), CS_DateDbFormat)
    strTblNm = Choose(Val(pTestDiv) + 1, T_LAB302, T_LAB351, T_LAB405)
    
    Select Case pPanelFg
    Case PN_Detail:
        SqlCumOldResult = " "
    Case "":
        SqlCumOldResult = "Select  a.vfydt as ColDt, substring(a.vfytm, 1, 4) as ColTm, " & _
                               "a.testcd, a.rstcd, a.rstunit, a.hidiv as hldiv, '' as dpdiv, a.spccd, '' as statfg, " & _
                               "a.workarea, '' as accdt, 0 as accseq, a.mesg as footnote, '' as rmkcd, " & _
                               "a.deptcd, '' as wardid, '' as hosilid, " & _
                               "'' as rstdiv, '' as txtfg, '' as rsttype, '' as detailfg, '' as eqpcd, " & _
                               "'' as TextResult, '' as RstCdNm " & _
                       "From " & TB_LAB999 & " a "
        SqlCumOldResult = SqlCumOldResult & _
                       "Where   a.ptid = " & pPtId & " " & _
                       "and     a.testcd = '" & pTestCd & "' " & _
                       "and     a.spccd = '" & pSpcCd & "' " & _
                       "and     a.vfydt >= '" & pFromDt & "' "
        SqlCumOldResult = SqlCumOldResult & " UNION ALL " & _
                       "Select  a.vfydt as ColDt, substring(a.vfytm, 1, 4) as ColTm, " & _
                               "rtrim(a.testcd) + rtrim(a.testseq) as testcd, a.rstcd, a.rstunit, a.hidiv as hldiv, '' as dpdiv, a.spccd, '' as statfg, " & _
                               "a.workarea, '' as accdt, 0 as accseq, a.mesg as footnote, '' as rmkcd, " & _
                               "a.deptcd, '' as wardid, '' as hosilid, " & _
                               "'' as rstdiv, '' as txtfg, '' as rsttype, '' as detailfg, '' as eqpcd, " & _
                               "'' as TextResult, '' as RstCdNm " & _
                       "From " & TB_LAB999 & " a "
        SqlCumOldResult = SqlCumOldResult & _
                       "Where   a.ptid = " & pPtId & " " & _
                       "and     rtrim(a.testcd) + rtrim(a.testseq) = '" & pTestCd & "' " & _
                       "and     a.spccd = '" & pSpcCd & "' " & _
                       "and     a.vfydt >= '" & pFromDt & "' "
    End Select
    
End Function



Public Function SqlWardBarReprint(ByVal pWorkDt As String, ByVal pWorkTm As String, ByVal pWardId As String) As String

    SqlWardBarReprint = "Select x.seq, c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                       "       a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                       "       a.deptcd, a.wardid, a.hosilid, d.field5 as SpcNm, " & _
                       "       a.workarea+'-'+substring(a.accdt,3,6)+'-'+convert(char,a.accseq) as LabNo, " & _
                       "       substring(a.coldt,5,2)+'/'+substring(a.coldt,7,2) as ColDt, " & _
                       "       b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                       "       f.ptnt_nm as PtNm, g.reqdt, g.reqtm, g.orddiv, '' fzfg " & _
                       "From " & T_LAB204 & " x, " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & _
                                 T_LAB032 & " d, " & T_HIS001 & " f, " & T_LAB101 & " g " & _
                       "Where  x.workdt = '" & pWorkDt & "' and x.wardid = '" & pWardId & "' and x.worktm = '" & pWorkTm & "' " & _
                       "and    a.workarea = x.workarea and a.accdt = x.accdt and a.accseq = x.accseq " & _
                       "and    a.multifg = '' " & _
                       "and    b.workarea = a.workarea  and    b.accdt = a.accdt " & _
                       "and    b.accseq = a.accseq      and    b.dcfg = '' " & _
                       "and    c.testcd = b.ordcd " & _
                       "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                       "and    d.cdindex = '" & CD2_Specimen & "'    and   d.cdval1 = a.spccd " & _
                       "and    g.ptid = b.ptid " & _
                       "and    g.orddt = b.orddt " & _
                       "and    g.ordno = b.ordno " & _
                       "and    f.ptnt_no = a.ptid "
   '
    SqlWardBarReprint = SqlWardBarReprint & _
                       "Union All " & _
                       "Select x.seq, c.testnm, c.abbrnm5, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                       "       a.spcyy, a.spcno, a.spccd, a.storecd, a.stscd, a.statfg, a.buildcd, " & _
                       "       a.deptcd, a.wardid, a.hosilid, d.field5 as SpcNm, " & _
                       "       a.workarea+'-'+substring(a.accdt,3,6)+'-'+convert(char,a.accseq) as LabNo, " & _
                       "       substring(a.coldt,5,2)+'/'+substring(a.coldt,7,2) as ColDt, " & _
                       "       b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                       "       h.ptnt_nm as PtNm, i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                       "From " & T_LAB204 & " x, " & T_LAB201 & " a, " & T_LAB102 & " b,  " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                                 T_LAB203 & " e, " & T_LAB031 & " f, " & T_HIS001 & " h, " & T_LAB101 & " i " & _
                       "Where  x.workdt = '" & pWorkDt & "' and x.wardid = '" & pWardId & "' and x.worktm = '" & pWorkTm & "' " & _
                       "and    a.workarea = x.workarea and a.accdt = x.accdt and a.accseq = x.accseq " & _
                       "and    e.workarea = a.workarea and e.accdt = a.accdt and e.accseq = a.accseq and e.ptid = a.ptid " & _
                       "and    b.ptid = e.ptid and b.orddt = e.orddt and b.ordno = e.ordno and b.ordseq = e.ordseq   and    b.dcfg = '' " & _
                       "and    a.multifg <> '' " & _
                       "and    f.cdindex = '" & CD1_Panel & "' " & _
                       "and    f.cdval1 = b.ordcd and    f.cdval2 = e.spcseq " & _
                       "and    c.testcd = f.field1 " & _
                       "and    c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd and applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                       "and    d.cdindex = '" & CD2_Specimen & "' and    d.cdval1 = a.spccd " & _
                       "and    i.ptid = b.ptid   and   i.orddt = b.orddt  and  i.ordno = b.ordno " & _
                       "and    h.ptnt_no = a.ptid "
    '해부병리
    SqlWardBarReprint = SqlWardBarReprint & _
                       "Union All " & _
                       "Select x.seq, c.testnm, c.abbrnm5, g.ptid, x.workarea, x.accdt, x.accseq, " & _
                       "       g.spcyy, g.spcno, g.spccd, '' storecd, b.stscd, b.statfg, '" & APS_BLDCD & "' as buildcd, " & _
                       "       a.deptcd, a.wardid, a.hosilid, d.spcnm5 as SpcNm, " & _
                       "       x.spcyy+'-'+convert(char,x.spcno) as LabNo, " & _
                       "       substring(g.coldt,5,2)+'/'+substring(g.coldt,7,2) as ColDt, " & _
                       "       b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                       "       f.ptnt_nm as PtNm, g.coldt reqdt, g.coltm reqtm, '" & APS_ORDDIV & "' orddiv, c.fzfg " & _
                       "From " & T_LAB204 & " x, " & T_LAB101 & " a, " & T_LAB102 & " b,  " & T_APS001 & " c, " & _
                                 T_APS002 & " d, " & T_HIS001 & " f, " & T_APS103 & " g " & _
                       "Where  x.workdt = '" & pWorkDt & "' and x.wardid = '" & pWardId & "' and x.worktm = '" & pWorkTm & "' " & _
                       "and    g.spcyy = x.spcyy        and    g.spcno = x.spcno " & _
                       "and    b.ptid = g.ptid " & _
                       "and    b.orddt = g.orddt " & _
                       "and    b.ordno = g.ordno " & _
                       "and    b.ordseq = g.ordseq " & _
                       "and    b.dcfg = '' " & _
                       "and    a.ptid = b.ptid   and   a.orddt = b.orddt  and  a.ordno = b.ordno " & _
                       "and    c.testcd = b.ordcd " & _
                       "and    d.spccd = g.spccd " & _
                       "and    f.ptnt_no = a.ptid "
   '혈액은행
   SqlWardBarReprint = SqlWardBarReprint & _
                       "Union All " & _
                       "Select x.seq, c.testnm, c.abbrnm5, a.ptid, '' workarea, '' accdt, 0 accseq, " & _
                       "       a.spcyy, a.spcno, '' spccd, '' storecd, '' stscd, '' statfg, a.buildcd, " & _
                       "       g.deptcd, g.wardid, g.hosilid, '혈액' as SpcNm, " & _
                       "       a.spcyy+'-'+convert(char,a.spcno) as LabNo, " & _
                       "       substring(a.coldt,5,2)+'/'+substring(a.coldt,7,2) as ColDt, " & _
                       "       b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                       "       f.ptnt_nm as PtNm, g.reqdt, g.reqtm, g.orddiv, '' fzfg " & _
                       "From " & T_LAB204 & " x, " & T_BBS201 & " a, " & T_LAB102 & " b,  " & _
                                 T_BBS001 & " c, " & T_HIS001 & " f, " & T_LAB101 & " g " & _
                       "Where  x.workdt = '" & pWorkDt & "' and x.wardid = '" & pWardId & "' and x.worktm = '" & pWorkTm & "' " & _
                       "and    a.spcyy = x.spcyy     and    a.spcno = x.spcno " & _
                       "and    g.ptid = a.ptid " & _
                       "and    g.orddiv = '" & C_WORKAREA & "' " & _
                       "and    g.orddt = (select max(orddt) from " & T_LAB101 & _
                                        " where ptid = g.ptid  and  orddiv = g.orddiv) " & _
                       "and    b.ptid = g.ptid    and    b.orddt = g.orddt " & _
                       "and    b.ordno = g.ordno  " & _
                       "and    b.dcfg = '' " & _
                       "and    c.testcd = b.ordcd " & _
                       "and    c.testdiv <> '3' " & _
                       "and    f.ptnt_no = a.ptid "
    
    SqlWardBarReprint = SqlWardBarReprint & " Order  By x.seq, LabNo, b.ordcd "
    
End Function
   

Public Function SqlGsOrder(ByVal ReqDt As String, ByVal ReqTm As String, ByVal DeptCd As String) As String
   
   SqlGsOrder = "Select b.deptcd as DeptCd, a.ptnt_no as PtId, a.ptnt_nm as PtNm, substring(a.ptnt_prsn_no,9,1) as Sex, c.statfg, c.spccd, " & _
                "          substring(a.ptnt_prsn_no,1,8) as Dob, b.bedindt, b.orddoct, b.majdoct, b.roomid, '' as bedid, b.hosilid " & _
                "From   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c " & _
                "Where  b.donefg = '0' " & _
                "and    b.reqdt = '" & ReqDt & "' " & _
                "and    b.deptcd in (" & DeptCd & ") " & _
                "and    b.bussdiv = '" & CS_BussOut & "' " & _
                "and    a.ptnt_no = b.ptid " & _
                "and    c.ptid = b.ptid    and   c.orddt = b.orddt  and   c.ordno = b.ordno  and    c.stscd = '0'  and   c.dcfg = '' " & _
                "and    c.ordcd = 'B400'  and c.spccd = '75' " & _
                "Order By DeptCd, PtId, c.statfg, c.spccd "
End Function


Public Function SqlInstrument()

    SqlInstrument = "SELECT a.eqpcd,a.eqpnm,a.remark" & _
      " FROM " & T_LAB006 & " a WHERE a.eqpdiv = 'E'" & _
      " ORDER BY eqpnm"

End Function
