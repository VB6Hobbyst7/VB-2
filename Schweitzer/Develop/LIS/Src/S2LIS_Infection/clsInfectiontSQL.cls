VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfectiontSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'%  Infection Test List
'%      - intDiv : 1 -  (의뢰검체 리스트), 2 - (균 리스트)
Public Function SqlItemList(ByVal intDiv As Integer) As String
   If intDiv = 1 Then
      SqlItemList = "Select cdval1, field1 " & _
                    "  From " & TB_LAB032 & _
                    " where cdindex = '" & CD2_OrdSpc & "' "
                    
   Else
      SqlItemList = "Select cdval1, field1 " & _
                    "  From " & TB_LAB032 & _
                    " where cdindex = '" & CD2_Micro & "' "
   End If
End Function

'%  Infection Specimen List
Public Function Spc_List() As String
                                            
    Spc_List = " select * from " & TB_LAB032 & _
               "  where cdindex = " & DBS(CD2_OrdSpc)
              
End Function

'%  Infection Temp List
Public Function TestMeth_List(ByVal pValue As String) As String
                                            
    TestMeth_List = " select * from " & TB_LAB034 & _
                    "  where cdindex = " & DBS(pValue)
              
End Function

'%  Infection Temp Insert/Update Flag (False = Insert, True = Update)
Public Function Template_Insert_Update_Flag(ByVal pValue As String, ByVal pTCode As String) As Boolean
    Dim RS      As New ADODB.Recordset
    Dim strSql  As String
    
    strSql = " select * from " & TB_LAB034 & _
             "  where cdindex = " & DBS(pValue) & _
             "    and cdval1  = " & DBS(pTCode)
                  
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Template_Insert_Update_Flag = True
    Else
        Template_Insert_Update_Flag = False
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

'%  Infection Temp Insert
Public Function Template_Insert(ByVal pValue As String, ByVal pTCode As String, _
                                ByVal pText1 As String) As String
                                            
    Template_Insert = " insert into " & TB_LAB034 & _
                      " Values( " & DBS(pValue) & "," & DBS(pTCode) & "," & _
                      "''" & "," & DBS(pText1) & "," & "''" & ")"

End Function

'%  Infection Temp Update
Public Function Template_Update(ByVal pValue As String, ByVal pTCode As String, _
                                ByVal pText1 As String) As String
                                            
    Template_Update = " update " & TB_LAB034 & _
                      "    set text1   = " & DBS(pText1) & _
                      "  where cdindex = " & DBS(pValue) & _
                      "    and cdval1  = " & DBS(pTCode)
              
End Function

'%  Infection Temp Delete
Public Function Template_Delete(ByVal pValue As String, ByVal pTCode As String) As String
                                            
    Template_Delete = " delete from " & TB_LAB034 & _
                      "  where cdindex = " & DBS(pValue) & _
                      "    and cdval1  = " & DBS(pTCode)
              
End Function


'%  Infection Insert_Update Flag (S2LAB315) (INSERT : True, UPDATE : False)
Public Function LAB315_INSERT_UPDATE_Status(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                            ByVal pColDt As String, ByVal pColTm As String) As Boolean
    Dim RS     As New ADODB.Recordset
    Dim strSql As String
    
    strSql = " select * from " & TB_LAB315 & _
             "  where workdt = " & DBS(pWorkDt) & _
             "    and deptcd = " & DBS(pDeptCd) & _
             "    and coldt = " & DBS(pColDt) & _
             "    and coltm = " & DBS(pColTm)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        LAB315_INSERT_UPDATE_Status = True
    Else
        LAB315_INSERT_UPDATE_Status = False
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

'%  Infection Insert (S2LAB315)
Public Function LAB315_INSERT_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                  ByVal pColDt As String, ByVal pColTm As String, ByVal pTestMeth As String, _
                                  ByVal pWaterFg As String, ByVal pWaterArea As String, ByVal pRstVal As String, _
                                  ByVal pStsCd As String, ByVal pMvfyDt As String, ByVal pMvfyTm As String, _
                                  ByVal pMvfyID As String, ByVal pFvfyDt As String, ByVal pFvfyTm As String, _
                                  ByVal pFvfyID As String, ByVal pMfyDt As String, ByVal pMfyTm As String, _
                                  ByVal pMfyID As String, ByVal pRptFg As String, ByVal pRptTxt As String, _
                                  ByVal pRptDt As String, ByVal pRptTm As String, ByVal pRptID As String) As String

    LAB315_INSERT_SQL = "Insert into " & _
                        "" & TB_LAB315 & " (workdt, deptcd, coldt, coltm, testmeth, waterfg, " & _
                        "waterarea, rstval, stscd, mvfyid, mvfydt, mvfytm, fvfyid, " & _
                        "fvfydt, fvfytm, mfyid, mfydt, mfytm, rptfg, rpttxt, rptdt, rpttm, rptid) " & _
                        "Values  (" & DBS(pWorkDt) & "," & DBS(pDeptCd) & _
                        "," & DBS(pColDt) & "," & DBS(pColTm) & "," & DBS(pTestMeth) & _
                        "," & DBS(pWaterFg) & "," & DBS(pWaterArea) & "," & DBS(pRstVal) & _
                        "," & DBS(pStsCd) & "," & DBS(pMvfyID) & "," & DBS(pMvfyDt) & _
                        "," & DBS(pMvfyTm) & "," & DBS(pFvfyID) & "," & DBS(pFvfyDt) & _
                        "," & DBS(pFvfyTm) & "," & DBS(pMfyID) & "," & DBS(pMfyDt) & _
                        "," & DBS(pMfyTm) & "," & DBS(pRptFg) & "," & DBS(pRptTxt) & _
                        "," & DBS(pRptDt) & "," & DBS(pRptTm) & "," & DBS(pRptID) & ")"

End Function

'%  Infection update (S2LAB315) : Final Result
Public Function LAB315_Modify_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, ByVal pColDt As String, _
                                  ByVal pColTm As String, ByVal pTestMeth As String, _
                                  ByVal pWaterFg As String, ByVal pWaterArea As String, ByVal pRstVal As String, _
                                  ByVal pStsCd As String, ByVal pMfyDt As String, ByVal pMfyTm As String, _
                                  ByVal pMfyID As String, ByVal pRptTxt As String) As String
    
    LAB315_Modify_SQL = " update " & TB_LAB315 & _
                        "    set testmeth = " & DBS(pTestMeth) & _
                        "," & " waterfg = " & DBS(pWaterFg) & "," & " waterarea = " & DBS(pWaterArea) & _
                        "," & " rstval = " & DBS(pRstVal) & "," & " stscd = " & DBS(pStsCd) & _
                        "," & " mfyid = " & DBS(pMfyID) & "," & " mfydt = " & DBS(pMfyDt) & _
                        "," & " mfytm = " & DBS(pMfyTm) & "," & " rpttxt = " & DBS(pRptTxt) & _
                        "  where workdt = " & DBS(pWorkDt) & _
                        "    and deptcd = " & DBS(pDeptCd) & _
                        "    and coldt = " & DBS(pColDt) & _
                        "    and coltm = " & DBS(pColTm)
    
End Function

'%  Infection update (S2LAB315) : Final Result
Public Function LAB315_FUPDATE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, ByVal pColDt As String, _
                                   ByVal pColTm As String, ByVal pTestMeth As String, _
                                   ByVal pWaterFg As String, ByVal pWaterArea As String, ByVal pRstVal As String, _
                                   ByVal pStsCd As String, ByVal pFvfyDt As String, ByVal pFvfyTm As String, _
                                   ByVal pFvfyID As String, ByVal pRptTxt As String) As String
    
    LAB315_FUPDATE_SQL = " update " & TB_LAB315 & _
                         "    set testmeth = " & DBS(pTestMeth) & _
                         "," & " waterfg = " & DBS(pWaterFg) & "," & " waterarea = " & DBS(pWaterArea) & _
                         "," & " rstval = " & DBS(pRstVal) & "," & " stscd = " & DBS(pStsCd) & _
                         "," & " fvfyid = " & DBS(pFvfyID) & "," & " fvfydt = " & DBS(pFvfyDt) & _
                         "," & " fvfytm = " & DBS(pFvfyTm) & "," & " rpttxt = " & DBS(pRptTxt) & _
                         "  where workdt = " & DBS(pWorkDt) & _
                         "    and deptcd = " & DBS(pDeptCd) & _
                         "    and coldt = " & DBS(pColDt) & _
                         "    and coltm = " & DBS(pColTm)
    
End Function

'%  Infection update (S2LAB315) : Middle Result
Public Function LAB315_MUPDATE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, ByVal pColDt As String, _
                                   ByVal pColTm As String, ByVal pTestMeth As String, _
                                   ByVal pWaterFg As String, ByVal pWaterArea As String, ByVal pRstVal As String, _
                                   ByVal pStsCd As String, ByVal pMvfyDt As String, ByVal pMvfyTm As String, _
                                   ByVal pMvfyID As String, ByVal pRptTxt As String) As String
    
    LAB315_MUPDATE_SQL = " update " & TB_LAB315 & _
                         "    set testmeth = " & DBS(pTestMeth) & _
                         "," & " waterfg = " & DBS(pWaterFg) & "," & " waterarea = " & DBS(pWaterArea) & _
                         "," & " rstval = " & DBS(pRstVal) & "," & " stscd = " & DBS(pStsCd) & _
                         "," & " mvfyid = " & DBS(pMvfyID) & "," & " mvfydt = " & DBS(pMvfyDt) & _
                         "," & " mvfytm = " & DBS(pMvfyTm) & "," & " rpttxt = " & DBS(pRptTxt) & _
                         "  where workdt = " & DBS(pWorkDt) & _
                         "    and deptcd = " & DBS(pDeptCd) & _
                         "    and coldt = " & DBS(pColDt) & _
                         "    and coltm = " & DBS(pColTm)
    
End Function

'%  Infection update (S2LAB315) : Report
Public Function LAB315_RFUPDATE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, ByVal pColDt As String, _
                                    ByVal pColTm As String, ByVal pTestMeth As String, ByVal pRstCd As String, _
                                    ByVal pWaterFg As String, ByVal pWaterArea As String, ByVal pRstVal As String, _
                                    ByVal pStsCd As String, ByVal pRptFg As String, ByVal pRptTxt As String, _
                                    ByVal pRptDt As String, ByVal pRptTm As String, ByVal pRptID As String) As String
    
    LAB315_RFUPDATE_SQL = " update " & TB_LAB315 & _
                         "    set testmeth = " & DBS(pTestMeth) & _
                         "," & " rstcd = " & DBS(pRstCd) & "," & " waterfg = " & DBS(pWaterFg) & _
                         "," & " waterarea = " & DBS(pWaterArea) & "," & " stscd = " & DBS(pStsCd) & _
                         "," & " rptfg = " & DBS(pRptFg) & "," & " rpttxt = " & DBS(pRptTxt) & _
                         "," & " rptdt = " & DBS(pRptDt) & "," & " rpttm = " & DBS(pRptTm) & _
                         "," & " rptid = " & DBS(pRptID) & _
                         "  where workdt = " & DBS(pWorkDt) & _
                         "    and deptcd = " & DBS(pDeptCd) & _
                         "    and coldt = " & DBS(pColDt) & _
                         "    and coltm = " & DBS(pColTm)
    
End Function

'%  Infection Insert_Update Flag (S2LAB316) (INSERT : True, UPDATE : False)
Public Function LAB316_INSERT_UPDATE_Status(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                            ByVal pColDt As String, ByVal pColTm As String, _
                                            ByVal pSpcCd As String, ByVal pRstCd As String) As Boolean
    Dim RS     As New ADODB.Recordset
    Dim strSql As String
    
    strSql = " select * from " & TB_LAB316 & _
             "  where workdt = " & DBS(pWorkDt) & _
             "    and deptcd = " & DBS(pDeptCd) & _
             "    and coldt = " & DBS(pColDt) & _
             "    and coltm = " & DBS(pColTm) & _
             "    and spccd = " & DBS(pSpcCd) & _
             "    and rstcd = " & DBS(pRstCd)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        LAB316_INSERT_UPDATE_Status = True
    Else
        LAB316_INSERT_UPDATE_Status = False
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

'%  Infection Insert (S2LAB316)
Public Function LAB316_INSERT_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                  ByVal pColDt As String, ByVal pColTm As String, _
                                  ByVal pSpcCd As String, ByVal pRstCd As String, ByVal pRstCnt As String, _
                                  ByVal pRstTxt As String, ByVal pRstFg As String) As String

    LAB316_INSERT_SQL = "Insert into " & _
                        "" & TB_LAB316 & " (workdt, deptcd, coldt, coltm, " & _
                        "spccd, rstcd,rstcount, rsttxt, rstfg) " & _
                        "Values  (" & DBS(pWorkDt) & "," & DBS(pDeptCd) & "," & DBS(pColDt) & _
                        "," & DBS(pColTm) & "," & DBS(pSpcCd) & "," & DBS(pRstCd) & "," & DBS(pRstCnt) & _
                        "," & DBS(pRstTxt) & "," & DBS(pRstFg) & ")"

End Function

'%  Infection Update (S2LAB316)
Public Function LAB316_UPDATE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                  ByVal pColDt As String, ByVal pColTm As String, _
                                  ByVal pSpcCd As String, ByVal pRstCd As String, ByVal pRstCnt As String, _
                                  ByVal pRstTxt As String, ByVal pRstFg As String) As String
    
    LAB316_UPDATE_SQL = " update " & TB_LAB316 & _
                        "    set rstcount = " & DBS(pRstCnt) & _
                        "," & " rstcd = " & DBS(pRstCd) & _
                        "," & " rsttxt = " & DBS(pRstTxt) & _
                        "," & " rstfg = " & DBS(pRstFg) & _
                        "  where workdt = " & DBS(pWorkDt) & _
                        "    and deptcd = " & DBS(pDeptCd) & _
                        "    and coldt  = " & DBS(pColDt) & _
                        "    and coltm  = " & DBS(pColTm) & _
                        "    and spccd  = " & DBS(pSpcCd) & _
                        "    and rstcd  = " & DBS(pRstCd)
    
End Function

'%  Infection Delete (S2LAB316)
Public Function LAB316_DELETE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                  ByVal pColDt As String, ByVal pColTm As String) As String

    LAB316_DELETE_SQL = " delete from " & TB_LAB316 & _
                        "  where workdt = " & DBS(pWorkDt) & _
                        "    and deptcd = " & DBS(pDeptCd) & _
                        "    and coldt  = " & DBS(pColDt) & _
                        "    and coltm  = " & DBS(pColTm)

End Function

'%  (Mid/Final) Verify Check (S2LAB315)
Public Function Mid_Verify_Check(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                 ByVal pColDt As String, ByVal pColTm As String, _
                                 Optional ByVal pStatus As String) As Boolean
    
    Dim RS     As New ADODB.Recordset
    Dim strSql As String
    
    strSql = " select * from " & TB_LAB315 & _
             "  where workdt = " & DBS(pWorkDt) & _
             "    and deptcd = " & DBS(pDeptCd) & _
             "    and coldt  = " & DBS(pColDt) & _
             "    and coltm  = " & DBS(pColTm) & _
             "    and stscd  = " & DBS(pStatus)
             
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Mid_Verify_Check = True
    Else
        Mid_Verify_Check = False
    End If
                 
    RS.Close
    Set RS = Nothing
    
End Function

'%  (S2LAB315) DELETE
Public Function LAB315_DELETE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                  ByVal pColDt As String, ByVal pColTm As String, _
                                  Optional ByVal pStatus As String) As String
    
    
    LAB315_DELETE_SQL = " delete from " & TB_LAB315 & _
                        "  where workdt = " & DBS(pWorkDt) & _
                        "    and deptcd = " & DBS(pDeptCd) & _
                        "    and coldt  = " & DBS(pColDt) & _
                        "    and coltm  = " & DBS(pColTm)
    
End Function

'%  Middle Result Find (S2LAB315)
Public Function Mid_Result_Find(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                ByVal pFDate As String, ByVal pTDate As String, _
                                Optional ByVal pStsCd As String = "") As String
    
    Mid_Result_Find = " select workdt, deptcd, coldt, coltm, testmeth, waterfg, " & _
                      "        waterarea, rstval, mvfydt, mvfytm, mvfyid, rpttxt " & _
                      "   from " & TB_LAB315 & _
                      "  where mvfydt between " & DBS(pFDate) & " and " & DBS(pTDate) & _
                      "    and stscd = " & DBS(pStsCd) & _
                      "  order by coldt desc "
                      
End Function

'%  Middle Body Result Find (S2LAB316)
Public Function Mid_Body_Result_Find(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                                     ByVal pColDt As String, ByVal pColTm As String, _
                                     Optional ByVal pStatus As String) As String

    Mid_Body_Result_Find = " select a.spccd, a.rstcd, a.rstcount, b.field1 as spcnm, a.rstcd as rstnm " & _
                           "   from " & TB_LAB316 & " a, " & TB_LAB032 & " b, " & TB_LAB032 & " c " & _
                           "  where a.workdt = " & DBS(pWorkDt) & _
                           "    and a.deptcd = " & DBS(pDeptCd) & _
                           "    and a.coldt  = " & DBS(pColDt) & _
                           "    and a.coltm  = " & DBS(pColTm) & _
                           "    and b.cdindex (+)= " & DBS(LC3_INFECTION) & _
                           "    and b.cdval1 (+)= a.spccd " & _
                           "    and c.cdindex (+)= " & DBS(LC3_INFECTIONTEST) & _
                           "    and c.cdval1 (+)= a.rstcd "

End Function

Public Function Mid_Result_Name(ByVal pMfyID As String) As String

    Mid_Result_Name = " select empnm from " & TB_COM006 & _
                      "  where empid = " & DBS(pMfyID)
                      
End Function

Public Function Dept_Name(ByVal pDeptCd As String) As String

    Dept_Name = " select deptnm from " & TB_Dept & _
                "  where dpcd = " & DBS(pDeptCd)
                      
End Function

Public Function Result_Info(ByVal pFDate As String, ByVal pTDate As String, _
                            Optional ByVal pFlag As String = "0") As String

    '** pFlag (0:의뢰일 기준, 1:보고일 기준)
    If pFlag = "1" Then
        Result_Info = " select a.*, b.deptnm from " & TB_LAB315 & " a, " & TB_Dept & " b " & _
                      "  where a.fvfydt between " & DBS(pFDate) & " and " & DBS(pTDate) & _
                      "    and b.dpcd = a.deptcd " & _
                      "    and a.stscd >= " & DBS(StsCd_LIS_FinRst) & _
                      "  order by mvfydt desc, mvfytm desc "
    Else
        Result_Info = " select a.*, b.deptnm from " & TB_LAB315 & " a, " & TB_Dept & " b " & _
                      "  where a.workdt between " & DBS(pFDate) & " and " & DBS(pTDate) & _
                      "    and b.dpcd = a.deptcd " & _
                      "    and a.stscd >= " & DBS(StsCd_LIS_FinRst) & _
                      "  order by workdt desc "
    End If
End Function

''%  Infection update (S2LAB316) : Final Result
'Public Function LAB316_FUPDATE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, _
'                                   ByVal pColDt As String, ByVal pColTm As String, _
'                                   ByVal pSpcCd As String, ByVal pRstCd As String, _
'                                   ByVal pRstTxt As String, ByVal pRstFg As String) As String
'
'    LAB316_FUPDATE_SQL = " update " & TB_LAB316 & _
'                         "    set testmeth = " & DBS(pTestMeth) & _
'                         "," & " waterfg = " & DBS(pWaterFg) & "," & " rstcount = " & DBS(pRstCnt) & _
'                         "," & " rstval = " & DBS(pRstVal) & "," & " stscd = " & DBS(pStsCd) & _
'                         "," & " fvfyid = " & DBS(pMvfyID) & "," & " fvfydt = " & DBS(pFvfyDt) & _
'                         "," & " fvfytm = " & DBS(pMvfyTm) & "," & " rpttxt = " & DBS(pRptTxt) & _
'                         "  where workdt = " & DBS(pWorkDt) & _
'                         "    and deptcd = " & DBS(pDeptCd) & _
'                         "    and coldt = " & DBS(pColDt) & _
'                         "    and coltm = " & DBS(pColTm)
'
'End Function
'
''%  Infection update (S2LAB316) : Report
'Public Function LAB316_RFUPDATE_SQL(ByVal pWorkDt As String, ByVal pDeptCd As String, ByVal pColDt As String, _
'                                    ByVal pColTm As String, ByVal pTestMeth As String, ByVal pRstCd As String, _
'                                    ByVal pWaterFg As String, ByVal pRstCnt As String, ByVal pRstVal As String, _
'                                    ByVal pStsCd As String, ByVal pRptFg As String, ByVal pRptTxt As String, _
'                                    ByVal pRptDt As String, ByVal pRptTm As String, ByVal pRptID As String) As String
'
'    LAB316_RFUPDATE_SQL = " update " & TB_LAB316 & _
'                         "    set testmeth = " & DBS(pTestMeth) & _
'                         "," & " rstcd = " & DBS(pRstCd) & "," & " waterfg = " & DBS(pWaterFg) & _
'                         "," & " rstcount = " & DBS(pRstCnt) & "," & " stscd = " & DBS(pStsCd) & _
'                         "," & " rptfg = " & DBS(pRptFg) & "," & " rpttxt = " & DBS(pRptTxt) & _
'                         "," & " rptdt = " & DBS(pRptDt) & "," & " rpttm = " & DBS(pRptTm) & _
'                         "," & " rptid = " & DBS(pRptID) & _
'                         "  where workdt = " & DBS(pWorkDt) & _
'                         "    and deptcd = " & DBS(pDeptCd) & _
'                         "    and coldt = " & DBS(pColDt) & _
'                         "    and coltm = " & DBS(pColTm)
'
'End Function
'
