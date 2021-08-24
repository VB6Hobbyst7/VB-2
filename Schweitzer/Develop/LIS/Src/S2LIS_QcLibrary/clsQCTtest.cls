VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQCTtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetTestItem(ByVal pCtrlCd As String, ByVal pLevel As String, ByVal pLotNo As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,b.abbrnm10 from " & T_LAB024 & " a, " & T_LAB001 & " b" & _
             " where " & DBW("ctrlcd=", pCtrlCd) & _
             " and " & DBW("levelcd=", pLevel) & _
             " and " & DBW("lotno=", pLotNo) & _
             " and a.testcd=b.testcd"

    Set GetTestItem = New Recordset
    GetTestItem.Open strSQL, dbconn

End Function

Public Function GetEqp(ByVal pCtrlCd As String, ByVal pLevel As String, pLotNo As String, _
                       ByVal pTestItem As String, ByVal pFdate As String, pTdate As String, _
                       ByVal pARdiv As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select distinct a.eqpcd,b.eqpnm from " & T_LAB026 & " a, " & T_LAB006 & " b " & _
             " where " & DBW("a.ctrlcd=", pCtrlCd) & _
             " and " & DBW("a.levelcd=", pLevel) & _
             " and " & DBW("a.lotno=", pLotNo) & _
             " and a.testcd in (" & pTestItem & ")" & _
             " and a.vfydt between " & DBS(pFdate) & " and " & DBS(pTdate) & _
             " and " & DBJ("a.eqpcd*=b.eqpcd") & pARdiv & _
             " order by a.eqpcd"
    
    Set GetEqp = New Recordset
    GetEqp.Open strSQL, dbconn
End Function

Public Function GetData(ByVal pCtrlCd As String, ByVal pLevel As String, ByVal pLotNo As String, _
                        ByVal pTestItem As String, ByVal pFdate As String, ByVal pTdate As String, _
                        ByVal pARdiv As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select a.workarea,a.accdt,a.accseq,a.vfydt,a.vfytm,a.testcd,c.abbrnm10," & _
             " a.eqpcd,b.eqpnm,a.rsttype,a.rstcd,a.rstval,a.radiv " & _
             " from " & T_LAB026 & " a, " & T_LAB006 & " b, " & T_LAB001 & " c" & _
             " where " & DBW("a.ctrlcd=", pCtrlCd) & _
             " and " & DBW("a.levelcd=", pLevel) & _
             " and " & DBW("a.lotno=", pLotNo) & _
             " and a.testcd in (" & pTestItem & ")" & _
             " and a.vfydt between " & DBS(pFdate) & " and " & DBS(pTdate) & _
             " and " & DBJ("a.eqpcd*=b.eqpcd") & _
             " and a.testcd=c.testcd " & pARdiv & _
             " order by a.testcd,a.vfydt,a.eqpcd  "
             
   Set GetData = New Recordset
   GetData.Open strSQL, dbconn
End Function

Public Function GetLotNo(ByVal pCtrlCd As String, ByVal pLevel As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select lotno from " & T_LAB023 & _
             " where " & DBW("ctrlcd =", pCtrlCd) & _
             " and " & DBW("levelcd=", pLevel)
             
    Set GetLotNo = New Recordset
    GetLotNo.Open strSQL, dbconn
End Function

Public Function GetComment(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                           ByVal pTestCd As String) As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT seq, text FROM " & T_LAB028 & " " & _
            " WHERE " & DBW("workarea =", pWorkArea) & _
            " AND " & DBW("accdt =", pAccDt) & _
            " AND " & DBW("accseq = ", pAccSeq) & _
            " AND " & DBW("testcd = ", pTestCd) & _
            " ORDER BY seq desc"
    
    Set GetComment = New Recordset
    GetComment.Open strSQL, dbconn
End Function

Public Function GetEqpNm(ByVal pEqpCd As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select eqpnm from " & T_LAB006 & _
            " where " & DBW("eqpcd=", pEqpCd)

    Set GetEqpNm = New Recordset
    GetEqpNm.Open strSQL, dbconn
End Function
