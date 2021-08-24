VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private AdoCn               As ADODB.Connection
Private AdoRs               As ADODB.Recordset
Private Call_Procedure      As String

Public Sub SetAdoCn(ByVal new_AdoCn As ADODB.Connection)
    Set AdoCn = new_AdoCn
End Sub

Private Sub Class_Initialize()
    Call DBErrClear
End Sub

Private Sub Class_Terminate()
    Set AdoCn = Nothing
End Sub

Public Function Get_TestDate(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset
    Dim strSql  As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TestDate() As ADODB.Recordset"
    
    strSql = "SELECT ACCDT FROM LIMAS101 WHERE ACCDT BETWEEN '" & FromDate & "' AND '" & ToDate & "' GROUP BY ACCDT"
    
    Set AdoRs = New ADODB.Recordset
    
    If GetRecordset(AdoCn, strSql, AdoRs, Call_Procedure) Then
        Set Get_TestDate = AdoRs
    Else
        Set Get_TestDate = Nothing
    End If
    
    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Call ErrMsgProc(Call_Procedure)

End Function

Public Function Get_TestName(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset
    Dim strSql  As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TestName() As ADODB.Recordset"
    
    strSql = "SELECT LIMAS102.TESTCD AS TESTCD, LIMAS001.TestNm AS TESTNM, LIMAS001.PrintSeq" & _
            " FROM LIMAS101, LIMAS102, LIMAS001" & _
            " WHERE (LIMAS101.ACCNO = LIMAS102.ACCNO) AND (LIMAS102.TESTCD = LIMAS001.TestCd)" & _
                " AND (LIMAS101.ACCDT BETWEEN '" & FromDate & "' AND '" & ToDate & "')" & _
            " GROUP BY LIMAS102.TESTCD, LIMAS001.TestNm,LIMAS001.PrintSeq" & _
            " ORDER BY LIMAS001.PrintSeq"
    
    Set AdoRs = New ADODB.Recordset
    
    If GetRecordset(AdoCn, strSql, AdoRs, Call_Procedure) Then
        Set Get_TestName = AdoRs
    Else
        Set Get_TestName = Nothing
    End If
    
    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

Public Function Get_TestCount(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset
    Dim strSql  As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TestCount(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset"
    
    strSql = "SELECT LIMAS101.ACCDT AS ACCDT, COUNT(LIMAS102.TESTCD) AS CNT, LIMAS102.TESTCD AS TESTCD" & _
            " FROM LIMAS101, LIMAS102, LIMAS001" & _
            " WHERE (LIMAS101.ACCNO = LIMAS102.ACCNO) AND (LIMAS102.TESTCD = LIMAS001.TestCd)" & _
                " AND (LIMAS101.ACCDT BETWEEN '" & FromDate & "' AND '" & ToDate & "')" & _
            " GROUP BY LIMAS101.ACCDT, LIMAS102.TESTCD" & _
            " ORDER BY LIMAS101.ACCDT, LIMAS102.TESTCD"
    
    Set AdoRs = New ADODB.Recordset
    
    If GetRecordset(AdoCn, strSql, AdoRs, Call_Procedure) Then
        Set Get_TestCount = AdoRs
    Else
        Set Get_TestCount = Nothing
    End If
    
    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

Public Function Get_SlipCount(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset
    Dim strSql  As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_SlipCount(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset"
    
    strSql = "SELECT ACCDT, COUNT(ACCNO) AS SLIP_CNT" & _
            " FROM LIMAS101" & _
            " WHERE (ACCDT BETWEEN '" & FromDate & "' AND '" & ToDate & "')" & _
            " GROUP BY ACCDT" & _
            " ORDER BY ACCDT"
    
    Set AdoRs = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, AdoRs, Call_Procedure) Then
        Set Get_SlipCount = AdoRs
    Else
        Set Get_SlipCount = Nothing
    End If
    
    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

Public Function Get_TotalTestCount(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset
    Dim strSql  As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Get_TotalTestCount(ByVal FromDate As String, ByVal ToDate As String) As ADODB.Recordset"
    
    strSql = "SELECT LIMAS101.ACCDT AS ACCDT, COUNT(LIMAS101.ACCNO) AS TEST_CNT" & _
            " FROM LIMAS101, LIMAS102" & _
            " WHERE (LIMAS101.ACCNO = LIMAS102.ACCNO)" & _
                " AND (LIMAS101.ACCDT BETWEEN '" & FromDate & "' AND '" & ToDate & "')" & _
            " GROUP BY LIMAS101.ACCDT" & _
            " ORDER BY LIMAS101.ACCDT"
    
    Set AdoRs = New ADODB.Recordset
    If GetRecordset(AdoCn, strSql, AdoRs, Call_Procedure) Then
        Set Get_TotalTestCount = AdoRs
    Else
        Set Get_TotalTestCount = Nothing
    End If
    
    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function
'SELECT LIMAS101.ACCDT, COUNT(LIMAS101.ACCNO) AS TEST_CNT FROM LIMAS101, LIMAS102 WHERE LIMAS101.ACCNO = LIMAS102.ACCNO GROUP BY LIMAS101.ACCDT
