VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISComFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
'Public Function GetEmpName(ByVal EmpId As String) As String
''직원이름을 가지고온다.
'    Dim RS      As Recordset
'    Dim SSQL    As String
'
'    SSQL = " SELECT empid,empnm" & _
'             " FROM " & T_COM006 & _
'             " WHERE" & _
'                      " " & DBW("empid", EmpId, 2)
'
'    Set RS = New Recordset
'    RS.Open SSQL, dbconn
'    If RS.EOF Then
'        GetEmpName = ""
'    Else
'        GetEmpName = RS.Fields("empnm").Value & ""
'    End If
'    Set RS = Nothing
'
'    If GetEmpName = "" Then
'        SSQL = " SELECT " & F_EMPNM & "  as enpnm FROM " & T_HIS005 & " WHERE" & DBW(F_EMPID, EmpId, 2)
'        Set RS = New Recordset
'        RS.Open SSQL, dbconn
'        If Not RS.EOF Then GetEmpName = RS.Fields("enpnm").Value & ""
'    End If
'
'    Set RS = Nothing
'
'End Function
'
'Public Function GetDeptName(ByVal qDeptCd As String) As String
'    Dim RS  As Recordset
'
'    Set RS = New Recordset
'    RS.Open GetDeptListSQL(qDeptCd), dbconn
'    If Not RS.EOF Then
'        GetDeptName = RS.Fields("deptnm").Value & ""
'    End If
'    Set RS = Nothing
'End Function
'
'
'Public Function GetWardName(ByVal qWardID As String) As String
'    Dim RS  As Recordset
'
'    Set RS = New Recordset
'    RS.Open GetWardListSQL(qWardID), dbconn
'    If Not RS.EOF Then
'        GetWardName = RS.Fields("wardnm").Value & ""
'    End If
'    Set RS = Nothing
'End Function
'
'Public Function GetDoctName(ByVal qDoctID As String) As String
'    Dim RS  As Recordset
'
'    Set RS = New Recordset
'    RS.Open GetDoctListSQL(qDoctID), dbconn
'    If Not RS.EOF Then
'        GetDoctName = RS.Fields("doctnm").Value & ""
'    End If
'    Set RS = Nothing
'
'End Function
'
'Public Function GetPatientName(ByVal qPtid As String) As String
'    Dim RS  As Recordset
'
'    Set RS = New Recordset
'    RS.Open GetPatientNameSQL(qPtid), dbconn
'    If Not RS.EOF Then
'        GetPatientName = RS.Fields("ptnm").Value & ""
'    End If
'    Set RS = Nothing
'
'End Function
'
'Public Function GetPatientInfo(ByVal qPtid As String, ByRef qPtnm As String, ByRef qSex As String, _
'                               ByRef qDob As String)
'
'    Dim RS  As Recordset
'
'    Set RS = New Recordset
'    RS.Open GetPatientNameSQL(qPtid), dbconn
'    If Not RS.EOF Then
'        qPtnm = RS.Fields("ptnm").Value & ""
'        qSex = RS.Fields("sex").Value & ""
'        qDob = RS.Fields("dob").Value & ""
'    End If
'    Set RS = Nothing
'End Function
'
'Public Function GetPatientNameSQL(ByVal qPtid As String) As String
'    GetPatientNameSQL = " SELECT " & F_PTID & " as ptid," & F_PTNM & " AS ptnm, " & _
'                                     F_SEX & " as sex," & F_DOB & " as dob" & _
'                        " FROM " & T_HIS001 & _
'                        " WHERE " & DBW(F_PTID, qPtid, 2)
'End Function
'
''%  11. Get Department List : 부서리스트
'Public Function GetDeptListSQL(Optional ByVal DeptCd As String) As String
'
'    If IsMissing(DeptCd) Then
'        GetDeptListSQL = " SELECT  " & F_DEPTCD & " as deptcd, " & _
'                                        F_DEPTNM & " as deptnm " & _
'                         " FROM " & T_HIS003 & _
'                         " ORDER BY rank"
'    Else
'        GetDeptListSQL = " SELECT  " & F_DEPTCD & " as deptcd, " & _
'                                            F_DEPTNM & " as deptnm " & _
'                         " FROM " & T_HIS003 & _
'                         " WHERE " & DBW(F_DEPTCD, DeptCd, 2) & _
'                         " ORDER BY rank"
'    End If
'End Function
'
''%  12. Get Ward Id List : 병동 리스트
'Public Function GetWardListSQL(Optional ByVal WardID As String) As String
'    If IsMissing(WardID) Then
'        GetWardListSQL = " SELECT distinct " & F_WARDID & " as wardid, " & _
'                                                  F_WARDID & " as wardnm " & _
'                         " FROM " & T_HIS004 & _
'                         " ORDER BY wardid"
'   Else
'        GetWardListSQL = " SELECT distinct " & F_WARDID & " as wardid, " & _
'                                            F_WARDID & " as wardnm" & _
'                         " FROM " & T_HIS004 & _
'                         " WHERE " & DBW(F_WARDID, WardID, 2) & _
'                         " ORDER BY wardid"
'   End If
'End Function
'
'
'Public Function GetEmpListSQL(Optional ByVal EmpId As String) As String
'    GetEmpListSQL = " SELECT " & F_EMPID & " as empid," & F_EMPNM & "  as enpnm FROM " & T_HIS005
'    If EmpId <> "" Then
'        GetEmpListSQL = GetEmpListSQL & " WHERE" & DBW(F_EMPID, EmpId, 2) & _
'                                        " AND  " & DBW(F_EXPDT, "Y", 2)
'    Else
'        GetEmpListSQL = GetEmpListSQL & " WHERE " & DBW(F_EXPDT, "Y", 2)
'    End If
'End Function
'
'Public Function GetDoctListSQL(Optional ByVal Doct As String) As String
'
'    GetDoctListSQL = " SELECT " & F_EMPID & " as doctid," & _
'                                  F_EMPNM & " as doctnm FROM " & T_HIS005
'
'    If Doct <> "" Then
'        GetDoctListSQL = GetDoctListSQL & _
'                         " WHERE" & DBW(F_EMPID, Doct, 2) & _
'                         " AND " & F_EMPDIV & " IN ('0','1')" & _
'                         " AND " & DBW(F_EXPDT, "Y", 2) '& _
'                         " ORDER BY rank"
'    Else
'        GetDoctListSQL = GetDoctListSQL & " WHERE " & F_EMPDIV & " IN ('0','1')" & _
'                        " AND " & DBW(F_EXPDT, "Y", 2) '& _
'                         " ORDER BY rank"
'    End If
'
'End Function
