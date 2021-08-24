VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBasisData"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'환자 마스터, 재원 마스터, 부서 마스터, 병상 마스터, 의사 마스터 및 직원 마스터,
'상병 마스터, 수가 마스터, 수술코드 마스터는 병원별로 테이블이 달라진다.

'이 클래스 모듈은 새로운 병원에 적용시 모두 또는 일부를 수정해야 된다.(반듯이)

Public Function GetSQLDept(Optional ByVal vDeptCd As String = "") As String
'부서 마스터
    GetSQLDept = " select distinct " & mF_DEPTCD & " as deptcd, " & mF_DEPTNM & " as deptnm " & _
                 " from " & mT_HIS003
    GetSQLDept = GetSQLDept & " where " & DBW(mF_DEPTCD, vDeptCd, 2)
End Function

Public Function GetSQLDeptList() As String
    GetSQLDeptList = " select distinct " & mF_DEPTCD & " as deptcd, " & mF_DEPTNM & " as deptnm " & _
                 " from " & mT_HIS003
End Function

Public Function GetDeptNm(ByVal vDeptCd As String) As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLDept(vDeptCd), DBConn
    
    If Rs.EOF = False Then
        GetDeptNm = Rs.Fields("deptnm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSQLWard(Optional ByVal vWardId As String = "") As String
'병상 마스터
    GetSQLWard = " select distinct " & mF_WARDID & " as wardid, " & mF_WARDNM & " as wardnm, " & _
                 " '10' as bldgb, '본원' as bldnm, 1 as bldno " & _
                 " from " & mT_HIS004 & _
                 " where (wardyn='Y' OR dpcd in ('HDR','HPC','DME','NNM')) and dpcd not in ('12','21','2S','3S','3E','3W','4E','4W','5E','5W','6E','6W','7W','BA','MID')"
    
    If vWardId <> "" Then
        GetSQLWard = GetSQLWard & " and " & DBW(mF_WARDID, vWardId, 2)
    End If
End Function

Public Function GetSQLWardList() As String
'    GetSQLWardList = " select distinct " & mF_WARDID & " as wardid, " & mF_WARDNM & " as wardnm, " & _
'                 " '10' as bldgb, '본원' as bldnm, 1 as bldno " & _
'                 " from " & mT_HIS004 & _
'                 " where wardyn='Y' "
    GetSQLWardList = " select distinct " & mF_WARDID & " as wardid, " & mF_WARDNM & " as wardnm, " & _
                 " '10' as bldgb, '본원' as bldnm, 1 as bldno, deptcd " & _
                 " from " & mT_HIS004 & _
                 " where wardyn='Y' order by deptcd desc"
End Function

Public Function GetWardNm(ByVal vWardId As String) As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLWard(vWardId), DBConn
    
    If Rs.EOF = False Then
        GetWardNm = Rs.Fields("wardnm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSQLDoct(Optional ByVal vDoctId As String = "") As String
'의사 마스터
    GetSQLDoct = " select distinct " & mF_DOCTID & " as doctid, " & mF_DOCTNM & " as doctnm, deptcd " & _
                 " from " & mT_HIS005
    GetSQLDoct = GetSQLDoct & " where " & DBW(mF_DOCTID, vDoctId, 2)
    GetSQLDoct = GetSQLDoct & " and  jikjong in ('HAA')"
End Function

Public Function GetSQLDoctList() As String
    GetSQLDoctList = " select distinct " & mF_DOCTID & " as doctid, " & mF_DOCTNM & " as doctnm, deptcd " & _
                 " from " & mT_HIS005
    GetSQLDoctList = GetSQLDoctList & " where jikjong in ('HAA')"
End Function

Public Function GetDoctNm(ByVal vDoctId As String) As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLDoct(vDoctId), DBConn
    
    If Rs.EOF = False Then
        GetDoctNm = Rs.Fields("doctnm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSQLEmp(Optional ByVal vEmpId As String = "") As String
'직원 마스터
    GetSQLEmp = " select empid,empnm from " & mT_COM006
    GetSQLEmp = GetSQLEmp & " where " & DBW("empid", vEmpId, 2)
End Function

Public Function GetSQLEmpList() As String
    GetSQLEmpList = " select empid,empnm from " & mT_COM006
End Function

Public Function GetEmpNm(ByVal vEmpId As String) As String
    Dim Rs As New Recordset
    
'COM006 LIS 직원 마스터에 불러와서 없으면 OCS마스터에서 불러온다.

    Rs.Open GetSQLEmp(vEmpId), DBConn
    
    If Rs.EOF Then
        GetEmpNm = GetHisEmpNm(vEmpId)
    Else
        GetEmpNm = Rs.Fields("empnm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSQLHisEmp(Optional ByVal vEmpId As String = "") As String
'HIS 직원 마스터

    GetSQLHisEmp = " select userid empid, username empnm from " & mT_HIS005
    GetSQLHisEmp = GetSQLHisEmp & " where jikjong in ('HAA','HAB','HAC','HABB')"
    GetSQLHisEmp = GetSQLHisEmp & " and " & DBW("userid", vEmpId, 2)
    GetSQLHisEmp = GetSQLHisEmp & " order by userid, username"
End Function

Public Function GetSQLHisEmpList() As String
    GetSQLHisEmpList = " select userid empid, username empnm from " & mT_HIS005
    GetSQLHisEmpList = GetSQLHisEmpList & " where jikjong in ('HAA','HAB','HAC','HABB')"
    GetSQLHisEmpList = GetSQLHisEmpList & " order by userid, username"
End Function

Public Function GetHisEmpNm(ByVal vEmpId As String) As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLHisEmp(vEmpId), DBConn
    
    If Rs.EOF = False Then
        GetHisEmpNm = Rs.Fields("empnm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSQLIcd(Optional ByVal vIcd As String = "") As String
'상병 마스터
    GetSQLIcd = " select   " & mF_ICD & " as icd, " & _
                             mF_IENM & " as ienm " & _
              "   from   " & mT_HIS006 & _
              " where " & DBW(mF_ICD, vIcd, 2) & _
              " order by " & mF_ICD
End Function

Public Function GetSQLIcdList() As String
    GetSQLIcdList = " select   " & mF_ICD & " as icd, " & _
                             mF_IENM & " as ienm " & _
                    "   from   " & mT_HIS006 & _
                    " order by " & mF_ICD
End Function

Public Function GetIcdNm(ByVal vIcd As String) As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLIcd(vIcd), DBConn
    
    If Rs.EOF = False Then
        GetIcdNm = Rs.Fields("ienm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSQLBuild(Optional ByVal vBuild As String = "") As String
    GetSQLBuild = " select cdval1 as buildcd, field1 as buildnm, field2 as buildno " & _
                " from " & mT_LAB032 & _
                " where " & DBW("cdindex", "C230", 2)
    GetSQLBuild = GetSQLBuild & " and " & DBW("cdval1=", vBuild)
End Function

Public Function GetSQLbuildList() As String
    GetSQLbuildList = " select cdval1 as buildcd, field1 as buildnm, field2 as buildno " & _
                    " from " & mT_LAB032 & _
                    " where " & DBW("cdindex", "C230", 2)
End Function

Public Function GetBuildNm(ByVal vBuild As String, Optional ByRef vBuildNo As String = "") As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLBuild(vBuild), DBConn
    
    If Rs.EOF = False Then
        GetBuildNm = Rs.Fields("buildnm").Value & ""
        vBuildNo = Rs.Fields("buildno").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetPtNm(ByVal vPtID As String) As String
    Dim objSQL As New clsPatient
    Dim Rs As New Recordset
    
    Rs.Open objSQL.GetSQLPt(vPtID), DBConn
    
    If Rs.EOF = False Then
        GetPtNm = Rs.Fields("ptnm").Value & ""
    End If
    
    Set Rs = Nothing
    Set objSQL = Nothing
End Function

Public Function GetSQLOP(Optional ByVal vOCd As String = "") As String
    GetSQLOP = " select " & mF_OCD & " as ocd," & mF_ONM & " as onm" & _
               " from  " & mT_HIS007
    GetSQLOP = GetSQLOP & " where " & DBW(mF_OCD, vOCd, 2)
End Function

Public Function GetSQLOPList() As String
    GetSQLOPList = " select " & mF_OCD & " as ocd," & mF_ONM & " as onm" & _
               " from  " & mT_HIS007
End Function

Public Function GetOPNm(ByVal vOCd As String) As String
    Dim Rs As New Recordset
    
    Rs.Open GetSQLOP(vOCd), DBConn
    
    If Rs.EOF = False Then
        GetOPNm = Rs.Fields("onm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

