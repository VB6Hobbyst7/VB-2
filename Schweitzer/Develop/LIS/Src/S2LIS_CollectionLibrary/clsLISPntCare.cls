VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISPntCare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetTestCd(ByRef pTestCd As String, ByRef pUnit As String) As Boolean
    Dim strSql As String
    Dim DrRs   As Recordset
    
    strSql = "SELECT cdval1 testcd,field1 unit " & _
             "FROM  " & T_LAB032 & " " & _
             "WHERE " & DBW("cdindex=", LC3_POCTestCd)
             
    Set DrRs = New Recordset
    Call DrRs.Open(strSql, DBConn)
    
    If DrRs.RecordCount < 1 Then
        GetTestCd = False
    Else
        GetTestCd = True
        pTestCd = DrRs.Fields("testcd").Value & ""
        pUnit = DrRs.Fields("unit").Value & ""
    End If
    
    Set DrRs = Nothing
End Function

Public Function GetTestTimes(ByRef pTime() As String, ByRef pDTime() As String) As Boolean
    Dim strSql As String
    Dim DrRs   As Recordset
    Dim i      As Long
    
    strSql = "SELECT cdval1,field1 " & _
             "FROM  " & T_LAB032 & " " & _
             "WHERE " & DBW("cdindex=", LC3_POCTimes) & " " & _
             "ORDER BY cdval1"
             
    Set DrRs = New Recordset
    Call DrRs.Open(strSql, DBConn)
    
    If DrRs.RecordCount < 1 Then
        GetTestTimes = False
    Else
        GetTestTimes = True
        Erase pTime
        Erase pDTime
        For i = 1 To DrRs.RecordCount
            ReDim Preserve pTime(i - 1)
            ReDim Preserve pDTime(i - 1)
            
            pTime(i - 1) = DrRs.Fields("cdval1").Value & ""
            pDTime(i - 1) = DrRs.Fields("field1").Value & ""
            
            DrRs.MoveNext
        Next i
    End If
    
    Set DrRs = Nothing
End Function

Public Function GetPtInfo(ByVal pPtid As String, _
                          ByRef pPtNm As String, ByRef pSex As String, ByRef pDOB As String, ByRef pAge As String) As Boolean
                          
    Dim strSql As String
    Dim DrRs   As Recordset
    
    strSql = "SELECT " & F_PTNM & " as ptnm," & F_SEX2 & " as sex," & F_DOB2 & " as dob " & _
             "FROM   " & T_HIS001 & " " & _
             "WHERE  " & DBW(F_PTID & "=", pPtid)
    Set DrRs = New Recordset
    Call DrRs.Open(strSql, DBConn)
    
    If DrRs.RecordCount < 1 Then
        GetPtInfo = False
    Else
        GetPtInfo = True
        pPtNm = DrRs.Fields("ptnm").Value & ""
        pSex = DrRs.Fields("sex").Value & ""
        pDOB = DrRs.Fields("dob").Value & ""
        pAge = medFindAge(pDOB, "Y")
    End If
    
    Set DrRs = Nothing
End Function

Public Function GetPtBedInfo(ByVal pPtid As String, _
                             ByRef pBedInDt As String, ByRef pBedOutDt As String, _
                             ByRef pDeptCd As String, ByRef pDoctCd As String, _
                             ByRef pWardId As String, ByRef pRoomId As String) As Boolean
                             
    Dim strSql As String
    Dim DrRs   As Recordset
    
    strSql = " SELECT " & F_BEDINDT & " as bedindt," & F_BEDOUTDT2 & " as bedoutdt," & _
                         F_PTDEPTCD & " as deptcd," & F_DOCTID & " as doctid," & _
                         F_PTWARDID & " as wardid," & F_PTROOMID & " as roomid" & _
             " FROM  " & T_HIS002 & " " & _
             " WHERE " & DBW(F_INPTID & "=", pPtid) & " " & _
             " AND  in_date= (SELECT max(in_date) " & _
                                       "FROM  " & T_HIS002 & " " & _
                                       "WHERE " & DBW(F_INPTID & "=", pPtid) & _
                                       ")"
    Set DrRs = New Recordset
    Call DrRs.Open(strSql, DBConn)
    
    If DrRs.RecordCount < 1 Then
        GetPtBedInfo = False
    Else
        GetPtBedInfo = True
        pBedInDt = DrRs.Fields("bedindt").Value & ""
        pBedOutDt = DrRs.Fields("bedoutdt").Value & ""
        pDeptCd = DrRs.Fields("deptcd").Value & ""
        pDoctCd = DrRs.Fields("doctid").Value & ""
        pWardId = DrRs.Fields("wardId").Value & ""
        pRoomId = DrRs.Fields("roomId").Value & ""
    End If
    
    Set DrRs = Nothing
End Function

Public Function Delete(ByVal pPtid As String, ByVal pVfydt As String) As Boolean
    Dim strSql As String
    
On Error GoTo Delete_error

    strSql = "delete FROM " & T_LAB195 & " " & _
             "WHERE " & DBW("ptid=", pPtid) & " " & _
             "AND   " & DBW("vfydt=", pVfydt)
    DBConn.Execute strSql
    Delete = True
    Exit Function
    
Delete_error:
    MsgBox Err.Description, vbExclamation
    Delete = False
End Function

Public Function Save(ByVal pPtid As String, _
                     ByVal pVfydt As String, _
                     ByVal pVfyTm As String, _
                     ByVal pVfyId As String, _
                     ByVal pTestCd As String, _
                     ByVal pRstCd As String, _
                     ByVal pRstUnit As String, _
                     ByVal pEntDt As String, _
                     ByVal pEntTm As String, _
                     ByVal pDeptCd As String, _
                     ByVal pWardId As String, _
                     ByVal pMajDoct As String) As Boolean
                     
    Dim strSql As String
    
    '----------------------------------------------------
    'ÀúÀå s2lab195    insert
    '----------------------------------------------------
On Error GoTo Save_error

    strSql = "insert into " & T_LAB195 & "(ptid,vfydt,vfytm,vfyid,testcd," & _
                                          "rstcd,rstunit,entdt,enttm,deptcd," & _
                                          "wardid,majdoct) " & _
             "values(" & DBV("ptid", pPtid) & "," & _
                         DBV("vfydt", pVfydt) & "," & _
                         DBV("vfytm", pVfyTm) & "," & _
                         DBV("vfyid", pVfyId) & "," & _
                         DBV("testcd", pTestCd) & "," & _
                         DBV("rstcd", pRstCd) & "," & _
                         DBV("rstunit", pRstUnit) & "," & _
                         DBV("entdt", pEntDt) & "," & _
                         DBV("enttm", pEntTm) & "," & _
                         DBV("deptcd", pDeptCd) & "," & _
                         DBV("wardid", pWardId) & "," & _
                         DBV("majdoct", pMajDoct) & ")"
    DBConn.Execute strSql
    Save = True
    Exit Function
    
Save_error:
    MsgBox Err.Description, vbExclamation
    
    Save = False
End Function
                     
Public Function LoadHistory(ByVal pPtid As String, ByVal pBedInDt As String, ByVal pBedOutDt As String, _
                            ByRef pVfydt() As String, ByRef pVfyTm() As String, ByRef pRstCd() As String, ByRef pVfyId() As String) As Long
                            
    Dim strSql  As String
    Dim DrRs    As Recordset
    Dim strFrDt As String
    Dim strToDt As String
    Dim i       As Long
    
    
    strFrDt = pBedInDt
    If pBedOutDt = "" Then
        strToDt = Format(GetSystemDate, "YYYYMMDD")
    Else
        strToDt = pBedOutDt
    End If
    
    strSql = "SELECT * " & _
             "FROM  " & T_LAB195 & " " & _
             "WHERE " & DBW("ptid=", pPtid) & " " & _
             "AND   " & DBW("vfydt>=", strFrDt) & " " & _
             "AND   " & DBW("vfydt<=", strToDt) & " " & _
             "AND   rstcd is not null " & _
             "ORDER BY ptid,vfydt,vfytm "
    Set DrRs = New Recordset
    
    Call DrRs.Open(strSql, DBConn)
    If DrRs.RecordCount < 1 Then
        LoadHistory = 0
    Else
        LoadHistory = DrRs.RecordCount
        For i = 1 To DrRs.RecordCount
            ReDim Preserve pVfydt(i - 1)
            ReDim Preserve pVfyTm(i - 1)
            ReDim Preserve pRstCd(i - 1)
            ReDim Preserve pVfyId(i - 1)
            
            pVfydt(i - 1) = DrRs.Fields("vfydt").Value & ""
            pVfyTm(i - 1) = DrRs.Fields("vfytm").Value & ""
            pRstCd(i - 1) = DrRs.Fields("rstcd").Value & ""
            pVfyId(i - 1) = DrRs.Fields("vfyid").Value & ""
            
            DrRs.MoveNext
        Next i
    End If
    
    Set DrRs = Nothing
End Function

Public Function GetResult(ByVal pPtid As String, ByVal pVfydt As String, _
                          ByRef pRstCd() As String, ByRef pVfyId() As String, ByRef pEntDt() As String, ByRef pEntTm() As String) As Long

    Dim strSql  As String
    Dim DrRs    As Recordset
    Dim i       As Long
    
    strSql = "SELECT * " & _
             "FROM  " & T_LAB195 & " " & _
             "WHERE " & DBW("ptid=", pPtid) & " " & _
             "AND   " & DBW("vfydt=", pVfydt) & " " & _
             "ORDER BY ptid,vfydt,vfytm "
    Set DrRs = New Recordset
    
    Call DrRs.Open(strSql, DBConn)
    If DrRs.RecordCount < 1 Then
        GetResult = 0
    Else
        GetResult = DrRs.RecordCount
        For i = 1 To DrRs.RecordCount
            ReDim Preserve pRstCd(i - 1)
            ReDim Preserve pVfyId(i - 1)
            ReDim Preserve pEntDt(i - 1)
            ReDim Preserve pEntTm(i - 1)
            
            pRstCd(i - 1) = DrRs.Fields("rstcd").Value & ""
            pVfyId(i - 1) = DrRs.Fields("vfyid").Value & ""
            pEntDt(i - 1) = DrRs.Fields("entdt").Value & ""
            pEntTm(i - 1) = DrRs.Fields("enttm").Value & ""
            
            DrRs.MoveNext
        Next i
    End If
    Set DrRs = Nothing
End Function
