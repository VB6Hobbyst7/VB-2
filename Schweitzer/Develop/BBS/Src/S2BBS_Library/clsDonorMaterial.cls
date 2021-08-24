VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Public Function GetDonorMaterial(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT ordcd,qty " & _
           " FROM " & T_BBS606 & " " & _
           " WHERE " & DBW("donorid", Donorid, 2) & _
           " AND " & DBW("donoraccdt", donoraccdt, 2)
           
    Set GetDonorMaterial = New Recordset
    Call GetDonorMaterial.Open(sSql, DBConn)
'    If GetDonorMaterial.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetDonorMaterial = Nothing
'    End If
End Function

Public Function GetDonorDonation(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT a.donationdt,a.bldsrc,a.bldyy,a.bldno,a.compocd,a.volumn,a.entfg,b.abbrnm as componm " & _
           " FROM " & T_BBS602 & " a," & T_BBS006 & " b " & _
           " WHERE " & DBW("a.donorid", Donorid, 2) & _
           "  AND " & DBW("a.donoraccdt", donoraccdt, 2) & _
           "  AND a.compocd=b.compocd "
    
    Set GetDonorDonation = New Recordset
    Call GetDonorDonation.Open(sSql, DBConn)
'    If GetDonorDonation.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetDonorDonation = Nothing
'    End If
End Function

Private Function GetNextBldNo(ByVal centercd As String) As String
    Dim sSql As String
    Dim DrRS As Recordset
    Dim Bldno As String
    
    sSql = " SELECT min(bldbundleno) bldno " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("centercd", centercd, 2) & _
           " AND " & DBW("usedt", "", 2) & _
           " AND " & DBW("returndt", "", 2)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        GetNextBldNo = ""
        Exit Function
    End If
    
    With DrRS
        If .RecordCount < 1 Then
            MsgBox "더 이상 할당할 헌혈증서가 없습니다", vbCritical, "오류"
            GetNextBldNo = ""
        Else
            Bldno = .Fields("bldno").Value & ""
            If Bldno = "" Then
                MsgBox "더 이상 할당할 헌혈증서가 없습니다", vbCritical, "오류"
                GetNextBldNo = ""
            Else
                GetNextBldNo = Bldno
            End If
        End If
    End With
    Set DrRS = Nothing
End Function

Public Function Save(ByVal Donorid As String, ByVal donoraccdt As String, _
                     ByVal DonationDt As String, ByVal pBldno As String, ByVal CompoCd As String, ByVal volumn As String, _
                     ByVal rec As String, ByVal centercd As String, ByVal EmpID As String) As Boolean
                     
    Dim bldbundleno As String
    Dim Bldsrc As String
    Dim Bldyy As String
    Dim Bldno As String
    
    Dim OrdCd As String
    Dim qty As String
    
    Dim tRec As String
    Dim usedt As String
    
    Dim dSql As String
    Dim sSql() As String
    Dim i As Long
    
    If pBldno = "" Then
        bldbundleno = GetNextBldNo(centercd)
        If bldbundleno = "" Then
            Save = False
            Exit Function
        Else
            Bldsrc = Mid(bldbundleno, 1, 2)
            Bldyy = Mid(bldbundleno, 5, 2)
            Bldno = Trim(Val(Mid(bldbundleno, 7)))
        End If
    Else
        Bldsrc = medGetP(pBldno, 1, "-")
        Bldyy = medGetP(pBldno, 2, "-")
        Bldno = medGetP(pBldno, 3, "-")
    End If
    
    usedt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    ReDim Preserve sSql(0)
    sSql(0) = " update " & T_BBS602 & " " & _
              " set " & DBW("donationdt", DonationDt, 3) & _
                        DBW("bldsrc", Bldsrc, 3) & _
                        DBW("bldyy", Bldyy, 3) & _
                        DBW("bldno", Bldno, 3) & _
                        DBW("volumn", volumn, 3) & _
                        DBW("compocd", CompoCd, 3) & _
                        DBW("entfg", "1", 2) & _
              " WHERE " & DBW("donorid", Donorid, 2) & _
              " AND " & DBW("donoraccdt", donoraccdt, 2)
              
    dSql = " delete FROM " & T_BBS606 & " " & _
           " WHERE " & DBW("donorid", Donorid, 2) & _
           " AND " & DBW("donoraccdt", donoraccdt, 2)

    i = 0
    Do
        i = i + 1
        tRec = medGetP(rec, i, LINE_DIV)
        If tRec = "" Then Exit Do
        
        OrdCd = medGetP(tRec, 1, COL_DIV)
        qty = medGetP(tRec, 2, COL_DIV)
        
        ReDim Preserve sSql(i)
        sSql(i) = "insert into " & T_BBS606 & "(donorid,donoraccdt,ordcd,qty) " & _
                  "values(" & _
                         DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & DBV("ordcd", OrdCd, 1) & DBV("qty", qty) & ")"
                  
    Loop
    
    ReDim Preserve sSql(i)
    sSql(i) = " update " & T_BBS411 & " " & _
              " set " & DBW("usedt", usedt, 3) & _
                  " " & DBW("useid", EmpID, 2) & _
              " WHERE " & _
                                  DBW("bldsrc", Bldsrc, 2) & _
                        " AND " & DBW("bldyy", Bldyy, 2) & _
                        " AND " & DBW("bldno", Bldno, 2) & _
                        " AND " & DBW("usedt", "", 2)
    
On Error GoTo Save_error
    DBConn.BeginTrans
    DBConn.Execute dSql
    For i = LBound(sSql) To UBound(sSql)
        DBConn.Execute sSql(i)
    Next i
    DBConn.CommitTrans
    Save = True
    Exit Function
    
Save_error:
'    DBConn.DisplayErrors
    DBConn.RollbackTrans
    Save = False
End Function

