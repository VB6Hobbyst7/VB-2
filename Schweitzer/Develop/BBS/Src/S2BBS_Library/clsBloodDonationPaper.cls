VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBloodDonationPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'속성값을 가질 지역 변수
Private mvarBldSrc As String '지역 복사
Private mvarBldYY As String
Private mvarBldNo As String '지역 복사
Private mvarRcvDt As String '지역 복사
Private mvarRcvId As String '지역 복사
Private mvarRcvNm As String '지역 복사
Private mvarDivCd As String '지역 복사
Private mvarCenterCd As String '지역 복사
Private mvarReturnDt As String '지역 복사
Private mvarReturnId As String '지역 복사
Private mvarUseDt As String '지역 복사
Private mvarUseId As String '지역 복사
Private mvarBldBundleNo As String '지역 복사

' 조회용 지역변수
Private mvarFrNo As String
Private mvarToNo As String
Private mvarStatus As String

Public Sub Clear()
    mvarBldSrc = ""
    mvarBldYY = ""
    mvarBldNo = ""
    mvarRcvDt = ""
    mvarRcvId = ""
    mvarRcvNm = ""
    mvarDivCd = ""
    mvarCenterCd = ""
    mvarReturnDt = ""
    mvarReturnId = ""
    mvarUseDt = ""
    mvarUseId = ""
    mvarBldBundleNo = ""
End Sub

Private Sub CheckBundleNo()
'    If mvarBldBundleNo = "" Then
        mvarBldBundleNo = mvarBldSrc & Mid(mvarRcvDt, 1, 4) & Format(mvarBldNo, "000000")
'    End If
End Sub


Public Property Let status(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.BldBundleNo = 5
    mvarStatus = vData
End Property


Public Property Get status() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.BldBundleNo
    status = mvarStatus
End Property


Public Property Let ToNo(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.BldBundleNo = 5
    mvarToNo = vData
End Property


Public Property Get ToNo() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.BldBundleNo
    ToNo = mvarToNo
End Property


Public Property Let FrNo(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.BldBundleNo = 5
    mvarFrNo = vData
End Property


Public Property Get FrNo() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.BldBundleNo
    FrNo = mvarFrNo
End Property




Public Property Let bldbundleno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.BldBundleNo = 5
    mvarBldBundleNo = vData
End Property


Public Property Get bldbundleno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.BldBundleNo
    bldbundleno = mvarBldBundleNo
End Property



Public Property Let useid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.UseId = 5
    mvarUseId = vData
End Property


Public Property Get useid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.UseId
    useid = mvarUseId
End Property



Public Property Let usedt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.UseDt = 5
    mvarUseDt = vData
End Property


Public Property Get usedt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.UseDt
    usedt = mvarUseDt
End Property



Public Property Let returnid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ReturnId = 5
    mvarReturnId = vData
End Property


Public Property Get returnid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ReturnId
    returnid = mvarReturnId
End Property



Public Property Let returndt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ReturnDt = 5
    mvarReturnDt = vData
End Property


Public Property Get returndt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ReturnDt
    returndt = mvarReturnDt
End Property



Public Property Let centercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CenterCd = 5
    mvarCenterCd = vData
End Property


Public Property Get centercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CenterCd
    centercd = mvarCenterCd
End Property



Public Property Let rcvnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.RcvNm = 5
    mvarRcvNm = vData
End Property


Public Property Get rcvnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.RcvNm
    rcvnm = mvarRcvNm
End Property



Public Property Let rcvid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.RcvId = 5
    mvarRcvId = vData
End Property


Public Property Get rcvid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.RcvId
    rcvid = mvarRcvId
End Property



Public Property Let DivCd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.DivCd = 5
    mvarDivCd = vData
End Property


Public Property Get DivCd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.DivCd
    DivCd = mvarDivCd
End Property



Public Property Let Bldno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.BldNo = 5
    mvarBldNo = vData
End Property


Public Property Get Bldno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.BldNo
    Bldno = mvarBldNo
End Property



Public Property Let rcvdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.RcvDt = 5
    mvarRcvDt = vData
End Property


Public Property Get rcvdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.RcvDt
    rcvdt = mvarRcvDt
End Property

Public Property Let Bldyy(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bldsrc = 5
    mvarBldYY = vData
End Property


Public Property Get Bldyy() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bldsrc
    Bldyy = mvarBldYY
End Property



Public Property Let Bldsrc(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bldsrc = 5
    mvarBldSrc = vData
End Property


Public Property Get Bldsrc() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bldsrc
    Bldsrc = mvarBldSrc
End Property




Private Sub Class_Initialize()
    Clear
End Sub








Public Function Delete(ByVal Bldsrc As String, ByVal rcvdt As String) As Boolean
    Dim sSql As String
    
    sSql = " delete FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
           " AND " & DBW("rcvdt", rcvdt, 2)
           
On Error GoTo Delete_error

    DBConn.Execute (sSql)
'        GoTo Delete_error
'    End If

    Delete = True
    Exit Function
    
Delete_error:
    Delete = False
End Function

Public Function Insert() As Boolean
    Dim sSql As String
    Dim rcvid As String
    Dim rcvnm As String
    Dim dSql As String
    
    
    CheckBundleNo
    
    rcvid = IIf(mvarDivCd = "0", mvarRcvId, "")
    rcvnm = IIf(mvarDivCd = "1", mvarRcvNm, "")
    
    
    dSql = " delete " & T_BBS411 & _
           " WHERE " & _
                     DBW("bldsrc=", mvarBldSrc) & _
           " AND " & DBW("bldyy=", mvarBldYY) & _
           " AND " & DBW("bldno=", mvarBldNo)
            

    sSql = "insert into " & T_BBS411 & _
           "(bldsrc,bldyy,bldno,rcvdt,rcvid," & _
            "rcvnm,divcd,centercd,returndt,returnid," & _
            "usedt,useid,bldbundleno) " & _
           "values(" & _
            DBV("bldsrc", mvarBldSrc) & "," & DBV("bldyy", mvarBldYY) & "," & DBV("bldno", mvarBldNo) & "," & _
            DBV("rcvdt", mvarRcvDt) & "," & DBV("rcvid", rcvid) & "," & _
            DBV("rcvnm", rcvnm) & "," & DBV("divcd", mvarDivCd) & "," & DBV("centercd", mvarCenterCd) & "," & _
            DBV("returndt", mvarReturnDt) & "," & DBV("returnid", mvarReturnId) & "," & _
            DBV("usedt", mvarUseDt) & "," & DBV("useid", mvarUseId) & "," & DBV("bldbundleno", mvarBldBundleNo) & ")"

    
On Error GoTo Insert_error

    DBConn.Execute (dSql)
    DBConn.Execute (sSql)
    
    Insert = True
    Exit Function
    
Insert_error:
    Insert = False
End Function

Public Function DoReturn() As Boolean
    Dim sSql As String
    
    sSql = " update " & T_BBS411 & " " & _
           " set   " & DBW("returndt", mvarReturnDt, 3) & _
           "       " & DBW("returnid", mvarReturnId, 2) & _
           " WHERE " & DBW("bldsrc", mvarBldSrc, 2) & _
           " AND   " & DBW("bldyy", mvarBldYY, 2) & _
           " AND   " & DBW("bldno", mvarBldNo, 2)
    
On Error GoTo DoReturn_error

    DBConn.Execute (sSql)
    
    DoReturn = True
    Exit Function
    
DoReturn_error:
    DoReturn = False
End Function

Public Function DoReturnAll(ByVal Bldsrc As String, ByVal rcvdt As String, ByVal returndt As String, ByVal returnid As String) As Boolean
    Dim sSql As String
    
    sSql = " update " & T_BBS411 & " " & _
           " set   " & DBW("usedt", returndt, 3) & _
           "       " & DBW("useid", returnid, 2) & _
           " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
           " AND   " & DBW("rcvdt", rcvdt, 2) & _
           " AND   " & DBW("returndt", "", 2) & _
           " AND   " & DBW("usedt", "", 2) & _
           " AND   " & DBW("divcd", "1", 2)

On Error GoTo DoReturnAll_error

    DBConn.Execute (sSql)
    
    DoReturnAll = True
    Exit Function
    
DoReturnAll_error:
    DoReturnAll = False
End Function

Public Function CancelReturn(ByVal Bldsrc As String, ByVal rcvdt As String, ByVal returndt As String) As Boolean
    Dim sSql1 As String
    Dim sSql2 As String
    
    sSql1 = " update " & T_BBS411 & " " & _
            " set " & DBW("returndt", "", 3) & _
            "     " & DBW("returnid", "", 2) & _
            " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
            " AND   " & DBW("rcvdt", rcvdt, 2) & _
            " AND   " & DBW("returndt", returndt, 2)
    
    sSql2 = " update " & T_BBS411 & " " & _
            " set " & DBW("usedt", "", 3) & _
            "     " & DBW("useid", "", 2) & _
            " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
            " AND   " & DBW("rcvdt", rcvdt, 2) & _
            " AND   " & DBW("usedt", returndt, 2) & _
            " AND   " & DBW("divcd", "1", 2)
            
On Error GoTo CancelReturn_error

    DBConn.Execute (sSql1)
    DBConn.Execute (sSql2)
    
    CancelReturn = True
    Exit Function
    
CancelReturn_error:
    CancelReturn = False
End Function

Public Function GetRcvDtList(dateList() As String) As Integer
    Dim DrRS As Recordset
    Dim sSql As String
    Dim i As Long
    
    Set DrRS = New Recordset
    
    sSql = "SELECT distinct rcvdt " & _
           "FROM " & T_BBS411 & " " & _
           "ORDER BY rcvdt desc"
           
    Call DrRS.Open(sSql, DBConn)
'    If DrRS.DBerror = True Then
    If DrRS.EOF Then
        GetRcvDtList = -1
        Set DrRS = Nothing
        Exit Function
    End If
    
    GetRcvDtList = DrRS.RecordCount
    
    If DrRS.RecordCount <= 0 Then Exit Function
    
    ReDim Preserve dateList(DrRS.RecordCount - 1)
    
    For i = 0 To DrRS.RecordCount - 1
        dateList(i) = DrRS.Fields("rcvdt").Value & ""
        DrRS.MoveNext
    Next i
    
    Set DrRS = Nothing
End Function

Public Function IsExistNotUsed(ByVal rcvdt As String) As Boolean
    Dim DrRS As Recordset
    Dim sSql As String
    
    sSql = " SELECT count(*) cnt " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("rcvdt<", rcvdt) & _
           " AND   (" & DBW("usedt", "", 2) & " AND " & DBW("returndt", "", 2) & ") "
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        IsExistNotUsed = True
        Set DrRS = Nothing
        Exit Function
    End If
    
    If DrRS.Fields("cnt").Value & "" > 0 Then
        IsExistNotUsed = True
    Else
        IsExistNotUsed = False
    End If
    
    Set DrRS = Nothing
End Function

Public Function IsExistUseable(ByVal rcvdt As String, ByVal fg As String) As Boolean
    Dim DrRS As Recordset
    Dim sSql As String
    
    sSql = " SELECT count(*) cnt " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("rcvdt" & fg, rcvdt) & _
           " AND   (" & DBW("usedt", "", 2) & " AND " & DBW("returndt", "", 2) & ") "
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        IsExistUseable = False
        Set DrRS = Nothing
        Exit Function
    End If
    
    If DrRS.Fields("cnt").Value & "" > 0 Then
        IsExistUseable = True
    Else
        IsExistUseable = False
    End If
    
    Set DrRS = Nothing
End Function

Public Function IsEnter(ByVal rcvdt As String) As Boolean
    Dim DrRS As Recordset
    Dim sSql As String
    
    sSql = " SELECT count(*) cnt " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("rcvdt", rcvdt, 2)
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        IsEnter = True
        Set DrRS = Nothing
        Exit Function
    End If
    
    If DrRS.Fields("cnt").Value & "" > 0 Then
        IsEnter = True
    Else
        IsEnter = False
    End If
    
    Set DrRS = Nothing
End Function

Public Function GetNotUsedCnt(ByVal rcvdt As String, Optional fg As String = "=") As Long
    Dim DrRS As Recordset
    Dim sSql As String
    
    sSql = " SELECT count(*) cnt " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("rcvdt" & fg, rcvdt) & _
           " AND   (" & DBW("usedt", "", 2) & " or " & DBW("returndt", "", 2) & ") "
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        GetNotUsedCnt = -1
    Else
        GetNotUsedCnt = DrRS.Fields("cnt").Value & ""
    End If
    
    Set DrRS = Nothing
End Function

Public Function GetReturnCnt(ByVal rcvdt As String, Optional fg As String = "=") As Long
    Dim DrRS As Recordset
    Dim sSql As String
    
    sSql = " SELECT count(*) cnt " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("rcvdt" & fg, rcvdt) & _
           " AND   (" & DBW("usedt", "", 2) & _
                    " AND " & DBW("returndt<>", "") & ") "
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        GetReturnCnt = -1
    Else
        GetReturnCnt = DrRS.Fields("cnt").Value & ""
'        DrRS.RsClose
    End If
    
    Set DrRS = Nothing
End Function

Public Function GetBldNo(ByVal centercd As String, ByVal usedt As String, ByVal useid As String, ByRef Bldno As String) As Boolean
    Dim sSql As String
    Dim tBldNo As String
    Dim DrRS As Recordset
    
    sSql = " SELECT min(bldbundleno) bldbundleno " & _
           " FROM " & T_BBS411 & " " & _
           " WHERE " & DBW("centercd", centercd, 2) & _
           " AND " & DBW("usedt", "", 2) & _
           " AND " & DBW("returndt", "", 2)
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        Set DrRS = Nothing
        Bldno = ""
        GetBldNo = False
        Exit Function
    Else
        tBldNo = DrRS.Fields("bldbundleno").Value & ""
        Bldno = Mid(tBldNo, 1, 2) & "-" & Mid(tBldNo, 5, 6) & "-" & Mid(tBldNo, 7)
    End If
    Set DrRS = Nothing
    
    sSql = " update " & T_BBS411 & " " & _
           " set " & DBW("usedt", usedt, 3) & _
           "     " & DBW("useid", useid, 2) & _
           " WHERE " & DBW("bldbundleno", tBldNo, 2)
           
    On Error GoTo ErrTrap
    
    DBConn.Execute (sSql)
    GetBldNo = True
        
    Exit Function
ErrTrap:
    GetBldNo = False
    Bldno = ""
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetBloodPaper(ByVal rcvdt As String) As String
    GetBloodPaper = " SELECT divcd,centercd,bldno,usedt,returndt,bldsrc,bldyy " & _
                    " FROM " & T_BBS411 & " " & _
                    " WHERE " & DBW("rcvdt", rcvdt, 2) & _
                    " ORDER BY divcd,centercd,bldno,bldyy"
End Function

Public Function GetEnterList(ByVal rcvdt As String) As String
    GetEnterList = " SELECT bldsrc,rcvdt,divcd,rcvid,rcvnm,centercd,min(bldno) frno,max(bldno) tono " & _
                   " FROM " & T_BBS411 & " " & _
                   " WHERE " & DBW("rcvdt", rcvdt, 2) & _
                   " GROUP BY bldsrc,rcvdt,divcd,rcvid,rcvnm,centercd " & _
                   " ORDER BY divcd,centercd "
End Function

