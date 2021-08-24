VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSMSTStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public TestCd  As String          ' 검사항목 코드
Public SpcCd As String            ' 검체 코드
Public ApplyDt As String          ' 적용일
Public ApplySex As String         ' 적용성별
Public AgeFrom As Long            ' 적용나이(FROM)
Public AgeTo As Long              ' 적용나이(To)
Public RefValFrom As Double       ' 참고치(FROM)
Public RefValTo As Double         ' 참고치(To)
Public RefCd As String            ' 참고치(Alpha)
Public ExpDt As String            ' 폐기일
Public RefTxt As String           ' 다중참고치


Public Function IsExistSpecmens(ByVal Centercd As String, ByVal leg As String, ByVal Row As Long, ByVal Col As Long) As Boolean
    Dim i       As Long
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS206 & " " & _
           "WHERE " & _
                             DBW("centercd=", Centercd) & _
                    " AND " & DBW("legcd =", leg) & _
                    " AND ((" & DBW("rowno=", Row) & " AND " & DBW("colno >", Col) & _
                            ") or (" & DBW("rowno >", Row) & "))"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        IsExistSpecmens = True
'        Set RS = Nothing
'    End If
    If RS.EOF Then Set RS = Nothing: Exit Function
    
    IsExistSpecmens = False
    For i = 1 To RS.RecordCount
        If (RS.Fields("stscd").Value & "") <> "0" Then
            IsExistSpecmens = True
            Exit For
        End If
        RS.MoveNext
    Next i
    
    Set RS = Nothing
End Function

'--검체보관장소 마스터
Public Function InsertBBS003(ByVal Centercd As String, ByVal leg As String, ByVal Row As Long, _
                             ByVal Col As Long, ByVal Rk As String, ByVal Chk As Boolean) As Boolean
    Dim i           As Long
    Dim j           As Long
    Dim idx         As Long
    Dim SSQL        As String
    Dim orgrowcnt   As Long
    Dim orgcolcnt   As Long
    Dim QuerySQL()      As String
    Dim RS          As Recordset
    
    
    If IsExistSpecmens(Centercd, leg, Row, Col) = True Then
        MsgBox "보관된 검체가 있어서 변경할 수 없습니다", vbCritical, "알림"
        InsertBBS003 = False
        Exit Function
    End If
    
    ReDim QuerySQL(0)
    idx = 0
    
    If Chk = True Then
        '신규등록 ---------------------------------------------------------------------------
        QuerySQL(0) = " INSERT INTO " & T_BBS003 & "(centercd,legcd, rowcnt, colcnt,rmk) VALUES(" & _
                 DBV("centercd", Centercd, 1) & DBV("legcd", leg, 1) & DBV("rowcnt", Row, 1) & DBV("colcnt", Col, 1) & DBV("rmk", Rk) & ")"
        '보관검체마스터생성------------------------------------------------------------------
        For i = 1 To Row
            For j = 1 To Col
                idx = idx + 1
                ReDim Preserve QuerySQL(idx)
                QuerySQL(idx) = "insert into " & T_BBS206 & "(centercd,legcd,rowno,colno,spcyy,spcno,stscd) " & _
                              "values(" & _
                                    DBV("centercd", Centercd, 1) & DBV("legcd", leg, 1) & DBV("rowno", i, 1) & _
                                    DBV("colno", j, 1) & DBV("spcyy", "", 1) & _
                                    DBV("spcno", "0", 1) & DBV("stscd", "0") & ")"
            Next j
        Next i
    
    Else
        '기존에 있는 정보를 읽어온다.----------------------------------------------------------
        SSQL = " SELECT * FROM " & T_BBS003 & _
               " WHERE " & _
                                 DBW("centercd =", Centercd) & _
                       " AND " & DBW("legcd=", leg)
                       
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        
'        If RS.DBerror = True Then
'            Set RS = Nothing
'            InsertBBS003 = False
'            Exit Function
'        End If
            
        If RS.RecordCount < 1 Then
            MsgBox "보관장소 정보가 손상되었읍니다.", vbCritical, "오류"
            Set RS = Nothing
            InsertBBS003 = False
        End If
        
        orgrowcnt = RS.Fields("rowcnt").Value & ""
        orgcolcnt = RS.Fields("colcnt").Value & ""
        
        Set RS = Nothing
        
        
        '수정---------------------------------------------------------------------------------
        QuerySQL(0) = " UPDATE " & T_BBS003 & " " & _
                    " SET " & DBW("rowcnt", Row, 3) & _
                              DBW("colcnt", Col, 3) & _
                              DBW("rmk", Rk, 2) & _
                    " WHERE " & _
                                      DBW("centercd=", Centercd) & _
                            " AND " & DBW("legcd=", leg)
                    
                    
        '보관검체마스터수정-------------------------------------------------------------------
        For i = 1 To IIf(orgrowcnt > Row, orgrowcnt, Row)
            For j = 1 To IIf(orgcolcnt > Col, orgcolcnt, Col)
                If i > orgrowcnt Then
                    '--- 생성 ---
                    idx = idx + 1
                    ReDim Preserve QuerySQL(idx)
                    QuerySQL(idx) = "insert into " & T_BBS206 & _
                                  "(centercd,legcd,rowno,colno,spcyy,spcno,stscd) " & _
                                  "values(" & _
                                  DBV("centercd", Centercd, 1) & DBV("legcd", leg, 1) & DBV("rowno", i, 1) & DBV("colno", j, 1) & _
                                  DBV("spcyy", "", 1) & DBV("spcno", "0", 1) & DBV("stscd", "0") & ")"
                End If
                If i <= orgrowcnt And j > orgcolcnt Then
                    '--- 생성 ---
                    idx = idx + 1
                    ReDim Preserve QuerySQL(idx)
                    QuerySQL(idx) = "insert into " & T_BBS206 & _
                                  "(centercd,legcd,rowno,colno,spcyy,spcno,stscd) " & _
                                  "values(" & _
                                  DBV("centercd", Centercd, 1) & DBV("legcd", leg, 1) & DBV("rowno", i, 1) & _
                                  DBV("colno", j, 1) & DBV("spcyy", "", 1) & DBV("spcno", "0", 1) & _
                                  DBV("stscd", "0") & ")"
                                  
                End If
                If i > Row Then
                    '---삭제---
                    idx = idx + 1
                    ReDim Preserve QuerySQL(idx)
                    QuerySQL(idx) = "delete FROM " & T_BBS206 & " " & _
                                  "WHERE " & _
                                                  DBW("centercd=", Centercd) & " " & _
                                         "AND " & DBW("legcd=", leg) & " " & _
                                         "AND " & DBW("rowno=", i) & " " & _
                                         "AND " & DBW("colno=", j) & " "
                End If
                If i <= Row And j > Col Then
                    '---삭제---
                    idx = idx + 1
                    ReDim Preserve QuerySQL(idx)
                    QuerySQL(idx) = "delete FROM " & T_BBS206 & " " & _
                                  "WHERE " & _
                                                  DBW("centercd=", Centercd) & " " & _
                                         "AND " & DBW("legcd=", leg) & " " & _
                                         "AND " & DBW("rowno=", i) & " " & _
                                         "AND " & DBW("colno=", j) & " "
                End If
            Next j
        Next i
        
    End If
    
On Error GoTo InsertBBS003_Error
    DBConn.BeginTrans
    
    For i = LBound(QuerySQL) To UBound(QuerySQL)
        DBConn.Execute QuerySQL(i)
    Next i
    
    DBConn.CommitTrans
    InsertBBS003 = True
    Exit Function

InsertBBS003_Error:
    DBConn.RollbackTrans
    InsertBBS003 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetBBS003(ByVal Centercd As String, Optional leg As String = "") As Recordset
    Dim SSQL As String
    
    If leg <> "" Then
    '내용을 가져가자...
        SSQL = "SELECT * FROM " & T_BBS003 & _
                 " WHERE " & _
                                   DBW("centercd=", Centercd) & _
                         " AND " & DBW("legcd=", leg)
    Else
        SSQL = "SELECT * FROM " & T_BBS003 & _
                 " WHERE " & DBW("centercd=", Centercd)
    End If
    
    Set GetBBS003 = New Recordset
    GetBBS003.Open SSQL, DBConn
    
'    If GetBBS003.DBerror = True Then
'        'dbconn.DisplayErrors
'        Exit Function
'    End If
End Function

'-- 검체보관일수 관리
Public Function InsertB001(ByVal Val As String, _
                           ByVal Field As String, ByVal Chk As Boolean) As Boolean
    Dim SSQL As String
    
    If Chk = True Then  '등록 시키자
        SSQL = " INSERT INTO " & T_COM003 & "(cdindex, cdval1, field1) VALUES(" & _
                   DBV("cdindex", "B001", 1) & _
                   DBV("cdval", Val, 1) & _
                   DBV("field1", Field) & ")"
    Else '수정 하자...
        SSQL = " UPDATE " & T_COM003 & " SET " & DBW("field1=", Field) & " " & _
                                         " WHERE " & DBW("cdindex=", "B001") & _
                                         " AND " & DBW("cdval1=", Val)
    End If
    
On Error GoTo InsertB001_Error
    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    InsertB001 = True
    Exit Function

InsertB001_Error:
    DBConn.RollbackTrans
    InsertB001 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetB001(Optional Val As String) As Recordset
    Dim SSQL As String
    
    If Val <> "" Then
        '내용을 가져가자...
        SSQL = "SELECT * FROM " & T_COM003 & _
                 " WHERE " & DBW("cdindex=", "B001") & _
                 " AND " & DBW("cdval1=", Val)
    Else
        SSQL = "SELECT * FROM " & T_COM003 & _
                 " WHERE " & DBW("cdindex=", "B001") & _
                 " ORDER BY cdval1 desc"
    End If
    
    Set GetB001 = New Recordset
    GetB001.Open SSQL, DBConn
    
'    If GetB001.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'-- X-M Step 관리
Public Function GetB002(ByVal Val As String) As Recordset
    Dim SSQL As String
    
    '내용을 가져가자...
    SSQL = "SELECT * FROM " & T_COM003 & _
             " WHERE " & DBW("cdindex=", "B002") & _
             " AND " & DBW("cdval1=", Val)
      
    Set GetB002 = New Recordset
    GetB002.Open SSQL, DBConn
    
'    If GetB002.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'-- 접수취소가 가능한 Status 관리
Public Function InsertB011(ByVal Val As String, ByVal Fld1 As String, ByVal Fld2 As String, ByVal Chk As Boolean) As Boolean
    Dim SSQL As String
    
    If Chk = True Then  '등록 시키자
        SSQL = " INSERT INTO " & T_COM003 & "(cdindex, cdval1, field1,field2) " & _
                 " VALUES(" & _
                         DBV("cdindex", "B011", 1) & _
                         DBV("cdval1", Val, 1) & _
                         DBV("field1", Fld1, 1) & _
                         DBV("field2", Fld2) & ") "
    Else '수정 하자...
        SSQL = " UPDATE " & T_COM003 & " SET " & DBW("field1=", Fld1, 3) & _
                                            "  " & DBW("field2=", Fld2, 2) & _
                                         "  WHERE " & _
                                                            DBW("cdindex=", "B011") & _
                                                  " AND " & DBW("cdval1=", Val)
    End If
    
On Error GoTo InsertB011_Error
    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    InsertB011 = True
    Exit Function

InsertB011_Error:
    DBConn.RollbackTrans
    InsertB011 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetB011(Optional Val As String) As Recordset
    Dim SSQL As String
    
    If Val <> "" Then
        '내용을 가져가자...
        SSQL = " SELECT * FROM " & T_COM003 & _
                 " WHERE " & DBW("cdindex=", "B011") & _
                 " AND " & DBW("cdval1=", Val)
    Else
        SSQL = " SELECT * FROM " & T_COM003 & _
                 " WHERE " & DBW("cdindex=", "B011") & _
                 " ORDER BY cdval1 desc"
    End If
    
    Set GetB011 = New Recordset
    GetB011.Open SSQL, DBConn
    
'    If GetB011.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function
Public Function InsertBBS001_Ghil(ByVal TCd As String, ByVal AppDt As String, _
                             ByVal TNm As String, ByVal ANm5 As String, ByVal ANm10 As String, _
                             ByVal CCd As String, ByVal Vol As Long, _
                             ByVal Tfg As String, ByVal Xm As String, _
                             ByVal MCd As String, ByVal NTdiv As String, _
                             ByVal Chk As Boolean, ByVal OrdDiv As String, Optional ByVal Edt As String) As Boolean
    Dim SSQL As String
    
    If Chk = True Then  '등록 시키자
        SSQL = " INSERT INTO " & T_BBS001 & "(testcd, applydt, testnm, abbrnm5, abbrnm10, compocd, " & _
                                              " volumn, testdiv, xmethod, expdt, matchcd,newtestdiv,orddiv) VALUES( " & _
                                              DBV("testcd", TCd, 1) & DBV("applydt", AppDt, 1) & DBV("testnm", TNm, 1) & _
                                              DBV("abbrnm5", ANm5, 1) & DBV("abbrnm10", ANm10, 1) & DBV("compocd", CCd, 1) & _
                                              DBV("volumn", Vol, 1) & DBV("testdiv", Tfg, 1) & DBV("xmethod", Xm, 1) & _
                                              DBV("expdt", Edt, 1) & DBV("matchcd", MCd, 1) & DBV("newtestdiv", NTdiv, 1) & DBV("orddiv", OrdDiv) & ") "
    Else '수정 하자...
        SSQL = " UPDATE " & T_BBS001 & " SET " & DBW("testnm", TNm, 3) & _
                                                  DBW("abbrnm5", ANm5, 3) & _
                                                  DBW("abbrnm10", ANm10, 3) & _
                                                  DBW("compocd", CCd, 3) & _
                                                  DBW("volumn", Vol, 3) & _
                                                  DBW("testdiv", Tfg, 3) & _
                                                  DBW("xmethod", Xm, 3) & _
                                                  DBW("expdt", Edt, 3) & _
                                                  DBW("matchcd", MCd, 3) & _
                                                  DBW("newtestdiv", NTdiv, 3) & _
                                                  DBW("orddiv", OrdDiv, 2) & _
                                         " WHERE " & _
                                                            DBW("testcd=", TCd) & _
                                                  " AND " & DBW("applydt=", AppDt)
    End If
    
On Error GoTo InsertBBS001_Error
    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    InsertBBS001_Ghil = True
    Exit Function

InsertBBS001_Error:
    DBConn.RollbackTrans
    InsertBBS001_Ghil = False
    MsgBox Err.Description, vbExclamation
End Function


'--수혈처방 마스터 등록
Public Function InsertBBS001(ByVal TCd As String, ByVal AppDt As String, _
                             ByVal TNm As String, ByVal ANm5 As String, ByVal ANm10 As String, _
                             ByVal CCd As String, ByVal Vol As Long, _
                             ByVal Tfg As String, ByVal Xm As String, _
                             ByVal MCd As String, ByVal NTdiv As String, _
                             ByVal Chk As Boolean, Optional ByVal Edt As String) As Boolean
    Dim SSQL As String
    
    If Chk = True Then  '등록 시키자
        SSQL = " INSERT INTO " & T_BBS001 & "(testcd, applydt, testnm, abbrnm5, abbrnm10, compocd, " & _
                                              " volumn, testdiv, xmethod, expdt, matchcd,newtestdiv) VALUES( " & _
                                              DBV("testcd", TCd, 1) & DBV("applydt", AppDt, 1) & DBV("testnm", TNm, 1) & _
                                              DBV("abbrnm5", ANm5, 1) & DBV("abbrnm10", ANm10, 1) & DBV("compocd", CCd, 1) & _
                                              DBV("volumn", Vol, 1) & DBV("testdiv", Tfg, 1) & DBV("xmethod", Xm, 1) & _
                                              DBV("expdt", Edt, 1) & DBV("matchcd", MCd, 1) & DBV("newtestdiv", NTdiv) & ") "
    Else '수정 하자...
        SSQL = " UPDATE " & T_BBS001 & " SET " & DBW("testnm", TNm, 3) & _
                                                  DBW("abbrnm5", ANm5, 3) & _
                                                  DBW("abbrnm10", ANm10, 3) & _
                                                  DBW("compocd", CCd, 3) & _
                                                  DBW("volumn", Vol, 3) & _
                                                  DBW("testdiv", Tfg, 3) & _
                                                  DBW("xmethod", Xm, 3) & _
                                                  DBW("expdt", Edt, 3) & _
                                                  DBW("matchcd", MCd, 3) & _
                                                  DBW("newtestdiv", NTdiv, 2) & _
                                         " WHERE " & _
                                                            DBW("testcd=", TCd) & _
                                                  " AND " & DBW("applydt=", AppDt)
    End If
    
On Error GoTo InsertBBS001_Error
    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    InsertBBS001 = True
    Exit Function

InsertBBS001_Error:
    DBConn.RollbackTrans
    InsertBBS001 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetB004(ByVal Val As String) As Recordset
    Dim SSQL As String
   
    '내용을 가져가자...
    SSQL = " SELECT * FROM " & T_COM003 & _
             " WHERE " & _
                               DBW("cdindex=", "B004") & _
                     " AND " & DBW(" cdval1=", Val)
   
    
    Set GetB004 = New Recordset
    GetB004.Open SSQL, DBConn
    
'    If GetB004.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function LoadBBS001(ByVal TCd As String) As Recordset
    Dim SSQL As String

    '내용을 가져가자....
    SSQL = " SELECT * FROM " & T_BBS001 & _
             " WHERE " & DBW("testcd=", TCd) & _
             " ORDER BY applydt DESC"
            
    Set LoadBBS001 = New Recordset
    LoadBBS001.Open SSQL, DBConn

'    If LoadBBS001.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function GetBBS001(ByVal TCd As String, ByVal AppDt As String) As Recordset
    Dim SSQL As String

    '내용을 가져가자...
    SSQL = " SELECT * FROM " & T_BBS001 & _
             " WHERE " & DBW("testcd=", TCd) & _
             " AND " & DBW("applydt=", AppDt)
             

    Set GetBBS001 = New Recordset
    GetBBS001.Open SSQL, DBConn

'    If GetBBS001.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function GetPopup(Optional ByVal busidiv As Integer = -1) As String
    'Popup String
    If busidiv = 0 Then
        GetPopup = "SELECT DISTINCT testcd, testnm FROM " & T_BBS001 & " ORDER BY testcd "
    Else
        GetPopup = "SELECT " & F_AMTCD & " amtcd," & F_AMTNM & " amtnm FROM  " & T_HIS008
    End If
End Function

Public Function GetcboBBS001(ByVal Fld As String) As Recordset
    Dim SSQL As String
    
    '내용을 가져가자...
    SSQL = " SELECT * FROM " & T_COM003 & _
             " WHERE " & DBW("cdindex=", "B004") & _
             " AND " & DBW("field1=", Fld) & ""
      
    Set GetcboBBS001 = New Recordset
    GetcboBBS001.Open SSQL, DBConn
    
'    If GetcboBBS001.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function LoadAllBBS001() As Recordset
    Dim SSQL As String

    '내용을 가져가자...
    SSQL = "SELECT * FROM " & T_BBS001 & " ORDER BY testcd "

    Set LoadAllBBS001 = New Recordset
    LoadAllBBS001.Open SSQL, DBConn

'    If LoadAllBBS001.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'--Doner Screening 검사 적격치 마스터

Public Function LoadPopup(Optional ByVal busidiv As Integer = -1) As String
    'Popup String
    If busidiv = 0 Then
        LoadPopup = "SELECT testcd,testnm FROM  " & T_LAB001 & " ORDER BY testcd"
    Else
        LoadPopup = "SELECT DISTINCT a.testcd,a.testnm FROM " & T_LAB001 & " a, " & T_BBS002 & " b " & _
                    "WHERE a.testcd=b.testcd "
    End If
End Function

Public Function GetBBS002() As Recordset
    Dim SSQL As String
    
    '내용을 가져가자...
    SSQL = "SELECT * FROM " & T_BBS002 & ""
    
    Set GetBBS002 = New Recordset
    GetBBS002.Open SSQL, DBConn
    
'    If GetBBS002.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function GetTestNm(sTestCd As String) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT a.testcd, a.testnm FROM " & T_LAB001 & " a " & _
             "WHERE (a.expdt = '' or a.expdt is null) " & _
             "AND   " & DBW("a.testcd=", sTestCd) & " " & _
             "ORDER BY a.testcd "
    
    Set GetTestNm = New Recordset
    GetTestNm.Open SSQL, DBConn
    
'    If GetTestNm.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function InsertBBS002(ByVal cRef As clsBBSMSTStatement) As Boolean
    Dim SSQL As String
    
    SSQL = "Insert into " & _
            "" & T_BBS002 & " (testcd, spccd, applydt, applysex, agefrom, ageto, " & _
            "refvalfrom, refvalto, refcd, expdt) values("
    
    SSQL = SSQL & " " & DBV("testcd", cRef.TestCd) & ","
    SSQL = SSQL & " " & DBV("spccd", cRef.SpcCd) & ","
    SSQL = SSQL & " " & DBV("applydt", cRef.ApplyDt) & ","
    SSQL = SSQL & " " & DBV("applysex", cRef.ApplySex) & ","
    SSQL = SSQL & " " & DBV("agefrom", cRef.AgeFrom) & ","
    SSQL = SSQL & " " & DBV("ageto", cRef.AgeTo) & ", "
    SSQL = SSQL & " " & DBV("refvalfrom", cRef.RefValFrom) & ","
    SSQL = SSQL & " " & DBV("refvalto", cRef.RefValTo) & ","
    SSQL = SSQL & " " & DBV("refcd", cRef.RefCd) & ","
    SSQL = SSQL & " " & DBV("expdt", cRef.ExpDt) & ") "
    
On Error GoTo InsertBBS002_Error
    DBConn.Execute SSQL
    InsertBBS002 = True
    Exit Function

InsertBBS002_Error:
    DBConn.RollbackTrans
    InsertBBS002 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function getSpcs(ByVal sTestCd As String) As Recordset
    Dim SSQL As String

    SSQL = " SELECT a.spccd as spccd, b.field3 as spcnm FROM " & T_LAB004 & " a," & T_LAB032 & " b " & _
             " WHERE " & DBW("a.testcd=", sTestCd) & _
             " AND " & DBW("b.cdindex=", "C215") & _
             " AND a.spccd=b.cdval1 ORDER BY spccd"
    
    Set getSpcs = New Recordset
    getSpcs.Open SSQL, DBConn
    
'    If getSpcs.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function getApplydt(ByVal sTestCd As String, ByVal sSpcCd As String) As Recordset
    Dim SSQL As String
        
    SSQL = " SELECT applydt FROM " & T_BBS002 & _
             " WHERE " & DBW("testcd=", sTestCd) & _
             " AND  " & DBW("spccd=", sSpcCd) & _
             " GROUP BY applydt ORDER BY applydt"
    
    Set getApplydt = New Recordset
    getApplydt.Open SSQL, DBConn
        
'    If getApplydt.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function getReference(sTestCd As String, sSpcCd As String, sAppDt As String) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_BBS002 & _
             " WHERE " & DBW("testcd=", sTestCd) & _
             " AND   " & DBW("spccd=", sSpcCd) & _
             " AND   " & DBW("applydt=", sAppDt) & _
             " ORDER BY applysex, agefrom"
        
    Set getReference = New Recordset
    getReference.Open SSQL, DBConn
        
'    If getReference.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function UpdateBBS002(ByVal cRef As clsBBSMSTStatement) As Boolean
    Dim SSQL As String
    
    SSQL = "update " & _
            "" & T_BBS002 & " "
    SSQL = SSQL & "Set " & DBW("refvalfrom", cRef.RefValFrom, 3) & DBW("refvalto", cRef.RefValTo, 3)
    SSQL = SSQL & "    " & DBW("refcd", cRef.RefCd, 3) & DBW("expdt", cRef.ExpDt, 2)
    SSQL = SSQL & "WHERE " & DBW("testcd=", cRef.TestCd)
    SSQL = SSQL & "AND   " & DBW("spccd=", cRef.SpcCd)
    SSQL = SSQL & "AND   " & DBW("applydt=", cRef.ApplyDt)
    SSQL = SSQL & "AND   " & DBW("applysex=", cRef.ApplySex)
    SSQL = SSQL & "AND   " & DBW("agefrom=", cRef.AgeFrom)
    SSQL = SSQL & "AND   " & DBW("ageto=", cRef.AgeTo)
        
On Error GoTo UpdateBBS002_Error
    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    UpdateBBS002 = True
    Exit Function

UpdateBBS002_Error:
    DBConn.RollbackTrans
    UpdateBBS002 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function DeleteBBS002(ByVal TCd As String, ByVal SCd As String, ByVal ADt As String) As Boolean
    Dim SSQL As String
    
    SSQL = "DELETE FROM " & _
            "" & T_BBS002 & " "
    
    SSQL = SSQL & "WHERE  " & DBW("testcd =", TCd) & " "
    SSQL = SSQL & "AND    " & DBW("spccd=", SCd) & " "
    SSQL = SSQL & "AND    " & DBW("applydt= ", ADt) & " "
    
        
    On Error GoTo DeleteBBS002_Error
    DBConn.Execute SSQL
    DeleteBBS002 = True
    Exit Function

DeleteBBS002_Error:
    DBConn.RollbackTrans
    DeleteBBS002 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function ChkBBS002(ByVal TCd As String, ByVal SCd As String, ByVal ADt As String) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT * FROM " & _
            "" & T_BBS002 & " "
    
    SSQL = SSQL & "WHERE  " & DBW("testcd=", TCd) & " "
    SSQL = SSQL & "AND    " & DBW("spccd=", SCd) & " "
    SSQL = SSQL & "AND    " & DBW("applydt=", ADt) & " "
    
        
    Set ChkBBS002 = New Recordset
    ChkBBS002.Open SSQL, DBConn
        
'    If ChkBBS002.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function LoadB002() As Recordset
    Dim SSQL As String

    '내용을 가져가자...
    SSQL = " SELECT * FROM " & T_COM003 & _
             " WHERE " & DBW("cdindex=", "B004") & _
             " ORDER BY cdval1 desc"

    Set LoadB002 = New Recordset
    LoadB002.Open SSQL, DBConn

'    If LoadB002.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function


'-- Doner Screening 검사항목 마스터

Public Function GetB003(ByVal sSpcCd As String) As Recordset
    Dim SSQL As String

    SSQL = " SELECT field3  FROM " & T_COM003 & " " & _
             " WHERE " & DBW("cdindex=", "C215") & _
             " AND " & DBW("cdval1=", sSpcCd) & _
             " ORDER BY cdval1"
    
    Set GetB003 = New Recordset
    GetB003.Open SSQL, DBConn
    
'    If GetB003.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function LoadB003(Optional sTestCd As String) As Recordset
    Dim SSQL As String
    
    If sTestCd <> "" Then
        SSQL = " SELECT * FROM " & T_COM003 & " " & _
                " WHERE " & DBW("cdindex=", "B003") & _
                " AND " & DBW("cdval1=", sTestCd) & ""
    Else
        SSQL = " SELECT * FROM " & T_COM003 & _
                " WHERE " & DBW("cdindex=", "B003") & _
                " ORDER BY cdval1 desc"
    End If
    
    Set LoadB003 = New Recordset
    LoadB003.Open SSQL, DBConn
    
'    If LoadB003.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

Public Function DeleteB003(ByVal sTestCd As String) As Boolean
    Dim SSQL As String

    SSQL = " DELETE FROM " & T_COM003 & " " & _
           " WHERE " & DBW("cdindex=", "B003") & _
           " AND   " & DBW("cdval1=", sTestCd) & ""
    
    On Error GoTo DeleteB003_Error
    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    DeleteB003 = True
    Exit Function

DeleteB003_Error:
    DBConn.RollbackTrans
    DeleteB003 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function InsertB003(ByVal Val As String, ByVal Fld1 As String, ByVal Fld2 As String, ByVal Chk As Boolean) As Boolean
    Dim strSql As String
    
    If Chk = True Then  '등록 시키자
        strSql = " INSERT INTO " & T_COM003 & "(cdindex, cdval1, field1,field2) VALUES(" & _
                                               DBV("cdindex", "B003", 1) & _
                                               DBV("cdval1", Val, 1) & _
                                               DBV("field1", Fld1, 1) & _
                                               DBV("field2", Fld2) & ")"
    Else '수정 하자...
        strSql = " UPDATE " & T_COM003 & " SET " & DBW("field1", Fld1, 3) & DBW("field2=", Fld2) & " " & _
                                         "  WHERE " & DBW("cdindex=", "B003") & _
                                         "  AND   " & DBW("cdval1=", Val) & " "
    End If
    
On Error GoTo InsertB003_Error
    DBConn.BeginTrans
    DBConn.Execute strSql
    DBConn.CommitTrans
    InsertB003 = True
    Exit Function

InsertB003_Error:
    DBConn.RollbackTrans
    InsertB003 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetABOList() As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT testcd,applydt,testnm,expdt " & _
           "FROM " & T_LAB001 & " " & _
           "WHERE " & DBW("testdiv=", "3")
           
    Set GetABOList = New Recordset
    Call GetABOList.Open(SSQL, DBConn)
'    If GetABOList.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetABOList = Nothing
'    End If
End Function

Public Function Save_frmCTControl(NewFg As Boolean, KeyCode As String, keyname As String, strFields As String, remark As String) As Boolean
    Dim SSQL As String
    
    If NewFg Then
        SSQL = "insert into " & T_COM001 & "(cdindex, cdval1,cdval2, field1, text1, text2) " & _
               " values (" & _
                       DBV("cdindex", "BC2", 1) & DBV("cdval1", KeyCode, 1) & DBV("cdval2", "0", 1) & _
                       DBV("field1", keyname, 1) & DBV("text1", strFields, 1) & DBV("text2", remark) & ")"
    Else
        SSQL = "update " & T_COM001 & " set " & DBW("field1", keyname, 3) & DBW("text1", strFields, 3) & _
                                                DBW("text2", remark, 2) & _
               " WHERE " & DBW("cdindex=", "BC2") & _
                       " AND " & DBW("cdval1=", KeyCode) & _
                       " AND " & DBW("cdval2=", "0")
    End If
    
On Error GoTo DBExecError

    DBConn.BeginTrans
    DBConn.Execute (SSQL)
    DBConn.CommitTrans
    Save_frmCTControl = True
    
    Exit Function

DBExecError:
    DBConn.RollbackTrans
    Save_frmCTControl = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Get_frmCtControlLoad() As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_COM001 & _
           " WHERE " & DBW("cdindex=", "BC2") & _
           " AND NOT " & DBW("cdval1=", "BC2") & _
           " AND " & DBW("cdval2=", "0")
           
    Set Get_frmCtControlLoad = New Recordset
    Call Get_frmCtControlLoad.Open(SSQL, DBConn)
'    If Get_frmCtControlLoad.DBerror Then
'        'dbconn.DisplayErrors
'        Set Get_frmCtControlLoad = Nothing
'    End If
End Function

Public Function Get_frmCtControlTable() As Recordset
    Dim SSQL As String
   
    SSQL = " SELECT text1 FROM " & T_COM001 & _
           " WHERE " & DBW("cdindex=", "BC2") & _
           " AND " & DBW("cdval1=", "BC2") & _
           " AND " & DBW("cdval2=", "0")
           
    Set Get_frmCtControlTable = New Recordset
    Call Get_frmCtControlTable.Open(SSQL, DBConn)
'    If Get_frmCtControlTable.DBerror Then
'        'dbconn.DisplayErrors
'        Set Get_frmCtControlTable = Nothing
'    End If
End Function
Public Function Get_frmCtControlTableKey(ssTableKey As String) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_COM001 & _
           " WHERE " & DBW("cdindex=", "BC2") & _
           " AND " & DBW("cdval1=", Trim(ssTableKey)) & _
           " AND " & DBW("cdval2=", "0")
           
    Set Get_frmCtControlTableKey = New Recordset
    Call Get_frmCtControlTableKey.Open(SSQL, DBConn)
'    If Get_frmCtControlTableKey.DBerror Then
'        'dbconn.DisplayErrors
'        Set Get_frmCtControlTableKey = Nothing
'    End If
End Function

Public Function Get_frmCtControlTableKeyCode(KeyCode As String) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_COM001 & _
           " WHERE " & DBW("cdindex=", "BC2") & _
           " AND " & DBW("cdval1=", Trim(KeyCode)) & _
           " AND " & DBW("cdval2=", "0")
    Set Get_frmCtControlTableKeyCode = New Recordset
    Call Get_frmCtControlTableKeyCode.Open(SSQL, DBConn)
    
'    If Get_frmCtControlTableKeyCode.DBerror Then
'        'dbconn.DisplayErrors
'        Set Get_frmCtControlTableKeyCode = Nothing
'    End If
End Function

Public Function Get_FrmBBS861(CDINDEX As String) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * " & _
           " FROM " & T_COM001 & " " & _
           " WHERE " & DBW("cdindex=", "BC2") & _
           " AND " & DBW("cdval1=", CDINDEX) & _
           " AND " & DBW("cdval2=", "0")
           
    Set Get_FrmBBS861 = New Recordset
    Call Get_FrmBBS861.Open(SSQL, DBConn)
'    If Get_FrmBBS861.DBerror Then
'        'dbconn.DisplayErrors
'        Set Get_FrmBBS861 = Nothing
'    End If
End Function


