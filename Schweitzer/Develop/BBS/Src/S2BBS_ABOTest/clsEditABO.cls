VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEditABO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private CODE_ABOFRONT As String
'Private CODE_ABOBACK As String
'Private CODE_RH As String
'Private CODE_ABOSUB As String
'Private CODE_RHSUB As String

'삭제대상~~~
'Public Function GetRemark() As Recordset

'    Dim strSQL As String
'
'    strSQL = "SELECT cdval1, field1,text1 " & _
'           "FROM " & T_LAB034 & " " & _
'           "WHERE " & DBW("cdindex = ", "C403") & " "
'
'    Set GetRemark = New Recordset
'    Call GetRemark.Open(strSQL, DBConn)
'End Function

'삭제대상~~~
'Public Function GetRemarkNm(ByVal RmkCd As String) As String
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'    strSQL = "SELECT cdval1, field1 " & _
'           "FROM " & T_LAB034 & " " & _
'           "WHERE " & DBW("cdindex = ", "C403") & " " & _
'           "AND   " & DBW("cdval1  = ", RmkCd) & " "
'
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
''        dbconn.DisplayErrors
'        GetRemarkNm = ""
'    Else
'        With Rs
'            If .RecordCount < 1 Then
'                GetRemarkNm = ""
'            Else
'                GetRemarkNm = .Fields("field1").Value & ""
'            End If
'        End With
'        Set Rs = Nothing
'    End If
'End Function

'삭제대상~~~~
'Public Function GetComment() As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT cdval1, text1 " & _
'           "FROM " & T_LAB034 & " " & _
'           "WHERE " & DBW("cdindex = ", "C404") & " "
'
'    Set GetComment = New Recordset
'    Call GetComment.Open(strSQL, DBConn)
'End Function

'삭제대상~~~~
'Public Function GetCommentText(ByVal code As String) As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT cdval1,field1,text1 " & _
'           "FROM " & T_LAB034 & " " & _
'           "WHERE " & DBW("cdindex = ", "C404") & " " & _
'           "AND   " & DBW("cdval1  = ", code) & " "
'
'    Set GetCommentText = New Recordset
'    Call GetCommentText.Open(strSQL, DBConn)
'End Function

'삭제대상~~~
'Public Function GetAccessInfo(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String) As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT workarea,accdt,accseq,ptid,ageday,sex,bedindt,deptcd,majdoct,stscd," & _
'                  "coldt,coltm,colid,rcvdt,rcvtm,rcvid," & _
'                  "reqtotcnt,reqinputcnt,wardid,statfg,rmkcd " & _
'           "FROM " & T_LAB201 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " "
'    Set GetAccessInfo = New Recordset
'    Call GetAccessInfo.Open(strSQL, DBConn)
'End Function

'Public Function GetOrderInfo(ByVal ptid As String, ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, ByVal ordcd As String) As Recordset
'    Dim strSQL As String
'
'    strSQL = " select examdt,examtm,examdoct from " & T_LAB101 & " a, " & T_LAB102 & " b " & _
'           " where a.ptid=b.ptid " & _
'           " and a.orddt=b.orddt " & _
'           " and a.ordno=b.ordno " & _
'           " and " & DBW("a.ptid=", ptid) & _
'           " and " & DBW("b.workarea=", workarea) & _
'           " and " & DBW("b.accdt=", accdt) & _
'           " and " & DBW("b.accseq=", accseq) & _
'           " and " & DBW("b.ordcd=", ordcd)
'    Set GetOrderInfo = New Recordset
'    GetOrderInfo.Open strSQL, DBConn
'End Function

'Public Function IsExistABO(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'    strSQL = "SELECT count(*) cnt /*a.testcd,b.testdiv*/ " & _
'           "FROM " & T_LAB302 & " a," & T_LAB001 & " b " & _
'           "WHERE " & DBW("a.workarea=", workarea) & " " & _
'           "AND   " & DBW("a.accdt   =", accdt) & " " & _
'           "AND   " & DBW("a.accseq  =", accseq) & " " & _
'           "AND    a.testcd=b.testcd " & _
'           "AND    b.applydt=(SELECT max(c.applydt) FROM " & T_LAB001 & " c WHERE b.testcd=c.testcd) " & _
'           "AND   " & DBW("b.testdiv=", "3") & " "
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
''        dbconn.DisplayErrors
'        IsExistABO = False
'    Else
'        With Rs
'            If .RecordCount < 1 Then
'                IsExistABO = False
'            Else
'                IsExistABO = ((.Fields("cnt").Value & "") > 0)
'            End If
'        End With
'    End If
'    Set Rs = Nothing
'End Function

'Public Function GetABOResultInfo(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String) As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT * " & _
'           "FROM " & T_BBS303 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " " & _
'           "AND seq=(SELECT max(seq) " & _
'                    "FROM " & T_BBS303 & " " & _
'                    "WHERE " & DBW("workarea=", workarea) & " " & _
'                    "AND   " & DBW("accdt   =", accdt) & " " & _
'                    "AND   " & DBW("accseq  =", accseq) & ")"
'
'    Set GetABOResultInfo = New Recordset
'    Call GetABOResultInfo.Open(strSQL, DBConn)
'End Function

'Public Function ModifyABOResult(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                ByVal typing As Long, _
'                                ByVal pABO As String, ByVal pRH As String, ByVal pABOSUB As String, ByVal pRHSUB As String, _
'                                ByVal EmpID As String, ByVal vfydt As String, ByVal vfytm As String) As Boolean
'    '-----------------------------------------------------------------------------------------
'    'PARAMETER 의 용도
'    '-----------------------------------------------------------------------------------------
'    'typing   = 0     : Front Typing
'    '         = 1     : Back typing
'    '
'    'isVerify = false : 저장만한다. (id만 입력)
'    '         = true  : verify한다.
'    '-----------------------------------------------------------------------------------------
'    '저장시 주의사항
'    '-----------------------------------------------------------------------------------------
'    'LAB302의 RSTTYPE을 'F'로 만든다
'    '-----------------------------------------------------------------------------------------
'
'    Dim astrSql() As String
'    Dim idx As Long
'    '-----------------------------------------------------------------------------------------
'    '진행순서
'    '-----------------------------------------------------------------------------------------
'    'STEP 1) CHECK
'    'SETP 3) ABO 저장
'    'STEP 2) REMARK, COMMENT 저장
'    'STEP 4) LAB302저장(필요하면)
'    '-----------------------------------------------------------------------------------------
'
'On Error GoTo ModifyABOResult_error
'
'    DBConn.BeginTrans
'
'    'BBS303 을 하나 더 생성(복사본)
'    If Not Exec_InsertABO(workarea, accdt, accseq, True, typing, pABO, pRH, pABOSUB, pRHSUB, EmpID, vfydt, vfytm) Then
'        GoTo ModifyABOResult_error
'    End If
'
'    'BBS308 수정내역 생성
'    If Exec_ModifyDoubleCheck(workarea, accdt, accseq, astrSql, idx) = False Then
'        GoTo ModifyABOResult_error
'    End If
'
'    'LAB102 Status를 수정으로 변경
'    If Exec_ModifyOrderStatus(workarea, accdt, accseq, vfydt, vfytm, EmpID) = False Then
'        GoTo ModifyABOResult_error
'    End If
'
'    DBConn.CommitTrans
'    ModifyABOResult = True
'    Exit Function
'
'ModifyABOResult_error:
'    DBConn.RollbackTrans
'    ModifyABOResult = False
'End Function

'Public Function SaveABOResult(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                              ByVal isVerify As Boolean, _
'                              ByVal typing As Long, _
'                              ByVal pABO As String, ByVal pRH As String, ByVal pABOSUB As String, ByVal pRHSUB As String, _
'                              ByVal remark As String, ByVal comment As String, _
'                              ByVal EmpID As String, ByVal vfydt As String, ByVal vfytm As String) As Boolean
'    '-----------------------------------------------------------------------------------------
'    'PARAMETER 의 용도
'    '-----------------------------------------------------------------------------------------
'    'typing   = 0     : Front Typing
'    '         = 1     : Back typing
'    '
'    'isVerify = false : 저장만한다. (id만 입력)
'    '         = true  : verify한다.
'    '-----------------------------------------------------------------------------------------
'    '저장시 주의사항
'    '-----------------------------------------------------------------------------------------
'    'LAB302의 RSTTYPE을 'F'로 만든다
'    '-----------------------------------------------------------------------------------------
'
'    Dim astrSql() As String
'    Dim idx As Long
'
'    '-----------------------------------------------------------------------------------------
'    '진행순서
'    '-----------------------------------------------------------------------------------------
'    'STEP 1) CHECK
'    'SETP 3) ABO 저장
'    'STEP 2) REMARK, COMMENT 저장
'    'STEP 4) LAB302저장(필요하면)
'    '-----------------------------------------------------------------------------------------
'
''이넘은 체크할 필요 없음
''    If Not CheckABOResult(workarea, accdt, accseq, isVerify, typing, pABO, pRH, pABOSUB, pRHSUB, EmpID) Then
''        SaveABOResult = False
''        Exit Function
''    End If
'
'On Error GoTo SaveABOResult_error
'
'    DBConn.BeginTrans
'
'    'BBS303저장
'    If Not Exec_UpdateABO(workarea, accdt, accseq, isVerify, typing, pABO, pRH, pABOSUB, pRHSUB, EmpID, vfydt, vfytm) Then
'        GoTo SaveABOResult_error
'    End If
'
'    'BBS201 검체리마크 저장
'    If remark <> "" Then
'        If Not Exec_UpdateRemark(workarea, accdt, accseq, remark) Then
'            GoTo SaveABOResult_error
'        End If
'    End If
'
'    'LAB304 코멘트 저장
'    If comment <> "" Then
'        If Not Exec_InsertComment(workarea, accdt, accseq, EmpID, comment) Then
'            GoTo SaveABOResult_error
'        End If
'    End If
'
'    'BBS303에 Front, Back 둘다 입력되어 있으면 Vftdt박아주고 LAB302 도 업뎃 해준다.
'    If Exec_DoubleCheck(workarea, accdt, accseq, astrSql, idx) = False Then
'        GoTo SaveABOResult_error
'    End If
'
'    'LAB302가 Vfy가 되었으면 LAB102 Status 반영
'    If Exec_OrderStatus(workarea, accdt, accseq, vfydt, vfytm, EmpID) = False Then
'        GoTo SaveABOResult_error
'    End If
'
'    'LAB302가 Vfy가 되었으면 LAB201 Status반영
'    If Exec_AccessStatus(workarea, accdt, accseq, EmpID, vfydt, vfytm) = False Then
'        GoTo SaveABOResult_error
'    End If
'
'    DBConn.CommitTrans
'    SaveABOResult = True
'    Exit Function
'
'SaveABOResult_error:
'    DBConn.RollbackTrans
'    SaveABOResult = False
'End Function

'Private Function Exec_AccessStatus(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                   ByVal VfyId As String, ByVal vfydt As String, ByVal vfytm As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As Recordset
'    Dim cnt As Long
'
'    strSQL = "SELECT count(*) cnt " & _
'           "FROM " & T_LAB302 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " " & _
'           "AND (vfyid is not null or " & DBW("vfyid<>", "") & ") "
'
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
'        GoTo Exec_AccessStatus_error
'    End If
'
'    With Rs
'        If .RecordCount <= 0 Then
'            cnt = 0
'        Else
'            cnt = .Fields("cnt").Value & ""
'        End If
'    End With
'    Set Rs = Nothing
'
'    strSQL = "update " & T_LAB201 & " " & _
'           "set   " & DBW("reqinputcnt=", cnt) & " " & _
'           "WHERE " & DBW("workarea   =", workarea) & " " & _
'           "AND   " & DBW("accdt      =", accdt) & " " & _
'           "AND   " & DBW("accseq     =", accseq) & " "
'
'    DBConn.Execute strSQL
'
'    strSQL = "update " & T_LAB201 & " " & _
'           "set " & DBW("stscd=", "5") & "," & _
'                    DBW("vfyid=", VfyId) & "," & _
'                    DBW("vfydt=", vfydt) & "," & _
'                    DBW("vfytm=", vfytm) & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt=", accdt) & " " & _
'           "AND   " & DBW("accseq=", accseq) & " " & _
'           "AND   reqtotcnt=reqinputcnt "
'
'    DBConn.Execute strSQL
'
'    Exec_AccessStatus = True
'    Exit Function
'
'Exec_AccessStatus_error:
'    Exec_AccessStatus = False
'    MsgBox Err.Description, vbExclamation
'End Function

'Private Function GetCountResult(ByVal ptid As String, ByVal orddt As String, ByVal ordno As String, ByVal ordseq As String) As Long
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'On Error GoTo GetCountResult_error
'
'    strSQL = "SELECT count(*) cnt " & _
'           "FROM " & T_LAB302 & " " & _
'           "WHERE " & DBW("ptid  =", ptid) & " " & _
'           "AND   " & DBW("orddt =", orddt) & " " & _
'           "AND   " & DBW("ordno =", ordno) & " " & _
'           "AND   " & DBW("ordseq=", ordseq) & " " & _
'           "AND (vfydt is null or " & DBW("vfydt=", "") & ") "
'
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
''        dbconn.DisplayErrors
'        GetCountResult = -1
'        Set Rs = Nothing
'        Exit Function
'    End If
'
'    With Rs
'        If .RecordCount < 1 Then
'            MsgBox "상태를 변경하기 위해 결과내역을 읽는중 오류가 발생하였읍니다.(2)", vbCritical, "오류"
'            Set Rs = Nothing
'            GetCountResult = -1
'            Exit Function
'        Else
'            GetCountResult = .Fields("cnt").Value & ""
'        End If
'    End With
'    Set Rs = Nothing
'
'    Exit Function
'
'GetCountResult_error:
''    DBConn.DisplayErrors
'    GetCountResult = -1
'End Function

'Private Function Exec_OrderStatus(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                  ByVal vfydt As String, ByVal vfytm As String, ByVal VfyId As String) As Boolean
'    Dim ptid As String
'    Dim orddt As String
'    Dim ordno As String
'    Dim ordseq As String
'
'    Dim Rs As Recordset
'    Dim strSQL As String
'    Dim cnt As Long
'    Dim i As Long
'
'On Error GoTo Exec_OrderStatus_Error
'
'    strSQL = "SELECT distinct ptid,orddt,ordno,ordseq " & _
'           "FROM " & T_LAB302 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " "
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
'        GoTo Exec_OrderStatus_Error
'    End If
'
'    With Rs
'        For i = 1 To .RecordCount
'            ptid = .Fields("ptid").Value & ""
'            orddt = .Fields("orddt").Value & ""
'            ordno = .Fields("ordno").Value & ""
'            ordseq = .Fields("ordseq").Value & ""
'
'            cnt = GetCountResult(ptid, orddt, ordno, ordseq)
'            If cnt <= 0 Then
''                strSQL = "update " & T_LAB102 & " " & _
''                       "set " & DBW("stscd   =", "5") & ", " & _
''                                DBW("examdt  =", vfydt) & "," & _
''                                DBW("examtm  =", vfytm) & "," & _
''                                DBW("examdoct=", VfyId) & " " & _
''                       "WHERE " & DBW("ptid  =", ptid) & " " & _
''                       "AND   " & DBW("orddt =", orddt) & " " & _
''                       "AND   " & DBW("ordno =", ordno) & " " & _
''                       "AND   " & DBW("ordseq=", ordseq) & " "
'
'                '예수병원용
'                strSQL = "update mdexmort " & _
'                       "set " & DBW("stscd   =", "5") & ", " & _
'                                " execdate=to_date(" & vfydt & vfytm & ",'yyyymmdd hh24:mi:ss') ," & _
'                                DBW("cofmdr=", VfyId) & " " & _
'                       "WHERE " & DBW("patno  =", ptid) & " " & _
'                       "AND   " & " orddate=to_date(" & orddt & ",'yyyymmdd') " & _
'                       "AND   " & DBW("ordseqno =", ordno) & " "
'
'                DBConn.Execute strSQL
'            End If
'
'            .MoveNext
'        Next i
'    End With
'    Set Rs = Nothing
'
'
'    Exec_OrderStatus = True
'    Exit Function
'
'Exec_OrderStatus_Error:
'    Exec_OrderStatus = False
'    MsgBox Err.Description, vbExclamation
'End Function

'Private Function Exec_ModifyOrderStatus(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                        ByVal vfydt As String, ByVal vfytm As String, ByVal VfyId As String) As Boolean
'    Dim ptid As String
'    Dim orddt As String
'    Dim ordno As String
'    Dim ordseq As String
'
'    Dim Rs As Recordset
'    Dim strSQL As String
'    Dim cnt As Long
'    Dim i As Long
'
'On Error GoTo Exec_ModifyOrderStatus_Error
'
'    strSQL = "SELECT distinct ptid,orddt,ordno,ordseq " & _
'           "FROM " & T_LAB302 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " "
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
'        GoTo Exec_ModifyOrderStatus_Error
'    End If
'
''    MsgBox "처방내역 업데이트", vbExclamation
'
'    With Rs
'        For i = 1 To .RecordCount
'            ptid = .Fields("ptid").Value & ""
'            orddt = .Fields("orddt").Value & ""
'            ordno = .Fields("ordno").Value & ""
'            ordseq = .Fields("ordseq").Value & ""
'
'            cnt = GetCountResult(ptid, orddt, ordno, ordseq)
'            If cnt <= 0 Then
''                strSQL = "update " & T_LAB102 & " " & _
''                       "set " & DBW("stscd    =", "6") & ", " & _
''                                DBW("examdt   =", vfydt) & "," & _
''                                DBW("examtm   =", vfytm) & "," & _
''                                DBW("examdoct =", VfyId) & " " & _
''                       "WHERE " & DBW("ptid   =", ptid) & " " & _
''                       "AND   " & DBW("orddt  =", orddt) & " " & _
''                       "AND   " & DBW("ordno  =", ordno) & " " & _
''                       "AND   " & DBW("ordseq =", ordseq) & " " & _
''                       "AND   " & DBW("stscd >=", "5") & " "
'
'                '예수병원용
'                strSQL = "update mdexmort " & _
'                       "set " & DBW("stscd    =", "6") & ", " & _
'                                " execdate=to_date(" & vfydt & vfytm & ",'yyyymmdd hh24:mi:ss')," & _
'                                DBW("cofmdr =", VfyId) & " " & _
'                       "WHERE " & DBW("patno   =", ptid) & " " & _
'                       "AND   orddate=to_date(" & orddt & ",'yyyymmdd')" & _
'                       "AND   " & DBW("ordseqno  =", ordno) & " " & _
'                       "AND   " & DBW("stscd >=", "5") & " "
'                DBConn.Execute strSQL
'            End If
'
'            .MoveNext
'        Next i
'    End With
'    Set Rs = Nothing
'
'
'    Exec_ModifyOrderStatus = True
'    Exit Function
'
'Exec_ModifyOrderStatus_Error:
'    Exec_ModifyOrderStatus = False
'    MsgBox Err.Description, vbExclamation
'End Function

'Private Function CopyLAB302(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, ByVal TestCd As String, _
'                            ByVal pvfyid As String, ByVal pvfydt As String, ByVal pvfytm As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'    Dim cnt As Long
'    Dim maxseq As Long
'    Dim rstval As String
'    Dim rstcd  As String
'    Dim rstunit As String
'    Dim hldiv   As String
'    Dim dpdiv   As String
'    Dim mfydt   As String
'    Dim mfytm   As String
'    Dim mfyid   As String
'    Dim vfydt   As String
'    Dim vfytm   As String
'    Dim VfyId   As String
'    Dim rsttype As String
'
'
'
'
'On Error GoTo CopyLAB302_error
'
'    strSQL = "SELECT * " & _
'           "FROM " & T_LAB302 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " " & _
'           "AND   " & DBW("testcd  =", TestCd) & " "
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
''        MsgBox "결과내역을 읽는 중 오류가 발생하였읍니다.", vbCritical, "오류"
'        Set Rs = Nothing
'        CopyLAB302 = False
'        Exit Function
'    End If
'
'    With Rs
'        cnt = 0
'        If .RecordCount > 0 Then
'            cnt = 1
'            rstval = .Fields("rstval").Value & ""
'            rstcd = .Fields("rstcd").Value & ""
'            rstunit = .Fields("rstunit").Value & ""
'            hldiv = .Fields("hldiv").Value & ""
'            dpdiv = .Fields("dpdiv").Value & ""
'            mfydt = pvfydt
'            mfytm = pvfytm
'            mfyid = pvfyid
'            vfydt = .Fields("vfydt").Value & ""
'            vfytm = .Fields("vfytm").Value & ""
'            VfyId = .Fields("vfyid").Value & ""
'            rsttype = .Fields("rsttype").Value & ""
'        End If
'    End With
'    Set Rs = Nothing
'
'    If cnt <= 0 Then Exit Function
'
'
'    strSQL = "SELECT max(seq) maxseq " & _
'           "FROM " & T_LAB308 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " " & _
'           "AND   " & DBW("testcd  =", TestCd) & " "
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
''        MsgBox "결과내역을 읽는 중 오류가 발생하였읍니다.(2)", vbCritical, "오류"
'        Set Rs = Nothing
'        CopyLAB302 = False
'        Exit Function
'    End If
'
'    With Rs
'        If .RecordCount > 1 Then
'            maxseq = .Fields("maxseq").Value & ""
'        Else
'            maxseq = 0
'        End If
'    End With
'    Set Rs = Nothing
'
'
'    strSQL = "insert into " & T_LAB308 & "(workarea,accdt,accseq,testcd,seq" & _
'                                       ",rstval,rstcd,rstunit,hldiv,dpdiv" & _
'                                       ",mfyid,mfydt,mfytm,vfyid,vfydt,vfytm" & _
'                                       ",rsttype) " & _
'           "values (" & _
'                    DBV("workarea", workarea) & "," & _
'                    DBV("accdt", accdt) & "," & _
'                    DBV("accseq", accseq) & "," & _
'                    DBV("testcd", TestCd) & "," & _
'                    DBV("seq", maxseq + 1) & "," & _
'                    DBV("rstval", rstval) & "," & _
'                    DBV("rstcd", rstcd) & "," & _
'                    DBV("rstunit", rstunit) & "," & _
'                    DBV("hldiv", hldiv) & "," & _
'                    DBV("dpdiv", dpdiv) & "," & _
'                    DBV("mfyid", mfyid) & "," & _
'                    DBV("mfydt", mfydt) & "," & _
'                    DBV("mfytm", mfytm) & "," & _
'                    DBV("vfyid", VfyId) & "," & _
'                    DBV("vfydt", vfydt) & "," & _
'                    DBV("vfytm", vfytm) & "," & _
'                    DBV("rsttype", rsttype) & ")"
'
'    DBConn.Execute strSQL
'
'    CopyLAB302 = True
'    Exit Function
'
'CopyLAB302_error:
'    CopyLAB302 = False
'    MsgBox Err.Description, vbExclamation
'End Function

'Private Function Exec_ModifyDoubleCheck(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, astrSql() As String, idx As Long) As Boolean
'    Dim Rs As Recordset
'    Dim strSQL As String
'
'    Dim abo1 As String
'    Dim rh1 As String
'    Dim abo2 As String
'    Dim rh2 As String
'    Dim ABOSub As String
'    Dim RhSub As String
'
'    Dim vfyid1 As String
'    Dim vfydt1 As String
'    Dim vfytm1 As String
'
'    Dim vfyid2 As String
'    Dim vfydt2 As String
'    Dim vfytm2 As String
'
'    Dim VfyId As String
'    Dim vfydt As String
'    Dim vfytm As String
'
'On Error GoTo Exec_ModifyDoubleCheck_error
'
'    Set Rs = GetABOResultInfo(workarea, accdt, accseq)
'    If Rs Is Nothing Then
'        Exec_ModifyDoubleCheck = False
'    End If
'
'    With Rs
'        If .RecordCount < 1 Then
'            Exec_ModifyDoubleCheck = False
'        Else
'            abo1 = Trim(.Fields("abo1").Value & "")
'            rh1 = Trim(.Fields("rh1").Value & "")
'            abo2 = Trim(.Fields("abo2").Value & "")
'            rh2 = Trim(.Fields("rh2").Value & "")
'            ABOSub = Trim(.Fields("abosub").Value & "")
'            RhSub = Trim(.Fields("rhsub").Value & "")
'
'            vfyid1 = Trim(.Fields("vfyid1").Value & "")
'            vfydt1 = Trim(.Fields("vfydt1").Value & "")
'            vfytm1 = Trim(.Fields("vfytm1").Value & "")
'
'            vfyid2 = Trim(.Fields("vfyid2").Value & "")
'            vfydt2 = Trim(.Fields("vfydt2").Value & "")
'            vfytm2 = Trim(.Fields("vfytm2").Value & "")
'
'            If vfyid1 = "0" Then vfyid1 = ""
'            If vfyid2 = "0" Then vfyid2 = ""
'        End If
'    End With
'
'
'    If vfydt1 <> "" And vfydt2 <> "" Then
'        If vfydt1 > vfydt2 Then
'            VfyId = vfyid1
'            vfydt = vfydt1
'            vfytm = vfytm1
'        ElseIf vfydt2 > vfydt1 Then
'            VfyId = vfyid2
'            vfydt = vfydt2
'            vfytm = vfytm2
'        Else
'            If vfytm1 > vfytm2 Then
'                VfyId = vfyid1
'                vfydt = vfydt1
'                vfytm = vfytm1
'            Else
'                VfyId = vfyid2
'                vfydt = vfydt2
'                vfytm = vfytm2
'            End If
'        End If
'    Else
'        Exec_ModifyDoubleCheck = True
'        Exit Function
'    End If
'
'    strSQL = "update " & T_BBS303 & " " & _
'           "set " & DBW("vfydt=", vfydt) & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND   " & DBW("accdt   =", accdt) & " " & _
'           "AND   " & DBW("accseq  =", accseq) & " " & _
'           "AND seq=(SELECT max(seq) " & _
'                    "FROM " & T_BBS303 & " " & _
'                    "WHERE " & DBW("workarea=", workarea) & " " & _
'                    "AND   " & DBW("accdt   =", accdt) & " " & _
'                    "AND   " & DBW("accseq  =", accseq) & ")"
'    '
'    DBConn.Execute strSQL
'
'
'    'LAB302 의 Copy본을 만든다.
'    If CODE_ABOFRONT <> "" Then
'        If Not CopyLAB302(workarea, accdt, accseq, CODE_ABOFRONT, VfyId, vfydt, vfytm) Then
'            Exec_ModifyDoubleCheck = False
'            Exit Function
'        End If
'        If abo1 <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd  =", abo1) & "," & _
'                            DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND   " & DBW("accdt   =", accdt) & " " & _
'                   "AND   " & DBW("accseq  =", accseq) & " " & _
'                   "AND   " & DBW("testcd  =", CODE_ABOFRONT) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   DBW("vfyid=", VfyId) & "," & _
'                   DBW("vfydt=", vfydt) & "," & _
'                   DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND   " & DBW("accdt   =", accdt) & " " & _
'               "AND   " & DBW("accseq  =", accseq) & " " & _
'               "AND   " & DBW("testcd  =", CODE_ABOFRONT) & " "
'        DBConn.Execute strSQL
'    End If
'
'    If CODE_ABOBACK <> "" Then
'        If CODE_ABOBACK <> CODE_ABOFRONT Then
'            If Not CopyLAB302(workarea, accdt, accseq, CODE_ABOBACK, VfyId, vfydt, vfytm) Then
'                Exec_ModifyDoubleCheck = False
'                Exit Function
'            End If
'        End If
'        If abo2 <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd", abo2, 3) & _
'                            DBW("rsttype", "F", 2) & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_ABOBACK) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_ABOBACK) & " "
'        DBConn.Execute strSQL
'    End If
'
'    If CODE_RH <> "" Then
'        If CODE_RH <> CODE_ABOFRONT And CODE_RH <> CODE_ABOBACK Then
'            If Not CopyLAB302(workarea, accdt, accseq, CODE_RH, VfyId, vfydt, vfytm) Then
'                Exec_ModifyDoubleCheck = False
'                Exit Function
'            End If
'        End If
'        If rh1 <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", rh1) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_RH) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_RH) & " "
'        DBConn.Execute strSQL
'    End If
'
'    If CODE_ABOSUB <> "" Then
'        If CODE_ABOSUB <> CODE_ABOFRONT And CODE_ABOSUB <> CODE_ABOBACK And CODE_ABOSUB <> CODE_RH Then
'            If Not CopyLAB302(workarea, accdt, accseq, CODE_ABOSUB, VfyId, vfydt, vfytm) Then
'                Exec_ModifyDoubleCheck = False
'                Exit Function
'            End If
'        End If
'        If ABOSub <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", ABOSub) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_ABOSUB) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_ABOSUB) & " "
'        DBConn.Execute strSQL
'    End If
'
'    If CODE_RHSUB <> "" Then
'        If CODE_RHSUB <> CODE_ABOFRONT And CODE_RHSUB <> CODE_ABOBACK And CODE_RHSUB <> CODE_RH And CODE_RHSUB = CODE_ABOSUB Then
'            If Not CopyLAB302(workarea, accdt, accseq, CODE_RHSUB, VfyId, vfydt, vfytm) Then
'                Exec_ModifyDoubleCheck = False
'                Exit Function
'            End If
'        End If
'        If RhSub <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", RhSub) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_RHSUB) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_RHSUB) & " "
'        DBConn.Execute strSQL
'    End If
'
'    Exec_ModifyDoubleCheck = True
'
'    Exit Function
'
'Exec_ModifyDoubleCheck_error:
'    Exec_ModifyDoubleCheck = False
'    MsgBox Err.Description, vbExclamation
'End Function
'
'삭제대상
'Public Function GetDetailFg(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, ByVal TestCd As String) As String
'End Function

'Public Function Exec_DetailVerify(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, ByVal TestCd As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'    Dim detailfg As String
'    Dim VfyId As String
'    Dim vfydt As String
'    Dim vfytm As String
'
'On Error GoTo Exec_DetailVerify_error
'
'    strSQL = "SELECT vfyid,vfydt,vfytm,detailfg " & _
'           "FROM " & T_LAB302 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " " & _
'           "AND " & DBW("testcd=", TestCd) & " "
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
'        Exec_DetailVerify = False
'        Set Rs = Nothing
'        Exit Function
'    End If
'
'    With Rs
'        If .RecordCount > 0 Then
'            detailfg = .Fields("detailfg").Value & ""
'            VfyId = .Fields("vfyid").Value & ""
'            vfydt = .Fields("vfydt").Value & ""
'            vfytm = .Fields("vfytm").Value & ""
'        Else
'            detailfg = ""
'        End If
'    End With
'    Set Rs = Nothing
'
'    'Detail항목이 아니다.----------------------------------
'    If detailfg = "" Then
'        Exec_DetailVerify = True
'        Exit Function
'    End If
'
'    strSQL = "update " & T_LAB302 & " " & _
'           "set " & DBW("vfyid=", VfyId) & ", " & _
'               "" & DBW("vfydt=", vfydt) & "," & _
'               "" & DBW("vfytm=", vfytm) & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " " & _
'           "AND " & DBW("detailfg=", detailfg) & " " & _
'           "AND " & DBW("rstdiv=", "*") & " "
'    DBConn.Execute strSQL
'
'    Exec_DetailVerify = True
'    Exit Function
'
'Exec_DetailVerify_error:
'    Exec_DetailVerify = False
'    MsgBox Err.Description, vbExclamation
'End Function
'
'Private Function Exec_DoubleCheck(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, astrSql() As String, idx As Long) As Boolean
'    Dim Rs As Recordset
'    Dim strSQL As String
'
'    Dim abo1 As String
'    Dim rh1 As String
'    Dim abo2 As String
'    Dim rh2 As String
'    Dim ABOSub As String
'    Dim RhSub As String
'
'    Dim vfyid1 As String
'    Dim vfydt1 As String
'    Dim vfytm1 As String
'
'    Dim vfyid2 As String
'    Dim vfydt2 As String
'    Dim vfytm2 As String
'
'    Dim VfyId As String
'    Dim vfydt As String
'    Dim vfytm As String
'
'On Error GoTo Exec_DoubleCheck_error
'
'    Set Rs = GetABOResultInfo(workarea, accdt, accseq)
'    If Rs Is Nothing Then
'        Exec_DoubleCheck = False
'    End If
'
'    With Rs
'        If .RecordCount < 1 Then
'            Exec_DoubleCheck = False
'        Else
'            abo1 = Trim(.Fields("abo1").Value & "")
'            rh1 = Trim(.Fields("rh1").Value & "")
'            abo2 = Trim(.Fields("abo2").Value & "")
'            rh2 = Trim(.Fields("rh2").Value & "")
'            ABOSub = Trim(.Fields("abosub").Value & "")
'            RhSub = Trim(.Fields("rhsub").Value & "")
'
'            vfyid1 = Trim(.Fields("vfyid1").Value & "")
'            vfydt1 = Trim(.Fields("vfydt1").Value & "")
'            vfytm1 = Trim(.Fields("vfytm1").Value & "")
'
'            vfyid2 = Trim(.Fields("vfyid2").Value & "")
'            vfydt2 = Trim(.Fields("vfydt2").Value & "")
'            vfytm2 = Trim(.Fields("vfytm2").Value & "")
'
'            If vfyid1 = "0" Then vfyid1 = ""
'            If vfyid2 = "0" Then vfyid2 = ""
'        End If
'    End With
'
'
'    If vfydt1 <> "" And vfydt2 <> "" Then
'        If vfydt1 > vfydt2 Then
'            VfyId = vfyid1
'            vfydt = vfydt1
'            vfytm = vfytm1
'        ElseIf vfydt2 > vfydt1 Then
'            VfyId = vfyid2
'            vfydt = vfydt2
'            vfytm = vfytm2
'        Else
'            If vfytm1 > vfytm2 Then
'                VfyId = vfyid1
'                vfydt = vfydt1
'                vfytm = vfytm1
'            Else
'                VfyId = vfyid2
'                vfydt = vfydt2
'                vfytm = vfytm2
'            End If
'        End If
'    Else
'        Exec_DoubleCheck = True
'        Exit Function
'    End If
'
'    strSQL = "update " & T_BBS303 & " " & _
'           "set " & DBW("vfydt=", vfydt) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND seq=(SELECT max(seq) " & _
'                        "FROM " & T_BBS303 & " " & _
'                        "WHERE " & DBW("workarea=", workarea) & " " & _
'                        "AND " & DBW("accdt=", accdt) & " " & _
'                        "AND " & DBW("accseq=", accseq) & ")"
'    DBConn.Execute strSQL
'
'
'    If CODE_ABOFRONT <> "" Then
'        If abo1 <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", abo1) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_ABOFRONT) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_ABOFRONT) & " "
'        DBConn.Execute strSQL
'        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
'        If Not Exec_DetailVerify(workarea, accdt, accseq, CODE_ABOFRONT) Then
'            GoTo Exec_DoubleCheck_error
'        End If
'    End If
'
'    If CODE_ABOBACK <> "" Then
'        If abo2 <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", abo2) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_ABOBACK) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_ABOBACK) & " "
'        DBConn.Execute strSQL
'        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
'        If Not Exec_DetailVerify(workarea, accdt, accseq, CODE_ABOBACK) Then
'            GoTo Exec_DoubleCheck_error
'        End If
'    End If
'
'    If CODE_RH <> "" Then
'        If rh1 <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", rh1) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_RH) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_RH) & " "
'        DBConn.Execute strSQL
'        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
'        If Not Exec_DetailVerify(workarea, accdt, accseq, CODE_RH) Then
'            GoTo Exec_DoubleCheck_error
'        End If
'    End If
'
'
'    If CODE_ABOSUB <> "" Then
'        If ABOSub <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", ABOSub) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_ABOSUB) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_ABOSUB) & " "
'        DBConn.Execute strSQL
'        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
'        If Not Exec_DetailVerify(workarea, accdt, accseq, CODE_ABOSUB) Then
'            GoTo Exec_DoubleCheck_error
'        End If
'    End If
'
'    If CODE_RHSUB <> "" Then
'        If RhSub <> "" Then
'            strSQL = "update " & T_LAB302 & " " & _
'                   "set " & DBW("rstcd=", RhSub) & "," & DBW("rsttype=", "F") & " " & _
'                   "WHERE " & DBW("workarea=", workarea) & " " & _
'                   "AND " & DBW("accdt=", accdt) & " " & _
'                   "AND " & DBW("accseq=", accseq) & " " & _
'                   "AND " & DBW("testcd=", CODE_RHSUB) & " "
'            DBConn.Execute strSQL
'        End If
'        strSQL = "update " & T_LAB302 & " " & _
'               "set " & _
'                   "" & DBW("vfyid=", VfyId) & "," & _
'                   "" & DBW("vfydt=", vfydt) & "," & _
'                   "" & DBW("vfytm=", vfytm) & " " & _
'               "WHERE " & DBW("workarea=", workarea) & " " & _
'               "AND " & DBW("accdt=", accdt) & " " & _
'               "AND " & DBW("accseq=", accseq) & " " & _
'               "AND " & DBW("testcd=", CODE_RHSUB) & " "
'        DBConn.Execute strSQL
'        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
'        If Not Exec_DetailVerify(workarea, accdt, accseq, CODE_RHSUB) Then
'            GoTo Exec_DoubleCheck_error
'        End If
'    End If
'
'    Exec_DoubleCheck = True
'    Exit Function
'
'Exec_DoubleCheck_error:
'    Exec_DoubleCheck = False
'    MsgBox Err.Description, vbExclamation
'End Function

'Public Function GetAccessComment(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String) As Recordset
'    Dim strSQL As String
'
'On Error GoTo GetAccessComment
'
'    strSQL = "SELECT * " & _
'           "FROM " & T_LAB304 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " " & _
'           "AND seq=(SELECT max(seq) " & _
'                    "FROM " & T_LAB304 & " " & _
'                    "WHERE " & DBW("workarea=", workarea) & " " & _
'                    "AND " & DBW("accdt=", accdt) & " " & _
'                    "AND " & DBW("accseq=", accseq) & ") "
'
'    Set GetAccessComment = New Recordset
'    Call GetAccessComment.Open(strSQL, DBConn)
''    If GetAccessComment.DBerror Then
''        dbconn.DisplayErrors
''        Set GetAccessComment = Nothing
''    End If
'    Exit Function
'
'GetAccessComment:
''    DBConn.DisplayErrors
'    Set GetAccessComment = Nothing
'End Function

'Private Function Exec_InsertComment(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                    ByVal VfyId As String, ByVal comment As String) As Boolean
'    '
'    Dim strSQL As String
'    Dim Rs As Recordset
'    Dim maxseq As Long
'    Dim rsttext As String
'
'On Error GoTo Exec_InsertComment_error
'    Exec_InsertComment = True
'
'    Set Rs = GetAccessComment(workarea, accdt, accseq)
'    If Rs Is Nothing Then
'        Exec_InsertComment = False
'        Exit Function
'    End If
'
'    With Rs
'        If .RecordCount <= 0 Then
'            maxseq = 0
'            rsttext = ""
'        Else
'            maxseq = .Fields("seq").Value & ""
'            rsttext = .Fields("rsttxt").Value & ""
'        End If
'    End With
'    Set Rs = Nothing
'
'    If Exec_InsertComment = False Then Exit Function
'
'    If Trim(rsttext) = Trim(comment) Then
'        Exec_InsertComment = True
'        Exit Function
'    End If
'
'    strSQL = "insert into " & T_LAB304 & "(workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
'           "values(" & _
'                    DBV("workarea", workarea) & "," & _
'                    DBV("accdt", accdt) & "," & _
'                    DBV("accseq", accseq) & "," & _
'                    DBV("seq", maxseq + 1) & "," & _
'                    DBV("vfyid", VfyId) & "," & _
'                    DBV("rsttxt", comment) & "" & _
'           ")"
'
'    DBConn.Execute strSQL
'    Exec_InsertComment = True
'    Exit Function
'
'Exec_InsertComment_error:
'    Exec_InsertComment = False
'    MsgBox Err.Description, vbExclamation
'End Function
'
'Private Function Exec_UpdateRemark(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                     ByVal remark As String) As Boolean
'    Dim strSQL As String
'
'On Error GoTo Exec_UpdateRemark_error
'
'    strSQL = "update " & T_LAB201 & " " & _
'           "set " & DBW("rmkcd=", remark) & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " "
'
'
'    DBConn.Execute strSQL
'    Exec_UpdateRemark = True
'    Exit Function
'
'Exec_UpdateRemark_error:
'    Exec_UpdateRemark = False
'    MsgBox Err.Description, vbExclamation
'End Function

'Private Function Exec_InsertABO(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                ByVal isVerify As Boolean, _
'                                ByVal typing As Long, _
'                                ByVal pABO As String, ByVal pRH As String, ByVal pABOSUB As String, ByVal pRHSUB As String, _
'                                ByVal pvfyid As String, ByVal pvfydt As String, ByVal pvfytm As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'    Dim seq As Long
'    Dim SpcCd As String
'    Dim statfg As String
'    Dim abo1 As String
'    Dim rh1 As String
'    Dim abo2 As String
'    Dim rh2 As String
'    Dim ABOSub As String
'    Dim RhSub As String
'    Dim vfyid1 As String
'    Dim vfydt1 As String
'    Dim vfytm1 As String
'    Dim vfyid2 As String
'    Dim vfydt2 As String
'    Dim vfytm2 As String
'    Dim vfydt As String
'    Dim ptid As String
'
'On Error GoTo Exec_InsertABO_error
'
'    strSQL = "SELECT * " & _
'           "FROM " & T_BBS303 & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " " & _
'           "AND seq=(SELECT max(seq) " & _
'                    "FROM " & T_BBS303 & " " & _
'                    "WHERE " & DBW("workarea=", workarea) & " " & _
'                    "AND " & DBW("accdt=", accdt) & " " & _
'                    "AND " & DBW("accseq=", accseq) & ")"
'
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
'        GoTo Exec_InsertABO_error
'    End If
'
'    With Rs
'        If .RecordCount > 0 Then
'            seq = .Fields("seq").Value & ""
'            SpcCd = .Fields("spccd").Value & ""
'            statfg = .Fields("statfg").Value & ""
'            abo1 = .Fields("abo1").Value & ""
'            rh1 = .Fields("rh1").Value & ""
'            abo2 = .Fields("abo2").Value & ""
'            rh2 = .Fields("rh2").Value & ""
'            ABOSub = .Fields("abosub").Value & ""
'            RhSub = .Fields("rhsub").Value & ""
'            vfyid1 = .Fields("vfyid1").Value & ""
'            vfydt1 = .Fields("vfydt1").Value & ""
'            vfytm1 = .Fields("vfytm1").Value & ""
'            vfyid2 = .Fields("vfyid2").Value & ""
'            vfydt2 = .Fields("vfydt2").Value & ""
'            vfytm2 = .Fields("vfytm2").Value & ""
'            vfydt = .Fields("vfydt").Value & ""
'            ptid = .Fields("ptid").Value & ""
'            Exec_InsertABO = True
'        Else
'            MsgBox "결과내역을 읽을 수 없습니다.전산실로 연락바랍니다.", vbCritical, "오류"
'            Exec_InsertABO = False
'        End If
'    End With
'    Set Rs = Nothing
'
'    If Exec_InsertABO = False Then Exit Function
'
'    '기존내역을 복사한다.(seq를 1증가시켜서)
'    strSQL = "insert into " & T_BBS303 & "(workarea,accdt,accseq,seq,spccd,statfg," & _
'                                        "abo1,rh1,abo2,rh2,abosub,rhsub," & _
'                                        "vfyid1,vfydt1,vfytm1,vfyid2,vfydt2,vfytm2," & _
'                                        "vfydt,ptid) " & _
'           "values(" & _
'                    DBV("workarea", workarea) & "," & _
'                    DBV("accdt", accdt) & "," & _
'                    DBV("accseq", accseq) & "," & _
'                    DBV("seq", seq + 1) & "," & _
'                    DBV("spccd", SpcCd) & "," & _
'                    DBV("statfg", statfg) & "," & _
'                    DBV("abo1", abo1) & "," & _
'                    DBV("rh1", rh1) & "," & _
'                    DBV("abo2", abo2) & "," & _
'                    DBV("rh2", rh2) & "," & _
'                    DBV("abosub", ABOSub) & "," & _
'                    DBV("rhsub", RhSub) & "," & _
'                    DBV("vfyid1", vfyid1) & "," & _
'                    DBV("vfydt1", vfydt1) & "," & _
'                    DBV("vfytm1", vfytm1) & "," & _
'                    DBV("vfyid2", vfyid2) & "," & _
'                    DBV("vfydt2", vfydt2) & "," & _
'                    DBV("vfytm2", vfytm2) & "," & _
'                    DBV("vfydt", pvfydt) & "," & _
'                    DBV("ptid", ptid) & ")"
'    DBConn.Execute strSQL
'
'    Exec_InsertABO = Exec_UpdateABO(workarea, accdt, accseq, True, typing, pABO, pRH, pABOSUB, pRHSUB, pvfyid, pvfydt, pvfytm)
'
'    'Double Check완료되었으면, 결과 2개 모두 수정
'    '                          LAB302의 복사본 마련, LAB308
'    If typing = 0 Then
'        If vfyid2 = "0" Then vfyid2 = ""
'        If vfyid2 <> "" Then
'            Exec_InsertABO = Exec_UpdateABO(workarea, accdt, accseq, True, "1", pABO, pRH, pABOSUB, pRHSUB, pvfyid, pvfydt, pvfytm)
'        End If
'    Else
'        If vfyid1 = "0" Then vfyid1 = ""
'        If vfyid1 <> "" Then
'            Exec_InsertABO = Exec_UpdateABO(workarea, accdt, accseq, True, "0", pABO, pRH, pABOSUB, pRHSUB, pvfyid, pvfydt, pvfytm)
'        End If
'    End If
'
'
'    Exit Function
'
'Exec_InsertABO_error:
'    Exec_InsertABO = False
'    MsgBox Err.Description, vbExclamation
'End Function
'
'Private Function Exec_UpdateABO(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                ByVal isVerify As Boolean, _
'                                ByVal typing As Long, _
'                                ByVal pABO As String, ByVal pRH As String, ByVal pABOSUB As String, ByVal pRHSUB As String, _
'                                ByVal EmpID As String, ByVal vfydt As String, ByVal vfytm As String) As Boolean
'
'    Dim strSQL As String
'
'On Error GoTo Exec_UpdateABO_error
'
'    If typing = 0 Then
'        strSQL = "update " & T_BBS303 & " " & _
'               "set " & DBW("abo1=", pABO) & "," & _
'                   "" & DBW("rh1=", pRH) & "," & _
'                   "" & DBW("abosub=", pABOSUB) & "," & _
'                   "" & DBW("rhsub=", pRHSUB) & ", " & _
'                   "" & DBW("vfyid1=", EmpID) & " "
'        If isVerify Then
'            strSQL = strSQL & "," & _
'                   "" & DBW("vfydt1=", vfydt) & "," & _
'                   "" & DBW("vfytm1=", vfytm) & " "
'        End If
'    Else
'        strSQL = "update " & T_BBS303 & " " & _
'               "set " & DBW("abo2=", pABO) & "," & _
'                   "" & DBW("rh2=", pRH) & "," & _
'                   "" & DBW("abosub=", pABOSUB) & "," & _
'                   "" & DBW("rhsub=", pRHSUB) & ", " & _
'                   "" & DBW("vfyid2=", EmpID) & " "
'        If isVerify Then
'            strSQL = strSQL & "," & _
'                   "" & DBW("vfydt2=", vfydt) & "," & _
'                   "" & DBW("vfytm2=", vfytm) & " "
'        End If
'    End If
'
'    strSQL = strSQL & " " & _
'           "WHERE " & DBW("workarea=", workarea) & " " & _
'           "AND " & DBW("accdt=", accdt) & " " & _
'           "AND " & DBW("accseq=", accseq) & " " & _
'           "AND seq=(SELECT max(seq) " & _
'                    "FROM " & T_BBS303 & " " & _
'                    "WHERE " & DBW("workarea=", workarea) & " " & _
'                    "AND " & DBW("accdt=", accdt) & " " & _
'                    "AND " & DBW("accseq=", accseq) & ")"
'
'
'    DBConn.Execute strSQL
'    Exec_UpdateABO = True
'    Exit Function
'
'Exec_UpdateABO_error:
'    Exec_UpdateABO = False
'    MsgBox Err.Description, vbExclamation
'End Function
'
'Private Function CheckABOResult(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, _
'                                ByVal isVerify As Boolean, _
'                                ByVal typing As Long, _
'                                ByVal pABO As String, ByVal pRH As String, _
'                                ByVal pABOSUB As String, ByVal pRHSUB As String, _
'                                ByVal EmpID As String) As Boolean
'    Dim Rs As Recordset
'
'    Dim abo1 As String
'    Dim rh1 As String
'    Dim abo2 As String
'    Dim rh2 As String
'
'    Dim vfyid1 As String
'    Dim vfydt1 As String
'    Dim vfytm1 As String
'
'    Dim vfyid2 As String
'    Dim vfydt2 As String
'    Dim vfytm2 As String
'
'
'    If Not isVerify Then
'        'Verify가 아니면 저장할 수 있다.
'        CheckABOResult = True
'        Exit Function
'    End If
'
'    Set Rs = GetABOResultInfo(workarea, accdt, accseq)
'    If Rs Is Nothing Then
'        CheckABOResult = False
'        Exit Function
'    End If
'
'    With Rs
'        abo1 = .Fields("abo1").Value & ""
'        rh1 = .Fields("rh1").Value & ""
'        abo2 = .Fields("abo2").Value & ""
'        rh2 = .Fields("rh2").Value & ""
'
'        vfyid1 = .Fields("vfyid1").Value & ""
'        vfydt1 = .Fields("vfydt1").Value & ""
'        vfytm1 = .Fields("vfytm1").Value & ""
'
'        vfyid2 = .Fields("vfyid2").Value & ""
'        vfydt2 = .Fields("vfydt2").Value & ""
'        vfytm2 = .Fields("vfytm2").Value & ""
'
'        If vfyid1 = "0" Then vfyid1 = ""
'        If vfyid2 = "0" Then vfyid2 = ""
'
'    End With
'    Set Rs = Nothing
'
''이넘 탈 확률 0%
'    If vfydt1 <> "" And vfydt2 <> "" Then
'        MsgBox "이미 Double Check 완료되었읍니다.", vbCritical
'        CheckABOResult = False
'        Exit Function
'    End If
'
'
'    Select Case typing
'        Case 0:     '--------------------------------------------------Front
'            '이넘 탈 확률 없음
'            If vfyid1 <> "" Then
'                If vfyid1 <> EmpID Then
'                    MsgBox "이미 다른 사용자가 Front Typing 결과를 입력하였읍니다.", vbCritical, "오류"
'                    CheckABOResult = False
'                    Exit Function
'                End If
'            End If
'            If vfyid2 = "" Then
'                'Back Typing이 입력이 안되어있다.
'                CheckABOResult = True
'                Exit Function
'            End If
'            If vfydt2 = "" Then
'                'Back Typing이 Verify되지 않았다.
'                CheckABOResult = True
'                Exit Function
'            End If
'            If abo2 <> pABO Or rh2 <> pRH Then
'                MsgBox "앞서 입력된 결과와 결과값이 다릅니다.", vbCritical, "오류"
'                CheckABOResult = False
'                Exit Function
'            End If
'            CheckABOResult = True
'        Case 1:     '--------------------------------------------------Back
'            '이넘 탈 확률 없음
'            If vfyid2 <> "" Then
'                If vfyid2 <> EmpID Then
'                    MsgBox "이미 다른 사용자가 Back Typing 결과를 입력하였읍니다.", vbCritical, "오류"
'                    CheckABOResult = False
'                    Exit Function
'                End If
'            End If
'            If vfyid1 = "" Then
'                'Front Typing이 입력이 안되어있다.
'                CheckABOResult = True
'                Exit Function
'            End If
'            If vfydt1 = "" Then
'                'Front Typing이 Verify되지 않았다.
'                CheckABOResult = True
'                Exit Function
'            End If
'            If abo1 <> pABO Or rh1 <> pRH Then
'                MsgBox "앞서 입력된 결과와 결과값이 다릅니다.", vbCritical, "오류"
'                CheckABOResult = False
'                Exit Function
'            End If
'            CheckABOResult = True
'        Case Else:  '--------------------------------------------------???
'            CheckABOResult = False
'    End Select
'End Function
'
'Public Function GetABOList() As Recordset
'    Dim objSQL As clsBBSMSTStatement
'
'    Set objSQL = New clsBBSMSTStatement
''    objSql.setDbConn DBConn
'    Set GetABOList = objSQL.GetABOList
'    Set objSQL = Nothing
'End Function
'
'Public Function GetMinMaxSeq(ByVal Workdt As String, ByVal workCd As String, min As Long, max As Long) As Boolean
'    Dim strSQL As String
'    Dim Rs As Recordset
'
'    strSQL = "SELECT min(workseq) minseq,max(workseq) maxseq " & _
'           "FROM " & T_LAB301 & " " & _
'           "WHERE " & DBW("workdt=", Workdt) & " " & _
'           "AND " & DBW("workcd=", workCd) & " "
'
'    min = 0
'    max = 0
'
'    Set Rs = New Recordset
'    Call Rs.Open(strSQL, DBConn)
'    If Rs.EOF Then
''        dbconn.DisplayErrors
'        GetMinMaxSeq = False
'        Exit Function
'    End If
'
'    With Rs
'        If .RecordCount <= 0 Then
'            GetMinMaxSeq = False
'        Else
'            min = Val(.Fields("minseq").Value & "")
'            max = Val(.Fields("maxseq").Value & "")
'            GetMinMaxSeq = True
'        End If
'    End With
'    Set Rs = Nothing
'
'End Function

'Public Function GetAccListinfo(ByVal workarea As String, ByVal accdt As String, ByVal txtS As String, ByVal txtE As String) As Recordset
'    Dim strSQL As String
'    strSQL = " SELECT * FROM " & T_BBS303 & _
'         " WHERE " & DBW("workarea=", workarea) & _
'         " AND " & DBW("accdt=", accdt) & _
'         " AND " & DBW("accseq>=", txtS) & _
'         " AND " & DBW("accseq<=", txtE)
'    Set GetAccListinfo = New Recordset
'    Call GetAccListinfo.Open(strSQL, DBConn)
'End Function

'Public Function GetWorkSheetInfo(ByVal Workdt As String, ByVal workCd As String, ByVal FrNo As Long, ByVal ToNo As Long) As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT * " & _
'           "FROM " & T_LAB301 & " " & _
'           "WHERE " & DBW("workdt=", Workdt) & " " & _
'           "AND " & DBW("workcd=", workCd) & " " & _
'           "AND " & DBW("workseq>=", FrNo) & " " & _
'           "AND " & DBW("workseq<=", ToNo) & " "
'    Set GetWorkSheetInfo = New Recordset
'    Call GetWorkSheetInfo.Open(strSQL, DBConn)
'End Function

'Public Function GetWorkSheetList(ByVal center As String) As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT cdval1,field1 " & _
'           "FROM " & T_LAB032 & " " & _
'           "WHERE " & DBW("cdindex=", "C228") & " " & _
'           "AND " & DBW("field2=", center) & " "
'    Set GetWorkSheetList = New Recordset
'    Call GetWorkSheetList.Open(strSQL, DBConn)
'End Function


'Private Function GetABOTEstListInAccess(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String, aryAbotest() As String) As Boolean
'    '----------------------------------------
'    '혹시 몰라서 만든 펑션
'    '해당 접수번호에 묶인 ABO검사 코드를 반환
'    '----------------------------------------
'End Function


'Private Sub Class_Initialize()
'    Dim objcom003 As clsCom003
'    Dim Rs As Recordset
'
'
'    Set objcom003 = New clsCom003
'    Set Rs = objcom003.OpenRecordSetDay(BC2_ABO_TEST)
'    Set objcom003 = Nothing
'
'    If Not (Rs Is Nothing) Then
'        With Rs
'            If .RecordCount > 0 Then
'                CODE_ABOFRONT = .Fields("field1").Value & ""
'                CODE_ABOBACK = .Fields("field2").Value & ""
'                CODE_RH = .Fields("field3").Value & ""
'                CODE_ABOSUB = .Fields("field4").Value & ""
'                CODE_RHSUB = .Fields("text1").Value & ""
'            Else
'                CODE_ABOFRONT = ""
'                CODE_ABOBACK = ""
'                CODE_RH = ""
'                CODE_ABOSUB = ""
'                CODE_RHSUB = ""
'            End If
'        End With
'        Set Rs = Nothing
'    End If
'End Sub

'Public Function GetLabNoBySpcNo(ByVal pSpcYy As String, ByVal pSpcNo As String) As String

'    Dim Rs As Recordset
'    Dim SqlStmt As String
'    Dim strMsk As String
'
'    SqlStmt = " SELECT workarea, " & FUNC_SUBSTR & "(accdt,3,6) as qaccdt, accseq " & _
'               " FROM " & T_LAB201 & _
'               " WHERE  " & DBW("spcyy =", pSpcYy) & _
'               " AND    " & DBW("spcno =", pSpcNo)
'
'    Set Rs = New Recordset
'    Rs.Open SqlStmt, DBConn
'
'    GetLabNoBySpcNo = ""
'    If Not Rs.EOF Then
'        GetLabNoBySpcNo = "" & Rs.Fields("workarea").Value & "-" & _
'                          "" & Rs.Fields("qaccdt").Value & "-" & _
'                          "" & Rs.Fields("accseq").Value
'    End If
'
'    Set Rs = Nothing

'End Function

