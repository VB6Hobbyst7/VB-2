VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsABOSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetAccComment(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT * " & _
           "FROM " & T_LAB304 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " " & _
           "AND seq=(SELECT max(seq) " & _
                    "FROM " & T_LAB304 & " " & _
                    "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                    "AND " & DBW("accdt=", vAccdt) & " " & _
                    "AND " & DBW("accseq=", vAccseq) & ") "
                    
    Set GetAccComment = New Recordset
    Call GetAccComment.Open(strSQL, DBConn)
End Function

Public Function IsExistABO(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    strSQL = "SELECT count(*) cnt /*a.testcd,b.testdiv*/ " & _
           "FROM " & T_LAB302 & " a," & T_LAB001 & " b " & _
           "WHERE " & DBW("a.workarea=", vWorkarea) & " " & _
           "AND   " & DBW("a.accdt   =", vAccdt) & " " & _
           "AND   " & DBW("a.accseq  =", vAccseq) & " " & _
           "AND    a.testcd=b.testcd " & _
           "AND    b.applydt=(SELECT max(c.applydt) FROM " & T_LAB001 & " c WHERE b.testcd=c.testcd) " & _
           "AND   " & DBW("b.testdiv=", "3") & " "
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    
    If Rs.EOF Then
        IsExistABO = False
    Else
        IsExistABO = (Rs.Fields("cnt").Value & "") > 0
    End If
    
    Set Rs = Nothing
End Function

Public Function GetABOResultInfo(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT * " & _
           "FROM " & T_BBS303 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " " & _
           "AND seq=(SELECT max(seq) " & _
                    "FROM " & T_BBS303 & " " & _
                    "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                    "AND   " & DBW("accdt   =", vAccdt) & " " & _
                    "AND   " & DBW("accseq  =", vAccseq) & ")"

    Set GetABOResultInfo = New Recordset
    Call GetABOResultInfo.Open(strSQL, DBConn)
End Function

Public Function LoadRemark() As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT cdval1, field1,text1 " & _
           "FROM " & T_LAB034 & " " & _
           "WHERE " & DBW("cdindex = ", "C403") & " "
           
    Set LoadRemark = New Recordset
    Call LoadRemark.Open(strSQL, DBConn)
End Function

Public Function GetRemarkNm(ByVal vRmkCd As String) As String
    Dim strSQL As String
    Dim Rs As Recordset
    
    strSQL = "SELECT cdval1, field1 " & _
           "FROM " & T_LAB034 & " " & _
           "WHERE " & DBW("cdindex = ", "C403") & " " & _
           "AND   " & DBW("cdval1  = ", vRmkCd) & " "
           
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    
    If Rs.EOF Then
        GetRemarkNm = ""
    Else
        GetRemarkNm = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function LoadComment() As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT cdval1, text1 " & _
           "FROM " & T_LAB034 & " " & _
           "WHERE " & DBW("cdindex = ", "C404") & " "
           
    Set LoadComment = New Recordset
    Call LoadComment.Open(strSQL, DBConn)
End Function

Public Function GetAccessInfo(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT workarea,accdt,accseq,ptid,ageday,sex,bedindt,deptcd,majdoct,stscd," & _
                  "coldt,coltm,colid,rcvdt,rcvtm,rcvid," & _
                  "reqtotcnt,reqinputcnt,wardid,statfg,rmkcd " & _
           "FROM " & T_LAB201 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " "
    Set GetAccessInfo = New Recordset
    Call GetAccessInfo.Open(strSQL, DBConn)
End Function

Public Function GetOrderInfo(ByVal vPtid As String, ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, ByVal vOrdcd As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select examdt,examtm,examdoct from " & T_LAB101 & " a, " & T_LAB102 & " b " & _
           " where a.ptid=b.ptid " & _
           " and a.orddt=b.orddt " & _
           " and a.ordno=b.ordno " & _
           " and " & DBW("a.ptid=", vPtid) & _
           " and " & DBW("b.workarea=", vWorkarea) & _
           " and " & DBW("b.accdt=", vAccdt) & _
           " and " & DBW("b.accseq=", vAccseq) & _
           " and " & DBW("b.ordcd=", vOrdcd)
    Set GetOrderInfo = New Recordset
    GetOrderInfo.Open strSQL, DBConn
End Function

Public Function GetWSNm(ByVal vWsCd As String, ByVal vCenter As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT cdval1,field1 " & _
           " FROM " & T_LAB032 & " " & _
           " WHERE " & DBW("cdindex=", "C228") & _
           " and " & DBW("cdval1=", vWsCd) & _
           " AND " & DBW("field2=", vCenter)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetWSNm = ""
    Else
        GetWSNm = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function LoadWorksheet(ByVal vCenter As String) As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT cdval1,field1 " & _
           "FROM " & T_LAB032 & " " & _
           "WHERE " & DBW("cdindex=", "C228") & " " & _
           "AND " & DBW("field2=", vCenter) & " "
    Set LoadWorksheet = New Recordset
    Call LoadWorksheet.Open(strSQL, DBConn)
End Function

Public Function LoadWsInfo(ByVal vWorkdt As String, ByVal vWorkCd As String, ByVal vFrNo As Long, ByVal vToNo As Long) As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_LAB201 & " a, " & T_LAB301 & " b "
    strSQL = strSQL & " where a.workarea=b.workarea"
    strSQL = strSQL & " and a.accdt=b.accdt"
    strSQL = strSQL & " and a.accseq=b.accseq"
    strSQL = strSQL & " and " & DBW("b.workdt=", vWorkdt)
    strSQL = strSQL & " and " & DBW("b.workcd=", vWorkCd)
    strSQL = strSQL & " and " & DBW("b.workseq>=", vFrNo)
    strSQL = strSQL & " and " & DBW("b.workseq<=", vToNo)
    strSQL = strSQL & " and a.stscd in ('2','3','4') "
    
'    strSQL = "SELECT * " & _
'           "FROM " & T_LAB301 & " " & _
'           "WHERE " & DBW("workdt=", vWorkdt) & " " & _
'           "AND " & DBW("workcd=", vWorkCd) & " " & _
'           "AND " & DBW("workseq>=", vFrNo) & " " & _
'           "AND " & DBW("workseq<=", vToNo) & " "
    Set LoadWsInfo = New Recordset
    Call LoadWsInfo.Open(strSQL, DBConn)
End Function

Public Function LoadAccList(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vFrNo As String, ByVal vToNo As String) As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_LAB201 & " a, " & T_BBS303 & " b "
    strSQL = strSQL & " where a.workarea=b.workarea"
    strSQL = strSQL & " and a.accdt=b.accdt"
    strSQL = strSQL & " and a.accseq=b.accseq"
    strSQL = strSQL & " and " & DBW("a.workarea=", vWorkarea)
    strSQL = strSQL & " and " & DBW("a.accdt=", vAccdt)
    strSQL = strSQL & " and " & DBW("a.accseq>=", vFrNo) 'a.accseq between 1 and 100"
    strSQL = strSQL & " and " & DBW("a.accseq<=", vToNo) 'a.accseq between 1 and 100"
    strSQL = strSQL & " and a.stscd in ('2','3','4')"
    
'    strSQL = " SELECT * FROM " & T_BBS303 & _
'         " WHERE " & DBW("workarea=", vWorkarea) & _
'         " AND " & DBW("accdt=", vAccdt) & _
'         " AND " & DBW("accseq>=", vFrNo) & _
'         " AND " & DBW("accseq<=", vToNo)
    Set LoadAccList = New Recordset
    Call LoadAccList.Open(strSQL, DBConn)
End Function

'결과입력 및 결과등록, 결과수정 관련---------------------------------------------------------

Public Function SaveABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                              ByVal vVerify As Boolean, _
                              ByVal vTyping As Long, _
                              ByVal vABO As String, ByVal vRh As String, ByVal vABOSub As String, ByVal vRhSub As String, _
                              ByVal vRemark As String, ByVal vComment As String) As Boolean
    '-----------------------------------------------------------------------------------------
    'PARAMETER 의 용도
    '-----------------------------------------------------------------------------------------
    'vtyping   = 0     : Front Typing
    '         = 1     : Back typing
    '
    'vVerify = false : 저장만한다. (id만 입력)
    '         = true  : verify한다.
    '-----------------------------------------------------------------------------------------
    '저장시 주의사항
    '-----------------------------------------------------------------------------------------
    'LAB302의 RSTTYPE을 'F'로 만든다
    '-----------------------------------------------------------------------------------------
    
    Dim strVfydt As String
    Dim strVfytm As String
    
    strVfydt = Format(GetSystemDate, "YYYYMMDD")
    strVfytm = Format(GetSystemDate, "HHMMSS")
    
    '-----------------------------------------------------------------------------------------
    '진행순서
    '-----------------------------------------------------------------------------------------
    'STEP 1) CHECK
    'SETP 3) ABO 저장
    'STEP 2) REMARK, COMMENT 저장
    'STEP 4) LAB302저장(필요하면)
    '-----------------------------------------------------------------------------------------

'이넘은 체크할 필요 없음
'    If Not CheckABOResult(vWorkarea, vAccdt, vAccseq, isVerify, typing, pABO, pRH, pABOSUB, pRHSUB, EmpID) Then
'        SaveABOResult = False
'        Exit Function
'    End If
    
On Error GoTo SaveABOResult_Error

    DBConn.BeginTrans
    
    'BBS303저장
    If Not UpdateABO(vWorkarea, vAccdt, vAccseq, vVerify, vTyping, vABO, vRh, vABOSub, vRhSub, ObjMyUser.EmpID, strVfydt, strVfytm) Then
        GoTo SaveABOResult_Error
    End If
    
    'BBS201 검체리마크 저장
    If vRemark <> "" Then
        If Not UpdateRemark(vWorkarea, vAccdt, vAccseq, vRemark) Then
            GoTo SaveABOResult_Error
        End If
    End If
    
    'LAB304 코멘트 저장
    If vComment <> "" Then
        If Not InsertComment(vWorkarea, vAccdt, vAccseq, ObjMyUser.EmpID, vComment) Then
            GoTo SaveABOResult_Error
        End If
    End If
    
    'BBS303에 Front, Back 둘다 입력되어 있으면 Vfydt박아주고 LAB302 도 업뎃 해준다.
    If DoubleCheck(vWorkarea, vAccdt, vAccseq) = False Then
        GoTo SaveABOResult_Error
    End If
    
    'LAB302가 Vfy가 되었으면 LAB102 Status 반영
    If UpdateOrdStatus(vWorkarea, vAccdt, vAccseq, strVfydt, strVfytm, ObjMyUser.EmpID) = False Then
        GoTo SaveABOResult_Error
    End If
            
    'LAB302가 Vfy가 되었으면 LAB201 Status반영
    If UpdateAccStatus(vWorkarea, vAccdt, vAccseq, ObjMyUser.EmpID, strVfydt, strVfytm) = False Then
        GoTo SaveABOResult_Error
    End If
    
    DBConn.CommitTrans
    SaveABOResult = True
    Exit Function
    
SaveABOResult_Error:
    DBConn.RollbackTrans
    SaveABOResult = False
End Function

Private Function UpdateRemark(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                     ByVal vRemark As String) As Boolean
    Dim strSQL As String
    
On Error GoTo UpdateRemark_Error
    
    strSQL = "update " & T_LAB201 & " " & _
           "set " & DBW("rmkcd=", vRemark) & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " "

    DBConn.Execute strSQL
    UpdateRemark = True
    Exit Function
    
UpdateRemark_Error:
    UpdateRemark = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function ModifyABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                ByVal vTyping As Long, _
                                ByVal vABO As String, ByVal vRh As String, ByVal vABOSub As String, ByVal vRhSub As String) As Boolean
    '-----------------------------------------------------------------------------------------
    'PARAMETER 의 용도
    '-----------------------------------------------------------------------------------------
    'typing   = 0     : Front Typing
    '         = 1     : Back typing
    '
    'isVerify = false : 저장만한다. (id만 입력)
    '         = true  : verify한다.
    '-----------------------------------------------------------------------------------------
    '저장시 주의사항
    '-----------------------------------------------------------------------------------------
    'LAB302의 RSTTYPE을 'F'로 만든다
    '-----------------------------------------------------------------------------------------
    
    Dim strVfydt As String
    Dim strVfytm As String
    
    strVfydt = Format(GetSystemDate, "YYYYMMDD")
    strVfytm = Format(GetSystemDate, "HHMMSS")
    '-----------------------------------------------------------------------------------------
    '진행순서
    '-----------------------------------------------------------------------------------------
    'STEP 1) CHECK
    'SETP 3) ABO 저장
    'STEP 2) REMARK, COMMENT 저장
    'STEP 4) LAB302저장(필요하면)
    '-----------------------------------------------------------------------------------------
    
On Error GoTo ModifyABOResult_Error

    DBConn.BeginTrans
    
    'BBS303 을 하나 더 생성(복사본)
    If Not InsertABO(vWorkarea, vAccdt, vAccseq, True, vTyping, vABO, vRh, vABOSub, vRhSub, ObjMyUser.EmpID, strVfydt, strVfytm) Then
        GoTo ModifyABOResult_Error
    End If
    
    'BBS308 수정내역 생성
    If ModifyDoubleCheck(vWorkarea, vAccdt, vAccseq) = False Then
        GoTo ModifyABOResult_Error
    End If
    
    'LAB102 Status를 수정으로 변경
    If ModifyOrdStatus(vWorkarea, vAccdt, vAccseq, strVfydt, strVfytm, ObjMyUser.EmpID) = False Then
        GoTo ModifyABOResult_Error
    End If
    
    DBConn.CommitTrans
    ModifyABOResult = True
    Exit Function
    
ModifyABOResult_Error:
    DBConn.RollbackTrans
    ModifyABOResult = False
End Function

Private Function UpdateAccStatus(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                   ByVal VfyId As String, ByVal vfydt As String, ByVal vfytm As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    Dim cnt As Long
    
    strSQL = "SELECT count(*) cnt " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " " & _
           "AND (vfyid is not null or " & DBW("vfyid<>", "") & ") "

    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
        GoTo UpdateAccStatus_Error
    End If
    
    With Rs
        If .RecordCount <= 0 Then
            cnt = 0
        Else
            cnt = .Fields("cnt").Value & ""
        End If
    End With
    Set Rs = Nothing
    
    strSQL = "update " & T_LAB201 & " " & _
           "set   " & DBW("reqinputcnt=", cnt) & " " & _
           "WHERE " & DBW("workarea   =", vWorkarea) & " " & _
           "AND   " & DBW("accdt      =", vAccdt) & " " & _
           "AND   " & DBW("accseq     =", vAccseq) & " "
           
    DBConn.Execute strSQL
    
    strSQL = "update " & T_LAB201 & " " & _
           "set " & DBW("stscd=", "5") & "," & _
                    DBW("vfyid=", VfyId) & "," & _
                    DBW("vfydt=", vfydt) & "," & _
                    DBW("vfytm=", vfytm) & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt=", vAccdt) & " " & _
           "AND   " & DBW("accseq=", vAccseq) & " " & _
           "AND   reqtotcnt=reqinputcnt "
           
    DBConn.Execute strSQL
           
    UpdateAccStatus = True
    Exit Function
    
UpdateAccStatus_Error:
    UpdateAccStatus = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function GetCountResult(ByVal ptid As String, ByVal orddt As String, ByVal ordno As String, ByVal ordseq As String) As Long
    Dim strSQL As String
    Dim Rs As Recordset
    
On Error GoTo GetCountResult_Error

    strSQL = "SELECT count(*) cnt " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("ptid  =", ptid) & " " & _
           "AND   " & DBW("orddt =", orddt) & " " & _
           "AND   " & DBW("ordno =", ordno) & " " & _
           "AND   " & DBW("ordseq=", ordseq) & " " & _
           "AND (vfydt is null or " & DBW("vfydt=", "") & ") "
           
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        GetCountResult = -1
        Set Rs = Nothing
        Exit Function
    End If
    
    With Rs
        If .RecordCount < 1 Then
            MsgBox "상태를 변경하기 위해 결과내역을 읽는중 오류가 발생하였읍니다.(2)", vbCritical, "오류"
            Set Rs = Nothing
            GetCountResult = -1
            Exit Function
        Else
            GetCountResult = .Fields("cnt").Value & ""
        End If
    End With
    Set Rs = Nothing
    
    Exit Function
    
GetCountResult_Error:
'    DBConn.DisplayErrors
    GetCountResult = -1
End Function

Private Function UpdateOrdStatus(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                  ByVal vfydt As String, ByVal vfytm As String, ByVal VfyId As String) As Boolean
    Dim ptid As String
    Dim orddt As String
    Dim ordno As String
    Dim ordseq As String
    
    Dim Rs As Recordset
    Dim strSQL As String
    Dim cnt As Long
    Dim i As Long
    
On Error GoTo UpdateOrdStatus_Error

    strSQL = "SELECT distinct ptid,orddt,ordno,ordseq " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " "
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
        GoTo UpdateOrdStatus_Error
    End If
    
    With Rs
        For i = 1 To .RecordCount
            ptid = .Fields("ptid").Value & ""
            orddt = .Fields("orddt").Value & ""
            ordno = .Fields("ordno").Value & ""
            ordseq = .Fields("ordseq").Value & ""
                        
            cnt = GetCountResult(ptid, orddt, ordno, ordseq)
            If cnt <= 0 Then
'                strSQL = "update " & T_LAB102 & " " & _
'                       "set " & DBW("stscd   =", "5") & ", " & _
'                                DBW("examdt  =", vfydt) & "," & _
'                                DBW("examtm  =", vfytm) & "," & _
'                                DBW("examdoct=", VfyId) & " " & _
'                       "WHERE " & DBW("ptid  =", ptid) & " " & _
'                       "AND   " & DBW("orddt =", orddt) & " " & _
'                       "AND   " & DBW("ordno =", ordno) & " " & _
'                       "AND   " & DBW("ordseq=", ordseq) & " "
                       
                '예수병원용
                strSQL = "update mdexmort " & _
                       "set " & DBW("stscd   =", "5") & ", " & _
                                " execdate=to_date(" & vfydt & vfytm & ",'yyyymmdd hh24:mi:ss') ," & _
                                DBW("cofmdr=", VfyId) & " " & _
                       "WHERE " & DBW("patno  =", ptid) & " " & _
                       "AND   " & " orddate=to_date(" & orddt & ",'yyyymmdd') " & _
                       "AND   " & DBW("ordseqno =", ordno) & " "
                       
                DBConn.Execute strSQL
            End If
            
            .MoveNext
        Next i
    End With
    Set Rs = Nothing
    
    
    UpdateOrdStatus = True
    Exit Function
    
UpdateOrdStatus_Error:
    UpdateOrdStatus = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function ModifyOrdStatus(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                        ByVal vfydt As String, ByVal vfytm As String, ByVal VfyId As String) As Boolean
    Dim ptid As String
    Dim orddt As String
    Dim ordno As String
    Dim ordseq As String
    
    Dim Rs As Recordset
    Dim strSQL As String
    Dim cnt As Long
    Dim i As Long
    
On Error GoTo ModifyOrdStatus_Error

    strSQL = "SELECT distinct ptid,orddt,ordno,ordseq " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " "
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
        GoTo ModifyOrdStatus_Error
    End If
    
'    MsgBox "처방내역 업데이트", vbExclamation
    
    With Rs
        For i = 1 To .RecordCount
            ptid = .Fields("ptid").Value & ""
            orddt = .Fields("orddt").Value & ""
            ordno = .Fields("ordno").Value & ""
            ordseq = .Fields("ordseq").Value & ""
                        
            cnt = GetCountResult(ptid, orddt, ordno, ordseq)
            If cnt <= 0 Then
'                strSQL = "update " & T_LAB102 & " " & _
'                       "set " & DBW("stscd    =", "6") & ", " & _
'                                DBW("examdt   =", vfydt) & "," & _
'                                DBW("examtm   =", vfytm) & "," & _
'                                DBW("examdoct =", VfyId) & " " & _
'                       "WHERE " & DBW("ptid   =", ptid) & " " & _
'                       "AND   " & DBW("orddt  =", orddt) & " " & _
'                       "AND   " & DBW("ordno  =", ordno) & " " & _
'                       "AND   " & DBW("ordseq =", ordseq) & " " & _
'                       "AND   " & DBW("stscd >=", "5") & " "
                
                '예수병원용
                strSQL = "update mdexmort " & _
                       "set " & DBW("stscd    =", "6") & ", " & _
                                " execdate=to_date(" & vfydt & vfytm & ",'yyyymmdd hh24:mi:ss')," & _
                                DBW("cofmdr =", VfyId) & " " & _
                       "WHERE " & DBW("patno   =", ptid) & " " & _
                       "AND   orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                       "AND   " & DBW("ordseqno  =", ordno) & " " & _
                       "AND   " & DBW("stscd >=", "5") & " "
                DBConn.Execute strSQL
            End If
            
            .MoveNext
        Next i
    End With
    Set Rs = Nothing
    
    
    ModifyOrdStatus = True
    Exit Function
    
ModifyOrdStatus_Error:
    ModifyOrdStatus = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function CopyLAB302(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, ByVal vTestCd As String, _
                            ByVal vVfyid As String, ByVal vvfydt As String, ByVal vVfytm As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    Dim cnt As Long
    Dim maxseq As Long
    Dim rstval As String
    Dim rstcd  As String
    Dim rstunit As String
    Dim hldiv   As String
    Dim dpdiv   As String
    Dim mfydt   As String
    Dim mfytm   As String
    Dim mfyid   As String
    Dim vfydt   As String
    Dim vfytm   As String
    Dim VfyId   As String
    Dim rsttype As String

On Error GoTo CopyLAB302_Error

    strSQL = "SELECT * " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " " & _
           "AND   " & DBW("testcd  =", vTestCd) & " "
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        MsgBox "결과내역을 읽는 중 오류가 발생하였읍니다.", vbCritical, "오류"
        Set Rs = Nothing
        CopyLAB302 = False
        Exit Function
    End If
    
    With Rs
        cnt = 0
        If .RecordCount > 0 Then
            cnt = 1
            rstval = .Fields("rstval").Value & ""
            rstcd = .Fields("rstcd").Value & ""
            rstunit = .Fields("rstunit").Value & ""
            hldiv = .Fields("hldiv").Value & ""
            dpdiv = .Fields("dpdiv").Value & ""
            mfydt = vvfydt
            mfytm = vVfytm
            mfyid = vVfyid
            vfydt = .Fields("vfydt").Value & ""
            vfytm = .Fields("vfytm").Value & ""
            VfyId = .Fields("vfyid").Value & ""
            rsttype = .Fields("rsttype").Value & ""
        End If
    End With
    Set Rs = Nothing
    
    If cnt <= 0 Then Exit Function
    
    
    strSQL = "SELECT max(seq) maxseq " & _
           "FROM " & T_LAB308 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " " & _
           "AND   " & DBW("testcd  =", vTestCd) & " "
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        MsgBox "결과내역을 읽는 중 오류가 발생하였읍니다.(2)", vbCritical, "오류"
        Set Rs = Nothing
        CopyLAB302 = False
        Exit Function
    End If
    
    With Rs
        If .RecordCount > 1 Then
            maxseq = .Fields("maxseq").Value & ""
        Else
            maxseq = 0
        End If
    End With
    Set Rs = Nothing
    
    
    strSQL = "insert into " & T_LAB308 & "(vWorkarea,vAccdt,vAccseq,testcd,seq" & _
                                       ",rstval,rstcd,rstunit,hldiv,dpdiv" & _
                                       ",mfyid,mfydt,mfytm,vfyid,vfydt,vfytm" & _
                                       ",rsttype) " & _
           "values (" & _
                    DBV("workarea", vWorkarea) & "," & _
                    DBV("accdt", vAccdt) & "," & _
                    DBV("accseq", vAccseq) & "," & _
                    DBV("testcd", vTestCd) & "," & _
                    DBV("seq", maxseq + 1) & "," & _
                    DBV("rstval", rstval) & "," & _
                    DBV("rstcd", rstcd) & "," & _
                    DBV("rstunit", rstunit) & "," & _
                    DBV("hldiv", hldiv) & "," & _
                    DBV("dpdiv", dpdiv) & "," & _
                    DBV("mfyid", mfyid) & "," & _
                    DBV("mfydt", mfydt) & "," & _
                    DBV("mfytm", mfytm) & "," & _
                    DBV("vfyid", VfyId) & "," & _
                    DBV("vfydt", vfydt) & "," & _
                    DBV("vfytm", vfytm) & "," & _
                    DBV("rsttype", rsttype) & ")"

    DBConn.Execute strSQL
    
    CopyLAB302 = True
    Exit Function
    
CopyLAB302_Error:
    CopyLAB302 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function UpdateDetailItem(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                ByVal vTestCd As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    Dim detailfg As String
    Dim VfyId As String
    Dim vfydt As String
    Dim vfytm As String
    
On Error GoTo UpdateDetailItem_Error

    strSQL = "SELECT vfyid,vfydt,vfytm,detailfg " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " " & _
           "AND " & DBW("testcd=", vTestCd) & " "
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
        UpdateDetailItem = False
        Set Rs = Nothing
        Exit Function
    End If

    With Rs
        If .RecordCount > 0 Then
            detailfg = .Fields("detailfg").Value & ""
            VfyId = .Fields("vfyid").Value & ""
            vfydt = .Fields("vfydt").Value & ""
            vfytm = .Fields("vfytm").Value & ""
        Else
            detailfg = ""
        End If
    End With
    Set Rs = Nothing
        
    'Detail항목이 아니다.----------------------------------
    If detailfg = "" Then
        UpdateDetailItem = True
        Exit Function
    End If
    
    strSQL = "update " & T_LAB302 & " " & _
           "set " & DBW("vfyid=", VfyId) & ", " & _
               "" & DBW("vfydt=", vfydt) & "," & _
               "" & DBW("vfytm=", vfytm) & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " " & _
           "AND " & DBW("detailfg=", detailfg) & " " & _
           "AND " & DBW("rstdiv=", "*") & " "
    DBConn.Execute strSQL
    
    UpdateDetailItem = True
    Exit Function
    
UpdateDetailItem_Error:
    UpdateDetailItem = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function DoubleCheck(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
    Dim Rs As Recordset
    Dim strSQL As String
    
    Dim abo1 As String
    Dim rh1 As String
    Dim abo2 As String
    Dim rh2 As String
    Dim ABOSub As String
    Dim RhSub As String
    
    Dim strVfyid1 As String
    Dim strVfydt1 As String
    Dim strVfytm1 As String
    
    Dim strVfyid2 As String
    Dim strVfydt2 As String
    Dim strVfytm2 As String
    
    Dim VfyId As String
    Dim vfydt As String
    Dim vfytm As String
    
On Error GoTo DoubleCheck_Error

    Set Rs = GetABOResultInfo(vWorkarea, vAccdt, vAccseq)
    If Rs Is Nothing Then
        DoubleCheck = False
    End If
    
    With Rs
        If .RecordCount < 1 Then
            DoubleCheck = False
        Else
            abo1 = Trim(.Fields("abo1").Value & "")
            rh1 = Trim(.Fields("rh1").Value & "")
            abo2 = Trim(.Fields("abo2").Value & "")
            rh2 = Trim(.Fields("rh2").Value & "")
            ABOSub = Trim(.Fields("abosub").Value & "")
            RhSub = Trim(.Fields("rhsub").Value & "")
            
            strVfyid1 = Trim(.Fields("vfyid1").Value & "")
            strVfydt1 = Trim(.Fields("vfydt1").Value & "")
            strVfytm1 = Trim(.Fields("vfytm1").Value & "")
            
            strVfyid2 = Trim(.Fields("vfyid2").Value & "")
            strVfydt2 = Trim(.Fields("vfydt2").Value & "")
            strVfytm2 = Trim(.Fields("vfytm2").Value & "")
        End If
    End With
    
    
    If strVfydt1 <> "" And strVfydt2 <> "" Then
        If strVfydt1 > strVfydt2 Then
            VfyId = strVfyid1
            vfydt = strVfydt1
            vfytm = strVfytm1
        ElseIf strVfydt2 > strVfydt1 Then
            VfyId = strVfyid2
            vfydt = strVfydt2
            vfytm = strVfytm2
        Else
            If strVfytm1 > strVfytm2 Then
                VfyId = strVfyid1
                vfydt = strVfydt1
                vfytm = strVfytm1
            Else
                VfyId = strVfyid2
                vfydt = strVfydt2
                vfytm = strVfytm2
            End If
        End If
    Else
        DoubleCheck = True
        Exit Function
    End If
    
    strSQL = "update " & T_BBS303 & " " & _
           "set " & DBW("vfydt=", vfydt) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND seq=(SELECT max(seq) " & _
                        "FROM " & T_BBS303 & " " & _
                        "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                        "AND " & DBW("accdt=", vAccdt) & " " & _
                        "AND " & DBW("accseq=", vAccseq) & ")"
    DBConn.Execute strSQL
    
    
    If CODE_ABOFRONT <> "" Then
        If abo1 <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", abo1) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_ABOFRONT) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_ABOFRONT) & " "
        DBConn.Execute strSQL
        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
        If Not UpdateDetailItem(vWorkarea, vAccdt, vAccseq, CODE_ABOFRONT) Then
            GoTo DoubleCheck_Error
        End If
    End If
        
    If CODE_ABOBACK <> "" Then
        If abo2 <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", abo2) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_ABOBACK) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_ABOBACK) & " "
        DBConn.Execute strSQL
        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
        If Not UpdateDetailItem(vWorkarea, vAccdt, vAccseq, CODE_ABOBACK) Then
            GoTo DoubleCheck_Error
        End If
    End If

    If CODE_RH <> "" Then
        If rh1 <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", rh1) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_RH) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_RH) & " "
        DBConn.Execute strSQL
        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
        If Not UpdateDetailItem(vWorkarea, vAccdt, vAccseq, CODE_RH) Then
            GoTo DoubleCheck_Error
        End If
    End If
        
    
    If CODE_ABOSUB <> "" Then
        If ABOSub <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", ABOSub) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_ABOSUB) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_ABOSUB) & " "
        DBConn.Execute strSQL
        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
        If Not UpdateDetailItem(vWorkarea, vAccdt, vAccseq, CODE_ABOSUB) Then
            GoTo DoubleCheck_Error
        End If
    End If
        
    If CODE_RHSUB <> "" Then
        If RhSub <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", RhSub) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_RHSUB) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_RHSUB) & " "
        DBConn.Execute strSQL
        '만일 Detail항목으로 잡혀있으면 대표항목을 Verify시킨다.
        If Not UpdateDetailItem(vWorkarea, vAccdt, vAccseq, CODE_RHSUB) Then
            GoTo DoubleCheck_Error
        End If
    End If
    
    DoubleCheck = True
    Exit Function
    
DoubleCheck_Error:
    DoubleCheck = False
    MsgBox Err.Description, vbExclamation
End Function
Private Function InsertComment(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                    ByVal VfyId As String, ByVal comment As String) As Boolean
    '
    Dim strSQL As String
    Dim Rs As Recordset
    Dim maxseq As Long
    Dim rsttext As String
    
On Error GoTo InsertComment_Error
    InsertComment = True
    
    Set Rs = GetAccComment(vWorkarea, vAccdt, vAccseq)
    If Rs Is Nothing Then
        InsertComment = False
        Exit Function
    End If
    
    With Rs
        If .RecordCount <= 0 Then
            maxseq = 0
            rsttext = ""
        Else
            maxseq = .Fields("seq").Value & ""
            rsttext = .Fields("rsttxt").Value & ""
        End If
    End With
    Set Rs = Nothing
    
    If InsertComment = False Then Exit Function
    
    If Trim(rsttext) = Trim(comment) Then
        InsertComment = True
        Exit Function
    End If
    
    strSQL = "insert into " & T_LAB304 & "(vWorkarea,vAccdt,vAccseq,seq,vfyid,rsttxt) " & _
           "values(" & _
                    DBV("workarea", vWorkarea) & "," & _
                    DBV("accdt", vAccdt) & "," & _
                    DBV("accseq", vAccseq) & "," & _
                    DBV("seq", maxseq + 1) & "," & _
                    DBV("vfyid", VfyId) & "," & _
                    DBV("rsttxt", comment) & "" & _
           ")"
        
    DBConn.Execute strSQL
    InsertComment = True
    Exit Function
    
InsertComment_Error:
    InsertComment = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function InsertABO(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                ByVal vVerify As Boolean, _
                                ByVal vTyping As Long, _
                                ByVal vABO As String, ByVal vRh As String, ByVal vABOSub As String, ByVal vRhSub As String, _
                                ByVal vVfyid As String, ByVal vvfydt As String, ByVal vVfytm As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    Dim seq As Long
    Dim SpcCd As String
    Dim statfg As String
    Dim abo1 As String
    Dim rh1 As String
    Dim abo2 As String
    Dim rh2 As String
    Dim ABOSub As String
    Dim RhSub As String
    Dim strVfyid1 As String
    Dim strVfydt1 As String
    Dim strVfytm1 As String
    Dim strVfyid2 As String
    Dim strVfydt2 As String
    Dim strVfytm2 As String
    Dim vfydt As String
    Dim ptid As String
    
On Error GoTo InsertABO_Error

    strSQL = "SELECT * " & _
           "FROM " & T_BBS303 & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " " & _
           "AND seq=(SELECT max(seq) " & _
                    "FROM " & T_BBS303 & " " & _
                    "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                    "AND " & DBW("accdt=", vAccdt) & " " & _
                    "AND " & DBW("accseq=", vAccseq) & ")"
                    
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
        GoTo InsertABO_Error
    End If
    
    With Rs
        If .RecordCount > 0 Then
            seq = .Fields("seq").Value & ""
            SpcCd = .Fields("spccd").Value & ""
            statfg = .Fields("statfg").Value & ""
            abo1 = .Fields("abo1").Value & ""
            rh1 = .Fields("rh1").Value & ""
            abo2 = .Fields("abo2").Value & ""
            rh2 = .Fields("rh2").Value & ""
            ABOSub = .Fields("abosub").Value & ""
            RhSub = .Fields("rhsub").Value & ""
            strVfyid1 = .Fields("vfyid1").Value & ""
            strVfydt1 = .Fields("vfydt1").Value & ""
            strVfytm1 = .Fields("vfytm1").Value & ""
            strVfyid2 = .Fields("vfyid2").Value & ""
            strVfydt2 = .Fields("vfydt2").Value & ""
            strVfytm2 = .Fields("vfytm2").Value & ""
            vfydt = .Fields("vfydt").Value & ""
            ptid = .Fields("ptid").Value & ""
            InsertABO = True
        Else
            MsgBox "결과내역을 읽을 수 없습니다.전산실로 연락바랍니다.", vbCritical, "오류"
            InsertABO = False
        End If
    End With
    Set Rs = Nothing
    
    If InsertABO = False Then Exit Function
    
    '기존내역을 복사한다.(seq를 1증가시켜서)
    strSQL = "insert into " & T_BBS303 & "(vWorkarea,vAccdt,vAccseq,seq,spccd,statfg," & _
                                        "abo1,rh1,abo2,rh2,abosub,rhsub," & _
                                        "vfyid1,strVfydt1,strVfytm1,strVfyid2,strVfydt2,strVfytm2," & _
                                        "vfydt,ptid) " & _
           "values(" & _
                    DBV("workarea", vWorkarea) & "," & _
                    DBV("accdt", vAccdt) & "," & _
                    DBV("accseq", vAccseq) & "," & _
                    DBV("seq", seq + 1) & "," & _
                    DBV("spccd", SpcCd) & "," & _
                    DBV("statfg", statfg) & "," & _
                    DBV("abo1", abo1) & "," & _
                    DBV("rh1", rh1) & "," & _
                    DBV("abo2", abo2) & "," & _
                    DBV("rh2", rh2) & "," & _
                    DBV("abosub", ABOSub) & "," & _
                    DBV("rhsub", RhSub) & "," & _
                    DBV("vfyid1", strVfyid1) & "," & _
                    DBV("vfydt1", strVfydt1) & "," & _
                    DBV("vfytm1", strVfytm1) & "," & _
                    DBV("vfyid2", strVfyid2) & "," & _
                    DBV("vfydt2", strVfydt2) & "," & _
                    DBV("vfytm2", strVfytm2) & "," & _
                    DBV("vfydt", vvfydt) & "," & _
                    DBV("ptid", ptid) & ")"
    DBConn.Execute strSQL
    
    InsertABO = UpdateABO(vWorkarea, vAccdt, vAccseq, True, vTyping, vABO, vRh, vABOSub, vRhSub, vVfyid, vvfydt, vVfytm)
    
    'Double Check완료되었으면, 결과 2개 모두 수정
    '                          LAB302의 복사본 마련, LAB308
    If vTyping = 0 Then
        If strVfyid2 = "0" Then strVfyid2 = ""
        If strVfyid2 <> "" Then
            InsertABO = UpdateABO(vWorkarea, vAccdt, vAccseq, True, "1", vABO, vRh, vABOSub, vRhSub, vVfyid, vvfydt, vVfytm)
        End If
    Else
        If strVfyid1 = "0" Then strVfyid1 = ""
        If strVfyid1 <> "" Then
            InsertABO = UpdateABO(vWorkarea, vAccdt, vAccseq, True, "0", vABO, vRh, vABOSub, vRhSub, vVfyid, vvfydt, vVfytm)
        End If
    End If
    
    
    Exit Function
    
InsertABO_Error:
    InsertABO = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function UpdateABO(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                ByVal isVerify As Boolean, _
                                ByVal typing As Long, _
                                ByVal pABO As String, ByVal pRh As String, ByVal pABOSub As String, ByVal pRhSub As String, _
                                ByVal EmpID As String, ByVal vfydt As String, ByVal vfytm As String) As Boolean
    
    Dim strSQL As String
    
On Error GoTo UpdateABO_Error

    If typing = 0 Then
        strSQL = "update " & T_BBS303 & " " & _
               "set " & DBW("abo1=", pABO) & "," & _
                   "" & DBW("rh1=", pRh) & "," & _
                   "" & DBW("abosub=", pABOSub) & "," & _
                   "" & DBW("rhsub=", pRhSub) & ", " & _
                   "" & DBW("vfyid1=", EmpID) & " "
        If isVerify Then
            strSQL = strSQL & "," & _
                   "" & DBW("vfydt1=", vfydt) & "," & _
                   "" & DBW("vfytm1=", vfytm) & " "
        End If
    Else
        strSQL = "update " & T_BBS303 & " " & _
               "set " & DBW("abo2=", pABO) & "," & _
                   "" & DBW("rh2=", pRh) & "," & _
                   "" & DBW("abosub=", pABOSub) & "," & _
                   "" & DBW("rhsub=", pRhSub) & ", " & _
                   "" & DBW("vfyid2=", EmpID) & " "
        If isVerify Then
            strSQL = strSQL & "," & _
                   "" & DBW("vfydt2=", vfydt) & "," & _
                   "" & DBW("vfytm2=", vfytm) & " "
        End If
    End If
    
    strSQL = strSQL & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND " & DBW("accdt=", vAccdt) & " " & _
           "AND " & DBW("accseq=", vAccseq) & " " & _
           "AND seq=(SELECT max(seq) " & _
                    "FROM " & T_BBS303 & " " & _
                    "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                    "AND " & DBW("accdt=", vAccdt) & " " & _
                    "AND " & DBW("accseq=", vAccseq) & ")"

    
    DBConn.Execute strSQL
    UpdateABO = True
    Exit Function
    
UpdateABO_Error:
    UpdateABO = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function CheckABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                ByVal vVerify As Boolean, _
                                ByVal vTyping As Long, _
                                ByVal vABO As String, ByVal vRh As String, _
                                ByVal vABOSub As String, ByVal vRhSub As String, _
                                ByVal vEmpID As String) As Boolean
    Dim Rs As Recordset
    
    Dim strABO1 As String
    Dim strRh1 As String
    Dim strABO2 As String
    Dim strRh2 As String
    
    Dim strVfyid1 As String
    Dim strVfydt1 As String
    Dim strVfytm1 As String
    
    Dim strVfyid2 As String
    Dim strVfydt2 As String
    Dim strVfytm2 As String
    
    
    If Not vVerify Then
        'Verify가 아니면 저장할 수 있다.
        CheckABOResult = True
        Exit Function
    End If

    Set Rs = GetABOResultInfo(vWorkarea, vAccdt, vAccseq)
    If Rs Is Nothing Then
        CheckABOResult = False
        Exit Function
    End If

    With Rs
        strABO1 = .Fields("abo1").Value & ""
        strRh1 = .Fields("rh1").Value & ""
        strABO2 = .Fields("abo2").Value & ""
        strRh2 = .Fields("rh2").Value & ""
        
        strVfyid1 = .Fields("vfyid1").Value & ""
        strVfydt1 = .Fields("vfydt1").Value & ""
        strVfytm1 = .Fields("vfytm1").Value & ""
        
        strVfyid2 = .Fields("vfyid2").Value & ""
        strVfydt2 = .Fields("vfydt2").Value & ""
        strVfytm2 = .Fields("vfytm2").Value & ""
    End With
    Set Rs = Nothing

'이넘 탈 확률 0%
    If strVfydt1 <> "" And strVfydt2 <> "" Then
        MsgBox "이미 Double Check 완료되었읍니다.", vbCritical
        CheckABOResult = False
        Exit Function
    End If
    
    
    Select Case vTyping
        Case 0:     '--------------------------------------------------Front
            '이넘 탈 확률 없음
            If strVfyid1 <> "" Then
                If strVfyid1 <> vEmpID Then
                    MsgBox "이미 다른 사용자가 Front Typing 결과를 입력하였읍니다.", vbCritical, "오류"
                    CheckABOResult = False
                    Exit Function
                End If
            End If
            If strVfyid2 = "" Then
                'Back Typing이 입력이 안되어있다.
                CheckABOResult = True
                Exit Function
            End If
            If strVfydt2 = "" Then
                'Back Typing이 Verify되지 않았다.
                CheckABOResult = True
                Exit Function
            End If
            If strABO2 <> vABO Or strRh2 <> vRh Then
                MsgBox "앞서 입력된 결과와 결과값이 다릅니다.", vbCritical, "오류"
                CheckABOResult = False
                Exit Function
            End If
            CheckABOResult = True
        Case 1:     '--------------------------------------------------Back
            '이넘 탈 확률 없음
            If strVfyid2 <> "" Then
                If strVfyid2 <> vEmpID Then
                    MsgBox "이미 다른 사용자가 Back Typing 결과를 입력하였읍니다.", vbCritical, "오류"
                    CheckABOResult = False
                    Exit Function
                End If
            End If
            If strVfyid1 = "" Then
                'Front Typing이 입력이 안되어있다.
                CheckABOResult = True
                Exit Function
            End If
            If strVfydt1 = "" Then
                'Front Typing이 Verify되지 않았다.
                CheckABOResult = True
                Exit Function
            End If
            If strABO1 <> vABO Or strRh1 <> vRh Then
                MsgBox "앞서 입력된 결과와 결과값이 다릅니다.", vbCritical, "오류"
                CheckABOResult = False
                Exit Function
            End If
            CheckABOResult = True
        Case Else:  '--------------------------------------------------???
            CheckABOResult = False
    End Select
End Function

Public Function GetWsNoDuration(ByVal Workdt As String, ByVal workCd As String, min As Long, max As Long) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    strSQL = "SELECT min(workseq) minseq,max(workseq) maxseq " & _
           "FROM " & T_LAB301 & " " & _
           "WHERE " & DBW("workdt=", Workdt) & " " & _
           "AND " & DBW("workcd=", workCd) & " "
    
    min = 0
    max = 0
    
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        GetWsNoDuration = False
        Exit Function
    End If
    
    With Rs
        If .RecordCount <= 0 Then
            GetWsNoDuration = False
        Else
            min = Val(.Fields("minseq").Value & "")
            max = Val(.Fields("maxseq").Value & "")
            GetWsNoDuration = True
        End If
    End With
    Set Rs = Nothing
    
End Function
Private Function ModifyDoubleCheck(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
    Dim Rs As Recordset
    Dim strSQL As String
    
    Dim abo1 As String
    Dim rh1 As String
    Dim abo2 As String
    Dim rh2 As String
    Dim ABOSub As String
    Dim RhSub As String
    
    Dim strVfyid1 As String
    Dim strVfydt1 As String
    Dim strVfytm1 As String
    
    Dim strVfyid2 As String
    Dim strVfydt2 As String
    Dim strVfytm2 As String
    
    Dim VfyId As String
    Dim vfydt As String
    Dim vfytm As String
    
On Error GoTo ModifyDoubleCheck_Error

    Set Rs = GetABOResultInfo(vWorkarea, vAccdt, vAccseq)
    If Rs Is Nothing Then
        ModifyDoubleCheck = False
    End If
    
    With Rs
        If .RecordCount < 1 Then
            ModifyDoubleCheck = False
        Else
            abo1 = Trim(.Fields("abo1").Value & "")
            rh1 = Trim(.Fields("rh1").Value & "")
            abo2 = Trim(.Fields("abo2").Value & "")
            rh2 = Trim(.Fields("rh2").Value & "")
            ABOSub = Trim(.Fields("abosub").Value & "")
            RhSub = Trim(.Fields("rhsub").Value & "")
            
            strVfyid1 = Trim(.Fields("vfyid1").Value & "")
            strVfydt1 = Trim(.Fields("vfydt1").Value & "")
            strVfytm1 = Trim(.Fields("vfytm1").Value & "")
            
            strVfyid2 = Trim(.Fields("vfyid2").Value & "")
            strVfydt2 = Trim(.Fields("vfydt2").Value & "")
            strVfytm2 = Trim(.Fields("vfytm2").Value & "")
            
            If strVfyid1 = "0" Then strVfyid1 = ""
            If strVfyid2 = "0" Then strVfyid2 = ""
        End If
    End With
    
    
    If strVfydt1 <> "" And strVfydt2 <> "" Then
        If strVfydt1 > strVfydt2 Then
            VfyId = strVfyid1
            vfydt = strVfydt1
            vfytm = strVfytm1
        ElseIf strVfydt2 > strVfydt1 Then
            VfyId = strVfyid2
            vfydt = strVfydt2
            vfytm = strVfytm2
        Else
            If strVfytm1 > strVfytm2 Then
                VfyId = strVfyid1
                vfydt = strVfydt1
                vfytm = strVfytm1
            Else
                VfyId = strVfyid2
                vfydt = strVfydt2
                vfytm = strVfytm2
            End If
        End If
    Else
        ModifyDoubleCheck = True
        Exit Function
    End If
    
    strSQL = "update " & T_BBS303 & " " & _
           "set " & DBW("vfydt=", vfydt) & " " & _
           "WHERE " & DBW("workarea=", vWorkarea) & " " & _
           "AND   " & DBW("accdt   =", vAccdt) & " " & _
           "AND   " & DBW("accseq  =", vAccseq) & " " & _
           "AND seq=(SELECT max(seq) " & _
                    "FROM " & T_BBS303 & " " & _
                    "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                    "AND   " & DBW("accdt   =", vAccdt) & " " & _
                    "AND   " & DBW("accseq  =", vAccseq) & ")"
    '
    DBConn.Execute strSQL
    
    
    'LAB302 의 Copy본을 만든다.
    If CODE_ABOFRONT <> "" Then
        If Not CopyLAB302(vWorkarea, vAccdt, vAccseq, CODE_ABOFRONT, VfyId, vfydt, vfytm) Then
            ModifyDoubleCheck = False
            Exit Function
        End If
        If abo1 <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd  =", abo1) & "," & _
                            DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND   " & DBW("accdt   =", vAccdt) & " " & _
                   "AND   " & DBW("accseq  =", vAccseq) & " " & _
                   "AND   " & DBW("testcd  =", CODE_ABOFRONT) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   DBW("vfyid=", VfyId) & "," & _
                   DBW("vfydt=", vfydt) & "," & _
                   DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND   " & DBW("accdt   =", vAccdt) & " " & _
               "AND   " & DBW("accseq  =", vAccseq) & " " & _
               "AND   " & DBW("testcd  =", CODE_ABOFRONT) & " "
        DBConn.Execute strSQL
    End If
        
    If CODE_ABOBACK <> "" Then
        If CODE_ABOBACK <> CODE_ABOFRONT Then
            If Not CopyLAB302(vWorkarea, vAccdt, vAccseq, CODE_ABOBACK, VfyId, vfydt, vfytm) Then
                ModifyDoubleCheck = False
                Exit Function
            End If
        End If
        If abo2 <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd", abo2, 3) & _
                            DBW("rsttype", "F", 2) & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_ABOBACK) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_ABOBACK) & " "
        DBConn.Execute strSQL
    End If

    If CODE_RH <> "" Then
        If CODE_RH <> CODE_ABOFRONT And CODE_RH <> CODE_ABOBACK Then
            If Not CopyLAB302(vWorkarea, vAccdt, vAccseq, CODE_RH, VfyId, vfydt, vfytm) Then
                ModifyDoubleCheck = False
                Exit Function
            End If
        End If
        If rh1 <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", rh1) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_RH) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_RH) & " "
        DBConn.Execute strSQL
    End If
        
    If CODE_ABOSUB <> "" Then
        If CODE_ABOSUB <> CODE_ABOFRONT And CODE_ABOSUB <> CODE_ABOBACK And CODE_ABOSUB <> CODE_RH Then
            If Not CopyLAB302(vWorkarea, vAccdt, vAccseq, CODE_ABOSUB, VfyId, vfydt, vfytm) Then
                ModifyDoubleCheck = False
                Exit Function
            End If
        End If
        If ABOSub <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", ABOSub) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_ABOSUB) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_ABOSUB) & " "
        DBConn.Execute strSQL
    End If
        
    If CODE_RHSUB <> "" Then
        If CODE_RHSUB <> CODE_ABOFRONT And CODE_RHSUB <> CODE_ABOBACK And CODE_RHSUB <> CODE_RH And CODE_RHSUB = CODE_ABOSUB Then
            If Not CopyLAB302(vWorkarea, vAccdt, vAccseq, CODE_RHSUB, VfyId, vfydt, vfytm) Then
                ModifyDoubleCheck = False
                Exit Function
            End If
        End If
        If RhSub <> "" Then
            strSQL = "update " & T_LAB302 & " " & _
                   "set " & DBW("rstcd=", RhSub) & "," & DBW("rsttype=", "F") & " " & _
                   "WHERE " & DBW("workarea=", vWorkarea) & " " & _
                   "AND " & DBW("accdt=", vAccdt) & " " & _
                   "AND " & DBW("accseq=", vAccseq) & " " & _
                   "AND " & DBW("testcd=", CODE_RHSUB) & " "
            DBConn.Execute strSQL
        End If
        strSQL = "update " & T_LAB302 & " " & _
               "set " & _
                   "" & DBW("vfyid=", VfyId) & "," & _
                   "" & DBW("vfydt=", vfydt) & "," & _
                   "" & DBW("vfytm=", vfytm) & " " & _
               "WHERE " & DBW("workarea=", vWorkarea) & " " & _
               "AND " & DBW("accdt=", vAccdt) & " " & _
               "AND " & DBW("accseq=", vAccseq) & " " & _
               "AND " & DBW("testcd=", CODE_RHSUB) & " "
        DBConn.Execute strSQL
    End If
        
    ModifyDoubleCheck = True
    
    Exit Function
    
ModifyDoubleCheck_Error:
    ModifyDoubleCheck = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function LoadABOOrdList(ByVal vPtid As String, _
                                ByVal vFrDt As String, ByVal vToDt As String, _
                                ByVal vDayfg As String) As Recordset

    Dim strSQL As String
    
    '2005/05/31 modify by legends
    'lab032 테이블의 풀스캔과 뷰 풀스캔으로 인한 속도저하로 힌트 구문 사용
    ''
    
    If vPtid <> "" Then
        strSQL = "SELECT a.orddoct,a.ordtm," & _
                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
                      "c.testnm,d.field3 spcnm " & _
               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a " & _
               "where  " & DBW("a.ptid=", vPtid) & " " & _
               "AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'               "WHERE " & DBW("a.orddiv=", "L") & " " &
    Else
        strSQL = "SELECT /*+ index (d S2LAB032_PK) */ a.orddoct,a.ordtm," & _
                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
                      "c.testnm,d.field3 spcnm " & _
               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a " & _
               "where a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'               "WHERE " & DBW("a.orddiv=", "L") & " " &
    End If
    
    If vDayfg = 0 Then '접수일자로
        strSQL = strSQL & " " & _
               "AND " & DBW("b.rcvdt>=", vFrDt) & " AND " & DBW("b.rcvdt<=", vToDt) & " "
    Else '처방일자로
        strSQL = strSQL & " " & _
               "AND  a.orddt between '" & vFrDt & "' and '" & vToDt & "'" ' " & DBW("a.orddt>=", vFrDt) & " AND " & DBW("a.orddt<=", vToDt) & " "
    End If
    strSQL = strSQL & " order by orddt, ptid"
    
    Set LoadABOOrdList = New Recordset
    Call LoadABOOrdList.Open(strSQL, DBConn)
End Function

Public Function LoadABOOrdList_New(ByVal vPtid As String, _
                                ByVal vFrDt As String, ByVal vToDt As String, _
                                ByVal vDayfg As String, ByVal vBarcode As String) As Recordset

    Dim strSQL As String
    
    '2005/05/31 modify by legends
    'lab032 테이블의 풀스캔과 뷰 풀스캔으로 인한 속도저하로 힌트 구문 사용
    ''
    
    If vPtid <> "" Then
        strSQL = "SELECT a.orddoct,a.ordtm," & _
                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
                      "c.testnm,d.field3 spcnm " & _
               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a " & _
               "where  " & DBW("a.ptid=", vPtid) & " " & _
               "AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'               "WHERE " & DBW("a.orddiv=", "L") & " " &
    ElseIf vBarcode <> "" Then
        strSQL = "SELECT a.orddoct,a.ordtm," & _
                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
                      "c.testnm,d.field3 spcnm " & _
               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a," & T_LAB201 & " e " & _
               "where  " & DBW("e.spcyy=", Mid(vBarcode, 1, 2)) & " " & _
               "AND " & DBW("e.spcno=", Val(Mid(vBarcode, 3))) & " " & _
               "AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno AND a.ptid=e.ptid " & _
               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'               "WHERE " & DBW("a.orddiv=", "L") & " " &
    Else
        strSQL = "SELECT /*+ index (d S2LAB032_PK) */ a.orddoct,a.ordtm," & _
                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
                      "c.testnm,d.field3 spcnm " & _
               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a " & _
               "where a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'               "WHERE " & DBW("a.orddiv=", "L") & " " &
    End If
    
    If vDayfg = 0 Then '접수일자로
        strSQL = strSQL & " " & _
               "AND " & DBW("b.rcvdt>=", vFrDt) & " AND " & DBW("b.rcvdt<=", vToDt) & " "
    Else '처방일자로
        strSQL = strSQL & " " & _
               "AND  a.orddt between '" & vFrDt & "' and '" & vToDt & "'" ' " & DBW("a.orddt>=", vFrDt) & " AND " & DBW("a.orddt<=", vToDt) & " "
    End If
    strSQL = strSQL & " order by orddt, ptid"
    
    Set LoadABOOrdList_New = New Recordset
    Call LoadABOOrdList_New.Open(strSQL, DBConn)
End Function

Public Function GetABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT a.testnm,b.workarea,b.accdt,b.accseq,b.testcd,b.rstcd,b.rstunit,c.field1 as RstCdNm " & _
            " FROM " & T_LAB031 & " c," & T_LAB001 & " a," & T_LAB302 & " b" & _
            " WHERE " & DBW("b.workarea=", vWorkarea) & " " & _
            " AND " & DBW("b.accdt=", vAccdt) & " " & _
            " AND " & DBW("b.accseq=", vAccseq) & " " & _
            " AND a.testcd=b.testcd " & _
            " and a.testdiv='3' " & _
            " AND     " & DBJ(DBW("c.cdindex", LC2_ItemResult, 2)) & _
            " AND     " & DBJ("c.cdval1   =* b.testcd") & _
            " AND     " & DBJ("c.cdval2   =* b.rstcd")
           
    Set GetABOResult = New Recordset
    Call GetABOResult.Open(strSQL, DBConn)
End Function

Public Function GetAccNoDuration(ByVal vAccdt As String) As String
    Dim strSQL As String
    Dim Rs   As Recordset
    
    strSQL = " SELECT max(accseq) as maxseq,min(accseq) as minseq FROM " & T_BBS303 & _
                   " WHERE " & DBW("workarea=", BB_WORKAREA) & _
                   "and " & DBW("accdt=", vAccdt)
                 
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        GetAccNoDuration = Rs.Fields("maxseq").Value & "" & COL_DIV & Rs.Fields("minseq").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Private Sub Class_Initialize()
    Call LoadTestCd
End Sub
