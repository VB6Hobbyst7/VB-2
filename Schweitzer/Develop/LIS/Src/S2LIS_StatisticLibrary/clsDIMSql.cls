VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIMSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*--------------------------------------------------------
'*  기능 : 해당직원의 근무형태, 남은연차를 구함
'*  Parameter
'*      pEmpId : 직원ID
'*  Return : 근무형태, 남은연차
'*--------------------------------------------------------
Public Function GetEmpInfo(ByVal pEmpId As String) As String
    Dim SQL As String
    
    SQL = "SELECT shiftcd, remain FROM " & T_COM006 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId)
    
    GetEmpInfo = SQL
End Function

'*--------------------------------------------------------
'*  기능 : 해당일의 직원정보를 구함
'*  Parameter
'*      pDt : 날짜
'*  Return : 직원ID, 직원이름, 시간, 상태(출근, 지각, 퇴근...)
'*--------------------------------------------------------
Public Function GetEmpFlagInfo(ByVal pDt As String) As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    
    SQL1 = "SELECT A.empid, B.empnm, A.tm, A.flag "
    SQL1 = SQL1 & "FROM " & T_COM006 & " B, " & T_COM013 & " A "
    SQL1 = SQL1 & "WHERE" & DBW("A.dt=", pDt)
    SQL1 = SQL1 & "   AND A.empid=B.empid "
    SQL1 = SQL1 & "GROUP BY A.empid, B.empnm, A.tm, A.flag"
    
    SQL2 = "SELECT A.empid, B.empnm, '' AS tm, A.flag "
    SQL2 = SQL2 & "FROM " & T_COM006 & " B, " & T_COM014 & " A "
    SQL2 = SQL2 & "WHERE" & DBW("A.dt=", pDt)
    SQL2 = SQL2 & "   AND A.empid=B.empid "
    SQL2 = SQL2 & "GROUP BY A.empid, B.empnm, '', A.flag"
    
    SQL3 = "SELECT A.empid, B.empnm, '' AS tm, A.flag "
    SQL3 = SQL3 & "FROM " & T_COM006 & " B, " & T_COM015 & " A "
    SQL3 = SQL3 & "WHERE '" & pDt & "' BETWEEN fromdt AND todt"
    SQL3 = SQL3 & "     AND A.empid=B.empid "
    SQL3 = SQL3 & "GROUP BY A.empid, B.empnm, '', A.flag"
    
    SQL1 = SQL1 & " UNION ALL " & SQL2 & " UNION ALL " & SQL3 & " ORDER BY empid, flag ASC"
    GetEmpFlagInfo = SQL1
End Function

'*--------------------------------------------------------
'*  기능 : 해당직원의 정상근무 일수를 구함(토욜일은 제외)
'*  Parameter
'*      pEmpId : 직원ID
'*      pFromDt : 시작일
'*      pToDt : 종료일
'*      pFromTm : 근무시작시간
'*      pToTm : 근무종료시간
'*  Return : 정상근무 일수
'*--------------------------------------------------------
Public Function GetNormalWorkCnt01(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                                   ByVal pFromTm As String, ByVal pToTm As String) As Integer
    Dim RS As New Recordset
    Dim SQL As String
    
On Error GoTo ErrorMsg:
    SQL = "SELECT Count(empid) AS count FROM " & T_COM013 & " A "
    SQL = SQL & "WHERE" & DBW("A.empid=", pEmpId)
    SQL = SQL & "   AND A.dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL = SQL & "   AND A.flag='1'"
    SQL = SQL & "   AND A.tm<='" & pFromTm & "'"
    SQL = SQL & "   AND '" & pToTm & "' <= (SELECT B.tm FROM " & T_COM013 & " B "
    SQL = SQL & "                           WHERE" & DBW("B.empid=", pEmpId)
    SQL = SQL & "                               AND B.dt=A.dt"
    SQL = SQL & "                               AND B.flag='3')"
    
    RS.Open SQL, dbconn
    If Not (RS.BOF Or RS.EOF) Then
        GetNormalWorkCnt01 = CInt(RS.Fields("count").Value & "")
    End If
    
    Set RS = Nothing
    Exit Function
    
ErrorMsg:
    MsgBox Err.Description, vbOKOnly + vbCritical, "오류"
End Function

'*--------------------------------------------------------
'*  기능 : 해당직원이 토욜일에 정상근무를 했는지 확인
'*  Parameter
'*      pEmpId : 직원ID
'*      pDt : 날짜(토요일)
'*      pFromTm : 근무시작시간
'*      pToTm : 근무종료시간
'*  Return : True / False
'*--------------------------------------------------------
Public Function GetNormalWorkCnt02(ByVal pEmpId As String, ByVal pDt As String, _
                                   ByVal pFromTm As String, ByVal pToTm As String) As Boolean
    Dim RS As New Recordset
    Dim SQL As String
    
On Error GoTo ErrorMsg:
    SQL = "SELECT A.empid FROM " & T_COM013 & " A "
    SQL = SQL & "WHERE" & DBW("A.empid=", pEmpId)
    SQL = SQL & "   AND" & DBW("A.dt=", pDt)
    SQL = SQL & "   AND A.flag='1'"
    SQL = SQL & "   AND A.tm<='" & pFromTm & "'"
    SQL = SQL & "   AND '" & pToTm & "' <= (SELECT B.tm FROM " & T_COM013 & " B "
    SQL = SQL & "                           WHERE" & DBW("B.empid=", pEmpId)
    SQL = SQL & "                               AND B.dt=A.dt"
    SQL = SQL & "                               AND B.flag='3')"
    
    RS.Open SQL, dbconn
    If Not (RS.BOF Or RS.EOF) Then
        GetNormalWorkCnt02 = True
    Else
        GetNormalWorkCnt02 = False
    End If
    
    Set RS = Nothing
    Exit Function
    
ErrorMsg:
    MsgBox Err.Description, vbOKOnly + vbCritical, "오류"
End Function

'*--------------------------------------------------------
'*  기능 : 검색기간안의 연차, 휴가가 있는지 존재
'*  Parameter
'*      pEmpId : 직원ID
'*      pFromDt : 시작일
'*      pToDt : 종료일
'*      pFlag : 연차(9), 휴가(10) Flag
'*  Return : 시작일, 종료일
'*--------------------------------------------------------
Public Function GetHoliInfo(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                            ByVal pFlag As String) As String
    Dim SQL As String
    
    SQL = "SELECT fromdt, todt FROM " & T_COM015 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId) & " "
    SQL = SQL & "   AND ((fromdt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    SQL = SQL & "           OR todt BETWEEN '" & pFromDt & "' AND '" & pToDt & "')"
    SQL = SQL & "           OR ('" & pFromDt & "' BETWEEN fromdt AND todt"
    SQL = SQL & "           OR '" & pToDt & "' BETWEEN fromdt AND todt))"
    SQL = SQL & "   AND" & DBW("flag=", pFlag) & " "
    SQL = SQL & "ORDER BY fromdt ASC"
    
    GetHoliInfo = SQL
End Function

'*--------------------------------------------------------
'*  기능 : 해당직원의 기간별 통계를 구함
'*  Parameter
'*      pEmpId : 직원ID
'*      pFromDt : 시작일
'*      pToDt : 종료일
'*  Return : 각 Flag별 통계
'*--------------------------------------------------------
Public Function GetStaticsEmp(ByVal pEmpId As String, ByVal pFromDt As String, _
                                            ByVal pToDt As String) As String
    Dim SQL1 As String
    Dim SQL2 As String
    
    SQL1 = "SELECT flag, Count(flag) as count "
    SQL1 = SQL1 & "FROM " & T_COM013 & " "
    SQL1 = SQL1 & "WHERE" & DBW("empid=", pEmpId)
    SQL1 = SQL1 & "     AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    SQL1 = SQL1 & "GROUP BY flag"
    
    SQL2 = "SELECT flag, Count(flag) as count "
    SQL2 = SQL2 & "FROM " & T_COM014 & " "
    SQL2 = SQL2 & "WHERE" & DBW("empid=", pEmpId)
    SQL2 = SQL2 & "     AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    SQL2 = SQL2 & "GROUP BY flag"
    
    SQL1 = SQL1 & " UNION ALL " & SQL2 & " ORDER BY flag ASC"
    GetStaticsEmp = SQL1
End Function

'*--------------------------------------------------------
'*  기능 : 전체직원의 기간별 통계를 구함
'*  Parameter
'*      pFromDt : 시작일
'*      pToDt : 종료일
'*  Return : 직원ID, 직원이름, 상태(출근, 지각, 조퇴 등등..)
'*--------------------------------------------------------
Public Function GetStaticsEmpAll(ByVal pFromDt As String, ByVal pToDt As String) As String
    Dim SQL1 As String
    Dim SQL2 As String
    
    SQL1 = "SELECT B.empid, B.empnm, A.flag, Count(A.flag) AS count "
    SQL1 = SQL1 & "FROM " & T_COM006 & " B, " & T_COM013 & " A "
    SQL1 = SQL1 & "WHERE dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL1 = SQL1 & "      AND A.empid=B.empid "
    SQL1 = SQL1 & "GROUP BY B.empid, B.empnm, A.flag"
    
    SQL2 = "SELECT B.empid, B.empnm, A.flag, Count(A.flag) AS count "
    SQL2 = SQL2 & "FROM " & T_COM006 & " B, " & T_COM014 & " A "
    SQL2 = SQL2 & "WHERE dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL2 = SQL2 & "      AND A.empid=B.empid "
    SQL2 = SQL2 & "GROUP BY B.empid, B.empnm, A.flag"
    
    SQL1 = SQL1 & " UNION ALL " & SQL2 & " ORDER BY empid, flag ASC"
    GetStaticsEmpAll = SQL1
End Function

'*--------------------------------------------------------
'*  기능 : 해당직원을 지각날짜를 구함
'*  Parameter
'*      pEmpId : 직원ID
'*      pFromDt : 시작일
'*      pToDt : 종료일
'*  Return : 지각일, 시간
'*--------------------------------------------------------
Public Function GetLateDt(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String) _
                            As String
    Dim SQL As String
    
    SQL = "SELECT dt, tm FROM " & T_COM013 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId)
    SQL = SQL & "       AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL = SQL & "       AND flag='2' "
    SQL = SQL & "ORDER BY dt ASC"
    
    GetLateDt = SQL
End Function

'*--------------------------------------------------------
'*  기능 : 해당직원을 조퇴, 반휴, 생휴, 월차, 결근날짜를 구함
'*  Parameter
'*      pEmpId : 직원ID
'*      pFromDt : 시작일
'*      pToDt : 종료일
'*      pFlag :조퇴(4), 반휴(5), 생휴(6), 월차(7), 결근(8) Flag
'*  Return : 조퇴, 반휴, 생휴, 월차, 결근일자
'*--------------------------------------------------------
Public Function GetDate(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                            ByVal pFlag As String) As String
    Dim SQL As String
    
    SQL = "SELECT dt FROM " & T_COM014 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId)
    SQL = SQL & "       AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL = SQL & "       AND" & DBW("flag=", pFlag)
    SQL = SQL & "ORDER BY dt ASC"
    
    GetDate = SQL
End Function

'*--------------------------------------------------------
'*  기능 : 해당기간의 공휴일 일수를 구함
'*  Parameter
'*      PFromDt : 시작일
'*      pToDt : 종료일
'*  Return : 공휴일 일수
'*--------------------------------------------------------
Public Function GetHolidayCount(ByVal pFromDt As String, ByVal pToDt As String) As Integer
    Dim RS As New Recordset
    Dim SQL As String
    
On Error GoTo ErrorMsg:
    SQL = "SELECT Count(holidt) AS count FROM " & T_COM017 & " "
    SQL = SQL & "WHERE holidt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    
    RS.Open SQL, dbconn
    If Not (RS.BOF Or RS.EOF) Then
        GetHolidayCount = CInt(RS.Fields("count").Value & "")
    End If
    
    Set RS = Nothing
    Exit Function
    
ErrorMsg:
    MsgBox Err.Description, vbOKOnly + vbCritical, "오류"
End Function
