VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISHospital05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarStat    As String
Private mvarSpcCd   As String

Public Property Let Stat(ByVal vData As String)
    mvarStat = vData
End Property

Public Property Let SpcCd(ByVal vData As String)
    mvarSpcCd = vData
End Property

'*--------------------------------------------------------
'*  기능 : 임상병리과의 WorkArea를 구함
'*  Return : 임상병리과 WorkArea 이름(한글명)
'*--------------------------------------------------------
Public Function LoadWorkArea() As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT field1, cdval1 FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", lc3_workarea)
    
    Set LoadWorkArea = New Recordset
    LoadWorkArea.Open strSQL, DBConn
    
'    If LoadWorkArea.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function
'*--------------------------------------------------------
'*  기능 : 임상병리과의 검체코드를 구함
'*  Return : 임상병리과 검체코드 이름(한글명)
'*--------------------------------------------------------
Public Function LoadSpcCd() As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT cdval1, field5 FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", LC3_Specimen) & " order by field5"
    
    Set LoadSpcCd = New Recordset
    LoadSpcCd.Open strSQL, DBConn
    
'    If LoadSpcCd.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function
'*--------------------------------------------------------
'*  기능 : 해당 WorkArea의 검사코드, 검사명을 구함
'*  Parameter
'*      pWA : WorkArea
'*  Return : 검사코드, 검사명
'*--------------------------------------------------------
Public Function LoadTestCdNm(ByVal pWA As String) As Recordset
    Dim SQL As String
    
    SQL = "SELECT testcd, abbrnm10, testnm FROM " & T_LAB001 & " WHERE" & DBW("workarea=", pWA)
    
    Set LoadTestCdNm = New Recordset
    LoadTestCdNm.Open SQL, DBConn
    
'    If LoadTestCdNm.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : 각 WorkArea별 접수건수를 구함
'*  Parameter : 접수일자
'*  Return : 각 WorkArea별 접수건수
'*--------------------------------------------------------
Public Function GetWorkAccount(RcvDt As String) As Recordset
    Dim SQL As String
    
    SQL = "SELECT b.field1, Count(*) AS Cnt FROM " & T_LAB032 & " b, " & T_LAB201 & " a "
    SQL = SQL & "WHERE" & DBW("a.rcvdt=", RcvDt) & " AND "
    SQL = SQL & "a.stscd >'1' AND a.stscd<'7' AND "
    SQL = SQL & DBW("b.cdindex=", lc3_workarea) & " AND "
    SQL = SQL & "a.workarea=b.cdval1 "
    SQL = SQL & "GROUP BY b.field1"
    
    Set GetWorkAccount = New Recordset
    GetWorkAccount.Open SQL, DBConn
    
'    If GetWorkAccount.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : 각 WorkArea별 접수내역 조회(조건시간 이후부터)
'*  Parameter :
'*      1. 접수/보고일
'*      2. WorkArea
'*      3. 조건시간
'*  Return : 접수일자, 보고일자, 각 WorkArea별, 조건시간 이후의 접수내역
'*--------------------------------------------------------
Public Function GetAccInfo01(ByVal pDate As String, ByVal pWorkArea As String, ByVal pTime As String) _
                                        As Recordset
    Dim SQL As String
    
    SQL = "SELECT a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm "
    SQL = SQL & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, " & T_LAB032 & " C "
    SQL = SQL & "WHERE A.ptid=B." & F_PTID & " AND "
    SQL = SQL & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
    SQL = SQL & DBW("A.workarea=", pWorkArea) & " AND" & DBW("A.rcvdt=", pDate)
    SQL = SQL & " AND " & DBW("A.stscd>", enStsCd.StsCd_LIS_Collection) & _
                " AND " & DBW("A.stscd<", enStsCd.StsCd_LIS_Cancel) & _
                " AND " & DBW("A.rcvtm >=", pTime)
        '응급 비응급체크
    If mvarStat <> "" Then
        SQL = SQL & " AND " & DBW("A.statfg=", mvarStat)
    End If
    
    '검체체크
    If mvarSpcCd <> "" Then
        SQL = SQL & " AND " & DBW("A.spccd=", mvarSpcCd)
    End If
    
    
    SQL = SQL & "ORDER BY A.rcvtm, A.rcvdt, A.accseq ASC"

    Set GetAccInfo01 = New Recordset
    GetAccInfo01.Open SQL, DBConn

'    If GetAccInfo01.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : 각 WorkArea별 접수내역 조회(조건시간 이후부터)
'*  Parameter :
'*      1. 접수/보고일
'*      2. WorkArea
'*      3. 조건시간
'*  Return : 접수일자, 보고일자, 각 WorkArea별, 조건시간 이후의 접수내역
'*--------------------------------------------------------
Public Function GetAccInfo02(ByVal pDate As String, ByVal pWorkArea As String, ByVal pTime As String) _
                                        As Recordset
    Dim SQL As String
    
    SQL = "SELECT a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm "
    SQL = SQL & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, " & T_LAB032 & " C "
    SQL = SQL & "WHERE A.ptid=B." & F_PTID & " AND "
    SQL = SQL & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
    SQL = SQL & DBW("A.workarea=", pWorkArea) & " AND" & DBW("A.vfydt=", pDate)
    SQL = SQL & " AND " & DBW("A.stscd>", enStsCd.StsCd_LIS_Collection) & _
                " AND " & DBW("A.stscd<", enStsCd.StsCd_LIS_Cancel) & _
                " AND " & DBW("A.rcvtm >=", pTime)
        '응급 비응급체크
    If mvarStat <> "" Then
        SQL = SQL & " AND " & DBW("A.statfg=", mvarStat)
    End If
    
    '검체체크
    If mvarSpcCd <> "" Then
        SQL = SQL & " AND " & DBW("A.spccd=", mvarSpcCd)
    End If
    
    
    SQL = SQL & "ORDER BY A.vfytm, A.vfydt, A.accseq ASC"

    Set GetAccInfo02 = New Recordset
    GetAccInfo02.Open SQL, DBConn

'    If GetAccInfo01.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function


'*--------------------------------------------------------
'*  기능 : 각 WorkArea별 접수내역 조회(조건시간 이후부터)
'*  Parameter :
'*      1. 접수/보고일
'*      2. 접수/보고 구분자(1:접수일자, 2:보고일자)
'*      3. WorkArea
'*      4. 조건시간
'*  Return : 접수일자, 보고일자, 각 WorkArea별, 조건시간 이후의 접수내역
'*--------------------------------------------------------
Public Function GetAccInfoByTime(qDate As String, qKind As Integer, qWorkArea As String, qTime As String) _
                                                As Recordset
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    
    If qKind = 1 Then   '접수일자
        '일반검사 검사코드
        SQL1 = "SELECT A.spcyy, A.spcno, a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm, D.testcd "
        SQL1 = SQL1 & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, "
        SQL1 = SQL1 & T_LAB032 & " C, " & T_LAB302 & " D "
        SQL1 = SQL1 & "WHERE A.ptid=B." & F_PTID & " AND "
        SQL1 = SQL1 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
        SQL1 = SQL1 & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.rcvdt=", qDate)
        SQL1 = SQL1 & " AND " & DBW("A.stscd>", enStsCd.StsCd_LIS_Collection) & _
                      " AND " & DBW("A.stscd<", enStsCd.StsCd_LIS_Cancel) & _
                      " AND " & DBW("A.rcvtm >=", qTime)
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL1 = SQL1 & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        '기타검사 검사코드
        SQL2 = "SELECT A.spcyy, A.spcno, a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm, D.testcd "
        SQL2 = SQL2 & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, "
        SQL2 = SQL2 & T_LAB032 & " C, " & T_LAB404 & " D "
        SQL2 = SQL2 & "WHERE A.ptid=B." & F_PTID & " AND "
        SQL2 = SQL2 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
        SQL2 = SQL2 & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.rcvdt=", qDate)
        SQL2 = SQL2 & " AND A.stscd > '1' AND A.stscd < '7' AND A.rcvtm >= '" & qTime & "'"
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL2 = SQL2 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL2 = SQL2 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL2 = SQL2 & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        '미생물검사 검사코드
        SQL3 = "SELECT A.spcyy, A.spcno, a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm, D.testcd "
        SQL3 = SQL3 & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, "
        SQL3 = SQL3 & T_LAB032 & " C, " & T_LAB404 & " D "
        SQL3 = SQL3 & "WHERE A.ptid=B." & F_PTID & " AND "
        SQL3 = SQL3 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
        SQL3 = SQL3 & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.rcvdt=", qDate)
        SQL3 = SQL3 & " AND A.stscd > '1' AND A.stscd < '7' AND A.rcvtm >= '" & qTime & "'"
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL3 = SQL3 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL3 = SQL3 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL3 = SQL3 & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        SQL1 = SQL1 & " UNION " & SQL2 & " UNION " & SQL3 & " ORDER BY rcvtm, accseq ASC"
    ElseIf qKind = 2 Then   '보고일자
        SQL1 = "SELECT a.workarea, a.accdt, A.accseq, A.stscd, A.ptid,B." & F_PTNM & " as name, C.field5, A.rcvtm "
        SQL1 = SQL1 & "FROM " & T_LAB032 & " C, " & T_HIS001 & " B, " & T_LAB201 & " A "
        SQL1 = SQL1 & " WHERE " & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.vfydt=", qDate) & " AND"
        SQL1 = SQL1 & " A.stscd > '1' AND A.stscd < '7' AND A.vfytm >= '" & qTime & "' "
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL1 = SQL1 & " AND A.ptid=B." & F_PTID & " AND "
        SQL1 = SQL1 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " "
        SQL1 = SQL1 & "ORDER BY A.rcvtm, A.rcvdt, A.accseq ASC"
    End If
    
    Set GetAccInfoByTime = New Recordset
    GetAccInfoByTime.Open SQL1, DBConn
    
'    If GetAccInfoByTime.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : 각 WorkArea별 접수내역 조회(조건시간 이후부터)
'*  Parameter :
'*      1. 접수/보고일
'*      2. 접수/보고 구분자(1:접수일자, 2:보고일자)
'*      3. WorkArea
'*      4. 조건시간
'*  Return : 접수일자, 보고일자, 각 WorkArea별, 조건시간 이후의 접수내역
'*--------------------------------------------------------
Public Function GetVfyInfoByTime(qDate As String, qKind As Integer, qWorkArea As String, qTime As String) _
                                                As Recordset
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    
    If qKind = 1 Then   '보고일자
        '일반검사 검사코드
        SQL1 = "SELECT A.spcyy, A.spcno, a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm, D.testcd "
        SQL1 = SQL1 & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, "
        SQL1 = SQL1 & T_LAB032 & " C, " & T_LAB302 & " D "
        SQL1 = SQL1 & "WHERE A.ptid=B." & F_PTID & " AND "
        SQL1 = SQL1 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
        SQL1 = SQL1 & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.vfydt=", qDate)
        SQL1 = SQL1 & " AND " & DBW("A.stscd>", enStsCd.StsCd_LIS_Collection) & _
                      " AND " & DBW("A.stscd<", enStsCd.StsCd_LIS_Cancel) & _
                      " AND " & DBW("A.vfytm >=", qTime)
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL1 = SQL1 & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        '기타검사 검사코드
        SQL2 = "SELECT A.spcyy, A.spcno, a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm, D.testcd "
        SQL2 = SQL2 & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, "
        SQL2 = SQL2 & T_LAB032 & " C, " & T_LAB404 & " D "
        SQL2 = SQL2 & "WHERE A.ptid=B." & F_PTID & " AND "
        SQL2 = SQL2 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
        SQL2 = SQL2 & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.vfydt=", qDate)
        SQL2 = SQL2 & " AND A.stscd > '1' AND A.stscd < '7' AND A.vfytm >= '" & qTime & "'"
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL2 = SQL2 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL2 = SQL2 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL2 = SQL2 & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        '미생물검사 검사코드
        SQL3 = "SELECT A.spcyy, A.spcno, a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvtm, D.testcd "
        SQL3 = SQL3 & "FROM " & T_LAB201 & " A, " & T_HIS001 & " B, "
        SQL3 = SQL3 & T_LAB032 & " C, " & T_LAB404 & " D "
        SQL3 = SQL3 & "WHERE A.ptid=B." & F_PTID & " AND "
        SQL3 = SQL3 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " AND"
        SQL3 = SQL3 & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.vfydt=", qDate)
        SQL3 = SQL3 & " AND A.stscd > '1' AND A.stscd < '7' AND A.vfytm >= '" & qTime & "'"
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL3 = SQL3 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL3 = SQL3 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL3 = SQL3 & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        SQL1 = SQL1 & " UNION " & SQL2 & " UNION " & SQL3 & " ORDER BY rcvtm, accseq ASC"
    ElseIf qKind = 2 Then   '보고일자
        SQL1 = "SELECT a.workarea, a.accdt, A.accseq, A.stscd, A.ptid,B." & F_PTNM & " as name, C.field5, A.rcvtm "
        SQL1 = SQL1 & "FROM " & T_LAB032 & " C, " & T_HIS001 & " B, " & T_LAB201 & " A "
        SQL1 = SQL1 & " WHERE " & DBW("A.workarea=", qWorkArea) & " AND" & DBW("A.vfydt=", qDate) & " AND"
        SQL1 = SQL1 & " A.stscd > '1' AND A.stscd < '7' AND A.vfytm >= '" & qTime & "' "
        '응급 비응급체크
        If mvarStat <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.statfg=", mvarStat)
        End If
        
        '검체체크
        If mvarSpcCd <> "" Then
            SQL1 = SQL1 & " AND " & DBW("A.spccd=", mvarSpcCd)
        End If
        
        SQL1 = SQL1 & " AND A.ptid=B." & F_PTID & " AND "
        SQL1 = SQL1 & "A.spccd=C.cdval1 AND" & DBW("C.cdindex=", LC3_Specimen) & " "
        SQL1 = SQL1 & "ORDER BY A.rcvtm, A.rcvdt, A.accseq ASC"
    End If
    
    Set GetVfyInfoByTime = New Recordset
    GetVfyInfoByTime.Open SQL1, DBConn
    
'    If GetVfyInfoByTime.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : 각 WorkArea별 접수내역 조회(날자범위에 의한)
'*  Parameter :
'*      1. 접수/보고일
'*      2. 접수/보고 구분자(1:접수일자, 2:보고일자)
'*      3. 날자범위
'*  Return : 접수일자, 보고일자, 날자범위내 접수내역
'*--------------------------------------------------------
Public Function GetAccInfoByDate(qPtID As String, qFDate As String, qTDate As String, qKind As Integer, _
                                    qWorkArea As String, qTestCd As String) As Recordset
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    
    If qKind = 1 Then   '접수일자
        '일반검사 검사코드
        SQL1 = "SELECT A.spcyy, A.spcno, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvdt, D.testcd "
        SQL1 = SQL1 & "FROM " & T_LAB032 & " C, " & T_LAB302 & " D, " & T_HIS001 & " B, " & T_LAB201 & " A "
        SQL1 = SQL1 & "WHERE " & DBW("A.workarea=", qWorkArea) & " AND " & DBW("a.ptid =", qPtID)
        SQL1 = SQL1 & " AND A.rcvdt between '" & qFDate & "' AND '" & qTDate & "'"
        SQL1 = SQL1 & " AND " & DBW("C.cdindex=", LC3_Specimen) & " AND A.stscd in('4', '5', '6')"
        SQL1 = SQL1 & " AND " & DBW("d.testcd =", qTestCd) & " AND A.spccd=C.cdval1 "
        SQL1 = SQL1 & " AND A.ptid=B." & F_PTID & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "

        '기타검사 검사코드
        SQL2 = "SELECT A.spcyy, A.spcno, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvdt, D.testcd "
        SQL2 = SQL2 & "FROM " & T_LAB032 & " C, " & T_LAB404 & " D, " & T_HIS001 & " B, " & T_LAB201 & " A "
        SQL2 = SQL2 & "WHERE " & DBW("A.workarea=", qWorkArea) & " AND " & DBW("a.ptid =", qPtID)
        SQL2 = SQL2 & " AND A.rcvdt between '" & qFDate & "' AND '" & qTDate & "'"
        SQL2 = SQL2 & " AND " & DBW("C.cdindex=", LC3_Specimen) & " AND A.stscd in('4', '5', '6')"
        SQL2 = SQL2 & " AND " & DBW("d.testcd =", qTestCd) & " AND A.spccd=C.cdval1 "
        SQL2 = SQL2 & " AND A.ptid=B." & F_PTID & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "
        
        '미생물검사 검사코드
        SQL3 = "SELECT A.spcyy, A.spcno, a.accdt, A.accseq, A.stscd, A.ptid, B." & F_PTNM & " as name, C.field5, A.rcvdt, D.testcd "
        SQL3 = SQL3 & "FROM " & T_LAB032 & " C, " & T_LAB404 & " D, " & T_HIS001 & " B, " & T_LAB201 & " A "
        SQL3 = SQL3 & "WHERE " & DBW("A.workarea=", qWorkArea) & " AND " & DBW("a.ptid =", qPtID)
        SQL3 = SQL3 & " AND A.rcvdt between '" & qFDate & "' AND '" & qTDate & "'"
        SQL3 = SQL3 & " AND " & DBW("C.cdindex=", LC3_Specimen) & " AND A.stscd in('4', '5', '6')"
        SQL3 = SQL3 & " AND " & DBW("d.testcd =", qTestCd) & " AND A.spccd=C.cdval1 "
        SQL3 = SQL3 & " AND A.ptid=B." & F_PTID & " AND A.workarea=D.workarea AND A.accdt=D.accdt AND A.accseq=D.accseq "
        
        SQL1 = SQL1 & " UNION " & SQL2 & " UNION " & SQL3 & " ORDER BY rcvdt desc"
    ElseIf qKind = 2 Then   '보고일자
        SQL1 = "SELECT a.accdt, A.accseq, A.stscd, A.ptid,B." & F_PTNM & " as name, C.field5, A.vfydt "
        SQL1 = SQL1 & "FROM " & T_LAB032 & " C, " & T_HIS001 & " B, " & T_LAB201 & " A "
        SQL1 = SQL1 & "WHERE " & DBW("C.cdindex=", LC3_Specimen)
        SQL1 = SQL1 & " AND " & DBW("A.workarea=", qWorkArea) & " AND " & DBW("a.ptid =", qPtID)
        SQL1 = SQL1 & " AND a.vfydt between '" & qFDate & "' AND '" & qTDate & "'"
        SQL1 = SQL1 & " AND stscd in('2', '3', '4', '5', '6')" & " AND A.spccd=C.cdval1"
        SQL1 = SQL1 & " AND A.ptid=B." & F_PTID
        SQL1 = SQL1 & " ORDER BY a.vfydt desc"
    End If
    
    Set GetAccInfoByDate = New Recordset
    GetAccInfoByDate.Open SQL1, DBConn
    
'    If GetAccInfoByDate.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : WA별 접수내역 조회
'*  Parameter
'*      qDate : 접수일
'*      qWorkArea : WorkArea
'*  Return : 조회결과
'*--------------------------------------------------------
Public Function GetAccList(ByVal qDate As String, ByVal qTime As String, ByVal qWorkArea As String) _
                                        As Recordset
    Dim SQL As String

    SQL = "SELECT a.workarea, a.accdt, A.accseq, A.stscd, A.ptid, A.spcyy, A.spcno, A.rcvtm, B." & F_PTNM & " as name, C.field5, D.workcd "
    SQL = SQL & "FROM " & T_LAB301 & " D, " & T_LAB032 & " C, " & T_HIS001 & " B, " & T_LAB201 & " A "
    SQL = SQL & "WHERE" & DBW("D.workdt=", qDate) & " AND"
    SQL = SQL & DBW("d.workcd=", qWorkArea) & " AND "
    SQL = SQL & DBW("C.cdindex=", LC3_Specimen) & " AND "
    SQL = SQL & "A.rcvtm >= '" & qTime & "' AND  "
    SQL = SQL & DBW("A.stscd<>", enStsCd.StsCd_LIS_Cancel) & " AND "
    
    '응급 비응급체크
    If mvarStat <> "" Then
        SQL = SQL & DBW("a.statfg=", mvarStat) & " AND "
    End If
    
    '검체체크
    If mvarSpcCd <> "" Then
        SQL = SQL & DBW("a.spccd=", mvarSpcCd) & " AND "
    End If
    
    
    SQL = SQL & "A.ptid=B." & F_PTID & " AND "
    SQL = SQL & "A.spccd=C.cdval1 AND "
    SQL = SQL & "A.workarea=D.workarea" & " AND  "
    SQL = SQL & "A.accdt=D.accdt" & " AND "
    SQL = SQL & "A.accseq=D.accseq"
    SQL = SQL & " ORDER BY A.rcvtm, A.rcvdt, A.accseq ASC"
    
    Set GetAccList = New Recordset
    GetAccList.Open SQL, DBConn
    
'    If GetAccList.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'        Exit Function
'    End If
End Function

'*--------------------------------------------------------
'*  기능 : 결과조회
'*  Parameter :
'*      1. 접수/보고일
'*      2. 접수/보고 구분(rcvdt:접수일자, examdt:보고일자)
'*      3. WorkArea
'*      4. 환자ID
'*      5. WA별 접수번호
'*  Return : 검사결과
'*--------------------------------------------------------
Public Function GetResultQueryNew(ByVal qDate As String, ByVal qKind As String, ByVal qWorkArea As String, _
                                    ByVal qPtID As String, ByVal qaccseq As Long) As String

    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim tmpNotice As String
    Dim tmpStr As String
    Dim tmpOrderby As String
    Dim strSQL As String
    
    tmpNotice = " e.text1 notice, "
    tmpStr = "  " & FUNC_SUBSTR & "(" & qKind & ",3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(" & qKind & ",5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(" & qKind & ",7,2) as KeyDate "
    tmpOrderby = " Order by KeyDate desc, workarea, accdt, accseq, RptSeq, ordcd "
    
    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW(qKind & "=", qDate) & _
              " AND " & DBW("a.workarea=", qWorkArea) & " AND " & DBW("c.accseq=", qaccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")

    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW(qKind & "=", qDate) & _
              " AND " & DBW("x.workarea=", qWorkArea) & " AND " & DBW("x.accseq=", qaccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd")

    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW(qKind & "=", qDate) & _
              " AND " & DBW("x.workarea=", qWorkArea) & " AND " & DBW("x.accseq=", qaccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")

    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW(qKind & "=", qDate) & _
              " AND " & DBW("x.workarea=", qWorkArea) & " AND " & DBW("x.accseq=", qaccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")
    
    strSQL = result1 & " UNION ALL " & result2 & " UNION ALL " & result3
    strSQL = strSQL & tmpOrderby
    
    GetResultQueryNew = strSQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당접수번호의 결과내역 검색 - 5.0.4: 이상대(2005-07-12)
'          GetResultQueryNew 함수에서 접수일, 보고일 조건삭제
'          WorkSheet으로 조회할경우만 사용
'   인수 :
'       - pKind     : 접수일
'       - pPtId     : 환자ID
'       - pWorkarea : Workarea
'       - pAccSeq   : AccSeq
'-----------------------------------------------------------------------------'
Public Function GetResultQueryNewX(ByVal pKind As String, ByVal qPtID As String, _
                                   ByVal pWorkArea As String, ByVal paccseq As Long) As String

    Dim result1     As String
    Dim result2     As String
    Dim result3     As String
    Dim tmpStr      As String
    Dim tmpNotice   As String
    Dim tmpOrderby  As String
    Dim strSQL      As String
    
    tmpNotice = " e.text1 notice, "
    tmpStr = "  " & FUNC_SUBSTR & "(" & pKind & ",3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(" & pKind & ",5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(" & pKind & ",7,2) as KeyDate "
    tmpOrderby = " Order by KeyDate desc, workarea, accdt, accseq, RptSeq, ordcd "
    
    result1 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, a.rstunit, " & _
              "        a.hldiv, a.dpdiv, a.spccd, a.rstdiv, a.statfg, a.lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm as VfyTm, a.vfyid as VfyId, '0' as testdiv, c.footnotefg, c.rmkcd, " & _
              "        a.mfyfg, a.txtfg, a.rsttype, a.detailfg, a.eqpcd, j.field3 as SpcNm, '' as SenFg, " & _
              "        x.dcfg, x.mesg, a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid, k.roomid, " & _
              "        b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, " & tmpNotice & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & T_LAB034 & " e, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB302 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW("a.workarea=", pWorkArea) & _
              " AND " & DBW("c.accseq=", paccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & " AND " & DBJ("d.cdval1 =* a.testcd") & "  AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "   AND " & DBJ("j.cdval1 =* a.spccd") & _
              " AND  " & DBJ(DBW("e.cdindex = ", LC4_ClinicalNotice)) & "   AND " & DBJ("e.cdval1 =* a.testcd")

    result2 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, '' as rstcd, " & _
              "        '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, '' as rstdiv, x.statfg, '' as lastrst, " & _
                       DBV("lastvfyid", "0") & " as lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '1' as testdiv, " & _
              "        c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, a.txtfg, a.rsttype, '' as detailfg, " & _
              "        '' as eqpcd, j.field3 as SpcNm, '' as SenFg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
              "        '' as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, '' as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB201 & " c, " & T_LAB101 & " k, " & T_LAB351 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW("x.workarea=", pWorkArea) & _
              " AND " & DBW("x.accseq=", paccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "  AND " & DBJ("j.cdval1 =* c.spccd")

    result3 = " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, " & _
              "        a.lastvfyid, a.vfydt as VfyDt, a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & _
                       FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, b.detailfg as DetailItem, " & _
              "        c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW("x.workarea=", pWorkArea) & _
              " AND " & DBW("x.accseq=", paccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype = '" & MRT_Stain & "' " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt  AND  c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")

    result3 = result3 & " UNION ALL " & _
              " SELECT a.workarea as WorkArea, a.accdt as AccDt, a.accseq as AccSeq, a.testcd, a.rstcd, '' as rstunit, " & _
              "        '' as hldiv, '' as dpdiv, c.spccd, a.rstdiv, x.statfg, '' as lastrst, a.lastvfyid, a.vfydt as VfyDt, " & _
              "        a.vfytm VfyTm, a.vfyid VfyId, '2' as testdiv, c.footnotefg, c.rmkcd, " & FUNC_CONVERT("char", "a.mfyseq") & " as mfyfg, " & _
              "        '' as txtfg, a.rsttype, a.detailfg, '' as eqpcd, j.field3 as SpcNm, a.senfg, x.dcfg, x.mesg, " & _
              "        a.ptid, a.orddt, a.ordno, a.ordseq, x.ordcd, k.orddiv, k.wardid, k.hosilid,  k.roomid, b.testnm as TestLongNm, b.abbrnm10 as TestShortNm, " & _
              "        b.detailfg as DetailItem, c.deptcd, c.sex, c.ageday, k.orddoct, c.colid, c.rcvid, null notice, " & _
                       FUNC_SUBSTR & "(k.orddt ,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.orddt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(k.ordtm,3,2) as OrdDate, " & _
                       FUNC_SUBSTR & "(a.lastvfydt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfydt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(a.lastvfytm,3,2) as LstVfyDtTm, " & _
                       FUNC_SUBSTR & "(c.coldt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coldt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.coltm,3,2) as ColDtTm, " & _
                       FUNC_SUBSTR & "(c.rcvdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(c.rcvtm,3,2) as RcvDtTm, " & _
                       FUNC_SUBSTR & "(x.examdt,3,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,5,2)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examdt,7,2)" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,1,2)" & FUNC_CONCAT & "':'" & FUNC_CONCAT & "" & FUNC_SUBSTR & "(x.examtm,3,2) as VfyDtTm, " & _
              "        0 unitqty, k.reqdt, k.reqtm, d.field1 as RstCdNm, " & FUNC_CONVERT("num", "b.rptseq") & " as RptSeq, " & tmpStr & _
              " FROM " & T_LAB031 & " d, " & T_LAB032 & " j, " & T_LAB001 & " b, " & _
                         T_LAB101 & " k, " & T_LAB201 & " c, " & T_LAB404 & " a, " & T_LAB102 & " x " & _
              " WHERE  " & DBW("x.ptid = ", qPtID) & _
              " AND " & DBW("x.workarea=", pWorkArea) & _
              " AND " & DBW("x.accseq=", paccseq) & _
              " AND    a.ptid = x.ptid  AND  a.orddt = x.orddt  AND  a.ordno = x.ordno  AND  a.ordseq = x.ordseq  AND  a.rsttype <> '" & MRT_Stain & "' " & _
              " AND    k.ptid = x.ptid  AND  k.orddt = x.orddt  AND  k.ordno = x.ordno " & _
              " AND    b.testcd = a.testcd  AND  b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd AND applydt <= a.orddt ) " & _
              " AND    c.workarea = a.workarea  AND  c.accdt = a.accdt    AND    c.accseq = a.accseq " & _
              " AND  " & DBJ(DBW("d.cdindex = ", LC2_ItemResult)) & "  AND " & DBJ("d.cdval1 =* a.testcd") & " AND " & DBJ("d.cdval2 =* a.rstcd") & _
              " AND  " & DBJ(DBW("j.cdindex = ", LC3_Specimen)) & "    AND " & DBJ("j.cdval1 =* c.spccd")
    
    strSQL = result1 & " UNION ALL " & result2 & " UNION ALL " & result3
    strSQL = strSQL & tmpOrderby
    
    GetResultQueryNewX = strSQL
End Function
