Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\DIDIM_Interface\당진종합병원\" & REG_INSNAME

'공용테이블 장비 인덱스
Public Const IDX_STA        As String = "C202" 'LIMAS032 워크스테이션
Public Const IDX_SPC        As String = "C203" 'LIMAS032 검체
Public Const IDX_EQP        As String = "C209" 'LIMAS032 장비리스트
Public Const IDX_ROOM       As String = "C252" 'LIMAS032 검사실
Public Const IDX_SITE       As String = "C261" 'LIMAS032 사업장
Public Const IDX_TST        As String = "C604" 'LIMAS032 장비별 검사코드

'Visual Basic Color
Global Const vbLockColor = &HE0E0E0

'검사 타입
Public Const MSG_GEN As String = "G"        '일반
Public Const MSG_QCT As String = "Q"        'QC
Public Const MSG_ETC As String = "E"        '기타

'현재 사용자 정보
Public Const ELVELS_SUP  As String = "모든 권한"
Public Const ELVELS_RED  As String = "읽   기"
Public Const ELVELS_WRI  As String = "쓰   기"
Public Const ELVELS_RW   As String = "읽기,쓰기"
Public Const ELVELS_NOT  As String = "권한 없음"

Public Type UserInfo
    CuUserID    As String '사용자 ID
    CuUserNM    As String '사용자 이름
    CuUserPW    As String '사용자 비밀번호
    CuPower     As Authority  '사용자 권한
End Type

' 권 한
Public Enum Authority
    ELVEL_SUP = 1
    ELVEL_RED = 2
    ELVEL_WRI = 3
    ELVEL_RW = 4
    ELVEL_NOT = 5
End Enum

'현재 사용자 정보
Public CurrUser             As UserInfo
Public INS_CODE             As String       '장비코드
Public INS_NAME             As String       '장비명

Public DirPath              As String
Public MainForm             As MDIMain
Private TimerID             As Long


'-- 해당 티맥스 버전 : 3.1.4
'-- 티맥스 DLL  버전 : TMAX4GL.DLL
Public initcheck As Boolean

Public Function getSvrcInfo(ByVal strSrcNm As String, ByVal strArg1 As String, _
                            Optional ByVal strArg2 As String, Optional ByVal strArg3 As String, _
                            Optional ByVal strArg4 As String, Optional ByVal strArg5 As String, _
                            Optional ByVal strArg6 As String, Optional ByVal strArg7 As String, _
                            Optional ByVal strArg8 As String, Optional ByVal strArg9 As String) As String
    
    Dim iRet As Integer
    Dim strText1 As String, strText2 As String, strText3 As String, strText4 As String, strText5 As String
    Dim strText6 As String, strText7 As String, strText8 As String, strText9 As String, strText10 As String
    Dim ErrMsg As String
    Dim sendBuf As tuxbuf
    Dim rcvbuf As tuxbuf
    Dim rbuflen  As Long
    Dim strSvrNm As String
    Dim iRcvLen As Long
    Dim strGet(100) As String
    Dim intCnt As Integer
    Dim IntRow As Integer
    Dim intRecord As Integer
    
    Dim sendBuf1 As tuxbuf
    Dim rcvbuf1 As tuxbuf
    
    
    Dim Isendbuf As Long
    
    If initcheck = False Then
        MsgBox "TMAX연결이 실패한 상태입니다. 진행할 수 없습니다"
        Exit Function
    End If
    
    sendBuf.bufptr = tpalloc("FIELD", "", ByVal 1024&)
    rcvbuf.bufptr = tpalloc("FIELD", "", ByVal 1024&)
        
    If sendBuf.bufptr = 0 Or rcvbuf.bufptr = 0 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        Exit Function
    End If
    
    getSvrcInfo = ""
    strSvrNm = strSrcNm
    strText1 = strArg1
    strText2 = strArg2
    strText3 = strArg3
    strText4 = strArg4
    strText5 = strArg5
    strText6 = strArg6
    strText7 = strArg7
    strText8 = strArg8
    strText9 = strArg9
    
    Select Case strSvrNm
        '-- Dummy
        Case "CC_SYSDATE_S"
            IntRow = 1
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
        
        '-- 사용자 정보
        Case "SL_USERM_L1"
            IntRow = 2
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
            
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText1, 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
        
        '-- 접수정보 download (양방향용)
        Case "SL_INTFC_L2"
            IntRow = 15
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
            
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText1, 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
            
        '-- 접수 상태
        Case "SL_SPCMD_L12"
            IntRow = 2
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
            
            
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_FLAG1"), ByVal strText1, 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
            
            If strText2 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText2, 0) = -1 Then
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If
        
        Case "SL_RSLTUP_M1"
            IntRow = 1
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_STRING1"), ByVal strText1, 0) = -1 Then      '-- 장비코드
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                GoTo memoryfree
            End If
            
            If strText2 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText2, 0) = -1 Then   '-- 검체번호
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If
        
            If strText3 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_CODE1"), ByVal strText3, 0) = -1 Then    '-- 검사코드
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If
        
            If strText4 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_NO1"), ByVal strText4, 0) = -1 Then      '-- 처방순번
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If
        
            If strText5 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_TEXT1"), ByVal strText5, 0) = -1 Then    '-- 검사결과
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If
        
            If strText6 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("LOG_USERID"), ByVal strText6, 0) = -1 Then '-- 사번
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If
        
            If strText7 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("LOG_CLTNAME"), ByVal strText7, 0) = -1 Then '-- 로컬 IP
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX 에러"
                    GoTo memoryfree
                End If
            End If

    End Select
        
    iRet = tpcall(strSvrNm, ByVal sendBuf.bufptr, ByVal 0, rcvbuf.bufptr, iRcvLen, ByVal 0)
    
    If iRet = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        getSvrcInfo = ""
        GoTo memoryfree
    End If
    
    If strSvrNm = "CC_SYSDATE_S" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATETIME1"))
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                            
                iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATETIME1"), ByVal strGet(intCnt), 0)   '-- 성명
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1)
                End If
            Next
        Next
    
    ElseIf strSvrNm = "SL_USERM_L1" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"))
    
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                            
                If intCnt = 1 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"), ByVal strGet(intCnt), 0)   '-- 성명
                ElseIf intCnt = 2 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_TEXT1"), ByVal strGet(intCnt), 0)   '-- 비밀번호
                End If
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"
                End If
            Next
        Next

    
    ElseIf strSvrNm = "SL_INTFC_L2" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"))
    
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                                            
                Select Case intCnt
                Case 1
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"), ByVal strGet(intCnt), 0)       '-- 검사일자[yyyymmdd]
                Case 2
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NUMVAL2"), ByVal strGet(intCnt), 0)     '-- 수거번호
                Case 3
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NO1"), ByVal strGet(intCnt), 0)         '-- 처방순번
                Case 4
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE2"), ByVal strGet(intCnt), 0)       '-- 검사코드
                Case 5
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"), ByVal strGet(intCnt), 0)       '-- 검사명
                Case 6
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE3"), ByVal strGet(intCnt), 0)       '-- 검체코드
                Case 7
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME2"), ByVal strGet(intCnt), 0)       '-- 검체명
                Case 8
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_PATNO"), ByVal strGet(intCnt), 0)       '-- 등록번호
                Case 9
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_PATNAME"), ByVal strGet(intCnt), 0)     '-- 성명
                Case 10
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STRING1"), ByVal strGet(intCnt), 0)     '-- 성별
                Case 11
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STRING2"), ByVal strGet(intCnt), 0)     '-- 나이
                Case 12
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE4"), ByVal strGet(intCnt), 0)       '-- 진료과
                Case 13
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_TEXT1"), ByVal strGet(intCnt), 0)       '-- 직전결과
                Case 14
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_TEXT2"), ByVal strGet(intCnt), 0)       '-- 검사결과
                Case 15
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE2"), ByVal strGet(intCnt), 0)       '-- 직전결과보고일시[yyyymmdd hh24:mi:ss]
                End Select
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"
                End If
            Next
            getSvrcInfo = getSvrcInfo '& vbCrLf
        Next
    
    ElseIf strSvrNm = "SL_SPCMD_L12" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STAT1"))
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                
                If intCnt = 1 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE6"), ByVal strGet(intCnt), 0)
                    
                    If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                        getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"    '-- 검사코드
                    End If
                ElseIf intCnt = 2 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STAT1"), ByVal strGet(intCnt), 0)
                    
                    If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                        getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"    '-- 상태값 B:채취/접수,C:수거,E:검사,N:보고
                    End If
                End If
                
                
            Next
        Next

    ElseIf strSvrNm = "SL_RSLTUP_M1" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"))
    
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                                            
                Select Case intCnt
                Case 1
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"), ByVal strGet(intCnt), 0)       '--
                Case 2
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NUMVAL2"), ByVal strGet(intCnt), 0)
                Case 3
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NO1"), ByVal strGet(intCnt), 0)
                Case 4
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE2"), ByVal strGet(intCnt), 0)
                Case 5
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"), ByVal strGet(intCnt), 0)
                Case 6
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE3"), ByVal strGet(intCnt), 0)
                Case 7
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME2"), ByVal strGet(intCnt), 0)
                End Select
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"
                End If
            Next
            getSvrcInfo = getSvrcInfo & vbCrLf
        Next
    End If
    
'    Debug.Print getSvrcInfo

memoryfree:
    
    Call tpfree(ByVal sendBuf.bufptr)
    Call tpfree(ByVal rcvbuf.bufptr)
    
    Call tmaxexit

End Function



Sub Main()
    Dim ret     As Integer
    Dim lsndbuf As Long
    Dim tpinfo  As tpstart_t
    Dim strPtr  As Long
    Dim iRet    As Integer
    Dim lret    As Long
    Dim retVal  As String
    Dim strPath As String
    Dim strLbl  As String
    Dim ErrMsg  As String
    Dim strMsg As String
    Dim strSvrcCnt  As Variant
    Dim strSvrcData As Variant
    
    Dim lngConnect  As Long
    
    '두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "     Now Excute twice!", vbExclamation
       End
    End If
        
    'Registree Scan
    If Len(GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)) = 0 Then
        frmDB_JET.Show vbModal
    End If
    
    If Len(GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_SERVER)) = 0 Then
        frmDB_ORACLE.Show vbModal
    End If
    
    If Not DbConnect_Jet Then
        strMsg = "Local Batabase Not found! Do you want database search it? "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
            frmDB_JET.Show vbModal
        Else
            End
        End If
    End If
    
    If Not DbConnect_ORACLE Then
        strMsg = "Oracle Batabase Not found! Do you want database search it?   "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
            frmDB_ORACLE.Show vbModal
        Else
            End
        End If
    End If
    
    '실행 위치 저장
    DirPath = App.Path
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
        
    'Login Form 나타남
'    frmLogin.Show vbModal
    Set MainForm = New MDIMain
    MainForm.Show
    
End Sub

'Progressbar 설정
Public Sub SetProgress(ByVal lngMax As Long, ByVal CapStyle As CaptionStyles, ByVal strCaption As String, ByVal blnVisible As Boolean)
    'lngMax         : 최대값
    'CapStyle       : 켑션 스타일
    'strCaption     : 켑션
    'blnVisible     : 보임

    With MainForm.pgbMain
        .Max = lngMax
        .Visible = blnVisible
        .CaptionStyle = CapStyle
        .Caption = strCaption
        .Value = 0
    End With
End Sub

'Progressbar 값 설정
Public Sub ShowProgress(ByVal Values As Long, ByVal strCaption As String, ByVal blnVisible As Boolean)
    'Values         : 값
    'strCaption     : 켑션
    'blnVisible     : 나타남
    
    With MainForm.pgbMain
        .Visible = blnVisible
        .Caption = strCaption
        .Value = Values
    End With
End Sub

'상태 표시줄에 메시지 자동 지우기
Public Sub TimerProc(ByVal hwnd&, ByVal MSG&, ByVal ID&, ByVal nTime&)
    Call KillTimer(MainForm.hwnd, TimerID)
    With MainForm.stbMain
        .Panels("Output").text = ""
    End With
End Sub

'상태 표시줄에 메시지 나타내기
Public Sub ShowMessage(ByVal strMessage As String)
    'strMessage : 켑션
    
    Call KillTimer(MainForm.hwnd, TimerID)
    Call SetTimer(MainForm.hwnd, TimerID, 5000, AddressOf TimerProc)
    
    With MainForm
        With .pgbMain
            .Visible = False
        End With
        With .stbMain
            .Panels("Output").text = strMessage
        End With
    End With
    
End Sub

