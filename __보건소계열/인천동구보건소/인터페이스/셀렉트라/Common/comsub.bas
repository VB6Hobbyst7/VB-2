Attribute VB_Name = "comSUB"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  S_COMSUB = 공통함수 선언 Library                            *
'*                                                              *
'*  Designed  :                                                 *
'*  Coded     :                                                 *
'*  Modified  :                                                 *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *

Option Explicit
'*****************************************
'*** [Module List ]
'*****************************************
'  D0SUB_BIRTHDAY (ByVal PassPort_Id As String) As String
'  D0SUB_CDNAME_GET(length As Integer) As string
'  D0SUB_CITYNO_CHECK(para as string) as integer
'  D0SUB_DELETEFILE (DirName As String, SaveDate As String)
'  D0SUB_EXIST_FILE (para As String) As Integer
'  D0SUB_Exist_RECORD (para As String) As Integer
'  D0SUB_GRIDGETROW (grd As Grid, Col As Integer, para As String) As Integer
'  D0SUB_GRIDTOPROW (grd As Grid, Row As Integer, disRow As Integer)
'  D0SUB_LSPACE (src1 As String, len as integer) As string
'  D0SUB_MASK_DISP (Dest_Ctl As Control, w_data As String)
'  D0SUB_MASK_DISP(ctrl As control, src1 As string)
'  D0SUB_MESSAGE (ByVal para As Integer) As String
'  D0SUB_NULL_CHECK (para As Variant) As String
'  D0SUB_POSITION (Frm As Form, xpos As Long, YPos As Long)
'  D0SUB_RSPACE (src1 As String, len as integer) As string
'  D0SUB_SPREAD_CLEAR (spd As vaSpread, DispRow As String)
'  D0SUB_SPREADGETCOL (spd As vaSpread, Row As Integer, para As String) As Integer
'  D0SUB_SPREADGETROW (spd As vaSpread, Col As Integer, para As String) As Integer
'  D0SUB_SPREADHIGHLIGHT (spd As vaSpread, Row As Integer, OldRow As Integer) As Integer
'  D0SUB_SPREADTOPROW (spd As vaSpread, Row As Integer, disRow As Integer)
'  D0SUB_SYSTEMDATE(frm as form, optional Sqlconn as variant)

'*-------------------------------------------------*
'* req      : "1" ==> yyyymmdd return              *
'*          : "2" ==> yyyy-mm-dd return            *
'* para     : value                    (i)         *
'*-------------------------------------------------*
Function D0SUB_DATE_6TO8(ByVal para As String, ByVal req As String) As String

    Dim temp As String

    temp = Left(para, 2)                    '년도 저장
    If temp > "70" And temp <= "99" Then
        temp = "19" & para
    Else
        temp = "20" & para
    End If

    If req = "1" Then
        D0SUB_DATE_6TO8 = temp
    Else
        D0SUB_DATE_6TO8 = Format$(temp, "####-##-##")
    End If

End Function

'********************************************************
'*                                                      *
'*  주민등록번호 Check                                  *
'*  para  : 주민등록번호(9703211048211)                 *
'*                                                      *
'********************************************************
Function D0SUB_CITYNO_CHECK(ByVal para As Variant) As Integer

    Dim cValue  As Integer, eValue  As Integer
    Dim idx As Integer
    
    If Len(para) = 14 Then _
        para = Mid$(para, 1, 6) + Mid$(para, 8, 7)
    
    cValue = Val(Mid$(para, 1, 1)) * 2 _
           + Val(Mid$(para, 2, 1)) * 3 _
           + Val(Mid$(para, 3, 1)) * 4 _
           + Val(Mid$(para, 4, 1)) * 5 _
           + Val(Mid$(para, 5, 1)) * 6 _
           + Val(Mid$(para, 6, 1)) * 7 _
           + Val(Mid$(para, 7, 1)) * 8 _
           + Val(Mid$(para, 8, 1)) * 9 _
           + Val(Mid$(para, 9, 1)) * 2 _
           + Val(Mid$(para, 10, 1)) * 3 _
           + Val(Mid$(para, 11, 1)) * 4 _
           + Val(Mid$(para, 12, 1)) * 5
           
    eValue = 11 - (cValue Mod 11)
    
    If eValue > 9 Then eValue = eValue Mod 10
    
    D0SUB_CITYNO_CHECK = True
    
    If Not eValue = Val(Mid$(para, 13, 1)) Then D0SUB_CITYNO_CHECK = False
    
    If Mid$(para, 7, 1) <> "1" And Mid$(para, 7, 1) <> "2" Then _
       D0SUB_CITYNO_CHECK = False
    
    If IsDate(Format(Mid$(para, 1, 6), "##-##-##")) = False Then _
       D0SUB_CITYNO_CHECK = False
    
End Function


'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                              *
'*  vaSpread에서 Mult Select Color Setting..                    *
'*                                                              *
'*  spd          : vaSpread Name                                *
'*  chk          : chk= tru(선택), false(비선택)                *
'*  row          : Row                                          *
'*  bacC         : Spread의 바탕색                              *
'*                                                              *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Sub D0SUB_SpreadMultiSelect(ByVal spd As vaSpread, ByVal chk As Integer _
                        , ByVal Row As Long _
                        , Optional bacC As Variant)

    If Row < 1 Then Exit Sub
    
    If chk = False Then
        spd.Col = 1: spd.col2 = spd.MaxCols
        spd.Row = Row: spd.row2 = Row
        
        spd.BlockMode = True
    
        If IsMissing(bacC) Then
            spd.BackColor = &H80000005      '흰색
        Else
            spd.BackColor = bacC
        End If
    
        spd.BlockMode = False
    Else
        spd.Col = 1: spd.col2 = spd.MaxCols
        spd.Row = Row: spd.row2 = Row
        
        spd.BlockMode = True
    
        spd.BackColor = &HC0FFC0   '
    
        spd.BlockMode = False
    
    End If
    
End Sub


'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                              *
'*  해당 vaSpread에서 지정행에 HighLight한다.                *
'*                                                              *
'*  spd          : vaSpread Name                             *
'*  row          : Row Value                                    *
'*  oldrow       : 전의 Row Value                               *
'*  Return Value : HighList Row Value                           *
'*                                                              *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Function D0SUB_SPREAD_HIGHLIGHT(spd As vaSpread, Row As Integer _
                            , OldRow As Integer, Optional bacC As Variant) As Integer
    
    If Not OldRow < 1 Then
        spd.Col = 1: spd.col2 = spd.MaxCols
        spd.Row = OldRow: spd.row2 = OldRow
        
        spd.BlockMode = True
    
        If IsMissing(bacC) Then
            spd.BackColor = &H80000005      '흰색
        Else
            spd.BackColor = bacC
        End If
    
        spd.BlockMode = False
    End If
    
'/*--------------------------------------
    If Not Row < 0 Then
        spd.Col = 1: spd.col2 = spd.MaxCols
        spd.Row = Row: spd.row2 = Row
        
        spd.BlockMode = True
    
        spd.BackColor = &HC0FFC0   '
    
        spd.BlockMode = False
    
        D0SUB_SPREAD_HIGHLIGHT = Row
    End If
    
End Function


'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*   생년월일로 나이를 계산                      *
'*   passport_id   :  생년월일 변환대상 data     *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function D0SUB_BIRTHDAY(ByVal PassPort_Id As String) As String

    Dim cDte     As String
    Dim yy       As String
    Dim age      As Integer

    On Error GoTo D0SUB_BIRTHDAY

    Select Case Len(Left$(PassPort_Id, 6))
        Case 2, 3, 4, 5
            yy = Left$(PassPort_Id, 2) & "-01-01"
            age = DateDiff("yyyy", yy, Now)

        Case 6
            cDte = Format$(Now, "yyyymmdd")
            
            If Val(Right$(cDte, 6)) <= Val(PassPort_Id) Then
                yy = Trim(Val(Left$(cDte, 2)) - 1)
                yy = yy & Format$(PassPort_Id, "0#-##-##")
            Else
                yy = Left$(cDte, 2) & Format$(PassPort_Id, "0#-##-##")
            End If
            
            age = DateDiff("yyyy", yy, Now)
    End Select
        
    D0SUB_BIRTHDAY = Trim(Str$(age))
        
    On Error GoTo 0
    Exit Function
D0SUB_BIRTHDAY:

    Resume Next
        
End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* Table 로부터 코드명칭을 SELECT 하여           *
'*                       명칭 Control 에 표시    *
'*                                               *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub D0SUB_CDNAME_GET(frm As Form, Optional SqlConn As Variant)

    Dim SqlCode As Long
    Dim sStr  As String, ret As Integer
    Dim SData() As String, record As String

    If IsMissing(SqlConn) Then
        If QSqlOpen(D0COM_SERVER01, frm.hWnd, SqlCode) <> QSQL_SUCCESS _
            Then Exit Sub
    Else
        SqlCode = SqlConn
    End If

    sStr = "SELECT " & D0COM_name_col & " FROM " & D0COM_table
    sStr = sStr & " WHERE " & D0COM_code_col & " = '" & Trim(D0COM_code) & "' "
    If D0COM_cd_gbn <> "" Then
        sStr = sStr + " AND " + D0COM_cd_gbn
    End If

    ret = QSqlDBExec(sStr, SqlCode)
    If ret = QSQL_SUCCESS Then
        ret = QSqlGetRow(record, SqlCode)
        If ret = QSQL_SUCCESS Then
            QSqlGetField 1, record, SData()
    
            D0COM_name = D0SUB_RSPACE(SData(1), D0COM_length)
        Else
            D0COM_name = ""
        End If
    End If
    
    D0COM_ret = ret
    Call QSqlSelectFree(SqlCode)

    If IsMissing(SqlConn) Then Call Qsqlclose(SqlCode, ONECLOSE)

End Sub


'********************************************************
'*                                                      *
'*  저장기간이 지난 interface file를 지운다.            *
'*  DirName  : 지우고자 하는 Directory Name             *
'*  SaveDate : 저장기간                                 *
'*                                                      *
'********************************************************
Sub D0SUB_DELETE_FILE(DirName As String, SaveDate As String)

    Dim DelFile As String
    Dim FileName    As String
    Dim FileDate    As String

    Const ATTR_NORMAL = 0

    On Error GoTo D0SUB_DELETEFILE

    DelFile = Format$(Now - Val(SaveDate), "YY/MM/DD")

    DirName = Trim$(DirName)
    If Right$(DirName, 1) <> "\" Then DirName = DirName + "\"

    FileName = Dir(DirName, ATTR_NORMAL)
    Do While Len(FileName) <> 0
        FileDate = Left$(FileDateTime(DirName + FileName), 8)
    
        If CVDate(DelFile) >= CVDate(FileDate) Then
            Kill DirName + FileName
        End If
        
        FileName = Dir
    Loop
    
    On Error GoTo 0
    Exit Sub
D0SUB_DELETEFILE:
    
    Resume Next

End Sub

'*  *   *   *   *   *   *   *   *   *   *   *
'*                                          *
'*  파일의 존재여부을 파악한다.             *
'*  para : 파일명(경로명 포함)              *
'*  Return Value : true, false              *
'*                                          *
'*  *   *   *   *   *   *   *   *   *   *   *
Function D0SUB_EXIST_FILE(para As String) As Integer

    If Dir$(para) <> "" Then
        D0SUB_EXIST_FILE = True
    Else
        D0SUB_EXIST_FILE = False
    End If

End Function

'*------------------------------------------------------*
'*                                                      *
'*  Record의 존재하면 True, 아니면 False                *
'*  para : SQL 문                                       *
'*  SqlConn : SqlConn missing 이면 서버를 Open하여 사용 *
'*            D0COM_server = 서버                         *                                                      *
'*                                                      *
'*------------------------------------------------------*
Function D0SUB_EXIST_RECORD(frm As Form, para As String, Optional SqlConn As Variant) As Integer
    
    Dim SqlCode As Long, status  As Integer
    Dim tData() As String, recode    As String

    D0SUB_EXIST_RECORD = False
    
    If IsMissing(SqlConn) Then
        If QSqlOpen(D0COM_SERVER01, frm.hWnd, SqlCode) <> QSQL_SUCCESS _
            Then Exit Function
    Else
        SqlCode = SqlConn
    End If
        
    status = QSqlDBExec(para, SqlCode)
    If status = QSQL_SUCCESS Then
        status = QSqlGetRow(recode, SqlCode)
        If status = QSQL_SUCCESS Then
        
            QSqlGetField 1, recode, tData()

            If Val(tData(1)) <> 0 Then D0SUB_EXIST_RECORD = True
        
        End If
    End If
    Call QSqlSelectFree(SqlCode)

    If IsMissing(SqlConn) Then Call Qsqlclose(SqlCode, ONECLOSE)

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                              *
'*  grid에서 특정열의 데이타을 찾고 위치을 돌려줌..             *
'*                                                              *
'*  grd          : grid Name                                    *
'*  Col          : Column                                       *
'*  para         : Data                                         *
'*  Return Value : 데이타의 위치 행                             *
'*                 찾고자 하는 데이타가 없을 경우 -1를 SETTING  *
'*                                                              *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Function D0SUB_GRIDGETROW(grd As Object, Col As Integer, para As String) As Integer
    
    Dim code  As String

    Dim Row%

    For Row = 1 To grd.Rows - 1
        grd.Row = Row
        grd.Col = Col
        code = grd.Text

        If Trim$(code) = Trim$(para) Then
            D0SUB_GRIDGETROW = Row
            
            grd.HighLight = True                                            '입력라인 반전
            grd.SelStartRow = Row: grd.SelEndRow = Row
            grd.SelStartCol = grd.FixedCols: grd.SelEndCol = grd.Cols - 1
            
            Exit Function
        End If
    Next

    D0SUB_GRIDGETROW = -1

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                      *
'*  grid에서 특정행를 화면에 보여준다..                 *
'*  grd : grid Name                                     *
'*  Row : Row                                           *
'*  disRow : 화면에 보여줄 수 있는 최대 행의 수         *
'*                                                      *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
Sub D0SUB_GRIDTOPROW(grd As Object, Row As Integer, disRow As Integer)
    Dim CurrRow%

    CurrRow = Row Mod disRow

    On Error GoTo GridTopRowErr

    grd.TopRow = Row - CurrRow

    On Error GoTo 0
    
    Exit Sub

GridTopRowErr:
    If (Row - CurrRow) = 0 Then
        grd.TopRow = (Row + 1) - CurrRow
    Else
        grd.TopRow = grd.Rows - (disRow * (Int(grd.Rows / Row)) - 1)
    End If

    Resume Next

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                     *
'*  SYSTEM DATE, TIME SETTING                          *
'*  (D0COM_SYSDATE, D0COM_SYSTIME)                         *
'*  D0COM_SERVER = 서버(SqlConn ismissting)              *
'*                                                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub D0SUB_SYSTEMDATE(frm As Form, Optional SqlConn As Variant)

    Dim SqlCode As Long
    Dim sStr  As String, sql_ret As Integer
    Dim record  As String, SqlData()    As String
    
    D0COM_SYSDATE = ""
    D0COM_SYSTIME = ""
    
    If IsMissing(SqlConn) Then
        If QSqlOpen(D0COM_SERVER01, frm.hWnd, SqlCode) <> QSQL_SUCCESS _
            Then Exit Sub
    Else
        SqlCode = SqlConn
    End If

    '현재의 SYSTEM일자와 시간을 읽어온다.
    sStr = "select convert(char(12),getdate(),102), convert(char(12),getdate(),108)"
    If QSqlDBExec(sStr, SqlCode) = QSQL_SUCCESS Then
        If QSqlGetRow(record, SqlCode) = QSQL_SUCCESS Then

            QSqlGetField 2, record, SqlData()

            D0COM_SYSDATE = Mid$(SqlData(1), 1, 4) & Mid$(SqlData(1), 6, 2) & Mid$(SqlData(1), 9, 2)
            D0COM_SYSTIME = Format$(SqlData(2), "HHMMDD")
        End If
    End If
    
    Call QSqlSelectFree(SqlCode)

    If IsMissing(SqlConn) Then Call Qsqlclose(SqlCode, ONECLOSE)

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값이 Length보다 짧으면 Left Space 채움 *
'*    길면 자름(한글 마지막 글자 처리)           *
'*    w_text    :  표시 대상 data                *
'*    w_len     :  표시 길이                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function D0SUB_LSPACE(w_text As String, w_len As Integer) As String
    
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer

    s_len = Len(w_text)                              ' length 계산
    
    If s_len <= w_len Then                           ' Left SPACE 채움
        D0SUB_LSPACE = Space$(w_len - s_len) + w_text
        Exit Function
    Else
        st = 0                  ' 한글 짤림여부 reset
        For i = 1 To w_len
            ch = Asc(Mid$(w_text, i, 1))
            If ch < &H80 Then                        ' 한글/영문 check ?
                st = 0                               ' 한글 짤림여부 reset
            Else
                st = (st + 1) Mod 2                  ' 한글 짤림여부 set
            End If
        Next i

        If st = 0 Then                               ' 한글 짤림여부 check ?
            D0SUB_LSPACE = Left$(w_text, w_len)      ' 마지막 한글 정상 set
        Else                                         ' 마지막 한글 자름
            D0SUB_LSPACE = " " + Left$(w_text, w_len - 1)
        End If
    End If
            
End Function


'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*    해당메세지를  선택하여 화면에 표시하기     *
'*    para  :  해당메세지 번호                   *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function D0SUB_MESSAGE(ByVal para As Integer) As String

    Select Case para
        Case 1: D0SUB_MESSAGE = "정상적으로 입력되었습니다.!!"
        Case 2: D0SUB_MESSAGE = "정상적으로 수정되었습니다.!!"
        Case 3: D0SUB_MESSAGE = "정상적으로 삭제되었습니다.!!"
        Case 4: D0SUB_MESSAGE = "정상적으로 조회되었습니다.!!"
        Case 5: D0SUB_MESSAGE = "정상적으로 인쇄되었습니다.!!"
        Case 6: D0SUB_MESSAGE = "해당 자료가 존재하지 않습니다.!!"
        Case 7: D0SUB_MESSAGE = "키값이 변경되었습니다! 확인바랍니다.!!"
        Case 8: D0SUB_MESSAGE = "삭제가 취소되었습니다.!!"
        Case 9: D0SUB_MESSAGE = "조회 데이타가 없습니다.!!"
        Case 10: D0SUB_MESSAGE = """ ' "" 는 입력할 수 없는 문자입니다.!!"
        Case 11: D0SUB_MESSAGE = "항목을 선택하여 주십시오.!!"
        Case 12: D0SUB_MESSAGE = "정상적으로 처리되었습니다.!!"
        Case 13: D0SUB_MESSAGE = "입력된 USER-ID는 권한이 없으므로 등록할 수 없습니다.!!"
        Case 14: D0SUB_MESSAGE = "선행번호 보다 작을 수 없읍니다.!!"
        Case 15: D0SUB_MESSAGE = "날짜입력이 틀립니다.  확인하세요.!!"
        Case 16: D0SUB_MESSAGE = "선행날짜보다 작을 수 없습니다.  확인하세요.!!"
        Case 17: D0SUB_MESSAGE = "출력 중 입니다.  잠시만 기다리세요.!!"
        Case 18: D0SUB_MESSAGE = "출력이 완료 되었습니다.!!"
        
        Case 101: D0SUB_MESSAGE = "과코드가 존재하지 않습니다.  확인하세요.!!"
        Case 102: D0SUB_MESSAGE = "SLIP코드가 존재하지 않습니다.  확인하세요.!!"
        Case 103: D0SUB_MESSAGE = "검체코드가 존재하지 않습니다.  확인하세요.!!"
        Case 104: D0SUB_MESSAGE = "업종코드가 존재하지 않습니다.  확인하세요.!!"
        Case 105: D0SUB_MESSAGE = "수가코드가 존재하지 않습니다.  확인하세요.!!"
        Case 106: D0SUB_MESSAGE = "검사코드가 존재하지 않습니다.  확인하세요.!!"
        Case 107: D0SUB_MESSAGE = "민원코드가 존재하지 않습니다.  확인하세요.!!"
        Case 108: D0SUB_MESSAGE = "민원종목코드가 존재하지 않습니다.  확인하세요.!!"
        Case 109: D0SUB_MESSAGE = "상병명코드가 존재하지 않습니다.  확인하세요.!!"
        Case 110: D0SUB_MESSAGE = "우편번호가 존재하지 않습니다.  확인하세요.!!"
        
        Case 536: D0SUB_MESSAGE = "DB OPEN이 되어있지 않습니다.!!"

    End Select

    
End Function

'****************************************************
'*                                                  *
'*  파일의 존재여부를 파악한다.                     *
'*  para : 파일명                                   *
'*                                                  *
'****************************************************
Function D0SUB_NULL_CHECK(para As Variant) As String

    If IsNull(para) Then
        D0SUB_NULL_CHECK = ""
    Else
        D0SUB_NULL_CHECK = para
    End If

End Function


'*********************************************************
'** 코드/명칭 help 화면 표시위치 계산 및 이동          ***
'** xpos      : help field 의 left position            ***
'** ypos      : help field 의 top position             ***
'** return    : 0:FAIL 1:TRUE                          ***
'*********************************************************
Sub D0SUB_POSITION(frm As Form, xpos As Long, YPos As Long)

    If (Screen.Width - xpos - 100) < frm.Width Then
        xpos = Screen.Width - frm.Width - 100
    Else
        xpos = xpos + 5
    End If
    If (Screen.Height - YPos - 200) < frm.Height Then
        YPos = YPos - frm.Height - 320
    Else
        YPos = YPos + 5
    End If

    frm.Move xpos, YPos

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값이 Length보다 짧으면 right space 채움 *
'*    길면 자름(한글 마지막 글자 처리)           *
'*    w_text    :  표시 대상 data                *
'*    w_len     :  표시 길이                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function D0SUB_RSPACE(w_text As String, w_len As Integer) As String
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer

    s_len = Len(w_text)                              ' length 계산
    
    If s_len <= w_len Then                           ' right SPACE 채움
        D0SUB_RSPACE = w_text + Space$(w_len - s_len)
        Exit Function
    Else
        st = 0                                       ' 한글 짤림여부 reset
        For i = 1 To w_len
            ch = Asc(Mid$(w_text, i, 1))
            If ch < &H80 Then                        ' 한글/영문 check ?
                st = 0                               ' 한글 짤림여부 reset
            Else
                st = (st + 1) Mod 2                  ' 한글 짤림여부 set
            End If
        Next i

        If st = 0 Then                               ' 한글 짤림여부 check ?
            D0SUB_RSPACE = Left$(w_text, w_len)      ' 마지막 한글 정상 set
        Else                                         ' 마지막 한글 자름
            D0SUB_RSPACE = Left$(w_text, w_len - 1) + " "
        End If
    End If

End Function

Function D0SUB_Len(w_text As String) As String
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer
    Dim t As String
    Dim c As String


    t = Right(Trim(w_text), 1)
    For i = 1 To Len(w_text) * 2
        c = Mid$(w_text, i, 1)
        If c = "" Then
            Exit For
        End If

        ch = Asc(c)
        If ch < 0 Then
            s_len = s_len + 2
        Else
            s_len = s_len + 1
        End If
        
    Next i
    
    D0SUB_Len = Str(s_len)
    
    
    'MsgBox Str(s_len)
    's_len = Len(w_text)                              ' length 계산
    
    'If s_len <= w_len Then                           ' right SPACE 채움
    '    D0SUB_RSPACE = w_text + Space$(w_len - s_len)
    '    Exit Function
    'Else
     '   st = 0                                       ' 한글 짤림여부 reset
    '    For i = 1 To w_len
    '        ch = Asc(Mid$(w_text, i, 1))
    '        If ch < &H80 Then                        ' 한글/영문 check ?
    '            st = 0                               ' 한글 짤림여부 reset
    '        Else
    '            st = (st + 1) Mod 2                  ' 한글 짤림여부 set
    '        End If
    '    Next i

    '    If st = 0 Then                               ' 한글 짤림여부 check ?
    '        D0SUB_RSPACE = Left$(w_text, w_len)      ' 마지막 한글 정상 set
    '    Else                                         ' 마지막 한글 자름
    '        D0SUB_RSPACE = Left$(w_text, w_len - 1) + " "
    '    End If
    'End If

End Function

'*----------------------------------------------------------*
'*                                                          *
'*  vaSpread에 데이타을 Clear한다.                       *
'*  spd : vaSpread명,  DispRow : vaSpread에 라인수    *
'*                                                          *
'*----------------------------------------------------------*
Sub D0SUB_Spread_Clear(spd As vaSpread, DispRow As Integer _
                      , Optional Col As Variant _
                      , Optional col2 As Variant _
                      , Optional Row As Variant _
                      , Optional row2 As Variant)
    
    If IsMissing(Col) Then Col = 1
    If IsMissing(Row) Then Row = 1
    
    If IsMissing(col2) Then col2 = spd.MaxCols
    If IsMissing(row2) Then row2 = spd.MaxRows
    
    With spd
        .Col = Col
        .col2 = col2
        .Row = Row
        .row2 = row2
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .MaxRows = DispRow
        .BlockMode = False
    End With

End Sub

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                              *
'*  vaSpread에서 특정행의 데이타을 찾고 위치을 돌려줌..      *
'*                                                              *
'*  spd          : vaSpread Name                             *
'*  row          : Row                                          *
'*  para         : Data                                         *
'*  Return Value : 데이타의 위치 열                             *
'*                 찾고자 하는 데이타가 없을 경우 -1를 SETTING  *
'*                                                              *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Function D0SUB_SPREADGETCOL(spd As vaSpread, Row As Long, para As Variant) As Integer
    
    Dim code    As Variant
    Dim Col     As Integer, sp As Boolean

    For Col = 1 To spd.MaxCols
        sp = spd.GetText(Col, Row, code)
        
        If Trim$(code) = Trim$(para) Then
            D0SUB_SPREADGETCOL = Col
            Exit Function
        End If
    Next

    D0SUB_SPREADGETCOL = -1

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                              *
'*  vaSpread에서 특정열의 데이타을 찾고 위치을 돌려줌..      *
'*                                                              *
'*  spd          : vaSpread Name                             *
'*  Col          : Column                                       *
'*  para         : Data                                         *
'*  Return Value : 데이타의 위치 행                             *
'*                 찾고자 하는 데이타가 없을 경우 -1를 SETTING  *
'*                                                              *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Function D0SUB_SPREADGETROW(spd As vaSpread, Col As Long, para As String) As Integer
    
    Dim code  As Variant
    Dim Row As Long, sp As Boolean

    For Row = 1 To spd.MaxRows
        
        sp = spd.GetText(Col, Row, code)
        
        If Trim$(code) = Trim$(para) Then
            D0SUB_SPREADGETROW = Row
            Exit Function
        End If
    Next

    D0SUB_SPREADGETROW = -1

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                      *
'*  vaSpread에서 특정행를 화면에 보여준다..          *
'*  spd : vaSpread Name                              *
'*  Row : Row                                           *
'*  disRow : 화면에 보여줄 수 있는 최대 행의 수         *
'*                                                      *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
Sub D0SUB_SPREADTOPROW(spd As vaSpread, Row As Integer, disRow As Integer)
    
    If Row < 1 Then Exit Sub
    
    On Error GoTo SpreadTopRowErr

    spd.Col = 1: spd.Row = Row - disRow + 1
    
    spd.Action = SS_ACTION_GOTO_CELL
    
    On Error GoTo 0
    
    Exit Sub

SpreadTopRowErr:

    Resume Next

End Sub

